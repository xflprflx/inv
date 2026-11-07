import json
import logging
import os
import re
import sqlite3
import threading
import time
import webbrowser
from datetime import datetime
from threading import Timer
from flask import Flask, Response, jsonify, render_template, request

import impressao

# Importação dos módulos nativos do Windows para Clipboard e Automação COM
try:
    import pythoncom        # Necessário para inicializar o modelo de threads COM (STA)
    import win32com.client  # Conexão COM com o motor de scripting do SAP GUI
    import win32clipboard   # Manipulação da área de transferência do Windows
except ImportError:
    pythoncom = None
    win32com = None
    win32clipboard = None

app = Flask(__name__)
DB_NAME = "inventario.db"
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "execucao.log")

# Log de diagnóstico da automação SAP (sobrescrito a cada nova execução)
logger = logging.getLogger("mi31")
logger.setLevel(logging.INFO)

# A impressora é confirmada pelo usuário a cada inicialização do programa
printer_state = {"confirmada": False}


# ==============================================================================
# GERENCIADOR DE ESTADO DE DEBUG E EXECUÇÃO PASSO A PASSO
# ==============================================================================

class ExecutionAborted(Exception):
    """Lançada nos checkpoints quando o usuário solicita o fechamento da conexão."""


class DebugExecutionController:
    """
    Controla o estado da automação em segundo plano e gerencia as pausas
    (por tempo determinado ou aguardando o clique no botão 'Play' do usuário).
    """

    def __init__(self):
        self.is_running = False
        self.waiting_for_play = False
        self.current_step_num = 0
        self.total_steps = 9
        self.current_step_desc = "Aguardando início..."
        self.error_message = None
        self.created_ficha_ids = []
        self.abort_requested = False
        self.sap_connection_id = None
        self.step_event = threading.Event()

    def reset(self):
        """Reseta os indicadores antes de iniciar uma nova bateria na MI31."""
        self.is_running = True
        self.waiting_for_play = False
        self.current_step_num = 0
        self.current_step_desc = "Iniciando automação SAP..."
        self.error_message = None
        self.created_ficha_ids = []
        self.abort_requested = False
        self.sap_connection_id = None
        self.step_event.clear()

    def checkpoint(self, step_num: int, description: str, debug_config: dict):
        """
        Registra o passo atual da execução e aplica a pausa caso o Modo Debug esteja ativo.
        """
        if self.abort_requested:
            raise ExecutionAborted("Execução cancelada: conexão SAP fechada pelo usuário.")

        self.current_step_num = step_num
        self.current_step_desc = description
        logger.info(description)

        # Se o modo debug não estiver ativo, segue imediatamente sem pausar
        if not debug_config.get("enabled", False):
            return

        mode = debug_config.get("mode", "infinite")
        if mode == "infinite":
            # Pausa infinita até que o usuário acione o endpoint /api/debug/play
            self.waiting_for_play = True
            self.step_event.clear()
            self.step_event.wait()
            self.waiting_for_play = False
        else:
            # Pausa pelo tempo em segundos definido pelo usuário (ou avança antes se clicar em Play)
            seconds = float(debug_config.get("seconds", 2.0))
            self.waiting_for_play = True
            self.step_event.clear()
            self.step_event.wait(timeout=max(0.1, seconds))
            self.waiting_for_play = False

        if self.abort_requested:
            raise ExecutionAborted("Execução cancelada: conexão SAP fechada pelo usuário.")

    def trigger_play(self):
        """Libera imediatamente a trava do passo atual."""
        self.step_event.set()

    def request_abort(self):
        """Sinaliza o cancelamento e destrava qualquer pausa de debug pendente."""
        self.abort_requested = True
        self.step_event.set()


# Instância global do controlador de execução
exec_ctrl = DebugExecutionController()


# ==============================================================================
# CAMADA DE UTILITÁRIOS: ÁREA DE TRANSFERÊNCIA DO WINDOWS
# ==============================================================================

def set_windows_clipboard(text: str):
    """
    Injeta uma lista de textos no Clipboard do Windows com codificação Unicode.
    Permite colar centenas de códigos de materiais na seleção múltipla do SAP instantaneamente.
    """
    if not win32clipboard:
        raise RuntimeError("Módulo win32clipboard indisponível no ambiente.")

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


# ==============================================================================
# CAMADA DE DADOS: PERSISTÊNCIA SQLITE
# ==============================================================================

def get_db():
    """Conecta ao banco de dados SQLite local retornando linhas mapeadas por nome."""
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn


FICHAS_COLUNAS = {
    "numero_ficha": "TEXT NOT NULL DEFAULT ''",
    "modo_execucao": "TEXT NOT NULL DEFAULT ''",
    "centro": "TEXT NOT NULL DEFAULT ''",
    "deposito": "TEXT NOT NULL DEFAULT ''",
    "endereco_inicio": "TEXT",
    "endereco_fim": "TEXT",
    "total_materiais": "INTEGER DEFAULT 0",
    "itens_json": "TEXT DEFAULT '[]'",
    "max_posicoes": "TEXT",
    "data_criacao": "DATETIME DEFAULT CURRENT_TIMESTAMP",
}


def init_db():
    """
    Cria/atualiza as tabelas do SQLite.
    Bancos de versões anteriores recebem as colunas que faltam; se a tabela antiga
    tiver colunas obrigatórias incompatíveis, ela é preservada com outro nome
    (fichas_antigo_AAAAMMDDHHMMSS) e uma tabela nova é criada.
    """
    colunas_sql = ",\n                ".join(f"{nome} {tipo}" for nome, tipo in FICHAS_COLUNAS.items())
    create_fichas = f"""
            CREATE TABLE IF NOT EXISTS fichas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                {colunas_sql}
            )
        """
    with get_db() as conn:
        conn.execute(create_fichas)

        existentes = {row["name"]: row for row in conn.execute("PRAGMA table_info(fichas)")}
        incompativeis = [
            nome for nome, row in existentes.items()
            if nome != "id" and nome not in FICHAS_COLUNAS and row["notnull"] and row["dflt_value"] is None
        ]
        if incompativeis:
            backup = "fichas_antigo_" + datetime.now().strftime("%Y%m%d%H%M%S")
            conn.execute(f"ALTER TABLE fichas RENAME TO {backup}")
            conn.execute(create_fichas)
            print(f"[DB] Tabela antiga incompatível preservada como '{backup}'.")
        else:
            for nome, tipo in FICHAS_COLUNAS.items():
                if nome not in existentes:
                    tipo_alter = "TEXT" if nome == "data_criacao" else tipo  # SQLite não aceita DEFAULT não constante em ALTER
                    conn.execute(f"ALTER TABLE fichas ADD COLUMN {nome} {tipo_alter}")
                    print(f"[DB] Coluna '{nome}' adicionada à tabela fichas.")

        conn.execute("""
            CREATE TABLE IF NOT EXISTS config (
                chave TEXT PRIMARY KEY,
                valor TEXT
            )
        """)
        conn.commit()


def get_config(chave: str, default: str = "") -> str:
    with get_db() as conn:
        row = conn.execute("SELECT valor FROM config WHERE chave = ?", (chave,)).fetchone()
    return row["valor"] if row and row["valor"] is not None else default


def set_config(chave: str, valor: str):
    with get_db() as conn:
        conn.execute(
            "INSERT INTO config (chave, valor) VALUES (?, ?) "
            "ON CONFLICT(chave) DO UPDATE SET valor = excluded.valor",
            (chave, valor),
        )
        conn.commit()


def carregar_fichas(id_list: list) -> list:
    """Busca as fichas no SQLite já com a lista de itens decodificada."""
    if not id_list:
        return []
    placeholders = ",".join(["?"] * len(id_list))
    with get_db() as conn:
        rows = conn.execute(f"SELECT * FROM fichas WHERE id IN ({placeholders}) ORDER BY id ASC", id_list).fetchall()

    fichas = []
    for r in rows:
        try:
            itens = json.loads(r["itens_json"] or "[]")
        except Exception:
            itens = []
        fichas.append({
            "id": r["id"],
            "numero_ficha": r["numero_ficha"],
            "centro": r["centro"],
            "deposito": r["deposito"],
            "endereco_inicio": (r["endereco_inicio"] or "").strip(),
            "endereco_fim": (r["endereco_fim"] or "").strip(),
            "data_criacao": r["data_criacao"],
            "itens": itens,
        })
    return fichas


def parse_ids(raw) -> list:
    if isinstance(raw, list):
        raw = ",".join(str(x) for x in raw)
    return [int(x) for x in str(raw or "").split(",") if x.strip().isdigit()]


# ==============================================================================
# CAMADA DE INTEGRAÇÃO: AUTOMAÇÃO SAP GUI (MI31) E LEITURA DO RELATÓRIO
# ==============================================================================

class SapAutomator:
    """
    Encapsula a automação da MI31 e a extração de dados da tela final do SAP.
    """

    @staticmethod
    def dump_screen(session) -> None:
        """Registra no log tudo o que está na tela atual do SAP (para diagnóstico)."""
        try:
            wnd = session.findById("wnd[0]")
            logger.info("Transação: %s | Título da janela: %s", session.Info.Transaction, wnd.Text)
            logger.info("Barra de status: [%s] %s", session.findById("wnd[0]/sbar").MessageType,
                        session.findById("wnd[0]/sbar").Text)
            try:
                popup = session.findById("wnd[1]")
                logger.info("Popup aberto (wnd[1]): %s", popup.Text)
            except Exception:
                pass

            usr = session.findById("wnd[0]/usr")
            logger.info("Elementos em wnd[0]/usr: %d", usr.Children.Count)
            for i in range(min(usr.Children.Count, 400)):
                try:
                    el = usr.Children(i)
                    pos = ""
                    try:
                        pos = f"linha={el.CharTop} col={el.CharLeft} "
                    except Exception:
                        pass
                    logger.info("  %s %s%s = %r", el.Type, pos, el.Id.split("/usr/")[-1], str(el.Text)[:120])
                except Exception as e:
                    logger.info("  (falha ao ler elemento %d: %s)", i, e)
        except Exception as e:
            logger.info("Falha ao registrar a tela do SAP: %s", e)

    # Cabeçalhos da lista "Documentos inventário físico criados diretamente" (MI31)
    LIST_HEADERS = {
        "Pla.": "centro",
        "Dep.": "deposito",
        "Material": "material",
        "Lote": "lote",
        "TEs": "tipo_estoque",
        "Pos.dpst.": "posicao",
        "Grp.merc.": "grupo_mercadoria",
        "Doc.inv.": "documento",
    }

    @staticmethod
    def _read_visible_rows(usr_area) -> dict:
        """Retorna {linha_visivel: [(coluna, texto), ...]} com os GuiLabel da tela."""
        rows = {}
        for i in range(usr_area.Children.Count):
            try:
                elem = usr_area.Children(i)
                text = str(elem.Text).strip()
                if text:
                    rows.setdefault(elem.CharTop, []).append((elem.CharLeft, text))
            except Exception:
                continue
        return {y: sorted(cells) for y, cells in rows.items()}

    @classmethod
    def extract_docs_and_items_from_screen(cls, session, fallback_materials: list) -> dict:
        """
        Lê a lista de saída da MI31 usando as colunas do cabeçalho
        (Pla. | Dep. | Material | Lote | TEs | Pos.dpst. | Grp.merc. | Doc.inv.)
        e rola a tela até o fim para capturar todas as linhas.
        Retorna {"100061266": [{"material": ..., "lote": ..., "posicao": ...}, ...]}.
        """
        usr_area = session.findById("wnd[0]/usr")
        header_cols = None   # {coluna: campo}
        seen = set()
        docs_map = {}
        total_sap = None

        try:
            scrollbar = usr_area.VerticalScrollbar
            max_pos = int(scrollbar.Maximum)
            page = max(1, int(scrollbar.PageSize) - 1)
        except Exception:
            scrollbar, max_pos, page = None, 0, 1

        pos = 0
        while True:
            if scrollbar is not None:
                try:
                    scrollbar.Position = pos
                    usr_area = session.findById("wnd[0]/usr")
                except Exception as e:
                    logger.info("Falha ao rolar a lista para a posição %s: %s", pos, e)
                    break

            for _, cells in sorted(cls._read_visible_rows(usr_area).items()):
                texts = [t for _, t in cells]

                # Estatística do SAP: quantidade de itens gerados
                joined = " ".join(texts)
                if total_sap is None and "gerados diretam" in joined:
                    total_sap = "aguardando"
                    continue
                if total_sap == "aguardando":
                    m = re.search(r"(\d+)\s*$", joined)
                    total_sap = int(m.group(1)) if m else None

                # Linha de cabeçalho (define a posição de cada coluna)
                if "Doc.inv." in texts and "Material" in texts:
                    header_cols = {col: cls.LIST_HEADERS[t] for col, t in cells if t in cls.LIST_HEADERS}
                    continue
                if not header_cols:
                    continue

                # Associa cada célula ao cabeçalho mais próximo à esquerda
                header_positions = sorted(header_cols)
                row = {}
                for col, text in cells:
                    candidatos = [h for h in header_positions if h <= col]
                    if candidatos:
                        row[header_cols[candidatos[-1]]] = text

                doc_num = row.get("documento", "")
                if not re.fullmatch(r"\d{8,10}", doc_num) or not row.get("material"):
                    continue  # linhas de status ("Criar doc."), totais etc.

                key = tuple(row.get(k, "") for k in ("documento", "material", "lote", "tipo_estoque", "posicao"))
                if key in seen:
                    continue
                seen.add(key)
                docs_map.setdefault(doc_num, []).append({
                    "material": row.get("material", ""),
                    "lote": row.get("lote") or "-",
                    "posicao": row.get("posicao") or "-",
                    "tipo_estoque": row.get("tipo_estoque", ""),
                })

            if scrollbar is None or pos >= max_pos:
                break
            pos = min(pos + page, max_pos)

        # Volta a lista para o topo
        if scrollbar is not None:
            try:
                scrollbar.Position = 0
            except Exception:
                pass

        lidos = sum(len(v) for v in docs_map.values())
        logger.info("Itens lidos da lista: %d | Itens informados pelo SAP: %s", lidos, total_sap)
        if isinstance(total_sap, int) and total_sap != lidos:
            logger.info("ATENÇÃO: quantidade lida (%d) difere da estatística do SAP (%d).", lidos, total_sap)

        # Documento identificado sem itens: usa a lista de entrada como fallback
        for doc_num in docs_map:
            if not docs_map[doc_num] and fallback_materials:
                docs_map[doc_num] = [{"material": m, "lote": "-", "posicao": "-"} for m in fallback_materials]

        return docs_map

    @classmethod
    def run_mi31_workflow(cls, p: dict, debug_cfg: dict):
        """
        Executa todo o roteiro da transação MI31 em uma thread dedicada COM,
        respeitando os checkpoints do Modo Debug e salvando os resultados no SQLite.
        """
        if not win32com or not pythoncom:
            exec_ctrl.error_message = "Módulos pywin32 (win32com / pythoncom) indisponíveis. Execute em Windows."
            exec_ctrl.is_running = False
            return

        # Reinicia o log de diagnóstico desta execução
        for h in list(logger.handlers):
            logger.removeHandler(h)
            h.close()
        handler = logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", "%H:%M:%S"))
        logger.addHandler(handler)
        safe_params = {k: v for k, v in p.items() if k != "materials_list"}
        logger.info("Parâmetros: %s", json.dumps(safe_params, ensure_ascii=False))

        # Inicializa o subsistema COM para a thread atual
        pythoncom.CoInitialize()
        sap_gui_auto = application = connection = session = None

        try:
            # Conecta à sessão ativa do SAP GUI
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
            exec_ctrl.sap_connection_id = connection.Id

            # PASSO 1: Abrir transação MI31
            exec_ctrl.checkpoint(1, "Passo 1/9: Maximizando janela e abrindo transação /nMI31", debug_cfg)
            session.findById("wnd[0]").maximize()
            session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmi31"
            session.findById("wnd[0]").sendVKey(0)

            # Prepara lista de materiais informada pelo usuário
            materials_text = p.get("materials_list", "").strip()
            mat_items = [m.strip() for m in re.split(r"[\r\n,;]+", materials_text) if m.strip()] if materials_text else []

            # PASSO 2: Inserir lista de materiais na seleção múltipla (R_MATNR)
            exec_ctrl.checkpoint(2, f"Passo 2/9: Carregando lista de materiais ({len(mat_items)} itens)", debug_cfg)
            try:
                session.findById("wnd[0]/usr/btn%_R_MATNR_%_APP_%-VALU_PUSH").Press()
                try:
                    session.findById("wnd[1]/tbar[0]/btn[16]").Press()  # Limpar seleção existente
                except Exception:
                    pass

                if mat_items:
                    set_windows_clipboard("\r\n".join(mat_items))
                    session.findById("wnd[1]/tbar[0]/btn[24]").Press()  # Colar do Clipboard

                session.findById("wnd[1]/tbar[0]/btn[8]").Press()       # Confirmar (F8)
            except Exception:
                pass

            # PASSO 3: Preencher Centro (R_WERKS) e Depósito (R_LGORT)
            exec_ctrl.checkpoint(3, f"Passo 3/9: Definindo Centro ({p.get('plant')}) e Depósito ({p.get('storage_loc')})", debug_cfg)
            try:
                session.findById("wnd[0]/usr/btn%_R_WERKS_%_APP_%-VALU_PUSH").Press()
                session.findById("wnd[1]/tbar[0]/btn[16]").Press()  # Limpar seleção existente
                session.findById("wnd[1]/tbar[0]/btn[8]").Press()
            except Exception:
                pass
            session.findById("wnd[0]/usr/ctxtR_WERKS-LOW").Text = p.get("plant", "4404")

            try:
                session.findById("wnd[0]/usr/btn%_R_LGORT_%_APP_%-VALU_PUSH").Press()
                session.findById("wnd[1]/tbar[0]/btn[16]").Press()  # Limpar seleção existente
                session.findById("wnd[1]/tbar[0]/btn[8]").Press()
            except Exception:
                pass
            session.findById("wnd[0]/usr/ctxtR_LGORT-LOW").Text = p.get("storage_loc", "0001")

            # PASSO 4: Preencher Range de Endereços (R_LGPBE)
            bin_start = p.get("bin_start", "").strip()
            bin_end = p.get("bin_end", "").strip()
            exec_ctrl.checkpoint(4, f"Passo 4/9: Configurando Range de Endereços ({bin_start or 'Vazio'} até {bin_end or 'Vazio'})", debug_cfg)
            try:
                session.findById("wnd[0]/usr/btn%_R_LGPBE_%_APP_%-VALU_PUSH").Press()
                session.findById("wnd[1]/tbar[0]/btn[16]").Press()  # Limpar seleção existente
                session.findById("wnd[1]/tbar[0]/btn[8]").Press()
            except Exception:
                pass
            session.findById("wnd[0]/usr/txtR_LGPBE-LOW").Text = bin_start
            session.findById("wnd[0]/usr/txtR_LGPBE-HIGH").Text = bin_end

            # PASSO 5: Configurar Modo de Execução (CALL_FB vs LISTE) e Limite MAXPO
            exec_mode = p.get("exec_mode", "CALL_FB")
            exec_ctrl.checkpoint(5, f"Passo 5/9: Selecionando Modo ({exec_mode}) e Máx. Posições ({p.get('max_positions')})", debug_cfg)
            session.findById("wnd[0]/usr/chkXDELE").Selected = p.get("chkXDELE", True)

            if exec_mode == "CALL_FB":
                session.findById("wnd[0]/usr/radCALL_FB").Select()
                session.findById("wnd[0]").sendVKey(0)
                try:
                    session.findById("wnd[0]/usr/txtMAXPO").Text = str(p.get("max_positions", "333"))
                except Exception:
                    session.findById("wnd[0]/usr/txtMAXPO-LOW").Text = str(p.get("max_positions", "333"))
                if p.get("anziv"):
                    session.findById("wnd[0]/usr/txtANZIV").Text = str(p.get("anziv"))
            else:
                session.findById("wnd[0]/usr/radLISTE").Select()
                session.findById("wnd[0]").sendVKey(0)

            session.findById("wnd[0]/usr/chkXKEEP").Selected = p.get("chkXKEEP", False)
            session.findById("wnd[0]/usr/chkXPROT").Selected = p.get("chkXPROT", True)

            # PASSO 6: Parâmetros de Seleção de Estoque (PA_PBSTO)
            exec_ctrl.checkpoint(6, "Passo 6/9: Configurando Parâmetros de Estoque e Bloqueio (PA_PBSTO)", debug_cfg)
            session.findById("wnd[0]/usr/btnPA_PBSTO").Press()
            session.findById("wnd[0]/usr/chkXUMKE").Selected = p.get("chkXUMKE", False)
            session.findById("wnd[0]/usr/chkXUMKEW").Selected = p.get("chkXUMKEW", False)
            session.findById("wnd[0]/usr/chkXNULB").Selected = p.get("chkXNULB", False)
            session.findById("wnd[0]/usr/chkXONUL").Selected = p.get("chkXONUL", False)
            session.findById("wnd[0]/usr/chkXNEGB").Selected = p.get("chkXNEGB", False)
            session.findById("wnd[0]/usr/chkXIMAT").Selected = p.get("chkXIMAT", True)
            session.findById("wnd[0]/usr/chkXICHA").Selected = p.get("chkXICHA", True)
            session.findById("wnd[0]/usr/chkXINSME").Selected = p.get("chkXINSME", False)
            session.findById("wnd[0]/usr/chkXLABST").Selected = p.get("chkXLABST", True)
            session.findById("wnd[0]/usr/chkXSPEME").Selected = p.get("chkXSPEME", True)
            session.findById("wnd[0]/usr/chkXWART").Selected = p.get("chkXWART", False)
            session.findById("wnd[0]/usr/chkSPERR").Selected = p.get("chkSPERR", True)
            session.findById("wnd[0]/usr/chkXBUFI").Selected = p.get("chkXBUFI", False)
            if p.get("txtXBLNI"):
                session.findById("wnd[0]/usr/txtXBLNI").Text = str(p.get("txtXBLNI"))

            # PASSO 7: Parâmetros de Ordenação (PA_PBSRT)
            exec_ctrl.checkpoint(7, "Passo 7/9: Aplicando Ordenação por Posição no Depósito (PA_PBSRT)", debug_cfg)
            session.findById("wnd[0]/usr/btnPA_PBSRT").Press()
            session.findById("wnd[0]/usr/radXLGPBE").Select()
            session.findById("wnd[0]/usr/chkXSAVE1").Selected = p.get("chkXSAVE1", False)
            session.findById("wnd[0]/usr/chkXSAVE2").Selected = p.get("chkXSAVE2", False)
            session.findById("wnd[0]/usr/chkXGRUPP").Selected = p.get("chkXGRUPP", False)
            session.findById("wnd[0]/usr/radXSORT").Select()
            if p.get("pa_vari"):
                session.findById("wnd[0]/usr/ctxtPA_VARI").Text = str(p.get("pa_vari"))

            # PASSO 8: Executar F8 no SAP
            exec_ctrl.checkpoint(8, "Passo 8/9: Executando geração do lote (Tecla F8)", debug_cfg)
            session.findById("wnd[0]/tbar[1]/btn[8]").Press()

            # PASSO 9: Extração dos Documentos e Itens da Tela de Saída
            exec_ctrl.checkpoint(9, "Passo 9/9: Lendo fichas e materiais gerados na tela de saída", debug_cfg)
            docs_map = cls.extract_docs_and_items_from_screen(session, mat_items)
            status_text = session.findById("wnd[0]/sbar").Text

            if not docs_map:
                # Plano B: número do documento informado na barra de status do SAP
                for doc_num in re.findall(r"\b\d{8,10}\b", status_text or ""):
                    docs_map[doc_num] = [{"material": m, "lote": "-", "posicao": "-"} for m in mat_items]
                if docs_map:
                    logger.info("Fichas obtidas pela barra de status: %s", list(docs_map))

            if not docs_map:
                raise ValueError(f"Nenhuma ficha identificada na tela. Status SAP: '{status_text or 'Sem mensagem'}'")

            logger.info("Fichas identificadas: %s", {k: len(v) for k, v in docs_map.items()})

            # Grava cada ficha e sua lista de itens no SQLite para geração da folha A4
            created_ids = []
            with get_db() as conn:
                for doc_num, items_list in docs_map.items():
                    cursor = conn.execute(
                        """
                        INSERT INTO fichas (
                            numero_ficha, modo_execucao, centro, deposito,
                            endereco_inicio, endereco_fim, total_materiais,
                            itens_json, max_posicoes
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            doc_num,
                            exec_mode,
                            p.get("plant", "4404"),
                            p.get("storage_loc", "0001"),
                            bin_start,
                            bin_end,
                            len(items_list),
                            json.dumps(items_list, ensure_ascii=False),
                            str(p.get("max_positions", "333"))
                        )
                    )
                    created_ids.append(cursor.lastrowid)
                conn.commit()

            exec_ctrl.created_ficha_ids = created_ids
            exec_ctrl.current_step_desc = f"Concluído! Ficha(s) gerada(s): {', '.join(docs_map.keys())}"

        except Exception as e:
            exec_ctrl.error_message = f"{exec_ctrl.current_step_desc} -> {e}"
            logger.exception("Erro na execução")
            if session is not None and not isinstance(e, ExecutionAborted):
                cls.dump_screen(session)
        finally:
            logger.info("Fim da execução.")
            exec_ctrl.waiting_for_play = False
            exec_ctrl.is_running = False
            # Libera as referências COM antes de desinicializar a thread
            session = connection = application = sap_gui_auto = None
            pythoncom.CoUninitialize()


def close_sap_connection() -> str:
    """
    Fecha a conexão do SAP GUI usada pela automação (ou a primeira conexão aberta).
    Atenção: encerra a sessão logada no SAP, descartando dados não salvos.
    """
    if not win32com or not pythoncom:
        raise RuntimeError("Módulos pywin32 (win32com / pythoncom) indisponíveis. Execute em Windows.")

    pythoncom.CoInitialize()
    sap_gui_auto = application = connection = None
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application = sap_gui_auto.GetScriptingEngine

        if exec_ctrl.sap_connection_id:
            try:
                connection = application.findById(exec_ctrl.sap_connection_id)
            except Exception:
                connection = None
        if connection is None:
            if application.Children.Count == 0:
                return "Nenhuma conexão SAP aberta."
            connection = application.Children(0)

        description = connection.Description or connection.Id
        connection.CloseConnection()
        exec_ctrl.sap_connection_id = None
        return f"Conexão SAP '{description}' fechada."
    finally:
        connection = application = sap_gui_auto = None
        pythoncom.CoUninitialize()


# ==============================================================================
# ROTAS FLASK (INTERFACE PRINCIPAL, CONTROLE DEBUG E IMPRESSÃO A4)
# ==============================================================================

@app.route("/")
def index():
    """Renderiza o painel de controle e a tabela de histórico."""
    with get_db() as conn:
        fichas = conn.execute("SELECT * FROM fichas ORDER BY id DESC LIMIT 30").fetchall()
    return render_template(
        "index.html",
        fichas=fichas,
        solicitar_impressora=not printer_state["confirmada"],
        impressora_atual=get_config("impressora"),
    )


@app.route("/api/iniciar", methods=["POST"])
def api_iniciar():
    """
    Inicia a thread de automação SAP MI31 com os parâmetros da UI e configurações de Debug.
    """
    if exec_ctrl.is_running:
        return jsonify({"success": False, "message": "Já existe uma execução em andamento."}), 400

    p = request.json
    if not p.get("plant") or not p.get("storage_loc"):
        return jsonify({"success": False, "message": "Preencha obrigatoriamente o Centro e o Depósito."}), 400

    debug_cfg = {
        "enabled": bool(p.get("debug_enabled", False)),
        "mode": p.get("debug_mode", "infinite"),
        "seconds": float(p.get("debug_seconds", 2.0) or 2.0)
    }

    exec_ctrl.reset()
    worker = threading.Thread(target=SapAutomator.run_mi31_workflow, args=(p, debug_cfg), daemon=True)
    worker.start()

    return jsonify({"success": True})


@app.route("/api/debug/status", methods=["GET"])
def api_debug_status():
    """
    Retorna o estado em tempo real da execução para o painel de Debug da UI.
    """
    return jsonify({
        "is_running": exec_ctrl.is_running,
        "waiting_for_play": exec_ctrl.waiting_for_play,
        "step_num": exec_ctrl.current_step_num,
        "total_steps": exec_ctrl.total_steps,
        "step_desc": exec_ctrl.current_step_desc,
        "error": exec_ctrl.error_message,
        "created_ids": exec_ctrl.created_ficha_ids
    })


@app.route("/api/debug/play", methods=["POST"])
def api_debug_play():
    """
    Acionado pelo botão 'Play / Próximo Passo' na UI para destravar a pausa do Debug.
    """
    exec_ctrl.trigger_play()
    return jsonify({"success": True})


@app.route("/api/sap/fechar", methods=["POST"])
def api_sap_fechar():
    """
    Cancela a execução em andamento (se houver) e fecha a conexão do SAP GUI.
    """
    if exec_ctrl.is_running:
        exec_ctrl.request_abort()
        # Aguarda a thread chegar a um checkpoint e encerrar (máx. ~3s)
        for _ in range(30):
            if not exec_ctrl.is_running:
                break
            time.sleep(0.1)

    try:
        message = close_sap_connection()
    except Exception as e:
        return jsonify({"success": False, "message": f"Falha ao fechar conexão SAP: {e}"}), 500

    return jsonify({"success": True, "message": message})


@app.route("/log")
def ver_log():
    """Exibe o log de diagnóstico da última execução da MI31."""
    try:
        with open(LOG_FILE, encoding="utf-8") as fh:
            conteudo = fh.read()
    except FileNotFoundError:
        conteudo = "Nenhuma execução registrada ainda."
    return Response(conteudo, mimetype="text/plain; charset=utf-8")


@app.route("/api/impressoras", methods=["GET"])
def api_impressoras():
    """Lista as impressoras do Windows e a impressora atualmente selecionada."""
    impressoras = impressao.listar_impressoras()
    selecionada = get_config("impressora")
    if selecionada not in impressoras:
        selecionada = ""
    return jsonify({
        "disponivel": impressao.impressao_disponivel(),
        "impressoras": impressoras,
        "padrao": impressao.impressora_padrao(),
        "selecionada": selecionada,
    })


@app.route("/api/impressora", methods=["POST"])
def api_impressora():
    """Salva a impressora escolhida (vazio = continuar sem impressão direta)."""
    nome = (request.json or {}).get("impressora", "").strip()
    if nome and nome not in impressao.listar_impressoras():
        return jsonify({"success": False, "message": f"Impressora '{nome}' não encontrada."}), 400
    set_config("impressora", nome)
    printer_state["confirmada"] = True
    return jsonify({"success": True, "impressora": nome})


@app.route("/fichas/pdf")
def fichas_pdf():
    """
    Gera o PDF A4 das fichas (?ids=1,2,3).
    - Sem parâmetros extras: abre no visualizador de PDF do navegador.
    - Com ?download=1: força o download do arquivo.
    """
    fichas = carregar_fichas(parse_ids(request.args.get("ids", "")))
    if not fichas:
        return "Nenhuma ficha selecionada.", 400

    pdf_bytes = impressao.gerar_pdf_fichas(fichas)
    numeros = "_".join(f["numero_ficha"] for f in fichas[:5])
    if len(fichas) > 5:
        numeros += f"_e_mais_{len(fichas) - 5}"
    filename = f"fichas_MI31_{numeros}.pdf"
    disposition = "attachment" if request.args.get("download") else "inline"
    return Response(
        pdf_bytes,
        mimetype="application/pdf",
        headers={"Content-Disposition": f'{disposition}; filename="{filename}"'},
    )


@app.route("/api/imprimir", methods=["POST"])
def api_imprimir():
    """Imprime as fichas diretamente na impressora selecionada, sem diálogo."""
    body = request.json or {}
    fichas = carregar_fichas(parse_ids(body.get("ids", "")))
    if not fichas:
        return jsonify({"success": False, "message": "Nenhuma ficha selecionada."}), 400

    impressora = get_config("impressora")
    if not impressora:
        return jsonify({"success": False, "message": "Nenhuma impressora selecionada. Clique em 'Trocar impressora'."}), 400

    try:
        pdf_bytes = impressao.gerar_pdf_fichas(fichas)
        paginas = impressao.imprimir_pdf(pdf_bytes, impressora, titulo=f"Fichas MI31 ({len(fichas)})")
    except Exception as e:
        return jsonify({"success": False, "message": f"Falha ao imprimir: {e}"}), 500

    return jsonify({
        "success": True,
        "message": f"{len(fichas)} ficha(s) enviada(s) para '{impressora}' ({paginas} página(s)).",
    })


@app.route("/imprimir_a4")
def imprimir_a4():
    """Compatibilidade: links antigos passam a abrir o PDF."""
    return fichas_pdf()


def open_browser():
    """Abre o navegador padrão automaticamente ao inicializar o servidor."""
    webbrowser.open_new("http://127.0.0.1:5000")


if __name__ == "__main__":
    init_db()
    Timer(1.2, open_browser).start()
    # threaded=True é essencial para que as requisições de /api/debug/play respondam enquanto a thread SAP aguarda
    app.run(port=5000, debug=False, threaded=True)