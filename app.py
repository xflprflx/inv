import json
import os
import re
import sqlite3
import threading
import time
import webbrowser
from datetime import datetime
from threading import Timer
from flask import Flask, jsonify, render_template, request

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


# ==============================================================================
# GERENCIADOR DE ESTADO DE DEBUG E EXECUÇÃO PASSO A PASSO
# ==============================================================================

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
        self.step_event = threading.Event()

    def reset(self):
        """Reseta os indicadores antes de iniciar uma nova bateria na MI31."""
        self.is_running = True
        self.waiting_for_play = False
        self.current_step_num = 0
        self.current_step_desc = "Iniciando automação SAP..."
        self.error_message = None
        self.created_ficha_ids = []
        self.step_event.clear()

    def checkpoint(self, step_num: int, description: str, debug_config: dict):
        """
        Registra o passo atual da execução e aplica a pausa caso o Modo Debug esteja ativo.
        """
        self.current_step_num = step_num
        self.current_step_desc = description

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

    def trigger_play(self):
        """Libera imediatamente a trava do passo atual."""
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


def init_db():
    """
    Inicializa a tabela de fichas e garante a existência da coluna 'itens_json'
    para armazenar os itens detalhados que serão impressos na folha A4.
    """
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS fichas (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                numero_ficha TEXT NOT NULL,
                modo_execucao TEXT NOT NULL,
                centro TEXT NOT NULL,
                deposito TEXT NOT NULL,
                endereco_inicio TEXT,
                endereco_fim TEXT,
                total_materiais INTEGER DEFAULT 0,
                itens_json TEXT DEFAULT '[]',
                max_posicoes TEXT,
                data_criacao DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        # Migração segura caso o banco já exista de versões anteriores sem a coluna itens_json
        try:
            conn.execute("ALTER TABLE fichas ADD COLUMN itens_json TEXT DEFAULT '[]'")
        except sqlite3.OperationalError:
            pass
        conn.commit()


# ==============================================================================
# CAMADA DE INTEGRAÇÃO: AUTOMAÇÃO SAP GUI (MI31) E LEITURA DO RELATÓRIO
# ==============================================================================

class SapAutomator:
    """
    Encapsula a automação da MI31 e a extração de dados da tela final do SAP.
    """

    @staticmethod
    def extract_docs_and_items_from_screen(session, fallback_materials: list) -> dict:
        """
        Lê todos os controles GuiLabel da tela 'Documentos inventário físico criados diretamente',
        agrupa os elementos por linha (CharTop) e extrai o número da ficha e seus respectivos itens.
        Retorna um dicionário no formato:
        {
            "100061411": [
                {"material": "110.345", "lote": "RACV8", "posicao": "A.01.01.01"},
                ...
            ]
        }
        """
        usr_area = session.findById("wnd[0]/usr")
        rows_by_y = {}

        # 1. Agrupa todos os elementos de texto da tela pela coordenada vertical (linha Y)
        for i in range(usr_area.Children.Count):
            try:
                elem = usr_area.Children(i)
                text = str(elem.Text).strip()
                if not text:
                    continue
                y_coord = elem.CharTop
                x_coord = elem.CharLeft
                rows_by_y.setdefault(y_coord, []).append((x_coord, text))
            except Exception:
                continue

        docs_map = {}

        # 2. Percorre cada linha ordenada da esquerda para a direita (coordenada X)
        for y_coord in sorted(rows_by_y.keys()):
            row_cells = [val for _, val in sorted(rows_by_y[y_coord], key=lambda item: item[0])]
            
            # Procura na linha um número de documento de inventário MM (8 a 10 dígitos numéricos, ex: 100061411)
            doc_candidates = [c for c in row_cells if re.match(r"^\d{8,10}$", c)]
            if not doc_candidates:
                continue

            doc_num = doc_candidates[0]
            if doc_num not in docs_map:
                docs_map[doc_num] = []

            # Estrutura típica das colunas na linha de dados do relatório MI31:
            # [0]: Centro (4404) | [1]: Depósito (0001) | [2]: Material (110.345) | [3]: Lote (RACV8) ...
            material_code = row_cells[2] if len(row_cells) > 2 else ""
            batch_code = row_cells[3] if len(row_cells) > 3 and row_cells[3] != doc_num else "-"
            
            # Identifica posição de depósito na linha (se houver)
            bin_code = "-"
            if len(row_cells) >= 6:
                # Na tela padrão: [0]Pla, [1]Dep, [2]Material, [3]Lote, [4]TEs, [5]Pos.dpst.
                bin_code = row_cells[5]

            if material_code and material_code != doc_num:
                docs_map[doc_num].append({
                    "material": material_code,
                    "lote": batch_code,
                    "posicao": bin_code
                })

        # 3. Se algum documento foi identificado mas sem linhas detalhadas, usa a lista de entrada como fallback
        for doc_num in docs_map:
            if not docs_map[doc_num] and fallback_materials:
                docs_map[doc_num] = [
                    {"material": m, "lote": "-", "posicao": "-"} for m in fallback_materials
                ]

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

        # Inicializa o subsistema COM para a thread atual
        pythoncom.CoInitialize()

        try:
            # Conecta à sessão ativa do SAP GUI
            sap_gui_auto = win32com.client.GetObject("SAPGUI")
            application = sap_gui_auto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)

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

            if not docs_map:
                status_text = session.findById("wnd[0]/sbar").Text
                raise ValueError(f"Nenhuma ficha identificada na tela. Status SAP: '{status_text or 'Sem mensagem'}'")

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
            exec_ctrl.error_message = str(e)
        finally:
            exec_ctrl.waiting_for_play = False
            exec_ctrl.is_running = False
            pythoncom.CoUninitialize()


# ==============================================================================
# ROTAS FLASK (INTERFACE PRINCIPAL, CONTROLE DEBUG E IMPRESSÃO A4)
# ==============================================================================

@app.route("/")
def index():
    """Renderiza o painel de controle e a tabela de histórico."""
    with get_db() as conn:
        fichas = conn.execute("SELECT * FROM fichas ORDER BY id DESC LIMIT 30").fetchall()
    return render_template("index.html", fichas=fichas)


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


@app.route("/imprimir_a4")
def imprimir_a4():
    """
    Renderiza a página formatada para Folha A4 contendo:
    - Código de barras grande
    - Número da ficha grande abaixo do código
    - Range de endereço (apenas se definido)
    - Tabela de itens da ficha
    Aceita ?ids=1,2,3 para imprimir múltiplas fichas geradas no lote de uma só vez.
    """
    ids_param = request.args.get("ids", "")
    id_list = [int(x) for x in ids_param.split(",") if x.strip().isdigit()]

    if not id_list:
        return "Nenhuma ficha selecionada para impressão.", 400

    placeholders = ",".join(["?"] * len(id_list))
    with get_db() as conn:
        rows = conn.execute(f"SELECT * FROM fichas WHERE id IN ({placeholders}) ORDER BY id ASC", id_list).fetchall()

    fichas_formatadas = []
    for r in rows:
        try:
            itens = json.loads(r["itens_json"] or "[]")
        except Exception:
            itens = []

        fichas_formatadas.append({
            "id": r["id"],
            "numero_ficha": r["numero_ficha"],
            "centro": r["centro"],
            "deposito": r["deposito"],
            "endereco_inicio": (r["endereco_inicio"] or "").strip(),
            "endereco_fim": (r["endereco_fim"] or "").strip(),
            "data_criacao": r["data_criacao"],
            "itens": itens
        })

    return render_template("print_a4.html", fichas=fichas_formatadas)


def open_browser():
    """Abre o navegador padrão automaticamente ao inicializar o servidor."""
    webbrowser.open_new("http://127.0.0.1:5000")


if __name__ == "__main__":
    init_db()
    Timer(1.2, open_browser).start()
    # threaded=True é essencial para que as requisições de /api/debug/play respondam enquanto a thread SAP aguarda
    app.run(port=5000, debug=False, threaded=True)