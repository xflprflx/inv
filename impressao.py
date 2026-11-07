"""
Geração do PDF A4 das fichas de inventário e impressão direta em impressora do Windows.

- O PDF é gerado com reportlab (código de barras Code 128 nativo, sem depender de internet).
- A impressão direta renderiza cada página do PDF com pypdfium2 e envia a imagem
  para a impressora selecionada via GDI (pywin32), sem abrir diálogo de impressão.
"""

import io
import threading

from reportlab.graphics.barcode.code128 import Code128
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    Flowable,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from xml.sax.saxutils import escape

try:
    import win32con
    import win32gui
    import win32print
    import win32ui
except ImportError:
    win32con = win32gui = win32print = win32ui = None


# ==============================================================================
# GERAÇÃO DO PDF A4
# ==============================================================================

PAGE_MARGIN = 12 * mm
CONTENT_WIDTH = A4[0] - 2 * PAGE_MARGIN

_st_title = ParagraphStyle("title", fontName="Helvetica-Bold", fontSize=15, leading=18)
_st_sub = ParagraphStyle("sub", fontName="Helvetica", fontSize=9, leading=12)
_st_meta = ParagraphStyle("meta", fontName="Helvetica", fontSize=9.5, leading=13, alignment=TA_RIGHT)
_st_label = ParagraphStyle("label", fontName="Helvetica-Bold", fontSize=8.5, leading=11, alignment=TA_CENTER)
_st_giant = ParagraphStyle("giant", fontName="Helvetica-Bold", fontSize=34, leading=40, alignment=TA_CENTER)
_st_range = ParagraphStyle("range", fontName="Helvetica-Bold", fontSize=17, leading=21, alignment=TA_CENTER)
_st_section = ParagraphStyle("section", fontName="Helvetica-Bold", fontSize=10.5, leading=14, spaceAfter=4)
_st_cell = ParagraphStyle("cell", fontName="Helvetica", fontSize=9, leading=11)
_st_cell_bold = ParagraphStyle("cellb", fontName="Helvetica-Bold", fontSize=9, leading=11)
_st_th = ParagraphStyle("th", fontName="Helvetica-Bold", fontSize=8, leading=10)
_st_sign = ParagraphStyle("sign", fontName="Helvetica-Bold", fontSize=8.5, leading=11, alignment=TA_CENTER)


def _p(text, style):
    return Paragraph(escape(str(text if text is not None else "")), style)


def _barcode(value: str):
    """Code 128 dimensionado para ~120 mm de largura, centralizado."""
    target_width = 120 * mm
    probe = Code128(value, barWidth=1, barHeight=28 * mm, quiet=False)
    bar_width = min(target_width / max(probe.width, 1), 1.2 * mm)
    code = Code128(value, barWidth=bar_width, barHeight=28 * mm, quiet=False)
    holder = Table([[code]], colWidths=[CONTENT_WIDTH - 6 * mm])
    holder.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    return holder


class _MarcaFicha(Flowable):
    """Marcador invisível que informa ao documento qual ficha está sendo desenhada."""

    def __init__(self, numero: str):
        super().__init__()
        self.numero = numero

    def wrap(self, *args):
        return 0, 0

    def draw(self):
        doc = self.canv._doctemplate
        doc.ficha_atual = self.numero
        doc.pagina_ficha = 0


class _FichasDocTemplate(SimpleDocTemplate):
    """Documento que escreve 'Ficha X - página n' no rodapé de todas as páginas."""

    ficha_atual = ""
    pagina_ficha = 0

    def afterPage(self):
        self.pagina_ficha += 1
        c = self.canv
        c.saveState()
        c.setFont("Helvetica", 7.5)
        texto = f"Ficha {self.ficha_atual} - página {self.pagina_ficha}" if self.ficha_atual else f"Página {self.page}"
        c.drawRightString(A4[0] - PAGE_MARGIN, PAGE_MARGIN - 2 * mm, texto)
        c.restoreState()


def agrupar_itens(itens: list) -> list:
    """
    Agrupa os itens por material + posição, ignorando lote e tipo de estoque:
    um material com vários lotes aparece uma única vez na ficha.
    Mantém a ordem em que os itens vieram do SAP (ordenados por posição).
    """
    agrupados = {}
    for item in itens:
        material = str(item.get("material", "")).strip()
        posicao = str(item.get("posicao", "") or "-").strip()
        chave = (material, posicao)
        if chave not in agrupados:
            agrupados[chave] = {"material": material, "posicao": posicao}
    return list(agrupados.values())


def _ficha_flowables(f: dict) -> list:
    story = [_MarcaFicha(str(f["numero_ficha"]))]

    # Cabeçalho
    header = Table(
        [[
            [_p("FICHA DE INVENTÁRIO FÍSICO (MI31)", _st_title),
             _p("Leitura Obrigatória via Coletor / Auditoria Física", _st_sub)],
            [Paragraph(f"<b>Centro:</b> {escape(str(f['centro']))} | <b>Depósito:</b> {escape(str(f['deposito']))}", _st_meta),
             Paragraph(f"<b>Emissão:</b> {escape(str(f['data_criacao'] or ''))}", _st_meta)],
        ]],
        colWidths=[CONTENT_WIDTH * 0.6, CONTENT_WIDTH * 0.4],
    )
    header.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LINEBELOW", (0, 0), (-1, 0), 2, colors.black),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story += [header, Spacer(1, 6 * mm)]

    # Código de barras grande + número da ficha
    hero = Table(
        [[_p("DOCUMENTO DE INVENTÁRIO SAP", _st_label)],
         [_barcode(str(f["numero_ficha"]))],
         [_p(f["numero_ficha"], _st_giant)]],
        colWidths=[CONTENT_WIDTH],
    )
    hero.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1.5, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 10),
    ]))
    story += [hero, Spacer(1, 5 * mm)]

    # Range de endereços (apenas se definido)
    if f.get("endereco_inicio") or f.get("endereco_fim"):
        rng = Table(
            [[_p("ENDEREÇO INICIAL (DE)", _st_label), "", _p("ENDEREÇO FINAL (ATÉ)", _st_label)],
             [_p(f.get("endereco_inicio") or "INÍCIO", _st_range), _p("»", _st_range),
              _p(f.get("endereco_fim") or "FIM", _st_range)]],
            colWidths=[CONTENT_WIDTH * 0.44, CONTENT_WIDTH * 0.12, CONTENT_WIDTH * 0.44],
        )
        rng.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 1.5, colors.black),
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f1f5f9")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, 0), 7),
            ("BOTTOMPADDING", (0, -1), (-1, -1), 7),
        ]))
        story += [rng, Spacer(1, 5 * mm)]

    # Tabela de itens (cabeçalho repetido em cada página)
    itens = agrupar_itens(f.get("itens") or [])
    story.append(_p(f"LISTA DE ITENS PARA CONTAGEM ({len(itens)} REGISTROS)", _st_section))
    data = [[_p("#", _st_th), _p("CÓDIGO DO MATERIAL", _st_th),
             _p("POSIÇÃO NO DEPÓSITO", _st_th), _p("QTD. FÍSICA CONTADA", _st_th)]]
    if itens:
        for i, item in enumerate(itens, start=1):
            data.append([
                _p(i, _st_cell),
                _p(item.get("material", ""), _st_cell_bold),
                _p(item.get("posicao", ""), _st_cell),
                "",
            ])
    else:
        data.append([_p(f"Nenhum item individual listado. Consulte o coletor pelo código de barras da ficha {f['numero_ficha']}.", _st_cell), "", "", ""])

    col_w = [12 * mm, 62 * mm, 50 * mm]
    col_w.append(CONTENT_WIDTH - sum(col_w))
    items_table = Table(data, colWidths=col_w, repeatRows=1)
    style = [
        ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e2e8f0")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    if not itens:
        style.append(("SPAN", (0, 1), (-1, 1)))
    items_table.setStyle(TableStyle(style))
    story += [items_table, Spacer(1, 16 * mm)]

    # Assinaturas
    sign = Table(
        [[_p("Nome Legível / Assinatura do Auxiliar de Contagem", _st_sign), "",
          _p("Data / Hora de Conclusão e Visto da Liderança", _st_sign)]],
        colWidths=[CONTENT_WIDTH * 0.46, CONTENT_WIDTH * 0.08, CONTENT_WIDTH * 0.46],
    )
    sign.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (0, 0), 0.8, colors.black),
        ("LINEABOVE", (2, 0), (2, 0), 0.8, colors.black),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(KeepTogether(sign))
    return story


def gerar_pdf_fichas(fichas: list) -> bytes:
    """Gera um único PDF A4 com todas as fichas informadas (cada ficha inicia em nova página)."""
    buffer = io.BytesIO()
    doc = _FichasDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=PAGE_MARGIN,
        rightMargin=PAGE_MARGIN,
        topMargin=PAGE_MARGIN,
        bottomMargin=PAGE_MARGIN + 4 * mm,
        title="Fichas de Inventário MI31",
        author="Emissor de Fichas MI31",
    )

    story = []
    for idx, f in enumerate(fichas):
        if idx > 0:
            story.append(PageBreak())
        story += _ficha_flowables(f)

    doc.build(story)
    return buffer.getvalue()


# ==============================================================================
# IMPRESSORAS DO WINDOWS
# ==============================================================================

_print_lock = threading.Lock()


def impressao_disponivel() -> bool:
    return win32print is not None


def listar_impressoras() -> list:
    """Lista as impressoras locais e de rede instaladas no Windows."""
    if not win32print:
        return []
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    nomes = {p[2] for p in win32print.EnumPrinters(flags)}
    return sorted(nomes, key=str.lower)


def impressora_padrao() -> str:
    if not win32print:
        return ""
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def _criar_dc_a4(nome_impressora: str):
    """Cria o contexto de dispositivo da impressora forçando papel A4 em retrato."""
    devmode = None
    try:
        handle = win32print.OpenPrinter(nome_impressora)
        try:
            devmode = win32print.GetPrinter(handle, 2)["pDevMode"]
        finally:
            win32print.ClosePrinter(handle)
    except Exception:
        devmode = None

    if devmode is not None:
        devmode.PaperSize = win32con.DMPAPER_A4
        devmode.Orientation = win32con.DMORIENT_PORTRAIT
        devmode.Fields |= win32con.DM_PAPERSIZE | win32con.DM_ORIENTATION
        hdc_handle = win32gui.CreateDC("WINSPOOL", nome_impressora, devmode)
        return win32ui.CreateDCFromHandle(hdc_handle)

    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(nome_impressora)
    return hdc


def imprimir_pdf(pdf_bytes: bytes, nome_impressora: str, titulo: str = "Fichas MI31") -> int:
    """
    Envia o PDF diretamente para a impressora informada, sem diálogo.
    Retorna a quantidade de páginas enviadas.
    """
    if not win32print:
        raise RuntimeError("Impressão direta disponível apenas no Windows (pywin32).")
    if not nome_impressora:
        raise RuntimeError("Nenhuma impressora selecionada.")
    if nome_impressora not in listar_impressoras():
        raise RuntimeError(f"Impressora '{nome_impressora}' não encontrada no Windows.")

    import pypdfium2 as pdfium
    from PIL import ImageWin

    with _print_lock:
        hdc = _criar_dc_a4(nome_impressora)
        try:
            dpi = hdc.GetDeviceCaps(win32con.LOGPIXELSX) or 300
            phys_w = hdc.GetDeviceCaps(win32con.PHYSICALWIDTH)
            phys_h = hdc.GetDeviceCaps(win32con.PHYSICALHEIGHT)
            off_x = hdc.GetDeviceCaps(win32con.PHYSICALOFFSETX)
            off_y = hdc.GetDeviceCaps(win32con.PHYSICALOFFSETY)

            pdf = pdfium.PdfDocument(pdf_bytes)
            total = len(pdf)
            hdc.StartDoc(titulo)
            try:
                for i in range(total):
                    page = pdf[i]
                    img = page.render(scale=min(dpi, 600) / 72, grayscale=True).to_pil()
                    page.close()

                    # Ajusta à folha física mantendo a proporção e centralizando
                    ratio = min(phys_w / img.width, phys_h / img.height)
                    dw, dh = int(img.width * ratio), int(img.height * ratio)
                    x0 = (phys_w - dw) // 2 - off_x
                    y0 = (phys_h - dh) // 2 - off_y

                    hdc.StartPage()
                    ImageWin.Dib(img).draw(hdc.GetHandleOutput(), (x0, y0, x0 + dw, y0 + dh))
                    hdc.EndPage()
                hdc.EndDoc()
            except Exception:
                try:
                    hdc.AbortDoc()
                except Exception:
                    pass
                raise
            finally:
                pdf.close()
        finally:
            hdc.DeleteDC()
    return total
