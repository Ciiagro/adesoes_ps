"""Gera a Ficha de Adesão em PDF usando reportlab (Platypus).

Ao contrário do weasyprint (HTML/CSS -> PDF), o reportlab não depende de
bibliotecas gráficas do sistema operacional (Pango/Cairo/GTK) — é só Python
puro. Isso é obrigatório para rodar em ambientes serverless como o Vercel,
onde não dá pra instalar pacotes de sistema.
"""

import io
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    HRFlowable, KeepTogether, PageBreak, Paragraph, Spacer, Table, TableStyle,
    SimpleDocTemplate,
)

LARGURA_PAGINA, ALTURA_PAGINA = A4
MARGEM = 16 * mm
LARGURA_UTIL = LARGURA_PAGINA - 2 * MARGEM

COR_LINHA = colors.HexColor("#D8D0B8")
COR_FUNDO_SECAO = colors.HexColor("#F1EFE3")
COR_FUNDO_TOTAL = colors.HexColor("#FBFAF3")
COR_ROTULO = colors.HexColor("#6E6555")
COR_TEXTO = colors.HexColor("#2A2620")
COR_RODAPE = colors.HexColor("#8A8168")


def _esc(valor):
    """Converte para string, escapa XML e troca vazio/None por '-'."""
    if valor is None or valor == "":
        return "-"
    return escape(str(valor))


def _estilos():
    base = getSampleStyleSheet()
    e = {}
    e["subtitulo"] = ParagraphStyle(
        "subtitulo", parent=base["Normal"], fontSize=9, textColor=COR_ROTULO,
    )
    e["campo"] = ParagraphStyle(
        "campo", parent=base["Normal"], fontSize=8.3, leading=15, textColor=COR_TEXTO,
    )
    e["celula_tabela"] = ParagraphStyle(
        "celula_tabela", parent=base["Normal"], fontSize=9.4, alignment=TA_CENTER,
        textColor=COR_TEXTO, leading=12,
    )
    e["celula_tabela_total"] = ParagraphStyle(
        "celula_tabela_total", parent=e["celula_tabela"], fontName="Helvetica-Bold",
    )
    e["cabecalho_tabela"] = ParagraphStyle(
        "cabecalho_tabela", parent=e["celula_tabela"], fontName="Helvetica-Bold",
    )
    e["escola_titulo"] = ParagraphStyle(
        "escola_titulo", parent=base["Normal"], fontSize=10.8, leading=13,
        fontName="Helvetica-Bold",
    )
    e["escola_meta"] = ParagraphStyle(
        "escola_meta", parent=base["Normal"], fontSize=9.2, leading=13, textColor=COR_TEXTO,
    )
    e["termo"] = ParagraphStyle(
        "termo", parent=base["Normal"], fontSize=9.6, leading=14,
    )
    e["subsecao"] = ParagraphStyle(
        "subsecao", parent=base["Normal"], fontSize=9.3, textColor=COR_ROTULO,
        fontName="Helvetica-Bold", spaceBefore=6, spaceAfter=4,
    )
    e["assinatura_nome"] = ParagraphStyle(
        "assinatura_nome", parent=base["Normal"], fontSize=9.8, leading=12,
        fontName="Helvetica-Bold",
    )
    e["assinatura_cargo"] = ParagraphStyle(
        "assinatura_cargo", parent=base["Normal"], fontSize=8.6, textColor=COR_ROTULO,
        leading=11,
    )
    e["rodape"] = ParagraphStyle(
        "rodape", parent=base["Normal"], fontSize=8, textColor=COR_RODAPE,
    )
    return e


def _secao_header(texto, cor_tema, cor_tema_escura):
    """Uma faixa de título de seção com fundo claro e barra colorida à esquerda."""
    estilo = ParagraphStyle(
        "secao", fontSize=11.5, leading=14, textColor=cor_tema_escura,
        fontName="Helvetica-Bold",
    )
    tabela = Table([[Paragraph(texto, estilo)]], colWidths=[LARGURA_UTIL])
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), COR_FUNDO_SECAO),
        ("LINEBEFORE", (0, 0), (0, -1), 4, cor_tema),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    return tabela


def _campo_paragrafo(estilos, rotulo, valor):
    texto = (
        f"<font size=7.5 color='#6E6555'>{escape(rotulo.upper())}</font>"
        f"<br/><font size=10.3 color='#2A2620'><b>{_esc(valor)}</b></font>"
    )
    return Paragraph(texto, estilos["campo"])


def _grid_campos(estilos, pares):
    """Recebe lista de (rotulo, valor) e monta uma grade de 2 colunas."""
    linhas = []
    for i in range(0, len(pares), 2):
        esquerda = _campo_paragrafo(estilos, *pares[i])
        direita = (
            _campo_paragrafo(estilos, *pares[i + 1]) if i + 1 < len(pares) else ""
        )
        linhas.append([esquerda, direita])
    tabela = Table(linhas, colWidths=[LARGURA_UTIL / 2, LARGURA_UTIL / 2])
    tabela.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 22),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return tabela


def _tabela_series(estilos, colunas, valores, cor_tema):
    cabecalho = [Paragraph(escape(c), estilos["cabecalho_tabela"]) for c in colunas] + [
        Paragraph("Total", estilos["cabecalho_tabela"])
    ]
    total = sum(valores)
    linha = [Paragraph(str(v), estilos["celula_tabela"]) for v in valores] + [
        Paragraph(str(total), estilos["celula_tabela_total"])
    ]
    largura_col = LARGURA_UTIL / (len(colunas) + 1)
    tabela = Table([cabecalho, linha], colWidths=[largura_col] * (len(colunas) + 1))
    tabela.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.6, COR_LINHA),
        ("BACKGROUND", (0, 0), (-1, 0), COR_FUNDO_SECAO),
        ("BACKGROUND", (-1, 1), (-1, 1), COR_FUNDO_TOTAL),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return tabela


def _bloco_assinatura(estilos, nome, cargo, selo_html=None):
    largura = LARGURA_UTIL / 2 - 12
    elementos = [
        Spacer(1, 22),
        HRFlowable(width=largura, thickness=0.8, color=COR_TEXTO, spaceAfter=4),
        Paragraph(_esc(nome) if nome else "____________________________", estilos["assinatura_nome"]),
        Paragraph(cargo, estilos["assinatura_cargo"]),
    ]
    if selo_html:
        estilo_selo = ParagraphStyle(
            "selo", parent=estilos["assinatura_cargo"], fontSize=8.2,
            backColor=colors.HexColor("#EFF6F1"), borderColor=colors.HexColor("#3F6F4C"),
            borderWidth=0.6, borderPadding=4, spaceBefore=4,
        )
        elementos.append(Paragraph(selo_html, estilo_selo))
    return elementos


def _linha_dupla(bloco_a, bloco_b):
    largura = LARGURA_UTIL / 2
    tabela = Table([[bloco_a, bloco_b]], colWidths=[largura, largura])
    tabela.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 20),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    return tabela


def gerar_pdf_adesao_bytes(adesao, series, totais_series, gerado_em, cor_tema,
                            cor_tema_escura, assinaturas_por_papel, nome_programa):
    """Monta a Ficha de Adesão inteira (seções 1 a 6) e devolve os bytes do PDF."""
    e = _estilos()
    m = adesao.municipio
    c = adesao.coordenador
    story = []

    # --- Cabeçalho ---
    estilo_selo_topo = ParagraphStyle(
        "selo_topo", fontSize=8.5, textColor=colors.white, fontName="Helvetica-Bold",
        backColor=cor_tema, borderPadding=(3, 10, 3, 10),
    )
    selo_texto = "SENAR CEARÁ" if adesao.programa.nome == "AGRINHO" else "SEDUC / PARCEIROS"
    story.append(Paragraph(selo_texto, estilo_selo_topo))
    story.append(Spacer(1, 8))
    estilo_titulo = ParagraphStyle(
        "titulo", fontSize=17, leading=20, fontName="Helvetica-Bold", textColor=cor_tema_escura,
    )
    story.append(Paragraph(f"Ficha de Adesão &middot; {nome_programa} {adesao.ano}", estilo_titulo))
    story.append(Paragraph(
        "Documento gerado pelo Sistema de Adesões &middot; Dados preenchidos eletronicamente pelo município",
        e["subtitulo"],
    ))
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width=LARGURA_UTIL, thickness=2.2, color=cor_tema))
    story.append(Spacer(1, 12))

    # --- 1. Município ---
    story.append(_secao_header("1. Identificação do Município e Executivo", cor_tema, cor_tema_escura))
    story.append(Spacer(1, 6))
    story.append(_grid_campos(e, [
        ("Município", f"{m.nome} - {m.uf}"),
        ("Prefeito(a)", m.prefeito_nome),
        ("RG do Prefeito(a)", m.prefeito_rg),
        ("CPF do Prefeito(a)", m.prefeito_cpf),
        ("Endereço da Prefeitura", m.prefeitura_endereco),
        ("CEP", m.prefeitura_cep),
        ("Telefone", m.prefeitura_telefone),
        ("E-mail", m.prefeitura_email),
    ]))
    story.append(Spacer(1, 10))

    # --- 2. Secretaria ---
    story.append(_secao_header("2. Secretaria de Educação", cor_tema, cor_tema_escura))
    story.append(Spacer(1, 6))
    story.append(_grid_campos(e, [
        ("Secretário(a)", m.secretario_nome),
        ("CPF", m.secretaria_cep),
        ("Endereço", m.secretaria_endereco),
        ("Contato", m.secretaria_telefone),
        ("E-mail", m.secretaria_email),
    ]))
    story.append(Spacer(1, 10))

    # --- 3. Coordenador ---
    story.append(_secao_header("3. Coordenador(a) do Projeto", cor_tema, cor_tema_escura))
    story.append(Spacer(1, 6))
    telefones = ", ".join(filter(None, [
        c.telefone1 if c else None, c.telefone2 if c else None,
    ]))
    story.append(_grid_campos(e, [
        ("Nome", c.nome if c else None),
        ("E-mail", c.email if c else None),
        ("RG", c.rg if c else None),
        ("CPF", c.cpf if c else None),
        ("Telefone(s)", telefones or None),
    ]))
    story.append(Spacer(1, 10))

    # --- 4. Censo ---
    bloco_4 = [
        _secao_header("4. Censo da Rede Municipal", cor_tema, cor_tema_escura),
        Spacer(1, 6),
        _grid_campos(e, [
            ("Total de Escolas (informado)", adesao.total_escolas_informado),
            ("Total de Professores (informado)", adesao.total_professores_informado),
        ]),
        Spacer(1, 8),
        Paragraph(
            "<b>Quantidade de alunos matriculados por série (soma de todas as escolas aderidas)</b>",
            e["subsecao"],
        ),
        _tabela_series(
            e, [s.nome for s in series],
            [totais_series.get(s.id, 0) for s in series],
            cor_tema,
        ),
    ]
    story.append(KeepTogether(bloco_4))
    story.append(Spacer(1, 4))

    # --- 5. Escolas ---
    story.append(PageBreak())
    story.append(_secao_header("5. Escolas Participantes e Matrículas", cor_tema, cor_tema_escura))
    story.append(Spacer(1, 10))

    if not adesao.escolas:
        story.append(Paragraph("Nenhuma escola cadastrada nesta adesão.", e["termo"]))
    for i, ae in enumerate(adesao.escolas, start=1):
        esc_obj = ae.escola
        meta1 = (
            f"<b>Endereço:</b> {_esc(esc_obj.endereco)} &nbsp;&nbsp; "
            f"<b>Telefone:</b> {_esc(esc_obj.contato_telefone)} &nbsp;&nbsp; "
            f"<b>E-mail:</b> {_esc(esc_obj.contato_email)}"
        )
        meta2 = f"<b>Professores:</b> {ae.quantidade_professores or 0} &nbsp;&nbsp; <b>Gestor(a) 1:</b> {_esc(esc_obj.gestor1_nome)}"
        if esc_obj.gestor2_nome:
            meta2 += f" &nbsp;&nbsp; <b>Gestor(a) 2:</b> {_esc(esc_obj.gestor2_nome)}"

        nomes_series = [m_.serie.nome for m_ in ae.matriculas]
        valores_series = [m_.quantidade for m_ in ae.matriculas]

        bloco_escola = [
            Paragraph(f"{i}. {_esc(esc_obj.nome)}", e["escola_titulo"]),
            Spacer(1, 3),
            Paragraph(meta1, e["escola_meta"]),
            Paragraph(meta2, e["escola_meta"]),
            Spacer(1, 6),
        ]
        if nomes_series:
            bloco_escola.append(_tabela_series(e, nomes_series, valores_series, cor_tema))
        bloco_escola.append(Spacer(1, 14))
        story.append(KeepTogether(bloco_escola))

    # --- 6. Termo e assinaturas ---
    story.append(PageBreak())
    story.append(_secao_header("6. Termo de Adesão e Compromisso", cor_tema, cor_tema_escura))
    story.append(Spacer(1, 8))
    termo_texto = (
        f"Pelo presente termo, o município de <b>{_esc(m.nome)}</b>, por meio de seus "
        f"representantes abaixo assinados, formaliza sua adesão ao <b>{escape(nome_programa)}</b> "
        f"para o ano de <b>{adesao.ano}</b>, comprometendo-se a viabilizar a execução do "
        f"programa junto às escolas listadas neste documento, garantindo a participação de "
        f"gestores, professores e alunos nas atividades previstas, bem como o correto envio "
        f"das informações de acompanhamento solicitadas pela coordenação do projeto."
    )
    estilo_termo_caixa = ParagraphStyle(
        "termo_caixa", parent=e["termo"], backColor=COR_FUNDO_TOTAL,
        borderColor=COR_LINHA, borderWidth=0.8, borderPadding=10,
    )
    story.append(Paragraph(termo_texto, estilo_termo_caixa))
    story.append(Spacer(1, 14))

    def selo(papel):
        a = assinaturas_por_papel.get(papel)
        if a and a.assinado_em:
            return (
                f"✓ Assinado eletronicamente em {a.assinado_em.strftime('%d/%m/%Y %H:%M')} "
                f"— Protocolo {a.token[:8].upper()}"
            )
        return None

    story.append(Paragraph("<b>Representantes do Município</b>", e["subsecao"]))
    story.append(_linha_dupla(
        _bloco_assinatura(e, m.prefeito_nome, f"Prefeito(a) Municipal — CPF {_esc(m.prefeito_cpf)}", selo("prefeito")),
        _bloco_assinatura(e, m.secretario_nome, "Secretário(a) de Educação", selo("secretario")),
    ))
    story.append(Spacer(1, 6))
    story.append(_linha_dupla(
        _bloco_assinatura(e, c.nome if c else None, "Coordenador(a) do Projeto", selo("coordenador")),
        _bloco_assinatura(e, adesao.responsavel_preenchimento, "Responsável pelo Preenchimento"),
    ))
    story.append(Spacer(1, 16))

    story.append(Paragraph("<b>Assinaturas dos Membros de cada Escola Participante</b>", e["subsecao"]))
    for ae in adesao.escolas:
        esc_obj = ae.escola
        bloco = [
            Paragraph(_esc(esc_obj.nome), e["escola_titulo"]),
            _linha_dupla(
                _bloco_assinatura(e, esc_obj.gestor1_nome, "Gestor(a) Escolar 1"),
                _bloco_assinatura(e, esc_obj.gestor2_nome, "Gestor(a) Escolar 2"),
            ),
            Spacer(1, 12),
        ]
        story.append(KeepTogether(bloco))

    story.append(Spacer(1, 10))
    story.append(HRFlowable(width=LARGURA_UTIL, thickness=0.6, color=COR_LINHA))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        f"Gerado automaticamente pelo Sistema de Adesões em {gerado_em.strftime('%d/%m/%Y às %H:%M')}. "
        f"Responsável pelo preenchimento: {_esc(adesao.responsavel_preenchimento)}.",
        e["rodape"],
    ))

    def _rodape_pagina(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(COR_RODAPE)
        texto = f"Sistema de Adesões · {nome_programa} — página {doc.page}"
        canvas.drawCentredString(LARGURA_PAGINA / 2, 10 * mm, texto)
        canvas.restoreState()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        topMargin=MARGEM, bottomMargin=MARGEM + 4 * mm,
        leftMargin=MARGEM, rightMargin=MARGEM,
        title=f"Ficha de Adesão - {nome_programa} {adesao.ano} - {m.nome}",
    )
    doc.build(story, onFirstPage=_rodape_pagina, onLaterPages=_rodape_pagina)
    return buffer.getvalue()
