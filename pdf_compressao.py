"""Compressão automática de PDFs grandes (tipicamente ofícios escaneados
em resolução exagerada) antes de subir pro Supabase Storage.

Por quê: um ofício de 1 página, escaneado sem cuidado no celular ou numa
impressora multifuncional, facilmente passa de 15-20 MB — não porque o
documento precise disso, mas porque a resolução do scan é maior do que
qualquer tela ou impressão vai precisar. Isso pesa no Storage à toa.

Como funciona: só entra em ação em arquivos GRANDES (acima de
LIMIAR_PARA_COMPRIMIR_BYTES) — um PDF pequeno quase sempre é um documento
"nascido digital" (word/editor exportando pra PDF), com texto selecionável
de verdade, e mexer nele só perderia qualidade à toa. Já um PDF grande, na
prática, é quase sempre um scan de imagem — nesse caso, rasteriza cada
página numa resolução mais razoável (150 DPI, de sobra pra leitura em tela
ou impressão) e recomprime como JPEG.

Nunca arrisca o upload: se a compressão falhar por qualquer motivo, ou se
o resultado não ficar menor que o original, devolve o arquivo original."""

import pymupdf as fitz  # PyMuPDF; "fitz" é o nome antigo do pacote, descontinuado

LIMIAR_PARA_COMPRIMIR_BYTES = 2 * 1024 * 1024  # só comprime PDFs acima de 2 MB
DPI_SAIDA = 150
QUALIDADE_JPEG = 60


def comprimir_pdf_se_necessario(conteudo: bytes) -> bytes:
    """Recebe os bytes de um PDF. Se for pequeno, devolve sem mexer. Se for
    grande, tenta comprimir e devolve o resultado só se ele realmente
    ficou menor — nunca devolve algo pior que o original, nunca levanta
    exceção (qualquer erro aqui cai pro arquivo original, sem travar o
    upload de quem está preenchendo o formulário)."""
    if len(conteudo) <= LIMIAR_PARA_COMPRIMIR_BYTES:
        return conteudo

    try:
        origem = fitz.open(stream=conteudo, filetype="pdf")
        destino = fitz.open()
        zoom = DPI_SAIDA / 72
        matriz = fitz.Matrix(zoom, zoom)

        for pagina in origem:
            pixmap = pagina.get_pixmap(matrix=matriz, colorspace=fitz.csRGB)
            jpeg_bytes = pixmap.tobytes("jpeg", jpg_quality=QUALIDADE_JPEG)
            nova_pagina = destino.new_page(width=pagina.rect.width, height=pagina.rect.height)
            nova_pagina.insert_image(nova_pagina.rect, stream=jpeg_bytes)

        resultado = destino.tobytes(deflate=True, garbage=4)
        origem.close()
        destino.close()

        if 0 < len(resultado) < len(conteudo):
            return resultado
        return conteudo
    except Exception as erro:  # noqa: BLE001 — nunca trava o upload por causa da compressão
        print(f"[pdf_compressao] aviso: não consegui comprimir, mandando o original: {erro}")
        return conteudo
