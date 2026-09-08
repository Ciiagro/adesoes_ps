"""Upload e leitura de arquivos no Supabase Storage.

Por quê: guardar fotos/PDFs como bytea direto nas tabelas do Postgres faz
qualquer SELECT que traga a linha (inclusive listagens que nem mostram a
imagem) arrastar o binário inteiro pela rede — é isso que estava estourando
a cota de egress do Supabase. O Storage serve os arquivos por CDN, com cache
de navegador, e sai de uma cota separada da do banco.

Precisa de duas variáveis no .env (Project Settings > API, no painel do
Supabase):

    SUPABASE_URL=https://xfqtjuwovwvdfvbmiuwn.supabase.co
    SUPABASE_SERVICE_KEY=eyJ...   # a chave "service_role" (secreta!)

E de um bucket criado no painel (Storage > New bucket), marcado como
"Public" — por padrão este código usa o nome "sindicatos-fotos".
"""

import os
import mimetypes
import re
import unicodedata
import uuid

import requests

SUPABASE_URL = os.environ.get("SUPABASE_URL", "").rstrip("/")
SUPABASE_SERVICE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", "")
BUCKET_FOTOS_PRESIDENTES = os.environ.get("SUPABASE_BUCKET_FOTOS", "sindicatos-fotos")
BUCKET_OFICIOS_CARRETA = os.environ.get("SUPABASE_BUCKET_OFICIOS", "carreta-oficios")


class SupabaseStorageError(RuntimeError):
    pass


def _checar_config():
    if not SUPABASE_URL or not SUPABASE_SERVICE_KEY:
        raise SupabaseStorageError(
            "SUPABASE_URL / SUPABASE_SERVICE_KEY não configurados no .env "
            "— sem eles não dá pra falar com o Storage."
        )


def _extensao(tipo_mime, nome_original=""):
    ext = mimetypes.guess_extension(tipo_mime or "") or ""
    if not ext and "." in (nome_original or ""):
        ext = "." + nome_original.rsplit(".", 1)[-1]
    return ext or ".jpg"


def _sanitizar_nome_arquivo(nome):
    """O Supabase Storage rejeita certos caracteres na chave do objeto
    (acentos como 'ó', espaços, etc — dá erro 'InvalidKey'). Tira os
    acentos (ó -> o) e troca qualquer coisa que não seja letra/número/
    ponto/traço/underline por "_", pra virar uma chave sempre aceita."""
    nome = nome or "arquivo"
    sem_acento = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Za-z0-9._-]", "_", sem_acento) or "arquivo"


def upload_foto_presidente(sindicato_id, conteudo: bytes, tipo_mime: str, nome_original: str = "") -> str:
    """Sobe a foto do presidente de um sindicato pro Storage e devolve a
    URL pública. Cada upload usa um nome novo (com um sufixo aleatório) pra
    evitar que o cache do navegador/CDN sirva a foto antiga depois de trocar."""
    _checar_config()

    ext = _extensao(tipo_mime, nome_original)
    caminho = f"presidentes/{sindicato_id}-{uuid.uuid4().hex[:8]}{ext}"

    resposta = requests.put(
        f"{SUPABASE_URL}/storage/v1/object/{BUCKET_FOTOS_PRESIDENTES}/{caminho}",
        headers={
            "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
            "apikey": SUPABASE_SERVICE_KEY,
            "Content-Type": tipo_mime or "application/octet-stream",
            "x-upsert": "true",
            "Cache-Control": "public, max-age=31536000, immutable",
        },
        data=conteudo,
        timeout=30,
    )
    if resposta.status_code not in (200, 201):
        raise SupabaseStorageError(
            f"Falha ao enviar foto pro Storage ({resposta.status_code}): {resposta.text[:300]}"
        )

    return f"{SUPABASE_URL}/storage/v1/object/public/{BUCKET_FOTOS_PRESIDENTES}/{caminho}"


def excluir_arquivo_por_url(url: str):
    """Apaga do Storage o arquivo apontado por uma URL pública gerada por
    upload_foto_presidente (usa isso pra não acumular fotos antigas órfãs).
    Falha em silêncio (só avisa) — não vale travar o fluxo principal por
    causa de limpeza (nem se o Storage não estiver configurado, nem se
    der erro de rede)."""
    if not url or f"/storage/v1/object/public/{BUCKET_FOTOS_PRESIDENTES}/" not in url:
        return
    caminho = url.split(f"/storage/v1/object/public/{BUCKET_FOTOS_PRESIDENTES}/", 1)[1]
    try:
        _checar_config()
        requests.delete(
            f"{SUPABASE_URL}/storage/v1/object/{BUCKET_FOTOS_PRESIDENTES}/{caminho}",
            headers={
                "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
                "apikey": SUPABASE_SERVICE_KEY,
            },
            timeout=15,
        )
    except (requests.RequestException, SupabaseStorageError) as erro:
        print(f"[supabase_storage] aviso: não consegui apagar '{caminho}': {erro}")


def baixar_bytes(url: str) -> bytes:
    """Baixa o conteúdo de uma URL pública do Storage. Usado só no lugar que
    realmente precisa dos bytes crus (montar o PDF da ficha do sindicato) —
    aí sim vale a pena buscar o arquivo inteiro, é uma foto de cada vez."""
    resposta = requests.get(url, timeout=30)
    resposta.raise_for_status()
    return resposta.content


def upload_oficio_carreta(solicitacao_id, conteudo: bytes, nome_arquivo: str) -> str:
    """Sobe um ofício (PDF) anexado a uma solicitação de carreta pro bucket
    PRIVADO do Storage e devolve o CAMINHO interno (não uma URL pública —
    esse bucket não é público, porque o download hoje é restrito a quem
    está logado no admin; queremos manter esse controle de acesso, só
    tirando o peso do bytea de dentro do Postgres)."""
    _checar_config()
    caminho = f"solicitacoes/{solicitacao_id}/{uuid.uuid4().hex[:8]}-{_sanitizar_nome_arquivo(nome_arquivo)}"

    resposta = requests.put(
        f"{SUPABASE_URL}/storage/v1/object/{BUCKET_OFICIOS_CARRETA}/{caminho}",
        headers={
            "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
            "apikey": SUPABASE_SERVICE_KEY,
            "Content-Type": "application/pdf",
            "x-upsert": "true",
        },
        data=conteudo,
        timeout=30,
    )
    if resposta.status_code not in (200, 201):
        raise SupabaseStorageError(
            f"Falha ao enviar ofício pro Storage ({resposta.status_code}): {resposta.text[:300]}"
        )
    return caminho


def baixar_bytes_privado(caminho: str, bucket: str = BUCKET_OFICIOS_CARRETA) -> bytes:
    """Baixa os bytes de um arquivo de um bucket PRIVADO, usando a
    service_role key (que tem acesso mesmo sem o bucket ser público). Quem
    chama essa função já deve ter conferido que o usuário tem permissão de
    ver o arquivo (ex: checar login antes) — aqui não tem RLS nenhuma
    protegendo, a responsabilidade de checar acesso é de quem chama."""
    _checar_config()
    resposta = requests.get(
        f"{SUPABASE_URL}/storage/v1/object/{bucket}/{caminho}",
        headers={
            "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
            "apikey": SUPABASE_SERVICE_KEY,
        },
        timeout=30,
    )
    resposta.raise_for_status()
    return resposta.content


def excluir_arquivo_privado(caminho: str, bucket: str = BUCKET_OFICIOS_CARRETA):
    """Apaga um arquivo do bucket privado (ex: se a solicitação inteira for
    excluída). Falha em silêncio — não vale travar a exclusão da
    solicitação por causa de limpeza do Storage (nem se o Storage não
    estiver configurado, nem se der erro de rede)."""
    if not caminho:
        return
    try:
        _checar_config()
        requests.delete(
            f"{SUPABASE_URL}/storage/v1/object/{bucket}/{caminho}",
            headers={
                "Authorization": f"Bearer {SUPABASE_SERVICE_KEY}",
                "apikey": SUPABASE_SERVICE_KEY,
            },
            timeout=15,
        )
    except (requests.RequestException, SupabaseStorageError) as erro:
        print(f"[supabase_storage] aviso: não consegui apagar '{caminho}' do bucket '{bucket}': {erro}")
