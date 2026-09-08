"""Migração ÚNICA: move os ofícios em PDF da Carreta que ainda estão
guardados em binário (bytea) direto na tabela `carreta.arquivos` pro
Supabase Storage (bucket privado — não público, porque o download continua
exigindo login de admin), e limpa o campo antigo depois de confirmar que o
upload deu certo.

Antes de rodar, crie o bucket PRIVADO no painel do Supabase:
  Storage > New bucket > nome "carreta-oficios" > deixe DESMARCADO "Public
  bucket" (diferente do bucket de fotos, que é público).

Rode uma vez:

    python migrar_oficios_para_storage.py

É seguro rodar mais de uma vez — arquivos que já têm storage_path
preenchido são pulados.
"""

from database import SessionLocal
from models import ArquivoCarreta
import supabase_storage


def main():
    db = SessionLocal()
    pendentes = (
        db.query(ArquivoCarreta)
        .filter(ArquivoCarreta.conteudo.isnot(None))
        .filter(ArquivoCarreta.storage_path.is_(None))
        .all()
    )

    if not pendentes:
        print("Nada pra migrar — todos os ofícios já estão no Storage (ou não há nenhum).")
        return

    print(f"Encontrei {len(pendentes)} ofício(s) ainda em binário no banco.")

    ok, falhas = 0, []
    for arquivo in pendentes:
        try:
            caminho = supabase_storage.upload_oficio_carreta(
                arquivo.solicitacao_id, arquivo.conteudo, arquivo.nome_arquivo
            )
            arquivo.storage_path = caminho
            arquivo.conteudo = None
            db.commit()
            ok += 1
            print(f"  ✓ {arquivo.nome_arquivo} (solicitação #{arquivo.solicitacao_id}) -> {caminho}")
        except Exception as erro:  # noqa: BLE001
            db.rollback()
            falhas.append((arquivo.nome_arquivo, str(erro)))
            print(f"  ✗ {arquivo.nome_arquivo}: {erro}")

    print(f"\nMigrados: {ok}/{len(pendentes)}")
    if falhas:
        print("Falharam (ficaram como estavam, pode rodar o script de novo depois):")
        for nome, erro in falhas:
            print(f"  - {nome}: {erro}")


if __name__ == "__main__":
    main()
