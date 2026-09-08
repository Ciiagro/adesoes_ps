"""Diagnóstico rápido: mostra o que está salvo no banco pros arquivos de
uma solicitação específica — pra descobrir por que o ofício não aparece
pro admin.

Uso:
    python diagnosticar_arquivo_carreta.py 42

(troque 42 pelo número/ID da solicitação com problema — esse número
aparece na URL quando você abre o detalhe dela, tipo /admin/carreta/42)
"""
import sys

from database import SessionLocal
from models import SolicitacaoCarreta, ArquivoCarreta


def main():
    if len(sys.argv) < 2:
        print("Uso: python diagnosticar_arquivo_carreta.py <ID_DA_SOLICITACAO>")
        return

    solicitacao_id = int(sys.argv[1])
    db = SessionLocal()

    solicitacao = db.get(SolicitacaoCarreta, solicitacao_id)
    if not solicitacao:
        print(f"Não encontrei nenhuma solicitação com ID {solicitacao_id}.")
        return

    print(f"Solicitação #{solicitacao.id} — {solicitacao.evento} — {solicitacao.municipio_nome}")
    print(f"Status: {solicitacao.status}")
    print()

    arquivos = (
        db.query(ArquivoCarreta)
        .filter(ArquivoCarreta.solicitacao_id == solicitacao_id)
        .all()
    )

    if not arquivos:
        print("❌ NENHUM registro de arquivo encontrado no banco pra essa solicitação.")
        print("   Isso significa que o upload nunca chegou a ser salvo — o problema")
        print("   é na hora de CRIAR a solicitação, não na hora de exibir.")
        return

    print(f"Encontrei {len(arquivos)} arquivo(s) cadastrado(s):\n")
    for arq in arquivos:
        print(f"  ID: {arq.id}")
        print(f"  Nome: {arq.nome_arquivo}")
        print(f"  storage_path: {arq.storage_path!r}")
        print(f"  tem conteudo (bytea antigo)?: {bool(arq.conteudo)}")
        print(f"  drive_url: {arq.drive_url!r}")
        print(f"  tipo_mime: {arq.tipo_mime!r}")
        print(f"  enviado_em: {arq.enviado_em}")

        if not arq.storage_path and not arq.conteudo and not arq.drive_url:
            print("  ⚠️  PROBLEMA: nenhum dos três campos (storage_path/conteudo/drive_url)")
            print("      está preenchido — é por isso que não aparece o link na tela!")
        elif arq.storage_path:
            print("  Testando se consegue baixar do Storage...")
            try:
                import supabase_storage
                conteudo = supabase_storage.baixar_bytes_privado(arq.storage_path)
                print(f"  ✅ Baixou {len(conteudo)} bytes do Storage sem erro.")
            except Exception as erro:
                print(f"  ❌ ERRO ao baixar do Storage: {erro}")
        print()


if __name__ == "__main__":
    main()
