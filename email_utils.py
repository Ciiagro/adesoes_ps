import os
import smtplib
import ssl
from email.message import EmailMessage


def enviar_email(destinatario, assunto, corpo_html, usuario=None, senha=None, servidor=None, porta=None, anexos=None):
    """Envia um e-mail simples em HTML. Retorna True se enviou, False se falhou
    (por exemplo, se as credenciais não estiverem configuradas). Nunca
    levanta exceção — quem chamar deve sempre ter um plano B (mostrar o
    link na tela), pois o envio de e-mail pode falhar por vários motivos.

    Por padrão usa MAIL_USERNAME/MAIL_PASSWORD (a conta principal do
    sistema). Passe `usuario`/`senha` pra usar uma conta diferente — por
    exemplo, um módulo que precisa mandar e-mail de um endereço próprio.

    `anexos`, se informado, é uma lista de dicts {"nome", "conteudo" (bytes),
    "tipo_mime"} — cada um vira um arquivo anexado ao e-mail."""
    usuario = usuario or os.environ.get("MAIL_USERNAME")
    senha = senha or os.environ.get("MAIL_PASSWORD")
    servidor = servidor or os.environ.get("MAIL_SERVER", "smtp.gmail.com")
    porta = porta or int(os.environ.get("MAIL_PORT", "587"))

    if not destinatario:
        print("[email_utils] Não enviei: nenhum destinatário informado.")
        return False
    if not usuario or not senha:
        print(
            f"[email_utils] Não enviei pra {destinatario}: credenciais de e-mail não configuradas "
            f"(usuário {'ok' if usuario else 'FALTANDO'}, senha {'ok' if senha else 'FALTANDO'})."
        )
        return False

    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = usuario
    msg["To"] = destinatario
    msg.set_content("Este e-mail contém conteúdo em HTML. Abra em um cliente compatível.")
    msg.add_alternative(corpo_html, subtype="html")

    for anexo in anexos or []:
        tipo_mime = anexo.get("tipo_mime") or "application/octet-stream"
        maintype, _, subtype = tipo_mime.partition("/")
        msg.add_attachment(
            anexo["conteudo"],
            maintype=maintype or "application",
            subtype=subtype or "octet-stream",
            filename=anexo.get("nome") or "anexo",
        )

    try:
        contexto = ssl.create_default_context()
        with smtplib.SMTP(servidor, porta, timeout=15) as smtp:
            smtp.starttls(context=contexto)
            smtp.login(usuario, senha)
            smtp.send_message(msg)
        return True
    except Exception as e:  # nunca deixa o fluxo do usuário quebrar por causa do e-mail
        print(f"[email_utils] Falha ao enviar e-mail para {destinatario}: {e}")
        return False
