# Sistema de Adesões - Agrinho / Projeto Valores

Sistema Flask para os municípios preencherem, todo ano, a adesão aos programas
**Agrinho** (2º ao 9º ano) e **Projeto Valores** (Infantil 3, 4 e 5), com
cadastro de escolas e matrículas por série. Cada ano é registrado
separadamente, então escolas podem entrar/sair de um ano para o outro sem
perder o histórico.

## ⚠️ Segurança - leia antes de tudo

A senha do banco compartilhada no chat (`U4Wabi1pnGOCX3Z4`) deve ser
considerada exposta. Troque-a agora em:
**Supabase → Project Settings → Database → Reset database password**,
e só depois coloque a nova senha no seu `.env` (que nunca deve ir para o
GitHub — adicione `.env` ao `.gitignore`).

## Nota sobre o driver do banco

O projeto usa **pg8000** (driver Python puro) em vez do `psycopg2-binary`.
Isso evita o erro clássico do Windows `UnicodeDecodeError ... psycopg2-binary`,
que acontece quando o pip não acha um wheel pronto e tenta compilar do zero.
Com pg8000 não há compilação nenhuma — funciona igual em Windows, Mac e Linux.
Seu `.env` continua normal, com `DATABASE_URL=postgresql://...` (a troca para
`postgresql+pg8000://` é feita automaticamente pelo `database.py`).

## 1. Instalação

```bash
python3 -m venv venv
source venv/bin/activate          # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Configurar o `.env`

Copie `.env.example` para `.env` e preencha com seus dados reais
(senha nova do banco, senha de admin, etc.):

```bash
cp .env.example .env
```

## 3. Criar o schema no Supabase (uma vez só)

Abra o **SQL Editor** do seu projeto Supabase e rode o conteúdo do arquivo
`schema.sql` deste projeto. Ele cria o schema `educacao`, todas as tabelas,
e já cadastra os programas (Agrinho/Valores) e as séries de cada um.

## 4. Rodar o sistema

```bash
python app.py
```

Acesse:
- `http://localhost:5000/` — página inicial
- `http://localhost:5000/adesao/agrinho` — formulário do Agrinho
- `http://localhost:5000/adesao/valores` — formulário do Projeto Valores
- `http://localhost:5000/admin` — painel administrativo (pede `ADMIN_PASSWORD`
  ou `COMISSAO_PASSWORD` do `.env`)

## Como o banco está organizado (schema `educacao`)

| Tabela            | O que guarda                                                        |
|-------------------|----------------------------------------------------------------------|
| `programas`       | Agrinho / Valores (fixo)                                             |
| `series`          | As séries de cada programa (2º–9º ano / Infantil 3–5)                |
| `municipios`      | Dados do prefeito e da secretaria (cadastro único, reaproveitado)     |
| `coordenadores`   | Coordenador(a) do projeto informado em cada envio                    |
| `adesoes`         | **Uma linha por ano** de cada programa/município                     |
| `escolas`         | Cadastro de escolas (reaproveitado ano a ano)                        |
| `adesao_escolas`  | Liga uma escola a uma adesão (ano) específica — é aqui que entram as escolas novas de cada ano |
| `matriculas`      | Quantidade de alunos por série, por escola, dentro de cada adesão     |

Se em 2027 um município adere de novo, um novo registro é criado em
`adesoes` (ano=2027) e as escolas daquele ano são ligadas via
`adesao_escolas` — as escolas de 2026 continuam intactas no histórico.

## Rascunho e Assinaturas Eletrônicas

A ficha é longa e o Prefeito, o(a) Secretário(a) de Educação e o(a) Coordenador(a)
do Projeto são pessoas diferentes — cada uma com seu próprio e-mail. Por isso o
formulário tem dois botões no final:

- **💾 Salvar rascunho e continuar depois** — grava tudo o que já foi preenchido e
  gera um link único (`/adesao/<programa>?continuar=TOKEN`). Esse link é mostrado
  na tela e também enviado por e-mail para o(a) coordenador(a), se houver e-mail
  cadastrado. Quem abrir o link não precisa fazer login — o formulário volta
  pré-preenchido com tudo que já foi salvo (município, escolas, matrículas etc.),
  pronto para continuar de onde parou.

- **Enviar Adesão** — finaliza o preenchimento e dispara automaticamente um
  e-mail para **cada um dos três responsáveis** (prefeito, secretário, coordenador)
  com um link individual (`/assinar/TOKEN`). Nessa página, cada um vê só os dados
  da própria função, confirma que revisou e "assina eletronicamente" (nome digitado
  + data/hora + IP ficam registrados) — sem precisar acessar o sistema. Quando os
  três confirmam, a adesão passa para o status **concluída**.

O status de cada adesão (`rascunho` / `aguardando_assinaturas` / `concluida`) e o
andamento das três assinaturas aparecem no painel `/admin/adesao/<id>`, de onde dá
pra reenviar o e-mail de um responsável específico caso ele tenha perdido o link.

A **Ficha de Adesão em PDF** reflete isso automaticamente: o responsável que já
assinou eletronicamente aparece com um selo "✓ Assinado eletronicamente em
DD/MM/AAAA HH:MM — Protocolo XXXXXXXX" no lugar da assinatura física.

⚠️ **Envio de e-mails:** depende de `MAIL_USERNAME`/`MAIL_PASSWORD` no `.env`
(veja `.env.example`). Se o Gmail for usado, `MAIL_PASSWORD` precisa ser uma
["senha de app"](https://myaccount.google.com/apppasswords), não a senha normal
da conta. Se o e-mail falhar por qualquer motivo, o link aparece na tela mesmo
assim, para copiar e enviar manualmente.

## Publicando no Vercel

O projeto já vem pronto pra rodar no Vercel: os arquivos `vercel.json` e
`api/index.py` fazem a ponte entre o Flask (que continua em `app.py`, sem
nenhuma mudança de estrutura) e o formato que o Vercel espera.

1. Suba este projeto pro GitHub (repositório normal, com estes arquivos na raiz).
2. No [vercel.com](https://vercel.com), clique em **Add New → Project** e
   importe o repositório.
3. Em **Environment Variables**, adicione as mesmas variáveis do seu `.env`:
   `DATABASE_URL`, `SECRET_KEY`, `ADMIN_PASSWORD`, `COMISSAO_PASSWORD`,
   `MAIL_USERNAME`, `MAIL_PASSWORD`.
4. Clique em **Deploy**.

⚠️ **Use a connection string do "Transaction pooler" do Supabase (porta
`6543`), não a porta `5432`.** Você encontra em *Project Settings → Database →
Connection string → Transaction pooler*. Isso é importante porque, no Vercel,
cada requisição pode rodar num processo novo — a porta 6543 é feita
especificamente pra aguentar muitas conexões curtas desse tipo, e a 5432 pode
estourar o limite de conexões do banco rapidinho.

⚠️ **Sobre a geração de PDF:** a ficha de adesão é gerada com `reportlab`
(Python puro), não com `weasyprint` — de propósito. O weasyprint depende de
bibliotecas gráficas do sistema (Pango/Cairo) que não têm como ser instaladas
em ambientes serverless como o Vercel; o reportlab não precisa de nada além do
que já está no `requirements.txt`, então roda sem configuração extra.

## Ficha de Adesão em PDF (para assinatura)

Depois que o município envia o formulário, o sistema gera automaticamente uma
**Ficha de Adesão em PDF**, pronta para impressão, com:

- Seções 1 a 5 (município, secretaria, coordenador(a), censo e escolas/matrículas) —
  no mesmo formato dos formulários já usados;
- Uma seção final **"Termo de Adesão e Compromisso"**, com linhas de assinatura
  para o(a) prefeito(a), secretário(a) de educação e coordenador(a) do projeto;
- Uma linha de assinatura para **cada escola participante**, usando os nomes de
  Gestor(a) 1 e Gestor(a) 2 informados no cadastro da escola.

O link para baixar aparece na página de sucesso logo após o envio do formulário,
e também no painel administrativo (`/admin/adesao/<id>`), então o PDF pode ser
gerado novamente a qualquer momento a partir dos dados já salvos.

⚠️ **Sobre a instalação:** o PDF é gerado com `reportlab`, que é Python puro —
não precisa de nenhum programa extra instalado no Windows/Mac/Linux além do
`pip install -r requirements.txt`.

## Próximos passos sugeridos

- Exportar as adesões para Excel/CSV direto do painel admin.
- Enviar e-mail automático (usando `MAIL_USERNAME`/`MAIL_PASSWORD`) quando
  uma nova adesão chegar.
- Colocar autenticação por usuário (em vez de senha única) se mais de uma
  pessoa for usar a área admin.
