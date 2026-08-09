# Publicando o Aplicativo Web (acesso de qualquer lugar)

Este guia explica como colocar o sistema de laudos de OCT na internet,
como um site particular só seu — acessível do computador, celular ou
tablet, sem instalar nada. **Não é necessário usar linha de comando em
nenhum momento**: tudo é feito clicando em telas de sites (GitHub e
Render.com).

Tempo estimado: 20 a 30 minutos, feito uma única vez.

---

## Visão geral (o que vamos fazer)

1. **GitHub** — um "cofre" gratuito onde os arquivos do programa ficam
   guardados (você já usa isso, mesmo sem saber, sempre que baixa um
   programa de código aberto).
2. **Render.com** — o serviço que pega os arquivos do GitHub e os
   "liga", disponibilizando um endereço de internet (URL) que você
   acessa de qualquer navegador.
3. No final, você terá um endereço tipo
   `https://laudos-oct-ocular-oftalmologia.onrender.com` para
   favoritar no celular e no computador.

---

## Passo 1 — Criar conta no GitHub

1. Acesse https://github.com e clique em **Sign up**.
2. Crie a conta com seu e-mail (é gratuito).

## Passo 2 — Criar o repositório e enviar os arquivos

1. Depois de logado, clique no `+` no canto superior direito → **New repository**.
2. Dê um nome, por exemplo `laudos-oct` (pode deixar como **privado** — só você vai ver o código).
3. Clique em **Create repository**.
4. Na página do repositório recém-criado, clique no link **"uploading an existing file"**
   (ou vá em **Add file → Upload files**).
5. Arraste **todos os arquivos e pastas** deste sistema para a janela do navegador:
   - `app_web.py`
   - `nucleo_laudo.py`
   - `main.py`
   - `config_oct.json`
   - `template_laudo.docx`
   - `requirements.txt`
   - `requirements_web.txt`
   - `render.yaml`
   - `Procfile`
   - `.gitignore`
   - a pasta `web` (com o arquivo `index.html` dentro)
   - a pasta `exemplo` (opcional, não é necessária para o site funcionar)
6. Role até o final da página e clique em **Commit changes** (pode manter as opções padrão).

Pronto — o código está no GitHub. **Nunca envie sua chave de API nem
sua senha de acesso para o GitHub** (elas não devem aparecer em
nenhum arquivo — no passo a seguir você vai configurá-las direto no
Render, de forma segura).

## Passo 3 — Criar conta no Render

1. Acesse https://render.com e clique em **Get Started**.
2. É mais rápido escolher **"Sign up with GitHub"** — assim as duas contas já ficam conectadas.

## Passo 4 — Publicar o aplicativo

Este projeto já inclui um arquivo `render.yaml`, que descreve para o
Render exatamente como rodar o aplicativo — isso permite publicar em
poucos cliques usando o recurso **Blueprint**:

1. No painel do Render, clique em **New +** → **Blueprint**.
2. Selecione o repositório `laudos-oct` que você criou no GitHub
   (pode ser necessário autorizar o Render a acessar seus repositórios).
3. O Render vai detectar o `render.yaml` automaticamente e mostrar o
   serviço `laudos-oct-ocular-oftalmologia` pronto para criar.
4. Ele vai pedir para você preencher **duas variáveis secretas** — é
   aqui que sua chave e sua senha entram, com segurança (não ficam no
   GitHub, só no painel do Render):
   - `ANTHROPIC_API_KEY` → sua chave da API da Anthropic (a mesma que
     você usa no aplicativo de mesa).
   - `SENHA_ACESSO` → invente uma senha só sua, que você vai digitar
     no site sempre que for gerar um laudo. Escolha algo fácil de
     digitar no celular, mas que só você saiba.
5. Clique em **Apply** (ou **Create Web Service**).
6. Aguarde alguns minutos — o Render vai instalar as dependências e
   ligar o site. Você pode acompanhar o progresso na aba **Logs**.
7. Quando aparecer **"Live"** (bolinha verde), o site está no ar. A
   URL fica no topo da página do serviço, algo como
   `https://laudos-oct-ocular-oftalmologia.onrender.com`.

> Se preferir não usar o Blueprint, dá para criar manualmente em
> **New + → Web Service**, apontando para o mesmo repositório, com:
> - **Build Command:** `pip install -r requirements_web.txt`
> - **Start Command:** `uvicorn app_web:app --host 0.0.0.0 --port $PORT`
> - E configurar as mesmas duas variáveis de ambiente na aba **Environment**.

## Passo 5 — Testar e favoritar

1. Abra a URL do Render no navegador do computador e do celular.
2. Digite a `SENHA_ACESSO` que você escolheu.
3. Selecione alguns PDFs de teste e clique em **Gerar Laudo** — o
   download do `.docx` deve começar automaticamente.
4. **No celular (deixar com "cara de aplicativo"):**
   - **iPhone (Safari):** toque no ícone de compartilhar (quadrado com
     seta) → **Adicionar à Tela de Início**.
   - **Android (Chrome):** toque nos três pontinhos no canto → **Adicionar
     à tela inicial** (ou **Instalar app**, dependendo da versão).
   - Isso cria um ícone na tela do celular que abre o site direto,
     sem precisar digitar o endereço toda vez.

## Como funciona o "abrir direto no Word"

Quando você toca em "Gerar Laudo", o navegador baixa o arquivo
`.docx` automaticamente. No computador, ele vai para a pasta de
Downloads (alguns navegadores já oferecem "Abrir" ao terminar). No
celular, o sistema normalmente mostra uma notificação de download ou
uma opção para abrir direto no Word, Google Docs ou outro app
instalado — é um toque a mais, mas é o máximo que um site consegue
fazer: por segurança, nenhum site pode forçar a abertura automática de
um outro aplicativo no seu celular sem essa confirmação.

## Atualizando o sistema depois

Sempre que eu (ou você) alterar algum arquivo (ex: ajustar o template,
adicionar um campo), basta enviar o(s) arquivo(s) atualizado(s) de
novo para o mesmo repositório no GitHub (**Add file → Upload files**,
substituindo o antigo). O Render detecta a mudança e republica o site
sozinho, em 1-2 minutos, sem precisar fazer nada no painel do Render.

## Custos

- **Render:** o plano de workspace "Hobby" é gratuito. Para o
  serviço web ficar sempre pronto (sem "dormir" entre usos), o plano
  de computação recomendado é o **Starter, US$ 7/mês** (cobrado em
  dólar no cartão informado). Existe também um nível realmente
  gratuito, mas ele costuma "dormir" depois de um tempo sem uso — a
  primeira geração de laudo depois de um tempo parado pode demorar
  bem mais (30-60 segundos extras) enquanto o servidor "acorda". Para
  uso profissional no dia a dia, vale o US$ 7/mês.
- **Anthropic (Claude API):** cobrança por uso, a mesma do aplicativo
  de mesa — cada laudo gerado consome uma pequena quantidade de
  créditos, proporcional ao número de PDFs enviados.

## Segurança e privacidade

- O site só funciona com a `SENHA_ACESSO` correta — sem ela, ninguém
  consegue gerar laudos nem gastar sua cota de API.
- O Render fornece HTTPS automaticamente (conexão criptografada, o
  cadeado do navegador).
- Os PDFs enviados e o laudo gerado ficam em uma pasta temporária no
  servidor **apenas durante o processamento** — são apagados
  automaticamente logo depois que o download é enviado ao navegador.
- Ainda assim, como se trata de dados de pacientes, revise a política
  de privacidade da Anthropic e do Render, e as normas do seu
  conselho profissional/LGPD antes de usar com exames reais — assim
  como já vale para o aplicativo de mesa.

## Problemas comuns

| Situação | Causa provável | Solução |
|---|---|---|
| Página mostra "Internal Server Error" ou erro 500 sobre `ANTHROPIC_API_KEY`/`SENHA_ACESSO` | Uma das duas variáveis não foi configurada no Render | Vá em **Environment** no painel do serviço e confira as duas variáveis |
| "Senha incorreta" mesmo digitando certo | A senha configurada no Render tem espaços/caracteres diferentes | Reconfirme o valor de `SENHA_ACESSO` no painel do Render (aba Environment) |
| Deploy fica "travado" ou falha | Algum arquivo não foi enviado ao GitHub (ex: esqueceu a pasta `web/`) | Confira nos **Logs** do Render qual arquivo está faltando e reenvie |
| Primeira geração do dia demora muito | Serviço no plano gratuito "dormiu" | Normal no plano gratuito; migre para o Starter (US$ 7/mês) se isso incomodar |
| Laudo não abre sozinho no celular | Comportamento normal do navegador/sistema | Toque na notificação de download ou abra o app de Arquivos/Downloads e toque no `.docx` |
