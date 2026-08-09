"""
=================================================================
 Aplicativo Web — Automação de Laudos de OCT
 Ocular Oftalmologia - Governador Valadares/MG
=================================================================

Este é o aplicativo WEB: a mesma automação de main.py, só que
publicada como um site, acessível pelo navegador de qualquer
computador, celular ou tablet — sem precisar instalar Python nem
nada. Usa a mesma lógica de "nucleo_laudo.py" que o aplicativo de
mesa (main.py); só a "porta de entrada" muda.

Veja README_WEB.md para o passo a passo (sem usar linha de comando)
de como publicar este aplicativo na internet (Render.com).

Variáveis de ambiente necessárias no servidor (configuradas no painel
do serviço de hospedagem, NUNCA dentro do código):
    ANTHROPIC_API_KEY  -> sua chave da API da Anthropic (Claude)
    SENHA_ACESSO        -> uma senha de sua escolha, para impedir que
                            estranhos usem o site e gastem sua cota de API

Para rodar localmente e testar (opcional, para quem tem Python):
    pip install -r requirements_web.txt
    set ANTHROPIC_API_KEY=sua_chave      (Windows)
    set SENHA_ACESSO=uma_senha_de_teste
    uvicorn app_web:app --reload
    Depois abra http://localhost:8000 no navegador.
=================================================================
"""

from __future__ import annotations

import hmac
import os
import shutil
import tempfile
import traceback
import urllib.parse
from pathlib import Path
from typing import List

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask

from nucleo_laudo import ErroLaudoOCT, GeradorLaudoOCT

# -----------------------------------------------------------------
# Caminhos e configuração
# -----------------------------------------------------------------
PASTA_PRINCIPAL = Path(__file__).resolve().parent
CAMINHO_CONFIG = PASTA_PRINCIPAL / "config_oct.json"
CAMINHO_TEMPLATE = PASTA_PRINCIPAL / "template_laudo.docx"
CAMINHO_INDEX_HTML = PASTA_PRINCIPAL / "web" / "index.html"

app = FastAPI(title="Automação de Laudos OCT — Ocular Oftalmologia")

# Instância única do gerador, reaproveitada entre requisições (evita
# reabrir o config/template a cada laudo). É criada sob demanda, na
# primeira requisição, para dar uma mensagem clara se a chave de API
# não estiver configurada — em vez de o servidor inteiro falhar ao ligar.
_gerador: GeradorLaudoOCT = None


def obter_gerador() -> GeradorLaudoOCT:
    """Cria (uma única vez) e devolve a instância do gerador de laudos."""
    global _gerador
    if _gerador is None:
        chave_api = os.environ.get("ANTHROPIC_API_KEY")
        if not chave_api:
            raise HTTPException(
                status_code=500,
                detail=(
                    "O servidor não está configurado corretamente: falta a variável "
                    "de ambiente ANTHROPIC_API_KEY. Configure-a no painel do serviço "
                    "de hospedagem (veja README_WEB.md)."
                ),
            )
        _gerador = GeradorLaudoOCT(
            str(CAMINHO_CONFIG), str(CAMINHO_TEMPLATE), chave_api=chave_api
        )
    return _gerador


def verificar_senha(senha: str):
    """
    Compara a senha enviada pelo navegador com a variável de ambiente
    SENHA_ACESSO. Usa comparação em tempo constante (hmac.compare_digest)
    para não vazar informação sobre a senha por diferença de tempo de
    resposta — boa prática mesmo em um sistema de usuário único.
    """
    senha_correta = os.environ.get("SENHA_ACESSO")
    if not senha_correta:
        raise HTTPException(
            status_code=500,
            detail=(
                "O servidor não está configurado corretamente: falta a variável de "
                "ambiente SENHA_ACESSO. Configure-a no painel do serviço de "
                "hospedagem (veja README_WEB.md)."
            ),
        )
    if not senha or not hmac.compare_digest(senha, senha_correta):
        raise HTTPException(status_code=401, detail="Senha incorreta.")


# -----------------------------------------------------------------
# Rotas
# -----------------------------------------------------------------
@app.get("/", response_class=HTMLResponse)
def pagina_inicial():
    """Serve a página única do aplicativo (login + upload + geração)."""
    if not CAMINHO_INDEX_HTML.exists():
        raise HTTPException(status_code=500, detail="Arquivo web/index.html não encontrado.")
    return CAMINHO_INDEX_HTML.read_text(encoding="utf-8")


@app.get("/api/status")
def status():
    """Rota simples para conferir se o servidor está no ar (health check)."""
    return {"status": "ok"}


@app.post("/api/gerar-laudo")
async def gerar_laudo(senha: str = Form(...), arquivos: List[UploadFile] = File(...)):
    """
    Recebe a senha de acesso + um ou mais PDFs do mesmo exame, gera o
    laudo em Word e devolve o arquivo .docx pronto para download —
    exatamente o mesmo fluxo de GeradorLaudoOCT.processar_exame usado
    no aplicativo de mesa (main.py).
    """
    verificar_senha(senha)

    if not arquivos:
        raise HTTPException(status_code=400, detail="Nenhum arquivo PDF foi enviado.")

    for arquivo in arquivos:
        nome = arquivo.filename or ""
        if not nome.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail=f"O arquivo '{nome}' não é um PDF.")

    pasta_temporaria = Path(tempfile.mkdtemp(prefix="laudo_oct_"))
    try:
        # A criação do cliente da API (obter_gerador) também fica dentro
        # do try: se a chave de API estiver ausente/inválida ou houver
        # qualquer outro problema de configuração, o erro vira uma
        # resposta JSON amigável em vez de derrubar a requisição.
        gerador = obter_gerador()

        caminhos_pdfs = []
        for arquivo in arquivos:
            caminho = pasta_temporaria / arquivo.filename
            conteudo = await arquivo.read()
            caminho.write_bytes(conteudo)
            caminhos_pdfs.append(caminho)

        resultado = gerador.processar_exame(caminhos_pdfs, str(pasta_temporaria))

    except ErroLaudoOCT as erro:
        shutil.rmtree(pasta_temporaria, ignore_errors=True)
        raise HTTPException(status_code=400, detail=str(erro))
    except HTTPException:
        shutil.rmtree(pasta_temporaria, ignore_errors=True)
        raise
    except Exception as erro:
        shutil.rmtree(pasta_temporaria, ignore_errors=True)
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Erro inesperado no servidor: {erro}")

    caminho_docx = Path(resultado["caminho_docx"])

    # Os avisos (não bloqueantes) vão em um cabeçalho HTTP à parte,
    # já que a resposta principal é o próprio arquivo .docx. Precisam
    # ser "urlencoded" porque cabeçalhos HTTP não aceitam acentuação.
    avisos_texto = urllib.parse.quote(" | ".join(resultado["avisos"]), safe="")

    # Apaga a pasta temporária (PDFs + docx + json) depois que a
    # resposta terminar de ser enviada ao navegador — nada de dados de
    # paciente fica armazenado no servidor além do necessário.
    tarefa_limpeza = BackgroundTask(shutil.rmtree, pasta_temporaria, ignore_errors=True)

    return FileResponse(
        path=caminho_docx,
        filename=caminho_docx.name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"X-Avisos": avisos_texto},
        background=tarefa_limpeza,
    )


@app.exception_handler(HTTPException)
async def tratador_erro_http(request, exc: HTTPException):
    """Garante que erros da API sempre voltem como JSON (o front-end espera isso)."""
    return JSONResponse(status_code=exc.status_code, content={"detail": exc.detail})


# Serve arquivos estáticos adicionais da pasta "web" (ex: ícone/imagens), se existirem.
_pasta_web = PASTA_PRINCIPAL / "web"
if _pasta_web.exists():
    app.mount("/web", StaticFiles(directory=str(_pasta_web)), name="web")
