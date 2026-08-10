"""
=================================================================
 Núcleo do Sistema de Automação de Laudos (lógica compartilhada)
 Ocular Oftalmologia - Governador Valadares/MG
=================================================================

Este módulo contém TODA a lógica de extração/validação/geração de
laudos, sem NENHUMA dependência de interface gráfica. Ele é usado por
dois "front-ends" diferentes:

    - main.py     -> aplicativo de mesa (janela Tkinter, roda no seu
                      computador Windows/Mac/Linux)
    - app_web.py   -> aplicativo web (FastAPI), acessível pelo
                       navegador em qualquer computador, celular ou
                       tablet, depois de publicado na internet

Isso evita duplicar a lógica de extração/geração em dois lugares —
qualquer melhoria feita aqui (ex: um novo campo, uma correção na
extração) vale automaticamente para os dois aplicativos.

Arquitetura (classes principais):
    - ExamConfig      : carrega a configuração do tipo de exame (JSON)
    - PDFExtractor     : lê o(s) PDF(s) e extrai os dados via Claude API
    - ValidadorDados    : valida os dados extraídos contra a configuração
    - WordGenerator      : preenche o template Word com os dados
    - GeradorLaudoOCT     : orquestra todo o fluxo (extrair -> validar -> gerar)

PRONTO PARA ADICIONAR NOVOS TIPOS DE EXAME:
    Para suportar um novo exame (ex: Campo Visual, Topografia, Biometria),
    basta criar um novo arquivo "config_<tipo>.json" seguindo a mesma
    estrutura de "config_oct.json" e um novo "template_<tipo>.docx".
    Nenhuma classe abaixo precisa ser alterada — todas recebem o
    caminho da configuração/template como parâmetro.

Instale as dependências com: pip install -r requirements.txt
=================================================================
"""

from __future__ import annotations

import os
import re
import json
import base64
from pathlib import Path
from datetime import datetime, date

# -----------------------------------------------------------------
# Importação das bibliotecas externas.
# Se alguma não estiver instalada, o erro só é lançado quando a
# funcionalidade correspondente for realmente usada — assim quem
# estiver chamando (janela Tkinter ou rota web) consegue mostrar uma
# mensagem amigável em vez de travar com um erro técnico do Python.
# -----------------------------------------------------------------
try:
    import fitz  # PyMuPDF -> leitura de PDF e conversão de páginas em imagem
except ImportError:
    fitz = None

try:
    import docx  # python-docx -> geração/edição do laudo Word
except ImportError:
    docx = None

try:
    import anthropic  # SDK oficial da Anthropic -> chamada à Claude API
except ImportError:
    anthropic = None


# ===================================================================
# EXCEÇÕES PERSONALIZADAS
# Facilitam mostrar mensagens de erro claras e específicas, em vez de
# um traceback genérico do Python (tanto na janela quanto na web).
# ===================================================================
class ErroLaudoOCT(Exception):
    """Classe base para todos os erros previstos deste sistema."""
    pass


class ErroConfiguracao(ErroLaudoOCT):
    """Erro ao carregar/ler o arquivo de configuração (config_*.json)."""
    pass


class ErroExtracaoPDF(ErroLaudoOCT):
    """Erro ao ler o(s) PDF(s) ou ao extrair dados via Claude API."""
    pass


class ErroValidacaoDados(ErroLaudoOCT):
    """Erro quando dados obrigatórios não foram encontrados/estão inválidos."""
    pass


class ErroGeracaoDocumento(ErroLaudoOCT):
    """Erro ao preencher ou salvar o documento Word final."""
    pass


# ===================================================================
# REGISTRO DE EQUIPAMENTOS SUPORTADOS
# Cada entrada aponta para o arquivo de configuração e o template Word
# daquele equipamento/exame. É usado pelos dois "front-ends" (main.py
# e app_web.py) para montar o seletor de equipamento — assim os dois
# aplicativos sempre mostram exatamente as mesmas opções.
#
# PRONTO PARA ADICIONAR UM NOVO EQUIPAMENTO (ex: Topcon Triton):
#   1. Crie "config_<equipamento>.json" e "template_laudo_<equipamento>.docx"
#      seguindo o padrão dos exemplos já existentes.
#   2. Adicione uma linha abaixo, com uma chave curta (ex: "topcon").
#   Pronto — o aplicativo de mesa e o aplicativo web já passam a
#   oferecer essa opção automaticamente, sem mais nenhuma alteração.
# ===================================================================
EQUIPAMENTOS_SUPORTADOS = {
    "zeiss": {
        "nome_exibicao": "Zeiss Cirrus HD-OCT",
        "config": "config_oct.json",
        "template": "template_laudo.docx",
    },
    "nidek": {
        "nome_exibicao": "Nidek RS-3000 Advance",
        "config": "config_nidek.json",
        "template": "template_laudo_nidek.docx",
    },
    "pentacam": {
        "nome_exibicao": "Pentacam® (Oculus) — Tomografia de Córnea",
        "config": "config_pentacam.json",
        "template": "template_laudo_pentacam.docx",
    },
    "campo_visual": {
        "nome_exibicao": "Octopus 600 (Haag-Streit) — Campo Visual",
        "config": "config_campo_visual.json",
        "template": "template_laudo_campo_visual.docx",
    },
}


# ===================================================================
# FUNÇÕES UTILITÁRIAS (datas, idade, protocolo)
# Ficam fora das classes por serem funções puras/independentes,
# reaproveitáveis por qualquer parte do sistema (e fáceis de testar).
# ===================================================================

# Formatos de data aceitos na entrada, na ordem em que são tentados.
_FORMATOS_DATA_ACEITOS = ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y")


def tentar_parsear_data(texto: str):
    """
    Tenta interpretar uma string de data em vários formatos possíveis
    (o equipamento Zeiss Cirrus exporta datas no formato americano
    M/D/AAAA, mas a IA pode ou não convertê-las). Retorna um objeto
    'date' do Python em caso de sucesso, ou None se não conseguir.
    """
    if not texto or not isinstance(texto, str):
        return None
    texto = texto.strip()
    for formato in _FORMATOS_DATA_ACEITOS:
        try:
            return datetime.strptime(texto, formato).date()
        except ValueError:
            continue
    return None


def normalizar_data(texto: str) -> str:
    """
    Converte uma data para o formato brasileiro dd/mm/aaaa, aceitando
    também o formato americano mm/dd/aaaa (comum nas exportações do
    equipamento). Se não conseguir interpretar, devolve o texto original
    sem alteração (para não perder a informação).
    """
    data = tentar_parsear_data(texto)
    if data is None:
        return texto
    return data.strftime("%d/%m/%Y")


def calcular_idade(data_nascimento_texto: str, data_referencia_texto: str) -> str:
    """
    Calcula a idade (em anos completos) a partir da data de nascimento
    e da data do exame. Retorna uma string como "73 anos", ou None se
    não for possível calcular (datas ausentes/ilegíveis) — nesse caso
    o campo simplesmente fica em branco no laudo, sem travar o programa.
    """
    nascimento = tentar_parsear_data(data_nascimento_texto)
    referencia = tentar_parsear_data(data_referencia_texto) or date.today()
    if nascimento is None:
        return None
    anos = referencia.year - nascimento.year
    fez_aniversario = (referencia.month, referencia.day) >= (nascimento.month, nascimento.day)
    if not fez_aniversario:
        anos -= 1
    if anos < 0 or anos > 130:
        return None  # data provavelmente ilegível/errada — não arrisca mostrar um valor absurdo
    return f"{anos} anos"


# Lista padrão de palavras-chave, usada apenas como último recurso se
# um arquivo config_*.json não definir a própria lista em
# "protocolos_por_palavra_chave" (ver ExamConfig.protocolos_por_palavra_chave
# mais abaixo). Cada equipamento nomeia os arquivos de um jeito — por
# isso o ideal é sempre definir essa lista dentro do config de cada
# equipamento, não aqui no núcleo.
_PROTOCOLOS_PADRAO_FALLBACK = [
    ("Macular Thickness", "Espessura Macular (Retina)"),
    ("Ganglion Cell", "Complexo de Células Ganglionares (GCC)"),
    ("ONH and RNFL", "Nervo Óptico e RNFL (Glaucoma)"),
    ("HD 5 Line Raster", "Cortes de Alta Definição (HD Raster)"),
]


def detectar_protocolo(nomes_arquivos: list, mapa_protocolos: list = None) -> str:
    """
    Monta uma descrição textual de quais protocolos de varredura foram
    incluídos no exame, a partir dos nomes dos arquivos PDF selecionados
    (cada equipamento nomeia os arquivos de forma padronizada, incluindo
    o nome do protocolo). 'mapa_protocolos' é a lista de (palavra_chave,
    descrição) vinda do config_*.json do equipamento em uso — se não for
    informada, usa uma lista padrão de fallback. Se nada for reconhecido,
    retorna None.
    """
    mapa = mapa_protocolos or _PROTOCOLOS_PADRAO_FALLBACK
    protocolos_encontrados = []
    for nome in nomes_arquivos:
        for palavra_chave, descricao in mapa:
            if palavra_chave.lower() in nome.lower() and descricao not in protocolos_encontrados:
                protocolos_encontrados.append(descricao)
    if not protocolos_encontrados:
        return None
    return " + ".join(protocolos_encontrados)


def gerar_nome_base_arquivo(dados: dict) -> str:
    """
    Gera um nome de arquivo seguro (sem caracteres especiais) a partir
    do nome do paciente e da data do exame, usado para nomear o .docx
    e o .json de saída (já que não há mais um único PDF de origem).
    """
    nome_paciente = dados.get("nome_paciente") or "paciente"
    data_exame = dados.get("data_exame") or ""

    texto = f"{nome_paciente}_{data_exame}"
    texto = texto.strip().replace(" ", "_")
    # Remove qualquer caractere que não seja letra, número, underscore ou hífen
    texto = re.sub(r"[^A-Za-z0-9_\-]", "", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto or "laudo_oct"


def completar_campos_automaticos(dados: dict, nomes_arquivos: list, config: "ExamConfig") -> dict:
    """
    Preenche os campos que o programa calcula sozinho (não pedidos à
    IA): datas normalizadas, idade, protocolo do exame e dados do
    médico responsável (vindos do arquivo de configuração). Fica como
    função de módulo (em vez de método de GeradorLaudoOCT) para poder
    ser reaproveitada também por scripts de teste, sem precisar de
    chave de API nem de PDFs reais — ver exemplo/teste_geracao_laudo.py.
    """
    dados = dict(dados)

    if dados.get("data_nascimento"):
        dados["data_nascimento"] = normalizar_data(dados["data_nascimento"])
    if dados.get("data_exame"):
        dados["data_exame"] = normalizar_data(dados["data_exame"])

    idade_calculada = calcular_idade(dados.get("data_nascimento"), dados.get("data_exame"))
    if idade_calculada:
        dados["idade"] = idade_calculada

    protocolo = detectar_protocolo(nomes_arquivos, config.protocolos_por_palavra_chave)
    if protocolo:
        dados["protocolo"] = protocolo

    medico = config.medico_responsavel
    dados["nome_medico"] = medico.get("nome", "")
    dados["crm_medico"] = medico.get("crm", "")
    dados["data_geracao_laudo"] = datetime.now().strftime("%d/%m/%Y %H:%M")

    return dados


# ===================================================================
# CLASSE: ExamConfig
# Responsabilidade única: carregar e disponibilizar a configuração
# de um tipo de exame (campos, validações, prompt de extração).
# ===================================================================
class ExamConfig:
    """
    Carrega a configuração de um tipo de exame a partir de um arquivo
    JSON (ex: config_oct.json) e expõe métodos utilitários para as
    demais classes do sistema.
    """

    def __init__(self, caminho_config: str):
        self.caminho_config = Path(caminho_config)
        self.dados_config = self._carregar_config()

    def _carregar_config(self) -> dict:
        """Lê e valida o arquivo JSON de configuração."""
        if not self.caminho_config.exists():
            raise ErroConfiguracao(
                f"Arquivo de configuração não encontrado: '{self.caminho_config}'.\n"
                "Verifique se o arquivo config_oct.json está na mesma pasta do programa."
            )
        try:
            with open(self.caminho_config, "r", encoding="utf-8") as arquivo:
                config = json.load(arquivo)
        except json.JSONDecodeError as erro:
            raise ErroConfiguracao(
                f"O arquivo de configuração '{self.caminho_config.name}' possui um "
                f"JSON inválido (erro na linha {erro.lineno}). Verifique a formatação."
            )
        if "campos" not in config or not config["campos"]:
            raise ErroConfiguracao(
                f"O arquivo '{self.caminho_config.name}' não define nenhum campo em 'campos'."
            )
        return config

    # ---------------- Propriedades de acesso rápido ----------------
    @property
    def tipo_exame(self) -> str:
        return self.dados_config.get("tipo_exame", "DESCONHECIDO")

    @property
    def nome_exibicao(self) -> str:
        return self.dados_config.get("nome_exibicao", self.tipo_exame)

    @property
    def campos(self) -> list:
        return self.dados_config.get("campos", [])

    @property
    def modelo_claude(self) -> str:
        return self.dados_config.get("modelo_claude", "claude-sonnet-5")

    @property
    def opcoes_extracao(self) -> dict:
        return self.dados_config.get("opcoes_extracao", {})

    @property
    def medico_responsavel(self) -> dict:
        return self.dados_config.get("medico_responsavel", {"nome": "", "crm": ""})

    @property
    def protocolos_por_palavra_chave(self) -> list:
        """
        Lista de (palavra_chave, descrição) usada para reconhecer, pelo
        nome dos arquivos PDF, quais protocolos de varredura foram
        incluídos no exame — específica de cada equipamento/config.
        Formato no JSON: [["palavra", "descrição"], ...].
        """
        pares = self.dados_config.get("protocolos_por_palavra_chave", [])
        return [tuple(par) for par in pares]

    @property
    def palavras_chave_verificacao(self) -> list:
        """
        Lista de palavras-chave (em minúsculas) que devem aparecer no
        texto do campo 'equipamento' lido pela IA diretamente do exame,
        usada como confirmação de que os PDFs enviados realmente são
        deste equipamento — proteção contra o usuário selecionar o
        equipamento errado no seletor (ex: enviar PDFs do Pentacam com
        'Zeiss Cirrus' escolhido no aplicativo). Ver ValidadorDados.
        """
        return [p.lower() for p in self.dados_config.get("palavras_chave_verificacao", [])]

    # ---------------- Métodos utilitários ----------------
    def obter_campo(self, campo_id: str) -> dict:
        """Retorna a definição de um campo específico pelo seu id."""
        for campo in self.campos:
            if campo["id"] == campo_id:
                return campo
        return None

    def campos_obrigatorios(self) -> list:
        """Retorna a lista de ids dos campos marcados como obrigatórios."""
        return [campo["id"] for campo in self.campos if campo.get("obrigatorio")]

    def montar_prompt_extracao(self) -> str:
        """
        Monta o prompt textual enviado à Claude API, descrevendo com
        precisão cada campo que deve ser extraído do exame, e de qual
        protocolo/arquivo cada um normalmente é originado.
        """
        base = self.dados_config.get("prompt_extracao_base", "")

        linhas_campos = []
        exemplo_json = {}
        for campo in self.campos:
            descricao = f'- "{campo["id"]}": {campo["rotulo"]}'
            if campo.get("unidade"):
                descricao += f' (unidade: {campo["unidade"]})'
            if campo.get("origem"):
                descricao += f' (protocolo/arquivo de origem esperado: {campo["origem"]})'
            if campo.get("valores_permitidos"):
                valores = ", ".join(campo["valores_permitidos"])
                descricao += f' (valores possíveis: {valores})'
            if campo.get("exemplos"):
                exemplos = ", ".join(campo["exemplos"])
                descricao += f' (exemplos: {exemplos})'
            linhas_campos.append(descricao)
            exemplo_json[campo["id"]] = None

        campos_texto = "\n".join(linhas_campos)
        exemplo_json_texto = json.dumps(exemplo_json, ensure_ascii=False, indent=2)

        instrucoes_finais = (
            "\n\nRegras importantes:\n"
            "1. Responda ESTRITAMENTE com um objeto JSON válido, sem nenhum texto, "
            "comentário ou marcação (como ```) antes ou depois.\n"
            "2. Use exatamente os identificadores de campo indicados acima como chaves do JSON, "
            "e APENAS esses identificadores (não crie campos extras, não aninhe objetos/dicionários "
            "dentro de um único campo).\n"
            "3. Se um dado não estiver visível, não fizer parte de nenhum dos arquivos fornecidos, "
            "ou pertencer a um protocolo cujo arquivo não foi enviado, use o valor null — nunca "
            "invente ou estime um valor.\n"
            "4. Cada arquivo fornecido abaixo está identificado por 'Arquivo N de M: <nome_do_arquivo>'. "
            "Use o nome do arquivo (e o texto extraído dele) para saber de qual protocolo os dados de "
            "cada campo devem ser lidos, conforme indicado entre parênteses em cada campo acima.\n"
            "5. Números devem ser retornados sem unidade (a unidade já é conhecida pelo sistema) e "
            "usando ponto como separador decimal (ex: 0.63, não 0,63).\n"
            "6. Datas devem ser retornadas no formato dd/mm/aaaa.\n"
            "7. Não faça diagnóstico nem interpretação clínica: apenas transcreva os dados exibidos.\n\n"
            f"Formato exato esperado (chaves, com valores de exemplo nulos):\n{exemplo_json_texto}"
        )

        return (
            f"{base}\n\n"
            f"Extraia os seguintes campos do exame de {self.nome_exibicao}:\n"
            f"{campos_texto}"
            f"{instrucoes_finais}"
        )


# ===================================================================
# CLASSE: PDFExtractor
# Responsabilidade única: ler o(s) PDF(s) de um mesmo exame e obter
# os dados estruturados usando a Claude API (texto + imagens).
# ===================================================================
class PDFExtractor:
    """
    Lê um ou mais PDFs de um mesmo exame (texto selecionável + imagens
    das páginas) e envia esse conteúdo, em uma única chamada, à Claude
    API para extração estruturada dos dados definidos em ExamConfig.
    """

    def __init__(self, config: ExamConfig, chave_api: str = None):
        if fitz is None:
            raise ErroExtracaoPDF(
                "A biblioteca PyMuPDF não está instalada.\n"
                "Instale com: pip install -r requirements.txt"
            )
        if anthropic is None:
            raise ErroExtracaoPDF(
                "A biblioteca 'anthropic' não está instalada.\n"
                "Instale com: pip install -r requirements.txt"
            )

        self.config = config
        opcoes = config.opcoes_extracao
        self.dpi_imagem = opcoes.get("dpi_imagem", 200)
        self.max_paginas_por_pdf = opcoes.get("max_paginas_por_pdf", 2)
        self.max_tokens_resposta = opcoes.get("max_tokens_resposta", 3072)
        # As imagens são enviadas em JPEG (não PNG): mapas/gráficos de OCT
        # são cheios de gradientes coloridos, que o PNG comprime muito mal
        # (arquivos de 4-6 MB por página não é incomum) — isso pode fazer o
        # pedido ultrapassar o limite de tamanho da API da Anthropic quando
        # o exame tem vários PDFs (erro HTTP 413 "request_too_large"). Em
        # JPEG, com qualidade alta, o mesmo conteúdo fica 5-7x menor sem
        # perda perceptível de legibilidade dos números.
        self.qualidade_jpeg = opcoes.get("qualidade_jpeg", 85)

        chave = chave_api or os.environ.get("ANTHROPIC_API_KEY")
        if not chave:
            raise ErroExtracaoPDF(
                "Chave da API da Anthropic não encontrada.\n"
                "Defina a variável de ambiente ANTHROPIC_API_KEY ou informe a "
                "chave quando solicitado."
            )
        self.cliente = anthropic.Anthropic(api_key=chave)

    def _extrair_texto(self, caminho_pdf: Path) -> str:
        """Extrai o texto selecionável de todas as páginas do PDF."""
        try:
            paginas_texto = []
            with fitz.open(caminho_pdf) as documento:
                for pagina in documento:
                    paginas_texto.append(pagina.get_text())
            return "\n".join(paginas_texto).strip()
        except Exception as erro:
            raise ErroExtracaoPDF(f"Falha ao ler o texto do PDF '{caminho_pdf.name}': {erro}")

    def _converter_paginas_em_imagens(self, caminho_pdf: Path) -> list:
        """
        Converte as páginas do PDF em imagens JPEG (base64), pois a
        maior parte dos valores de um exame de OCT aparece dentro de
        mapas/gráficos coloridos (imagem), não como texto selecionável.
        JPEG é usado em vez de PNG por comprimir muito melhor esse tipo
        de conteúdo (gradientes de cor), mantendo o pedido dentro do
        limite de tamanho da API mesmo com vários PDFs no mesmo exame.
        """
        imagens_base64 = []
        try:
            with fitz.open(caminho_pdf) as documento:
                zoom = self.dpi_imagem / 72
                matriz = fitz.Matrix(zoom, zoom)
                for indice, pagina in enumerate(documento):
                    if indice >= self.max_paginas_por_pdf:
                        break
                    pixmap = pagina.get_pixmap(matrix=matriz)
                    imagem_bytes = pixmap.tobytes("jpg", jpg_quality=self.qualidade_jpeg)
                    imagens_base64.append(base64.b64encode(imagem_bytes).decode("utf-8"))
            return imagens_base64
        except Exception as erro:
            raise ErroExtracaoPDF(
                f"Falha ao converter as páginas do PDF '{caminho_pdf.name}' em imagem: {erro}"
            )

    def _parsear_json(self, texto_resposta: str, motivo_parada: str = None) -> dict:
        """Extrai e valida o bloco JSON contido na resposta da IA."""
        texto = texto_resposta.strip()

        # Remove blocos de código markdown (```json ... ```), caso a IA os inclua
        if texto.startswith("```"):
            texto = texto.strip("`")
            if texto.lower().startswith("json"):
                texto = texto[4:]

        inicio = texto.find("{")
        fim = texto.rfind("}")
        if inicio == -1 or fim == -1:
            # Diagnóstico: mostra um trecho da resposta real da IA (ou avisa
            # se veio vazia), para dar pistas concretas do que aconteceu —
            # em vez de só uma mensagem genérica.
            if not texto:
                detalhe = "A IA devolveu uma resposta vazia (nenhum texto)."
            else:
                trecho = texto[:600]
                detalhe = f"Texto recebido da IA (trecho, para diagnóstico):\n\"{trecho}\""

            dica_parada = ""
            if motivo_parada and motivo_parada != "end_turn":
                dica_parada = (
                    f"\n\nMotivo de parada informado pela API: '{motivo_parada}'. "
                    + (
                        "Isso normalmente indica que a resposta foi cortada por atingir "
                        "o limite de tokens — tente novamente enviando menos PDFs de uma "
                        "vez, ou aumente 'max_tokens_resposta' no arquivo de configuração."
                        if motivo_parada == "max_tokens"
                        else "Verifique se o pedido não foi bloqueado por algum filtro de segurança."
                    )
                )

            raise ErroExtracaoPDF(
                "A resposta da IA não contém um JSON reconhecível. "
                "Tente novamente ou verifique a qualidade dos PDFs.\n\n"
                f"{detalhe}{dica_parada}"
            )

        bloco_json = texto[inicio: fim + 1]
        try:
            return json.loads(bloco_json)
        except json.JSONDecodeError as erro:
            trecho = bloco_json[-400:] if len(bloco_json) > 400 else bloco_json
            dica_parada = ""
            if motivo_parada == "max_tokens":
                dica_parada = (
                    "\n\nA resposta parece ter sido cortada por atingir o limite de "
                    "tokens (max_tokens) — tente novamente enviando menos PDFs de uma "
                    "vez, ou aumente 'max_tokens_resposta' no arquivo de configuração "
                    "deste equipamento."
                )
            raise ErroExtracaoPDF(
                f"Não foi possível interpretar o JSON retornado pela IA: {erro}\n\n"
                f"Final do texto recebido (para diagnóstico):\n\"...{trecho}\""
                f"{dica_parada}"
            )

    def extrair_dados(self, caminhos_pdfs) -> dict:
        """
        Fluxo completo de extração: lê um ou mais PDFs do mesmo exame,
        monta uma única mensagem para a Claude API (texto + imagens de
        todos os arquivos, cada um identificado pelo nome) e retorna os
        dados combinados em formato dict.

        'caminhos_pdfs' pode ser um único caminho (str) ou uma lista de
        caminhos — um exame de OCT normalmente é composto por vários PDFs.
        """
        if isinstance(caminhos_pdfs, (str, Path)):
            caminhos_pdfs = [caminhos_pdfs]
        if not caminhos_pdfs:
            raise ErroExtracaoPDF("Nenhum arquivo PDF foi selecionado.")

        caminhos = [Path(caminho) for caminho in caminhos_pdfs]
        for caminho in caminhos:
            if not caminho.exists():
                raise ErroExtracaoPDF(f"Arquivo PDF não encontrado: '{caminho}'.")
            if caminho.suffix.lower() != ".pdf":
                raise ErroExtracaoPDF(f"O arquivo '{caminho.name}' não é um PDF.")

        prompt = self.config.montar_prompt_extracao()
        conteudo_mensagem = [{"type": "text", "text": prompt}]

        total_arquivos = len(caminhos)
        for indice, caminho in enumerate(caminhos, start=1):
            texto_pdf = self._extrair_texto(caminho)
            imagens_base64 = self._converter_paginas_em_imagens(caminho)

            cabecalho_arquivo = f"\n--- Arquivo {indice} de {total_arquivos}: {caminho.name} ---"
            if texto_pdf:
                cabecalho_arquivo += f"\nTexto extraído (pode estar incompleto):\n{texto_pdf}"
            conteudo_mensagem.append({"type": "text", "text": cabecalho_arquivo})

            for imagem_b64 in imagens_base64:
                conteudo_mensagem.append({
                    "type": "image",
                    "source": {
                        "type": "base64",
                        "media_type": "image/jpeg",
                        "data": imagem_b64,
                    },
                })

        # Verificação proativa de tamanho: a API da Anthropic rejeita
        # pedidos acima de um certo tamanho (erro HTTP 413
        # "request_too_large"). Em vez de deixar o usuário receber esse
        # erro técnico sem explicação, avisamos antes com uma sugestão
        # prática (reduzir a resolução das imagens ou dividir o exame em
        # duas gerações). O limite abaixo (25 MB de conteúdo base64) fica
        # com folga do limite real da API, já que o texto do prompt e a
        # própria requisição HTTP somam um pouco mais por cima.
        tamanho_estimado_bytes = sum(
            len(bloco.get("source", {}).get("data", "")) for bloco in conteudo_mensagem
            if bloco.get("type") == "image"
        )
        if tamanho_estimado_bytes > 25_000_000:
            raise ErroExtracaoPDF(
                f"Os PDFs selecionados, convertidos em imagem, somam cerca de "
                f"{tamanho_estimado_bytes / 1_000_000:.0f} MB — isso provavelmente "
                "ultrapassa o limite de tamanho de pedido da API da Anthropic e "
                "resultaria em erro 'request_too_large'.\n\n"
                "Sugestões: (1) gere o laudo em duas etapas, selecionando menos "
                "PDFs de cada vez e completando os campos manualmente depois, ou "
                "(2) reduza o valor de 'dpi_imagem' no arquivo de configuração "
                "deste equipamento (ex: de 220 para 150) para diminuir o tamanho "
                "das imagens enviadas."
            )

        try:
            resposta = self.cliente.messages.create(
                model=self.config.modelo_claude,
                max_tokens=self.max_tokens_resposta,
                messages=[{"role": "user", "content": conteudo_mensagem}],
            )
        except Exception as erro:
            raise ErroExtracaoPDF(
                f"Erro ao chamar a API da Claude: {erro}\n"
                "Verifique sua conexão com a internet e se a chave de API é válida."
            )

        texto_resposta = "".join(
            bloco.text for bloco in resposta.content if hasattr(bloco, "text")
        )
        motivo_parada = getattr(resposta, "stop_reason", None)
        return self._parsear_json(texto_resposta, motivo_parada=motivo_parada)


# ===================================================================
# CLASSE: ValidadorDados
# Responsabilidade única: validar os dados extraídos contra as
# regras definidas em ExamConfig (obrigatoriedade, valores permitidos).
# ===================================================================
class ValidadorDados:
    """Valida os dados extraídos de um exame contra a configuração."""

    def __init__(self, config: ExamConfig):
        self.config = config

    def validar(self, dados: dict) -> list:
        """
        Verifica os dados extraídos. Lança ErroValidacaoDados se algum
        campo obrigatório estiver ausente. Retorna uma lista de avisos
        (não bloqueantes) para inconsistências menores.
        """
        self._verificar_equipamento_correto(dados)

        campos_faltantes = []
        for campo_id in self.config.campos_obrigatorios():
            valor = dados.get(campo_id)
            if valor in (None, "", "null"):
                campo = self.config.obter_campo(campo_id)
                rotulo = campo["rotulo"] if campo else campo_id
                campos_faltantes.append(rotulo)

        if campos_faltantes:
            lista = ", ".join(campos_faltantes)
            raise ErroValidacaoDados(
                f"Não foi possível identificar os seguintes campos obrigatórios no "
                f"exame: {lista}. Verifique a qualidade dos PDFs ou preencha manualmente "
                "no laudo gerado."
            )

        avisos = []
        for campo in self.config.campos:
            valor = dados.get(campo["id"])
            valores_permitidos = campo.get("valores_permitidos")
            if valor and valores_permitidos:
                if str(valor).strip().upper() not in [v.upper() for v in valores_permitidos]:
                    avisos.append(
                        f"Campo '{campo['rotulo']}' retornou o valor '{valor}', fora do "
                        f"esperado {valores_permitidos}. Confira manualmente."
                    )
        return avisos

    def _verificar_equipamento_correto(self, dados: dict):
        """
        Confere se o texto de identificação do equipamento, lido pela IA
        diretamente das imagens do exame (campo 'equipamento'), é
        compatível com o equipamento selecionado no aplicativo. Isso
        evita que um exame de um equipamento (ex: Pentacam) seja
        processado com o modelo/config de outro (ex: Zeiss Cirrus) só
        porque o usuário esqueceu de trocar o seletor — nesse caso, os
        campos clínicos ficam todos em branco no laudo, o que é confuso
        e só é percebido depois de já ter sido gerado.

        Se o exame não tiver essa informação legível (campo vazio/null),
        a verificação é pulada — não bloqueia o fluxo normal por causa
        de uma imagem de baixa qualidade.
        """
        palavras_chave = self.config.palavras_chave_verificacao
        if not palavras_chave:
            return

        texto_equipamento = str(dados.get("equipamento") or "").strip()
        if not texto_equipamento or texto_equipamento.lower() == "null":
            return

        texto_lower = texto_equipamento.lower()
        if any(palavra in texto_lower for palavra in palavras_chave):
            return

        raise ErroValidacaoDados(
            f"Os PDFs selecionados podem não ser do equipamento escolhido "
            f"('{self.config.nome_exibicao}'). O texto de identificação do "
            f"equipamento lido dentro do próprio exame foi: '{texto_equipamento}'.\n\n"
            "Verifique se você selecionou o equipamento correto no seletor do "
            "aplicativo antes de gerar o laudo (ex: se o exame é do Pentacam, "
            "selecione 'Pentacam® (Oculus)', não 'Zeiss Cirrus HD-OCT' ou "
            "'Nidek RS-3000 Advance'), e gere o laudo novamente."
        )


# ===================================================================
# CLASSE: WordGenerator
# Responsabilidade única: preencher o template Word com os dados
# já extraídos e validados, gerando o laudo final.
# ===================================================================
class WordGenerator:
    """
    Preenche um template .docx substituindo marcadores no formato
    {{id_do_campo}} pelos valores extraídos do exame.
    """

    def __init__(self, caminho_template: str):
        if docx is None:
            raise ErroGeracaoDocumento(
                "A biblioteca python-docx não está instalada.\n"
                "Instale com: pip install -r requirements.txt"
            )
        self.caminho_template = Path(caminho_template)
        if not self.caminho_template.exists():
            raise ErroGeracaoDocumento(
                f"Template Word não encontrado: '{self.caminho_template}'.\n"
                "Verifique se o arquivo template_laudo.docx está na mesma pasta do programa."
            )

    def _valor_para_texto(self, valor) -> str:
        """
        Converte um valor extraído em texto para exibição no laudo.
        Valores ausentes (None/vazio) aparecem como travessão "—", para
        não deixar células em branco de forma confusa no documento.
        """
        if valor is None:
            return "—"
        texto = str(valor).strip()
        return texto if texto and texto.lower() != "null" else "—"

    def _aplicar_substituicoes(self, texto: str, dados: dict) -> str:
        """Substitui todos os marcadores {{campo}} presentes em um texto."""
        resultado = texto
        for chave, valor in dados.items():
            marcador = "{{" + chave + "}}"
            if marcador in resultado:
                resultado = resultado.replace(marcador, self._valor_para_texto(valor))

        # Rede de segurança: se sobrar algum marcador {{...}} sem
        # correspondência em 'dados' (ex: a IA não retornou aquele
        # campo, ou o protocolo não foi reconhecido), troca por "—" em
        # vez de deixar o texto "{{campo}}" cru e confuso no laudo.
        if "{{" in resultado:
            resultado = re.sub(r"\{\{[a-zA-Z0-9_]+\}\}", "—", resultado)
        return resultado

    def _substituir_em_paragrafo(self, paragrafo, dados: dict):
        """
        Substitui marcadores dentro de um parágrafo, preservando a
        formatação do primeiro "run" (trecho de texto) do parágrafo.
        """
        texto_completo = "".join(run.text for run in paragrafo.runs)
        if "{{" not in texto_completo:
            return
        novo_texto = self._aplicar_substituicoes(texto_completo, dados)
        if novo_texto == texto_completo or not paragrafo.runs:
            return
        paragrafo.runs[0].text = novo_texto
        for run in paragrafo.runs[1:]:
            run.text = ""

    def gerar_laudo(self, dados: dict, caminho_saida: str):
        """Gera o documento Word final preenchido e salva no caminho indicado."""
        try:
            documento = docx.Document(self.caminho_template)
        except Exception as erro:
            raise ErroGeracaoDocumento(f"Falha ao abrir o template Word: {erro}")

        # Substitui marcadores em parágrafos "soltos" do documento
        for paragrafo in documento.paragraphs:
            self._substituir_em_paragrafo(paragrafo, dados)

        # Substitui marcadores dentro de tabelas (identificação, parâmetros etc.)
        for tabela in documento.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for paragrafo in celula.paragraphs:
                        self._substituir_em_paragrafo(paragrafo, dados)

        caminho_saida = Path(caminho_saida)
        caminho_saida.parent.mkdir(parents=True, exist_ok=True)
        try:
            documento.save(caminho_saida)
        except PermissionError:
            raise ErroGeracaoDocumento(
                f"Não foi possível salvar '{caminho_saida.name}'. Feche o arquivo caso "
                "ele já esteja aberto no Word e tente novamente."
            )
        except Exception as erro:
            raise ErroGeracaoDocumento(f"Falha ao salvar o documento Word: {erro}")


# ===================================================================
# CLASSE: GeradorLaudoOCT (Orquestrador)
# Junta ExamConfig + PDFExtractor + ValidadorDados + WordGenerator
# em um fluxo único, do(s) PDF(s) de entrada ao laudo Word de saída.
#
# PRONTO PARA ADICIONAR NOVOS TIPOS DE EXAME: esta mesma classe pode
# ser reaproveitada para outro exame apenas trocando os caminhos de
# configuração e template passados no construtor.
# ===================================================================
class GeradorLaudoOCT:
    """Orquestra o fluxo completo: extração -> validação -> geração do laudo."""

    def __init__(self, caminho_config: str, caminho_template: str, chave_api: str = None):
        self.config = ExamConfig(caminho_config)
        self.extrator = PDFExtractor(self.config, chave_api=chave_api)
        self.validador = ValidadorDados(self.config)
        self.gerador_word = WordGenerator(caminho_template)

    def processar_exame(self, caminhos_pdfs, pasta_saida: str, callback_status=None) -> dict:
        """
        Executa o fluxo completo para um exame (um ou mais PDFs do mesmo
        paciente/exame) e retorna um dicionário com os caminhos gerados,
        avisos e dados extraídos.
        """
        def status(mensagem):
            if callback_status:
                callback_status(mensagem)

        if isinstance(caminhos_pdfs, (str, Path)):
            caminhos_pdfs = [caminhos_pdfs]
        caminhos_pdfs = [Path(caminho) for caminho in caminhos_pdfs]

        status(
            f"Lendo {len(caminhos_pdfs)} arquivo(s) PDF e extraindo dados com a "
            "Claude API (pode levar alguns segundos)..."
        )
        dados = self.extrator.extrair_dados(caminhos_pdfs)

        status("Calculando idade, protocolo e demais campos automáticos...")
        nomes_arquivos = [caminho.name for caminho in caminhos_pdfs]
        dados = completar_campos_automaticos(dados, nomes_arquivos, self.config)

        status("Validando os dados extraídos...")
        avisos = self.validador.validar(dados)

        nome_base = gerar_nome_base_arquivo(dados)
        pasta_saida = Path(pasta_saida)
        pasta_saida.mkdir(parents=True, exist_ok=True)

        caminho_json = pasta_saida / f"{nome_base}_dados.json"
        caminho_docx = pasta_saida / f"{nome_base}_laudo.docx"

        status("Salvando dados extraídos em JSON (backup/auditoria)...")
        with open(caminho_json, "w", encoding="utf-8") as arquivo:
            json.dump(dados, arquivo, ensure_ascii=False, indent=2)

        status("Preenchendo o laudo em Word...")
        self.gerador_word.gerar_laudo(dados, caminho_docx)

        status("Laudo gerado com sucesso!")
        return {
            "caminho_docx": str(caminho_docx),
            "caminho_json": str(caminho_json),
            "avisos": avisos,
            "dados": dados,
            "arquivos_processados": nomes_arquivos,
        }
