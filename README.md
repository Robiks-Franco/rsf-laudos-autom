# Automação de Laudos de Exames Oftalmológicos — Ocular Oftalmologia

Sistema para extrair automaticamente os dados de exames oftalmológicos
(OCT, tomografia de córnea, campo visual, e outros que forem adicionados) usando a
API da Claude (Anthropic) e gerar o laudo já preenchido em Word
(.docx), junto com um arquivo JSON de auditoria/backup dos dados
extraídos.

**Equipamentos suportados atualmente:**

| Equipamento | Tipo de exame | Chave interna | Status |
|---|---|---|---|
| Zeiss Cirrus HD-OCT | OCT (retina/nervo óptico) | `zeiss` | ✅ Disponível |
| Nidek RS-3000 Advance (RetinaScan Advance) | OCT (retina/nervo óptico) | `nidek` | ✅ Disponível |
| Pentacam® (Oculus) | Tomografia de córnea (Scheimpflug) | `pentacam` | ✅ Disponível |
| Octopus 600 (Haag-Streit) | Campo Visual (perimetria computadorizada) | `campo_visual` | ✅ Disponível |
| Topcon Triton (DRI OCT Triton / Triton Plus) | OCT Swept-Source + retinografia | `topcon` | 🕓 Planejado — aguardando PDFs de exemplo para configurar |

Ao abrir o programa (de mesa ou web), você escolhe o equipamento usado
naquele exame em um seletor — o restante do fluxo é idêntico. Cada
equipamento tem seu próprio arquivo de configuração (`config_*.json`)
e seu próprio modelo de laudo (`template_laudo_*.docx`); veja a seção
9 para saber como adicionar um novo (seja um novo equipamento de OCT,
seja um exame completamente diferente, como já foi feito com o
Pentacam e com o Campo Visual).

**Um exame completo normalmente é composto por vários PDFs** — cada
equipamento exporta um arquivo separado por protocolo de varredura ou
por olho (no Zeiss Cirrus: Macular Cube - Macular Thickness, Macular
Cube - Ganglion Cell, Optic Disc Cube - ONH and RNFL, HD 5 Line Raster
OD e OS; no Nidek RS-3000: Macula Map OD/OE, Disc Map, Macula Radial,
Macula Line, Angio OCT OD/OE; no Pentacam: um PDF para o olho direito
e outro para o olho esquerdo; no Octopus 600/Campo Visual: um PDF para
o olho direito e outro para o olho esquerdo). O programa permite
selecionar todos esses PDFs de uma vez e combina os dados de todos
eles em um único laudo, com os parâmetros de OD e OE lado a lado.

Não é necessário usar terminal no dia a dia: o programa abre em uma
janela simples, com um seletor de equipamento e botões para selecionar
os PDFs e gerar o laudo.

> **Quer usar de qualquer lugar (celular, tablet, outro computador),
> sem precisar instalar Python?** Este mesmo sistema também existe
> como aplicativo web (`app_web.py`) — veja **README_WEB.md** para o
> passo a passo de como publicá-lo na internet.

---

## 1. Estrutura dos arquivos

```
nucleo_laudo.py       -> TODA a lógica (extração, validação, geração) — sem interface
                          inclui o registro EQUIPAMENTOS_SUPORTADOS (ver seção 9)
main.py                 -> aplicativo de mesa (janela Tkinter), usa nucleo_laudo.py
app_web.py                -> aplicativo web (FastAPI), usa nucleo_laudo.py — ver README_WEB.md
web/index.html               -> página do aplicativo web (upload + download do laudo)
config_oct.json                 -> configuração dos campos — Zeiss Cirrus HD-OCT
template_laudo.docx               -> modelo do laudo em Word — Zeiss Cirrus HD-OCT
config_nidek.json                   -> configuração dos campos — Nidek RS-3000 Advance
template_laudo_nidek.docx             -> modelo do laudo em Word — Nidek RS-3000 Advance
config_pentacam.json                    -> configuração dos campos — Pentacam® (Oculus)
template_laudo_pentacam.docx              -> modelo do laudo em Word — Pentacam® (Oculus)
config_campo_visual.json                    -> configuração dos campos — Octopus 600 (Campo Visual)
template_laudo_campo_visual.docx              -> modelo do laudo em Word — Octopus 600 (Campo Visual)
requirements.txt                        -> dependências do aplicativo de mesa
requirements_web.txt                      -> dependências do aplicativo web (inclui as de cima + FastAPI)
exemplo/                                    -> dados fictícios e scripts de teste sem precisar de PDF/API
```

`main.py` e `app_web.py` são só duas "portas de entrada" diferentes
para o mesmo motor (`nucleo_laudo.py`) — qualquer melhoria feita no
núcleo (um novo campo, uma correção) vale para os dois automaticamente.

Os laudos gerados **não** ficam na pasta do programa: eles são salvos
automaticamente em `Documentos\LaudosOCT\laudos_gerados`, na pasta
"Documentos" do seu usuário do Windows (essa pasta é criada
automaticamente na primeira vez que você gerar um laudo). Cada laudo é
nomeado a partir do nome do paciente e da data do exame (ex:
`Antonio_Carlos_Pereira_07082026_laudo.docx`), já que agora ele pode
vir de vários PDFs, não de um único arquivo.

## 2. Instalação

Pré-requisitos: Python 3.9 ou superior instalado no computador
(em Windows, marque a opção "Add Python to PATH" durante a instalação).

Abra um terminal na pasta do programa e instale as dependências com:

```
pip install -r requirements.txt
```

> Observação: em algumas instalações de Linux, o Tkinter (interface
> gráfica) não vem junto com o Python e precisa ser instalado à parte,
> por exemplo: `sudo apt install python3-tk`. No Windows e no macOS,
> normalmente já vem incluído.

## 3. Chave da API da Anthropic (Claude)

O sistema usa a API da Claude para "ler" o exame (texto e imagem) e
extrair os dados. Você precisa de uma chave de API própria:

1. Crie uma conta em https://console.anthropic.com
2. Gere uma chave de API (API Key) na seção correspondente.
3. Você pode informar essa chave de duas formas:
   - **Variável de ambiente** `ANTHROPIC_API_KEY` (recomendado, não
     precisa digitar a chave toda vez); ou
   - **Digitando na hora**, quando o programa pedir, na primeira vez
     que você clicar em "Gerar Laudo".

A chave de API é paga por uso (conforme a tabela de preços da
Anthropic) — cada laudo gerado consome uma quantidade de créditos
proporcional ao número de PDFs e páginas enviados.

## 4. Como usar

1. Rode o programa (duplo clique ou, em terminal: `python main.py`).
2. No seletor **"Equipamento"**, escolha o equipamento usado naquele
   exame (ex: "Zeiss Cirrus HD-OCT", "Nidek RS-3000 Advance", "Pentacam®
   (Oculus)" ou "Octopus 600 (Haag-Streit) — Campo Visual").
3. Clique em **"1. Selecionar PDFs do Exame"** e marque **todos** os
   PDFs daquele exame (você pode selecionar vários arquivos de uma vez
   na mesma janela — segure Ctrl e clique em cada um, ou Shift para
   selecionar um intervalo).
4. Clique em **"2. Gerar Laudo"**.
5. Aguarde a barra de progresso (a extração pode levar alguns
   segundos a mais de um minuto, dependendo da quantidade de PDFs).
6. Ao final, o programa mostra onde os arquivos foram salvos, dentro de
   `Documentos\LaudosOCT\laudos_gerados`:
   - `<paciente>_<data>_laudo.docx` — o laudo pronto em Word.
   - `<paciente>_<data>_dados.json` — os dados extraídos, em formato
     JSON, para conferência e auditoria.
7. **Sempre revise o laudo gerado antes de assinar/entregar** — o
   sistema apenas transcreve os dados visíveis no exame; a
   responsabilidade pela interpretação clínica, pela conclusão
   diagnóstica e pela conferência dos dados é do médico. Campos que a
   IA não conseguiu identificar aparecem como um travessão "—" no
   laudo, para ficar claro que precisam ser preenchidos manualmente.

## 5. Testando sem PDFs reais (dados fictícios)

Para validar a instalação sem precisar de um exame real nem de chave
de API, use os scripts de teste com dados fictícios — um para cada
equipamento:

```
python exemplo/teste_geracao_laudo.py                  (Zeiss Cirrus HD-OCT)
python exemplo/teste_geracao_laudo_nidek.py             (Nidek RS-3000 Advance)
python exemplo/teste_geracao_laudo_pentacam.py           (Pentacam® Oculus)
python exemplo/teste_geracao_laudo_campo_visual.py        (Octopus 600 — Campo Visual)
```

> **Importante:** esses scripts precisam estar **dentro da pasta
> `exemplo/`** para funcionar — eles descobrem o caminho da pasta
> principal do programa (`config_*.json`, `template_laudo_*.docx`) a
> partir da própria localização do arquivo (`Path(__file__).parent`).
> Se algum desses scripts (ou o `dados_exemplo_*.json` correspondente)
> acabar sendo enviado para a raiz do repositório por engano, mova-o
> para dentro de `exemplo/` antes de rodar, ou ele não vai encontrar a
> configuração/template corretos.

Cada script pega os dados de `exemplo/dados_exemplo_oct.json` (ou
`dados_exemplo_nidek.json`/`dados_exemplo_pentacam.json`/`dados_exemplo_campo_visual.json`),
calcula os campos automáticos (idade, protocolo quando aplicável),
valida e gera o laudo `.docx` + `.json` correspondente na pasta
`exemplo/` — útil para testar o template Word sem chamar a API.

## 6. Como funciona (visão geral)

```
PDFs do exame (1 a 5 arquivos, um por protocolo de varredura)
   │
   ▼
PDFExtractor  ──►  extrai texto + converte páginas em imagem de CADA pdf
   │                  ▼
   │            Claude API — UMA chamada com todos os PDFs juntos,
   │            cada um identificado pelo nome do arquivo
   ▼
dados extraídos (JSON em memória, com campos _od e _os lado a lado)
   │
   ▼
completar_campos_automaticos ──► idade, protocolo, médico responsável,
   │                              datas normalizadas (calculado pelo
   │                              programa, não pela IA — mais confiável)
   ▼
ValidadorDados ──► confere campos obrigatórios e valores esperados
   │
   ▼
WordGenerator  ──► preenche template_laudo.docx
   │
   ▼
laudo_final.docx  +  dados.json (auditoria)
```

O que cada campo extrai está definido em `config_oct.json` — nada
disso está "escondido" dentro do código. Campos como idade, protocolo
do exame, nome/CRM do médico e data de geração **não** são pedidos à
IA: são calculados pelo próprio programa (mais confiável e sem custo
de API), na função `completar_campos_automaticos` em `main.py`.

## 7. Estrutura dos dados do laudo (seções clínicas)

Esta seção descreve `config_oct.json` (Zeiss Cirrus). O Nidek RS-3000
Advance segue a mesma lógica em `config_nidek.json`, com uma seção a
mais (Angiografia OCT/AngioScan) e nomenclatura de campos própria do
equipamento (ex: SQI/SSI em vez de Signal Strength) — abra o arquivo
para ver a lista completa de campos. O Pentacam (`config_pentacam.json`)
e o Campo Visual/Octopus 600 (`config_campo_visual.json`) seguem a
mesma estrutura geral, mas com seções e campos próprios do tipo de
exame (ceratometria/paquimetria/índices de ectasia no Pentacam;
índices de confiabilidade e índices globais MD/sLV/DD/LD/MS no Campo
Visual).

Os campos em `config_oct.json` seguem a estrutura de um laudo real de
OCT (Zeiss Cirrus), organizada em 3 seções, com valores separados por
olho (`_od` / `_os`):

- **Mácula/Retina** (`categoria: "macula_retina"`) — CST, volume
  macular, espessura média do cubo. Vem do PDF "Macular Thickness OU
  Analysis".
- **Nervo Óptico/RNFL** (`categoria: "nervo_optico"`) — RNFL global e
  por quadrante, simetria, área do disco/rima, relação C/D, volume da
  escavação. Vem do PDF "ONH and RNFL OU Analysis".
- **Complexo de Células Ganglionares** (`categoria: "ganglion_cell"`)
  — espessura média/mínima GCL-IPL. Vem do PDF "Ganglion Cell OU
  Analysis".

Cada campo tem um atributo `"origem"` indicando de qual protocolo
Zeiss ele deveria vir — isso ajuda a IA a não misturar dados de
arquivos diferentes. Se um PDF de um determinado protocolo não for
selecionado, os campos daquela seção simplesmente ficam em branco
("—") no laudo — nada quebra.

O **médico responsável** (nome e CRM que aparecem na assinatura) fica
definido centralmente em `config_oct.json`, na chave
`"medico_responsavel"` — edite esse arquivo se precisar trocar.

## 8. Personalizando o template do laudo

Abra `template_laudo.docx` normalmente no Word (ou LibreOffice). Os
textos entre chaves duplas, como `{{nome_paciente}}`, são os
marcadores que o programa substitui automaticamente. Você pode:

- Mudar cores, fontes, logotipo, cabeçalho e rodapé livremente.
- Mover os marcadores de posição no documento.
- Editar o texto padrão da conclusão ("Correlacionar a outros achados
  clínicos.") — esse texto é fixo no template (não é gerado pela IA de
  propósito, já que a conclusão diagnóstica é responsabilidade do
  médico); edite-o manualmente em cada laudo gerado.
- **Não altere o texto dentro das chaves** (ex: `{{nome_paciente}}`),
  pois é exatamente esse texto que o programa procura para substituir.
  Se precisar editar o marcador, edite-o inteiro de uma vez (apague e
  digite novamente `{{id_do_campo}}`), para evitar que o Word quebre o
  texto em pedaços que o programa não reconheça.

A lista de todos os marcadores disponíveis está em `config_oct.json`,
no campo `"id"` de cada item de `"campos"`, mais os campos calculados
automaticamente listados em `"campos_calculados_automaticamente"`
(`idade`, `protocolo`, `nome_medico`, `crm_medico`,
`data_geracao_laudo`).

## 9. Adicionando novos equipamentos ou tipos de exame

O sistema foi desenhado para ser reaproveitado em outros equipamentos
de OCT (ex: Topcon Triton) ou em exames completamente diferentes de
OCT — foi exatamente assim que o Pentacam® (tomografia de córnea) e o
Campo Visual/Octopus 600 (perimetria computadorizada) foram
adicionados, sem alterar uma linha da lógica em `nucleo_laudo.py`. O
mesmo caminho serve para Biometria, ou qualquer outro exame que gere
PDF. Tudo passa por um único registro central,
`EQUIPAMENTOS_SUPORTADOS`, no topo de `nucleo_laudo.py`:

```python
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
```

Para adicionar um equipamento novo (ex: Topcon Triton):

1. Copie um `config_*.json` existente (ex: `config_nidek.json`) para
   `config_topcon.json` e ajuste a lista `"campos"` com os dados que
   você quer extrair daquele equipamento (rótulo, tipo, se é
   obrigatório, unidade, etc.), além de `"protocolos_por_palavra_chave"`
   (como reconhecer o protocolo pelo nome dos arquivos PDF exportados).
2. Copie um `template_laudo_*.docx` existente para
   `template_laudo_topcon.docx` e ajuste o layout e os marcadores
   `{{...}}` para os novos campos.
3. Adicione uma linha nova em `EQUIPAMENTOS_SUPORTADOS` (em
   `nucleo_laudo.py`), com uma chave curta (ex: `"topcon"`) apontando
   para os dois arquivos acima.

Pronto — o seletor de equipamento passa a mostrar essa opção
automaticamente, tanto no aplicativo de mesa (`main.py`) quanto no
aplicativo web (`app_web.py` + `web/index.html`), sem mais nenhuma
alteração de código. É recomendável também criar um
`exemplo/dados_exemplo_<equipamento>.json` e um
`exemplo/teste_geracao_laudo_<equipamento>.py` (copiando os do Nidek
ou do Pentacam como modelo) para validar o template sem precisar de
PDFs reais — **e lembre de enviar os dois para dentro da pasta
`exemplo/` no GitHub, não para a raiz do repositório** (veja o aviso
na seção 5).

Todas as classes de `nucleo_laudo.py` (`ExamConfig`, `PDFExtractor`,
`ValidadorDados`, `WordGenerator`, `GeradorLaudoOCT`) já recebem o
caminho de configuração/template como parâmetro — nenhuma delas é
específica de um equipamento, e todas já suportam receber vários PDFs
de um mesmo exame. Isso vale tanto para o aplicativo de mesa quanto
para o web.

> **Topcon Triton (DRI OCT Triton / Triton Plus):** planejado, mas
> ainda não configurado — falta receber PDFs de exemplo desse
> equipamento para mapear os campos corretamente (o Triton, por ser
> swept-source e ter retinógrafo integrado, deve trazer campos
> adicionais como imagem de fundo de olho e OCT de coroide/EDI).

## 10. Mensagens de erro mais comuns

| Mensagem | Causa provável | O que fazer |
|---|---|---|
| "Chave da API da Anthropic não encontrada" | Nenhuma chave foi informada | Defina `ANTHROPIC_API_KEY` ou digite a chave quando solicitado |
| "Arquivo de configuração não encontrado" | `config_oct.json` foi movido/apagado | Verifique se ele está na mesma pasta de `main.py` |
| "Não foi possível identificar os seguintes campos obrigatórios" | Os PDFs não traziam esses dados de forma legível (nome, data de nascimento ou data do exame), ou a extração falhou | Confira os PDFs (qualidade da digitalização) ou preencha manualmente no laudo gerado |
| "A resposta da IA não contém um JSON reconhecível" | Falha pontual da API ou PDFs muito ruins | Tente gerar novamente |
| "Não foi possível salvar..." | O arquivo `.docx` de saída está aberto no Word | Feche o arquivo e gere novamente |
| Campos de uma seção aparecem todos como "—" | O PDF daquele protocolo (ex: Ganglion Cell) não foi selecionado | Selecione também esse PDF na próxima geração |

Se os números extraídos parecerem imprecisos (ex: confundir um dígito
em um mapa colorido), aumente a resolução das imagens enviadas à IA
editando `"dpi_imagem"` em `config_oct.json` (padrão: 200) — isso deixa
os números mais legíveis, ao custo de mensagens um pouco maiores/mais
lentas.

## 11. Privacidade dos dados dos pacientes

Os PDFs do exame (texto e imagens das páginas) são enviados à API da
Anthropic para a extração dos dados. Consulte a política de
privacidade e os termos de uso de dados da Anthropic antes de usar
este sistema com dados reais de pacientes, e avalie a necessidade de
termo de consentimento/anonimização de acordo com a LGPD e as normas
do seu conselho profissional.
