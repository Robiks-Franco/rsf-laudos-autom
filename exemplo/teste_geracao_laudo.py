"""
=================================================================
 Teste de ponta a ponta (SEM precisar de PDFs reais nem de chave de API)
=================================================================

Este script demonstra o funcionamento do sistema usando dados
fictícios de um exame de OCT (arquivo dados_exemplo_oct.json), pulando
apenas a etapa de extração via Claude API (que exige PDFs reais e uma
chave de API paga).

Fluxo demonstrado aqui:
    dados fictícios (JSON)
        -> completar_campos_automaticos (idade, protocolo, médico)
        -> ValidadorDados
        -> WordGenerator
        -> laudo .docx

Rode com:
    python exemplo/teste_geracao_laudo.py

(a partir da pasta principal do programa, onde está o main.py)
=================================================================
"""

import sys
import json
from pathlib import Path

# Garante que conseguimos importar nucleo_laudo.py, que está uma
# pasta acima deste script (pasta principal do programa). Repare que
# importamos de "nucleo_laudo", não de "main" — main.py é só a janela
# Tkinter; toda a lógica (usada também pelo app web) está no núcleo.
PASTA_EXEMPLO = Path(__file__).resolve().parent
PASTA_PRINCIPAL = PASTA_EXEMPLO.parent
sys.path.insert(0, str(PASTA_PRINCIPAL))

from nucleo_laudo import (  # noqa: E402
    ExamConfig,
    ValidadorDados,
    WordGenerator,
    ErroLaudoOCT,
    completar_campos_automaticos,
)

# Nomes de arquivos fictícios, só para simular a detecção automática
# de protocolo (ver função detectar_protocolo em main.py) — em um uso
# real, esses seriam os nomes dos PDFs exportados pelo Zeiss Cirrus.
NOMES_ARQUIVOS_SIMULADOS = [
    "PACIENTE_EXEMPLO_Macular Cube 512x128_OU_Macular Thickness OU Analysis.pdf",
    "PACIENTE_EXEMPLO_Macular Cube 512x128_OU_Ganglion Cell OU Analysis.pdf",
    "PACIENTE_EXEMPLO_Optic Disc Cube 200x200_OU_ONH and RNFL OU Analysis.pdf",
]


def main():
    print("=" * 70)
    print("TESTE DE PONTA A PONTA — Sistema de Laudos de OCT (dados fictícios)")
    print("=" * 70)

    caminho_config = PASTA_PRINCIPAL / "config_oct.json"
    caminho_template = PASTA_PRINCIPAL / "template_laudo.docx"
    caminho_dados_exemplo = PASTA_EXEMPLO / "dados_exemplo_oct.json"

    caminho_docx_saida = PASTA_EXEMPLO / "laudo_exemplo.docx"
    caminho_json_saida = PASTA_EXEMPLO / "laudo_exemplo_dados.json"

    try:
        print(f"\n1) Carregando configuração: {caminho_config.name}")
        config = ExamConfig(str(caminho_config))
        print(f"   Tipo de exame: {config.nome_exibicao}")
        print(f"   Total de campos definidos: {len(config.campos)}")

        print(f"\n2) Carregando dados fictícios: {caminho_dados_exemplo.name}")
        with open(caminho_dados_exemplo, "r", encoding="utf-8") as arquivo:
            dados = json.load(arquivo)
        # Remove a chave de observação (não é um campo do laudo)
        dados.pop("_observacao", None)
        print(f"   Paciente (exemplo): {dados.get('nome_paciente')}")

        print("\n3) Completando campos automáticos (idade, protocolo, médico)...")
        dados = completar_campos_automaticos(dados, NOMES_ARQUIVOS_SIMULADOS, config)
        print(f"   Idade calculada: {dados.get('idade')}")
        print(f"   Protocolo detectado: {dados.get('protocolo')}")
        print(f"   Médico responsável: {dados.get('nome_medico')} ({dados.get('crm_medico')})")

        print("\n4) Validando os dados (ValidadorDados)...")
        validador = ValidadorDados(config)
        avisos = validador.validar(dados)
        if avisos:
            print("   Avisos encontrados:")
            for aviso in avisos:
                print(f"   - {aviso}")
        else:
            print("   Nenhum aviso. Todos os campos obrigatórios estão presentes.")

        print(f"\n5) Gerando o laudo em Word a partir de: {caminho_template.name}")
        gerador_word = WordGenerator(str(caminho_template))
        gerador_word.gerar_laudo(dados, str(caminho_docx_saida))

        print("\n6) Salvando cópia dos dados em JSON (auditoria)...")
        with open(caminho_json_saida, "w", encoding="utf-8") as arquivo:
            json.dump(dados, arquivo, ensure_ascii=False, indent=2)

        print("\n" + "=" * 70)
        print("SUCESSO! O sistema está funcionando corretamente.")
        print(f"Laudo gerado: {caminho_docx_saida}")
        print(f"JSON gerado:  {caminho_json_saida}")
        print("=" * 70)
        print(
            "\nObservação: este teste NÃO chamou a Claude API — ele usou dados "
            "fictícios prontos apenas para validar a geração do laudo Word. Para "
            "testar a extração real a partir de PDFs de um exame, use a interface "
            "gráfica (main.py) com uma chave de API válida, selecionando todos os "
            "PDFs do exame de uma vez."
        )

    except ErroLaudoOCT as erro:
        print(f"\nERRO: {erro}")
        sys.exit(1)
    except FileNotFoundError as erro:
        print(f"\nERRO: Arquivo não encontrado: {erro}")
        sys.exit(1)


if __name__ == "__main__":
    main()
