"""
=================================================================
 Teste de Geração de Laudo — Topcon Maestro2 3D Wide Report
 (SEM PDF real, SEM API — apenas dados fictícios)
=================================================================
"""

from __future__ import annotations

import sys
import json
from pathlib import Path
from datetime import datetime

# Descobre o caminho da pasta principal
PASTA_PROGRAMA = Path(__file__).parent.parent

# Adiciona a pasta principal ao caminho
sys.path.insert(0, str(PASTA_PROGRAMA))

# Importa
from nucleo_laudo import WordGenerator


def main():
    """Teste de geração de laudo com dados fictícios."""

    print("=" * 65)
    print(" Teste de Geração de Laudo — Topcon Maestro2 3D Wide Report")
    print("=" * 65)
    print()

    # Caminhos
    caminho_template = PASTA_PROGRAMA / "template_laudo_topcon_maestro2_wide.docx"
    caminho_dados = Path(__file__).parent / "dados_exemplo_topcon_maestro2.json"
    pasta_saida = PASTA_PROGRAMA / "laudos_gerados"

    print(f"✓ Pasta do programa: {PASTA_PROGRAMA}")
    print(f"✓ Pasta de saída: {pasta_saida}")
    print()

    if not caminho_template.exists():
        print(f"✗ ERRO: Template não encontrado em {caminho_template}")
        sys.exit(1)
    print(f"✓ Template encontrado: {caminho_template.name}")

    if not caminho_dados.exists():
        print(f"✗ ERRO: Dados não encontrados em {caminho_dados}")
        sys.exit(1)
    print(f"✓ Dados fictícios encontrados: {caminho_dados.name}")
    print()

    # Carrega os dados
    print("Carregando dados fictícios...")
    with open(caminho_dados, "r", encoding="utf-8") as f:
        dados_extraidos = json.load(f)
    print(f"✓ Dados carregados: {len(dados_extraidos)} campos")
    print()

    # Gera o laudo
    print("Gerando laudo em Word...")
    try:
        gerador = WordGenerator(str(caminho_template))
        caminho_docx, caminho_json = gerador.gerar_laudo(
            dados_extraidos,
            str(pasta_saida)
        )
        print(f"✓ Laudo gerado com sucesso!")
        print()
        print(f"  Word: {caminho_docx}")
        print(f"  JSON (auditoria): {caminho_json}")
        print()
        print("=" * 65)
        print(" ✓ TESTE COMPLETO — Arquivo pronto para revisão!")
        print("=" * 65)
    except Exception as e:
        print(f"✗ ERRO ao gerar laudo: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()