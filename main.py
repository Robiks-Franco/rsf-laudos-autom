"""
=================================================================
 Sistema de Automação de Laudos de Exames Oftalmológicos
 Ocular Oftalmologia - Governador Valadares/MG
 (Aplicativo de mesa — janela Tkinter)
=================================================================

Este é o aplicativo de MESA (roda no seu computador Windows/Mac/
Linux). Toda a lógica de extração/validação/geração de laudos está em
"nucleo_laudo.py" — este arquivo cuida apenas da janela/botões.

Se você quer usar de qualquer lugar (celular, tablet, outro
computador), veja "app_web.py" e o guia "README_WEB.md": é o mesmo
sistema, só que publicado como um site que você acessa pelo navegador.

IMPORTANTE: um exame de OCT completo no equipamento Zeiss Cirrus
normalmente gera VÁRIOS arquivos PDF — um para cada protocolo de
varredura (ex: Macular Cube - Macular Thickness, Macular Cube -
Ganglion Cell, Optic Disc Cube - ONH and RNFL, HD 5 Line Raster OD/OS).
Por isso, este programa permite selecionar vários PDFs de uma vez,
todos referentes ao mesmo exame/paciente, e combina os dados de todos
eles em um único laudo.

Instale as dependências com: pip install -r requirements.txt
=================================================================
"""

from __future__ import annotations

import os
import sys
import threading
import traceback
from pathlib import Path

from nucleo_laudo import (
    ErroLaudoOCT,
    GeradorLaudoOCT,
)

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, simpledialog, ttk
except ImportError:
    # O Tkinter não vem por padrão em algumas instalações mínimas de
    # Linux (precisa de "sudo apt install python3-tk"). Deixamos o
    # import falhar de forma silenciosa aqui e avisamos o usuário só
    # se ele tentar realmente abrir a janela (função main() abaixo).
    tk = None


# ===================================================================
# INTERFACE GRÁFICA (Tkinter)
# Interface simples, sem terminal: selecionar PDFs -> gerar laudo,
# com barra de progresso e área de status.
# ===================================================================
class InterfaceGrafica:
    """Janela principal do sistema — pensada para uso sem terminal."""

    def __init__(self, raiz: tk.Tk):
        self.raiz = raiz
        self.raiz.title("Automação de Laudos OCT — Ocular Oftalmologia")
        self.raiz.geometry("640x520")
        self.raiz.minsize(580, 440)

        pasta_programa = Path(__file__).resolve().parent
        self.caminho_config = pasta_programa / "config_oct.json"
        self.caminho_template = pasta_programa / "template_laudo.docx"

        # Os laudos gerados ficam salvos em uma pasta fixa dentro de
        # "Documentos", independente de onde o programa esteja
        # instalado/rodando — assim fica sempre fácil de encontrar.
        self.pasta_saida = Path.home() / "Documents" / "LaudosOCT" / "laudos_gerados"

        self.caminhos_pdf_selecionados = []
        self.gerador = None  # criado sob demanda, após informar a chave de API

        self._montar_layout()

    # ---------------- Construção da interface ----------------
    def _montar_layout(self):
        fonte_titulo = ("Segoe UI", 14, "bold")
        fonte_normal = ("Segoe UI", 10)

        quadro_topo = tk.Frame(self.raiz, padx=16, pady=12)
        quadro_topo.pack(fill="x")

        tk.Label(
            quadro_topo,
            text="Automação de Laudos de OCT",
            font=fonte_titulo,
        ).pack(anchor="w")
        tk.Label(
            quadro_topo,
            text="Selecione todos os PDFs do exame (um exame Zeiss Cirrus geralmente\n"
                 "gera vários arquivos) e gere o laudo em Word automaticamente.",
            font=fonte_normal,
            fg="#444444",
            justify="left",
        ).pack(anchor="w", pady=(2, 0))

        quadro_arquivo = tk.Frame(self.raiz, padx=16, pady=8)
        quadro_arquivo.pack(fill="x")

        self.botao_selecionar = tk.Button(
            quadro_arquivo,
            text="1. Selecionar PDFs do Exame",
            command=self.selecionar_pdfs,
            width=26,
        )
        self.botao_selecionar.pack(side="left")

        self.rotulo_arquivo = tk.Label(
            quadro_arquivo, text="Nenhum arquivo selecionado", font=fonte_normal, fg="#666666"
        )
        self.rotulo_arquivo.pack(side="left", padx=10)

        quadro_acao = tk.Frame(self.raiz, padx=16, pady=8)
        quadro_acao.pack(fill="x")

        self.botao_gerar = tk.Button(
            quadro_acao,
            text="2. Gerar Laudo",
            command=self.iniciar_geracao,
            width=26,
            bg="#2f6fdb",
            fg="white",
        )
        self.botao_gerar.pack(side="left")

        self.barra_progresso = ttk.Progressbar(
            quadro_acao, mode="indeterminate", length=260
        )
        self.barra_progresso.pack(side="left", padx=10)

        quadro_status = tk.Frame(self.raiz, padx=16, pady=8)
        quadro_status.pack(fill="both", expand=True)

        tk.Label(quadro_status, text="Andamento:", font=fonte_normal).pack(anchor="w")

        quadro_texto = tk.Frame(quadro_status)
        quadro_texto.pack(fill="both", expand=True)

        barra_rolagem = tk.Scrollbar(quadro_texto)
        barra_rolagem.pack(side="right", fill="y")

        self.texto_status = tk.Text(
            quadro_texto,
            height=14,
            state="disabled",
            wrap="word",
            yscrollcommand=barra_rolagem.set,
            font=("Consolas", 9),
        )
        self.texto_status.pack(side="left", fill="both", expand=True)
        barra_rolagem.config(command=self.texto_status.yview)

        self._log(f"Pasta do programa: {Path(__file__).resolve().parent}")
        self._log(f"Os laudos gerados serão salvos em: {self.pasta_saida}")
        self._log("Pronto. Selecione os PDFs do exame para começar.")

    # ---------------- Ações da interface ----------------
    def selecionar_pdfs(self):
        caminhos = filedialog.askopenfilenames(
            title="Selecione TODOS os PDFs do exame (pode marcar vários arquivos)",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
        )
        if not caminhos:
            return
        self.caminhos_pdf_selecionados = list(caminhos)
        quantidade = len(self.caminhos_pdf_selecionados)
        self.rotulo_arquivo.config(
            text=f"{quantidade} arquivo(s) selecionado(s)", fg="#000000"
        )
        self._log(f"{quantidade} arquivo(s) selecionado(s):")
        for caminho in self.caminhos_pdf_selecionados:
            self._log(f"   • {Path(caminho).name}")

    def _obter_chave_api(self) -> str:
        """Busca a chave da API no ambiente ou solicita ao usuário via janela."""
        chave = os.environ.get("ANTHROPIC_API_KEY")
        if chave:
            return chave
        return simpledialog.askstring(
            "Chave da API da Anthropic",
            "Não foi encontrada a variável de ambiente ANTHROPIC_API_KEY.\n"
            "Informe sua chave da API da Anthropic (Claude):",
            show="*",
            parent=self.raiz,
        )

    def iniciar_geracao(self):
        if not self.caminhos_pdf_selecionados:
            messagebox.showwarning(
                "Nenhum arquivo selecionado",
                "Selecione todos os PDFs do exame antes de gerar o laudo.",
            )
            return

        chave_api = self._obter_chave_api()
        if not chave_api:
            messagebox.showerror(
                "Chave de API necessária",
                "É necessária uma chave de API válida da Anthropic para continuar.",
            )
            return

        self.botao_gerar.config(state="disabled")
        self.botao_selecionar.config(state="disabled")
        self.barra_progresso.start(12)

        thread = threading.Thread(
            target=self._gerar_laudo_em_segundo_plano, args=(chave_api,), daemon=True
        )
        thread.start()

    # ---------------- Execução em segundo plano ----------------
    def _gerar_laudo_em_segundo_plano(self, chave_api: str):
        """
        Executa a geração do laudo em uma thread separada, para que a
        janela não trave (fique "sem resposta") durante a chamada à API.
        """
        try:
            if self.gerador is None:
                self.gerador = GeradorLaudoOCT(
                    str(self.caminho_config), str(self.caminho_template), chave_api=chave_api
                )
            resultado = self.gerador.processar_exame(
                self.caminhos_pdf_selecionados,
                str(self.pasta_saida),
                callback_status=self._log_thread_segura,
            )
            self.raiz.after(0, self._ao_finalizar_com_sucesso, resultado)
        except ErroLaudoOCT as erro:
            self.raiz.after(0, self._ao_finalizar_com_erro, str(erro))
        except Exception as erro:
            detalhe = f"Erro inesperado: {erro}\n\n{traceback.format_exc()}"
            self.raiz.after(0, self._ao_finalizar_com_erro, detalhe)

    def _log_thread_segura(self, mensagem: str):
        """Agenda a atualização do log na thread principal do Tkinter."""
        self.raiz.after(0, self._log, mensagem)

    def _log(self, mensagem: str):
        self.texto_status.config(state="normal")
        self.texto_status.insert("end", f"• {mensagem}\n")
        self.texto_status.see("end")
        self.texto_status.config(state="disabled")

    def _restaurar_botoes(self):
        self.barra_progresso.stop()
        self.botao_gerar.config(state="normal")
        self.botao_selecionar.config(state="normal")

    def _ao_finalizar_com_sucesso(self, resultado: dict):
        self._restaurar_botoes()
        texto_avisos = ""
        if resultado["avisos"]:
            texto_avisos = "\n\nAvisos:\n- " + "\n- ".join(resultado["avisos"])
            for aviso in resultado["avisos"]:
                self._log(f"AVISO: {aviso}")

        self._log("Laudo gerado com sucesso!")
        messagebox.showinfo(
            "Laudo gerado com sucesso",
            "O laudo foi gerado com sucesso!\n\n"
            f"Word: {resultado['caminho_docx']}\n"
            f"JSON (auditoria): {resultado['caminho_json']}"
            f"{texto_avisos}",
        )

    def _ao_finalizar_com_erro(self, mensagem: str):
        self._restaurar_botoes()
        self._log(f"ERRO: {mensagem}")
        messagebox.showerror("Erro ao gerar laudo", mensagem)


def main():
    """Ponto de entrada do programa: abre a janela principal."""
    if tk is None:
        print(
            "ERRO: o Tkinter (interface gráfica) não está instalado neste Python.\n"
            "No Linux, instale com: sudo apt install python3-tk\n"
            "No Windows/macOS, reinstale o Python marcando o componente Tkinter/tcl-tk."
        )
        sys.exit(1)
    raiz = tk.Tk()
    InterfaceGrafica(raiz)
    raiz.mainloop()


if __name__ == "__main__":
    main()
