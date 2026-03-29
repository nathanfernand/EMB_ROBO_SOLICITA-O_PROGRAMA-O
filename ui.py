# =============================================================================
# ui.py — Interface gráfica Tkinter do Robô SAP
# =============================================================================

import os
from tkinter import Button, Label, Tk, mainloop, messagebox

from sap_connection import conectar_sap
from orchestrator import Orchestrator

_EXCEL_PATH = r"C:\base\ZLOLMM025.xlsx"


def criar_janela() -> None:
    """
    Cria e exibe a janela principal do robô.

    Componentes:
        - Label de status dinâmico (aguardando / em execução / concluído)
        - 3 Labels de resumo (total / sucesso / erro) — limpas a cada execução
        - Botão "Iniciar" desabilitado durante o processamento e reabilitado
          ao terminar, permitindo nova demanda sem reabrir o programa
    """
    window = Tk()
    window.title("Robô SAP - ZLOLMM025")
    window.geometry("430x195")
    window.resizable(False, False)

    # --- Label de status dinâmico ---
    lbl_status = Label(
        window,
        text="Pressione 'Iniciar' para processar a planilha.",
        font=("Segoe UI", 10),
        fg="gray",
    )
    lbl_status.pack(pady=(14, 2))

    # --- Labels de resumo ---
    lbl_total   = Label(window, text="", font=("Segoe UI", 11))
    lbl_sucesso = Label(window, text="", font=("Segoe UI", 11), fg="green")
    lbl_erro    = Label(window, text="", font=("Segoe UI", 11), fg="red")
    lbl_total.pack(pady=2)
    lbl_sucesso.pack(pady=2)
    lbl_erro.pack(pady=2)

    def iniciar() -> None:
        """
        Callback do botão 'Iniciar / Iniciar nova demanda'.

        Fluxo:
            1. Desabilita o botão e exibe estado 'em execução'
            2. Limpa os resultados da execução anterior
            3. Força atualização visual da janela antes da chamada bloqueante
            4. Conecta ao SAP, instancia o Orchestrator e dispara fazer_login()
            5. Exibe o resumo e reabilita o botão
            6. Agenda a abertura do Excel 3 s após o término
        """
        botao.config(state="disabled", text="Processando...")
        lbl_status.config(text="Robô em execução, aguarde...", fg="blue")
        lbl_total.config(text="")
        lbl_sucesso.config(text="")
        lbl_erro.config(text="")
        window.update()

        try:
            session = conectar_sap()
            if session is None:
                raise RuntimeError("Não foi possível conectar ao SAP.")
            resultado = Orchestrator(session).fazer_login()
            total, sucesso, erros = resultado if resultado else (0, 0, 0)
        except Exception as e:
            messagebox.showerror("Erro", str(e))
            total = sucesso = erros = 0

        lbl_status.config(text="✔ Processamento concluído!", fg="black")
        lbl_total.config(text=f"Total de ordens processadas :  {total}")
        lbl_sucesso.config(text=f"Total de ordens com sucesso :  {sucesso}")
        lbl_erro.config(text=f"Total de ordens com erro        :  {erros}")

        botao.config(state="normal", text="Iniciar nova demanda")
        window.after(3000, lambda: os.startfile(_EXCEL_PATH))

    botao = Button(window, text="Iniciar", command=iniciar, width=22, height=2)
    botao.pack(pady=(6, 12))

    mainloop()
