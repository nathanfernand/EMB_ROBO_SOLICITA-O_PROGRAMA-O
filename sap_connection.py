# =============================================================================
# sap_connection.py — Conexão e sessão SAP via COM (win32com)
# =============================================================================

import subprocess
import time
from typing import Any

import win32com.client
from tkinter import messagebox

from config import SAP_GUI_PATH, SAP_SERVER_NAME


def conectar_sap() -> Any | None:
    """
    Estabelece (ou reutiliza) uma sessão SAP GUI via COM.

    Fluxo:
        1. Tenta reutilizar uma sessão SAP já aberta (GetObject).
        2. Se não houver sessão ativa, abre o SAP Logon Pad e cria
           uma nova conexão com o servidor configurado em SAP_SERVER_NAME.

    Retorna:
        Objeto de sessão SAP (GuiSession) em caso de sucesso, ou None em falha.
    """
    # ------------------------------------------------------------------
    # PASSO 1 — Tenta reutilizar sessão existente
    # ------------------------------------------------------------------
    try:
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application  = sap_gui_auto.GetScriptingEngine

        if application.Children.Count > 0:
            connection = application.Children(0)
            if connection.Children.Count > 0:
                session = connection.Children(0)
                session.findById("wnd[0]").maximize()
                print("✓ Sessão SAP existente reutilizada.")
                return session

    except Exception:
        pass  # SAP não está aberto — segue para abertura normal

    # ------------------------------------------------------------------
    # PASSO 2 — Abre o SAP e cria nova conexão
    # ------------------------------------------------------------------
    try:
        subprocess.Popen(SAP_GUI_PATH)
        time.sleep(1)

        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        application  = sap_gui_auto.GetScriptingEngine

        connection = application.OpenConnection(SAP_SERVER_NAME, True)
        time.sleep(1.5)

        session = connection.Children(0)
        session.findById("wnd[0]").maximize()
        print("✓ Nova sessão SAP aberta e conectada.")
        return session

    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao conectar ao SAP: {e}")
        print(f"Erro COM: {e}")
        return None
