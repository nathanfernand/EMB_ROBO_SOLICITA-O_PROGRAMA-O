# =============================================================================
# transactions.py — Transações SAP: ZDPQPL126 (Montador/Fabricante),
#                   ZLOLMM027 (MTS) e ZLOBMM001 (fechamento)
# =============================================================================

import csv
import os
import time
from typing import Any

import pandas as pd

from config import Status, TXT_EXPORT_PATH
from excel_manager import ExcelManager


class TransactionHandler:
    """
    Executa as transações SAP para cada linha de ordem de produção.

    Parâmetros:
        session       : Objeto GuiSession ativo do SAP GUI Scripting.
        excel_manager : Instância de ExcelManager para persistência de STATUS.
    """

    def __init__(self, session: Any, excel_manager: ExcelManager) -> None:
        self.session = session
        self.excel   = excel_manager

    # =========================================================================
    # HELPERS
    # =========================================================================

    def reset_transacao(self) -> None:
        """Retorna à tela inicial do SAP enviando o comando /N."""
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
        self.session.findById("wnd[0]").sendVKey(0)

    def _tratar_falha(
        self,
        data: pd.DataFrame,
        index: int,
        data_126: list,
        status_msg: str,
        log_msg: str,
    ) -> None:
        """
        Tratamento padronizado de falhas durante o processamento de uma linha.

        Ações (nesta ordem):
            1. Imprime a mensagem de log
            2. Grava o STATUS de erro no Excel
            3. Limpa a lista temporária data_126
            4. Reseta a transação SAP
        """
        print(log_msg)
        print(f"[FALHA] Linha={index + 1} | status={status_msg!r}")
        salvo = self.excel.atualizar_status(data, index, status_msg)
        if not salvo:
            print(f"⚠ Não foi possível persistir STATUS da linha {index + 1} no Excel.")
        data_126.clear()
        self.reset_transacao()

    def _exportar_e_ler_txt(
        self,
        data: pd.DataFrame,
        index: int,
        data_126: list,
        op_val: str,
    ) -> bool:
        """
        Etapas 2 e 3 compartilhadas por Montador e Fabricante:
            - Filtra pela operação na grid do SAP e exporta para TXT
            - Lê o TXT e extrai o valor da Linha 7, Coluna 5

        Retorna:
            bool: True se o valor foi extraído com sucesso, False em falha.
                  Em falha, o STATUS já é atualizado internamente.
        """
        # ── Passo 2: filtrar e exportar ──────────────────────────────────────
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").setCurrentCell(-1, "VORNR_OP")
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectColumn("VORNR_OP")
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarButton("&MB_FILTER")
        self.session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").text = op_val
        self.session.findById("wnd[1]/usr/ssub%_SUBSCREEN_FREESEL:SAPLSSEL:1105/ctxt%%DYN001-LOW").caretPosition = 4
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").pressToolbarContextButton("&MB_EXPORT")
        self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").selectContextMenuItem("&PC")
        self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        self.session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\base"
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ZDPQPL126.txt"
        self.session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 13
        self.session.findById("wnd[1]/tbar[0]/btn[11]").press()

        # ── Passo 3: ler o TXT ──────────────────────────────────────────────
        time.sleep(2)
        caminho = TXT_EXPORT_PATH

        try:
            if os.path.exists(caminho):
                linhas: list[list[str]] = []
                with open(caminho, "r", encoding="latin-1") as f:
                    for linha in csv.reader(f, delimiter="|"):
                        linhas.append(linha)

                if len(linhas) >= 7:
                    linha_7 = linhas[6]
                    if len(linha_7) >= 5:
                        valor = linha_7[4].strip()
                        print(f"✓ Valor extraído — Linha 7, Coluna 5: {valor}")
                        data_126.append(valor)
                        return True
                    else:
                        print("❌ Coluna 5 não existe na Linha 7")
                        self.excel.atualizar_status(data, index, Status.OPERACAO_INCORRETA)
                        self.reset_transacao()
                        return False
                else:
                    print(f"❌ Linha 7 não existe (total={len(linhas)})")
                    self.excel.atualizar_status(data, index, Status.TXT_INCOMPLETO)
                    self.reset_transacao()
                    return False
            else:
                print(f"❌ Arquivo TXT não encontrado: {caminho}")
                self.excel.atualizar_status(data, index, Status.TXT_NAO_ENCONTRADO)
                self.reset_transacao()
                return False

        except Exception as e:
            print(f"❌ Erro ao ler o arquivo TXT: {e}")
            self.excel.atualizar_status(data, index, Status.ERRO_LEITURA_TXT)
            self.reset_transacao()
            return False

    def _confirmar_popups(self, data: pd.DataFrame, index: int) -> bool:
        """
        Tenta confirmar até 10 popups de finalização no SAP.

        Retorna:
            bool: True se pelo menos 1 popup foi confirmado, False caso contrário.
        """
        confirmacoes = 0
        for tentativa in range(10):
            try:
                self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
                confirmacoes += 1
                print(f"Confirmação no popup executada ({tentativa + 1}/10)")
                time.sleep(0.3)
            except Exception:
                break
        return confirmacoes > 0

    # =========================================================================
    # TRANSAÇÃO ZDPQPL126 — MONTADOR
    # =========================================================================

    def processar_montador(self, row: pd.Series, index: int, data: pd.DataFrame) -> None:
        """
        Processa uma linha de ordem do tipo MONTADOR via ZDPQPL126 + ZLOLMM025.

        Etapas:
            1. Acessa ZDPQPL126 com a OP informada
            2. Filtra pela operação e exporta para TXT
            3. Lê o TXT e extrai a data de início (Linha 7, Col 5)
            4. Captura o valor adicional da shell SAP
            5. Navega para ZLOLMM025 (P_LINHA="M") e preenche os campos
            6. Confirma os popups de finalização
            7. Grava STATUS = "ORDEM SOLICITADA!"

        Em falha atualiza o STATUS e retorna sem relançar.
        """
        data_126: list[str] = []

        print(index, row["CARRO"], row["OP"], row["OPERAÇÃO"], row["STATUS"])
        op_val = str(row["OPERAÇÃO"]).strip().zfill(4)
        print(f"Operação formatada: '{op_val}'")

        # ── Passo 1: acessar ZDPQPL126 ────────────────────────────────────
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NZDPQPL126"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/txtSP$00003-LOW").text = row["OP"]
            self.session.findById("wnd[0]/usr/ctxt%LAYOUT").text = "/FRBARRO"
            self.session.findById("wnd[0]/usr/ctxt%LAYOUT").setFocus()
            self.session.findById("wnd[0]/usr/ctxt%LAYOUT").caretPosition = 8
            self.session.findById("wnd[0]").sendVKey(0)
        except Exception:
            self._tratar_falha(data, index, data_126, Status.ORDEM_NAO_SINC,
                               "Falha ao acessar ZDPQPL126 (Montador)")
            return

        # ── Passos 2 e 3: exportar TXT e extrair valor ──────────────────
        if not self._exportar_e_ler_txt(data, index, data_126, op_val):
            return  # STATUS já atualizado por _exportar_e_ler_txt

        # ── Passo 4: capturar valor da shell ─────────────────────────────
        valor_shell = self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").text
        print(f"✓ Valor da shell SAP: {valor_shell}")
        data_126.append(valor_shell)

        # ── Passo 5: ZLOLMM025 com linha "M" ─────────────────────────────
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NZLOLMM025"
        self.session.findById("wnd[0]").sendVKey(0)

        try:
            self.session.findById("wnd[0]/usr/ctxtS_CENTRO-LOW").text = "BOT1"
            self.session.findById("wnd[0]/usr/ctxtP_AS").text = "AS-B28"
            self.session.findById("wnd[0]/usr/ctxtP_LINHA").text = "M"
            self.session.findById("wnd[0]/usr/ctxtP_LINHA").setFocus()
            self.session.findById("wnd[0]/usr/ctxtP_LINHA").caretPosition = 1
            self.session.findById("wnd[0]").sendVKey(0)
        except Exception:
            self._tratar_falha(data, index, data_126, Status.ORDEM_CONGELADA,
                               "Falha ao preencher ZLOLMM025 (Montador — etapa navegação)")
            return

        try:
            time.sleep(1)
            self.session.findById("wnd[0]/usr/ctxtS_PROGR2-LOW").text = "0200"
            self.session.findById("wnd[0]/usr/txtP_TAKT2").text = "1"
            self.session.findById("wnd[0]/usr/ctxtP_INICA2").text = data_126[0]
            self.session.findById("wnd[0]/usr/ctxtS_ORDEM2-LOW").text = row["OP"]
            self.session.findById("wnd[0]/usr/ctxtS_ORDEM2-LOW").setFocus()
            self.session.findById("wnd[0]/usr/ctxtS_ORDEM2-LOW").caretPosition = 8
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/usr/tblZLOLMM025TC_TL100/btnT_TL100-D01[5,0]").setFocus()
            self.session.findById("wnd[0]/usr/tblZLOLMM025TC_TL100/btnT_TL100-D01[5,0]").press()
            self.session.findById("wnd[0]/usr/tblZLOLMM025TC_TL100/btnT_TL100-D01[5,0]").press()
            self.session.findById("wnd[0]/usr/tblZLOLMM025TC_TL100/btnT_TL100-D01[5,0]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[25]").press()
        except Exception:
            self._tratar_falha(data, index, data_126, Status.ORDEM_CONGELADA,
                               "Falha ao preencher ZLOLMM025 (Montador — etapa campos)")
            return

        # ── Passo 6: confirmar popups ─────────────────────────────────────
        if not self._confirmar_popups(data, index):
            print("Falha ao confirmar popup (Montador).")
            self.reset_transacao()
            return

        # ── Passo 7: STATUS de sucesso ────────────────────────────────────
        self.excel.atualizar_status(data, index, Status.SOLICITADA)
        print(f"STATUS atualizado: {Status.SOLICITADA}")
        self.reset_transacao()
        print("✓ Linha processada com sucesso (MONTADOR)")

    # =========================================================================
    # TRANSAÇÃO ZDPQPL126 — FABRICANTE
    # =========================================================================

    def processar_fabricante(self, row: pd.Series, index: int, data: pd.DataFrame) -> bool:
        """
        Processa uma linha de ordem do tipo FABRICANTE via ZDPQPL126 + ZLOLMM025.

        Análogo a processar_montador, com P_LINHA="F" e IDs específicos da
        aba Fabricante (txtP_TAKT4, ctxtS_PERID4-LOW, ctxtS_ORDEM4-LOW).

        Retorna:
            bool: True em sucesso, False em qualquer falha (sem relançar).
        """
        data_126: list[str] = []

        print(index, row["CARRO"], row["OP"], row["OPERAÇÃO"], row["STATUS"])
        op_val = str(row["OPERAÇÃO"]).strip().zfill(4)
        print(f"Operação formatada: '{op_val}'")

        # ── Passo 1: acessar ZDPQPL126 ────────────────────────────────────
        try:
            time.sleep(1)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NZDPQPL126"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/txtSP$00003-LOW").text = row["OP"]
            self.session.findById("wnd[0]/usr/ctxt%LAYOUT").text = "/FRBARRO"
            self.session.findById("wnd[0]/usr/ctxt%LAYOUT").setFocus()
            self.session.findById("wnd[0]/usr/ctxt%LAYOUT").caretPosition = 8
            self.session.findById("wnd[0]").sendVKey(0)
        except Exception:
            time.sleep(1)
            self._tratar_falha(data, index, data_126, Status.ORDEM_NAO_SINC,
                               "Falha ao acessar ZDPQPL126 (Fabricante)")
            return False

        # ── Passos 2 e 3: exportar TXT e extrair valor ──────────────────
        if not self._exportar_e_ler_txt(data, index, data_126, op_val):
            return False  # STATUS já atualizado

        # ── Passo 4: capturar valor da shell ─────────────────────────────
        valor_shell = self.session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell").text
        print(f"✓ Valor da shell SAP: {valor_shell}")
        data_126.append(valor_shell)

        # ── Passo 5: ZLOLMM025 com linha "F" ─────────────────────────────
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NZLOLMM025"
        self.session.findById("wnd[0]").sendVKey(0)

        try:
            self.session.findById("wnd[0]/usr/ctxtS_CENTRO-LOW").text = "BOT1"
            self.session.findById("wnd[0]/usr/ctxtP_AS").text = "AS-B28"
            self.session.findById("wnd[0]/usr/ctxtP_LINHA").text = "F"
            self.session.findById("wnd[0]/usr/ctxtP_LINHA").setFocus()
            self.session.findById("wnd[0]/usr/ctxtP_LINHA").caretPosition = 1
            self.session.findById("wnd[0]").sendVKey(0)
        except Exception:
            self._tratar_falha(data, index, data_126, Status.ORDEM_CONGELADA,
                               "Falha ao preencher ZLOLMM025 (Fabricante — etapa navegação)")
            return False

        try:
            time.sleep(1)
            self.session.findById("wnd[0]/usr/txtP_TAKT4").text = "1"
            self.session.findById("wnd[0]/usr/ctxtS_PERID4-LOW").text = data_126[0]
            self.session.findById("wnd[0]/usr/ctxtS_ORDEM4-LOW").text = row["OP"]
            self.session.findById("wnd[0]/usr/ctxtS_ORDEM4-LOW").setFocus()
            self.session.findById("wnd[0]/usr/ctxtS_ORDEM4-LOW").caretPosition = 8
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[0]/usr/tblZLOLMM025TC_TL100/btnT_TL100-D01[5,0]").setFocus()
            self.session.findById("wnd[0]/usr/tblZLOLMM025TC_TL100/btnT_TL100-D01[5,0]").press()
            self.session.findById("wnd[0]/usr/tblZLOLMM025TC_TL100/btnT_TL100-D01[5,0]").press()
            self.session.findById("wnd[0]/usr/tblZLOLMM025TC_TL100/btnT_TL100-D01[5,0]").press()
            self.session.findById("wnd[0]/tbar[1]/btn[25]").press()
        except Exception:
            self._tratar_falha(data, index, data_126, Status.ORDEM_CONGELADA,
                               "Falha ao preencher ZLOLMM025 (Fabricante — etapa campos)")
            return False

        # ── Passo 6: confirmar popups ─────────────────────────────────────
        if not self._confirmar_popups(data, index):
            print("Falha ao confirmar popup (Fabricante).")
            self.excel.atualizar_status(data, index, Status.FALHA_POPUP)
            self.reset_transacao()
            return False

        # ── Passo 7: STATUS de sucesso ────────────────────────────────────
        self.excel.atualizar_status(data, index, Status.SOLICITADA)
        print(f"STATUS atualizado: {Status.SOLICITADA}")
        self.reset_transacao()
        print("✓ Linha processada com sucesso (FABRICANTE)")
        return True

    # =========================================================================
    # TRANSAÇÃO ZLOLMM027 — MTS (fallback / roteamento direto)
    # =========================================================================

    def processar_zlolmm027(self, row: pd.Series, index: int, data: pd.DataFrame) -> None:
        """
        Processa uma linha via transação ZLOLMM027 (MTS).

        Campos: P_PICK (código do carro), P_AUFNR (OP), S_VORNR-LOW (operação).

        O STATUS de sucesso ("ORDEM SOLICITADA VIA ZLOLMM027") NÃO é gravado
        aqui — é responsabilidade do orquestrador gravá-lo após retorno bem-sucedido.

        Levanta:
            Exception: Qualquer falha é relançada para o orquestrador manter
                       o STATUS de falha sem sobrescrevê-lo.
        """
        data_126: list[str] = []

        print(index, row["CARRO"], row["OP"], row["OPERAÇÃO"], row["STATUS"])
        op_val = str(row["OPERAÇÃO"]).strip().zfill(4)
        print(f"Operação formatada: '{op_val}'")

        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NZLOLMM027"
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/usr/ctxtP_PICK").text = row["CARRO"]
            self.session.findById("wnd[0]/usr/ctxtP_AUFNR").text = row["OP"]
            self.session.findById("wnd[0]/usr/txtS_VORNR-LOW").text = row["OPERAÇÃO"]
            self.session.findById("wnd[0]/usr/txtS_VORNR-LOW").setFocus()
            self.session.findById("wnd[0]/usr/txtS_VORNR-LOW").caretPosition = 4
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
            self.session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            self._tratar_falha(data, index, data_126, Status.FALHA_MTS,
                               "Falha ao acessar ZLOLMM027")
            # Reset adicional para garantir estado limpo
            try:
                self.session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
                self.session.findById("wnd[0]").sendVKey(0)
            except Exception:
                pass
            raise  # Relança para o orquestrador

        print("✓ Linha processada com sucesso (ZLOLMM027 / MTS)")

    # =========================================================================
    # TRANSAÇÃO ZLOBMM001 — Fechamento final
    # =========================================================================

    def executar_zlobmm001(self) -> None:
        """Executa o fechamento final via transação ZLOBMM001."""
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/NZLOBMM001"
        self.session.findById("wnd[0]").sendVKey(0)
        self.session.findById("wnd[0]/tbar[1]/btn[17]").press()
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
        self.session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
        self.session.findById("wnd[0]/tbar[1]/btn[8]").press()
        self.session.findById("wnd[0]/tbar[0]/okcd").text = "/N"
        self.session.findById("wnd[0]").sendVKey(0)
