# =============================================================================
# orchestrator.py — Orquestra o processamento de todas as ordens de produção
# =============================================================================

import os
import sys
import time
from typing import Any

import pandas as pd
from tkinter import messagebox

from config import EXCEL_PATH, TIPO_CARRO, CARRO_ZLOLMM027, Status
from excel_manager import ExcelManager
from transactions import TransactionHandler


class Orchestrator:
    """
    Coordena o fluxo completo de processamento das ordens.

    Responsabilidades:
        - Ler a planilha de controle
        - Rotear cada linha para a transação correta (Montador / ZLOLMM027 / Fabricante)
        - Garantir que o STATUS seja gravado em todos os caminhos de erro
        - Executar o fechamento final (ZLOBMM001) e retornar o resumo

    Parâmetros:
        session : Objeto GuiSession ativo do SAP GUI Scripting.
    """

    def __init__(self, session: Any) -> None:
        self.session = session
        self.excel   = ExcelManager(EXCEL_PATH)
        self.tx      = TransactionHandler(session, self.excel)

    # =========================================================================
    # HELPER DE ROTEAMENTO PARA ZLOLMM027
    # =========================================================================

    def _rotear_para_zlolmm027(
        self,
        row: pd.Series,
        index: int,
        data: pd.DataFrame,
    ) -> None:
        """
        Aciona ZLOLMM027 e atualiza STATUS em caso de sucesso.

        O STATUS de erro já é definido internamente por processar_zlolmm027
        se a transação falhar; neste método tratamos apenas o caminho feliz.
        """
        try:
            self.tx.processar_zlolmm027(row, index, data)
            self.excel.atualizar_status(data, index, Status.SOLICITADA_MTS)
            print(f"✓ ZLOLMM027_MTS concluído | STATUS: {Status.SOLICITADA_MTS}")
        except Exception as e:
            print(f"❌ Erro em ZLOLMM027_MTS: {e}")
            # STATUS já definido internamente por processar_zlolmm027

    # =========================================================================
    # FLUXO DE LOGIN
    # =========================================================================

    def fazer_login(self) -> tuple[int, int, int]:
        """
        Conclui o fluxo de login interativo do SAP e inicia o processamento.

        Etapas:
            1. Aguarda a tela de login carregar
            2. Confirma os popups iniciais de autenticação
            3. Delega o processamento das ordens a processar_ordens()

        Retorna:
            tuple(int, int, int): (total, sucesso, erros).
        """
        try:
            time.sleep(0.5)
            self.session.findById("wnd[0]").sendVKey(0)
            self.session.findById("wnd[1]/usr/btnENTER").press()
            self.session.findById("wnd[1]/usr/btnSEL_BUTTON").press()
        except Exception:
            print(f"Erro no login: {sys.exc_info()[0]}")

        time.sleep(1)
        return self.processar_ordens()

    # =========================================================================
    # ORQUESTRADOR PRINCIPAL
    # =========================================================================

    def processar_ordens(self) -> tuple[int, int, int]:
        """
        Lê a planilha e processa cada linha, roteando para a transação correta.

        Fluxo de roteamento:
            ┌─ CARRO em TIPO_CARRO?      → processar_montador
            ├─ CARRO em CARRO_ZLOLMM027? → processar_zlolmm027  (roteamento direto)
            └─ Nenhum dos anteriores     → processar_fabricante

        Retorna:
            tuple(int, int, int): (total, sucesso, erros) contados pela
                                  releitura do Excel ao final.
        """
        if not os.path.exists(EXCEL_PATH):
            messagebox.showerror("Erro", f"Arquivo não encontrado: {EXCEL_PATH}")
            return 0, 0, 0

        try:
            data = self.excel.ler_planilha()

            print("\n" + "=" * 80)
            print("INICIANDO PROCESSAMENTO DE ORDENS")
            print("=" * 80)

            for index, row in data.iterrows():
                try:
                    carro = row["CARRO"].strip()
                    print(f"\n[Linha {index + 1}] Processando CARRO: {carro}")

                    # ----------------------------------------------------------
                    if carro in TIPO_CARRO:
                        print(f"✓ '{carro}' → MONTADOR")
                        try:
                            self.tx.processar_montador(row, index, data)
                            print("✓ ZDPQPL126_MONTADOR concluído")
                        except Exception as e:
                            print(f"❌ Erro em ZDPQPL126_MONTADOR: {e}")

                    elif carro in CARRO_ZLOLMM027:
                        print(f"✓ '{carro}' → ZLOLMM027_MTS (roteamento direto)")
                        self._rotear_para_zlolmm027(row, index, data)

                    else:
                        print(f"✗ '{carro}' → FABRICANTE")
                        try:
                            sucesso = self.tx.processar_fabricante(row, index, data)
                            if sucesso:
                                print("✓ ZDPQPL126_FABRICANTE concluído")
                            else:
                                print("❌ ZDPQPL126_FABRICANTE finalizado com falha")
                        except Exception as e_fab:
                            print(f"❌ Exceção não mapeada no FABRICANTE: {e_fab}")
                            self.excel.atualizar_status(data, index, Status.ORDEM_NAO_SINC)
                            self.tx.reset_transacao()
                            continue
                    # ----------------------------------------------------------

                except KeyError as e:
                    print(f"❌ Coluna não encontrada no DataFrame: {e}")
                    continue
                except Exception as e:
                    print(f"❌ Erro inesperado na linha {index + 1}: {e}")
                    self.excel.atualizar_status(data, index, Status.ORDEM_NAO_SINC)
                    self.tx.reset_transacao()
                    continue

            print("\n" + "=" * 80)
            print("PROCESSAMENTO DE ORDENS FINALIZADO")
            print("=" * 80)

            # -----------------------------------------------------------------
            # Fechamento final
            # -----------------------------------------------------------------
            try:
                print("\nExecutando ZLOBMM001...")
                self.tx.executar_zlobmm001()
                print("✓ ZLOBMM001 executada com sucesso")
            except Exception as e:
                print(f"❌ Falha ao executar ZLOBMM001: {e}")

            # Flush de STATUS que falharam durante o processamento
            self.excel.persistir_pendentes(data)

            # -----------------------------------------------------------------
            # Contagem final — relê o Excel já persistido
            # "SOLICITADA" cobre "ORDEM SOLICITADA!" e "ORDEM SOLICITADA VIA ZLOLMM027"
            # -----------------------------------------------------------------
            data_final = pd.read_excel(EXCEL_PATH, sheet_name="DataBase").astype(str)
            total   = len(data_final)
            sucesso = int(data_final["STATUS"].str.contains("SOLICITADA", na=False).sum())
            erros   = total - sucesso

            print(f"\n  Total processadas : {total}")
            print(f"  Com sucesso       : {sucesso}")
            print(f"  Com erro          : {erros}")

            return total, sucesso, erros

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar planilha: {e}")
            print(f"❌ Erro geral: {e}")
            return 0, 0, 0
