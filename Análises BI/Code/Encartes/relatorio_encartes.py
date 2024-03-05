import os
import pandas as pd
from datetime import datetime, timedelta
from extrair_encartes import ExtratorDeEncartes
from config import CAMINHO_ENCARTE_SMJ, CAMINHO_ENCARTE_STT, ARQUIVO_EXCEL


class ManipuladorDeDatas:
    @staticmethod
    def calcular_datas_extremas(considerar_penultima=True):
        hoje = datetime.now()
        dia_semana = hoje.weekday()
        if considerar_penultima and dia_semana == 4:
            sexta = hoje - timedelta(7)
            quinta = hoje - timedelta(1)
        else:
            sexta = hoje - timedelta(days=(dia_semana + 3) % 7)
            quinta = hoje + timedelta(days=(3 - dia_semana) % 7)
        return sexta.strftime('%d/%m/%Y'), quinta.strftime('%d/%m/%Y')

class ManipuladorExcel:
    @staticmethod
    def ler_penultima_linha(arquivo):
        df = pd.read_excel(arquivo, engine='xlrd')
        if len(df) > 1:
            penultima_linha = df.iloc[[-2]]
            colunas_nao_nulas = penultima_linha.dropna(axis=1)
            return colunas_nao_nulas
        else:
            return pd.DataFrame()

    @staticmethod
    def renomear_colunas(df):
        novos_nomes_colunas = [
            'Quantidade', 'Custo', 'Venda', 'Lucro Bruto',
            'Lucro Líquido', 'Margem Bruta', 'Margem Líquida',
            'Participação', 'Contribuição'
        ]
        df.columns = novos_nomes_colunas[:len(df.columns)]
        return df

    @staticmethod
    def atualizar_ou_adicionar_linha_excel(arquivo_excel, df_nova):
        try:
            tabela_existente = pd.read_excel(arquivo_excel, engine='openpyxl')
        except FileNotFoundError:
            tabela_existente = pd.DataFrame()

        colunas_necessarias = ['LOJA', 'DATA INICIAL', 'DATA FINAL']
        if all(coluna in tabela_existente.columns for coluna in colunas_necessarias):
            for _, nova_linha in df_nova.iterrows():
                indices = tabela_existente[
                    (tabela_existente['LOJA'] == nova_linha['LOJA']) &
                    (tabela_existente['DATA INICIAL'] == nova_linha['DATA INICIAL']) & 
                    (tabela_existente['DATA FINAL'] == nova_linha['DATA FINAL'])
                ].index
                
                if len(indices) > 0:
                    for col in df_nova.columns:
                        tabela_existente.at[indices[0], col] = nova_linha[col]
                else:
                    tabela_existente = pd.concat([tabela_existente, pd.DataFrame([nova_linha])], ignore_index=True)
        else:
            tabela_existente = pd.concat([tabela_existente, df_nova], ignore_index=True)
        
        with pd.ExcelWriter(arquivo_excel) as writer:
            tabela_existente.to_excel(writer, index=False, sheet_name="ENCARTES")
            ManipuladorExcel.ajustar_largura_colunas(writer, tabela_existente)

    @staticmethod
    def ajustar_largura_colunas(writer, df):
        worksheet = writer.sheets["ENCARTES"]
        for i, col in enumerate(df.columns):
            col_width = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.set_column(i, i, col_width)

class RelatorioEncartes:
    def gerar_relatorio(self):
        dfs = []
        sexta, quinta = ManipuladorDeDatas.calcular_datas_extremas()

        for caminho in [CAMINHO_ENCARTE_SMJ, CAMINHO_ENCARTE_STT]:
            loja = "Desconhecido"
            if "SMJ" in caminho.upper():
                loja = "SMJ"
            elif "STT" in caminho.upper():
                loja = "STT"
            
            df = ManipuladorExcel.ler_penultima_linha(caminho)
            df_renomeado = ManipuladorExcel.renomear_colunas(df)
            df_final = pd.DataFrame({
                "LOJA": [loja],
                "DATA INICIAL": [sexta],
                "DATA FINAL": [quinta]
                
            })
            for coluna in df_renomeado.columns:
                df_final[coluna] = df_renomeado.iloc[0][coluna]
            
            dfs.append(df_final)
            os.remove(caminho)

        relatorio_final = pd.concat(dfs, ignore_index=True)
        ManipuladorExcel.atualizar_ou_adicionar_linha_excel(ARQUIVO_EXCEL, relatorio_final)

if __name__ == "__main__":
    extrair = ExtratorDeEncartes()
    relatorio = RelatorioEncartes()

    extrair.extrair_encartes()
    relatorio.gerar_relatorio()