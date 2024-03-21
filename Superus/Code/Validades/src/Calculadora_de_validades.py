import os
import re
import sys
import logging
import pyperclip
import pandas as pd
from time import sleep
import pyautogui as pg
import pygetwindow as gw
from datetime import datetime
from Validades.src.querry.precos_4_lojas import verificar_e_atualizar_csv_precos
from Validades.src.querry.relatorio_tresmann import verificar_e_atualizar_csv_tresmann

def abrir_janela(nome_da_janela):
    """ Seleciona a janela especificada e a ativa. """
    try:
        janelas = gw.getWindowsWithTitle(nome_da_janela)
        if janelas:
            janela = janelas[0]
            if janela.isMinimized:
                janela.restore()
            janela.activate()
            sleep(1)
        else:
            print(f"Janela {nome_da_janela} não encontrada.")
            sys.exit()
    except Exception as e:
        print(f"Erro ao ativar a janela: {e}")
        abrir_janela(nome_da_janela)

def procurar_botao(imagem, clicar):
    """ Localiza um botão na tela e clica nele se especificado. """
    while True:
        try:
            localizacao = pg.locateOnScreen(imagem, confidence=0.9)
            if localizacao:
                if clicar:
                    x, y, width, height = localizacao
                    pg.click(x + width // 2, y + height // 2, duration=0.1)
                break
            sleep(1)
        except Exception:
            continue

def copiar_para_area_de_transferencia(codigo, nome_produto, dias_a_vencer, vcto_medio, estoque_atual, venda_media_30_dias, projecao_de_venda, perda_financeira_projetada, resultado):
    # Converter todas as variáveis para strings
    codigo = str(codigo)
    dias_a_vencer = str(dias_a_vencer)
    vcto_medio = str(vcto_medio)
    estoque_atual = str(estoque_atual)
    venda_media_30_dias = str(venda_media_30_dias)
    projecao_de_venda = str(projecao_de_venda)

    # Use condicionais para definir a situação
    if resultado == "RECUSAR":
        situacao = f"*SITUAÇÃO: {resultado}* @Werllen Agoraa 2"
    elif resultado == "RECEBER":
        situacao = f"*SITUAÇÃO: {resultado}*"

    # Agora crie a lista de dados formatados
    dados_formatados = [
        f"*{nome_produto}*",
        f"CÓDIGO: {codigo}",
        f"DIAS A VENCER: {dias_a_vencer}",
        f"VCTO MÉDIO: {vcto_medio}",
        f"ESTOQUE ATUAL: {estoque_atual}",
        f"VENDA MEDIA(30D): {venda_media_30_dias}",
        f"PROJEÇÃO DE VENDA: {projecao_de_venda}",
        f"ESTIMATIVA DE PERDA: {perda_financeira_projetada}",
        situacao
    ]

    # Juntar os dados em uma única string com quebras de linha
    dados_formatados_str = "\n".join(dados_formatados)

    # Copiar para a área de transferência
    pyperclip.copy(dados_formatados_str)

    pg.hotkey("Ctrl", "v"); sleep(1)
    pg.press("enter", presses=2, interval=1)
    sleep(2)

def buscar_info_produto(df, codigo, loja, coluna):
    result = df.loc[(df['CODIGO'] == codigo) & (df['LOJA'] == loja), coluna]

    return result.iloc[0] if not result.empty else 0


# Funções específicas baseadas na função genérica acima
def buscar_nome_produto(codigo, loja, df):
    return buscar_info_produto(df, codigo, loja, 'NOME')

def buscar_fornecedor_produto(codigo, loja, df):
    return buscar_info_produto(df, codigo, loja, 'NOME_FANTASIA')

def buscar_vcto_produto(codigo, loja, df):
    return buscar_info_produto(df, codigo, loja, 'VCTO_MEDIO')

def buscar_estoque_produto(codigo, loja, df):
    return buscar_info_produto(df, codigo, loja, 'ESTOQUE ATUAL')

def buscar_venda_media_produto(codigo, loja, df):
    return buscar_info_produto(df, codigo, loja, 'VENDA MEDIA')

# Função ajustada para usar o dicionário
def processar_arquivo(caminho_arquivo):
    df_final_tresmann = verificar_e_atualizar_csv_tresmann()
    df_final_precos = verificar_e_atualizar_csv_precos()

    lojas_dict = {
    "smj": "TRESMANN - SMJ",
    "stt": "TRESMANN - STT",
    "vix": "TRESMANN - VIX"
    }

    nome_arquivo, _ = os.path.splitext(os.path.basename(caminho_arquivo))
    nome_arquivo_limpo = re.sub(r'[^a-zA-Z0-9 ]', '', nome_arquivo)
    nome_arquivo_limpo = nome_arquivo_limpo.lower()
    palavras_nome_arquivo = nome_arquivo_limpo.split()

    loja = "Nome da Loja Não Encontrado"

    for palavra in palavras_nome_arquivo:
        if palavra in lojas_dict:
            loja = lojas_dict[palavra]
            loja_abreviada = palavra.upper()
            break

    if loja == "Nome da Loja Não Encontrado":
        return

    abrir_janela("WhatsApp")
    if loja == "TRESMANN - VIX":
        procurar_botao(r"Imgs\recebimentovix.png", clicar=True)
    elif loja == "TRESMANN - SMJ":
        procurar_botao(r"Imgs\recebimentosmj.png", clicar=True)
    elif loja == "TRESMANN - STT":
        procurar_botao(r"Imgs\recebimentostt.png", clicar=True)

    procurar_botao(r"Imgs\digite.png", clicar=True)

    primeira_iteracao = True

    # Lendo o arquivo .xlsx com pandas
    df = pd.read_excel(caminho_arquivo)

    # Iterando sobre as linhas do DataFrame
    for index, linha in df.iterrows():
        # Supondo que a primeira coluna seja o código, a sexta a quantidade e a sétima a data de validade
        codigo = int(linha['Código'])
        quantidade = float(str(linha['Quantidade']).replace(',', '.'))
        data_validade = linha['Validade']

        df_final_tresmann = verificar_e_atualizar_csv_tresmann()
        df_final_precos = verificar_e_atualizar_csv_precos()

        # Trabalhando com datas
        data_validade = pd.to_datetime(data_validade, format='%d/%m/%Y', errors='coerce')
        data_hoje = pd.to_datetime(datetime.now().strftime('%d/%m/%Y'), format='%d/%m/%Y')
        dias_a_vencer = (data_validade - data_hoje).days

        # Cálculos de projeção e estoque
        venda_media_30_dias = (buscar_venda_media_produto(codigo, loja, df_final_tresmann))
        projecao_de_venda = ((venda_media_30_dias / 30) * dias_a_vencer) - 2
        estoque_atual = buscar_estoque_produto(codigo, loja, df_final_tresmann)  
        primeira_condicao = (estoque_atual + quantidade) if projecao_de_venda < 0 else (estoque_atual + quantidade - projecao_de_venda)
        resultado = 0 if primeira_condicao < 0 else primeira_condicao

        # Cálculo do custo líquido
        def buscar_custoliquido_produto(codigo, loja, resultado):
            result = df_final_precos.loc[(df_final_precos['CODIGO'] == codigo) & (df_final_precos['NOME_LOJA'] == loja), 'CUSTOLIQUIDO']
            if result.empty:
                return None
            custo_liquido = result.iloc[0] * resultado
            return f'R${custo_liquido:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

        perda_financeira_projetada = buscar_custoliquido_produto(codigo, loja, resultado)

        # Lógica condicional final
        vcto_medio = buscar_vcto_produto(codigo, loja, df_final_tresmann)
        if projecao_de_venda > (quantidade + estoque_atual):
            resultado = "RECEBER"
        elif dias_a_vencer > (vcto_medio - (vcto_medio * 0.10)):
            resultado = "RECEBER"
        else:
            resultado = "RECUSAR"

        nome_produto = buscar_nome_produto(codigo, loja, df_final_tresmann)

        if primeira_iteracao:
            fornecedor = buscar_fornecedor_produto(codigo, loja, df_final_tresmann)
            logging.info("Enviando Validade de %s - %s", fornecedor, loja_abreviada)
            pg.write(f"*{fornecedor} - {loja_abreviada}*"); sleep(1)
            pg.press("enter", presses=2, interval=1)
            primeira_iteracao = False

        logging.info("Produto %s, Situação %s", nome_produto, resultado)
        copiar_para_area_de_transferencia(codigo, nome_produto, dias_a_vencer, int(vcto_medio), int(estoque_atual), int(venda_media_30_dias), int(projecao_de_venda), perda_financeira_projetada, resultado)

    abrir_janela("Visual Studio Code")