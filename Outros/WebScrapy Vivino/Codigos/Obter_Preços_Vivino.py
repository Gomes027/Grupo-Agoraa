import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def obter_preco_e_informacoes_produto(numero_produto, url, xpath, codigo_produto, nome_produto):
    driver.get(url)

    try:
        elemento_preco = driver.find_element('xpath', xpath)
        preco_com_simbolo = elemento_preco.text
        palavras = preco_com_simbolo.split()

        if palavras[-1].startswith("R$"):
            preco = palavras[-1]
        else:
            preco = " "

    except:
        preco = " "

    print("URL:", url)
    print(f"Produto {numero_produto}:", nome_produto)
    print("Código:", codigo_produto)
    print("Preço:", preco)
    print("="*80)

    preco = preco.replace("R$", "").strip()
    indice_produto = nomes_produtos.index(nome_produto)
    planilha.at[indice_produto, 'Preços Vivino'] = preco
        
planilha_1 = r"F:\COMPRAS\Automações.Compras\Automações\Superus\Vivino x Tresmann\Planilhas\Base de dados_CSV.xlsx"
planilha_2 = r"F:\COMPRAS\Automações.Compras\Automações\Superus\Vivino x Tresmann\Planilhas\Tresmann - Sem Preço.xlsx"
planilha_3 = r"F:\COMPRAS\Automações.Compras\Automações\Superus\Vivino x Tresmann\Planilhas\Vivino - Sem Preço.xlsx"

try:
    os.remove(planilha_1)
    os.remove(planilha_2)
    os.remove(planilha_3)
except:
    pass

planilha = pd.read_excel(r"F:\COMPRAS\Automações.Compras\Automações\Superus\Vivino x Tresmann\Outros\Dados\Base de dados.xlsx", sheet_name="BASE", engine='openpyxl')

nomes_produtos = planilha.iloc[:, 0].tolist()
codigos = planilha.iloc[:, 1].tolist()
urls = planilha.iloc[:, 2].tolist()
xpaths = planilha.iloc[:, 3].tolist()

chrome_options = Options()
chrome_options.add_argument("--incognito")
driver = webdriver.Chrome(options=chrome_options)

os.system("cls")

for numero, (url, xpath, codigo, nome) in enumerate(zip(urls, xpaths, codigos, nomes_produtos), start=1):
    obter_preco_e_informacoes_produto(numero, url, xpath, codigo, nome)

planilha.to_excel(r"F:\COMPRAS\Automações.Compras\Automações\Superus\Vivino x Tresmann\Planilhas\Base de dados_CSV.xlsx", sheet_name="CSV", index=False)

original_file_path = r'F:\COMPRAS\Automações.Compras\Automações\Superus\Vivino x Tresmann\Planilhas\Base de dados_CSV.xlsx'

df_original = pd.read_excel(original_file_path)
df_space_in_fifth_col = df_original[df_original.iloc[:, 4] == " "]
output_file_path = r'F:\COMPRAS\Automações.Compras\Automações\Superus\Vivino x Tresmann\Planilhas\Tresmann - Sem Preço.xlsx'
df_space_in_fifth_col.to_excel(output_file_path, index=False)
df_original = df_original[df_original.iloc[:, 4] != " "]
df_original.to_excel(original_file_path, index=False)
print("Linhas com espaço na quinta coluna foram movidas para o arquivo 'Tresmann - Sem Preço.xlsx'.")


df_original = pd.read_excel(original_file_path)
df_empty_sixth_col = df_original[df_original.iloc[:, 5].isna()]
output_file_path = r'F:\COMPRAS\Automações.Compras\Automações\Superus\Vivino x Tresmann\Planilhas\Vivino - Sem Preço.xlsx'
df_empty_sixth_col.to_excel(output_file_path, index=False)
df_original = df_original.dropna(subset=[df_original.columns[5]])
df_original.to_excel(original_file_path, index=False)
print("Linhas vazias na sexta coluna foram movidas para o arquivo 'Vivino - Sem Preço.xlsx'.")

print("="*80)

driver.quit()
