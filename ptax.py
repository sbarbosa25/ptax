import pandas as pd
import os
from PyQt5.QtWidgets import QApplication, QFileDialog
import requests
import json

# Passo 1: Realizar a consulta à API do Banco Central do Brasil
# Aqui utilizo a URL fornecida no código M para buscar os dados JSON.
url = "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoMoedaPeriodo(moeda=@moeda,dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)?@moeda='USD'&@dataInicial='01-10-2024'&@dataFinalCotacao='12-29-2025'&$top=1000&$format=json&$select=cotacaoVenda,dataHoraCotacao,tipoBoletim"
response = requests.get(url)
data = response.json()

# Passo 2: Transformar os dados JSON em um DataFrame
# Extraio a lista de valores do JSON e converto para um DataFrame.
values = data["value"]
df = pd.DataFrame(values)

# Passo 3: Ajustar as colunas
# Aqui extraio os primeiros 19 caracteres da coluna "dataHoraCotacao" e ajusto os tipos de dados.
df["dataHoraCotacao"] = pd.to_datetime(df["dataHoraCotacao"].str[:19])
df["cotacaoVenda"] = pd.to_numeric(df["cotacaoVenda"], errors="coerce")
df["Moeda"] = "USD"  # Adiciono uma coluna fixa indicando a moeda.

# Passo 4: Exibir a janela para selecionar onde salvar o arquivo
# Uso o QFileDialog do PyQt5 para permitir ao usuário escolher o local onde salvar o arquivo.
app = QApplication([])
options = QFileDialog.Options()
options |= QFileDialog.DontUseNativeDialog
file_path, _ = QFileDialog.getSaveFileName(
    None, 
    "Salvar Arquivo", 
    "", 
    "Arquivos Excel (*.xlsx);;Todos os Arquivos (*)", 
    options=options
)

# Passo 5: Salvar o arquivo
# Caso o usuário selecione um local, salvo o DataFrame como um arquivo Excel.
if file_path:
    df.to_excel(file_path, index=False)
    print(f"Arquivo salvo com sucesso em: {file_path}")
else:
    print("Nenhum local selecionado. Operação cancelada.")

# Passo 6: Finalizar a aplicação Qt
# Fecho o aplicativo PyQt5.
app.exit()
