import pandas as pd

# Caminho do arquivo de entrada e saída
entrada = r"C:\Users\arnaldo.silva\Desktop\Inclusão_TELEMEDICINA_VIDAS_ MEDICAR ATUAL 21-07-2025_copia.csv"
saida = r"C:\Users\arnaldo.silva\Desktop\Layout_vitace_220725_corrigido.csv"

# Carregar o arquivo CSV usando separador ;
df = pd.read_csv(entrada, dtype=str, sep=";")

# Limpar valores das colunas R e S (deixar apenas os cabeçalhos)
if len(df.columns) >= 19:
    df.iloc[:, 17] = ""
    df.iloc[:, 18] = ""

# Excluir colunas indesejadas, se existirem
colunas_excluir = ['Unnamed: 23', 'Unnamed: 24', 'Unnamed: 25', 'OBSERVACAO']
df = df.drop(columns=[col for col in colunas_excluir if col in df.columns])

# Excluir a coluna CODTIT se existir
if 'CODTIT' in df.columns:
    df = df.drop(columns=['CODTIT'])

# Preencher COD TIT para titulares e dependentes
if 'TIPO' in df.columns and 'CPF' in df.columns:
    if 'COD TIT' not in df.columns:
        df['COD TIT'] = ""
    # Titulares recebem seu próprio CPF
    df.loc[df['TIPO'] == '1', 'COD TIT'] = df.loc[df['TIPO'] == '1', 'CPF']
    # Dependentes recebem o CPF do titular mais próximo acima
    last_titular_cpf = None
    for idx, row in df.iterrows():
        if row['TIPO'] == '1':
            last_titular_cpf = row['CPF']
        df.at[idx, 'COD TIT'] = last_titular_cpf

# Preencher CODIGODEP para dependentes
if 'TIPO' in df.columns and 'CPF' in df.columns:
    if 'CODIGODEP' not in df.columns:
        df['CODIGODEP'] = ""
    df.loc[df['TIPO'] == '2', 'CODIGODEP'] = df.loc[df['TIPO'] == '2', 'CPF']

# Exportar para CSV
df.to_csv(saida, index=False, sep=";")