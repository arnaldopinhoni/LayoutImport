import pandas as pd
import os

# Caminho completo para o arquivo Excel que contém as abas
caminho_do_arquivo_origem = r'X:\CONTRATOS PRIVADOS\IMPORTAÇÃO DE CONTRATOS\Processamento de Arquivos.xlsx'

# --- Variáveis para a exportação ---
nome_arquivo_csv = 'controle_clientes_com_datas_e_max_colE.csv'
caminho_salvar_csv = "" # Deixe vazio "" para salvar na mesma pasta do script

# Constrói o caminho completo para o arquivo CSV
if caminho_salvar_csv:
    caminho_completo_csv = os.path.join(caminho_salvar_csv, nome_arquivo_csv)
else:
    caminho_completo_csv = nome_arquivo_csv

try:
    # 1. Obter os nomes das abas do arquivo de origem
    xls = pd.ExcelFile(caminho_do_arquivo_origem)
    nomes_das_abas = xls.sheet_names

    # Dicionário para armazenar as datas mais recentes por cliente/aba
    datas_por_cliente = {}
    # Dicionário para armazenar o maior valor da 5ª coluna (coluna E) por cliente/aba
    max_valor_quinta_coluna_por_cliente = {} # Renomeado para maior clareza

    print("Processando abas para extrair a data mais recente e o maior valor da 5ª coluna (E)...")

    # 2. Iterar sobre cada aba para encontrar as informações necessárias
    for aba_nome in nomes_das_abas:
        try:
            # Carrega a aba atual para um DataFrame temporário
            temp_df = pd.read_excel(caminho_do_arquivo_origem, sheet_name=aba_nome)

            # --- Lógica para Extração da Data Mais Recente (mantida) ---
            max_date_in_aba = pd.NaT
            converted_df_dates = temp_df.apply(lambda x: pd.to_datetime(x, errors='coerce'))
            stacked_dates = converted_df_dates.stack()
            if not stacked_dates.empty:
                max_date_in_aba = stacked_dates.max()

            if pd.isna(max_date_in_aba):
                datas_por_cliente[aba_nome] = None
                print(f"  - Aviso: Nenhuma data válida encontrada na aba '{aba_nome}'. Data será vazia.")
            else:
                datas_por_cliente[aba_nome] = max_date_in_aba.strftime('%Y-%m-%d')
                print(f"  - Data mais recente para '{aba_nome}': {datas_por_cliente[aba_nome]}")

            # --- NOVO: Lógica para Extração do Maior Valor da 5ª Coluna (Coluna E) ---
            # Verifica se a aba tem pelo menos 5 colunas (índices 0 a 4)
            if temp_df.shape[1] >= 5: # shape[1] é o número de colunas
                # Acessa a 5ª coluna (índice 4) usando iloc
                col_e_data = temp_df.iloc[:, 4] 
                
                # Tenta converter a coluna para numérica. Valores não numéricos viram NaN.
                col_e_numerica = pd.to_numeric(col_e_data, errors='coerce')
                # Encontra o valor máximo na coluna numérica (ignorando NaNs)
                max_value_e = col_e_numerica.max()

                # Verifica se o valor máximo encontrado é NaN (coluna vazia ou sem números válidos)
                if pd.isna(max_value_e):
                    max_valor_quinta_coluna_por_cliente[aba_nome] = None
                    print(f"  - Aviso: 5ª coluna (E) na aba '{aba_nome}' não contém valores numéricos válidos. Max Valor Coluna E será vazio.")
                else:
                    max_valor_quinta_coluna_por_cliente[aba_nome] = max_value_e
                    print(f"  - Maior valor da 5ª coluna (E) para '{aba_nome}': {max_value_e}")
            else:
                max_valor_quinta_coluna_por_cliente[aba_nome] = None
                print(f"  - Aviso: A aba '{aba_nome}' não possui uma 5ª coluna (E). Max Valor Coluna E será vazio.")

        except Exception as e_aba:
            print(f"  - Erro geral ao processar a aba '{aba_nome}': {e_aba}. Datas e Max Valor Coluna E serão vazios.")
            datas_por_cliente[aba_nome] = None
            max_valor_quinta_coluna_por_cliente[aba_nome] = None

    # 3. Criar o DataFrame de controle com a coluna 'clientes'
    df_controle = pd.DataFrame(nomes_das_abas, columns=['clientes'])

    # 4. Mapear as datas encontradas para a coluna 'DATAS'
    df_controle['DATAS'] = df_controle['clientes'].map(datas_por_cliente)

    # 5. Mapear os maiores valores da 5ª coluna para a nova coluna
    df_controle['Max Valor Coluna E'] = df_controle['clientes'].map(max_valor_quinta_coluna_por_cliente)

    print("\nDataFrame 'df_controle' com datas e maiores valores da Coluna 'E' preenchidos:")
    print(df_controle.head())

    # 6. Exportar o DataFrame para CSV
    df_controle.to_csv(caminho_completo_csv, index=False, encoding='utf-8')

    print(f"\nDataFrame exportado com sucesso para: {caminho_completo_csv}")
    print("O arquivo CSV agora inclui a data mais recente de cada aba e o maior valor da 5ª coluna (E).")

except FileNotFoundError:
    print(f"Erro: O arquivo '{caminho_do_arquivo_origem}' não foi encontrado.")
    print("Por favor, verifique se o caminho está correto e se você tem permissão de acesso à pasta 'C:\Users\arnaldo.silva\Desktop\'.")
except Exception as e:
    print(f"Ocorreu um erro geral ao processar o arquivo: {e}")
    print("Certifique-se de que a biblioteca 'openpyxl' está instalada ('pip install openpyxl').")