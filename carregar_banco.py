import os
import pandas as pd
from sqlalchemy import create_engine

# 1. Configurar a conexão com o MySQL
# Formato: mysql+pymysql://USUARIO:SENHA@HOST:PORTA/BANCO
engine = create_engine("mysql+pymysql://root:SUA_SENHA_AQUI@127.0.0.1:3306/olist_db")

# 2. Caminho onde estão os seus arquivos CSV baixados
caminho_pasta = r"C:\Projetos\archive"

# 3. Dicionário com os arquivos e os nomes das tabelas no MySQL
arquivos = {
    "pedidos": "olist_orders_dataset.csv",
    "itens_pedidos": "olist_order_items_dataset.csv",
    "clientes": "olist_customers_dataset.csv",
    "avaliacoes": "olist_order_reviews_dataset.csv",
}

# 4. Importar cada arquivo substituindo as tabelas incompletas
for tabela, arquivo_csv in arquivos.items():
    caminho_completo = os.path.join(caminho_pasta, arquivo_csv)
    print(f"Carregando {arquivo_csv} para a tabela '{tabela}'...")

    df = pd.read_csv(caminho_completo)
    df.to_sql(
        tabela, con=engine, if_exists="replace", index=False, chunksize=10000
    )
    print(f"Sucesso: {len(df):,} linhas inseridas na tabela '{tabela}'.\n")

print("Importação completa finalizada!")