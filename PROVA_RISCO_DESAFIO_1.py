#DESAFIO: QUESTÃO 1
"""
Crie um programa em Python que utilize o modelo de precificação de opções BlackScholes para avaliar o preço de uma opção de compra ou venda. O programa deve ser capaz de receber informações sobre o preço atual do ativo subjacente, a taxa livre de risco, a volatilidade do ativo e o tempo restante até o vencimento da opção. 
"""

import pandas as pd
import numpy as np
import datetime as dt
from xbbg import blp

#FUNÇÕES
def cotacoes_BBG(lista_ativos, parametro, dt_di, dt_df):
    
    df = blp.bdh(lista_ativos, parametro, dt_di, dt_df, currency="BRL") # Requisição na BBG
    print("Conexão OK")
    df.columns = df.columns.droplevel(1)
    
    return df

def beta(df):  
    
    result = {}
    for i in lista_ativos[:-1]:
        df2 = df.pct_change()
        df2 = df2[[i, "IBOV INDEX"]].dropna()

        cov_ativo_bench = np.cov(df2[i], df2["IBOV INDEX"])[1][0]
        variancia_bench = np.var(df2["IBOV INDEX"])
        beta = cov_ativo_bench/variancia_bench
        result[i] = beta

    result = pd.DataFrame([result]).T
    return result

#PARÂMETROS QUE PODEM SER PASSADOS POR UM ARQUIVO DE CONFIGURAÇÃO
dt_di = dt.datetime(2015,1,1)
dt_df = dt.datetime(2023,3,16)
parametro = "PX_LAST"
lista_ativos = ["BRPR3 BZ EQUITY","ALPA4 BZ EQUITY", "LREN3 BZ EQUITY" ,"EQTL3 BZ EQUITY", "ELET3 BZ EQUITY", "NTCO3 BZ EQUITY",
                "RAIL3 BZ EQUITY", "HAPV3 BZ EQUITY", "IGTI11 BZ EQUITY", "BRML3 BZ EQUITY", "GGPS3 BZ EQUITY", "VBBR3 BZ EQUITY",
                "RECV3 BZ EQUITY", "IBOV INDEX"]


#CHAMA AS FUNÇÕES COM OS PARÂMETROS ESTABELECIDOS
df_ativo = cotacoes_BBG(lista_ativos, parametro, dt_di, dt_df)
beta(df_ativo)