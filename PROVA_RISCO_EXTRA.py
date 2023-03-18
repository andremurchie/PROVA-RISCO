# %%
import numpy as np
from scipy.stats import norm
import json
import datetime as dt
import pandas as pd
from openpyxl import load_workbook, Workbook
import sys
import datetime
from IPython.display import display

N = norm.cdf
Nm = norm.pdf

def BS(spot, strike, time, juros, volatilite, flag):
    if flag == "CALL":
        flag = 1
    else:
        flag = -1
    
    time = time/252

    vp_k = strike/(1+juros)**time
    d1 = np.log(spot/vp_k) / (volatilite * np.sqrt(time)) + (volatilite * np.sqrt(time))/2
    d2 = np.log(spot/vp_k) / (volatilite * np.sqrt(time)) - (volatilite * np.sqrt(time))/2
    premio = flag * spot * N(flag * d1) - flag * vp_k * N(flag * d2) 
    
    return premio

def DELTA(spot, strike, time, juros, volatilite, flag):
    if flag == "CALL":
        flag = 1
    else:
        flag = -1

    time = time/252
    vp_k = strike/(1+juros)**time
    d = np.log(spot/strike) / (volatilite * np.sqrt(time)) + (volatilite * np.sqrt(time)/2)

    if flag == 1:
        delta = N(d) * np.exp(-juros * time)
    else:
        delta = N(d - 1) * np.exp(-juros * time)
    
    return delta

def GAMMA(spot, strike, time, juros, volatilite, flag):
    time = time/252
    vp_k = strike / (1+juros)**time
    d = np.log(spot/vp_k) / (volatilite * np.sqrt(time)) + (volatilite * np.sqrt(time)/2)
    gamma = Nm(d) / spot / volatilite / np.sqrt(time)
    
    return gamma

def VEGA(spot, strike, time, juros, volatilite, flag):
    time = time/252
    vp_k = strike/(1+juros)**time
    d = np.log(spot/vp_k) / (volatilite * np.sqrt(time)) + (volatilite * np.sqrt(time)/2)
    vega = N(d) * spot * np.sqrt(time)
    return vega
    
#################################################################################################

def op():
    print("Buscando JSON de opções")
    config = open("PATH JSON")
    config = json.load(config)

    choque_volatilidade = float(config["CHOQUE_VOLATILIDADE"].replace(",","."))
    choque_spot = float(config["CHOQUE_SPOT"].replace(",", "."))
    nome_estrututa = (config["NOME_ESTRUTURA"])
    nome_arquivo = (config["NOME_ARQUIVO"])
    ativo_objeto = (config["ATIVO_OBJETO"])
    pl = float(config["PL"].replace(",", "."))
    data_hoje = datetime.date.today()

    tabela_titulo =  pd.DataFrame(index= ["Ativo-Objeto","Patrimônio Líquido","Data", "Estrutura"])
    tabela_titulo["A"] = [ativo_objeto, pl, data_hoje, nome_estrututa] 
    


    lista_tipo = []
    lista_lado = []
    lista_strike = []
    lista_volatilidade = []
    lista_juros = []
    lista_spot = []
    lista_vencimento = []
    lista_quantidade = []

    tabela_infos = pd.DataFrame(index= ["CALL/PUT", "BUT/SELL", "STRIKE", "SPOT", "VOLATILIDADE", "JUROS", "DIAS P/ VENCIMENTO", "QUANTIDADE"])

    for i in range(1,5): # Cria listas com as variáveis
        tipo = config[f"TIPO{i}"]
        lado = config[f"LADO{i}"]
        strike = config[f"STRIKE{i}"]
        volatilidade = config[f"VOLATILIDADE{i}"]
        juros = config[f"JUROS{i}"]
        spot = config[f"SPOT{i}"]
        vencimento = config[f"VENCIMENTO{i}"]
        quantidade = config[f"QUANTIDADE{i}"]
        

        if tipo != "" and lado != "":
            strike = float(strike.replace(",", "."))
            juros = float(juros.replace(",", "."))
            volatilidade = float(volatilidade.replace(",", "."))
            spot = float(spot.replace(",", "."))
            vencimento = float(vencimento.replace(",","."))
            quantidade = float(quantidade.replace(",","."))
            

            tabela_infos[f"OPC{i}"] = [tipo, lado, strike, spot,  volatilidade,  juros,  vencimento, quantidade] 

            lista_tipo.append(tipo)
            lista_lado.append(lado)
            lista_strike.append(strike)
            lista_volatilidade.append(volatilidade)
            lista_juros.append(juros)
            lista_spot.append(spot)
            lista_vencimento.append(vencimento)
            lista_quantidade.append(quantidade)

    display(tabela_infos)
    print("-----------------------------------------------------------------")

    #LOOP 0
    lista_df_s_m = []
    lista_df_v_m = []
    lista_df_s_v_0 = []
    lista_df_s_v_1 = []
    lista_df_s_v_2 = []
    lista_df_s_v_3 = []
    lista_df_s_v_4 = []
    lista_df_s_v_5 = []

    lista_delta_s_m = []
    lista_delta_v_m = []
    lista_delta_s_v_0 = []
    lista_delta_s_v_1 = []
    lista_delta_s_v_2 = []
    lista_delta_s_v_3 = []
    lista_delta_s_v_4 = []
    lista_delta_s_v_5 = []

    lista_gamma_s_m = []
    lista_gamma_v_m = []
    lista_gamma_s_v_0 = []
    lista_gamma_s_v_1 = []
    lista_gamma_s_v_2 = []
    lista_gamma_s_v_3 = []
    lista_gamma_s_v_4 = []
    lista_gamma_s_v_5 = []

    for i in range(0,len(lista_tipo)): 
                
        flag = lista_tipo[i]
        lado = lista_lado[i]
        strike = lista_strike[i]
        volatilidade = lista_volatilidade[i]
        juros = lista_juros[i]
        spot = lista_spot[i]
        vencimento = lista_vencimento[i]
        quantidade = lista_quantidade[i]
        
        
        lista_choque_spot = [spot]
        lista_choque_volatilidade = [volatilidade]
        lista_choque_dias = [vencimento]

        lista_choque_volatilidade2 = [0]

        for y in range(1,5): #Cria lista com as Volatilidades e Spots dentro do range especificado

            choque_volatilidade_up = volatilidade * (1 + (choque_volatilidade * y))
            choque_volatilidade_down = volatilidade * (1 - (choque_volatilidade * y))

            choque_volatilidade_up2 = choque_volatilidade * y
            choque_volatilidade_down2 = choque_volatilidade * (-y)


            
            dim = int(vencimento/5)
            choque_dias = vencimento - (dim * y)

            lista_choque_volatilidade.append(choque_volatilidade_down)
            lista_choque_volatilidade.append(choque_volatilidade_up)
            lista_choque_dias.append(choque_dias)


            lista_choque_volatilidade2.append(choque_volatilidade_down2)
            lista_choque_volatilidade2.append(choque_volatilidade_up2)
            
        for y in range(1,9):
            choque_spot_up = spot * (1 + (choque_spot*y))
            choque_spot_down = spot * (1-(choque_spot*y))
            lista_choque_spot.append(choque_spot_down)
            lista_choque_spot.append(choque_spot_up)

        lista_choque_spot = sorted(lista_choque_spot)
        lista_choque_volatilidade = sorted(lista_choque_volatilidade)
        lista_choque_dias.append(1)
        lista_choque_dias = sorted(lista_choque_dias)

        lista_choque_volatilidade2 = sorted(lista_choque_volatilidade2)
        
        #####################################################################################################################
        #   Nesse ponto temos todas as váriaveis fixas (strike, dias para vencimento, lado, juros) e as listas de variáveis não-fixas (spot, volatilidade) 

        # PRÊMIO: Variando Maturity e Volatilidade. (Spot Constante)
        df_v_m = pd.DataFrame(columns=lista_choque_dias, index = lista_choque_volatilidade2)
        delta_v_m = pd.DataFrame(columns=lista_choque_dias, index = lista_choque_volatilidade2)
        gamma_v_m = pd.DataFrame(columns=lista_choque_dias, index = lista_choque_volatilidade2)
        for v, vol in enumerate(lista_choque_volatilidade):
            for m, maturity in enumerate(lista_choque_dias):
                df_v_m.iloc[v, m] = BS(spot, strike, maturity, juros, vol, flag)
                delta_v_m.iloc[v,m] = DELTA(spot, strike, maturity, juros, vol, flag)
                gamma_v_m.iloc[v,m] = GAMMA(spot, strike, maturity, juros, vol, flag)

        # PRÊMIO: Variando Maturity e Spot. (Volatilidade Constante)
        df_s_m = pd.DataFrame(columns=lista_choque_dias, index = lista_choque_spot)
        delta_s_m = pd.DataFrame(columns=lista_choque_dias, index = lista_choque_spot)
        gamma_s_m = pd.DataFrame(columns=lista_choque_dias, index = lista_choque_spot)
        for s, spot in enumerate(lista_choque_spot):
            for m, maturity in enumerate(lista_choque_dias):
                df_s_m.iloc[s, m] = BS(spot, strike, maturity, juros, volatilidade, flag)
                delta_s_m.iloc[s,m] = DELTA(spot, strike, maturity, juros, volatilidade, flag)
                gamma_s_m.iloc[s,m] = GAMMA(spot, strike, maturity, juros, volatilidade, flag)


        
        # PRÊMIO: Um gráfico por Maturity variando SPOT e VOLATILIDADE
        df_s_v_0 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        df_s_v_1 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        df_s_v_2 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        df_s_v_3 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        df_s_v_4 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        df_s_v_5 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        
        delta_s_v_0 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        delta_s_v_1 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        delta_s_v_2 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        delta_s_v_3 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        delta_s_v_4 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        delta_s_v_5 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)

        gamma_s_v_0 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        gamma_s_v_1 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        gamma_s_v_2 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        gamma_s_v_3 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        gamma_s_v_4 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
        gamma_s_v_5 = pd.DataFrame(columns=lista_choque_volatilidade2, index = lista_choque_spot)
    

        for m, maturity in enumerate(lista_choque_dias):
            for s, spot in enumerate(lista_choque_spot):
                for v, vol in enumerate(lista_choque_volatilidade):
                                     
                    if m == 0:
                        df_s_v_0.iloc[s, v] = BS(spot, strike, maturity, juros, vol, flag)
                        delta_s_v_0.iloc[s,v] = DELTA(spot, strike, maturity, juros, vol, flag)
                        gamma_s_v_0.iloc[s,v] = GAMMA(spot, strike, maturity, juros, vol, flag)

                    elif m == 1:
                        df_s_v_1.iloc[s, v] = BS(spot, strike, maturity, juros, vol, flag)
                        delta_s_v_1.iloc[s, v] = DELTA(spot, strike, maturity, juros, vol, flag)
                        gamma_s_v_1.iloc[s, v] = GAMMA(spot, strike, maturity, juros, vol, flag)

                    elif m == 2:
                        df_s_v_2.iloc[s, v] = BS(spot, strike, maturity, juros, vol, flag)
                        delta_s_v_2.iloc[s, v] = DELTA(spot, strike, maturity, juros, vol, flag)
                        gamma_s_v_2.iloc[s, v] = GAMMA(spot, strike, maturity, juros, vol, flag)
                        
                    elif m == 3:
                        df_s_v_3.iloc[s, v] = BS(spot, strike, maturity, juros, vol, flag)
                        delta_s_v_3.iloc[s, v] = DELTA(spot, strike, maturity, juros, vol, flag)
                        gamma_s_v_3.iloc[s, v] = GAMMA(spot, strike, maturity, juros, vol, flag)
                        
                    elif m == 4:
                        df_s_v_4.iloc[s, v] = BS(spot, strike, maturity, juros, vol, flag)
                        delta_s_v_4.iloc[s, v] = DELTA(spot, strike, maturity, juros, vol, flag)
                        gamma_s_v_4.iloc[s, v] = GAMMA(spot, strike, maturity, juros, vol, flag)
                        
                    elif m == 5:
                        # Dias Totais que Faltam
                        df_s_v_5.iloc[s, v] = BS(spot, strike, maturity, juros, vol, flag)
                        delta_s_v_5.iloc[s, v] = DELTA(spot, strike, maturity, juros, vol, flag)
                        gamma_s_v_5.iloc[s, v] = GAMMA(spot, strike, maturity, juros, vol, flag)
                        

        lista_df_s_m.append(df_s_m)
        lista_df_v_m.append(df_v_m)
        lista_df_s_v_0.append(df_s_v_0)
        lista_df_s_v_1.append(df_s_v_1)
        lista_df_s_v_2.append(df_s_v_2)
        lista_df_s_v_3.append(df_s_v_3)
        lista_df_s_v_4.append(df_s_v_4)
        lista_df_s_v_5.append(df_s_v_5)
        
        lista_delta_s_m.append(delta_s_m)
        lista_delta_v_m.append(delta_v_m)
        lista_delta_s_v_0.append(delta_s_v_0)
        lista_delta_s_v_1.append(delta_s_v_1)
        lista_delta_s_v_2.append(delta_s_v_2)
        lista_delta_s_v_3.append(delta_s_v_3)
        lista_delta_s_v_4.append(delta_s_v_4)
        lista_delta_s_v_5.append(delta_s_v_5)

        lista_gamma_s_m.append(gamma_s_m)
        lista_gamma_v_m.append(gamma_v_m)
        lista_gamma_s_v_0.append(gamma_s_v_0)
        lista_gamma_s_v_1.append(gamma_s_v_1)
        lista_gamma_s_v_2.append(gamma_s_v_2)
        lista_gamma_s_v_3.append(gamma_s_v_3)
        lista_gamma_s_v_4.append(gamma_s_v_4)
        lista_gamma_s_v_5.append(gamma_s_v_5)

        if i == 0:
            if lado == "BUY":
                table_s_m = lista_df_s_m[i]
                table_v_m = lista_df_v_m[i]
                table_s_v_0 = lista_df_s_v_0[i]
                table_s_v_1 = lista_df_s_v_1[i]
                table_s_v_2 = lista_df_s_v_2[i]
                table_s_v_3 = lista_df_s_v_3[i]
                table_s_v_4 = lista_df_s_v_4[i]
                table_s_v_5 = lista_df_s_v_5[i]

                delta_table_s_m = lista_delta_s_m[i]
                delta_table_v_m = lista_delta_v_m[i]
                delta_table_s_v_0 = lista_delta_s_v_0[i]
                delta_table_s_v_1 = lista_delta_s_v_1[i]
                delta_table_s_v_2 = lista_delta_s_v_2[i]
                delta_table_s_v_3 = lista_delta_s_v_3[i]
                delta_table_s_v_4 = lista_delta_s_v_4[i]
                delta_table_s_v_5 = lista_delta_s_v_5[i]

                gamma_table_s_m = lista_gamma_s_m[i]
                gamma_table_v_m = lista_gamma_v_m[i]
                gamma_table_s_v_0 = lista_gamma_s_v_0[i]
                gamma_table_s_v_1 = lista_gamma_s_v_1[i]
                gamma_table_s_v_2 = lista_gamma_s_v_2[i]
                gamma_table_s_v_3 = lista_gamma_s_v_3[i]
                gamma_table_s_v_4 = lista_gamma_s_v_4[i]
                gamma_table_s_v_5 = lista_gamma_s_v_5[i]

            else:
                table_s_m = -lista_df_s_m[i]
                table_v_m = -lista_df_v_m[i]
                table_s_v_0 = -lista_df_s_v_0[i]
                table_s_v_1 = -lista_df_s_v_1[i]
                table_s_v_2 = -lista_df_s_v_2[i]
                table_s_v_3 = -lista_df_s_v_3[i]
                table_s_v_4 = -lista_df_s_v_4[i]
                table_s_v_5 = -lista_df_s_v_5[i]

                delta_table_s_m = -lista_delta_s_m[i]
                delta_table_v_m = -lista_delta_v_m[i]
                delta_table_s_v_0 = -lista_delta_s_v_0[i]
                delta_table_s_v_1 = -lista_delta_s_v_1[i]
                delta_table_s_v_2 = -lista_delta_s_v_2[i]
                delta_table_s_v_3 = -lista_delta_s_v_3[i]
                delta_table_s_v_4 = -lista_delta_s_v_4[i]
                delta_table_s_v_5 = -lista_delta_s_v_5[i]

                gamma_table_s_m = -lista_gamma_s_m[i]
                gamma_table_v_m = -lista_gamma_v_m[i]
                gamma_table_s_v_0 = -lista_gamma_s_v_0[i]
                gamma_table_s_v_1 = -lista_gamma_s_v_1[i]
                gamma_table_s_v_2 = -lista_gamma_s_v_2[i]
                gamma_table_s_v_3 = -lista_gamma_s_v_3[i]
                gamma_table_s_v_4 = -lista_gamma_s_v_4[i]
                gamma_table_s_v_5 = -lista_gamma_s_v_5[i]

        else:
            if lado == "BUY":
                table_s_m += lista_df_s_m[i]
                table_v_m += lista_df_v_m[i]
                table_s_v_0 += lista_df_s_v_0[i]
                table_s_v_1 += lista_df_s_v_1[i]
                table_s_v_2 += lista_df_s_v_2[i]
                table_s_v_3 += lista_df_s_v_3[i]
                table_s_v_4 += lista_df_s_v_4[i]
                table_s_v_5 += lista_df_s_v_5[i]

                delta_table_s_m += lista_delta_s_m[i]
                delta_table_v_m += lista_delta_v_m[i]
                delta_table_s_v_0 += lista_delta_s_v_0[i]
                delta_table_s_v_1 += lista_delta_s_v_1[i]
                delta_table_s_v_2 += lista_delta_s_v_2[i]
                delta_table_s_v_3 += lista_delta_s_v_3[i]
                delta_table_s_v_4 += lista_delta_s_v_4[i]
                delta_table_s_v_5 += lista_delta_s_v_5[i]
                
                gamma_table_s_m += lista_gamma_s_m[i]
                gamma_table_v_m += lista_gamma_v_m[i]
                gamma_table_s_v_0 += lista_gamma_s_v_0[i]
                gamma_table_s_v_1 += lista_gamma_s_v_1[i]
                gamma_table_s_v_2 += lista_gamma_s_v_2[i]
                gamma_table_s_v_3 += lista_gamma_s_v_3[i]
                gamma_table_s_v_4 += lista_gamma_s_v_4[i]
                gamma_table_s_v_5 += lista_gamma_s_v_5[i]


            else:
                table_s_m -= lista_df_s_m[i]
                table_v_m -= lista_df_v_m[i]
                table_s_v_0 -= lista_df_s_v_0[i]
                table_s_v_1 -= lista_df_s_v_1[i]
                table_s_v_2 -= lista_df_s_v_2[i]
                table_s_v_3 -= lista_df_s_v_3[i]
                table_s_v_4 -= lista_df_s_v_4[i]
                table_s_v_5 -= lista_df_s_v_5[i]

                delta_table_s_m -= lista_delta_s_m[i]
                delta_table_v_m -= lista_delta_v_m[i]
                delta_table_s_v_0 -= lista_delta_s_v_0[i]
                delta_table_s_v_1 -= lista_delta_s_v_1[i]
                delta_table_s_v_2 -= lista_delta_s_v_2[i]
                delta_table_s_v_3 -= lista_delta_s_v_3[i]
                delta_table_s_v_4 -= lista_delta_s_v_4[i]
                delta_table_s_v_5 -= lista_delta_s_v_5[i]

                gamma_table_s_m -= lista_gamma_s_m[i]
                gamma_table_v_m -= lista_gamma_v_m[i]
                gamma_table_s_v_0 -= lista_gamma_s_v_0[i]
                gamma_table_s_v_1 -= lista_gamma_s_v_1[i]
                gamma_table_s_v_2 -= lista_gamma_s_v_2[i]
                gamma_table_s_v_3 -= lista_gamma_s_v_3[i]
                gamma_table_s_v_4 -= lista_gamma_s_v_4[i]
                gamma_table_s_v_5 -= lista_gamma_s_v_5[i]


    # Colando no Template  
    book = load_workbook("PATH DO TEMPLATE")
    writer = pd.ExcelWriter(f"PATH DE SALVAMENTO", engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

    tabela_titulo.to_excel(writer, "Payoff & Gregas", startrow=0, startcol=0, header=False, index=True)
    tabela_infos.to_excel(writer, "Payoff & Gregas", startrow=5, startcol=0, header=False, index=True)
    
    table_s_m.sort_index(axis=1, ascending = False, inplace=True)
    table_v_m.sort_index(axis=1, ascending = False, inplace=True)
    delta_table_s_m.sort_index(axis=1, ascending = False, inplace=True)
    delta_table_v_m.sort_index(axis=1, ascending = False, inplace=True)
    gamma_table_s_m.sort_index(axis=1, ascending = False, inplace=True)
    gamma_table_v_m.sort_index(axis=1, ascending = False, inplace=True)
    
    
    table_s_m.to_excel(writer, "Premios", startrow=25, startcol=3, header=True, index=True)
    table_v_m.to_excel(writer, "Premios", startrow=46, startcol=3, header=True, index=True)
    table_s_v_5.to_excel(writer, "Premios", startrow=67, startcol=3, header=True, index=True)
    table_s_v_4.to_excel(writer, "Premios", startrow=88, startcol=3, header=True, index=True)
    table_s_v_3.to_excel(writer, "Premios", startrow=109, startcol=3, header=True, index=True)
    table_s_v_2.to_excel(writer, "Premios", startrow=130, startcol=3, header=True, index=True)
    table_s_v_1.to_excel(writer, "Premios", startrow=151, startcol=3, header=True, index=True)
    table_s_v_0.to_excel(writer, "Premios", startrow=172, startcol=3, header=True, index=True)
    
    delta_table_s_m.to_excel(writer, "Payoff & Gregas", startrow=26, startcol=34, header=False, index=True)
    delta_table_v_m.to_excel(writer, "Payoff & Gregas", startrow=47, startcol=34, header=False, index=True)
    delta_table_s_v_5.to_excel(writer, "Payoff & Gregas", startrow=67, startcol=34, header=True, index=True)
    delta_table_s_v_4.to_excel(writer, "Payoff & Gregas", startrow=88, startcol=34, header=True, index=True)
    delta_table_s_v_3.to_excel(writer, "Payoff & Gregas", startrow=109, startcol=34, header=True, index=True)
    delta_table_s_v_2.to_excel(writer, "Payoff & Gregas", startrow=130, startcol=34, header=True, index=True)
    delta_table_s_v_1.to_excel(writer, "Payoff & Gregas", startrow=151, startcol=34, header=True, index=True)
    delta_table_s_v_0.to_excel(writer, "Payoff & Gregas", startrow=172, startcol=34, header=True, index=True)
    
    gamma_table_s_m.to_excel(writer, "Payoff & Gregas", startrow=26, startcol=65, header=False, index=True)
    gamma_table_v_m.to_excel(writer, "Payoff & Gregas", startrow=47, startcol=65, header=False, index=True)
    gamma_table_s_v_5.to_excel(writer, "Payoff & Gregas", startrow=67, startcol=65, header=True, index=True)
    gamma_table_s_v_4.to_excel(writer, "Payoff & Gregas", startrow=88, startcol=65, header=True, index=True)
    gamma_table_s_v_3.to_excel(writer, "Payoff & Gregas", startrow=109, startcol=65, header=True, index=True)
    gamma_table_s_v_2.to_excel(writer, "Payoff & Gregas", startrow=130, startcol=65, header=True, index=True)
    gamma_table_s_v_1.to_excel(writer, "Payoff & Gregas", startrow=151, startcol=65, header=True, index=True)
    gamma_table_s_v_0.to_excel(writer, "Payoff & Gregas", startrow=172, startcol=65, header=True, index=True)
    
    writer.close()
    print("FIM")
op()


