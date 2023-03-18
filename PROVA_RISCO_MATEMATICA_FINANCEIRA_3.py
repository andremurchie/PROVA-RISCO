# MATEMÁTICA FINANCEIRA: QUESTÃO 3
"""
Uma opção de compra tem um preço de exercício de R$ 50, um prazo de 6 meses e uma volatilidade implícita de 20%. Qual é o valor teórico da opção se o ativo subjacente estiver sendo negociado a R$ 60? 
"""

import numpy as np
from scipy.stats import norm

def BS(spot, strike, time, juros, volatilite, flag):
    N = norm.cdf

    if flag == "CALL":
        flag = 1
    else:
        flag = -1
    
    time = time/252

    vp_k = strike/(1+juros)**time
    d1 = np.log(spot/vp_k) / (volatilite * np.sqrt(time)) + (volatilite * np.sqrt(time))/2
    d2 = np.log(spot/vp_k) / (volatilite * np.sqrt(time)) - (volatilite * np.sqrt(time))/2
    premio = flag * spot * N(flag * d1) - flag * vp_k * N(flag * d2) 
    
    return (f"O valor justo da opção é: {premio}")

spot = 60
strike = 50
time = 126
juros = 0.1375
volatilite = 0.2
flag = "CALL"

BS(spot, strike, time, juros, volatilite, flag)