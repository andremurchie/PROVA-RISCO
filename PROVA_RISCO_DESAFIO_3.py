#DESAFIO: QUESTÃO 3

"""
A Dual Digital Option é uma opção exótica que envolve dois ativos subjacentes e é paga quando ambos atingem as barreiras de preço. Especificamente, a opção paga um valor fixo se os preços dos dois ativos subjacentes estiverem dentro de um intervalo específico de preços em um determinado momento do tempo. Para que o pagamento seja realizado, ambos os preços dos ativos subjacentes devem "tocar" as barreiras de preço simultaneamente durante a vigência da opção. 

A precificação da Dual Digital Option é um processo matemático complexo que envolve a modelagem da dinâmica dos preços dos dois ativos subjacentes, levando em consideração a volatilidade, tempo até a expiração, correlação entre os preços dos ativos subjacentes e taxas de juros livres de risco. Suponha que você deseja precificar, antes do vencimento, a opção exótica determinada acima. Escreva um código em Python para calcular o preço teórico dessa opção. O output deve ser o preço teórico de uma determinada opção em um determinada tempo.


MODELO DE PRECIFICAÇÃO DUAL DIGITAL OPTION
Bibliografia: Opções, Futuros e Outros Derivativos, HULL
            : Mercado de Derivativos no Brasil, BESSADA

Fórmulas:
u = e^(vol*sqtr(delta-t))  ... taxa de subida do spot
d = e^(-vol*sqtr(delta-t)) ou d = 1/u ... taxa de decaimento do spot
a = e^(r*delta-t) ... apreçamento da taxa livre de risco

p = (a-d) / (u-d) ... probabilidade neutra ao risco

f = 0 ou S-K ... para qualquer um dos últimos nós ...
f = (e^(-r*delta-t)) * [p * fu + (1-p)fd] ... para cada nó do meio e nó inicial
"""
import pandas as pd
import numpy as np
import math
from scipy.special import comb

def digital_option(spot, strike, volatilidade, juros, pagamento, maturidade, passos):
    
    preco_justo = 0

    #Fórmulas Binomiais
    delta_t = (maturidade/252)/(passos)
    u = math.exp(volatilidade*np.sqrt(delta_t))
    d = 1/u
    a = math.exp(juros*delta_t)
    q = (a-d)/(u-d)
    print(u,d,q, delta_t)

    # Construindo a árvore binomial
    arv_b = [[0.0 for j in range(1+i)] for i in range (passos+1)]
    arv_probabilidade = [[0.0 for j in range(1+i)] for i in range (passos+1)]
    

    for j in range(passos+1):
        # Calcula a probabilidade final de cada nó binomial
        arv_probabilidade[passos][j] = comb(passos, j) * (1-q)**(passos-j) * (q)**(j)
    
        #Para cada valor final da arvore binomial: se esse valor for maior que o strike, paga-se o prêmio fixo, caso contrário a opção valerá 0
        if(spot * u**j * d**(passos-j) - strike) > 0: 
            arv_b[passos][j] = 10
            
        else:
            arv_b[passos][j] = 0

        # Média ponderada entre a probabilidade de chegar no nó final correspondente com o valor do prêmio nesse nó
        preco_justo += arv_b[passos][j] * arv_probabilidade[passos][j]
        print(arv_b[passos][j], arv_probabilidade[passos][j])    
    
    preco_justo = preco_justo * math.exp(-juros*delta_t)
    
    return f"O Preço Justo calculado foi de {preco_justo}"


# Mercado de Derivativos no Brasil, BESSADA - Pag. 289 Exemplo 1
spot = 100
strike = 100
volatilidade = 0.2
juros = 0.12 #a.a.
pagamento = 10
maturidade = 189 #dias úteis
passos = 3

digital_option(spot, strike, volatilidade, juros, pagamento, maturidade, passos)