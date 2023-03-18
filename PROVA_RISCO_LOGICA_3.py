# LÓGICA: QUESTÃO 3
"""
Um cassino oferece um jogo no qual o jogador joga uma moeda justa. Se a moeda cair cara, o jogador ganha R$ 1. Se a moeda cair coroa, o jogador perde R$ 1. O jogador começa com R$ 10 e decide jogar 100 vezes. Qual é a probabilidade de que o jogador tenha pelo menos R$ 15 no final?

P(X ≥ k) = (n¦k) * p^k * (1-p)^(n-k)
# P(X≥53) = P(x=53) + P(x=54) + P(x=55) + ⋯ + P(x=100)
"""
from scipy.stats import binom  

n =100
p = 0.5
result = 0

#FORMA 1
for x in range(53,101):
    p_x = binom.pmf(x,n,p)
    result += p_x

print(f"Resultado forma 1: {result}")

#FORMA 2
x = range(53,101)
p_x = (binom.pmf(x,n,p))
result = sum(p_x)
print(f"Resultado forma 2: {result}")