#DESAFIO: QUESTÃO 3

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