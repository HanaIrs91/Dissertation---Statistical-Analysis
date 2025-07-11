#Spearman Correlation, Hanna Arshid, 10//07/2025
import pandas as pd
import numpy as np
import math

file_path = "/Users/hanairshaid/Desktop/Likert Scale Answers.xlsx"
df = pd.read_excel(file_path)

mapping = {
    'Strongly Disagree': 1,
    'Disagree': 2,
    'Neutral': 3,
    'Agree': 4,
    'Strongly Agree': 5
}

df_numeric = df.replace(mapping)

group1_cols = df_numeric.columns[:12]
group2_cols = df_numeric.columns[12:23]

df_numeric['Group1_Mean'] = df_numeric[group1_cols].mean(axis=1)
df_numeric['Group2_Mean'] = df_numeric[group2_cols].mean(axis=1)

def spearman_corr(x, y):
    rx = pd.Series(x).rank()
    ry = pd.Series(y).rank()
    return rx.corr(ry)

def t_dist_sf(t, df): 
    def betacf(a, b, x, max_iter=200, tol=1e-10):
        """ Continued fraction for incomplete beta function """
        am, bm = 1.0, 1.0
        az = 1.0
        qab = a + b
        qap = a + 1.0
        qam = a - 1.0
        bz = 1.0 - qab * x / qap
        for m in range(1, max_iter + 1):
            em = float(m)
            tem = em + em
            d = em * (b - m) * x / ((qam + tem) * (a + tem))
            ap = az + d * am
            bp = bz + d * bm
            d = -(a + em) * (qab + em) * x / ((a + tem) * (qap + tem))
            app = ap + d * az
            bpp = bp + d * bz
            aold = az
            am = ap / bpp
            bm = bp / bpp
            az = app / bpp
            bz = 1.0
            if abs(az - aold) < tol * abs(az):
                return az
        return az

    def betai(a, b, x):
        if x == 0.0 or x == 1.0:
            return x
        ln_beta = math.lgamma(a) + math.lgamma(b) - math.lgamma(a + b)
        front = math.exp(a * math.log(x) + b * math.log(1 - x) - ln_beta) / a
        cf = betacf(a, b, x)
        return front * cf

    x = df / (df + t * t)
    ib = betai(df/2, 0.5, x)
    return 0.5 * ib

def spearman_pvalue(rho, n):
    t_stat = rho * np.sqrt((n - 2) / (1 - rho**2))
    p = 2 * t_dist_sf(abs(t_stat), n - 2)
    return p

corr = spearman_corr(df_numeric['Group1_Mean'], df_numeric['Group2_Mean'])
n = len(df_numeric)
p_value = spearman_pvalue(corr, n)

print("Spearman correlation:", corr)
print("Approximate p-value:", p_value)
