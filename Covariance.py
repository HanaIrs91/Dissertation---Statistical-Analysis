#Covariance, Hanna Arshid, 10//07/2025
import pandas as pd

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

group1_cols = [f for f in df_numeric.columns[:12]]
group2_cols = [f for f in df_numeric.columns[12:23]]

df_numeric['Group1_Mean'] = df_numeric[group1_cols].mean(axis=1)
df_numeric['Group2_Mean'] = df_numeric[group2_cols].mean(axis=1)

covariance = df_numeric[['Group1_Mean', 'Group2_Mean']].cov().iloc[0,1]

print(covariance)
