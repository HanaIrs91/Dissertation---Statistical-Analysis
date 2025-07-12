#Sample variance, Hanna Arshid, 12/07/2025
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.font_manager as fm

file_path = "/Users/hanairshaid/Desktop/Likert Scale Answers.xlsx"
df = pd.read_excel(file_path)

likert_scale_order = ['Strongly Disagree', 'Disagree', 'Neutral', 'Agree', 'Strongly Agree']
mapping = {
    'Strongly Disagree': 1,
    'Disagree': 2,
    'Neutral': 3,
    'Agree': 4,
    'Strongly Agree': 5
}
df_numeric = df.replace(mapping)

font_path = "/System/Library/Fonts/Supplemental/Times New Roman.ttf"
font_properties = fm.FontProperties(fname=font_path)

variances = {}
for question in df_numeric.columns:
    series = df_numeric[question].dropna()
    variances[question] = series.var()

question_labels = [f"Q{i+1}" for i in range(len(df_numeric.columns))]
variance_values = list(variances.values())

plt.figure(figsize=(10, 6))
sns.barplot(x=variance_values, y=question_labels, palette="coolwarm", orient='h')
plt.xlabel("Sample Variance", fontproperties=font_properties)
plt.title("Sample Variance per Question", fontproperties=font_properties)
plt.xticks(fontproperties=font_properties)
plt.yticks(fontproperties=font_properties)
plt.tight_layout(rect=[0, 0, 1, 0.88])
plt.show()

variance_df = pd.DataFrame({
    "Question": question_labels,
    "Sample Variance": variance_values
})
output_path = "/Users/hanairshaid/Desktop/Sample Variance.xlsx"
variance_df.to_excel(output_path, index=False)