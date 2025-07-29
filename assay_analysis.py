import pandas as pd
import seaborn as sns
import numpy as np
import matplotlib.pyplot as plt

# load sheets into DataFrames
xls = pd.ExcelFile(r"C:\Users\ganatrasr\Downloads\AAIC_PowerBI_Tables - Copy.xlsx")
concordance_df = xls.parse('Concordance')
apoe_df = xls.parse('Assay Results')
etiology_df = xls.parse('Primary Etiology')

# merge concodrance with APOE using pt ID
merged_df_1 = concordance_df.merge(apoe_df, on='Subject ID', how='inner')

# merge etiology and risk score
merged_df_2 = merged_df_1.merge(etiology_df, on='Subject ID', how='inner')

merged_df_2['E4 Carrier'] = merged_df_2['Neurocode APOE'].apply(
    lambda x: 'Carrier' if '4' in str(x).replace(' ', '').upper() else 'Non-Carrier'
)

# drop rows with missing values in key columns
merged_df_2.dropna(subset=['Concordance', 'Neurocode APOE', 'Neurocode AD Risk Score', 'Primary Etiology (Per referral)_x'], inplace=True)

# melt data for all -217 assay results
assay_cols = ['Lucent -217 Result', 'ALZPath -217', 'Quest -217 Result']
melted_df = merged_df_2.melt(
    id_vars = 'E4 Carrier',
    value_vars = assay_cols,
    var_name = 'Assay',
    value_name = 'Result Category' # holds 'positive', etc.
)

melted_df = melted_df.dropna(subset=['Result Category'])

melted_df['E4 Carrier'] = melted_df['E4 Carrier'].str.strip().str.title()

grouped = (
    melted_df.groupby(['Assay', 'Result Category', 'E4 Carrier'])
    .size()
    .reset_index(name='Count')
)

# pivot to get stacked values
pivot_df = grouped.pivot_table(
    index=['Assay', 'Result Category'],
    columns='E4 Carrier',
    values='Count',
    fill_value=0
).reset_index()

bar_width = 0.25
result_order = ['Positive', 'Intermediate', 'Negative']
assays = pivot_df['Assay'].unique()

import matplotlib.pyplot as plt

# format assay names - remove ' Result' and extra spaces
pivot_df['Assay'] = (
    pivot_df['Assay']
    .str.replace(' Result', '', regex=False)
    .str.replace(' ', '', regex=False)
)

# ensure standard casing
pivot_df['Result Category'] = pivot_df['Result Category'].str.title()
result_order = ['Positive', 'Intermediate', 'Negative']
assays = pivot_df['Assay'].unique()

# create subplots
fig, axes = plt.subplots(1, 3, figsize=(15, 6), sharey=True)

for i, assay in enumerate(assays):
    ax = axes[i]
    assay_df = pivot_df[pivot_df['Assay'] == assay]

    # dynamically determine available result categories
    result_order = ['Positive', 'Intermediate', 'Negative']
    available_results = [res for res in result_order if res in assay_df['Result Category'].values]

    for j, result in enumerate(available_results):
        row = assay_df[assay_df['Result Category'] == result]
        carrier = row['Carrier'].values[0] if 'Carrier' in row and not row.empty else 0
        noncarrier = row['Non-Carrier'].values[0] if 'Non-Carrier' in row and not row.empty else 0
        total = carrier + noncarrier

        x = j
        bar1 = ax.bar(x, carrier, color="#0b8dea")
        bar2 = ax.bar(x, noncarrier, bottom=carrier, color="#e20cc2")

        # add percent labels
        if total > 0:
            if carrier > 0:
                ax.text(x, carrier / 2, f"{carrier/total:.0%}", ha='center', va='center', fontsize=9, color='white')
            if noncarrier > 0:
                ax.text(x, carrier + noncarrier / 2, f"{noncarrier/total:.0%}", ha='center', va='center', fontsize=9, color='white')

    # set x-axis
    ax.set_xticks(range(len(available_results)))
    ax.set_xticklabels(available_results)
    ax.set_title(f"{assay}")
    if i == 0:
        ax.set_ylabel("Participant Count")

# final plot formatting
fig.suptitle("pTau-217 Assay Results by APOE ε4 Carrier Status", fontsize=15)
blue_patch = plt.matplotlib.patches.Patch(color="#0b8dea", label='Carrier')
pink_patch = plt.matplotlib.patches.Patch(color="#e20cc2", label='Non-Carrier')
fig.legend(handles=[blue_patch, pink_patch], loc='upper right', ncol=2)
plt.tight_layout(rect=[0, 0, 1, 0.95])  # leave space for title
# plt.show()

# ------ plot #2 ------

sheet1_df = pd.read_excel(xls, sheet_name="Sheet1")

# clean column names: strip whitespace and fix typo
sheet1_df.columns = sheet1_df.columns.str.strip()
sheet1_df.rename(columns={"Diagnostic Certainity": "Diagnostic Certainty"}, inplace=True)

# ensure 'Primary Etiology' column is string
sheet1_df['Primary Etiology (Per referral)'] = sheet1_df['Primary Etiology (Per referral)'].astype(str)

# drop missing values
plot_df = sheet1_df.dropna(subset=['Primary Etiology (Per referral)', 'Diagnostic Certainty'])

# plot
plt.figure(figsize=(10, 6))
sns.violinplot(
    data=plot_df,
    x='Primary Etiology (Per referral)',
    y='Diagnostic Certainty',
    inner='box',
    palette='Set2'
)
plt.xticks(rotation=0)
plt.title("Diagnostic Certainty by Primary Etiology at Referral")
plt.tight_layout()
plt.show()
