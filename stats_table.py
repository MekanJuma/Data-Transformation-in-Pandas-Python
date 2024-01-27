import pandas as pd

input_file = '/content/data.xlsx'
sheet_name = 'Query Data'
use_cols = [
    'merchant_customer_id', 'shipped_units', 'sfp_dea_decision', 
    'sfp_cr_decision', 'sfp_vtr_decision', 'sfp_gvs_std_decision', 
    'sfp_gvs_os_decision', 'sfp_gvs_xl_decision', 'gms'
]
main_df = pd.read_excel(input_file, sheet_name=sheet_name, usecols=use_cols)

metric_columns = [
    'sfp_dea_decision', 'sfp_cr_decision', 'sfp_vtr_decision', 
    'sfp_gvs_std_decision', 'sfp_gvs_os_decision', 'sfp_gvs_xl_decision'
]
main_df['Action'] = main_df[metric_columns].apply(lambda x: x.first_valid_index(), axis=1)
main_df['Metric'] = main_df['Action'].str.extract(r'_(\w+)_decision')
main_df['Action'] = main_df['Action'].str.extract(r'(\w+)_')

agg_funcs = {
    'merchant_customer_id': 'nunique', 
    'shipped_units': 'sum', 
    'gms': 'sum'
}
grouped = main_df.groupby(['Metric', 'Action']).agg(agg_funcs).reset_index()

# Calculate percentages
total = grouped.groupby('Metric').sum().rename(lambda x: x + ' Total')
grouped = grouped.merge(total, on='Metric', suffixes=('', ' Total'))
for col in ['merchant_customer_id', 'shipped_units', 'gms']:
    grouped[col + ' %'] = grouped[col] / grouped[col + ' Total']

# Reorder and clean up columns
cols_order = [
    'Metric', 'Action', 'merchant_customer_id', 'merchant_customer_id %', 
    'shipped_units', 'shipped_units %', 'gms', 'gms %'
]
grouped = grouped[cols_order]

# Add totals and percentages for overall data
overall_totals = grouped[grouped['Action'] == 'Total']
grouped = pd.concat([overall_totals, grouped[grouped['Action'] != 'Total']])

grouped.to_excel('/content/output.xlsx', index=False)
