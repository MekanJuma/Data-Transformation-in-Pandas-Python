import pandas as pd


FILE_2023 = '/content/2023.xlsx'
FILE_2024 = '/content/2024.xlsx'
SHEET_NAME = 'Summary'
HEADER_ROW_2023 = 0
HEADER_ROW_2024 = 1
COLS = ['Week', 'Total sellers reviewed', 'Sellers passed trial']

def read_and_process_data(file_name, year, header_row=None):
    df = pd.read_excel(file_name, sheet_name=SHEET_NAME, header=header_row)
    df = (df.dropna(subset=['Week'])
          .query("Week != 'Total'")
          .replace({'Week': 53}, 1)
          [COLS])
    df['Sellers passed trial'] = df['Sellers passed trial'].fillna(0)
    if year == 2023:
        df['Year'] = df['Week'].apply(lambda x: year+1 if x in [1, 2] else year)
    else:
        df['Year'] = year
    
    return df

df_2023 = read_and_process_data(FILE_2023, 2023, HEADER_ROW_2023)
df_2024 = read_and_process_data(FILE_2024, 2024, HEADER_ROW_2024)

df_union = pd.concat([df_2023, df_2024]).sort_values(by=['Year', 'Week'])

df_grouped = (df_union.groupby(['Year', 'Week'])
                      .sum()
                      .assign(Cumulated_grad_total=lambda x: x['Sellers passed trial'].cumsum())
                      .rename(columns={'Total sellers reviewed': '# of Completed', 
                                       'Sellers passed trial': '# of Graduated'}))


df_grouped = df_grouped.reset_index()
df_grouped.to_excel('/content/output.xlsx', index=False)