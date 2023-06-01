#%%
import pandas as pd
from prettytable import PrettyTable

def create_table_from_excel(excel_file_path):
    # 讀取Excel檔案
    df = pd.read_excel(excel_file_path)

    # 建立PrettyTable物件
    table = PrettyTable()

    # 設定表格欄位名稱
    table.field_names = df.columns

    # 將資料加入表格
    for _, row in df.iterrows():
        table.add_row(row)

    # 設定表格的框線樣式
    table.border = True

    # 設定表格內容對齊方式
    table.align = 'l'

    return table.get_string()

# 指定Excel檔案路徑
excel_file_path = './test_worksheet.xlsx'

# 建立表格並輸出成文字檔案
table = create_table_from_excel(excel_file_path)
with open('output2.txt', 'w') as file:
    file.write(table)



# %%
