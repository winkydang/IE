# 方法一：使用tabula库，可以抽出多张表，效果还不错
import re

import tabula
import pandas as pd

in_filename = '山东能源集团有限公司2022年面向专业投资者公开发行可续期公司债券(第一期)信用评级报告.pdf'
file_path = f"../../file/{in_filename}"

# Read the PDF file and extract the table
tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
df = tables  # Select the first dataframe containing the table

# Convert the extracted table to a df_list
df_list = []
for item in df:
    df_list.append(item)

# save table to excel
out_filename = '01-煤炭-山东能源集团有限公司.xlsx'
output_file_path = f"../../output_file/{out_filename}"

# Write the dataframes to different sheets of the same Excel file
with pd.ExcelWriter(output_file_path) as writer:
    for i, df in enumerate(df_list):
        for index, row in df.iterrows():  # row: ('评级结果:', '主体长期信用等级:AAA') ('评级观点', '联合资信评估股份有限公司(以下简称“联合资信”)对')  index: 0
            if '吨煤' in str(df.iloc[index][0]):
                sheet_name = f"Sheet {i + 1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                break




