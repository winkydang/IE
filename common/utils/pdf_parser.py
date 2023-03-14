# 方法一：使用tabula库，可以抽出多张表，效果还不错
import tabula
import pandas as pd

file_path = "../../file/山东能源集团有限公司2022年面向专业投资者公开发行可续期公司债券(第一期)信用评级报告.pdf"

# Read the PDF file and extract the table
tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)
df = tables  # Select the first dataframe containing the table

# # Convert the extracted table to a pands dataframe
# df = pd.concat(df)

# Convert the extracted table to a df_list
df_list = []
for item in df:
    df_list.append(item)

# save table to excel
output_file_path = "../../output_file/method1-02-煤炭-山东能源集团有限公司.xlsx"

# Write the dataframes to different sheets of the same Excel file
with pd.ExcelWriter(output_file_path) as writer:
    for i, df in enumerate(df_list):
        sheet_name = f"Sheet {i+1}"
        df.to_excel(writer, sheet_name=sheet_name, index=False)  # 注：抽到了我想要的内容，在sheet16里

# # save the table to Excel format
# df.to_excel(output_file_path, index=False)

# for i, item in enumerate(df):
#     item.to_excel(output_file_path, sheet_name=f'sheet{i}', index=False)  # the 'index=False' argument is used to remove the row numbers from the output Excel file

