# # 一个excel也没抽出来
# import PyPDF2
# import pandas as pd
#
# # Open the PDF file and read its contents
# pdf_file = open('../../file/山东能源集团有限公司2022年面向专业投资者公开发行可续期公司债券(第一期)信用评级报告.pdf', 'rb')
# pdf_reader = PyPDF2.PdfReader(pdf_file)
# # page = pdf_reader.getPage(0)
# page = pdf_reader.pages[0]
# page_content = page.extract_text()
#
# # Split the content by new line and remove empty lines
# content_lines = [line for line in page_content.split('\n') if line.strip()]
#
# # Find the table header and rows
# table_header = content_lines[0].split('\t')
# table_rows = [row.split('\t') for row in content_lines[1:]]
#
# # Convert the table rows into a pandas dataframe
# df = pd.DataFrame(table_rows, columns=table_header)
#
# # Output the dataframe into Excel format
# df.to_excel('../../output_file/method2-煤炭-山东能源集团有限公司.xlsx', index=False)
#

import re

s = '2022-07-29-012105084.IB-21潞安化工SCP005-潞安化工集团有限公司主体长期信用评级报告.pdf'
pattern = r'\d{9}\.\w{2}'

match = re.search(pattern, s)
if match:
    print(match.group())
else:
    print('未匹配到')
