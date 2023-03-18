# 方法一：使用tabula库，可以抽出多张表，效果还不错
import os
import re

import tabula
import pandas as pd
from pandas import DataFrame

from common.utils import read_excel

path = os.getcwd()


def df_to_excel(df, save_path):
    df.to_excel(save_path, index=False)
    # df.to_excel(save_path, index=False, header=False)


def pdf2excel(file_path, save_path, in_filename, out_filename):
    file_path = f"{file_path}/{in_filename}"
    # Read the PDF file and extract the table
    tables = tabula.read_pdf(file_path, pages='all', multiple_tables=True)

    # Convert the extracted table to a df_list
    df_list = []
    for item in tables:
        df_list.append(item)
    output_file_path = f"{save_path}/{out_filename}"

    # for i, df in enumerate(df_list):
    #     with pd.ExcelWriter(output_file_path) as writer:
    #         sheet_name = f"Sheet{i + 1}"
    #         df.to_excel(writer, sheet_name=sheet_name, index=False)
    #         # writer.save()

    with pd.ExcelWriter(output_file_path) as writer:
        try:
            for i, df in enumerate(df_list):  # 列名 row: ('评级结果:', '主体长期信用等级:AAA') ('评级观点', '联合资信评估股份有限公司(以下简称“联合资信”)对')  行名 index: 0
                sheet_name = f"Sheet{i + 1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                # for index, row in df.iterrows():
                #     if '产量' in str(df.iloc[index][0]):
                #         sheet_name = f"Sheet{i + 1}"
                #         df.to_excel(writer, sheet_name=sheet_name, index=False)
                #         break
        except IndexError:
            print("IndexError: At least one sheet must be visible")


if __name__ == '__main__':
    file_path = os.path.abspath(os.path.join(path, "../../file/file_path"))
    # print(file_path)  # /Users/pc/PycharmProjects/IE/common/file/file_path
    save_path = os.path.abspath(os.path.join(path, "../../output_file/save_path"))

    example_dir = os.path.abspath(os.path.join(path, "../../file/example/example.xlsx"))
    company_code = read_excel.read_excel(example_dir)
    company_code.pop('公司简称')
    code_company = {v: k for k, v in company_code.items()}

    pattern = r'\d{3,12}\.\w{2}'  # 匹配 3-9 位数字、一个点号和 2 个字母数字字符的字符串
    files = os.listdir(file_path)  # 得到文件夹下所有的文件名称

    df_columns = ['序号', '公司简称', '债券代码', '全员工效', '吨煤成本', '吨煤均价', '商品煤销量', '商品煤产量', '来源报告']
    df = pd.DataFrame(columns=df_columns)
    for i, file in enumerate(files):  # 遍历文件夹
        match = re.search(pattern, file)
        if match:
            company_name = code_company[match.group()]
            com_code = company_code[company_name]
        else:
            company_name = f'{file}'
            com_code = f'tem{i}'
        dic_ = {'序号': i,
                '公司简称': company_name,
                '债券代码': com_code,
                '来源报告': file}
        df = df.append(dic_, ignore_index=True)
    res_dir = os.path.abspath(os.path.join(path, "../../output_file/result/0317_res.xlsx"))
    df_to_excel(df, res_dir)
    # for i, file in enumerate(files):  # 遍历文件夹
    #     match = re.search(pattern, file)
    #     if match:
    #         company_name = code_company[match.group()]
    #         out_filename = f'{i}-{company_name}-{company_code[company_name]}.xlsx'
    #     else:
    #         out_filename = f'{i}-{file}.xlsx'
    #         print('未匹配到')
    #     pdf2excel(file_path, save_path, file, out_filename)
