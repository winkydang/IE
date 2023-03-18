# 2023.02.23 周三 读取excel中合并的单元格内容
import pandas as pd
import numpy as np
import xlrd
from openpyxl import load_workbook


# from openpyxl.reader.excel import load_workbook


def excel_to_df(file_path):  # 读取excel文件到DataFrame
    # df = pd.read_excel('file/预警池企业.xlsx')
    file_path = './file/' + file_path
    data = pd.read_excel(file_path)
    output = pd.DataFrame(columns=data.columns[0:3])
    for index in range(len(data)):
        output = output.append({data.columns[0]: data.loc[index, data.columns[0]], data.columns[1]: data.loc[index, data.columns[1]], data.columns[2]: data.loc[index, data.columns[2]]}, ignore_index=True)
    # for indexs in df.index:
    #     output = output.append(df.loc[indexs].values)
    print(output[0:10])
    return output


# 读取合并的单元格
# ！！！ 不能用，报错。xlrd.biffh.XLRDError: Excel xlsx file; not supported
# def excel_merged_cell(file_path):
#     # 读取Excel文件
#     file_path = './file/' + file_path
#     book = xlrd.open_workbook(file_path)
#     sheet = book.sheet_by_index(0)
#
#     # 获取所有合并单元格的位置
#     merged_cells = sheet.merged_cells
#
#     # 读取Excel表格
#     data = []
#     for row in range(sheet.nrows):
#         row_data = []
#         for col in range(sheet.ncols):
#             value = sheet.cell_value(row, col)
#             for (rlo, rhi, clo, chi) in merged_cells:
#                 if row >= rlo and row < rhi and col >= clo and col < chi:
#                     value = sheet.cell_value(rlo, clo)
#                     break
#             row_data.append(value)
#         data.append(row_data)
#     # 将数据转换为DataFrame
#     df = pd.DataFrame(data)
#     print(df[0:10])
#     return df
def read_merge_cell(file_path):
    file_path = './file/' + file_path
    df = pd.read_excel(file_path, engine='openpyxl', sheet_name='Sheet1', header=None)
    for index, row in df.iterrows():
        for i, cell in enumerate(row):
            if pd.isna(cell):
                # 若该单元格的值为nan，则将其上面的值赋给它
                df.iloc[index, i] = df.iloc[index - 1, i]
    # 显示DataFrame
    # print(df)
    return df


def df_to_excel(df, file_path):
    df.to_excel('./out_file/new_' + file_path, index=False, header=False)


def list_to_excel(empty_comp, file_path):
    df = pd.DataFrame(empty_comp, columns=["company_name"])
    df.to_excel("./out_file/new_" + file_path, index=False)


def space_to_(str):
    str.replace(' ', '_')

if __name__ == '__main__':
    # excel_to_df('1. 舆情指标分类20220715.xlsx')
    filename = '1. 舆情指标分类20220715.xlsx'
    df = read_merge_cell('1. 舆情指标分类20220715.xlsx')
    # df_to_excel(df, filename)
    dict_ = {}
    df = df[1:]
    for items in df.values:
        i3 = items[3].replace(' ', '_')
        i4 = [items[0].replace(' ', '_'), items[1].replace(' ', '_'), items[2].replace(' ', '_')]
        d1 = {i3: i4}
        dict_.update(d1)
    # print(dict_)
    filename_2 = '仅发行私募债主体近三个月公告舆情.xlsx'
    df_yq = read_merge_cell(filename_2)
    # print(df_yq[0:20])
    # df_yq = df_yq[1:]
    df_yq['一级分类'] = ''
    df_yq['二级分类'] = ''
    df_yq['三级分类'] = ''

    list_r = []
    for index, row in df_yq.iterrows():
        if index == 0:
            df_yq.iloc[index, len(row) - 3 + 0] = '一级分类'
            df_yq.iloc[index, len(row) - 3 + 1] = '二级分类'
            df_yq.iloc[index, len(row) - 3 + 2] = '三级分类'
            continue
        if row[3] in dict_.keys():
            # list_r.append((row[3]))
            # print(len(list_r)+1, row[3])
            # if len(list_r)+1 == 495:
            #     a = 1
            index_all = dict_[row[3]]
            index_1 = index_all[0]
            index_2 = index_all[1]
            index_3 = index_all[2]
            df_yq.iloc[index, len(row)-3+0] = index_1
            df_yq.iloc[index, len(row)-3+1] = index_2
            df_yq.iloc[index, len(row)-3+2] = index_3
        else:
            continue
    # print(df_yq[0:10])
    df_to_excel(df_yq, filename_2)
