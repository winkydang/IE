import os
import pandas as pd

path = os.getcwd()
example_dir = os.path.abspath(os.path.join(path, "../../file/example/example.xlsx"))


def read_excel(file_name):
    df = pd.read_excel(example_dir, sheet_name='煤炭示例')
    # print(df)

    company_list = {}
    for index in df.index:
        company_list[df.iloc[index].values[0]] = df.iloc[index].values[1]
    # print(company_list)
    return company_list


