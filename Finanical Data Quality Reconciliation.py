import pandas as pd
import datetime
from pathlib import Path
from dateutil.relativedelta import relativedelta


def main():
    datas = pd.DataFrame()
    day_time = (datetime.datetime.now() + relativedelta(days=-1))

    # 1.读取数据
    for file_path in Path(folder_path).iterdir():
        if '~$' in str(file_path):
            continue
        file_name = str(file_path).split('\\')[-1]
        new_path = "{}\\{}".format(folder_path1, file_name)
        print(new_path)
        writer = pd.ExcelWriter(new_path, engine='xlsxwriter')
        workbook = writer.book
        # 设置格式
        format1 = workbook.add_format({...})
        format2 = workbook.add_format({...})
        format3 = workbook.add_format({...})
        format4 = workbook.add_format({...})
        name1, contract_no1 = '', ''

        # 获取每个sheet
        try:
            sheets = list(pd.read_excel(file_path, sheet_name=None).keys())
        except:
            print('读取异常文件:', file_path)
            continue

        # 2.获取指定数值
        for sheet in sheets:
            if '逾期信息' in sheet:
                data = pd.read_excel(file_path, sheet_name=sheet)
                data = data.fillna('Naa')
                name1 = data.iloc[[0], [1]].values[0][0]  # 客户姓名
                contract_no1 = data.iloc[[0], [3]].values[0][0]  # 合同号
                money1 = data.iloc[[0], [5]].values[0][0]  # 借款金额
                time1 = data.iloc[[1], [1]].values[0][0]  # 期限
                rate1 = data.iloc[[1], [3]].values[0][0]  # 利率
                time2 = data.iloc[[1], [5]].values[0][0]  # 起息日
                time3 = data.iloc[[2], [1]].values[0][0]  # 到期日
                time4 = data.iloc[[2], [3]].values[0][0]  # 每月还款日
                sum_fee2 = data.iloc[[4], [5]].values[0][0]  # 逾期利息和 y
                rent_num = data.iloc[[5], [1]].values[0][0]  # 应还租金合计
