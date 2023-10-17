import pandas as pd
import re


def split_personnel(input_path, output_path):
    # 读取Excel文件的“A类成本”表
    df = pd.read_excel(input_path, sheet_name="A类成本")

    # 创建一个新的DataFrame来存储处理后的数据
    new_rows = []

    for _, row in df.iterrows():
        # 如果费用大类为项目差旅费
        if row['费用大类'] == '项目差旅费' and pd.notna(row['出差人']):
            # 尝试使用多种分隔符来拆分出差人
            travelers = re.split(r'[,;|/、\s]+', str(row['出差人']))
            num_travelers = len(travelers)

            for traveler in travelers:
                new_row = row.copy()
                new_row['出差人'] = traveler
                new_row['实际费用金额'] = row['费用金额'] / num_travelers
                new_rows.append(new_row)
        else:
            # 对于其他情况，实际费用金额等于费用金额
            row['实际费用金额'] = row['费用金额']
            new_rows.append(row)

    # 从新行创建DataFrame
    new_df = pd.DataFrame(new_rows)

    # 保存处理后的数据到新的Excel文件
    new_df.to_excel(output_path, index=False)


if __name__ == "__main__":
    input_path = '项目成本表-202309-AB类汇总-Z.xlsx'
    output_path = '项目成本表-202309-AB类汇总-Z-拆分.xlsx'
    split_personnel(input_path, output_path)
