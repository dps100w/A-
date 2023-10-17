from A类差旅人员拆分 import split_personnel
from A类成本初分 import initial_cost_distribution


def main():
    # 定义输入和输出文件的路径
    input_path_1 = '项目成本表-202309-AB类汇总-Z.xlsx'
    output_path_1 = '项目成本表-202309-AB类汇总-Z-拆分.xlsx'

    input_path_2 = output_path_1
    output_path_2 = 'processed_项目成本表-202309-AB类汇总-Z.xlsx'

    # 调用A类差旅人员拆分.py中的函数
    split_personnel(input_path_1, output_path_1)

    # 使用split_personnel的输出作为initial_cost_distribution的输入
    initial_cost_distribution(input_path_2, output_path_2)


if __name__ == "__main__":
    main()
