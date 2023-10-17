import pandas as pd


def initial_cost_distribution(input_path='项目成本表-202309-AB类汇总-Z-拆分.xlsx',
                              output_path='processed_项目成本表-202309-AB类汇总-Z.xlsx'):
    df = pd.read_excel(input_path)
    # 创建一个新的列"初步划分"
    df["初步划分"] = ""

    # 定义集团职能部门列表
    集团职能部门 = ["杭州管理中心品牌中心", "运营中心", "集团管理中心", "业财中心", "行政中心", "青岛管理中心",
                    "上海管理中心", "重庆管理中心", "厦门管理中心", "集团", "图审中心", "技术中心", "总裁办", "效果图",
                    "合约团队"]

    # 根据您的需求进行划分
    for index, row in df.iterrows():
        大类 = row["费用大类"]
        归属部门 = row["归属部门"]
        A类费用分摊特殊标记 = row["A类费用分摊特殊标记"]
        小类 = row.get("费用小类", None)

        # 逻辑1
        if 大类 == "项目差旅费":
            if 归属部门 == "集团":
                df.at[index, "初步划分"] = "前端团队"
            else:
                df.at[index, "初步划分"] = 归属部门

        # 逻辑2
        elif 大类 == "项目制作费":
            if A类费用分摊特殊标记 == "集团活动":
                df.at[index, "初步划分"] = "公共费用"
            elif 小类 == "晒图费":
                df.at[index, "初步划分"] = "后端团队"
            else:
                df.at[index, "初步划分"] = "前端团队"

        # 逻辑3
        elif 大类 == "项目概预算":
            if A类费用分摊特殊标记 in ["估算", "成本咨询"]:
                df.at[index, "初步划分"] = "前端团队"
            elif A类费用分摊特殊标记 == "概算":
                df.at[index, "初步划分"] = "后端团队"
            elif A类费用分摊特殊标记 in ["专项咨询", "幕墙", "景观", "精装", "智能化", "平战转换"]:
                df.at[index, "初步划分"] = "专项团队"

        # 逻辑4-6
        elif 大类 in ["项目会务费", "项目快递费", "项目履约保函手续费"]:
            df.at[index, "初步划分"] = "前端团队"

        # 逻辑7
        elif 大类 == "项目品宣费用":
            if 归属部门 in 集团职能部门:
                df.at[index, "初步划分"] = "公共费用"
            else:
                df.at[index, "初步划分"] = 归属部门

        # 逻辑8
        elif 大类 == "项目图审费":
            df.at[index, "初步划分"] = "后端团队"

        # 逻辑9
        elif 大类 == "项目招待费":
            if A类费用分摊特殊标记 == "施工图图审专家费":
                df.at[index, "初步划分"] = "后端团队"
            elif A类费用分摊特殊标记 == "方案阶段专家费":
                df.at[index, "初步划分"] = "前端团队"
            elif A类费用分摊特殊标记 == "超限审查专家费":
                df.at[index, "初步划分"] = "结构团队"
            elif A类费用分摊特殊标记 == "专项发生的专家费":
                df.at[index, "初步划分"] = "专项团队"
            else:
                df.at[index, "初步划分"] = 归属部门

        # 逻辑10-11
        elif 大类 in ["项目招投标/入库/备案费用", "项目诉讼费"]:
            df.at[index, "初步划分"] = "前端团队"

        # 逻辑12
        elif 大类 == "项目其他费用":
            df.at[index, "初步划分"] = 归属部门

    # 最后，您可以将处理后的数据保存回Excel文件
    df.to_excel(output_path, index=False)

if __name__ == "__main__":
    initial_cost_distribution()
