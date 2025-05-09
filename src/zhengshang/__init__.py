# -*- coding: utf-8 -*-
"""
@author: Jonathan Huang
"""

import requests, os, openpyxl, time
from datetime import datetime
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("fox_weixinjielong")

@mcp.tool()
def save_shangji(r_date: str, unit: str, manager: str, customer: str, b_type: str, b_desc: str, b_date: str,
                 b_from: str, s_id) -> str:
    """
    Description:从商客商机接龙中提取关键信息并保存到excel表，不要自行发挥。
    Args:
        r_date: 接龙填报日期
        unit: 划小营服（必须转换为如下规范名称之一："珠江新城北"、"天河东"、"天河北"、"天河中"、"体育西"、"软件园"、"沙河"、"智慧城"、"珠江新城东"、"金融城"、"环五山"、"广汕"、"珠江新城西"）
        manager: 客户经理（例：张三）
        customer: 客户名称（例：广州***公司，注意客户名称中的*或x代表脱敏要保留）
        b_type: 晒单类型（例：商机单）
        b_desc: 业务类型（例：新装xxx套餐）
        b_date: 预计转化时间（例：5月7日，注意要明确的日期（月日））
        b_from: 商机来源（例：收单点介绍）
        s_id: 531平台工单编号（例：s123456，如无编号则填入“无”）
    """

    units = ["珠江新城北", "天河东", "天河北", "天河中", "体育西", "软件园", "沙河", "智慧城", "珠江新城东", "金融城",
             "环五山", "广汕", "珠江新城西"]
    if unit not in units:
        return "unit参数不合法，请重新按要求输入"

    data_path = os.getcwd()  # 放置输入数据与输出结果
    template_path = os.path.join(data_path, 'template_zhengshang.xlsx')

    current_date = datetime.now().strftime("%Y%m%d")
    output_file_name = f"商机晒单接龙汇总_{current_date}.xlsx"
    output_file_path = os.path.join(data_path, output_file_name)
    if not os.path.exists(output_file_path):
        # 文件不存在，从模板复制
        print(f"File '{output_file_path}' does not exist, copying from template '{template_path}'")
        wb = openpyxl.load_workbook(template_path)
    else:
        wb = openpyxl.load_workbook(output_file_path)
    ws = wb["登记清单"]
    max_row = ws.max_row
    s_id = str(s_id)
    data_to_add = [r_date, max_row, unit, manager, customer, b_type, b_desc, b_date, b_from, s_id]  # 单元格的数据
    ws.append(data_to_add)
    wb.save(output_file_path)
    print(f"Saved '{output_file_name}'")
    return "已完成接龙信息提取并保存到excel文件中"

def main() -> None:
    mcp.run()
