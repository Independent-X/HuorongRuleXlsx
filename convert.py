#!/usr/bin/python
# -*- coding: utf-8 -*-

import json
import openpyxl
from typing import *

def json_to_xlsx(json_path:str, xlsx_path:str, json_encoding="utf-8") -> None:
    '''该函数有以下参数：
json_path: 从火绒导出的 json 规则文件路径
xlsx_path: 想要创建的 xlsx 文件路径和名称
json_encoding: 从火绒导出的 json 规则文件编码格式

使用该函数，可以将火绒的 json 规则文件转换为 xlsx 文件。

如果希望将 xlsx 文件重新转换为 json 规则文件，则在编辑 xlsx 文件时，要求该文件第一个 sheet 的前5列、其格式不能被改变。
'''
    with open(json_path, mode="r", encoding=json_encoding) as file:
        data = json.load(file)
    workbook = openpyxl.Workbook()
    sheet = workbook["Sheet"]
    sheet.append(["title","res_path","montype","action_type","treatment"])
    for swdir in data["data"]:
        for treatment in data["data"][swdir]:
            sheet.append([swdir,treatment["res_path"],treatment["montype"],treatment["action_type"],treatment["treatment"]])
    workbook.save(xlsx_path)

def xlsx_to_json(json_path:str, xlsx_path:str, json_encoding="utf-8") -> None:
    '''该函数有以下参数：
json_path: 想要创建的 json 文件路径和名称
xlsx_path: 用 json_to_xlsx 函数导出的 xlsx 规则文件路径。
json_encoding: json 规则文件编码格式

使用该函数，可以将 xlsx 文件转换为 json 文件。
'''
    data = list(openpyxl.load_workbook(xlsx_path).worksheets[0].values)[1:]
    init_data = {"ver":"5.0","tag":"hipsuser_auto","data":{}}
    for treatment in data:
        if treatment[0] in init_data["data"]:
            init_data["data"][treatment[0]].append({
                "res_path":treatment[1],
                "montype":treatment[2],
                "action_type":treatment[3],
                "treatment":treatment[4]
            })
        else:
            init_data["data"][treatment[0]] = [{
                "res_path":treatment[1],
                "montype":treatment[2],
                "action_type":treatment[3],
                "treatment":treatment[4]
            }]
    with open(json_path, 'w', encoding=json_encoding) as file:
        file.write(json.dumps(init_data, indent=4, ensure_ascii=False).replace(": ",":"))
