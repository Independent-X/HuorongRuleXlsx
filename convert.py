#!/usr/bin/python
# -*- coding: utf-8 -*-

import json
import openpyxl
from typing import *

def json_to_xlsx(json_path:str, xlsx_path:str, json_encoding="utf-8") -> None:
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
