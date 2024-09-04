# HuorongRultXlsx
将火绒高级防护中导出的“自动处理”规则转换为xlsx，或者将符合格式的xlsx列表转换为可以导入进火绒"自动处理"的json。

基于 Python 和 Openpyxl。

# 使用方法
## 安装 Openpyxl
`pip install openpyxl`
## 下载 convert.py 到你的项目文件夹
## 按照 convert.py 中的文档字符串进行操作

# 关于各项设置
action_type:
 1 -> create
 2 -> read
 4 -> edit
 8 -> delete
treatment:
 0 -> auto_allow
 3 -> auto_ban
