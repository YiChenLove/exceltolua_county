#coding:utf-8

import sys
import os
import xlrd
import re

reload(sys)
sys.setdefaultencoding("utf-8")


# 当前脚本路径
curpath = os.path.dirname(os.path.abspath(sys.argv[0]))

# 文件头描述格式化文本
lua_file_head_format_desc = '''--[[

        %s
        exported by excel2lua.py
        from file:%s

--]]\n\n'''

# 将数据导出到tgt_lua_path
def excel2lua(src_excel_path, tgt_lua_path):
    # print('[file] %s -> %s' % (src_excel_path, tgt_lua_path))s
    # load excel data
    excel_data_src = xlrd.open_workbook(src_excel_path, encoding_override = 'utf-8')
    # print('[excel] Worksheet name(s):', excel_data_src.sheet_names())
    excel_sheet = excel_data_src.sheet_by_index(0)
    # print('[excel] parse sheet: %s (%d row, %d col)' % (excel_sheet.name, excel_sheet.nrows, excel_sheet.ncols))
    
    lua_str = "local country_data = {\n"

    prvnend = False
    cityend = False
    Countyend = False
    for row in range(2, excel_sheet.nrows):
        cell2 = excel_sheet.cell(row, 1)
        cell3 = excel_sheet.cell(row, 2)
        if cell3.value == 1 :
            prvnKey = cell2.value
            if prvnend == True :
                prvnend = False
                if cityend == True :
                    cityend = False
                    lua_str =  lua_str + "\n\t\t},\n\t},\n\t[\"" + prvnKey + "\"] = {\n"
                else :
                    lua_str =  lua_str + "\t},\n\t[\"" + prvnKey + "\"] = {\n"
            else :
                lua_str =  lua_str + "\t[\""+ prvnKey + "\"] = {\n"
        if cell3.value == 2 :
            cityKey = cell2.value
            if cityend == True :
                cityend = False
                lua_str =  lua_str + "\n\t\t},\n\t\t[\"" + cityKey + "\"] = {\n\t\t\t"
            else :
                lua_str =  lua_str + "\t\t[\"" + cityKey + "\"] = {\n\t\t\t"
        if cell3.value == 3 :
            countyKey = cell2.value
            lua_str =  lua_str +"\""+ countyKey + "\","
            prvnend = True
            cityend = True


    lua_str = lua_str + "\n\t\t},\n\t},\n\n}\nreturn country_data" 

    # 正则搜索lua文件名 不带后缀 用作table的名称 练习正则的使用
    searchObj = re.search(r'([^\\/:*?"<>|\r\n]+)\.\w+$', tgt_lua_path, re.M|re.I)
    lua_table_name = searchObj.group(1)
    # print('正则匹配:', lua_table_name, searchObj.group(), searchObj.groups())

    # 这个就直接获取文件名了
    src_excel_file_name = os.path.basename(src_excel_path)
    tgt_lua_file_name = os.path.basename(tgt_lua_path)

    # file head desc
    lua_file_head_desc = lua_file_head_format_desc % (tgt_lua_file_name, src_excel_file_name)

    # export to lua file
    lua_export_file = open(tgt_lua_path, 'w')
    lua_export_file.write(lua_file_head_desc)
    lua_export_file.write(lua_str)

    lua_export_file.close()


# Make a script both importable and executable (∩_∩)
if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('python excel2lua.py <excel_input_path> <lua_output_path>')
        exit(1)

    excel2lua(os.path.join(curpath, sys.argv[1]), os.path.join(curpath, sys.argv[2]))

    exit(0)