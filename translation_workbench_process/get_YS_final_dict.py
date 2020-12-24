import time

import xlrd
import xlwt
from xlutils.copy import copy


import os

from tqdm.auto import tqdm
import logging
import yaml
import logging.config
#from translation_workbench_excel_process import *

def get_YS_final_dict(YS_final_path, YS_final_files, YS_dict_txt_path, field_app_to_be_confirmed_txt_path, index0, index1):
    # YS_final_files 是需要处理的各领域的文件夹名称（不是所有文件夹都需要处理）
    field_to_file_dict = {}
    field_app_to_be_confirmed_file = open(field_app_to_be_confirmed_txt_path, 'w') # 不确定目标文件的条目

    for final_file in YS_final_files:
        if os.path.exists(os.path.join(YS_final_path, final_file)):
            for file in listdir_nohidden(os.path.join(YS_final_path, final_file)):
                if not file.endswith(".xlsx") and not file.endswith(".xls"):
                    # xlsx与csv文件内容相同，只需遍历其中一个
                    continue
                # win linux 路径斜杠不兼容，用os.path.join，而不是写死左斜杠还是右斜杠
                #temp_path = YS_final_path + '/'+ final_file +'/'+ file
                temp_path = os.path.join(YS_final_path, final_file, file)
                workbook = open_excel(temp_path)  # 打开工作簿
                sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
                table = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格

                # file为此文件名，需记录在file_to_field_dict中

                field_index = getColumnIndex(table,index0)
                app_index = getColumnIndex(table,index1)
                # 部分文件的表头没有这两项
                if field_index and app_index:
                    field_list = table.col_values(field_index)
                    app_list = table.col_values(app_index)
                    # 得到这一文件中的所有“领域：应用”list（已去重，且若其中一项为空，则结果为None）
                    field_app_list = list(set(list(map(lambda x:x[0]+':'+x[1] if x[0] and x[1] else None, zip(field_list, app_list)))))
                    field_app_list.remove('领域名称:应用名称')
                    # TODO field_to_file_dict，并去None，注意重复值的处理
                    for field_app in field_app_list:
                        if field_app is None:
                            continue
                        if field_app not in field_to_file_dict:
                            field_to_file_dict[field_app] = temp_path
                        else:
                            # 若有重复，则删除，保证字典中的条目都是已确定的，不确定的条目在确认其所属文件后，直接加入txt
                            del field_to_file_dict[field_app]
                            field_app_to_be_confirmed_file.write(str(field_app)+'\n')
        else:
            
            print('dir '+ os.path.join(YS_final_path, final_file) + ' not exists')
    field_app_to_be_confirmed_file.close()        
    dict2txt(YS_dict_txt_path, field_to_file_dict)      
    return field_to_file_dict

# 忽略隐藏文件的listdir，如果excel文件正在被打开，则自动保存~$开头的临时文件
def listdir_nohidden(path):
    for f in os.listdir(path):
        if not f.startswith('~$'):
            yield f

#打开excel文件
def open_excel(file):
    try:
        print(file)
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))
        logging.error('Error', exc_info=True)
    return

def getColumnIndex(table, columnName):
    columnIndex = None  
    for i in range(table.ncols):        
        if(table.cell_value(0, i) == columnName):
            columnIndex = i
            break
    return columnIndex

def dict2txt(path,dict_temp):
    # 先创建并打开一个文本文件
    file = open(path, 'w',encoding='utf-8') 

    # 遍历字典的元素，将每项元素的key和value分拆组成字符串，注意添加分隔符和换行符
    file.write("""# 第一次生成field2file字典时，使用get_YS_final_dict方法生成\n# 之后使用txt2dict方法直接读取txt得到field2file字典\n# 若对字典有增改，直接在txt末尾按格式加入新的key:value即可\n""")
    for k,v in dict_temp.items():
        # 文件名中有空格，因此txt中用中文的冒号：，将field与file名隔开
        file.write(str(k)+'：'+str(v)+'\n')

    # 注意关闭文件
    file.close()

def txt2dict(path):
    # 声明一个空字典，来保存文本文件数据
    dict_temp = {}

    # 打开文本文件
    file = open(path,'r',encoding='utf-8')

    # 遍历文本文件的每一行，strip可以移除字符串头尾指定的字符（默认为空格或换行符）或字符序列
    for line in file.readlines():
        if not line.startswith('#'):
            line = line.strip()
            line = line.split('：')
            k = line[0]
            v = line[1]
            dict_temp[k] = v

    # 关闭文件
    file.close()
    return dict_temp

if __name__=='__main__':
    index0 = '领域名称'
    index1 = '应用名称'
    YS_final_files = ['Collab-HR-SCM-Purchasing','Finance','Marketing','数字化建模']
    YS_final_path = r'\\172.20.56.15\d\YS-Final'
    YS_dict_txt_path = 'YS_dict.txt'
    field_app_to_be_confirmed_txt_path = 'field_app_to_be_confirmed.txt'
    dict1 = get_YS_final_dict(YS_final_path, YS_final_files, YS_dict_txt_path, field_app_to_be_confirmed_txt_path, index0, index1)
    dict2 = txt2dict(YS_dict_txt_path)
    print(dict2)