#!/usr/bin/env python
# coding: utf-8

import Ipynb_importer


from utils import *
import time

import xlrd
import xlwt
from xlutils.copy import copy


import os

from tqdm.auto import tqdm
import logging
import yaml
import logging.config


def get_YS_final_dict(YS_final_path, YS_final_files, YS_dict_txt_path, field_app_to_be_confirmed_txt_path, index0, index1):
    # YS_final_files 是需要处理的各领域的文件夹名称（不是所有文件夹都需要处理）
    field_to_file_dict = {}
    field_app_to_be_confirmed_file = open(field_app_to_be_confirmed_txt_path, 'w') # 不确定目标文件的条目

    for final_file in YS_final_files:
        if os.path.exists(os.path.join(YS_final_path, final_file)):
            for file in listdir_nohidden(os.path.join(YS_final_path, final_file)):
                if not file.endswith(".xls"):
                    # xlsx/xls与csv文件内容相同，只需遍历其中一个，且还有zip文件等
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
                    # 打印没有这两个字段的文件名，需要手动处理
                    logging.error('文件 %s 不存在 %s 和 %s 这两个字段', os.path.abspath(file), index0, index1)
        else:
            logging.error('路径 %s 不存在', os.path.join(YS_final_path, final_file))
            print('dir '+ os.path.join(YS_final_path, final_file) + ' not exists')
    field_app_to_be_confirmed_file.close()        
    dict2txt(YS_dict_txt_path, field_to_file_dict)      
    return field_to_file_dict

if __name__ == '__main__':
    yaml_path = 'config.yaml'
    read_config = YamlHandler(yaml_path).read_yaml()
    excel_config = read_config['YS_final_excel']
    YS_final_files = excel_config['YS_final_files']
    YS_final_path = excel_config['YS_final_path']
    YS_dict_txt_path = excel_config['YS_dict_txt_path']
    field_app_to_be_confirmed_txt_path = excel_config['field_app_to_be_confirmed_txt_path']
    index0 = excel_config['index0']
    index1 = excel_config['index1']
    
    file_to_field_dict = get_YS_final_dict(YS_final_path, YS_final_files, YS_dict_txt_path, field_app_to_be_confirmed_txt_path, index0, index1)

