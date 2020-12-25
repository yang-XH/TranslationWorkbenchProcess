#!/usr/bin/env python
# coding: utf-8

# 1. 第一次操作时：
#     - 执行xlsx2xls.py
#         - YS_final_xlsx2xls函数将所有xlsx文件转换为xls格式（以使用formatting_info属性）
#         - remove_xlsx函数将存在xls格式的xlsx文件删除
#     - 执行excel_style_process.py，将所有xls文件的格式替换为指定格式
# 
# 2. 后续操作时：
#     - 执行translation_workbench_excel_process.py，直接按指定格式写入数据，且原数据格式不变
#         - 第一次获取file_to_field_dict时用get_YS_final_dict方法
#             - file_to_field_dict = get_YS_final_dict(YS_final_path, YS_final_files, YS_dict_txt_path, field_app_to_be_confirmed_txt_path, index0, index1)
#             - field_app_to_be_confirmed.txt中是有冲突的，待确认的应用领域、应用名称所属文件的数据，需手动添加至file_to_field_dict
#         - 程序默认通过读取YS_dict_txt_path获取file_to_field_dict

# In[1]:


#!/usr/bin/python
# -*- coding: UTF-8 -*-
import time
import xlrd # xlrd 版本1.2.0，高版本不能同时支持xlsx
import xlwt
from xlutils.copy import copy

import os

from tqdm.auto import tqdm
import logging
import yaml
import logging.config

import sys
import Ipynb_importer # 可直接import ipynb

import ntpath
from excel_style_process import *
from utils import *
from xlsx2xls import *


# In[16]:


from config_handler import YamlHandler  # config_handler 是ipynb文件，需先导入Ipynb_importer解析


# In[17]:


'''
# 向已存在excel中写入多行数据
def write_xls_append(path, value, style=xlwt.XFStyle()):
    path_xls = xlsx2xls(file_path)
    rows_append_num = len(value)  # 获取需要写入数据的行数
    workbook = open_excel(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_origin_num = worksheet.nrows  # 获取表格中已存在的数据的行数
    # xlrd的open函数的formatting_info=True只能保证未更改的数据保持原有的excel格式，更改的数据不可，需要用xlutils.XLWTWriter返回原excel的style信息
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    logging.info('#################################################################################################################')
    for i in tqdm(range(0, rows_append_num)):
        logging.info(value[i])
        for j in range(0, len(value[i])):
            new_worksheet.write(i+rows_origin_num, j, value[i][j], style)  # 追加写入数据，注意是从i+rows_old行开始写入，且是style格式
    new_workbook.save(path_xls)  # 保存工作簿
    print(os.path.abspath(path_xls) + " 写入数据成功！\n")
    print('###########################################################################################################')
    logging.info('%s 写入数据成功', os.path.abspath(path_xls))
    logging.info('#################################################################################################################')
    return
'''


# In[18]:


# 备份数据，并将新ecxel数据分类并入已存在的ecxel文件中 by_index：表的索引
def excel_append_process(file,file_temp,backup_path ,backup_file , XlwtStyleWriter, head_style,data_style,by_index=0):
    data = open_excel(file) #打开excel文件
    print(data)
    table = data.sheets()[by_index] #根据sheet序号来获取excel中的sheet
    nrows = table.nrows #行数
    ncols = table.ncols #列数 
    data_list =[] #待确认数据的序列
    target_file_data_dict = {} # 待加入对应文件的数据dict，key=文件名，value=待加入的数据
    field_index = getColumnIndex(table,index0)
    app_index = getColumnIndex(table,index1)
    '''
    field_index = file_to_field_colIndex_dict['file'].split(':')[0]
    app_index = file_to_field_colIndex_dict['file'].split(':')[1]
    '''
    for rownum in range(1,nrows): #遍历每一行的内容

        row = table.row_values(rownum) #根据行号获取行
        if row: #如果行存在
            field = table.cell_value(rownum, field_index)
            app = table.cell_value(rownum, app_index)
            field_app_index = field + ':' + app
            if field_app_index in field_to_file_dict:
                target_file = field_to_file_dict[field_app_index]
                # target_file_name = target_file.split('.')[0] 不可以用这种简单的方式划分基本文件名称和其目录，因为其目录中可能包含172.等路径
                # win下的路径操作，用ntpath，os.path在linux系统下运行时，相当于posixpath
                # target_file_name = ntpath.basename(target_file)
                '''
                # 文件名中存在“-”、空格等，作为list名称进行赋值时，相当于运算符在等号左方，报错SyntaxError: can't assign to operator，将其去掉
                #target_file_name = '_'.join(target_file_name.split('-'))
                #target_file_name = '_'.join(target_file_name.split(' '))
                # 文件名中有中文，无法直接作为list名称，改为将其存储为{字符串：[]}形式
                # TODO 去重
                if target_file_name+'_list' in globals():
                    eval(target_file_name+'_list').append(row)
                else:
                    #动态变量名赋值
                    exec('{}_list = []'.format(target_file_name))
                    eval(target_file_name+'_list').append(row)
                '''
                if target_file in target_file_data_dict:
                    target_file_data_dict[target_file].append(row)
                else:
                    target_file_data_dict[target_file] = [row]
            # 领域名称、应用名称为空，或者是新的待确认的规则，放入新的excel表格中
            else:
                 data_list.append(row)
    if data_list:
        # 若存在待处理数据，写入表头
        #data_list.insert(0,table.row_values(0))
        #head = table.row_values(0)
        # 创建file_temp文件
        create_excel(file_temp)
        XlwtStyleWriter.write_xls_append(file_temp, [XlwtStyleWriter.headings], head_style, set_panes_frozen=True)
        XlwtStyleWriter.write_xls_append(file_temp, data_list, data_style)
        logging.info('#################################################################################################################')
        print('有数据待处理，文件位于:{}'.format(os.path.abspath(file_temp)))
        logging.info('文件: %s 中有数据待处理', os.path.abspath(file_temp))
    for file in target_file_data_dict:
        # 备份需要更改的文件
        excel_backup(file, backup_path, backup_file)
        XlwtStyleWriter.write_xls_append(file, target_file_data_dict[file], data_style)     
    
    return 


# In[19]:


def get_file_temp_name(file_path, pending_data_path):
    '''
    # 待处理数据的文件路径不应与正式数据的路径相同
    path_list = file_path.split('/')[:-1]
    path = '/'.join(path_list)
    '''
    basename = ntpath.basename(file)
    basename_xls = xlsx2xls(basename)
    return os.path.join(pending_data_path, '待处理数据__in__'+ basename_xls)


# In[20]:


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


# In[21]:


if __name__ == '__main__':
    yaml_path = 'config.yaml'
    read_config = YamlHandler(yaml_path).read_yaml()
    excel_config = read_config['YS_final_excel']
    excel_style_config = read_config['excel_style']
    # TODO 控制台输入待处理文件路径
    #file: ../translation_workbench_data/未翻译内容1201+to+trans.xlsx
    '''
    控制台输入
    Python3.x 中 raw_input( ) 和 input( ) 进行了整合，去除了 raw_input( )，仅保留了 input( ) 函数，
    其接收任意任性输入，将所有输入默认为字符串处理，并返回字符串类型。
    '''
    file = input('请输入待处理excel文件的路径：')
     # 读取yaml配置文件，变量名称与配置文件中变量名称相同
    #config_generate()
    index0 = excel_config['index0']
    index1 = excel_config['index1']

    log_path = excel_config['log_path'] # log文件的路径（与需增添数据的文件路径不同）   
    pending_data_path = excel_config['pending_data_path'] # 待处理数据的文件路径
    backup_path = excel_config['backup_path']
    YS_final_files = excel_config['YS_final_files']
    YS_final_path = excel_config['YS_final_path']
    YS_dict_txt_path = excel_config['YS_dict_txt_path']
    field_app_to_be_confirmed_txt_path = excel_config['field_app_to_be_confirmed_txt_path']

    current_time = time.strftime('%Y%m%d%H%M', time.localtime(time.time()))
    log_name = os.path.join(log_path, "excel_append_"+current_time+".log" )  #日志文件路径
    setup_logging(yaml_path, log_file_name=log_name)
    file_temp = get_file_temp_name(file, pending_data_path)
    backup_file = current_time + 'backup'
    
    field_to_file_dict = txt2dict(YS_dict_txt_path)
    print(field_to_file_dict)
    
    head_height = eval(excel_style_config['head_height']) # yaml文件中读出的 20*26 是 str类型，需要转换
    height = eval(excel_style_config['height'])
    width = [eval(x) for x in excel_style_config['width']]
    head_style_dict = excel_style_config['head_style']
    data_style_dict = excel_style_config['other_style']
    headings = excel_config['headings']

    a = XlwtStyleWriter()
    a.height = height
    a.hheight = head_height
    a.headings = headings
    a.width = width
    head_style =  XlwtStyleWriter._convert_to_style(head_style_dict)
    data_style = XlwtStyleWriter._convert_to_style(data_style_dict)
 
    #field_to_file_dict = {'人力资源:绩效管理前端':'../translation_workbench_data/test1.xls'} # './translation_workbench_data/test.xlsx'
    print('本次处理的文件为: {}\n'.format(os.path.abspath(file)))
    print('若存在处理不成功的数据，数据写入: {}\n'.format(os.path.abspath(file_temp)))
    
    excel_append_process(file,file_temp,backup_path , backup_file, a, head_style, data_style, by_index=0)
    


# In[ ]:




