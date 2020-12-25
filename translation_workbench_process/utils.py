#!/usr/bin/env python
# coding: utf-8

import Ipynb_importer # 可直接import ipynb

import logging
import ntpath
import os
import time
import xlrd # xlrd 版本1.2.0，高版本不能同时支持xlsx
import xlwt
from xlutils.copy import copy
from config_handler import *
from tqdm.auto import tqdm
import yaml
import logging.config
import sys

def getColumnIndex(table, columnName):
    columnIndex = None  
    for i in range(table.ncols):        
        if(table.cell_value(0, i) == columnName):
            columnIndex = i
            break
    return columnIndex

# 忽略隐藏文件的listdir，因为如果excel文件正在被打开，则自动保存~$开头的临时文件
def listdir_nohidden(path):
    for f in os.listdir(path):
        if not f.startswith('~$'):
            yield f


def xlsx2xls(file):
    # win下路径的file
    file_without_postfix = ntpath.splitext(file)[0]
    file_xls = file_without_postfix + '.xls'
    return file_xls

#打开excel文件
def open_excel(file):
    try:
        data = xlrd.open_workbook(file, formatting_info=True) # 保证copy时，未更改数据保持原有excel的格式,formatting_info仅对xls文件有效
        return data
    except Exception as e:
        print(str(e))
        logging.error('Error', exc_info=True)
    return

# 创建指定路径excel文件，指定格式
def create_excel(file_path):
    path_xls = xlsx2xls(file_path)
    workbook = xlwt.Workbook(encoding='utf-8')       #新建工作簿
    sheet1 = workbook.add_sheet("sheet1")            #新建sheet
    workbook.save(path_xls)                         #保存
    return

# 创建目录
def mkdir(path):
    # 去除首位空格
    path=path.strip()
    # 去除尾部 \ 符号
    path=path.rstrip("\\")
    # 判断路径是否存在
    isExists=os.path.exists(path)
    # 判断结果
    if not isExists:
        # 如果不存在则创建目录,创建目录操作函数
        '''
        os.mkdir(path)与os.makedirs(path)的区别是,当父目录不存在的时候os.mkdir(path)不会创建，os.makedirs(path)则会创建父目录
        '''
        #此处路径最好使用utf-8解码，否则在磁盘中可能会出现乱码的情况
        #os.makedirs(path.decode('utf-8'))
        os.makedirs(path)
        print (path+' 创建成功')
        logging.info('%s 创建成功', path)
        return True
    else:
        # 如果目录存在则不创建，并提示目录已存在
        print (path+' 目录已存在')
        #logging.info('%s 目录已存在', path)
        return False

def excel_backup(file, backup_path, backup_file):
    workbook = open_excel(file)  # 打开工作簿
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    logging.info('#################################################################################################################')
    # 创建文件夹，文件名应和原文件一样，而不是backup_file
    mkdir(os.path.join(backup_path, backup_file))
    # os.path.basename 得到路径中的除去目录的基本文件名称
    # 保存为xls，防止WPS编辑过导致模板不对使得xlsx打不开的问题
    basename = ntpath.basename(file)
    path_xls = xlsx2xls(basename)
    new_workbook.save(os.path.join(backup_path, backup_file, path_xls)) # 保存工作簿
    logging.info('%s 备份数据成功', os.path.abspath(file))
    print(os.path.abspath(file) + " 备份数据成功！\n")
    print('###########################################################################################################')

def dict2txt(path,dict_temp):
    # 先创建并打开一个文本文件
    mkdir(path)
    file = open(path, 'w', encoding='utf-8') # 指定编码格式，否则读取时中文乱码

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

    # 打开文本文件,encoding='utf-8'中文
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

def setup_logging(default_path='config.yaml', default_level=logging.INFO, log_file_name = 'excel_append.log'):
    yaml_path = default_path
    if os.path.exists(yaml_path):
        read_config = YamlHandler(yaml_path).read_yaml()
        logging_config = read_config['logging']
        # 更改配置中file handlers中的log文件名信息，使之由时间信息命名
        logging_config['handlers']['file']['filename'] = log_file_name
        logging.config.dictConfig(logging_config)
    else:
        logging.basicConfig(level=default_level)

# TODO :跑不通
def config_generate(yaml_path):
    read_config = YamlHandler(yaml_path).read_yaml()
    excel_config = read_config['YS_final_excel']
    excel_style_config = read_config['excel_style']
    for key in excel_config:
        #print(excel_config)
        #print(type(key))
        #print(excel_config[key])
        #print(type(excel_config[key]))
        print(key)
        exec('{}=1'.format(key))


