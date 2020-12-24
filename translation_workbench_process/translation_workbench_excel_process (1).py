#!/usr/bin/env python
# coding: utf-8

# In[376]:


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
import Ipynb_importer

import ntpath


# In[343]:


from config_handler import YamlHandler  # config_handler 是ipynb文件，需先导入Ipynb_importer解析


# In[316]:


#打开excel文件
def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        print(str(e))
        logging.error('Error', exc_info=True)
    return


# In[317]:


# 创建指定路径excel文件
def create_excel(file_path):
    workbook = xlwt.Workbook(encoding='utf-8')       #新建工作簿
    sheet1 = workbook.add_sheet("sheet1")            #新建sheet
    workbook.save(file_path)                         #保存
    return


# In[318]:


# 向已存在excel中写入多行数据
'''
fpath='mcw_test.xlsx'
valueli=[["3","明明如月","女","听歌","2030.07.01"],
         ["4","志刚志强","男","学习","2019.07.01"],]
'''
def write_excel_append(path, value):
    path_xls = ntpath.splitext(path)[0]+'.xls'
    rows_append_num = len(value)  # 获取需要写入数据的行数
    workbook = open_excel(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_origin_num = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    logging.info('#################################################################################################################')
    for i in tqdm(range(0, rows_append_num)):
        logging.info(value[i])
        for j in range(0, len(value[i])):
            new_worksheet.write(i+rows_origin_num, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path_xls)  # 保存工作簿
    print(os.path.abspath(path_xls) + " 写入数据成功！\n")
    print('###########################################################################################################')
    logging.info('%s 写入数据成功', os.path.abspath(path_xls))
    logging.info('#################################################################################################################')
    return


# In[359]:


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


# In[361]:


def excel_backup(file, backup_path, backup_file):
    workbook = open_excel(file)  # 打开工作簿
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    logging.info('#################################################################################################################')
    # TODO 创建文件夹，文件名应和原文件一样，而不是backup_file
    mkdir(os.path.join(backup_path, backup_file))
    # os.path.basename 得到路径中的除去目录的基本文件名称
    # 保存为xls，防止WPS编辑过导致模板不对使得xlsx打不开的问题
    basename = ntpath.basename(file)
    basename_no_postfix = ntpath.splitext(basename)[0]
    new_workbook.save(os.path.join(backup_path, backup_file, basename_no_postfix+'.xls')) # 保存工作簿
    logging.info('%s 备份数据成功', os.path.abspath(file))
    print(os.path.abspath(file) + " 备份数据成功！\n")
    print('###########################################################################################################')


# In[321]:


# 备份数据，并将新ecxel数据分类并入已存在的ecxel文件中 by_index：表的索引
def excel_append_process(file= '../translation_workbench_data/未翻译内容1201+to+trans.xlsx',file_temp = '../translation_workbench_data/待确认的数据.xls',backup_path = r'\\172.20.56.15\d\YS-Final\backup',backup_file = 'backup_file', by_index=0):
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
        data_list.insert(0,table.row_values(0))
        # 创建file_temp文件
        create_excel(file_temp)
        write_excel_append(file_temp, data_list)
        logging.info('#################################################################################################################')
        print('有数据待处理，文件位于:{}'.format(os.path.abspath(file_temp)))
        logging.info('文件: %s 中有数据待处理', os.path.abspath(file_temp))
    for file in target_file_data_dict:
        # 备份需要更改的文件
        excel_backup(file, backup_path, backup_file)
        write_excel_append(file, target_file_data_dict[file])     
    
    return 


# In[322]:


def getColumnIndex(table, columnName):
    columnIndex = None  
    for i in range(table.ncols):        
        if(table.cell_value(0, i) == columnName):
            columnIndex = i
            break
    return columnIndex


# In[323]:


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


# In[338]:


def get_file_temp_name(file_path, pending_data_path):
    '''
    # 待处理数据的文件路径不应与正式数据的路径相同
    path_list = file_path.split('/')[:-1]
    path = '/'.join(path_list)
    '''
    basename = ntpath.basename(file)
    basename_no_postfix = ntpath.splitext(basename)[0]
    return os.path.join(pending_data_path, '待处理数据__in__'+ basename_no_postfix + '.xls')


# In[353]:


def dict2txt(path,dict_temp):
    # 先创建并打开一个文本文件
    file = open(path, 'w', encoding='utf-8') # 指定编码格式，否则读取时中文乱码

    # 遍历字典的元素，将每项元素的key和value分拆组成字符串，注意添加分隔符和换行符
    file.write("""# 第一次生成field2file字典时，使用get_YS_final_dict方法生成\n# 之后使用txt2dict方法直接读取txt得到field2file字典\n# 若对字典有增改，直接在txt末尾按格式加入新的key:value即可\n""")
    for k,v in dict_temp.items():
        # 文件名中有空格，因此txt中用中文的冒号：，将field与file名隔开
        file.write(str(k)+'：'+str(v)+'\n')

    # 注意关闭文件
    file.close()


# In[363]:


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


# In[327]:


# 忽略隐藏文件的listdir，如果excel文件正在被打开，则自动保存~$开头的临时文件
def listdir_nohidden(path):
    for f in os.listdir(path):
        if not f.startswith('~$'):
            yield f


# In[328]:


def get_YS_final_dict(YS_final_path, YS_final_files, YS_dict_txt_path, field_app_to_be_confirmed_txt_path, index0, index1):
    # YS_final_files 是需要处理的各领域的文件夹名称（不是所有文件夹都需要处理）
    field_to_file_dict = {}
    field_app_to_be_confirmed_file = open(field_app_to_be_confirmed_txt_path, 'w') # 不确定目标文件的条目

    for final_file in YS_final_files:
        if os.path.exists(os.path.join(YS_final_path, final_file)):
            for file in listdir_nohidden(os.path.join(YS_final_path, final_file)):
                # TODO
                if not file.endswith(".xlsx") and not file.endswith(".xls"):
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


# In[364]:


if __name__ == '__main__':
    yaml_path = 'config.yaml'
    if os.path.exists(yaml_path):
        with open(yaml_path, 'r', encoding='utf-8') as f:
            config = yaml.load(f)
            # 更改配置中file handlers中的log文件名信息，使之由时间信息命名
            excel_config = config['YS_final_excel']
    read_config = YamlHandler(yaml_path).read_yaml()
    excel_config = read_config['YS_final_excel']
    # TODO 控制台输入待处理文件路径
    #file: ../translation_workbench_data/未翻译内容1201+to+trans.xlsx
    '''
    控制台输入
    Python3.x 中 raw_input( ) 和 input( ) 进行了整合，去除了 raw_input( )，仅保留了 input( ) 函数，
    其接收任意任性输入，将所有输入默认为字符串处理，并返回字符串类型。
    '''
    file = input('请输入待处理excel文件的路径：')
    
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
    
    # TODO
    
    # 第一次获取file_to_field_dict时用get_YS_final_dict方法
    #file_to_field_dict = get_YS_final_dict(YS_final_path, YS_final_files, YS_dict_txt_path, field_app_to_be_confirmed_txt_path, index0, index1)
    # 之后通过读取YS_dict_txt_path获取file_to_field_dict
    field_to_file_dict = txt2dict(YS_dict_txt_path)
    #file_to_field_dict
    #file_to_field_colIndex_dict
    #field_to_file_dict = {'人力资源:绩效管理前端':'../translation_workbench_data/test1.xls'} # './translation_workbench_data/test.xlsx'
    
    
    print('本次处理的文件为: {}\n'.format(os.path.abspath(file)))
    print('若存在处理不成功的数据，数据写入: {}\n'.format(os.path.abspath(file_temp)))
    
    excel_append_process(file=file,file_temp=file_temp,backup_path = backup_path, backup_file=backup_file, by_index=0)
    


# In[ ]:





# In[ ]:




