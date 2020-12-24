#!/usr/bin/env python
# coding: utf-8

# In[2]:


import Ipynb_importer


# In[3]:


from utils import *
import os
import xlrd


# In[8]:


def YS_final_xlsx2xls(YS_final_path, YS_final_files):
    # YS_final_files 是需要处理的各领域的文件夹名称（不是所有文件夹都需要处理）
    for final_file in YS_final_files:
        if os.path.exists(os.path.join(YS_final_path, final_file)):
            for file in listdir_nohidden(os.path.join(YS_final_path, final_file)):
                if file.endswith(".xlsx"):
                    temp_path = os.path.join(YS_final_path, final_file, file)
                    temp_path_xls = xlsx2xls(temp_path)
                    workbook = xlrd.open_workbook(temp_path)  # 打开工作簿
                    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
                    
                    # os.path.basename 得到路径中的除去目录的基本文件名称
                    # 保存为xls，防止WPS编辑过导致模板不对使得xlsx打不开的问题

                    new_workbook.save(temp_path_xls) # 保存工作簿
        else:
            logging.error('路径 %s 不存在', os.path.join(YS_final_path, final_file))
            print('dir '+ os.path.join(YS_final_path, final_file) + ' not exists')    
    return 


# In[ ]:


def remove_xlsx(YS_final_path, YS_final_files):
    for final_file in YS_final_files:
        if os.path.exists(os.path.join(YS_final_path, final_file)):
            for file in listdir_nohidden(os.path.join(YS_final_path, final_file)):
                if file.endswith(".xlsx"):
                    file_xlsx_path = os.path.join(YS_final_path, final_file, file)
                    file_xls = xlsx2xls(file_xlsx_path)
                    # 若此文件的xls格式文件存在，则删除此xlsx
                    if os.path.exists(file_xls):
                        os.remove(file_xlsx_path)
                    else:
                        logging.info('%s 的xls格式文件不存在', file_xlsx_path)
                        print('{} 的xls格式文件不存在'.format( file_xlsx_path))


# In[ ]:


if __name__ == '__main__':
    yaml_path = 'config.yaml'
    read_config = YamlHandler(yaml_path).read_yaml()
    excel_config = read_config['YS_final_excel']
    YS_final_files = excel_config['YS_final_files']
    YS_final_path = excel_config['YS_final_path']
    print('将xlsx转换为xls格式')
    logging.info('将xlsx转换为xls格式')
    YS_final_xlsx2xls(YS_final_path, YS_final_files)
    logging.info('删除存在xls格式的xlsx文件')
    print('删除存在xls格式的xlsx文件')
    remove_xlsx(YS_final_path, YS_final_files)

