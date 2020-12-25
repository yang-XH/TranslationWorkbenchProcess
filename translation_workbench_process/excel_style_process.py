#!/usr/bin/env python
# coding: utf-8

import Ipynb_importer


import yaml
import xlwt
import ntpath
from utils import *
from config_handler import *
'''
在配置文件中固定excel各风格，每次写入数据时按照风格写入，而不是保持原有的风格
若保持原有的风格，需提取原有风格，仅仅formatting_info=True只能保证复制的数据的风格一致，不能保证新写入的或修改的数据的风格一致
保持原有风格：https://cloud.tencent.com/developer/ask/65204
'''

class XlwtStyleWriter(object):
# https://github.com/awslabs/predictive-maintenance-using-machine-learning/blob/master/source/predictive_maintenance/pandas/io/excel.py

    def __init__(self):
        super(XlwtStyleWriter, self).__init__()

        '''
        # 首行冻结
        if _validate_freeze_panes(freeze_panes):
            wks.set_panes_frozen(True)
            wks.set_horz_split_pos(freeze_panes[1])
            #wks.set_vert_split_pos(freeze_panes[1])
        '''

    @property
    def headings(self):
        return self._headings
    
    @headings.setter
    def headings(self,value):
        if not isinstance(value, list):
            raise ValueError('headings must be a list!')
        self._headings = value
        
    @property
    def height(self):
        return self._height
    
    @height.setter
    def height(self,value):
        if value < 0:
            raise ValueError('height must greater than 0!')
            logging.error('height must greater than 0!')
        self._height = value
        
    @property
    def hheight(self):
        return self._hheight
    
    @hheight.setter
    def hheight(self,value):
        if value < 0:
            raise ValueError('head height must greater than 0!')
            logging.error('head height must greater than 0!')
        self._hheight = value
        
    @property
    def width(self):
        return self._width
    
    @width.setter
    def width(self,value):
        
        if not isinstance(value, list):
            raise ValueError('width must be a list!')
            logging.error('width must be a list!')
        
        self._width = value
        
    @classmethod
    def _style_to_xlwt(cls, item, firstlevel=True, field_sep=',',line_sep=';'):
        """helper which recursively generate an xlwt easy style string
        for example:
            hstyle = {"font": {"bold": True},
            "border": {"top": "thin",
                    "right": "thin",
                    "bottom": "thin",
                    "left": "thin"},
            "align": {"horiz": "center"}}
            will be converted to
            font: bold on; \
                    border: top thin, right thin, bottom thin, left thin; \
                    align: horiz center;
        """
        if hasattr(item, 'items'):
            if firstlevel:
                it = ["{key}: {val}"
                      .format(key=key, val=cls._style_to_xlwt(value, False))
                      for key, value in item.items()]
                out = "{sep} ".format(sep=(line_sep).join(it))
                return out
            else:
                it = ["{key} {val}"
                      .format(key=key, val=cls._style_to_xlwt(value, False))
                      for key, value in item.items()]
                out = "{sep} ".format(sep=(field_sep).join(it))
                return out
        else:
            item = "{item}".format(item=item)
            item = item.replace("True", "on")
            item = item.replace("False", "off")
            return item

    @classmethod
    def _convert_to_style(cls, style_dict):
        """
        converts a style_dict to an xlwt style object
        Parameters
        ----------
        style_dict : style dictionary to convert
        num_format_str : optional number format string
        """
        import xlwt

        if style_dict:
            xlwt_stylestr = cls._style_to_xlwt(style_dict)
            style = xlwt.easyxf(xlwt_stylestr, field_sep=',', line_sep=';')
            print(xlwt_stylestr)
        else:
            style = xlwt.XFStyle()

        return style
    
    def write_xls(self, file_name, headings, data, heading_xf, data_xfs, set_panes_frozen=True):
        # 写入名为filename的新文件
        book = xlwt.Workbook()
        sheet = book.add_sheet('Sheet1')
        rowx = 0
        for colx, value in enumerate(headings):
            sheet.write(rowx, colx, value, heading_xf)
        if set_panes_frozen:
            sheet.set_panes_frozen(True) # frozen headings instead of split panes
            sheet.set_horz_split_pos(1) # in general, freeze after last heading row
            #sheet.set_remove_splits(True) # if user does unfreeze, don't leave a split there
        for i, width_i in enumerate(self._width):
            sheet.col(i).width = width_i
        sheet.row(0).height_mismatch = True # tell xlwt row height and default font height do not match
        sheet.row(0).height = self._hheight
        for row in data:
            rowx += 1
            sheet.row(rowx).height_mismatch = True # tell xlwt row height and default font height do not match
            sheet.row(rowx).height = self._height
            for colx, value in enumerate(row):
                sheet.write(rowx, colx, value, data_xfs)
        book.save(file_name)
        
    # 向已存在excel中写入多行数据
    def write_xls_append(self, path, value, style=xlwt.XFStyle(),set_panes_frozen=False):
        path_xls = xlsx2xls(path)
        rows_append_num = len(value)  # 获取需要写入数据的行数
        cols = len(value[0]) # 列数
        workbook = open_excel(path)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_origin_num = worksheet.nrows  # 获取表格中已存在的数据的行数
        # xlrd的open函数的formatting_info=True只能保证未更改的数据保持原有的excel格式，更改的数据不可，需要用xlutils.XLWTWriter返回原excel的style信息
        new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
        new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
        logging.info('#################################################################################################################')
        if cols != len(self._width):
            logging.info('%s 文件的列数与配置文件中设定好列宽的列数不等', os.path.abspath(path))
            logging.error('%s 文件的列数与配置文件中设定好列宽的列数不等', os.path.abspath(path))
            print('{}文件的列数与配置文件中设定好列宽的列数不等'.format(os.path.abspath(path)))
        for i, width_i in enumerate(self._width):
            new_worksheet.col(i).width = width_i
        new_worksheet.row(0).height_mismatch = True # tell xlwt row height and default font height do not match
        new_worksheet.row(0).height = self._hheight
        for i in tqdm(range(0, rows_append_num)):
            logging.info(value[i])
            new_worksheet.row(i+rows_origin_num).height_mismatch = True
            new_worksheet.row(i+rows_origin_num).height = self._height
            for j in range(0, len(value[i])):
                new_worksheet.write(i+rows_origin_num, j, value[i][j], style)  # 追加写入数据，注意是从i+rows_old行开始写入，且是style格式
        if set_panes_frozen:
            new_worksheet.set_panes_frozen(True) # frozen headings instead of split panes
            new_worksheet.set_horz_split_pos(1) # in general, freeze after last heading row
        new_workbook.save(path_xls)  # 保存工作簿
        print(os.path.abspath(path_xls) + " 写入数据成功！\n")
        print('###########################################################################################################')
        logging.info('%s 写入数据成功', os.path.abspath(path_xls))
        logging.info('#################################################################################################################')
        return
    
    def existing_xls_style_process(self, path, headings, head_style, data_style, set_panes_frozen=True):
        data = open_excel(path) #打开excel文件
        table = data.sheets()[0] #根据sheet序号来获取excel中的sheet
        nrows = table.nrows #行数
        ncols = table.ncols #列数 
        data_list =[] #待确认数据的序列
        #head = table.row_values(0)
        for rownum in range(1,nrows): #遍历每一行的内容
            row = table.row_values(rownum) #根据行号获取行
            if row: #如果行存在
                data_list.append(row)
        self.write_xls(path, headings, data_list, head_style, data_style, set_panes_frozen)
        return
    
    def YS_final_existing_xls_style_process(self, YS_final_path, YS_final_files,headings, head_style, data_style, set_panes_frozen=True):
        for final_file in YS_final_files:
            if os.path.exists(os.path.join(YS_final_path, final_file)):
                for file in listdir_nohidden(os.path.join(YS_final_path, final_file)):
                    if file.endswith(".xls"):
                        temp_path = os.path.join(YS_final_path, final_file, file)
                        self.existing_xls_style_process(temp_path, headings, head_style, data_style, set_panes_frozen)


'''
def generate_style(name, height, wrap, pattern_fore_colour, bold=False):
    style = xlwt.XFStyle()
    ########## 这部分设置字体 #########
    font = xlwt.Font()
    font.name = name 
    font.bold = bold
    font.height = height
    style.font = font
    ########## 这部分设置居中格式 ############
    alignment = xlwt.Alignment()
    #alignment.horz = xlwt.Alignment.HORZ_CENTER    #水平居中
    #alignment.vert = xlwt.Alignment.VERT_CENTER    #垂直居中
    alignment.wrap = wrap # 自动换行
    style.alignment = alignment
    ########## 设置单元格背景颜色 ############
    pattern = xlwt.Pattern()                 # 创建一个模式                                        
    pattern.pattern = pattern.SOLID_PATTERN     # 设置其模式为实型              
    pattern.pattern_fore_colour = pattern_fore_colour                        
    style.pattern = pattern  
    ######### 还可以添加几个设置颜色，边框的部分 ##########

    return style
'''

if __name__=='__main__':
    
    yaml_path = 'config.yaml'
    read_config = YamlHandler(yaml_path).read_yaml()
    excel_config = read_config['YS_final_excel']
    excel_style_config = read_config['excel_style']
    
    YS_final_files = excel_config['YS_final_files']
    YS_final_path = excel_config['YS_final_path']
    
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
    a.YS_final_existing_xls_style_process( YS_final_path, YS_final_files,a.headings, head_style, data_style, set_panes_frozen=True)


