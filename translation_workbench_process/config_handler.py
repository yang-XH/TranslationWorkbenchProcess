#!/usr/bin/env python
# coding: utf-8

import yaml

class YamlHandler:
    def __init__(self,file):
        self.file = file

    def read_yaml(self,encoding='utf-8'):
        """
        读取yaml数据
        load：将yaml流转化为python字典
        """
        with open(self.file, encoding=encoding) as f:
            return yaml.load(f.read(), Loader=yaml.FullLoader)
            
    def write_yaml(self, data, encoding='utf-8'):
        """
        向yaml文件写入数据
        dump：将python对象转化为yaml流
        """
        with open(self.file, encoding=encoding, mode='w') as f:
            return yaml.dump(data, stream=f, allow_unicode=True)

if __name__ == '__main__':
    data = {
        "user":{
            "username": "vivi",
            "password": "123456"
        }
    }
    # 读取config.yaml配置文件数据
    read_data = YamlHandler('config.yaml').read_yaml()
    # 将data数据写入config1.yaml配置文件
    # write_data = YamlHandler('../translation_workbench_data/config.yaml').write_yaml(data)
    print(read_data)


