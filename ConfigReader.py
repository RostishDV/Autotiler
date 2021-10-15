import json

class ConfigReader:
    def __init__(self):
        with open('resources/config.txt', 'r', encoding='utf-8') as f:
            conf = f.read()
        self.config = json.loads(conf)
        print(f"in_New = {self.config['name']}")
    
    def get_connect_config(self):
        if 'connect' in self.config:
            return self.config['connect']
        print('Не найдено заданное значение подключения в конфигурации, используется стандартное')
        return {'server': 'LAPTOP-88HGVMDK', 'db': 'dev.nornickel'}


    def get_is_zno(self):
        if 'is_zno' in self.config:
            return self.config['is_zno']
        print('Не найдено заданное значение ЗНО в конфигурации, используется значение ЗНО')
        return True


    def get_name(self):
        if 'name' in self.config:
            return self.config['name']
        print('Не найдено заданное значение подключения в конфигурации, используется пустая строка')
        return ''