from openpyxl import load_workbook


class ExcelTileParser:
    def __init__(self, filename):
        self.wb = load_workbook(filename)
        self.rules_action = {}
        self.scripts_map = {}
        # print(self.wb.sheetnames)


    def read_rules_actions(self, sheet_name):
        ws = self.wb[sheet_name]
        need_not_to_do = ['#Н/Д', '#N/A']
        header = ws[1]
        tile_field_col = 0
        politic_name_col = 0
        is_required_col = 0
        is_read_only_col = 0
        is_visible_col = 0
        i = 0
        for cell in header:
            if cell.value == 'ВПР живые поля с плитки/заголовка':
                tile_field_col = i
            if cell.value == 'Политика интерфейса пользователя':
                politic_name_col = i
            if cell.value == 'Видимое':
                is_visible_col = i
            if cell.value == 'Только для чтения':
                is_read_only_col = i
            if cell.value == 'Обязательный':
                is_required_col = i
            i += 1
        for i in range(3, ws.max_row + 1):
            row = ws[i]
            if row[tile_field_col].value in need_not_to_do:
                continue
            tile_field = row[tile_field_col].value
            visible = self.get_column_right_version(row, is_visible_col, 'Видимость')
            required = self.get_column_right_version(row, is_required_col, 'Обязательнотсь')
            enabled = self.get_column_right_version(row, is_read_only_col, 'Только_для_чтения')
            info = [visible, required, enabled]
            if tile_field in self.rules_action:
                self.rules_action[tile_field][row[politic_name_col].value] = info
            else:
                self.rules_action[tile_field] = {}
                self.rules_action[tile_field][row[politic_name_col].value] = info

    def read_rules(self, sheet_name):
        ws = self.wb[sheet_name]
        rule_col = 0
        rule_name_col = 0
        header = ws[1]
        i = 0
        for cell in header:
            if cell.value == 'Условия каталога':
                rule_col = i
            if cell.value == 'Краткое описание':
                rule_name_col = i
            i += 1

        for i in range(3, ws.max_row + 1):
            row = ws[i]
            rule = row[rule_col].value
            rule_name = row[rule_name_col].value
            self.add_in_rules_actions(rule_name, rule)

    def get_column_right_version(self, row, column, property_name):
        not_touch = 'Не трогать'
        yes = 'Верно'
        no = 'Неверно'
        if row[column].value == not_touch:
            return ' '
        if row[column].value == yes:
            return property_name
        if row[column].value == no:
            return f'!{property_name}'

    def add_in_rules_actions(self, rule_name, rule):
        for field_name in self.rules_action.keys():
            if rule_name in self.rules_action[field_name]:
                self.rules_action[field_name][rule_name].append(rule)
    
    def try_fill_full_field_name(self, sheet_name):
        ws = self.wb[sheet_name]
        field_name_col = 6
        full_field_name_col = 7
        header = ws[1]
        i = 0
        for i in range(3, ws.max_row + 1):
            row = ws[i]
            field = row[field_name_col].value
            if field in self.rules_action:
                full_name = row[full_field_name_col].value
                for rule in self.rules_action[field]:
                    self.rules_action[field][rule].append(full_name)
        # self.print_rules_actions()
        

    def write_in_file(self):
        with open('out/БП.txt', 'w', encoding='utf-8') as f:
            for field_name in self.rules_action.keys():
                for rule_name in self.rules_action[field_name]:
                    info = self.rules_action[field_name][rule_name]
                    if len(info) >= 4:
                        visible = info[0]
                        required = info[1]
                        read_only = info[2]
                        rule = info[3]
                        full_name = info[4] if len(info) == 5 else 'NOT_FOUND_IN_IO'
                        f.write(f'{rule_name}\n\n')
                        if rule != '':
                            f.write(f'\t{rule}\n\n')
                        f.write(f'\t{field_name} {full_name} \t\t{visible} {required} {read_only}\n\n')
                    else:
                        f.write(f'\terror in {info}\n\n')
                f.write('================================================\n\n')


    def fill_source_file (self, sc_cat_item_sheet_name, io_sheet_name):
        ws_sc_cat = self.wb[sc_cat_item_sheet_name]
        ws_io = self.wb[io_sheet_name]
        category_group = 'u_group_of_categories'
        top_category = 'u_top_category'
        code_col = 56
        field_name_col = 6
        value_col = 9
        sc_cat_io_row_number = 3
        sc_cat_io_row = ws_sc_cat[sc_cat_io_row_number]
        service_now_id =  sc_cat_io_row[code_col].value
        category_id = 00000000-0000-0000-0000-000000000000
        top_id = 00000000-0000-0000-0000-000000000000
        for i in range(3, ws_io.max_row + 1):
            row = ws_io[i]
            field = row[field_name_col].value
            if field == category_group:
                category_id = row[value_col].value
            if field == top_category:
                top_id = row[value_col].value
        with open('resources/source.txt', 'w', encoding='utf-8') as f:
            f.write(f'IteTile; {service_now_id}\nIteGroupCategory; {category_id}\nIteTopCategories; {top_id}')


    def fill_scripts_map (self, script_sheet_name):
        ws = self.wb[script_sheet_name]
        variable_name_col = 8
        script_name_col = 15
        script_action_col = 17 
        for i in range(3, ws.max_row + 1):
            row = ws[i]
            variable_name = row[variable_name_col].value
            if len(variable_name) > 1:
                variable_name = variable_name[3:]
            script_name = row[script_name_col].value
            script_action = row[script_action_col].value
            self.add_in_script_map(variable_name, script_name, script_action)

    def add_in_script_map(self, variable_name, script_name, script_action): 
        script_info = {
                'name': script_name,
                'action': script_action
            }
        if not self.scripts_map.keys().__contains__(variable_name):
            self.scripts_map[variable_name] = {
                'name': '',
                'script_list': []
            }
        self.scripts_map[variable_name]['script_list'].append(script_info)


    def try_to_fill_field_names(self, io_sheet_name):
        io_sheet = self.wb[io_sheet_name]
        field_id_col = 5
        field_name_col = 6
        for row_number in range(4, io_sheet.max_row + 1):
            row = io_sheet[row_number]
            var_id = row[field_id_col].value
            # print(var_id)
            # print(self.scripts_map.keys())
            if self.scripts_map.keys().__contains__(var_id):
                field_name = row[field_name_col].value
                self.scripts_map[var_id]['name'] = field_name


    def write_in_actions_file(self):
        with open('out/scripts.txt', 'w', encoding='utf-8') as f:
            for variable in self.scripts_map.keys():
                var_name = self.scripts_map[variable]['name']
                scripts_of_variable = self.scripts_map[variable]['script_list']
                f.write(f'{var_name}\n{variable}\n')
                for script_info in scripts_of_variable:
                    name = script_info['name']
                    action = script_info['action']
                    f.write(f'{name}\n\n{action}\n\n-----------------------------------------------------\n')
                f.write(f'=====================================================\n')                    

    # def print_rules_actions(self):
    #     for field in self.rules_action:
    #         print(f'{field}:')
    #         for rule in self.rules_action[field]:
    #             print(f'\t{rule}:')
    #             print(f'\t\t{self.rules_action[field][rule]}')

    def execute(self, rule_action_sheet_name, rule_sheet_name, io_sheet_name, sc_cat_item_sheet_name, script_sheet_name):
        self.fill_source_file(sc_cat_item_sheet_name, io_sheet_name)
        self.read_rules_actions(rule_action_sheet_name)
        self.read_rules(rule_sheet_name)
        self.try_fill_full_field_name(io_sheet_name)
        self.write_in_file()
        self.fill_scripts_map(script_sheet_name)
        self.try_to_fill_field_names(io_sheet_name)
        self.write_in_actions_file()

# sheet.max_row
# sheet.max_column
def main():
    excelparser = ExcelTileParser('./resources/выборка_корпоративные_порталы.xlsx')
    # Для старых выгрузок
    excelparser.execute(
        rule_action_sheet_name='catalog_ui_policy_action_дей.по', 
        rule_sheet_name='catalog_ui_policy_политики', 
        io_sheet_name='item_option_new', 
        sc_cat_item_sheet_name='sc_cat_item_плитка'
        )
    # Для новых выгрузок
    #excelparser.execute(rule_action_sheet_name='действия политик', rule_sheet_name='политики', io_sheet_name='переменные', sc_cat_item_sheet_name='sc_cat_item')

if __name__ == '__main__':
    main()