from openpyxl import load_workbook


class ExcelTileParser:
    def __init__(self, filename):
        self.wb = load_workbook(filename)
        self.rules_action = {}
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
        self.print_rules_actions()
        

    def write_in_file(self):
        with open('out/БП.txt', 'w', encoding='utf-8') as f:
            for field_name in self.rules_action.keys():
                for rule_name in self.rules_action[field_name]:
                    info = self.rules_action[field_name][rule_name]
                    visible = info[0]
                    required = info[1]
                    read_only = info[2]
                    rule = info[3]
                    full_name = info[4] if len(info) == 5 else 'NOT_FOUND_IN_IO'
                    f.write(f'{rule_name}\n\n')
                    if rule != '':
                        f.write(f'\t{rule}\n\n')
                    f.write(f'\t{field_name} {full_name} \t\t{visible} {required} {read_only}\n\n')
                f.write('================================================\n\n')

    def print_rules_actions(self):
        for field in self.rules_action:
            print(f'{field}:')
            for rule in self.rules_action[field]:
                print(f'\t{rule}:')
                print(f'\t\t{self.rules_action[field][rule]}')

    def execute(self, rule_action_sheet_name, rule_sheet_name, io_sheet_name):
        self.read_rules_actions(rule_action_sheet_name)
        self.read_rules(rule_sheet_name)
        self.try_fill_full_field_name(io_sheet_name)
        self.write_in_file()

# sheet.max_row
# sheet.max_column
def main():
    excelparser = ExcelTileParser('./resources/выборка_удаленный доступ.xlsx')
    excelparser.execute(rule_action_sheet_name='действия политик', rule_sheet_name='политики', io_sheet_name='переменные')

if __name__ == '__main__':
    main()