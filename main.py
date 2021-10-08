from Executor import Executor
from ExcelParser import ExcelTileParser

def main():
    executor = Executor()
    # set 1 if first time
    method = 1
    
    # write True if new and False if not new
    is_new = True

    if method == 1:
        excelparser = ExcelTileParser('./resources/выборка_Запрос на подтверждение или изменение информации об ИТ-активе (1).xlsx')
        if is_new:
            # Для новых выгрузок
            print('new')
            excelparser.execute(
                    rule_action_sheet_name='действия политик', 
                    rule_sheet_name='политики', 
                    io_sheet_name='переменные', 
                    sc_cat_item_sheet_name='sc_cat_item',
                    script_sheet_name='скрипты'
                )
        else:
            # Для старых выгрузок
            print('old')
            excelparser.execute(
                    rule_action_sheet_name='catalog_ui_policy_action_дей.по', 
                    rule_sheet_name='catalog_ui_policy_политики', 
                    io_sheet_name='item_option_new', 
                    sc_cat_item_sheet_name='sc_cat_item_плитка',
                    script_sheet_name=''
                )
        
        lines = executor.read_source_file()
        executor.execute_seource_lines(lines)
        executor.try_set_categories_for_tile()
    if method == 2:
        executor.select_other_rows()


if __name__ == '__main__':
	main()