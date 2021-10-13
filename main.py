from Executor import Executor


def main():
    
    # set 1 if first time
    method = 1
    
    # write True if new and False if not new
    is_new = True

    # True если ЗнО и False если Инцидент
    is_zno = False
    name = 'выборка_Неполадки в АСУ МТР (ЗСК-ПСМК).xlsx'
    executor = Executor(excel_file_name=name,is_new=is_new,is_zno=is_zno)
    if method == 1:
        executor.execute()
    if method == 2:
        executor.select_other_rows()


if __name__ == '__main__':
	main()