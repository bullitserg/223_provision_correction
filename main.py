import xlrd
from queries import *
from collections import OrderedDict
from os.path import normpath, join
from os import listdir
from shutil import move
from ets.ets_mysql_lib import MysqlConnection as mc
from ets.ets_xml_worker import found_procedure_223_db
from config import *


SEPARATE_LINE = '----------------------------------------------------'


def setting_parameters_from_file(input_file):
    # получаем наборы данных из файла
    rb = xlrd.open_workbook(input_file)

    sheet = rb.sheet_by_index(0)
    names_row = sheet.row_values(0, start_colx=0, end_colx=9)
    keys_row = sheet.row_values(1, start_colx=0, end_colx=9)
    data_row = sheet.row_values(2, start_colx=0, end_colx=9)

    # функция проверяет количество возвращенных id и при необходимости выводит сообщение и завершает программу
    def return_value_with_len_check(iter_data, value_text):
        iter_data_len = len(iter_data)
        if iter_data_len == 1:
            return iter_data[0][0]
        elif iter_data_len == 0:
            print('%s не найден' % value_text)
        elif iter_data_len > 0:
            print('Найдено несколько значений %s: %s' % (value_text, iter_data))
        print('Выход')
        exit(1)

    # так как по умолчанию все строки из чисел преобразуются во float - конвертируем обратно
    str_data_row = []
    for cell in data_row:
        if type(cell) is float:
            cell = str(int(cell))
        cell = cell.strip().replace('\'', '"')
        str_data_row.append(cell)

    # собираем словарь
    data_dict = OrderedDict(zip(keys_row, str_data_row))

    # проверим, есть ли незаполненные данные
    empty_values = []
    for pos in range(len(names_row)):
        if not len(str_data_row[pos]):
            empty_values.append(names_row[pos])

    # если обеспечение контракта установлено, то получаем сведения о БИК
    if data_dict['CONTRACT_PROVISION_RUB']:
        cn94 = mc(connection=mc.MS_94_1_CONNECT)
        with cn94.open():
            bank_info = cn94.execute_query(get_bank_info % data_dict)
        if bank_info:
            data_dict['CONTRACT_BANK_NAME'], data_dict['CONTRACT_BANK_ADDRESS'] = bank_info[0]
        else:
            print('''Банк с БИК %(CONTRACT_BIC)s не найден''' % data_dict)
            exit(0)

    # находим бд, в которой содержатся данные об указанной процедуре
    procedure_db_info = found_procedure_223_db(data_dict['PROCEDURE_NUMBER_CH'])
    if not procedure_db_info:
        print('''Процедура %(PROCEDURE_NUMBER_CH)s не найдена''' % data_dict)
        exit(0)

    # выводим основные параметры на консоль
    print('Тип процедуры: \'%s\' (%s)' % (procedure_db_info['name'], procedure_db_info['db']))
    print(SEPARATE_LINE)
    for pos in range(len(names_row)):
        print('''%s: '%s\'''' % (names_row[pos], str_data_row[pos]))

    # если обеспечение контракта установлено, то выводим сведения о БИК
    if data_dict['CONTRACT_PROVISION_RUB']:
        print('Наименование банка: \'%(CONTRACT_BANK_NAME)s\'' % data_dict)
        print('Адрес банка: \'%(CONTRACT_BANK_ADDRESS)s\'' % data_dict)

    print(SEPARATE_LINE)

    # оповещаем о незаполненных полях
    if empty_values:
        print('Не заполнены поля: %s\n%s' % ('; '.join(empty_values), SEPARATE_LINE))

    # если не будет заполнено обеспечение заявки или контракта - так же указываем
    if not data_dict['REQUEST_PROVISION_RUB']:
        print('Обеспечение заявки не будет установлено')
    if not data_dict['CONTRACT_PROVISION_RUB']:
        print('Обеспечение контракта не будет установлено')

    if input('Установить указанные параметры? Y/n: ') not in 'YН':
        print('Выход без установки параметров')
        exit(0)

    print('Установка указанных параметров')
    cn = mc(connection=procedure_db_info['connection'])

    cn.connect()

    data_dict['PROCEDURE_ID'] = return_value_with_len_check(
        cn.execute_query(get_procedure_id % data_dict), 'procedure_id')
    data_dict['LOT_CUSTOMER_ID'] = return_value_with_len_check(
        cn.execute_query(get_lot_customer_id % data_dict), 'lot_customer_id')

    # установка гарантийного взноса
    cn.execute_query(procedure_guarantee_update % data_dict)

    # если обеспечение заявки установлено в файле, то устанавливаем его в БД
    if data_dict['REQUEST_PROVISION_RUB']:
        # установка обеспечения заявок
        cn.execute_query(request_provision_insert % data_dict)

        data_dict['REQUEST_PROVISION_ID'] = cn.execute_query(get_last_insert_id)[0][0]

        # установка requestProvisionId в таблице lotCustomer
        cn.execute_query(set_request_provision_id_update % data_dict)
    else:
        data_dict['REQUEST_PROVISION_ID'] = 'Не устанавливался'

    # если обеспечение контракта установлено в файле, то устанавливаем его в БД
    if data_dict['CONTRACT_PROVISION_RUB']:
        # установка обеспечения исполнения контракта
        cn.execute_query(provision_insert % data_dict)

        data_dict['CONTRACT_PROVISION_ID'] = cn.execute_query(get_last_insert_id)[0][0]

        # установка contractProvisionId в таблице lotCustomer
        cn.execute_query(contract_provision_id_update % data_dict)
    else:
        data_dict['CONTRACT_PROVISION_ID'] = 'Не устанавливался'

    # правка отображения на форме
    data_dict['enabledRequestProvision'] = 0
    data_dict['enabledContractProvision'] = 0

    if data_dict['REQUEST_PROVISION_RUB']:
        data_dict['enabledRequestProvision'] = 1
    if data_dict['CONTRACT_PROVISION_RUB']:
        data_dict['enabledContractProvision'] = 1

    cn.execute_query(lot_customer_view_update % data_dict)

    cn.disconnect()

    print(SEPARATE_LINE)

    for key in 'PROCEDURE_ID', 'LOT_CUSTOMER_ID', 'REQUEST_PROVISION_ID', 'CONTRACT_PROVISION_ID':
        print('%s: %s' % (key, data_dict[key]))
    print(SEPARATE_LINE)

    print('Параметры установлены')

    file_done_location = move(file, done_dir)
    print('Файл перемещен в %s' % file_done_location)

if __name__ == '__main__':

    input_dir = join(normpath(work_dir), input_dir)
    done_dir = join(normpath(work_dir), done_dir)

    input_files = [join(input_dir, file) for file in listdir(input_dir) if file.endswith('.xls')]

    if not input_files:
        print('Файлы для обработки отсутствуют')
        input('Нажмите ENTER для выхода')
        exit(0)

    for file in input_files:
        print('Обрабатываемый файл: %s' % file)
        if input('Обработать файл? Y/n: ') not in 'YН':
            continue
        else:
            print('Обработка файла')
            setting_parameters_from_file(file)
    print('Все файлы обработаны')

input('Нажмите ENTER для выхода')
exit(0)
