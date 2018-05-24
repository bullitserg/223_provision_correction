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


def get_parameters_from_file(input_file):
    # получаем наборы данных из файла
    rb = xlrd.open_workbook(input_file)

    sheet = rb.sheet_by_index(0)
    names_row = sheet.row_values(0, start_colx=0, end_colx=9)
    keys_row = sheet.row_values(1, start_colx=0, end_colx=9)
    line_number = 2
    data_dicts = []

    while True:
        try:
            data_row = sheet.row_values(line_number, start_colx=0, end_colx=9)
        except IndexError:
            break

        if not data_row[0]:
            break

        # так как по умолчанию все строки из чисел преобразуются во float - конвертируем обратно
        str_data_row = []

        for cell in data_row:
            if type(cell) is float:
                cell = str(int(cell))
            cell = cell.strip().replace('\'', '"')
            str_data_row.append(cell)

        # собираем словарь
        data_dict = OrderedDict(zip(keys_row, str_data_row))
        data_dicts.append(data_dict)
        line_number += 1

    return names_row, data_dicts


def setting_parameters_from_data_line(names_line, data_line):
    # упорядоченный лист значений словаря
    data_line_values = list(data_line.values())

    # проверим, есть ли незаполненные данные
    empty_values = []
    for pos in range(len(data_line)):
        if not len(data_line_values[pos]):
            empty_values.append(names_line[pos])

    # если обеспечение контракта установлено, то получаем сведения о БИК
    if data_line['CONTRACT_PROVISION_RUB']:
        cn94 = mc(connection=mc.MS_94_1_CONNECT)
        with cn94.open():
            bank_info = cn94.execute_query(get_bank_info % data_line)
        if bank_info:
            data_line['CONTRACT_BANK_NAME'], data_line['CONTRACT_BANK_ADDRESS'] = bank_info[0]
        else:
            print('''Банк с БИК %(CONTRACT_BIC)s не найден''' % data_line)
            exit(0)

    # находим бд, в которой содержатся данные об указанной процедуре
    procedure_db_info = found_procedure_223_db(data_line['PROCEDURE_NUMBER_CH'])
    if not procedure_db_info:
        print('''Процедура %(PROCEDURE_NUMBER_CH)s не найдена''' % data_line)
        exit(0)

    # выводим основные параметры на консоль
    print('Тип процедуры: \'%s\' (%s)' % (procedure_db_info['name'], procedure_db_info['db']))
    print(SEPARATE_LINE)
    for pos in range(len(names_line)):
        print('''%s: '%s\'''' % (names_line[pos], data_line_values[pos]))

    # если обеспечение контракта установлено, то выводим сведения о БИК
    if data_line['CONTRACT_PROVISION_RUB']:
        print('Наименование банка: \'%(CONTRACT_BANK_NAME)s\'' % data_line)
        print('Адрес банка: \'%(CONTRACT_BANK_ADDRESS)s\'' % data_line)

    print(SEPARATE_LINE)

    # оповещаем о незаполненных полях
    if empty_values:
        print('Не заполнены поля: %s\n%s' % ('; '.join(empty_values), SEPARATE_LINE))

    # если не будет заполнено обеспечение заявки или контракта - так же указываем
    if not data_line['REQUEST_PROVISION_RUB']:
        print('Обеспечение заявки не будет установлено')
    if not data_line['CONTRACT_PROVISION_RUB']:
        print('Обеспечение контракта не будет установлено')

    # пропускаем установку для данной процедуры, если необходимо
    if input('Установить указанные параметры? Y/n: ') not in 'YН':
        print("Пропуск установки для процедуры %(PROCEDURE_NUMBER_CH)s" % data_line)
        return 0

    print('Установка указанных параметров')
    cn = mc(connection=procedure_db_info['connection'])

    cn.connect()

    data_line['PROCEDURE_ID'] = return_value_with_len_check(
        cn.execute_query(get_procedure_id % data_line), 'procedure_id')
    data_line['LOT_CUSTOMER_ID'] = return_value_with_len_check(
        cn.execute_query(get_lot_customer_id % data_line), 'lot_customer_id')

    # если обеспечение заявки установлено в файле, то устанавливаем его в БД
    if data_line['REQUEST_PROVISION_RUB']:
        # установка обеспечения заявок
        cn.execute_query(request_provision_insert % data_line)

        data_line['REQUEST_PROVISION_ID'] = cn.execute_query(get_last_insert_id)[0][0]

        # установка requestProvisionId в таблице lotCustomer
        cn.execute_query(set_request_provision_id_update % data_line)
    else:
        data_line['REQUEST_PROVISION_ID'] = 'Не устанавливался'

    # если обеспечение контракта установлено в файле, то устанавливаем его в БД
    if data_line['CONTRACT_PROVISION_RUB']:
        # установка обеспечения исполнения контракта
        cn.execute_query(provision_insert % data_line)

        data_line['CONTRACT_PROVISION_ID'] = cn.execute_query(get_last_insert_id)[0][0]

        # установка contractProvisionId в таблице lotCustomer
        cn.execute_query(contract_provision_id_update % data_line)
    else:
        data_line['CONTRACT_PROVISION_ID'] = 'Не устанавливался'

    # правка отображения на форме
    data_line['enabledRequestProvision'] = 0
    data_line['enabledContractProvision'] = 0

    if data_line['REQUEST_PROVISION_RUB']:
        data_line['enabledRequestProvision'] = 1
    if data_line['CONTRACT_PROVISION_RUB']:
        data_line['enabledContractProvision'] = 1

    cn.execute_query(lot_customer_view_update % data_line)

    print(SEPARATE_LINE)

    # выводим список полезных id
    for key in 'PROCEDURE_ID', 'LOT_CUSTOMER_ID', 'REQUEST_PROVISION_ID', 'CONTRACT_PROVISION_ID':
        print('%s: %s' % (key, data_line[key]))
    print(SEPARATE_LINE)

    print('Параметры установлены')


if __name__ == '__main__':

    input_dir = join(normpath(work_dir), input_dir)
    done_dir = join(normpath(work_dir), done_dir)

    input_files = [join(input_dir, file) for file in listdir(input_dir) if file.endswith('.xls')]

    if not input_files:
        print('Файлы для обработки отсутствуют')
        input('Нажмите ENTER для выхода')
        exit(0)

    for file in input_files:
        if input('Обработать файл %s? Y/n: ' % file) not in 'YН':
            continue
        else:
            print('Обработка файла')
            # Получаем из файла данные о наименованиях и list данных по всем аукционам
            n_line, data = get_parameters_from_file(file)
            print('Обнаружено %s аукционов для корректировки' % len(data))

            for d_line in data:
                setting_parameters_from_data_line(n_line, d_line)
                print(SEPARATE_LINE)

        file_done_location = move(file, done_dir)
        print('Файл перемещен в %s' % file_done_location)

    print('Все файлы обработаны')

input('Нажмите ENTER для выхода')
exit(0)
