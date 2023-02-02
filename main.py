import openpyxl as op
import os
import time

import openpyxl.styles.numbers


def get_info(wb):
    sum_in_information = 0
    sum_in_registration = 0
    sum_out_information = 0
    sum_out_registration = 0
    sheet = wb['Прием обращений']
    max_row = sheet.max_row
    for i in range(2, max_row + 1):
        if 'Предоставление' in sheet.cell(row=i, column=5).value:
            sum_in_information += sheet.cell(row=i, column=3).value
        else:
            sum_in_registration += sheet.cell(row=i, column=3).value

    sheet = wb['Выданные обращения']
    max_row = sheet.max_row
    for i in range(2, max_row + 1):
        if 'Предоставление' in sheet.cell(row=i, column=6).value:
            sum_out_information += sheet.cell(row=i, column=5).value
        else:
            sum_out_registration += sheet.cell(row=i, column=5).value
    return int(sum_in_information), int(sum_in_registration), int(sum_out_information), int(sum_out_registration)

# Раздел с формированием имен выводимых файлов по месяцу
mon_list = {1: 'Январь', 2: 'Февраль', 3: 'Март', 4: 'Апрель', 5: 'Май', 6: 'Июнь', 7: 'Июль', 8: 'Август', 9: 'Сентябрь', 10: 'Октярбь', 11: 'Ноябрь', 12: 'Декабрь'}
now_time = time.localtime()
if now_time.tm_mon == 1:
    pred_mon = 12
else:
    pred_mon = now_time.tm_mon-1
mon_for_name = mon_list[pred_mon]


file_list = os.listdir()    #Формирование списка файлов для работы
file_names = ['','','','']
for el in file_list:
    if 'ленин' in el.lower() and 'минэконом' not in el.lower():
        file_names[0] = el      # 0 ПВД по ЛЕНИНА
    elif 'ильич' in el.lower() and 'минэконом' not in el.lower():
        file_names[1] = el      # 1 ПВД по ИЛЬИЧА
    elif 'свер' in el.lower():
        file_names[2] = el      # 2 Общий ОТЧЕТ по сверкам для директора
    elif 'минэконом' in el.lower() and 'пвд' not in el.lower():
        file_names[3] = el      # 3 Отчет по ТОСП из Минэкономразвития по ЛЕНИНА


filename = file_names[0]
wb = op.load_workbook(filename, data_only=True)
sum_105_in_information = 0
sum_105_in_registration = 0
sum_105_out_information = 0
sum_105_out_registration = 0
sum_105_in_information, sum_105_in_registration, sum_105_out_information, sum_105_out_registration = get_info(wb)
print('ОТЧЕТ ПВД')
print('По офису Ленина (Администрация):')
print()
print('Принято сведений из ЕГРН:', sum_105_in_information)
print('Принято регистраций:', sum_105_in_registration)
print('***************************')
print('Выдано сведений из ЕГРН:', sum_105_out_information)
print('Выдано регистраций:', sum_105_out_registration)
print()


filename = file_names[1]
wb = op.load_workbook(filename, data_only=True)
sum_172_in_information = 0
sum_172_in_registration = 0
sum_172_out_information = 0
sum_172_out_registration = 0
sum_172_in_information, sum_172_in_registration, sum_172_out_information, sum_172_out_registration = get_info(wb)
print('По офису Ильича (Станция):')
print()
print('Принято сведений из ЕГРН:', sum_172_in_information)
print('Принято регистраций:', sum_172_in_registration)
print('***************************')
print('Выдано сведений из ЕГРН:', sum_172_out_information)
print('Выдано регистраций:', sum_172_out_registration)



wb = op.Workbook()
sheet = wb.active
sheet.cell(row=1, column=1).value = 'По офису на Ленина:'
sheet.cell(row=1, column=1).font = openpyxl.styles.Font(size=14, bold=True)
sheet.cell(row=2, column=1).value = 'Принято сведений из ЕГРН:'
sheet.cell(row=3, column=1).value = 'Принято регистраций:'
sheet.cell(row=2, column=2).value = sum_105_in_information
sheet.cell(row=3, column=2).value = sum_105_in_registration
sheet.cell(row=4, column=1).value = 'Выдано сведений из ЕГРН:'
sheet.cell(row=5, column=1).value = 'Выдано регистраций:'
sheet.cell(row=4, column=2).value = sum_105_out_information
sheet.cell(row=5, column=2).value = sum_105_out_registration

sheet.cell(row=9, column=1).value = 'По офису на Ильича:'
sheet.cell(row=9, column=1).font = openpyxl.styles.Font(size=14, bold=True)
sheet.cell(row=10, column=1).value = 'Принято сведений из ЕГРН:'
sheet.cell(row=11, column=1).value = 'Принято регистраций:'
sheet.cell(row=10, column=2).value = sum_172_in_information
sheet.cell(row=11, column=2).value = sum_172_in_registration
sheet.cell(row=12, column=1).value = 'Выдано сведений из ЕГРН:'
sheet.cell(row=13, column=1).value = 'Выдано регистраций:'
sheet.cell(row=12, column=2).value = sum_172_out_information
sheet.cell(row=13, column=2).value = sum_172_out_registration
wb.save(f'Готовый отчет за {mon_for_name} для директора по ПВД.xlsx')



print()
print('*********************************************************************************')
print()
print('ОТЧЕТ ПО СВЕРКАМ')
filename = file_names[2]
wb = op.load_workbook(filename, data_only=True)
sheet = wb['Отчёт']
max_row = sheet.max_row
bd_sv = ['mfc-kashira', 'mfc-kashira-ilicha']
lenina_dic = {}
ilicha_dic = {}
for i in range(2, max_row + 1):
    if 'mfc-kashira' == sheet.cell(row=i, column=4).value:
        if sheet.cell(row=i, column=7).value not in lenina_dic.keys():
            lenina_dic[sheet.cell(row=i, column=7).value] = 1
        else:
            lenina_dic [sheet.cell(row=i, column=7).value] += 1
    elif 'mfc-kashira-ilicha' == sheet.cell(row=i, column=4).value:
        if sheet.cell(row=i, column=7).value not in ilicha_dic.keys():
            ilicha_dic[sheet.cell(row=i, column=7).value] = 1
        else:
            ilicha_dic[sheet.cell(row=i, column=7).value] += 1
print('По Ленина услуг =', len(lenina_dic), 'и', sum(lenina_dic.values()), 'сверок в общей сложности')
print('По Ленина услуг =', len(ilicha_dic), 'и', sum(ilicha_dic.values()), 'сверок в общей сложности')
print()
print('По Ленина:')
print(*sorted(lenina_dic.items()), sep='\n')
print()
print('По Ильича:')
print(*sorted(ilicha_dic.items()), sep='\n')

wb = op.Workbook()
sheet = wb.active
v_row = 1
sheet.cell(row=v_row, column=1).value = 'Сверок по офису на Ленина:'
sheet.cell(row=v_row, column=1).font = openpyxl.styles.Font(size=14, bold=True)
v_row += 1
for key, value in sorted(lenina_dic.items()):
    sheet.cell(row=v_row, column=1).value = key
    sheet.cell(row=v_row, column=2).value = value
    v_row += 1
v_row += 1
sheet.cell(row=v_row, column=1).value = 'ИТОГО СВЕРОК НА ЛЕНИНА:'
sheet.cell(row=v_row, column=1).font = openpyxl.styles.Font(size=12, bold=True)
sheet.cell(row=v_row, column=2).value = sum(lenina_dic.values())
sheet.cell(row=v_row, column=2).font = openpyxl.styles.Font(size=12, bold=True)

v_row += 3
sheet.cell(row=v_row, column=1).value = 'Сверок по офису на Ильича:'
sheet.cell(row=v_row, column=1).font = openpyxl.styles.Font(size=14, bold=True)
v_row += 1
for key, value in sorted(ilicha_dic.items()):
    sheet.cell(row=v_row, column=1).value = key
    sheet.cell(row=v_row, column=2).value = value
    v_row += 1
v_row += 1
sheet.cell(row=v_row, column=1).value = 'ИТОГО СВЕРОК НА ИЛЬИЧА:'
sheet.cell(row=v_row, column=1).font = openpyxl.styles.Font(size=12, bold=True)
sheet.cell(row=v_row, column=2).value = sum(ilicha_dic.values())
sheet.cell(row=v_row, column=2).font = openpyxl.styles.Font(size=12, bold=True)

wb.save(f'Готовый отчет за {mon_for_name} для директора по сверкам.xlsx')

# Отчет по сверкам
filename = file_names[3]
wb = op.load_workbook(filename, data_only=True)
sheet = wb['Сводные данные (ТОСП)']
max_row = sheet.max_row
temp = {}
name_urm =[['УРМ Ожерелье', 'ожерелье', 'ожерелье'], ['УРМ Базаровское (Зендиково)', 'базаровское', 'зендиково'], ['УРМ Колтовское (Тарасково)', 'колтовское', 'тарасково'], ['УРМ Топкановское (Богатищево)', 'топкановское',  'богатищево']]

# Заполнение данными из таблицы
in_urlic = 0
out_urlic = 0
for i in range(5, max_row + 1):
    name_urm_new = ''
    for j in range(len(name_urm)):
        if name_urm[j][1] in (sheet.cell(row=i, column=1).value).lower():
            name_urm_new = name_urm[j][0]
            break
    else:
        name_urm_new = sheet.cell(row=i, column=1).value + ' УЖЕ ЗАКРЫТ!!!'
    temp.setdefault(name_urm_new, {}).setdefault(sheet.cell(row=i, column=2).value.rstrip(), [0,0,0,0,0])
    temp[name_urm_new][sheet.cell(row=i, column=2).value.rstrip()][0] += int(sheet.cell(row=i, column=4).value)  # Поступивших от физиков
    temp[name_urm_new][sheet.cell(row=i, column=2).value.rstrip()][1] += int(sheet.cell(row=i, column=5).value)  # Поступивших от юриков (по идее 0)
    temp[name_urm_new][sheet.cell(row=i, column=2).value.rstrip()][2] += int(sheet.cell(row=i, column=7).value)  # Выданых физикам
    temp[name_urm_new][sheet.cell(row=i, column=2).value.rstrip()][3] += int(sheet.cell(row=i, column=8).value)  # Выданых юрикам (по идее 0)
    temp[name_urm_new][sheet.cell(row=i, column=2).value.rstrip()][4] += int(sheet.cell(row=i, column=10).value)  # Консультаций
    if int(sheet.cell(row=i, column=5).value) > 0:
        in_urlic += int(sheet.cell(row=i, column=5).value)
    if int(sheet.cell(row=i, column=8).value) > 0:
        out_urlic += int(sheet.cell(row=i, column=8).value)

# Как мне кажется, красивый вывод результата в консоль
print()
print('*********************************************************************************')
print()
print('ОТЧЕТ ПО ТОСП')
for k, v in temp.items():
    print(k, ':')
    for k_in, val_in in v.items():
        if len(k_in)<=100:
            print(f'   {k_in}{"."*(100-len(k_in))}: Прин. от Физ - {val_in[0]}\tВыд. Физ. - {val_in[2]}\tКонсультаций - {val_in[4]}')
        else:
            count_iter = len(k_in)//100
            temp_len = len(k_in) % 100
            for i in range(count_iter+1):
                temp_str = k_in[100 * i:100 * (i + 1)]
                if i != count_iter:
                    print(f'   {temp_str}')
                else:
                    print(f'   {temp_str}{"." * (100 - temp_len)}: Прин. от Физ - {val_in[0]}\tВыд. Физ. - {val_in[2]}\tКонсультаций - {val_in[4]}')
    print()
    if in_urlic == 0 and out_urlic == 0:
        print('ПРИЕМА И ВЫДАЧИ ЮРЛИЦАМ НЕ БЫЛО!!!')
    else:
        if in_urlic > 0:
            print('Был прием от ЮРЛЦ, посмотри в таблице сама!')
        if out_urlic > 0:
            print('Была выдача ЮРЛИЦАМ, посмотри в таблице сама!')

# Вывод результата в файл
wb = op.Workbook()
sheet = wb.active
#fontStyle_big = openpyxl.styles.numbers.FORMAT_TEXT(size = "10")
v_row = 2
sheet.cell(row=1, column=2).value = 'ПРИЕМ ФИЗ'
sheet.cell(row=1, column=3).value = 'ВЫДАЧА ФИЗ'
sheet.cell(row=1, column=4).value = 'КОНСУЛЬТАЦИИ'
for k, v in temp.items():
    sheet.cell(row=v_row, column=1).value = k
    sheet.cell(row=v_row, column=1).font = openpyxl.styles.Font(size=14, bold=True)
    v_row += 1
    for k_in, val_in in v.items():
        sheet.cell(row=v_row, column=1).value = k_in
        sheet.cell(row=v_row, column=2).value = val_in[0]
        sheet.cell(row=v_row, column=3).value = val_in[2]
        sheet.cell(row=v_row, column=4).value = val_in[4]
        v_row += 1
    v_row += 3

wb.save(f'Готовый отчет за {mon_for_name} по ТОСП для меня.xlsx')