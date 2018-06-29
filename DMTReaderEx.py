import xlrd

# PODAJ NAZWĘ PLIKU DMT
ExcelFileName = 'DMT002 128.xlsx'

# PODAJ NAZWĘ BAZY DANYCH
DatabaseName = "ET_ZT7_DEF"

# PODAJ NR PIERWSZEGO ISTOTNEGO WIERSZA
# (NR LINII EXCELA)

FirstRow = 8

# PODAJ NUMERY ISTOTNYCH KOLUMN W ZAKŁADCE,
# LICZĄC OD 0

ListColumns = (67, 68, 70, 71, 73)

# PODAJ NR PIERWSZEJ I OSTATNIEJ ISTOTNEJ ZAKŁADKI EXCELA,
# LICZĄC OD 1, NP. DLA DMT003 BĘDZIE TO 6 DO 9
# !! UWAGA !! PROGRAM BIERZE POD UWAGĘ UKRYTE ZAKŁADKI.
# JEŚLI WYSTĄPI BŁĄD, SPRAWDŹ, CZY W EXCELU NIE MA
# UKRYTYCH ZAKŁADEK I JEŚLI SĄ, SKASUJ JE.

MinZakladka = 6
MaxZakladka = 20

workbook = xlrd.open_workbook(ExcelFileName)
table_list = []
table_dict = {}
required_data = {}
result_dict = {}
output_msg = ""

def przelec_zakladke(ktora_zakladka):
    worksheet = workbook.sheet_by_index(ktora_zakladka)
    num_rows = worksheet.nrows
    zakladka = worksheet.name

    result_data = []
    global table_list, table_dict, required_data, result_dict, output_msg

    for curr_row in range(FirstRow-1, num_rows, 1):
        row_data = []

        try:
            data = worksheet.cell_value(curr_row, ListColumns[0])
        except:
            print(zakladka, curr_row, curr_col)

        if data and data!='N/D':

            row_data.append(data)

            for cell in ListColumns[1:]:
                data = worksheet.cell_value(curr_row, cell)
                row_data.append(data)
        if row_data:
            result_data.append(row_data)

    # print(result_data)
    for lista in result_data:

        try:
            if float(lista[3]):
                lista[3] = "DL_" + str(lista[3])

        except ValueError:
            pass

        try:
            if lista[4] and float(lista[4]):
                lista[3] = "DL_" + str(lista[4])

        except ValueError:
            pass

        except IndexError:
            print("Wybuchła:", zakladka)
            pass


        for i in range(len(lista)):
            lista[i] = str(lista[i]).replace("/", "")
            lista[i] = str(lista[i]).replace("(", "")
            lista[i] = str(lista[i]).replace(")", "")
            lista[i] = str(lista[i]).replace(",", "_")
            lista[i] = str(lista[i]).replace(" ", "_")
            lista[i] = str(lista[i]).replace(".", "_")
            lista[i] = str(lista[i]).replace(":", "_")

    # print(result_data)
    # print("Wymagane", required_data)



    # zmiana na słownik tabel z polami
    for lista in result_data:
        table_name = lista[0]
        table_field = lista[1:]
        if table_name not in table_dict:
            table_dict.update({table_name:[table_field]})
        if table_field not in table_dict[table_name]:
            table_dict[table_name].append(table_field)

    print('TD:',table_dict)


    # stworzenie słownika wymaganych pól w danej tabeli
    for key in table_dict.keys():
        for i in table_dict[key]:
            if i[1]=="Y" and key not in required_data:
                required_data.update({key:[i[0]]})
            elif i[1]=="Y" and key in required_data: #and i[0] not in required_data.values():
                required_data[key].append(i[0])

    for key in required_data:
        required_data[key] = list(set(required_data[key]))







    for key in table_dict.keys():
        output_msg += key + " = \n"
        output_msg += str(table_dict[key]) + "\n"





    print("Ostatecznie:", table_dict)


przelec_zakladke(6)
przelec_zakladke(7)
print('Req:', required_data)
print('RD:',result_dict)
print(output_msg)