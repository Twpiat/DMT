import xlrd

# PODAJ NAZWĘ PLIKU DMT I PLIKU WSADOWEGO SQL
ExcelFileName = '../DMT002 133.xlsx'

# PO URUCHOMIENIU PROGRAM WYGENERUJE JEDEN WSAD
# ZE WSZYSTKIMI ZAPYTANIAMI DO WSZYSTKICH ZNALEZIONYCH TABEL EKSPORTOWYCH.
# UWAGA: OSTATNIE ZAPYTANIE SIĘ DUBLUJE, BEZ WPŁYWU NA DZIAŁANIE PROGRAMU.
output_file = "../EX SQL wsad.txt"

# PODAJ NAZWĘ BAZY DANYCH
DatabaseName = "ET_ZT8_DEF"

# PODAJ NR PIERWSZEGO ISTOTNEGO WIERSZA
# (NR LINII EXCELA)

FirstRow = 8

# PODAJ NUMERY ISTOTNYCH KOLUMN W ZAKŁADCE,
# LICZĄC OD 0

ListColumns = (67, 68, 70, 71, 73)
# ListColumns = (37, 38, 40, 41, 43) # Z niebieskiej tabeli, daje dziwne wyniki

# PODAJ NR PIERWSZEJ I OSTATNIEJ ISTOTNEJ ZAKŁADKI EXCELA,
# LICZĄC OD 1, NP. DLA DMT003 BĘDZIE TO 6 DO 9
# !! UWAGA !! PROGRAM BIERZE POD UWAGĘ UKRYTE ZAKŁADKI.
# JEŚLI WYSTĄPI BŁĄD, SPRAWDŹ, CZY W EXCELU NIE MA
# UKRYTYCH ZAKŁADEK I JEŚLI SĄ, SKASUJ JE LUB PRZENIEŚ
# NA KONIEC LISTY ZAKŁADEK I ZMODYFIKUJ ZAKRES.

MinZakladka = 6
MaxZakladka = 34

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

    for curr_row in range(FirstRow - 1, num_rows, 1):
        row_data = []

        try:
            data = worksheet.cell_value(curr_row, ListColumns[0])
        except:
            print(zakladka, curr_row, curr_col)

        if data and data != 'N/D':

            row_data.append(data)

            for cell in ListColumns[1:]:
                data = worksheet.cell_value(curr_row, cell)
                row_data.append(data)
        if row_data:
            result_data.append(row_data)

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

    # zmiana na słownik tabel z polami
    for lista in result_data:
        table_name = lista[0]
        table_field = lista[1:]
        if table_name not in table_dict:
            table_dict.update({table_name: [table_field]})
        if table_field not in table_dict[table_name]:
            table_dict[table_name].append(table_field)

    # stworzenie słownika wymaganych pól w danej tabeli
    for key in table_dict.keys():
        for i in table_dict[key]:
            if i[1] == "Y" and key not in required_data:
                required_data.update({key: [i[0]]})
            elif i[1] == "Y" and key in required_data:
                required_data[key].append(i[0])

    for key in required_data:
        required_data[key] = list(set(required_data[key]))

    for key in table_dict.keys():
        if key not in output_msg:
            output_msg += key + " = \n"
            output_msg += str(table_dict[key]) + "\n"


def generuj_sql():
    global table_dict, required_data
    file = open(output_file, "w")

    for key in table_dict.keys():
        counter = 0
        try:
            if key in required_data.keys():
                str_req = ", ".join(required_data[key])
        except KeyError:
            print("Wybuchło", required_data)
        for field in table_dict[key]:
            counter += 1
            output_msg = "--" + str(key) + "/" + str(field[0]) \
                         + " Zapytanie " + str(counter) + " / " + str(len(table_dict[key])) + ":\n"
            output_msg += "SELECT MAX(LENGTH(TRIM(" + str(field[0]) + "))) AS " + str(field[0]) + \
                          "_" + str(field[2]) + " FROM " + DatabaseName + "." + str(key) + ";\n\n"

            file.write(output_msg)

        if counter == len(table_dict[key]) and key in required_data:
            output_msg += "--Zapytanie o REQ = Y:\nSELECT DISTINCT "
            output_msg += str_req
            output_msg += "\n\tFROM " + DatabaseName + "." + str(key) + "\nWHERE"
            for field in required_data[key]:
                if field != required_data[key][len(required_data[key]) - 1]:
                    output_msg += "\n\t" + field + " IS NULL OR"
                else:
                    output_msg += "\n\t" + field + " IS NULL;\n\n"

        file.write(output_msg)
    file.close()
    print("Plik wsadowy wygenerowany.")


def read_excel():
    for i in range(MinZakladka - 1, MaxZakladka):
        print("Przetwarzam zakladke nr", i + 1)
        przelec_zakladke(i)


read_excel()
generuj_sql()
