import xlrd

# PODAJ NAZWĘ PLIKU DMT
ExcelFileName = 'DMT051 155.xlsx'

# PODAJ NAZWĘ BAZY DANYCH
DatabaseName = "ET_ZT7_DEF"

# PODAJ NR PIERWSZEGO ISTOTNEGO WIERSZA
# (NR LINII EXCELA)

FirstRow = 8

# PODAJ NUMERY ISTOTNYCH KOLUMN W ZAKŁADCE,
# LICZĄC OD 0

ListColumns = (8, 10, 11, 12, 13)

# PODAJ NR PIERWSZEJ I OSTATNIEJ ISTOTNEJ ZAKŁADKI EXCELA,
# LICZĄC OD 1, NP. DLA DMT003 BĘDZIE TO 6 DO 9
# !! UWAGA !! PROGRAM BIERZE POD UWAGĘ UKRYTE ZAKŁADKI.
# JEŚLI WYSTĄPI BŁĄD, SPRAWDŹ, CZY W EXCELU NIE MA
# UKRYTYCH ZAKŁADEK I JEŚLI SĄ, SKASUJ JE.

MinZakladka = 6
MaxZakladka = 20



# PRZEJDŹ TERAZ NA KONIEC PLIKU!


workbook = xlrd.open_workbook(ExcelFileName)

def przelec_zakladke(ktora_zakladka):
    worksheet = workbook.sheet_by_index(ktora_zakladka)
    num_rows = worksheet.nrows
    zakladka = worksheet.name

    result_data = []
    required_data = []

    for curr_row in range(FirstRow-1, num_rows, 1):
        row_data = []

        for curr_col in ListColumns:
            try:
                data = worksheet.cell_value(curr_row, curr_col)
            except:
                print(zakladka, curr_row, curr_col)

            if curr_col == ListColumns[0] and data == "N/D" or data == "":
                break
            row_data.append(data)

        if row_data:
            result_data.append(row_data)

    for lista in result_data:
        try:
            if lista[1] == "Y":
                required_data.append(lista[0])
        except:
            print("Wybuch:",zakladka)

        else:
            lista[1] = "N"

        try:
            if float(lista[2]):
                lista[2] = "DL_" + str(lista[2])

        except ValueError:
            pass

        try:
            if float(lista[4]):
                lista[4] = "DL_" + str(lista[4])

        except ValueError:
            pass

        for i in range(len(lista)):
            lista[i] = str(lista[i]).replace("/", "")
            lista[i] = str(lista[i]).replace("(", "")
            lista[i] = str(lista[i]).replace(")", "")
            lista[i] = str(lista[i]).replace(",", "_")
            lista[i] = str(lista[i]).replace(" ", "_")
            lista[i] = str(lista[i]).replace(".", "_")
            lista[i] = str(lista[i]).replace(":", "_")


    output_msg = zakladka + " = [\n"
    for i, j, k, l, m in result_data:
        output_msg += "('{}', '{}', '{}', '{}'),\n".format(i, k, l, m)

    output_msg += "]\n\n"

    if required_data:
        output_msg += zakladka + "_REQ = [\n"
        for pole in required_data:
            output_msg += "'" + pole + "', "
        output_msg += "\n]\n"

    file = open("wsad_" + zakladka + ".py", "w")
    file.write(output_msg)
    file.close()
    print("Przetwarzanie", zakladka, "zakończone.")
    return zakladka



def generuj_sql(tabela, nazwa, tabela_req):
    licznik = 0
    str_req = ", ".join(tabela_req)
    file = open("SQL wsad " + nazwa + ".txt", "w")
    for lista in tabela:
        licznik += 1

        str_to_write = "--" + str(lista[0]) + ": Zapytanie " + str(licznik) + " / " + str(len(tabela)) + ":\n"
        if str(lista[3])[0:len(lista[1])] == str(lista[1]):
            str_precision = "_" + str(lista[3][-1])
        else:
            str_precision = ""

        str_to_write += "SELECT DISTINCT\n" + \
                        "\tCOL.COLUMN_NAME AS " + str(lista[0]) + ",\n" \
                        "\tCOL.DATA_TYPE AS " + str(lista[2]) + ",\n" \
                        "\tCOL.NUMERIC_PRECISION AS " + str(lista[3]) + ",\n" \
                        "\tCOL.NUMERIC_SCALE AS PRECYZJA" + str_precision + ",\n" \
                        "\tCOL.CHARACTER_MAXIMUM_LENGTH AS " + str(lista[1]) + "\n" \
                        "FROM INFORMATION_SCHEMA.COLUMNS COL\n" \
                        "WHERE COL.TABLE_NAME = '" + nazwa + "' \n" \
                        "AND COL.COLUMN_NAME = '" + str(lista[0]) + "';\n\n"

        if licznik == len(tabela):
            str_to_write += "--Zapytanie o REQ = Y:\nSELECT DISTINCT "
            str_to_write += str_req
            str_to_write += "\n\tFROM " + DatabaseName + "." + nazwa + "\nWHERE"
            for pole in tabela_req:
                if pole != tabela_req[len(tabela_req) - 1]:
                    str_to_write += "\n\t" + pole + " IS NULL OR"
                else:
                    str_to_write += "\n\t" + pole + " IS NULL;"

        file.write(str_to_write)

    file.close()
    print("Plik: SQL wsad " + nazwa + ".txt wygenerowany. Skopiuj zawartość pliku i uruchom go w DBVisualizer.")


def przelec_excela():
    counter = 0
    imports = []
    output_imports = ""
    output_args = ""
    for i in range(MinZakladka-1, MaxZakladka):
        imports.append(przelec_zakladke(i))
        output_imports += "from wsad_" + str(imports[counter])+" import *\n"
        output_args += "generuj_sql("+str(imports[counter])+", 'NAZWA_TABELI', "+str(imports[counter])+"_REQ)\n"
        counter+=1

    print("Przetworzono", counter, "zakładek. Oto importy do dodania:\n")
    print("Zaimportuj:\n"+output_imports)
    print("Wygeneruj:\n"+output_args)


#przelec_excela()
#przelec_zakladke(2)

# *************** CZYTAJ TUTAJ *************** CZYTAJ TUTAJ ***************

# ZAIMPORTUJ WSADY WYGENEROWANE W przelec_excela(),
# NAZWANE wsad_NAZWA_ZAKŁADKI W PLIKU ExcelFileName
# NP. DLA ZAKŁADKI "DMT_TAHOLD" KOMENDA TO:
# from wsad_DMT_TAHOLD import *
# W RAZIE BŁĘDÓW KODOWANIA RĘCZNIE POPRAW PLIK WSADU

# from wsad_DMT_TXI00301 import *
# from wsad_DMT_TXI00501 import *
# from wsad_DMT_TXI00701 import *
# from wsad_DMT_TXI00401 import *
# from wsad_DMT_TXI00601 import *
# from wsad_DMT_TXI01401 import *
# from wsad_DMT_TXI01501 import *
# from wsad_DMT_TEGI06101 import *
# from wsad_DMT_TEGI04501 import *
# from wsad_DMT_TEGI04601 import *
# from wsad_DMT_TEGI01101 import *
# from wsad_DMT_TXI07501 import *
# from wsad_DMT_TEGI06001 import *
# from wsad_DMT_TXI01301 import *
# from wsad_DMT_TXI01001 import *
from EX_wsad_DMT_TMH0071 import *


# DLA WSZYSTKICH WSADÓW WYGENERUJ SQL PONIŻSZĄ FUNKCJĄ.
# JEJ PARAMETRY TO: NAZWA_LISTY_Z_WSADEM, NAZWA_TABLICY_W_BD, NAZWA_LISTY_Z_POLAMI_WYMAGANYMI,
# NP. generuj_sql(DMT_TAHOLD, "XTAHOLD", DMT_TAHOLD_REQ)

# generuj_sql(DMT_TXI00301, 'XEGMAIN', DMT_TXI00301_REQ)
# generuj_sql(DMT_TXI00501, 'XEG00501', DMT_TXI00501_REQ)
# generuj_sql(DMT_TXI00701, 'XEG00701', DMT_TXI00701_REQ)
# generuj_sql(DMT_TXI00401, 'XEG00401', DMT_TXI00401_REQ)
# generuj_sql(DMT_TXI00601, 'XEG00601', DMT_TXI00601_REQ)
# generuj_sql(DMT_TXI01401, 'XEG01401', DMT_TXI01401_REQ)
# generuj_sql(DMT_TXI01501, 'XEG01501', DMT_TXI01501_REQ)
# generuj_sql(DMT_TEGI06101, 'XEG06101', DMT_TEGI06101_REQ)
# generuj_sql(DMT_TEGI04501, 'XEG04501', DMT_TEGI04501_REQ)
# generuj_sql(DMT_TEGI04601, 'XEG04601', DMT_TEGI04601_REQ)
# generuj_sql(DMT_TEGI01101, 'XEG01101', DMT_TEGI01101_REQ)
# generuj_sql(DMT_TXI07501, 'XEG07501', DMT_TXI07501_REQ)
# generuj_sql(DMT_TEGI06001, 'XEG06001', DMT_TEGI06001_REQ)
# generuj_sql(DMT_TXI01301, 'XEG01301', DMT_TXI01301_REQ)
# generuj_sql(DMT_TXI01001, 'XEG01001', DMT_TXI01001_REQ)
generuj_sql(DMT_TMH0071, 'XTMH00701', DMT_TMH0071_REQ)

# PO WYGENEROWANIU SQL-i SKOPIUJ CAŁOŚĆ PLIKU SQL WSAD,
# NP. "SQL wsad XTAHOLD.txt" DO DBVisualizera I URUCHOM.
