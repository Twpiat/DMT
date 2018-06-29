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

ListColumns = (67, 68, 70, 71, 73)

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
            if row_data[1]!="N/D" and row_data not in result_data:
                result_data.append(row_data)


    for lista in result_data:
        try:
            if lista[2] == "Y":
                required_data.append(lista[1])
        except:
            print("Wybuch:",zakladka)

        else:
            lista[2] = "N"

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


    output_msg = zakladka + " = [\n"

    for i, j, k, l, m in result_data:

        output_msg += "('{}', '{}', '{}', '{}', '{}'),\n".format(i, j, k, l, m)

    output_msg += "]\n\n"

    if required_data:
        output_msg += zakladka + "_REQ = [\n"
        for pole in required_data:
            output_msg += "'" + pole + "', "
        output_msg += "\n]\n"

    file = open("EX_wsad_" + zakladka + ".py", "w")
    file.write(output_msg)
    file.close()
    print("Przetwarzanie", zakladka, "zakończone.")
    return zakladka



def generuj_sql(tabela, nazwa, tabela_req):
    licznik = 0
    str_req = ", ".join(tabela_req)
    file = open("EX SQL wsad " + nazwa + ".txt", "w")
    for lista in tabela:
        licznik += 1

        str_to_write = "--" + str(lista[1]) + ": Zapytanie " + str(licznik) + " / " + str(len(tabela)) + ":\n"

        str_to_write += "SELECT MAX(LENGTH(TRIM("+str(lista[1]) + "))) AS "\
                        +str(lista[1])+"_"+str(lista[3])+"\nFROM "+DatabaseName+"."+str(lista[0])+";\n\n"

        if licznik == len(tabela):
            str_to_write += "--Zapytanie o REQ = Y:\nSELECT DISTINCT "
            str_to_write += str_req
            str_to_write += "\n\tFROM " + DatabaseName + "." + str(lista[0]) + "\nWHERE"
            for pole in tabela_req:
                if pole != tabela_req[len(tabela_req) - 1]:
                    str_to_write += "\n\t" + pole + " IS NULL OR"
                else:
                    str_to_write += "\n\t" + pole + " IS NULL;"

        file.write(str_to_write)

    file.close()
    print("Plik: EX SQL wsad " + nazwa + ".txt wygenerowany. Skopiuj zawartość pliku i uruchom go w DBVisualizer.")


def przelec_excela():
    counter = 0
    imports = []
    output_imports = ""
    output_args = ""
    name = ""
    for i in range(MinZakladka-1, MaxZakladka):
        imports.append(przelec_zakladke(i))
        name = str(imports[counter])
        output_imports += "from EX_wsad_" + name +" import *\n"
        output_args += "generuj_sql("+ name +", '" + name + "', " + name + "_REQ)\n"
        counter+=1

    print("Przetworzono", counter, "zakładek. Oto importy do dodania:\n")
    print("Zaimportuj:\n"+output_imports)
    print("Wygeneruj:\n"+output_args)


#przelec_excela()
przelec_zakladke(2)

# *************** CZYTAJ TUTAJ *************** CZYTAJ TUTAJ ***************

# ZAIMPORTUJ WSADY WYGENEROWANE W przelec_excela(),
# NAZWANE wsad_NAZWA_ZAKŁADKI W PLIKU ExcelFileName
# NP. DLA ZAKŁADKI "DMT_TAHOLD" KOMENDA TO:
# from EX_wsad_DMT_TAHOLD import *
# W RAZIE BŁĘDÓW KODOWANIA RĘCZNIE POPRAW PLIK WSADU

from EX_wsad_DMT_TXI00301 import *
from EX_wsad_DMT_TXI00501 import *
from EX_wsad_DMT_TXI00701 import *
from EX_wsad_DMT_TXI00401 import *

from EX_wsad_DMT_TXI01501 import *
from EX_wsad_DMT_TEGI06101 import *
from EX_wsad_DMT_TEGI04501 import *
from EX_wsad_DMT_TEGI04601 import *

from EX_wsad_DMT_TXI07501 import *
from EX_wsad_DMT_TEGI06001 import *
from EX_wsad_DMT_TXI01301 import *
from EX_wsad_DMT_TMH0071 import *

# DLA WSZYSTKICH WSADÓW WYGENERUJ SQL PONIŻSZĄ FUNKCJĄ.
# JEJ PARAMETRY TO: NAZWA_LISTY_Z_WSADEM, NAZWA_TABLICY_W_BD, NAZWA_LISTY_Z_POLAMI_WYMAGANYMI,
# NP. generuj_sql(DMT_TAHOLD, "XTAHOLD", DMT_TAHOLD_REQ)

# generuj_sql(DMT_TXI00301, 'DMT_TXI00301', DMT_TXI00301_REQ)
# generuj_sql(DMT_TXI00501, 'DMT_TXI00501', DMT_TXI00501_REQ)
# generuj_sql(DMT_TXI00701, 'DMT_TXI00701', DMT_TXI00701_REQ)
# generuj_sql(DMT_TXI00401, 'DMT_TXI00401', DMT_TXI00401_REQ)
#
#
# generuj_sql(DMT_TXI01501, 'DMT_TXI01501', DMT_TXI01501_REQ)
# generuj_sql(DMT_TEGI06101, 'DMT_TEGI06101', DMT_TEGI06101_REQ)
# generuj_sql(DMT_TEGI04501, 'DMT_TEGI04501', DMT_TEGI04501_REQ)
# generuj_sql(DMT_TEGI04601, 'DMT_TEGI04601', DMT_TEGI04601_REQ)
#
# generuj_sql(DMT_TXI07501, 'DMT_TXI07501', DMT_TXI07501_REQ)
# generuj_sql(DMT_TEGI06001, 'DMT_TEGI06001', DMT_TEGI06001_REQ)
# generuj_sql(DMT_TXI01301, 'DMT_TXI01301', DMT_TXI01301_REQ)
# generuj_sql(DMT_TXI01001, 'DMT_TXI01001', DMT_TXI01001_REQ)

generuj_sql(DMT_TMH0071, "XTMMAIN", DMT_XTMMAIN_REQ)
# PO WYGENEROWANIU SQL-i SKOPIUJ CAŁOŚĆ PLIKU SQL WSAD,
# NP. "EX SQL wsad XTAHOLD.txt" DO DBVisualizera I URUCHOM.
