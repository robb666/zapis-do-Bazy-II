import os
import re
import pdfplumber
from datetime import datetime, timedelta
import win32com.client
from win32com.client import Dispatch
from regon_api import get_regon_data
import time


path = os.getcwd()
obj = input(r'Podaj polis(ę/y) w formacie .pdf do rejestracji: ')# + '\\'

# directory = r'C:\Users\Robert\Desktop\python\excel\zapis w Bazie\polisy\\'


def words_separately(text):
    """Tokenizuje tekst całej polisy."""
    make_tokens = re.compile(r"((?:(?<!'|\w)(?:\w-?'?)+(?<!-))|(?:(?<='|\w)(?:\w-?'?)+(?=')))")
    return make_tokens.findall(text)


def polisa(pdf):
    """Tekst całej polisy."""
    with pdfplumber.open(pdf) as policy:
        page_1 = policy.pages[0].extract_text()  # Tylko pierwsza strona
    return page_1, words_separately(page_1.lower())


def polisa_box(pdf):
    """Tekst wybranego fragmentu polisy."""
    with pdfplumber.open(pdf) as policy:
        width = policy.pages[0].width
        height = policy.pages[0].height
        box_left = (0, 0, 150, 140)
        box_center = (width - 400, 0, width - 200, 140)
        box_right = (150, 0, width, 140)
        page_1_box = policy.pages[0].crop(box_center, relative=True).extract_text()
    return words_separately(page_1_box.lower())


def pesel_checksum(p):
    """Suma kontrolna nr pesel."""
    l = int(p[10])
    suma = ((1*int(p[0])) + 3*int(p[1]) + (7*int(p[2])) + (9*int(p[3])) + (1*int(p[4])) + (3*int(p[5])) +
            (7*int(p[6])) + (9*int(p[7])) + (1*int(p[8])) + (3*int(p[9])))
    lm = (suma % 10)  # dzielenie wyniku modulo 10
    kontrola=(10 - lm)  # sprawdzenie ostatniej liczby kontrolnej
    if kontrola == 10 or l == kontrola:  # w przypadku liczby kontrolnej 10 i 0 sa jednoznaczne a 0 moze byc wynikiem odejmowania
        return 1
    else:
        # print('Niepoprawny pesel!')
        return 0


def regon_checksum(r: int):
    """Waliduje regon sprawdzając sumę kontrolną."""
    regon = list(str(r))
    suma = (int(regon[0])*8 + int(regon[1])*9 + int(regon[2])*2 + int(regon[3])*3 + int(regon[4])*4 +
            int(regon[5])*5 + int(regon[6])*6 + int(regon[7])*7) % 11
    if suma == int(regon[-1]) or suma == 10 and int(regon[-1]) == 0:
        return 1
    else:
        return 0


def regon(pesel_regon):
    """API Regon"""
    if len(pesel_regon) == 9:
        print('\nCzekam na dane z bazy REGON...')

        osoba = get_regon_data(pesel_regon)['forma']
        imie, nazwisko = '', ''
        if osoba == 'Jednoosobowa dz.g.':
            imie = get_regon_data(pesel_regon)['imie']
            nazwisko = get_regon_data(pesel_regon)['nazwisko']
        nazwa_firmy = get_regon_data(pesel_regon)['nazwa'].title()
        nip = get_regon_data(pesel_regon)['nip']
        ulica_f = get_regon_data(pesel_regon)['ul'].lstrip('ul. ')
        nr_ulicy_f = get_regon_data(pesel_regon)['nr_ul']
        nr_lok = get_regon_data(pesel_regon)['nr_lok']
        kod_poczt_f = get_regon_data(pesel_regon)['kod_poczt']
        miasto_f = get_regon_data(pesel_regon)['miasto']
        pkd = get_regon_data(pesel_regon)['pkd']
        opis_pkd = get_regon_data(pesel_regon)['opis pkd']
        data_rozp = get_regon_data(pesel_regon)['data rozpoczęcia']
        tel = get_regon_data(pesel_regon)['tel']
        email = get_regon_data(pesel_regon)['email'].lower()

        return nazwa_firmy, ulica_f, nr_ulicy_f, nr_lok, kod_poczt_f, miasto_f, tel, email
    else:
        return '', '', '', '', '', '', '', ''


"""Funkcje odpowiadają kolumnom w bazie."""
def nazwisko_imie(d):
    """Zwraca imię i nazwisko Klienta."""
    with open(path + '\\imiona.txt') as content:
        all_names = content.read().split('\n')
        if 'tuz' in d.values():
            name = [f'{d[k + 1].title()} {v.title()}' for k, v in d.items() if k > 10 and v.title() in all_names]
        elif 'warta' in d.values():
            name = [f'{d[k - 1].title()} {v.title()}' for k, v in d.items() if v.title() in all_names]
        else:
            name = [f'{d[k + 1].title()} {v.title()}' for k, v in d.items() if v.title() in all_names]
    if name:
        return name[0].split()[0], name[0].split()[1]
    else:
        return '', ''


def pesel_regon(d):
    """Zapisuje pesel/regon."""
    pesel = [pesel for k, pesel in d.items() if k < 200 and len(pesel) == 11 and re.search('\d{11}', pesel) and
                                                                                                  pesel_checksum(pesel)]
    regon = [regon for k, regon in d.items() if k < 150 and len(regon) == 9 and re.search('\d{9}', regon) and
                                                                                                  regon_checksum(regon)]
    if pesel:
        return 'p' + pesel[0]
    elif regon:
        return 'r' + regon[0]
    else:
        return ''


def prawo_jazdy(): pass


def adres():
    """Tylko w przypadku regon (API)."""
    pass


def kod_pocztowy(page_1):
    data = page_1.split()
    # print(data)
    dystans = [data[data.index(adres) - 10: data.index(adres) + 17] for adres in data
               if re.search('adres?\w+', adres, re.I) or re.search('kontakt?\w+', adres, re.I)
               or adres.lower() == 'pocztowy'][0]
    # print(dystans)
    kod_pocztowy = [kod for kod in dystans if re.search('\d{2}[-|\xad]\d{3}', kod)][0]
    return kod_pocztowy


def data_wystawienia():
    one_day = timedelta(1)
    today = datetime.strptime(datetime.now().strftime('%y-%m-%d'), '%y-%m-%d')# + one_day
    return today


def TU():
    """W funkcji numer_polisy(page_1)"""
    pass


def numer_polisy(page_1):
    nr_polisy = ''
    if 'Allianz' in page_1 and (nr_polisy := re.search('Polisa nr (\d+)', page_1)):
        return 'ALL', nr_polisy.group(1)
    if 'AXA' in page_1 and (nr_polisy := re.search('Numer polisy (\d{4}-\d+)', page_1)):
        return 'AXA', nr_polisy.group(1)
    if 'Compensa' in page_1 and (nr_polisy := re.search('typ polisy: *\s*(\d+),numer: *\s*(\d+)', page_1)):
        return 'COM', nr_polisy.group(1) + nr_polisy.group(2)
    if 'Generali' in page_1 and (nr_polisy := re.search('POLISA NR\s*(\d+)', page_1, re.I)):
        return 'GEN', nr_polisy.group(1)
    if 'Hestia' in page_1 and (nr_polisy := re.search('Polisa\s.*\s(\d+)', page_1, re.I)):
        return 'HES', nr_polisy.group(1)
    if 'LINK4' in page_1 and (nr_polisy := re.search('Numer\s(\w\d+)', page_1, re.I)):
        return 'LIN', nr_polisy.group(1)
    if 'PZU' in page_1 and (nr_polisy := re.search('Nr *(\d+)', page_1)):
        return 'PZU', nr_polisy.group(1)
    if 'TUW' in page_1 and (nr_polisy := re.search('Wniosko-Polisa\snr\s*(\d+)', page_1)):
        return 'TUW', nr_polisy.group(1)
    if 'TUZ' in page_1 and (nr_polisy := re.search('WNIOSEK seria (\w+) nr (\d+)', page_1)):
        return 'TUZ', nr_polisy.group(1) + nr_polisy.group(2)
    if 'WARTA' in page_1 and (nr_polisy := re.search('POLISA NR: *(\d+)', page_1)):
        return 'WAR', nr_polisy.group(1)
    if 'Wiener' in page_1 and (nr_polisy := re.search('Seria i numer (\w+\d+)', page_1)):
        return 'WIE', nr_polisy.group(1)
    else:
        return 'Nie rozpoznałem polisy!'


def tacka_na_polisy(obj):
    if obj.endswith('.pdf'):
        yield from rozpoznanie_danych(obj)
    else:
        for file in os.listdir(obj):
            if file.endswith('.pdf'):
                pdf = obj + '\\' + file
                yield from rozpoznanie_danych(pdf)


def rozpoznanie_danych(pdf):
    page_1, page_1_tok = polisa(pdf)[0], polisa(pdf)[1]
    page_1_box = polisa_box(pdf)
    d = dict(enumerate(page_1_tok))

    nazwisko, imie = nazwisko_imie(d)
    p_lub_r = pesel_regon(d)
    nazwa_firmy, ulica_f, nr_ulicy_f, nr_lok, kod_poczt_f, miasto_f, tel, email = regon(p_lub_r[1:])
    ulica_f_edit = f'{ulica_f} {nr_ulicy_f}' if not nr_lok else f'{ulica_f} {nr_ulicy_f} m {nr_lok}'
    kod_poczt_f_edit = f'{kod_poczt_f[:2]}-{kod_poczt_f[2:]}' if '-' not in kod_poczt_f else kod_poczt_f
    kod_poczt = kod_pocztowy(page_1)
    data_wyst = data_wystawienia()
    tow_ub = numer_polisy(page_1)[0]
    nr_polisy = numer_polisy(page_1)[1]

    return nazwa_firmy, nazwisko, imie, p_lub_r, ulica_f_edit, kod_poczt, miasto_f, tel, email, data_wyst, \
           tow_ub, nr_polisy


for i in tacka_na_polisy(obj):
    if i != '' and i != ' ':
        print(i)
print()

nazwa_firmy, nazwisko, imie, p_lub_r, ulica_f_edit, kod_poczt, miasto_f, tel, email, data_wyst, tow_ub, nr_polisy \
    = rozpoznanie_danych(obj)






"""Zapisanie w Bazie"""
#
# # Sprawdza czy arkusz jest otwarty
# try:
#     ExcelApp = win32com.client.GetActiveObject('Excel.Application')
#     wb = ExcelApp.Workbooks("DTESTY.xlsx")
#     ws = wb.Worksheets("Arkusz1")
#     # workbook = ExcelApp.Workbooks("Baza.xlsx")
#
# # Jeżeli arkusz jest zamknięty, otwiera go
# except:
#     ExcelApp = Dispatch("Excel.Application")
#     wb = ExcelApp.Workbooks.Open(path + "\\TESTY.xlsx")
#     ws = wb.Worksheets("Arkusz1")
#
#
# """Rozpoznaje kolejny wiersz, który może zapisać."""
# row_to_write = wb.Worksheets(1).Cells(wb.Worksheets(1).Rows.Count, 30).End(-4162).Row + 1
#
#
# # ExcelApp.Cells(row_to_write, 1).Value = data_wystawienia()[:2] # Komórka tylko do testów
# ExcelApp.Cells(row_to_write, 7).Value = 'Robert'
# ExcelApp.Cells(row_to_write, 10).Value = 'Grzelak'
# ExcelApp.Cells(row_to_write, 11).Value = nazwa_firmy
# ExcelApp.Cells(row_to_write, 12).Value = nazwisko_imie(d).split()[0] if nazwisko_imie(d) else ''
# ExcelApp.Cells(row_to_write, 13).Value = nazwisko_imie(d).split()[1] if nazwisko_imie(d) else ''
# ExcelApp.Cells(row_to_write, 14).Value = pesel_regon(d)
# # ExcelApp.Cells(row_to_write, 15).Value = data_pr_j
# ExcelApp.Cells(row_to_write, 16).Value = f'{ulica_f} {nr_ulicy_f}' if not nr_lok else f'{ulica_f} {nr_ulicy_f} m {nr_lok}'
# ExcelApp.Cells(row_to_write, 17).Value = kod_pocztowy(page_1) if not kod_poczt_f else kod_poczt_f_edit
# ExcelApp.Cells(row_to_write, 18).Value = miasto_f
# ExcelApp.Cells(row_to_write, 19).Value = tel
# ExcelApp.Cells(row_to_write, 20).Value = email
# # ExcelApp.Cells(row_to_write, 23).Value = marka
# # ExcelApp.Cells(row_to_write, 24).Value = model
# # ExcelApp.Cells(row_to_write, 25).Value = nr_rej
# # ExcelApp.Cells(row_to_write, 26).Value = rok_prod
# # ExcelApp.Cells(row_to_write, 29).Value = int(ile_dni) + 1
#
# # ExcelApp.Cells(row_to_write, 30).NumberFormat = 'yy-mm-dd'
# ExcelApp.Cells(row_to_write, 30).Value = data_wystawienia()
# # ExcelApp.Cells(row_to_write, 31).Value = data_pocz
# # ExcelApp.Cells(row_to_write, 32).Value = data_konca
# # ExcelApp.Cells(row_to_write, 36).Value = 'SPÓŁKA'
# # tor = ExcelApp.Cells(row_to_write, 37).Value = tow
# # ExcelApp.Cells(row_to_write, 38).Value = tow
# # ExcelApp.Cells(row_to_write, 39).Value = rodzaj
# # ExcelApp.Cells(row_to_write, 40).Value = nr_polisy
# # ExcelApp.Cells(row_to_write, 41).Value = nowa_wzn
# # ExcelApp.Cells(row_to_write, 42).Value = nr_wzn
# # if wzn_idx:
# #     ExcelApp.Cells(row_to_write, 41).Value = 'W'
# #     ExcelApp.Cells(row_to_write, 42).Value = nowa_wzn
# # else:
# #     ExcelApp.Cells(row_to_write, 41).Value = 'N'
# #     ExcelApp.Cells(row_to_write, 42).Value = ''
#
# # ryzyko = ExcelApp.Cells(row_to_write, 46).Value = 'b/d'
# # ExcelApp.Cells(row_to_write, 48).Value = przypis
# # ExcelApp.Cells(row_to_write, 49).Value = ter_platnosci
# # if I_rata_data:
# #     ExcelApp.Cells(row_to_write, 49).Value = I_rata_data
# # ExcelApp.Cells(row_to_write, 50).Value = przypis
# # if I_rata_data:
# #     ExcelApp.Cells(row_to_write, 50).Value = I_rata_wart
# # ExcelApp.Cells(row_to_write, 51).Value = f_platnosci
# #
# # ExcelApp.Cells(row_to_write, 52).Value = ilosc_rat
# # ExcelApp.Cells(row_to_write, 53).Value = ilosc_rat
# # data_inkasa = ExcelApp.Cells(row_to_write, 54).Value = ter_platnosci
# # ExcelApp.Cells(row_to_write, 55).Value = przypis
# # ExcelApp.Cells(row_to_write, 59).Value = tow
#
#
# # if II_rata_data:
# #     owca = ExcelApp.Cells(row_to_write + 1, 7).Value = 'Robert'
# #     podpis = ExcelApp.Cells(row_to_write + 1, 10).Value = 'Grzelak'
# #     ExcelApp.Cells(row_to_write + 1, 13).Value = imie
# #     ExcelApp.Cells(row_to_write + 1, 12).Value = nazwisko
# #     ExcelApp.Cells(row_to_write + 1, 14).Value = 'p' + pesel
# #     ExcelApp.Cells(row_to_write + 1, 15).Value = data_pr_j
# #     ExcelApp.Cells(row_to_write + 1, 16).Value = ulica
# #     ExcelApp.Cells(row_to_write + 1, 17).Value = kod_poczt
# #     ExcelApp.Cells(row_to_write + 1, 18).Value = miasto
# #     # tel = ExcelApp.Cells(row_to_write, 19).Value = int('5001900')
# #     # email = ExcelApp.Cells(row_to_write, 20).Value = 'malpa@gmail.pl'
# #     ExcelApp.Cells(row_to_write + 1, 23).Value = marka
# #     ExcelApp.Cells(row_to_write + 1, 24).Value = model
# #     ExcelApp.Cells(row_to_write + 1, 25).Value = nr_rej
# #     ExcelApp.Cells(row_to_write + 1, 26).Value = rok_prod
# #     # data_podpi = ExcelApp.Cells(row_to_write, 30).Value = '18.02.2019'
# #     ExcelApp.Cells(row_to_write + 1, 31).Value = data_pocz
# #     ExcelApp.Cells(row_to_write + 1, 32).Value = data_konca
# #     firma = ExcelApp.Cells(row_to_write, 36).Value = 'SPÓŁKA'
# #     # tor = ExcelApp.Cells(row_to_write, 37).Value = 'GEN'
# #     # tow = ExcelApp.Cells(row_to_write, 38).Value = 'GEN'
# #     # rodz = ExcelApp.Cells(row_to_write, 39).Value = 'kom'
# #     ExcelApp.Cells(row_to_write + 1, 40).Value = nr_polisy
# #     # nowa_wzn = ExcelApp.Cells(row_to_write, 41).Value = 'N'
# #     # nr_wzn = ExcelApp.Cells(row_to_write, 42).Value = '908568823555'
# #     # ryzyko = ExcelApp.Cells(row_to_write, 46).Value = 'b/d'
# #     ExcelApp.Cells(row_to_write + 1, 48).Value = ''
# #     ExcelApp.Cells(row_to_write + 1, 49).Value = ter_platnosci
# #     # ExcelApp.Cells(row_to_write + 1, 49).Value = I_rata_data   ###
# #     ExcelApp.Cells(row_to_write + 1, 49).Value = II_rata_data
# #     ExcelApp.Cells(row_to_write + 1, 50).Value = II_rata_wart
# #     ExcelApp.Cells(row_to_write + 1, 51).Value = f_platnosci
#
#
#
# """Opcje zapisania"""
# # ExcelApp.DisplayAlerts = False
# # wb.SaveAs(path + "\\TESTY.xlsx")
# # wb.Close()
# # ExcelApp.DisplayAlerts = True
#
#
#
#
#
#
#
#
#
#



# def postal_code(data, nazwisko):
#     # kod_poczt = re.compile('\d{,2}-\d{3,}')
#     dystans = [data[data.index(nazwisk) - 6: data.index(nazwisk) + 15] for nazwisk in data if nazwisk.lower() == nazwisko.lower()][0]