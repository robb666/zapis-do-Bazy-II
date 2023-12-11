import requests
import json
import win32com.client
from win32com.client import Dispatch
import time
from openpyxl import load_workbook
from creds import key
from icecream import ic


start_time = time.time()

wb = load_workbook(filename="M:/Agent baza/Login_Hasło.xlsm", read_only=True)
ws = wb['Aplikacje']

in_l = ws['F21'].value
in_h = ws['G21'].value

headers = {
    "Authorization": f"Bearer {key}",
    "Content-Type": "application/json"
}

payload = {
    "username": in_l,
    "password": in_h,
    "ajax_url": "/api/policy/getpolicy",
    "output": "json",
    "policy_oid": "10340324",
    # "policy_oid": "10274225",
    "return_objects": "1"
}


try:
    response = requests.post('https://magro2-api.insly.pl/api/policy/getpolicy',
                             headers=headers,
                             json=payload)
    r = response.json()
    ic(r)

except requests.exceptions.HTTPError as errh:
    print(f"Http Error: {errh}")
except requests.exceptions.ConnectionError as errc:
    print(f"Error Connecting: {errc}")
except requests.exceptions.Timeout as errt:
    print(f"Timeout Error: {errt}")
except requests.exceptions.RequestException as err:
    print(f"Error: {err}")



def regon_checksum(r: str):
    if len(r) == 9:
        regon = list(str(r))
        suma = (int(regon[0]) * 8 + int(regon[1]) * 9 + int(regon[2]) * 2 + int(regon[3]) * 3 + int(regon[4]) * 4 +
                int(regon[5]) * 5 + int(regon[6]) * 6 + int(regon[7]) * 7) % 11
        if suma == int(regon[-1]) or suma == 10 and int(regon[-1]) == 0:
            return True
        else:
            return False
    else:
        return False


def pesel_checksum(p):
    if len(p) == 11:
        l = int(p[10])
        suma = 1 * int(p[0]) + 3 * int(p[1]) + 7 * int(p[2]) + 9 * int(p[3]) + 1 * int(p[4]) + 3 * int(p[5]) + \
               7 * int(p[6]) + 9 * int(p[7]) + 1 * int(p[8]) + 3 * int(p[9]) + 1 * int(p[10])
        lm = suma % 10  # dzielenie wyniku modulo 10
        kontrola = 10 - lm  # sprawdzenie ostatniej liczby kontrolnej
        # w przypadku liczby kontrolnej 10 i 0 sa jednoznaczne a 0 moze byc wynikiem odejmowania
        if (kontrola == 10 or l == kontrola) and p[2:4] != '00':
            return True
        else:
            return False
    else:
        return False


pesel = pesel_checksum(r['customer_idcode'])
regon = regon_checksum(r['customer_idcode'])

nazwa_firmy = r['customer_name'] if regon else ''
nazwisko = r['customer_name'].split()[-1] if nazwa_firmy == '' else ''
imie = r['customer_name'].split()[0] if nazwa_firmy == '' else ''

p_lub_r = r['customer_idcode'] if pesel else r['customer_idcode'] if regon else ''
ulica = r['address'][0]['customer_address_street']
kod_poczt = r['address'][0]['customer_address_zip']
miasto = r['address'][0]['customer_address_city']
tel = r['customer_mobile'] if r['customer_mobile'] != '' else r['customer_phone']
email = r['customer_email']

marka = r['objects'][0]['vehicle_make']


print(nazwa_firmy)
print(nazwisko)
print(imie)
print(p_lub_r)
print(ulica)
print(kod_poczt)
print(miasto)
print(tel)
print(email)
print(marka)


# nazwa_firmy, nazwisko, imie, p_lub_r, pr_j, ulica_f_edit, kod_poczt, miasto_f, tel, email, marka, kod, model, \
# miasto, nr_rej, adres, rok, data_wyst, data_konca, tow_ub_tor, tow_ub, nr_polisy, przypis, ter_platnosci, rata_I, \
# f_platnosci, ilosc_rat, nr_raty, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV = dane_polisy





# try:
#     """Sprawdza czy arkusz jest otwarty."""
#     ExcelApp = win32com.client.GetActivefolderect('Excel.Application')
#     wb = ExcelApp.Workbooks("2014 BAZA MAGRO.xlsm")
#     ws = wb.Worksheets("BAZA 2014")
#     # workbook = ExcelApp.Workbooks("Baza.xlsx")
#
# except:
#     """Jeżeli arkusz jest zamknięty, otwiera go."""
#     ExcelApp = Dispatch("Excel.Application")
#
#     # Exec
#     # wb = ExcelApp.Workbooks.OpenXML("M:\\Agent baza\\2014 BAZA MAGRO.xlsm")
#     wb = ExcelApp.Workbooks.OpenXML("C:\\Users\\PipBoy3000\\Desktop\\2014 BAZA MAGRO.xlsm")
#
#     # Testy
#     # wb = ExcelApp.Workbooks.OpenXML(path + "\\2014 BAZA MAGRO.xlsm")
#
#     ws = wb.Worksheets("BAZA 2014")
#
# ExcelApp.Visible = True
#
# """Jesienne Bazie"""
# # try:
# for dane_polisy in tacka_na_polisy(folder):
#     nazwa_firmy, nazwisko, imie, p_lub_r, pr_j, ulica_f_edit, kod_poczt, miasto_f, tel, email, marka, kod, model, \
#     miasto, nr_rej, adres, rok, data_wyst, data_konca, tow_ub_tor, tow_ub, nr_polisy, przypis, ter_platnosci, rata_I, \
#     f_platnosci, ilosc_rat, nr_raty, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV = dane_polisy
#     print(dane_polisy)
#
#     """Rozpoznaje kolejny wiersz, który może zapisać."""
#     row_to_write = wb.Worksheets(1).Cells(wb.Worksheets(1).Rows.Count, 30).End(-4162).Row + 1
#
#     # Rok_przypisu = ExcelApp.Cells(row_to_write, 1).Value = data_wyst[:2] # Komórka tylko do testów
#     Rozlicz = ExcelApp.Cells(row_to_write, 7).Value = 'Robert'
#     Podpis = ExcelApp.Cells(row_to_write, 10).Value = 'Grzelak'
#     FIRMA = ExcelApp.Cells(row_to_write, 11).Value = nazwa_firmy
#     Nazwisko = ExcelApp.Cells(row_to_write, 12).Value = nazwisko
#     Imie = ExcelApp.Cells(row_to_write, 13).Value = imie
#     Pesel_Regon = ExcelApp.Cells(row_to_write, 14).Value = p_lub_r
#     ExcelApp.Cells(row_to_write, 15).Value = pr_j
#     ExcelApp.Cells(row_to_write,
#                    16).Value = ulica_f_edit  # f'{ulica_f} {nr_ulicy_f}' if not nr_lok else f'{ulica_f} {nr_ulicy_f} m {nr_lok}'
#     ExcelApp.Cells(row_to_write, 17).Value = kod_poczt  # kod_pocztowy(page_1) if not kod_poczt_f else kod_poczt_f_edit
#     ExcelApp.Cells(row_to_write, 18).Value = miasto_f
#     ExcelApp.Cells(row_to_write, 19).Value = tel
#     ExcelApp.Cells(row_to_write, 20).Value = email.lower() if email else ''
#     ExcelApp.Cells(row_to_write, 23).Value = marka if marka else kod
#     ExcelApp.Cells(row_to_write, 24).Value = model if model else miasto
#     ExcelApp.Cells(row_to_write, 25).Value = nr_rej if nr_rej else adres
#     ExcelApp.Cells(row_to_write, 26).Value = rok
#     # ExcelApp.Cells(row_to_write, 29).Value = int(ile_dni) + 1
#     # ExcelApp.Cells(row_to_write, 30).NumberFormat = 'yy-mm-dd'
#     ExcelApp.Cells(row_to_write, 30).Value = data_wyst
#     # ExcelApp.Cells(row_to_write, 31).Value = data_pocz
#     ExcelApp.Cells(row_to_write, 32).Value = data_konca
#     ExcelApp.Cells(row_to_write, 36).Value = 'SPÓŁKA'
#     tor = ExcelApp.Cells(row_to_write, 37).Value = tow_ub_tor
#     ExcelApp.Cells(row_to_write, 38).Value = tow_ub
#     # ExcelApp.Cells(row_to_write, 39).Value = rodzaj
#     ExcelApp.Cells(row_to_write, 40).Value = nr_polisy
#     # ExcelApp.Cells(row_to_write, 41).Value = nowa_wzn
#     # ExcelApp.Cells(row_to_write, 42).Value = nr_wzn
#     # if wzn_idx:
#     #     ExcelApp.Cells(row_to_write, 41).Value = 'W'
#     #     ExcelApp.Cells(row_to_write, 42).Value = nowa_wzn
#     # else:
#     #     ExcelApp.Cells(row_to_write, 41).Value = 'N'
#     #     ExcelApp.Cells(row_to_write, 42).Value = ''
#     # ExcelApp.Cells(row_to_write, 46).Value = ryzyko
#     ExcelApp.Cells(row_to_write, 48).Value = przypis
#     ExcelApp.Cells(row_to_write, 49).Value = ter_platnosci
#     # if I_rata_data:
#     #     ExcelApp.Cells(row_to_write, 49).Value = I_rata_data
#     if rata_I:
#         ExcelApp.Cells(row_to_write, 50).Value = rata_I
#     else:
#         ExcelApp.Cells(row_to_write, 50).Value = przypis
#     ExcelApp.Cells(row_to_write, 51).Value = f_platnosci
#     ExcelApp.Cells(row_to_write, 52).Value = ilosc_rat
#     ExcelApp.Cells(row_to_write, 53).Value = nr_raty
#     data_inkasa = ExcelApp.Cells(row_to_write, 54).Value = ter_platnosci
#     if rata_I:
#         ExcelApp.Cells(row_to_write, 55).Value = rata_I
#     else:
#         ExcelApp.Cells(row_to_write, 55).Value = przypis
#     ExcelApp.Cells(row_to_write, 58).Value = 'aut'
#     ExcelApp.Cells(row_to_write, 60).Value = tow_ub_tor
#
#     if rata_II:
#         ws.Range(f'A{row_to_write}:BH{row_to_write}').Copy()
#         ws.Range(f'A{row_to_write + 1}').PasteSpecial()
#
#         ExcelApp.Cells(row_to_write + 1, 48).Value = ''
#         ExcelApp.Cells(row_to_write + 1, 49).Value = termin_II
#         ExcelApp.Cells(row_to_write + 1, 50).Value = rata_II
#         ExcelApp.Cells(row_to_write + 1, 53).Value = 2
#         data_inkasa = ExcelApp.Cells(row_to_write + 1, 54).Value = ''
#         ExcelApp.Cells(row_to_write + 1, 55).Value = ''
#
#         if rata_III:
#             ws.Range(f'A{row_to_write + 1}:BH{row_to_write + 1}').Copy()
#             ws.Range(f'A{row_to_write + 2}').PasteSpecial()
#
#             ExcelApp.Cells(row_to_write + 2, 49).Value = termin_III
#             ExcelApp.Cells(row_to_write + 2, 50).Value = rata_III
#             ExcelApp.Cells(row_to_write + 2, 53).Value = 3
#
#             if rata_IV:
#                 ws.Range(f'A{row_to_write + 2}:BH{row_to_write + 2}').Copy()
#                 ws.Range(f'A{row_to_write + 3}').PasteSpecial()
#
#                 ExcelApp.Cells(row_to_write + 3, 49).Value = termin_IV
#                 ExcelApp.Cells(row_to_write + 3, 50).Value = rata_IV
#                 ExcelApp.Cells(row_to_write + 3, 53).Value = 4
#
# """Opcje zapisania"""
#
# # Exec
# ExcelApp.DisplayAlerts = False
#
# ## Bez zapisania
# # wb.SaveAs("M:\\Agent baza\\2014 BAZA MAGRO.xlsm")
#
# # Testy
# # wb.SaveAs(path + "\\2014 BAZA MAGRO.xlsm")
#
# """Zamknięcie narazie wyłączone..."""
# # wb.Close()
# ExcelApp.DisplayAlerts = True
#
# # ExcelApp.Application.Quit()
#
# end_time = time.time() - start_time
# print('Czas wykonania: {:.2f} sekund'.format(end_time))