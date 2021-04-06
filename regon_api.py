"""pip install RegonAPI==1.1.0"""
from RegonAPI import RegonAPI
import pprint as pp

API_KEY = "f98f22f8c4cc439ca677"
api = ''
try:
    api = RegonAPI(bir_version="bir1.1", is_production=True) # BIR version 1.1
except:
    print(f'Sprawdź czy baza REGON jest dostępna. \n\n{e}')
api.authenticate(key=API_KEY)


def validate_regon(r: int):
    """Waliduje sprawdzając sumę kontrolną regon."""
    regon = list(str(r))
    suma = (int(regon[0])*8 + int(regon[1])*9 + int(regon[2])*2 + int(regon[3])*3 + int(regon[4])*4 +
            int(regon[5])*5 + int(regon[6])*6 + int(regon[7])*7) % 11
    if suma == int(regon[-1]) or suma == 10 and int(regon[-1]) == 0:
        return r


def os_praw(REGON9):
    res_os_prawna = api.dataDownloadFullReport(REGON9, "BIR11OsPrawna")
    res_os_prawna_pkd = api.dataDownloadFullReport(REGON9, "BIR11OsPrawnaPkd")

    gus_praw_nazwa = res_os_prawna[0]['praw_nazwa']
    gus_praw_nip = res_os_prawna[0]['praw_nip']
    gus_praw_ul = res_os_prawna[0]['praw_adSiedzUlica_Nazwa']
    gus_praw_nr_ul = res_os_prawna[0]['praw_adSiedzNumerNieruchomosci']
    gus_praw_nr_lok = res_os_prawna[0]['praw_adSiedzNumerLokalu']
    gus_praw_kod_poczt = res_os_prawna[0]['praw_adSiedzKodPocztowy']
    gus_praw_miasto = res_os_prawna[0]['praw_adSiedzMiejscowoscPoczty_Nazwa']
    gus_praw_tel = res_os_prawna[0]['praw_numerTelefonu']
    gus_praw_mail = res_os_prawna[0]['praw_adresEmail']
    gus_praw_pkd = res_os_prawna_pkd[0]['praw_pkdKod']
    gus_praw_pkd_opis = res_os_prawna_pkd[0]['praw_pkdNazwa']
    gus_data_rozp = res_os_prawna[0]['praw_dataRozpoczeciaDzialalnosci']

    # gus_praw_adres = f'{gus_praw_ul} {gus_praw_nr_ul} {gus_praw_nr_lok}'
    # if gus_praw_adres.startswith('ul. '):
    #     gus_praw_adres = gus_praw_adres.replace('ul. ', '')

    return gus_praw_nazwa, gus_praw_nip, gus_praw_ul, gus_praw_nr_ul, gus_praw_nr_lok, gus_praw_kod_poczt, gus_praw_miasto, \
           gus_praw_tel, gus_praw_mail, gus_praw_pkd, gus_praw_pkd_opis, gus_data_rozp


def os_fiz(REGON9):
    gus_fiz_ceidg = api.dataDownloadFullReport(REGON9, "BIR11OsFizycznaDzialalnoscCeidg")
    gus_fiz = api.dataDownloadFullReport(REGON9, "BIR11OsFizycznaDaneOgolne")
    gus_fiz_pkd = api.dataDownloadFullReport(REGON9, "BIR11OsFizycznaPkd")

    gus_fiz_nazwa = gus_fiz_ceidg[0]['fiz_nazwa']
    gus_fiz_nip = gus_fiz[0]['fiz_nip']
    gus_imie = gus_fiz[0]['fiz_imie1']
    gus_nazwisko = gus_fiz[0]['fiz_nazwisko']
    gus_fiz_ul = gus_fiz_ceidg[0]['fiz_adSiedzUlica_Nazwa']
    gus_fiz_nr_ul = gus_fiz_ceidg[0]['fiz_adSiedzNumerNieruchomosci']
    gus_fiz_nr_lok = gus_fiz_ceidg[0]['fiz_adSiedzNumerLokalu']
    gus_fiz_kod_poczt = gus_fiz_ceidg[0]['fiz_adSiedzKodPocztowy']
    gus_fiz_miasto = gus_fiz_ceidg[0]['fiz_adSiedzMiejscowoscPoczty_Nazwa']
    gus_fiz_tel = gus_fiz_ceidg[0]['fiz_numerTelefonu']
    gus_fiz_mail = gus_fiz_ceidg[0]['fiz_adresEmail']
    gus_fiz_pkd_ = gus_fiz_pkd[0]['fiz_pkd_Kod']
    gus_fiz_pkd_opis = gus_fiz_pkd[0]['fiz_pkd_Nazwa']
    gus_fiz_data_rozp = gus_fiz_ceidg[0]['fiz_dataRozpoczeciaDzialalnosci']
    # wsio = gus_fiz_ceidg

    # gus_fiz_adres = f'{gus_fiz_ul} {gus_fiz_nr_ul} {gus_fiz_nr_lok}'
    # if gus_fiz_adres.startswith('ul. '):
    #     gus_fiz_adres = gus_fiz_adres.replace('ul. ', '')

    return gus_fiz_nazwa, gus_fiz_nip, gus_nazwisko, gus_imie, gus_fiz_ul, gus_fiz_nr_ul, gus_fiz_nr_lok, gus_fiz_kod_poczt, \
           gus_fiz_miasto, gus_fiz_tel, gus_fiz_mail, gus_fiz_pkd_, gus_fiz_pkd_opis, gus_fiz_data_rozp


def get_regon_data(r):
    """Pobiera dane z GUS. Jak nie łączyć z API za każdym wywołaniem zmiennej?"""
    if validate_regon(r):
        REGON9 = r
        try:
            return {'forma': 'Osoba prawna',
                    'nazwa': os_praw(REGON9)[0],
                    'nip': os_praw(REGON9)[1],
                    'ul': os_praw(REGON9)[2],
                    'nr_ul': os_praw(REGON9)[3],
                    'nr_lok': os_praw(REGON9)[4],
                    'kod_poczt': os_praw(REGON9)[5],
                    'miasto': os_praw(REGON9)[6],
                    'tel': os_praw(REGON9)[7],
                    'email': os_praw(REGON9)[8],
                    'pkd': os_praw(REGON9)[9],
                    'opis pkd': os_praw(REGON9)[10],
                    'data rozpoczęcia': os_praw(REGON9)[11],
                    }

        except:
            return {'forma': 'Jednoosobowa dz.g.',
                    'nazwa': os_fiz(REGON9)[0],
                    'nip': os_fiz(REGON9)[1],
                    'imie': os_fiz(REGON9)[3],
                    'nazwisko': os_fiz(REGON9)[2],
                    'ul': os_fiz(REGON9)[4],
                    'nr_ul': os_fiz(REGON9)[5],
                    'nr_lok': os_fiz(REGON9)[6],
                    'kod_poczt': os_fiz(REGON9)[7],
                    'miasto': os_fiz(REGON9)[8],
                    'tel': os_fiz(REGON9)[9],
                    'email': os_fiz(REGON9)[10],
                    'pkd': os_fiz(REGON9)[11],
                    'opis pkd': os_fiz(REGON9)[12],
                    'data rozpoczęcia': os_fiz(REGON9)[13],
                    }
    else:
        return print('\nBaza REGON\n==========\nPodany numer REGON jest błędny!\n')


# pp.pprint(get_regon_data(100416810))
