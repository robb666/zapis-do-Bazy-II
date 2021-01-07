import os
import re
import pdfplumber
from datetime import datetime, timedelta
import win32com.client
from win32com.client import Dispatch
from regon_api import get_regon_data
import time

start_time = time.time()
path = os.getcwd()
# obj = input('Podaj polisę/y w formacie .pdf do rejestracji: ')
obj = r'C:\Users\ROBERT\Desktop\IT\PYTHON\PYTHON 37 PROJEKTY\excel\zapis do Bazy II\polisy\II partia\Policy_BPPAP_98439.pdf'

one_day = timedelta(1)

def words_separately(text):
    """Tokenizuje tekst całej polisy."""
    make_tokens = re.compile(r"((?:(?<!'|\w)(?:\w-?'?)+(?<!-))|(?:(?<='|\w)(?:\w-?'?)+(?=')))")
    return make_tokens.findall(text)


def polisa(pdf):
    """Tekst całej 1 str. polisy."""
    with pdfplumber.open(pdf) as policy:
        page_1 = policy.pages[0].extract_text()  # Tylko pierwsza strona
    return page_1, words_separately(page_1.lower())

def polisa_str(pdf):
    """Tekst 3 str. polisy."""
    with pdfplumber.open(pdf) as policy:
        page_1 = policy.pages[0].extract_text()
        page_2 = policy.pages[1].extract_text()
        page_3 = policy.pages[2].extract_text()
    return page_1 + page_2 + page_3


def polisa_box(pdf, left, top, right, bottom):
    """Tekst wybranego fragmentu polisy."""
    with pdfplumber.open(pdf) as policy:
        # width = policy.pages[0].width
        # height = policy.pages[0].height
        # box_left = (0, 0, 150, 140)
        box_center = (left, top, right, bottom)
        # box_right = (150, 0, width, 140)
        page_1_box = policy.pages[0].within_bbox(box_center, relative=False).extract_text()
    return page_1_box #words_separately(page_1_box.lower())


def pesel_checksum(p):
    """Suma kontrolna nr pesel."""
    l = int(p[10])
    suma = ((1*int(p[0])) + 3*int(p[1]) + (7*int(p[2])) + (9*int(p[3])) + (1*int(p[4])) + (3*int(p[5])) +
            (7*int(p[6])) + (9*int(p[7])) + (1*int(p[8])) + (3*int(p[9])))
    lm = (suma % 10)  # dzielenie wyniku modulo 10
    kontrola = (10 - lm)  # sprawdzenie ostatniej liczby kontrolnej
    # w przypadku liczby kontrolnej 10 i 0 sa jednoznaczne a 0 moze byc wynikiem odejmowania
    if (kontrola == 10 or l == kontrola) and p[2:4] != '00':
        return 1
    else:
        return 0


def regon_checksum(r: int):
    """Waliduje regon sprawdzając sumę kontrolną."""
    regon = list(str(r))
    if len(regon) == 9:
        suma = (int(regon[0])*8 + int(regon[1])*9 + int(regon[2])*2 + int(regon[3])*3 + int(regon[4])*4 +
                int(regon[5])*5 + int(regon[6])*6 + int(regon[7])*7) % 11
        if suma == int(regon[-1]) or suma == 10 and int(regon[-1]) == 0:
            return r
        else:
            return 0
    else:
        return ''


def regon(regon_checksum):
    """API Regon"""
    if len(regon_checksum) == 9:
        print('\nCzekam na dane z bazy REGON...')

        osoba = get_regon_data(regon_checksum)['forma']
        imie, nazwisko = '', ''
        if osoba == 'Jednoosobowa dz.g.':
            imie = get_regon_data(regon_checksum)['imie']
            nazwisko = get_regon_data(regon_checksum)['nazwisko']
        nazwa_firmy = get_regon_data(regon_checksum)['nazwa'].title()
        nip = get_regon_data(regon_checksum)['nip']
        ulica_f = get_regon_data(regon_checksum)['ul'].lstrip('ul. ')
        nr_ulicy_f = get_regon_data(regon_checksum)['nr_ul']
        nr_lok = get_regon_data(regon_checksum)['nr_lok']
        kod_poczt_f = get_regon_data(regon_checksum)['kod_poczt']
        miasto_f = get_regon_data(regon_checksum)['miasto']
        pkd = get_regon_data(regon_checksum)['pkd']
        opis_pkd = get_regon_data(regon_checksum)['opis pkd']
        data_rozp = get_regon_data(regon_checksum)['data rozpoczęcia']
        tel = get_regon_data(regon_checksum)['tel']
        email = get_regon_data(regon_checksum)['email'].lower()

        return nazwa_firmy, ulica_f, nr_ulicy_f, nr_lok, kod_poczt_f, miasto_f, tel, email
    else:
        return '', '', '', '', '', '', '', ''


"""Funkcje odpowiadają kolumnom w bazie."""
def nazwisko_imie(d):
    """Zwraca imię i nazwisko Klienta."""
    with open(path + '\\imiona.txt') as content:
        all_names = content.read().split('\n')
        if 'euroins' in d.values():
            name = []
            for k, v in d.items():
                if v.title() in all_names and not re.search('\d', d[k + 4]):
                    name.append(f'{d[k + 4].title()} {v.title()}')
                if v.title() in all_names and re.search('\d', d[k + 4]):
                    name.append(f'{d[k + 5].title()} {v.title()}')
        elif 'tuz' in d.values():
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
    regon = [regon for k, regon in d.items() if k < 200 and len(regon) == 9 and re.search('\d{9}', regon) and
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
    # print(page_1)
    c = re.compile('(adres\w*|(?<!InterRisk) kontakt\w*|pocztowy|ubezpieczony)', re.I)
    # print(c)

    if (f := c.search(page_1)):
        adres = f.group().strip()
        # print(adres)

    data = page_1.split()
    # print(data)
    dystans = [data[data.index(split) - 10: data.index(split) + 33] for split in data if adres in split][0]
    # print(dystans)
    kod_pocztowy = [kod for kod in dystans if re.search('\d{2}[-|\xad]\d{3}', kod)][0]
    return kod_pocztowy


def data_wystawienia():
    # one_day = timedelta(1)
    today = datetime.strptime(datetime.now().strftime('%y-%m-%d'), '%y-%m-%d') + one_day
    return today # datetime.today().date().strftime('%y-%m-%d')


def koniec_ochrony(page_1):
    # one_day = timedelta(1)
    daty = re.compile(r'(\b\d{2}[-|.|/]\d{2}[-|.|/]\d{4}|\b\d{4}[-|.|/]\d{2}[-|.|/]\d{2})')
    lista_dat = [re.sub('[^0-9]', '-', data) for data in daty.findall(page_1)]
    jeden_format = [re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', date) for date in lista_dat]
    koniec = max(datetime.strptime(data, '%Y-%m-%d') for data in jeden_format)
    koniec_absolutny = datetime.strptime(datetime.strftime(koniec, '%y-%m-%d'), '%y-%m-%d') + one_day
    if koniec_absolutny:
        return koniec_absolutny
    else:
        return ''


def TU():
    """W funkcji numer_polisy(page_1)"""
    pass


def numer_polisy(page_1):
    nr_polisy = ''
    if 'Allianz' in page_1 and (nr_polisy := re.search('Polisa nr (\d+)', page_1)):
        return 'ALL', 'ALL', nr_polisy.group(1)
    if 'AXA' in page_1 and (nr_polisy := re.search('(\d{4}-\d+)', page_1)):
        return 'AXA', 'AXA', nr_polisy.group(1)
    if 'Compensa' in page_1 and (nr_polisy := re.search('typ polisy: *\s*(\d+),numer: *\s*(\d+)', page_1)):
        return 'COM', 'COM', nr_polisy.group(1) + nr_polisy.group(2)
    if 'EUROINS' in page_1 and (nr_polisy := re.search('Polisa ubezpieczenia nr: (\d+)', page_1)):
        return 'EIN', 'EIN', nr_polisy.group(1)
    if 'Generali' in page_1 and not 'Proama' in page_1 and (nr_polisy := re.search('POLISA NR\s*(\d+)', page_1, re.I)):
        return 'GEN', 'GEN', nr_polisy.group(1)
    if 'HDI' in page_1 and (nr_polisy := re.search('POLISA NR\s?: *(\d+)', page_1)):
        return 'WAR', 'HDI', nr_polisy.group(1)
    if 'Hestia' in page_1  and not 'MTU' in page_1 and (nr_polisy := re.search('Polisa\s.*\s(\d+)', page_1, re.I)):
        return 'HES', 'HES', nr_polisy.group(1)
    if 'INTER' and (nr_polisy := re.search('polisa\s*seria\s*(\w*)\s*numer\s*(\d*)', page_1)):
        return 'INT', 'INT', nr_polisy.group(1) + nr_polisy.group(2)
    if 'InterRisk' in page_1 and (nr_polisy := re.search('Polisa seria?\s(.*)\snumer\s(\d+)', page_1, re.I)):
        return 'RIS', 'RIS', nr_polisy.group(1) + nr_polisy.group(2)
    if (nr_polisy := re.search('Numer\s(\w\d+)', page_1)):
        return 'LIN', 'LIN',  nr_polisy.group(1)
    if 'MTU' in page_1 and (nr_polisy := re.search('Polisa\s.*\s(\d+)', page_1, re.I)):
        return 'AZ', 'MTU', nr_polisy.group(1)
    if 'Proama' in page_1 and (nr_polisy := re.search('POLISA NR\s*(\d+)', page_1, re.I)):
        return 'GEN', 'PRO', nr_polisy.group(1)
    if 'PZU' in page_1 and (nr_polisy := re.search('Nr *(\d+)', page_1)):
        return 'PZU', 'PZU', nr_polisy.group(1)
    if 'TUW' in page_1 and (nr_polisy := re.search('Wniosko-Polisa\snr\s*(\d+)', page_1, re.I)):
        return 'TUW', 'TUW', nr_polisy.group(1)
    if 'TUZ' in page_1 and (nr_polisy := re.search('WNIOSEK seria (\w+) nr (\d+)', page_1)):
        return 'TUZ', 'TUZ', nr_polisy.group(1) + nr_polisy.group(2)
    if 'UNIQA' in page_1 and (nr_polisy := re.search('Nr (\d{6,})', page_1)):
        return 'UNI', 'UNI', nr_polisy.group(1)
    if 'WARTA' in page_1 and (nr_polisy := re.search('POLISA NR\s?: *(\d+)', page_1)):
        return 'WAR', 'WAR', nr_polisy.group(1)
    if 'Wiener' in page_1 and (nr_polisy := re.search('Seria i numer (\w+\d+)', page_1)):
        return 'WIE', 'WIE', nr_polisy.group(1)
    else:
        return 'Nie rozpoznałem polisy!'


def przypis_daty_raty(pdf, page_1):
    # one_day = timedelta(1)
    total, termin_I, rata_I, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV = \
                                                                         '', '', '', '', '', '', '', '', ''

    if 'AXA' in page_1:
        box = polisa_box(pdf, 0, 250, 590, 650)
        print(box)
        (total := re.search(r'(Składka:|łącznie:) (\d*\s?\d+)', box).group(2))

        if 'Wpłata przelewem' in box or 'Nr konta' in box:
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    if 'Compensa' in page_1:
        box = polisa_box(pdf, 0, 260, 590, 650)
        (total := re.search(r'Składka ogółem: (\d*\s?\d+)', box, re.I))
        total = int(re.sub(r'\xa0', '', total.group(1)))

        if 'składki: kwartalna' in box:
            (rata_I := re.search(r'I rata -  \d{2}.\d{2}.\d{4} - (\d+)', box, re.I).group(1))
            (rata_II := re.search(r'II rata.* - (\d+)', box, re.I).group(1))
            (rata_III := re.search(r'[|III] rata.* - (\d+)', box, re.I).group(1))
            (rata_IV := re.search(r'\n- (\d+)', box, re.I).group(1))

            def terminy(termin):
                zamiana_sep = re.sub('[^0-9]', '-', termin.group(1))
                return re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', zamiana_sep)

            termin_I = terminy(re.search(r'I\srata\s-\s+(\d{2}.\d{2}.\d{4})', box, re.I))
            termin_II = terminy(re.search(r'II rata -  (\d{2}.\d{2}.\d{4})', box, re.I))
            termin_III = terminy(re.search(r',   rata -  (\d{2}.\d{2}.\d{4})', box, re.I))
            termin_IV = terminy(re.search(r'IV rata -  (\d{2}.\d{2}.\d{4})', box, re.I))

            return total, termin_I, rata_I, 'P', 4, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    if 'EUROINS' in page_1:
        box = polisa_box(pdf, 0, 400, 590, 750)
        (total := re.search(r'Łączna składka do zapłaty (\d+,\d*)', box, re.I))
        total = re.sub(r',', '.', total.group(1))

        (termin_I := re.search(r'1. (\d{4}-\d{2}-\d{2})', box, re.I).group(1))

        if 'jednorazowo' in box and 'przelew' in box:
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    if 'Generali' in page_1 and not 'Proama' in page_1:
        box = polisa_box(pdf, 0, 300, 590, 530)
        print(box)

        (total := re.search(r'(RAZEM:|Składka) (\d*\s?\d+)', box, re.I))
        print()
        print(total)
        total = int(re.sub(r' ', '', total.group(2)))

        if 'przelewem' in box:
            (termin := re.search(r'płatna\s?do\s?(\d{2}.\d{2}.\d{4})', box, re.I))
            termin_I = re.sub('[^0-9]', '-', termin.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'została pobrana' in box or 'została opłacona' in box:
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    if 'HDI' in page_1 or '„WARTA” S.A. POTWIERDZA' in page_1:
        box = polisa_box(pdf, 0, 200, 590, 530)
        (total := re.search(r'[ŁĄCZNA SKŁADKA|Składka łączna] (\d*\s?\d+)', box, re.I))
        total = int(re.sub(r' ', '', total.group(1)))

        if 'JEDNORAZOWO' in box and 'PRZELEW' in box:
            (rata_I := re.search(r'kwota: (\d*\s?\d+)', box, re.I).group(1))
            (termin_I := re.search(r'termin płatności: (\d{4}-\d{2}-\d{2})', box, re.I).group(1))
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if '2 RATACH' in box:
            (rata_I := re.search(r'kwota: (\d*\s?\d+)', box, re.I).group(1))
            (rata_II := re.search(r'kwota: (\d*\s?\d+) PLN (\d*\s?\d+)', box, re.I).group(2))

            (termin_I := re.search(r'termin płatności: (\d{4}-\d{2}-\d{2})', box, re.I).group(1))
            (termin_II := re.search(r'termin płatności: (\d{4}-\d{2}-\d{2})\s?(\d{4}-\d{2}-\d{2})', box, re.I).group(2))

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

    if 'Hestia' in page_1 and not 'MTU' in page_1:
        box = polisa_box(pdf, 0, 220, 590, 600)
        print(box)
        (total := re.search(r'DO ZAPŁATY (\d*\s?\d+)', box, re.I))
        total = int(re.sub(r' ', '', total.group(1)))

        if not 'II rata' in box and 'gotówka' in box:
            (termin_I := re.search(r'płatności (\d{4}-\d{2}-\d{2})', box, re.I).group(1))
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if not 'II rata' in box and 'przelew' in box:
            (termin_I := re.search(r'płatności (\d{4}-\d{2}-\d{2})', box, re.I).group(1))
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'II rata' in box and 'przelew' in box:
            (termin_I := re.search(r'płatności I rata (\d{4}-\d{2}-\d{2})', box, re.I).group(1))
            (termin_II := re.search(r'II rata (\d{4}-\d{2}-\d{2})', box, re.I).group(1))
            (rata_I := re.search(fr'{termin_I},  (\d*\s?\d+)', box, re.I).group(1))
            (rata_II := re.search(fr'{termin_II},  (\d*\s?\d+)', box, re.I).group(1))

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

    if 'INTER' in page_1:
        box = polisa_box(pdf, 0, 220, 590, 490)
        print(box)
        (total := re.search(r'kwota składki: (\d*\s?\d+)', box, re.I))
        total = int(re.sub(r' ', '', total.group(1)))

        if re.findall(r'(?=.*jednorazowo)(?=.*przelewem).*', box, re.I | re.DOTALL):
            (termin := re.search(r'i termin płatności:.*\(do (\d{2}.\d{2}.\d{4})', box, re.I))
            termin_I = re.sub('[^0-9]', '-', termin.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    if 'InterRisk' in page_1:

        pdf_str3 = polisa_str(pdf)[1900:-2600]
        print(pdf_str3)

        total_match = re.compile(r'(Składka\słączna:\s*|WYSOKOŚĆ\sSKŁADKI\sŁĄCZNEJ:\n)(\d*\s?\d+)')

        (total := re.search(total_match, pdf_str3))

        print(total)
        total = int(re.sub(r'\xa0', '', total.group(2)))

        if re.findall(r'(?=.*jednorazow[o|a])(?=.*płatności:\s*przelewem).*', pdf_str3, re.I | re.DOTALL):
            (termin_I := re.search(r'płatna\sdo\sdnia:\s(\d{4}-\d{2}-\d{2})', pdf_str3, re.I).group(1))
            print(termin_I)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    # Link4
    if (nr_polisy := re.search('Numer\s(\w\d+)', page_1)):
        box = polisa_box(pdf, 0, 300, 590, 600)
        print(box)
        (total := re.search(r'[\w\s()](\d*\s?\d+,\d+)', box, re.I))
        print()
        print(total.group(1))
        total = float(total.group(1).replace(',', '.').replace(' ', ''))

        if re.findall(r'(?=.*Metoda płatności Karta).*', box):
            (termin := re.search(r'Termin płatności (\d{2}/\d{2}/\d{4})', box, re.I))
            termin_I = re.sub('[^0-9]', '-', termin.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if re.findall(r'(?=.*Przelew)(?=.*Kolejne raty).*', box, re.I | re.DOTALL):

            (raty := re.search(r'Termin Kwota raty.* (\d*\s?\d+,\d{2}).* (\d*\s?\d+,\d{2}).* (\d*\s?\d+,\d{2}).* '
                                                   r'(\d*\s?\d+,\d{2})', box, re.I | re.DOTALL))
            # print(raty.group(2))
            rata_I = float(raty.group(1).replace(',', '.').replace(' ', ''))
            rata_II = float(raty.group(2).replace(',', '.').replace(' ', ''))
            rata_III = float(raty.group(3).replace(',', '.').replace(' ', ''))
            rata_IV = float(raty.group(4).replace(',', '.').replace(' ', ''))

            def termin(terminy, n):
                zamiana_sep = re.sub('[^0-9]', '-', terminy.group(n))
                return re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', zamiana_sep)

            (terminy := re.search(r'Termin Kwota raty\n(\d{2}/\d{2}/\d{4}).*(\d{2}/\d{2}/\d{4}).*(\d{2}/\d{2}/\d{4}).*'
                                                     r'(\d{2}/\d{2}/\d{4}).*', box, re.I | re.DOTALL))
            print(terminy)
            termin_I = datetime.strptime(termin(terminy, 1), '%Y-%m-%d') + one_day
            termin_II = datetime.strptime(termin(terminy, 2), '%Y-%m-%d') + one_day
            termin_III = datetime.strptime(termin(terminy, 3), '%Y-%m-%d') + one_day
            termin_IV = datetime.strptime(termin(terminy, 4), '%Y-%m-%d') + one_day

            return total, termin_I, rata_I, 'P', 4, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    if 'MTU' in page_1:
        box = polisa_box(pdf, 0, 200, 590, 400)

        (total := re.search(r'RAZEM DO ZAPŁATY (\d*\s?\d+)', box, re.I))
        total = int(re.sub(r' ', '', total.group(1)))

        if 'przelew' in box:
            (termin := re.search(r'i kwoty płatności (\d{4}‑\d{2}‑\d{2})', box, re.I))
            termin_I = re.sub('[^0-9]', '-', termin.group(1))

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

    if 'Proama' in page_1:
        box = polisa_box(pdf, 0, 250, 590, 450)

        (total := re.search(r'RAZEM: (\d*\s?\d+)', box, re.I))
        total = int(re.sub(r'\xa0', '', total.group(1)))

        if 'przelewem' in box:
            (termin := re.search(r'płatna\s?do (\d{2}.\d{2}.\d{4})', box, re.I))
            termin_I = re.sub('[^0-9]', '-', termin.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'została pobrana' in box:
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

    if 'PZU' in page_1:
        pdf_str = polisa_str(pdf)[1900:4600]
        (total := re.search(r'Składka łączna: (\d*\s?\d+)', pdf_str, re.I))
        total = int(re.sub(r' ', '', total.group(1)))


        if 'została opłacona w całości.' in pdf_str:
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'Jednorazowo' in pdf_str and 'tytule przelewu' in pdf_str:
            (termin_I := re.search(r'Termin płatności (\d{2}.\d{2}.\d{4})', pdf_str, re.I))
            termin_I = re.sub('[^0-9]', '-', termin_I.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'Rata 1 2 3 4' in pdf_str:
            (raty := re.search(r'Kwota w złotych (\d*\s?\d+,\d{2}) (\d*\s?\d+,\d{2}) (\d*\s?\d+,\d{2}) (\d*\s?\d+,\d{2})',
                               pdf_str, re.I))
            rata_I = float(raty.group(1).replace(',', '.').replace(' ', ''))
            rata_II = float(raty.group(2).replace(',', '.').replace(' ', ''))
            rata_III = float(raty.group(3).replace(',', '.').replace(' ', ''))
            rata_IV = float(raty.group(4).replace(',', '.').replace(' ', ''))

            def termin(terminy, n):
                zamiana_sep = re.sub('[^0-9]', '-', terminy.group(n))
                return re.sub(r'(\d{2})-(\d{2})-(\d{2})', r'\3-\2-\1', zamiana_sep)

            terminy = re.search(r'Termin płatności (\d{2}.\d{2}.\d{2}) (\d{2}.\d{2}.\d{2}) (\d{2}.\d{2}.\d{2}) '
                                r'(\d{2}.\d{2}.\d{2})', pdf_str, re.I)
            termin_I = datetime.strptime(termin(terminy, 1), '%y-%m-%d') + one_day
            termin_II = datetime.strptime(termin(terminy, 2), '%y-%m-%d') + one_day
            termin_III = datetime.strptime(termin(terminy, 3), '%y-%m-%d') + one_day
            termin_IV = datetime.strptime(termin(terminy, 4), '%y-%m-%d') + one_day

            return total, termin_I, rata_I, 'P', 4, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

    if 'TUW' in page_1 and not 'TUZ' in page_1:
        box = polisa_box(pdf, 0, 400, 400, 600)
        (total := re.search(r'Składka łączna: (\d*\s?\d+) PLN', box, re.I))
        total = int(re.sub(r'\xa0', '', total.group(1)))

        (termin := re.search(r'Termin płatności.*(\d{2,4}-\d{2}-\d{2,4})', box, re.I))
        termin_I = re.sub('[^0-9]', '-', termin.group(1))
        termin_I = datetime.strptime(re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I), '%y-%m-%d') + one_day

        if 'JEDNORAZOWO' in box and 'PRZELEW' in box:
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

    if 'TUZ' in page_1:
        pdf_str = polisa_str(pdf)[2000:6500]
        (total := re.search(r'[kwota|Składka] do zapłaty .* (\d*\s?\d+)', pdf_str, re.I))
        total = int(re.sub(r' ', '', total.group(1)))

        if re.findall(r'(?=.*JEDNORAZOWA)(?=.*Przelew).*', pdf_str, re.I | re.DOTALL):
            (termin_I := re.search(r'płatny do dnia (\d{4}-\d{2}-\d{2})', pdf_str, re.I).group(1))

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'PÓŁROCZNA' in pdf_str and 'Przelew' in pdf_str:
            (rata_I := re.search(r'Kwota wpłaty w zł (\d+)', pdf_str, re.I).group(1))
            (rata_II := re.search(r'Kwota wpłaty w zł (.+) (\d*)', pdf_str, re.I).group(2))

            (termin_I := re.search(r'Termin płatności (\d{4}-\d{2}-\d{2})', pdf_str, re.I).group(1))
            (termin_II := re.search(r'Termin płatności (.*)(\d{4}-\d{2}-\d{2})', pdf_str, re.I).group(2))

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

    if 'UNIQA' in page_1:
        box = polisa_box(pdf, 0, 300, 590, 470)
        print(box)
        (total := re.search(r'Składka łączna: (\d*\xa0?\d+)', box, re.I))
        total = int(re.sub(r'\xa0', '', total.group(1)))

        if re.findall(r'(?=.*przelewem)(?=.*jednorazowo).*', box, re.I | re.DOTALL):
            (termin_I := re.search(r'do dnia: (\d{2}.\d{2}.\d{4})', box))
            termin_I = re.sub('[^0-9]', '-', termin_I.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


        if 'przelewem' in box and 'w ratach' in box:
            if 'II.' in box:
                (rata_I := re.search(r'I. (\d+)', box, re.I).group(1))
                (rata_II := re.search(r'II. (\d+)', box, re.I).group(1))

                (termin := re.search(r'/(\d{2}.\d{2}.\d{4})', box))
                termin_I = re.sub('[^0-9]', '-', termin.group(1))
                termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

                (termin_II := re.search(r'II. (.*)/(\d{2}.\d{2}.\d{4})', box))
                termin_II = re.sub('[^0-9]', '-', termin_II.group(2))
                termin_II = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_II)

                return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

    if 'WARTA' in page_1:
        pdf_str = polisa_str(pdf)[1200:5000]
        print(pdf_str)
        (total := re.search(r'SKŁADKA ŁĄCZNA|Kwota\s: (\d*\s?\.?\d+)', pdf_str, re.I))
        total = int(total.group(1).replace('\xa0', '').replace('.', ''))

        if re.findall(r'(?=.*JEDNORAZOWO)(?=.*PRZELEW).*', pdf_str, re.I | re.DOTALL):
            (termin_I := re.search(r'Termin:|DO DNIA.*(\d{4}-\d{2}-\d{2})', pdf_str).group(1))

        return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


"""Koniec arkusza EXCEL"""
def rozpoznanie_danych(tacka_na_polisy):
    pdf = tacka_na_polisy

    page_ = polisa(pdf)
    page_1, page_1_tok = page_[0], page_[1]

    d = dict(enumerate(page_1_tok))
    # print(d)
    p_lub_r = pesel_regon(d)
    # nazwisko, imie = '', '' if regon_checksum(p_lub_r[1:]) else nazwisko_imie(d)

    nazwisko_imie_ = nazwisko_imie(d)
    nazwisko = '' if regon_checksum(p_lub_r[1:]) else nazwisko_imie_[0]
    imie = '' if regon_checksum(p_lub_r[1:]) else nazwisko_imie_[1]
    regon_ = regon(p_lub_r[1:])
    nazwa_firmy, ulica_f, nr_ulicy_f, nr_lok, kod_poczt_f, miasto_f, tel, email = regon_
    ulica_f_edit = f'{ulica_f} {nr_ulicy_f}' if not nr_lok else f'{ulica_f} {nr_ulicy_f} m {nr_lok}'
    kod_poczt_f_edit = f'{kod_poczt_f[:2]}-{kod_poczt_f[2:]}' if '-' not in kod_poczt_f else kod_poczt_f
    kod_poczt = kod_pocztowy(page_1) if kod_pocztowy(page_1) else kod_poczt_f_edit
    data_wyst = data_wystawienia()
    data_konca = koniec_ochrony(page_1)

    numer_polisy_ = numer_polisy(page_1)
    tow_ub_tor = numer_polisy_[0]
    tow_ub = numer_polisy_[1]
    nr_polisy = numer_polisy_[2]


    przypis_daty_raty_ = przypis_daty_raty(pdf, page_1)
    print(przypis_daty_raty_)

    przypis = przypis_daty_raty_[0]
    termin_I = przypis_daty_raty_[1] if przypis_daty_raty_[1] else data_wyst # gotówka
    rata_I = przypis_daty_raty_[2]
    f_platnosci = przypis_daty_raty_[3]
    ilosc_rat = przypis_daty_raty_[4]
    nr_raty = przypis_daty_raty_[5]

    termin_II = przypis_daty_raty_[6]
    rata_II = przypis_daty_raty_[7]
    termin_III = przypis_daty_raty_[8]
    rata_III = przypis_daty_raty_[9]
    termin_IV = przypis_daty_raty_[10]
    rata_IV = przypis_daty_raty_[11]
    # print(przypis(pdf, page_1))

    return nazwa_firmy, nazwisko, imie, p_lub_r, ulica_f_edit, kod_poczt, miasto_f, tel, email, data_wyst, \
            data_konca, tow_ub_tor, tow_ub, nr_polisy, przypis, termin_I, rata_I, f_platnosci, ilosc_rat, nr_raty, \
            termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

def tacka_na_polisy(obj):
    if obj.endswith('.pdf'):
        yield rozpoznanie_danych(obj)
    else:
        for file in os.listdir(obj):
            if file.endswith('.pdf'):
                pdf = obj + '\\' + file
                yield rozpoznanie_danych(pdf)


"""Sprawdza czy arkusz jest otwarty."""
"""Jeżeli arkusz jest zamknięty, otwiera go."""

try:
    ExcelApp = win32com.client.GetActiveObject('Excel.Application')
    wb = ExcelApp.Workbooks("DTESTY.xlsx")
    ws = wb.Worksheets("BAZA 2014")
    # workbook = ExcelApp.Workbooks("Baza.xlsx")

except:
    ExcelApp = Dispatch("Excel.Application")
    wb = ExcelApp.Workbooks.Open(path + "\\DTESTY.xlsx")
    ws = wb.Worksheets("BAZA 2014")


"""Jesienne Bazie"""

for dane_polisy in tacka_na_polisy(obj):
    nazwa_firmy, nazwisko, imie, p_lub_r, ulica_f_edit, kod_poczt, miasto_f, tel, email, data_wyst, data_konca, \
    tow_ub_tor, tow_ub, nr_polisy, przypis, ter_platnosci, rata_I, f_platnosci, ilosc_rat, nr_raty, termin_II, \
    rata_II, termin_III, rata_III, termin_IV, rata_IV = dane_polisy
    print(dane_polisy)

    """Rozpoznaje kolejny wiersz, który może zapisać."""
    row_to_write = wb.Worksheets(1).Cells(wb.Worksheets(1).Rows.Count, 30).End(-4162).Row + 1

    # Rok_przypisu = ExcelApp.Cells(row_to_write, 1).Value = data_wyst[:2] # Komórka tylko do testów
    Rozlicz = ExcelApp.Cells(row_to_write, 7).Value = 'Robert'
    Podpis = ExcelApp.Cells(row_to_write, 10).Value = 'Grzelak'
    FIRMA = ExcelApp.Cells(row_to_write, 11).Value = nazwa_firmy
    Nazwisko = ExcelApp.Cells(row_to_write, 12).Value = nazwisko
    Imie = ExcelApp.Cells(row_to_write, 13).Value = imie
    Pesel_Regon = ExcelApp.Cells(row_to_write, 14).Value = p_lub_r
    # ExcelApp.Cells(row_to_write, 15).Value = data_pr_j
    ExcelApp.Cells(row_to_write, 16).Value = ulica_f_edit # f'{ulica_f} {nr_ulicy_f}' if not nr_lok else f'{ulica_f} {nr_ulicy_f} m {nr_lok}'
    ExcelApp.Cells(row_to_write, 17).Value = kod_poczt # kod_pocztowy(page_1) if not kod_poczt_f else kod_poczt_f_edit
    ExcelApp.Cells(row_to_write, 18).Value = miasto_f
    ExcelApp.Cells(row_to_write, 19).Value = tel
    ExcelApp.Cells(row_to_write, 20).Value = email
    # ExcelApp.Cells(row_to_write, 23).Value = marka
    # ExcelApp.Cells(row_to_write, 24).Value = model
    # ExcelApp.Cells(row_to_write, 25).Value = nr_rej
    # ExcelApp.Cells(row_to_write, 26).Value = rok_prod
    # ExcelApp.Cells(row_to_write, 29).Value = int(ile_dni) + 1
    # ExcelApp.Cells(row_to_write, 30).NumberFormat = 'yy-mm-dd'
    ExcelApp.Cells(row_to_write, 30).Value = data_wyst
    # ExcelApp.Cells(row_to_write, 31).Value = data_pocz
    ExcelApp.Cells(row_to_write, 32).Value = data_konca
    ExcelApp.Cells(row_to_write, 36).Value = 'SPÓŁKA'
    tor = ExcelApp.Cells(row_to_write, 37).Value = tow_ub_tor
    ExcelApp.Cells(row_to_write, 38).Value = tow_ub
    # ExcelApp.Cells(row_to_write, 39).Value = rodzaj
    ExcelApp.Cells(row_to_write, 40).Value = nr_polisy
    # ExcelApp.Cells(row_to_write, 41).Value = nowa_wzn
    # ExcelApp.Cells(row_to_write, 42).Value = nr_wzn
    # if wzn_idx:
    #     ExcelApp.Cells(row_to_write, 41).Value = 'W'
    #     ExcelApp.Cells(row_to_write, 42).Value = nowa_wzn
    # else:
    #     ExcelApp.Cells(row_to_write, 41).Value = 'N'
    #     ExcelApp.Cells(row_to_write, 42).Value = ''
    # ryzyko = ExcelApp.Cells(row_to_write, 46).Value = 'b/d'
    ExcelApp.Cells(row_to_write, 48).Value = przypis
    ExcelApp.Cells(row_to_write, 49).Value = ter_platnosci
    # if I_rata_data:
    #     ExcelApp.Cells(row_to_write, 49).Value = I_rata_data
    if rata_I:
        ExcelApp.Cells(row_to_write, 50).Value = rata_I
    else:
        ExcelApp.Cells(row_to_write, 50).Value = przypis
    ExcelApp.Cells(row_to_write, 51).Value = f_platnosci
    ExcelApp.Cells(row_to_write, 52).Value = ilosc_rat
    ExcelApp.Cells(row_to_write, 53).Value = nr_raty
    data_inkasa = ExcelApp.Cells(row_to_write, 54).Value = ter_platnosci
    if rata_I:
        ExcelApp.Cells(row_to_write, 55).Value = rata_I
    else:
        ExcelApp.Cells(row_to_write, 55).Value = przypis
    ExcelApp.Cells(row_to_write, 60).Value = tow_ub_tor


    if rata_II:
        ws.Range(f'A{row_to_write}:BH{row_to_write}').Copy()
        ws.Range(f'A{row_to_write + 1}').PasteSpecial()

        ExcelApp.Cells(row_to_write + 1, 48).Value = ''
        ExcelApp.Cells(row_to_write + 1, 49).Value = termin_II
        ExcelApp.Cells(row_to_write + 1, 50).Value = rata_II
        ExcelApp.Cells(row_to_write + 1, 53).Value = 2
        data_inkasa = ExcelApp.Cells(row_to_write + 1, 54).Value = ''
        ExcelApp.Cells(row_to_write + 1, 55).Value = ''

        if rata_IV:
            ws.Range(f'A{row_to_write + 1}:BH{row_to_write + 1}').Copy()
            ws.Range(f'A{row_to_write + 2}').PasteSpecial()

            ExcelApp.Cells(row_to_write + 2, 49).Value = termin_III
            ExcelApp.Cells(row_to_write + 2, 50).Value = rata_III
            ExcelApp.Cells(row_to_write + 2, 53).Value = 3

            ws.Range(f'A{row_to_write + 2}:BH{row_to_write + 2}').Copy()
            ws.Range(f'A{row_to_write + 3}').PasteSpecial()

            ExcelApp.Cells(row_to_write + 3, 49).Value = termin_IV
            ExcelApp.Cells(row_to_write + 3, 50).Value = rata_IV
            ExcelApp.Cells(row_to_write + 3, 53).Value = 4



"""Opcje zapisania"""
# ExcelApp.DisplayAlerts = False
# wb.SaveAs(path + "\\DTESTY.xlsx")
# wb.Close()
# ExcelApp.DisplayAlerts = True


end_time = time.time() - start_time
print('Czas wykonania: {:.2f} sekund'.format(end_time))
