import os
import re
import pdfplumber
from datetime import datetime, timedelta
import win32com.client
from win32com.client import Dispatch
import time
from regon_api import get_regon_data

start_time = time.time()
path = os.getcwd()
one_day = timedelta(1)

# obj = input('Podaj polisę/y w formacie .pdf do rejestracji: ')
obj = r'M:\Agent baza\Skrzynka na polisy'


# obj = r'M:\zSkrzynka na polisy'


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
    page_1, page_2, page_3 = '', '', ''
    with pdfplumber.open(pdf) as policy:
        try:
            (page_1 := policy.pages[0].extract_text())
        except:
            pass
        try:
            if policy.pages[1].extract_text():
                page_2 = policy.pages[1].extract_text()
        except:
            pass
        try:
            (page_3 := policy.pages[2].extract_text())
        except:
            pass
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
    return page_1_box  # words_separately(page_1_box.lower())


def pesel_checksum(p):
    """Suma kontrolna nr pesel."""
    l = int(p[10])
    suma = 1 * int(p[0]) + 3 * int(p[1]) + 7 * int(p[2]) + 9 * int(p[3]) + 1 * int(p[4]) + 3 * int(p[5]) + \
           7 * int(p[6]) + 9 * int(p[7]) + 1 * int(p[8]) + 3 * int(p[9]) + 1 * int(p[10])
    lm = suma % 10  # dzielenie wyniku modulo 10
    kontrola = 10 - lm  # sprawdzenie ostatniej liczby kontrolnej
    # w przypadku liczby kontrolnej 10 i 0 sa jednoznaczne a 0 moze byc wynikiem odejmowania
    if (kontrola == 10 or l == kontrola) and p[2:4] != '00':
        return 1
    else:
        return 0


def regon_checksum(r: int):
    """Waliduje regon sprawdzając sumę kontrolną."""
    regon = list(str(r))
    if len(regon) == 9:
        suma = (int(regon[0]) * 8 + int(regon[1]) * 9 + int(regon[2]) * 2 + int(regon[3]) * 3 + int(regon[4]) * 4 +
                int(regon[5]) * 5 + int(regon[6]) * 6 + int(regon[7]) * 7) % 11
        if suma == int(regon[-1]) or suma == 10 and int(regon[-1]) == 0:
            return r
        else:
            return 0
    else:
        return ''


def regon(regon_checksum):
    """API Regon"""
    if len(regon_checksum) == 9:
        try:
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
        except Exception as e:
            print('Wystąpił błąd numeru REGON.')
            return '', '', '', '', '', '', '', ''
    else:
        return '', '', '', '', '', '', '', ''


"""Funkcje odpowiadają kolumnom w bazie."""


def nazwisko_imie(d, page_1, pdf):
    """Zwraca imię i nazwisko Klienta."""
    agent = {'Robert': 'Grzelak Robert', 'Maciej': 'Grzelak Maciej', 'MAGRO': 'Magro Maciej',
             'Wpierdalata': 'Kozłowska-Chyła Beata'}
    with open(path + '\\imiona.txt') as content:
        all_names = content.read().split('\n')
        if 'euroins' in d.values():
            name = []
            for k, v in d.items():
                if v.title() in all_names and not re.search('pesel', d[k + 1], re.I):
                    name.append(f'{d[k + 1].title()} {v.title()}')
                if v.title() in all_names and re.search('\d', d[k + 4]):
                    name.append(f'{d[k + 5].title()} {v.title()}')
                if v.title() in all_names and not re.search('(\d|telefon$|Adres:?)', d[k + 4]):
                    name.append(f'{d[k + 4].title()} {v.title()}')

        elif 'tuz' in d.values():
            name = [f'{d[k + 1].title()} {v.title()}' for k, v in d.items() if k > 10 and v.title() in all_names]
        elif 'warta' in d.values():
            name = [f'{d[k - 1].title()} {v.title()}' for k, v in d.items() if v.title() in all_names]
        elif 'Wiener' in page_1 and 'полис' in page_1:
            name = [f'{name} BRAK!' for name in all_names if re.search(name, page_1, re.I)]
        else:
            name = [f'{d[k + 1].title()} {v.title()}' for k, v in d.items() if v.title() in all_names
                    and f'{d[k + 1].title()} {v.title()}' not in agent.values() and not re.search('\d', d[k + 1])
                    and v.title() not in przedmiot_ub(page_1, pdf)]

    if name:
        last_name = name[0].split()[0]
        first_name = name[0].split()[1]
        return name[0].split()[0], name[0].split()[1]
    else:
        return '', ''


def pesel_regon(d, page_1):
    """Zapisuje pesel/regon."""
    nr_reg_TU = {'AXA': '140806789', 'PKO BP': '016298263', 'Marian Toś': '602776660'}
    pesel = [pesel for k, pesel in d.items() if k < 200 and len(pesel) == 11 and re.search('(?<!\+)\d{11}', pesel)
             and pesel_checksum(pesel)]
    regon = [regon for k, regon in d.items() if k < 200 and len(regon) == 9 and re.search('\d{9}', regon) and regon
             not in nr_reg_TU.values() and regon_checksum(regon) and not 'INTER PARTNER' in page_1]
    if pesel:
        return 'p' + pesel[0]
    elif regon:
        return 'r' + regon[0]
    else:
        return ''


def prawo_jazdy(page_1, pdf):
    data_pr_j = ''
    if 'Allianz' in page_1 and (data_pr_j := re.search('(prawa jazdy:) (\d{4})', page_1, re.I)):
        return data_pr_j.group(2)
    if 'UNIQA' in page_1 and (data_pr_j := re.search('(Prawo jazdy od) (\d{4})', page_1, re.I)):
        return data_pr_j.group(2)
    if 'Generali' in page_1 and (data_pr_j := re.search('(rok uzyskania prawa jazdy:) (\d{4})', page_1, re.I)):
        return data_pr_j.group(2)
    if 'Hestia' in page_1 and (data_pr_j := re.search('(data uzyskania prawa jazdy) (\d{4})', page_1, re.I)):
        return data_pr_j.group(2)
    # Link4
    if re.search('Numer\s(\w\d+)', page_1) and (data_pr_j := re.search('uzyskania prawa jazdy (\d{4})', page_1, re.I)):
        return data_pr_j.group(1)


def adres():
    """Tylko w przypadku regon (API)."""
    pass


def kod_pocztowy(page_1, pdf):
    # print(page_1)
    kod = re.compile('(adres\w*|(?<!InterRisk) kontakt\w*|pocztowy|ubezpiecz(ony|ający))', re.I)
    if 'UNIQA' in page_1 or 'TUW' in page_1 and not 'TUZ' in page_1:
        page_1 = polisa_str(pdf)[0:-200]

    if (wiener := re.search('wiener', page_1, re.I)):
        kod_pocztowy = re.search('(Adres.*|Siedziba.*|ŁÓDŹ)\s?(\d{2}[-\xad]\d{3})', page_1)
        return kod_pocztowy.group(2)

    if (f := kod.search(page_1)):
        adres = f.group().strip()
        data = page_1.split()
        dystans = [data[data.index(split) - 10: data.index(split) + 33] for split in data if adres in split][0]

        if kod_pocztowy := [kod for kod in dystans if re.search('^\d{2}[-\xad]\d{3}$', kod)]:
            return kod_pocztowy[0]
    return ''


def tel_mail(page_1, pdf, d, nazwisko):
    regon = pesel_regon(d, page_1)[1:]
    tel, mail = '', ''
    tel_mail_off = {'tel Robert': '606271169', 'mail Robert': 'ubezpieczenia.magro@gmail.com',
                    'tel Maciej': '602752893', 'mail Maciej': 'magro@ubezpieczenia-magro.pl',
                    'tel MAGRO': '572810576', 'mail AXA': 'obsluga@axaubezpieczenia.pl',
                    'mail UNIQA': 'centrala@uniqa.pl', 'p_lub_r': regon}

    if 'Allianz' in page_1:
        try:
            tel = ''.join([tel for tel in re.findall(r'tel.*([0-9 .\-\(\)]{8,}[0-9])', page_1) if
                           tel not in tel_mail_off.values()][0])
        except:
            pass
        try:
            mail = ''.join([mail for mail in re.findall(r'([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)', page_1) if
                            mail not in tel_mail_off.values()][0])
        except:
            pass
        return tel, mail

    elif 'EUROINS' in page_1:
        if tel := re.search(r'telefon: (\+48|0048)?\s?([0-9.\-\(\)\s]{9,})?', page_1):
            tel = tel.group(2)
        if mail := re.search(r'email: ([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', page_1):
            mail = mail.group(1)
        return tel, mail

    elif 'Compensa' in page_1:
        try:
            tel = re.search(r'kom:\s(\+48|0048)?\s?([0-9.\-\(\)\s]{9,})?', page_1).group(2)
        except:
            pass
        try:
            mail = re.search(r'mail:\s([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', page_1).group(1)
        except:
            pass
        return tel, mail

    elif 'Generali' in page_1 and not 'Proama' in page_1:
        try:
            tel = re.search(r'telefon: (\+48|0048)?\s?([0-9.\-\(\)\s]{9,})?', page_1).group(2)
        except:
            pass
        try:
            mail = re.search(r'email: ([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', page_1).group(1)
        except:
            pass
        return tel, mail

    elif 'Hestia' in page_1:
        try:
            if (t := re.search('TEL. ([0-9.\-\(\)\s]{9})', page_1)):
                tel = t.group(1)
        except:
            pass
        try:
            if (m := re.search(r' ([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)', page_1)):
                mail = m.group(1).lower()
        except:
            pass
        return tel, mail

    elif 'InterRisk' in page_1:
        try:
            tel = re.search(r'Telefon: (\+48|0048)?\s?([0-9.\-\(\)\s]{9,}\n)', page_1).group(2)
        except:
            pass
        try:
            mail = re.search(r'Email: ([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', page_1).group(1)
        except:
            pass
        return tel.rstrip('.\xa0'), mail

    elif 'INTER' in page_1:
        try:
            tel = re.search(r'Telefon kontaktowy: (\+48|0048)?\s?([0-9.\-\(\)\s]{9})?', page_1).group(2)
        except:
            pass
        try:
            mail = re.search(r'Adres e-mail: ([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', page_1).group(1)
        except:
            pass
        return tel, mail

    elif 'Proama' in page_1:
        try:
            tel = re.search(r'telefon:\s*(\+48|0048)\s?([0-9.\-\(\)\s]{9,})', page_1).group(2)
        except:
            pass
        try:
            mail = re.search(r'email:\s*([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)', page_1).group(1)
        except:
            pass
        return tel, mail

    elif 'PZU' in page_1:
        try:
            tel = re.search(r'Telefon:\s*(\+48|0048)?\s?([0-9.\-\(\)\s]{9})?', page_1, re.I).group(2)
        except:
            pass
        try:
            mail = re.search(r'E\s?-\s?mail:\s*([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', page_1, re.I).group(1)
        except:
            pass
        return tel, mail

    elif 'TUW' in page_1 and not 'TUZ' in page_1:
        pdf_str3 = polisa_str(pdf)[0:-500]
        try:
            tel = re.search(r'(Tel:) (\+48|0048)?\s?([0-9.\-\(\)]{9,})?', pdf_str3).group(3)
        except:
            pass
        try:
            mail = re.search(r'(mail:|e-mail:)\s?([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', pdf_str3).group(2)
            if not mail:
                mail = re.search(r'(Właściciel)\s?([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', pdf_str3).group(2)
        except:
            pass
        return tel, mail

    elif 'TUZ' in page_1:
        try:
            tel = re.search(rf'(\d+)(\n?Imię i nazwisko)', page_1, re.I).group(1)
            tel.replace('\n', '')
        except:
            pass
        try:
            mail = re.search(r'(email:).*([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', page_1, re.DOTALL).group(2)
        except:
            pass
        return tel, mail

    elif 'UNIQA' in page_1:
        pdf_str2 = polisa_str(pdf)
        try:
            mail = ''.join([mail for mail in re.findall(r'([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)', page_1) if
                            mail.casefold() not in tel_mail_off.values()][0])
        except:
            pass
        try:
            # tel = re.search(r'(komórkowy:)?\n?(?<![-\w])([0-9]{9})\s', pdf_str2, re.M).group(2)
            tel = ''.join([tel for tel in re.findall(r'\s([0-9]{9}),?\s', page_1) if tel not in
                           tel_mail_off.values()][0]) and not re.search(r'REGON:\s([0-9]{9}),?\s', page_1)
        except:
            pass
        return tel, mail

    elif 'WARTA' in page_1:
        pdf_str3 = polisa_str(pdf)[0:-3000]
        try:
            tel = re.search(r'Telefon komórkowy:\s?(\+48|0048)?\s?([0-9.\-\(\)]{9,})?', pdf_str3).group(2)
        except Exception:
            pass
        try:
            mail = re.search(r'E-?mail:\s?([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)?', pdf_str3, re.I).group(1)
        except Exception:
            pass

        return tel, mail

    elif 'Wiener' in page_1:
        try:
            tel = re.search(r'(Telefon kontaktowy:?|Telefon komórkowy:)\s?(\+48|0048)?\s?([0-9.\-\(\)]{9,})?',
                            page_1).group(3)
            if '-' in tel:
                tel = tel.replace('-', '')
        except:
            pass
        try:
            mail = re.search(r'(E-mail:?)\s?([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)', page_1, re.I).group(2)
        except:
            pass
        return tel, mail

    else:
        tel_comp, email_comp = re.compile(r'\s([0-9]{9}),?\n?'), re.compile(r'([A-z0-9._+-]+@[A-z0-9-]+\.[A-z0-9.-]+)')
        tel_list = re.findall(tel_comp, page_1)
        email_list = re.findall(email_comp, page_1)
        try:
            tel = [phony for phony in tel_list if not regon_checksum(phony) and phony not in tel_mail_off.values()][0]
        except:
            pass
        try:
            mail = [email.casefold() for email in email_list if email.casefold() not in tel_mail_off.values()][0]

        except:
            pass

        return tel, mail

    return tel, mail


def przedmiot_ub(page_1, pdf):
    marka, kod, model, miasto, nr_rej, adres, rok = '', '', '', '', '', '', ''
    with open(path + '\\marki.txt') as content:
        makes = content.read().split('\n')
        if 'Allianz' in page_1:
            if 'Marka / model pojazdu' in page_1:
                marka = re.search('(Marka / model pojazdu) ([\w-]+)', page_1, re.I).group(2)
                model = re.search('(Marka / model pojazdu) ([\w-]+) (\w+)', page_1).group(3)
                nr_rej = re.search('(NR REJESTRACYJNY) ([\w\d.]+),?', page_1).group(2)
                rok = re.search('(Rok produkcji) (\d+),?', page_1).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok

            elif 'MÓJ DOM' in page_1:
                pdf_str2 = polisa_box(pdf, 0, 100, 320, 400)
                kod = re.search('(Miejsce ubezpieczenia).*\n?.*,\s?(\d{2}-\d{3})', page_1).group(2)
                miasto = re.search(f'{kod}.*\n?(\w*)', page_1).group(1)
                adres = re.search('(Miejsce ubezpieczenia) (ul.) ([\w\d/]+)', page_1).group(3)
                if re.search('(Rok budowy) (\d+)', page_1):
                    rok = re.search('(Rok budowy) (\d+)', page_1).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'UNIQA' in page_1:
            pdf_str2 = polisa_str(pdf)[0:-300]
            if re.search('Pojazd', pdf_str2, re.I) and not 'Firma & Planowanie' in pdf_str2:
                for make in makes:
                    for line_num, line in enumerate(pdf_str2.split('\n')):
                        if make in (lsplt := line.split()):
                            marka = make
                            model = lsplt[lsplt.index(marka) + 1] if lsplt[lsplt.index(marka) + 1] not in \
                                                                     ('Model:', '-') else lsplt[lsplt.index(marka) + 2]
                        if re.search(f'(Twój zakres ubezpieczenia\n?[\w\d]+)', pdf_str2, re.I):
                            nr_rej = re.search(f'Twój zakres ubezpieczenia\n?([\w\d]+)', pdf_str2).group(1)
                        elif re.search('er rejestracyjny: ([\w\d]+)', pdf_str2):
                            nr_rej = re.search('er rejestracyjny: ([\w\d]+)', pdf_str2).group(1)
                        elif re.search(f'(Twój zakres ubezpieczenia\n?[\w\d]+)', pdf_str2, re.I):
                            nr_rej = re.search(f'([\w\d]+) {marka}', pdf_str2).group(1)
                        else:
                            pass
                        if line.startswith('Rok produkcji'):
                            rok = re.search('\d{4}', line).group()
                        elif re.search('Rok produkcji: (\d{4})', pdf_str2):
                            rok = re.search('Rok produkcji: (\d{4})', pdf_str2).group(1)
                        elif re.search('\d{4}-\d{2}-\d{2} (\d{4})', pdf_str2):
                            rok = re.search('\d{4}-\d{2}-\d{2} (\d{4})', pdf_str2).group(1)
                        else:
                            pass
                return marka, kod, model, miasto, nr_rej, adres, rok

            elif 'Przedmiot ubezpieczenia: Mieszkanie' in pdf_str2:
                kod = re.search('(Adres:).*\s(\d{2}[-\xad]\d{3})\s', pdf_str2, re.I).group(2)
                miasto = re.search(f'{kod}\s(\w+)', pdf_str2).group(1)
                adres = re.search('(Adres:)\s(.*),', pdf_str2).group(2)
                rok = re.search('(Rok budowy:)\s(\d{4})', pdf_str2).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok

            elif 'Adres miejsca ubezpieczenia ' in pdf_str2:

                kod = re.search('(Adres miejsca ubezpieczenia).*\s(\d{2}[-\xad]\d{3})\s', pdf_str2, re.I).group(2)
                miasto = re.search(f'{kod}\s(\w+)', pdf_str2).group(1)
                adres = re.search('miejsca ubezpieczenia\s([A-z0-9\s\./]+),', pdf_str2).group(1)
                # rok = re.search('(Rok\sbudowy:)\s(\d{4})', pdf_str2).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'Compensa' in page_1:
            if 'Dane pojazdu' in page_1:
                for make in makes:
                    if make in page_1:
                        marka = make
                        model = re.search(rf'{marka}\s(\w+)', page_1).group(1)
                nr_rej = re.search('(Numer rejestracyjny:) ([\w\d.]+),?', page_1).group(2)
                rok = re.search('(produkcji:).*\n?([0-9]{4}),', page_1, re.M).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'EUROINS' in page_1:
            if 'Dane pojazdu' in page_1:
                marka = re.search('(Marka, model:) (\w+)', page_1, re.I).group(2)
                model = re.search(f'{marka}\s+([\w\d./]+)', page_1).group(1)
                nr_rej = re.search('(rejestracyjny:) ([\w\d.]+)', page_1).group(2)
                rok = re.search('(Rok[\s\n]?produkcji:) (\d+),?', page_1).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'Generali' in page_1 or 'Proama' in page_1:
            if 'DANE POJAZDU' in page_1:
                marka = re.search('(Marka / Model) (\w+)', page_1, re.I).group(2)
                model = re.search('(Marka / Model) (\w+-?\w+) /? ([\w\d./]+)', page_1).group(3)
                nr_rej = re.search('(Numer rejestracyjny / VIN) ([\w\d.]+)', page_1).group(2)
                rok = re.search('(Rok produkcji) (\d+),?', page_1).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok

            elif 'UBEZPIECZENIE MIESZKANIA' in page_1:
                kod = re.search('(Miejsce ubezpieczenia).*\n?(\d{2}[-\xad]\d{3})', page_1, re.I).group(2)
                miasto = re.search(f'{kod} (\w+)', page_1).group(1)
                adres = re.search('(Miejsce ubezpieczenia) ([\w \d/.]+),', page_1).group(2)
                rok = re.search('(Rok budowy) (\d+)', page_1).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'Hestia' in page_1 and not 'MTU' in page_1:
            if 'Ubezpieczony pojazd' in page_1:
                try:
                    marka = re.search(r'(Ubezpieczony pojazd).*?(\w+), (\w+-?\w+)', page_1, re.I | re.DOTALL).group(3)
                    model = re.search(rf'(?<={marka}),? (\w+)', page_1, re.I).group(1)
                    nr_rej = re.search(rf'([A-Z0-9]+)\s?(?=, ROK)', page_1).group(1)
                    rok = re.search(r'(ROK PRODUKCJI:?) (\d{4})', page_1).group(2)
                    return marka, kod, model, miasto, nr_rej, adres, rok
                except:
                    pass


        elif 'InterRisk' in page_1:
            if 'DANE POJAZDU' in page_1:
                marka = re.search(r'Marka/typ/model: ([\w./-]+)', page_1, re.I).group(1)
                model = re.search(rf'(?<={marka})\s(\w+)', page_1, re.I).group(1)
                nr_rej = re.search(r'Nr rejestracyjny: ([A-Z0-9]+)', page_1).group(1)
                rok = re.search(r'Rok produkcji: (\d{4})', page_1).group(1)

                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'HDI' in page_1 and not 'PZU' in page_1 or '„WARTA” S.A. POTWIERDZA' in page_1:

            if re.search(r'DANE UBEZPIECZONEGO POJAZDU|Marka|rejestracyjny', page_1, re.I):
                for make in makes:
                    for line in page_1.split('\n'):
                        if make in (lsplt := line.split()):
                            marka = make
                            model = lsplt[lsplt.index(marka) + 1] if lsplt[lsplt.index(marka) + 1] not in \
                                                                     ('Model:', '-') else lsplt[lsplt.index(marka) + 2]
                nr_rej = re.search(r'(Nr|Numer) rejestracyjny: ([A-Z0-9]+)', page_1).group(2)
                rok = re.search(r'Rok produkcji: (\d{4})', page_1).group(1)

                return marka, kod, model, miasto, nr_rej, adres, rok


        # Link4
        elif re.search('Numer:?\s\n?(\w\d+)', page_1, re.I) and not 'WARTA' in page_1 or 'LINK4' in page_1:
            if 'Marka / Model' in page_1 or 'DANE POJAZDU' in page_1:
                marka = re.search(r'(Marka / Model|Marka) ([\w./]+)', page_1, re.I).group(2)
                model = re.search(rf'({marka})\n?\s?[A-z]+\s?(\w+)', page_1, re.I).group(2)
                nr_rej = re.search(r'rejestracyjny ([A-Z0-9]+)', page_1).group(1)
                rok = re.search(r'Rok produkcji (\d{4})', page_1).group(1)
                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'MTU' in page_1:
            if 'Ubezpieczony pojazd' in page_1:
                marka = re.search(r'(Ubezpieczony pojazd).*?(\w+), (\w+-?\w+)', page_1, re.I | re.DOTALL).group(3)
                model = re.search(rf'(?<={marka}) (\w+)', page_1, re.I).group(1)
                nr_rej = re.search(rf'([A-Z0-9]+)(?=, ROK)', page_1).group(1)
                rok = re.search(r'(ROK PRODUKCJI:) (\d{4})', page_1).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'PZU' in page_1:
            if 'Ubezpieczony pojazd' in page_1:
                marka = re.search(r'Marka:\s([\w./-]+)', page_1, re.I).group(1)
                model = re.search(r'(typ pojazdu:|Model:|Typ:)\s((?!Model:)(?!Typ:)\w+)', page_1, re.I).group(2)
                if nr_rej := re.search(r'nr\srejestracyjny:?\s([A-Z0-9]+)', page_1, re.I):
                    nr_rej = nr_rej.group(1)
                rok = re.search(r'Rok\sprodukcji:\s(\d{4})', page_1, re.I).group(1)
                return marka, kod, model, miasto, nr_rej, adres, rok

            if 'Miejsce ubezpieczenia:' in page_1:
                kod = re.search('(Miejsce ubezpieczenia:).*(\d{2}[-\xad]\d{3})', page_1, re.I).group(2)
                miasto = re.search(f'{kod} (\w+)', page_1).group(1)
                adres = re.search('(Miejsce ubezpieczenia:)\s([\w\s\-\d/\.]+),', page_1).group(2)
                rok = re.search('(rok budowy:)? (\d+)?', page_1).group(1)
                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'TUW' in page_1 and not 'TUZ' in page_1:
            pdf_str3 = polisa_str(pdf)[0:6500]
            if 'pojazdu:' in pdf_str3:
                nr_rej = re.search(r'numer\s*rejestracyjny:\s*([A-Z0-9]+)', pdf_str3).group(1)
                try:
                    marka_model = re.search(rf'({nr_rej})\n?.*?((?<=marka/model/ typ: )?.*?([\w-/]*))',
                                            pdf_str3, re.I | re.DOTALL).group(2).split('/')
                    marka, model = marka_model[0], marka_model[1]
                except Exception:
                    splt = re.split(' |/|\n', pdf_str3)
                    for make in makes:
                        if make in splt:
                            marka, model = make, splt[splt.index(make) + 1]
                rok = re.search(r'rok produkcji:\s?(\d{4})', pdf_str3).group(1)
                return marka, kod, model, miasto, nr_rej, adres, rok

            if 'Miejsce ubezpieczenia:' in page_1:
                kod = re.search('(Miejsce ubezpieczenia:).*(\d{2}[-\xad]\d{3})', page_1, re.I).group(2)
                miasto = re.search(f'{kod} (\w+)', page_1).group(1)
                adres = re.search('(Miejsce ubezpieczenia:) ([\w \d/]+),', page_1).group(2)
                # rok = re.search('(Rok budowy) (\d+)', page_1).group(2)
                return marka, kod, model, miasto, nr_rej, adres, rok


        elif 'TUZ' in page_1:
            pdf_str3 = polisa_str(pdf)[0:6500]
            print(pdf_str3)
            if 'Dane pojazdu' in pdf_str3:
                try:
                    nr_rej = re.search(r'Dane pojazdu\n?\n?([A-Z0-9]+)', pdf_str3, re.DOTALL).group(1)
                    # marka = re.search(rf'{nr_rej}.*?([\w./]+)', page_1, re.I).group(1)
                    for make in makes:
                        for line in pdf_str3.split('\n'):
                            if make in line.split() or make[:8] in line.split():
                                marka = make if make else make[:8]
                                try:
                                    model = re.search(rf'{marka}\s([\w./\d]+)', page_1, re.I | re.DOTALL).group(1)
                                except Exception:
                                    print(f'TUZ krótka "wersja" marki {Exception}')
                                    model = re.search(rf'{make[:8]}.*([\w./\d]+)', page_1, re.I | re.DOTALL).group(1)
                    rok = re.search(r'SAMOCHÓD (\w+) (\d{4})', page_1).group(2)
                except Exception:
                    print(f'TUZ błąd w danych pojazdu {Exception}')
                return marka, kod, model, miasto, nr_rej, adres, rok

            """Uniqa sie z AXA"""
        elif 'UNIQA' in page_1:
            if 'POJAZD' in page_1:
                # print(page_1)
                marka = re.search(r'(Marka i model:|Pojazd Symbol) ([\w./-]+)', page_1, re.I).group(2)
                model = re.search(rf'(Marka i model:)?\s?{marka}\s?([\w./\d]+)', page_1, re.I).group(2)
                nr_rej = re.search(r'(Numer rejestracyjny:|Liczba miejsc)\n?\s?([A-Z0-9]+)', page_1).group(2)
                try:
                    rok = re.search(rf'(Rok produkcji:|{model})' + '\D(\d{4})\D', page_1).group(2)
                except:
                    pass
                return marka, kod, model, miasto, nr_rej, adres, rok

        elif 'WARTA' in page_1:
            if re.search('Nr rejestracyjny', page_1) or re.search('VIN', page_1):
                for make in makes:
                    for line in page_1.split('\n'):
                        if make in (lsplt := line.split()):
                            marka = make
                            model = lsplt[lsplt.index(marka) + 1] if lsplt[lsplt.index(marka) + 1] not in \
                                                                     ('Model:', '-') else lsplt[
                                lsplt.index(marka) + 2]
                nr_rej = re.search(r'(Nr|Numer) rejestracyjny: ([A-Z0-9]+)', page_1).group(2)
                rok = re.search(r'Rok produkcji: (\d{4})', page_1).group(1)

                return marka, kod, model, miasto, nr_rej, adres, rok

            if 'WARTA DOM' in page_1:
                kod = re.search('(?s:.*ADRES MIEJSCA UBEZPIECZENIA:?)\n(\d{2}[-\xad]\d{3})\s', page_1, re.I).group(
                    1)  # zamieszkania/korespondencyjny:
                miasto = re.search(f'(?s:.*){kod}\s(\w+)', page_1).group(1)
                adres = re.search(rf'(?s:.*){miasto},\n?(.*)\n', page_1).group(1).lstrip()
                return marka, kod, model, miasto, nr_rej, adres, rok

        elif 'Wiener' in page_1:
            if 'DANE POJAZDU' in page_1:
                marka = re.search(r'Marka pojazdu: ([\w./]+)', page_1, re.I).group(1)
                model = re.search(rf'Model pojazdu:\s*([\w-]+)', page_1, re.I).group(1)
                nr_rej = re.search(r'Numer rejestracyjny: ([A-Z0-9]+)', page_1).group(1)
                rok = re.search(r'Rok produkcji: (\d{4})', page_1).group(1)
                return marka, kod, model, miasto, nr_rej, adres, rok

        return marka, kod, model, miasto, nr_rej, adres, rok


def data_wystawienia():
    today = datetime.strptime(datetime.now().strftime('%y-%m-%d'), '%y-%m-%d') + one_day
    return today  # datetime.today().date().strftime('%y-%m-%d')


def koniec_ochrony(page_1, pdf):
    daty = re.compile(r'(?<!\w)(\d{2}[-\./]\d{2}[-\./]\d{4}|t?o?\d{4}[-\./]\d{2}[-\./]\d{2})')
    if 'UNIQA' in page_1 or 'TUW' in page_1 and not 'TUZ' in page_1:
        page_1 = polisa_str(pdf)[0:-1]
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


def numer_polisy(page_1, pdf):
    # print(page_1)
    nr_polisy = ''
    if 'Allianz' in page_1 and (nr_polisy := re.search(r'(Polisa nr|NUMER POLISY) (\d*-?\d+)', page_1)) or \
            'Globtroter' in page_1 and nr_polisy:
        return 'ALL', 'ALL', nr_polisy.group(2)
    elif 'Compensa' in page_1 and (nr_polisy := re.search('typ polisy: *\s*(\d+),numer: *\s*(\d+)', page_1)):
        return 'COM', 'COM', nr_polisy.group(1) + nr_polisy.group(2)
    elif 'EUROINS' in page_1 and (nr_polisy := re.search('Polisa ubezpieczenia nr: (\d+)', page_1)):
        return 'EIN', 'EIN', nr_polisy.group(1)
    elif 'Generali' in page_1 and not 'Proama' in page_1 and (
    nr_polisy := re.search('POLISA NR\s*(\d+)', page_1, re.I)):
        return 'GEN', 'GEN', nr_polisy.group(1)
    elif 'HDI' in page_1 and (nr_polisy := re.search('POLISA NR\s?: *(\d+)', page_1)):
        return 'WAR', 'HDI', nr_polisy.group(1)
    elif 'Hestia' in page_1 and not 'MTU' in page_1 and (nr_polisy := re.search('Polisa[^0-689]+(\d+)', page_1, re.I)):
        return 'HES', 'HES', nr_polisy.group(1)
    elif 'INTER' and (nr_polisy := re.search('polisa\s*seria\s*(\w*)\s*numer\s*(\d*)', page_1)):
        return 'INT', 'INT', nr_polisy.group(1) + nr_polisy.group(2)
    elif 'InterRisk' in page_1 and (nr_polisy := re.search('Polisa seria?\s(.*)\snumer\s(\d+)', page_1, re.I)):
        return 'RIS', 'RIS', nr_polisy.group(1) + nr_polisy.group(2)

    elif 'MTU' in page_1 and (nr_polisy := re.search('Polisa\s.*\s(\d+)', page_1, re.I)):
        return 'AZ', 'MTU', nr_polisy.group(1)
    elif 'Proama' in page_1 and (nr_polisy := re.search('POLISA NR\s*(\d+)', page_1, re.I)):
        return 'GEN', 'PRO', nr_polisy.group(1)
    elif 'PZU' in page_1 and (nr_polisy := re.search('(Nr|nr polisy:?)\s(\d+)', page_1, re.I)):
        return 'PZU', 'PZU', nr_polisy.group(2)
    elif 'TUW' in page_1 and not 'TUZ' in page_1:
        page_str3 = polisa_str(pdf)[0:-600]
        if (nr_polisy := re.search('Wniosko-Polisa\snr:?\s?(\d+)', page_str3, re.I)):
            return 'TUW', 'TUW', nr_polisy.group(1)
    elif 'TUZ' in page_1 and (nr_polisy := re.search('WNIOSEK seria (\w+) nr (\d+)', page_1)):
        return 'TUZ', 'TUZ', nr_polisy.group(1) + nr_polisy.group(2)
    elif 'UNIQA' in page_1:
        page_1 = polisa_str(pdf)[0:4600]
        if (nr_polisy := re.search('(Numer\spolisy:)?\s(\d{4}-\d{7,}|\d{8,14})', page_1)):
            return 'UNI', 'UNI', nr_polisy.group(2)
    # if 'UNIQA' in page_1 and (nr_polisy := re.search('Nr (\d{6,})', page_1)):  # było skomentowane po AXA
    #     return 'UNI', 'UNI', nr_polisy.group(1)
    elif 'WARTA' in page_1 and (nr_polisy := re.search(
            '(POLISA NR\s?:|WARTA DOM\s\w*\s?NR:|PLUS NR:|TRAVEL NR:)\s*(\d+)', page_1)):
        return 'WAR', 'WAR', nr_polisy.group(2)
    elif 'Wiener' in page_1 and (nr_polisy := re.search('(Seria\s?i\s?numer|полис)\s*(\w+\d+)', page_1)):
        return 'WIE', 'WIE', nr_polisy.group(2)

    # Link4 powinien być na końcu - brak "Link4" na polisie.
    elif (nr_polisy := re.search('(Numer:?|POLISA DLA PANA:?|NR)\s?\n?(\w\d+)', page_1, re.I)) and \
            not 'Travel' in page_1 and not 'WARTA' in page_1 and not 'PZU' in page_1 and not 'Wiener' in page_1:
        return 'ULIN', 'LIN', nr_polisy.group(2)
    else:
        return 'NIE ROZPOZNANE !', '', ''


def zam_spacji(kwota):
    replacements = [(' ', ''), (',', '.'), ('\xa0', '')]
    for pattern, replacement in replacements:
        kwota = kwota.replace(pattern, replacement)
    return float(kwota)


def zamiana_sep(termin):
    return re.sub('[^0-9]', '-', termin)


def term_pln(termin, group_n):
    if termin:
        zamiana_sep = re.sub('[^0-9]', '-', termin.group(group_n))
        return re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', zamiana_sep)


def przypis_daty_raty(pdf, page_1):
    total, termin_I, rata_I, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV = \
        '', '', '', '', '', '', '', '', ''

    if 'Allianz' in page_1 or 'Globtroter' in page_1:
        # box = polisa_box(pdf, 0, 320, 590, 780)
        pdf_str2 = polisa_str(pdf)[0:-2600]
        total = re.search(r'(Składka:|łącznie:|za 3 lata:|za rok:|Razem)\s(\d*\s?\d+)', pdf_str2)
        if total:
            total = int(re.sub(r'\xa0', '', total.group(2)))

        if re.findall(r'(?=.*gotówka)(?!.*rat[ay]).*', pdf_str2, re.I | re.DOTALL) or \
                re.findall(r'(?=.*opłacono gotówką)(?=.*I rata).*', pdf_str2, re.I | re.DOTALL):
            termin_I = re.search(r'Dane płatności:\ndo (\d{2}.\d{2}.\d{4})', pdf_str2, re.I)
            termin_I = re.sub('[^0-9]', '-', termin_I.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'I rata przelew' in page_1:

            termin_I = re.search(r'Dane płatności:\ndo (\d{2}.\d{2}.\d{4})', pdf_str2, re.I)
            termin_I = re.sub('[^0-9]', '-', termin_I.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


        elif 'przelew' in page_1 and 'II rata' in page_1 and not 'III rata' in page_1 or \
                (raty_bezrat := re.search('do (\d{2}.\d{2}.\d{4}).*\ndo (\d{2}.\d{2}.\d{4})', page_1)):

            termin = re.search(r'do (\d{2}.\d{2}.\d{4}).*\ndo\s?(\d{2}.\d{2}.\d{4}).*', pdf_str2, re.I | re.DOTALL)

            termin_I = term_pln(termin, 1)
            termin_II = term_pln(termin, 2)
            rata_I = re.search(r'do (\d{2}.\d{2}.\d{4}) r. (\d*\s?\d+)', pdf_str2, re.I).group(2)
            rata_II = re.search(rf'{raty_bezrat.group(2)}.*r\.\s(\d*\s?\d+)', pdf_str2, re.I).group(1)

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


        elif 'Twoja składka za 3 lata' in page_1:
            termin = re.search('do (\d{2}.\d{2}.\d{4}).*\n?do (\d{2}.\d{2}.\d{4}).*\n?do (\d{2}.\d{2}.\d{4})', pdf_str2,
                               re.I)

            termin_I = term_pln(termin, 1)
            termin_II = term_pln(termin, 2)
            termin_III = term_pln(termin, 3)

            rata_I = re.search(f'{termin.group(1)} r. (\d*\s?\d+)', pdf_str2, re.I).group(1)
            rata_II = re.search(f'{termin.group(2)} r. (\d*\s?\d+)', pdf_str2, re.I).group(1)
            rata_III = re.search(f'{termin.group(3)} r. (\d*\s?\d+)', pdf_str2, re.I).group(1)

            return total, termin_I, rata_I, 'P', 3, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        return total, termin_I, rata_I, '', '', '', termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'Compensa' in page_1:

        box = polisa_box(pdf, 0, 260, 590, 750)
        try:
            total = re.search(r'Składka ogółem: (\d*\s?\d+)', box, re.I)
            total = int(re.sub(r'\xa0', '', total.group(1)))
        except:
            pass

        if re.findall(r'(?=.*przelew)(?=.*jednorazowa).*', box, re.I | re.DOTALL):
            termin_I = term_pln(re.search(r'I\srata\s-\s+(\d{2}.\d{2}.\d{4})', box, re.I), 1)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*przelew)(?=.*półroczna).*', box, re.I | re.DOTALL):
            rata_I = re.search(r'I rata -  \d{2}.\d{2}.\d{4} - (\d*\s?\d+)', box, re.I).group(1)
            rata_II = re.search(r'II rata.* - (\d*\s*\d+)', box, re.I).group(1)

            termin_I = term_pln(re.search(r'I\srata\s-\s+(\d{2}.\d{2}.\d{4})', box, re.I), 1)
            termin_II = term_pln(re.search(r'II\srata\s-\s+(\d{2}.\d{2}.\d{4})', box, re.I), 1)
            return total, termin_I, zam_spacji(rata_I), 'P', 2, 1, termin_II, zam_spacji(rata_II), \
                   termin_III, rata_III, termin_IV, rata_IV

        elif 'składki: kwartalna' in box:
            (rata_I := re.search(r'I rata -  \d{2}.\d{2}.\d{4} - (\d+)', box, re.I).group(1))
            (rata_II := re.search(r'II rata.* - (\d+)', box, re.I).group(1))
            (rata_III := re.search(r'[|III] rata.* - (\d+)', box, re.I).group(1))
            (rata_IV := re.search(r'\n- (\d+)', box, re.I).group(1))

            termin_I = term_pln(re.search(r'I\srata\s-\s+(\d{2}.\d{2}.\d{4})', box, re.I), 1)
            termin_II = term_pln(re.search(r'II rata -  (\d{2}.\d{2}.\d{4})', box, re.I), 1)
            termin_III = term_pln(re.search(r',   rata -  (\d{2}.\d{2}.\d{4})', box, re.I), 1)
            termin_IV = term_pln(re.search(r'IV rata -  (\d{2}.\d{2}.\d{4})', box, re.I), 1)
            return total, termin_I, rata_I, 'P', 4, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'EUROINS' in page_1:
        box = polisa_box(pdf, 0, 400, 590, 750)
        total = re.search(r'Łączna składka do zapłaty ([\d ]+,\d*)', box, re.I)
        total = float(total.group(1).replace(r',', '.').replace(' ', ''))

        if re.search('I Rata', box, re.I) and not re.search('II Rata', box, re.I) and re.search('przelew', box, re.I):
            termin_I = re.search(r'1. (\d{4}-\d{2}-\d{2})', box, re.I).group(1)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if re.search('II Rata', box, re.I) and re.search('przelew', box, re.I):
            termin_I = re.search(r'\n1. (\d{4}-\d{2}-\d{2})', box).group(1)
            termin_II = re.search(r'\n2. (\d{4}-\d{2}-\d{2})', box).group(1)

            rata_I = re.search(rf'{termin_I}\s' + r'(\d*\s?\d+.\d*)', box).group(1)
            rata_II = re.search(rf'{termin_II}\s' + r'(\d*\s?\d+.\d*)', box).group(1)

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'Generali' in page_1 and not 'Proama' in page_1:
        box = polisa_box(pdf, 0, 200, 590, 730)
        total = re.search(r'(RAZEM:|Składka|\(TOTAL\))(?!GRUPĘ) (\d*\s?\d+,\d*)\s?zł', box, re.I)
        total = float(total.group(2).replace(' ', '').replace(',', '.'))

        if re.findall(r'(?=.*jednorazow[o|a])(?=.*przele[w|wem]).*', box, re.I | re.DOTALL) and not 'III rata' in box \
                or re.search(r'przelew[em]', box, re.I) and not 'II rata' in box:
            termin = re.search(r'(płatna\s?do|płatności)\s?(\d{2}.\d{2}.\d{4})', box, re.I)
            termin_I = re.sub('[^0-9]', '-', termin.group(2))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif ('została pobrana' in box or 'została opłacona' in box) and not 'II rata' in page_1:
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*przele[w|em])(?=.*II rata)(?!.III rata).*', box,
                        re.I | re.DOTALL):  # and not 'III rata' in box:
            rata_I = float(re.search(r'I rata (\d*\s?\d+,?\d*)', box, re.I).group(1).replace(' ', '').replace(',', '.'))
            rata_II = float(
                re.search(r'II rata (\d*\s?\d+,?\d*)', box, re.I).group(1).replace(' ', '').replace(',', '.'))

            terminy = re.search(r' I rata .* płatna do (\d{2}.\d{2}.\d{4}).*(\d{2}.\d{2}.\d{4})', box, re.I)
            termin_I = datetime.strptime(term_pln(terminy, 1), '%Y-%m-%d') + one_day
            termin_II = datetime.strptime(term_pln(terminy, 2), '%Y-%m-%d') + one_day

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif 'III rata' in page_1:

            (rata_I := re.search(r'I rata (\d*\s?\d+)', box, re.I).group(1))
            (rata_II := re.search(r'II rata (\d*\s?\d+)', box, re.I).group(1))
            (rata_III := re.search(r'III rata (\d*\s?\d+)', box, re.I).group(1))

            def termin(terminy, n):
                zamiana_sep = re.sub('[^0-9]', '-', terminy.group(n))
                return re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', zamiana_sep)

            terminy = re.search(r' I rata .* płatna do (\d{2}.\d{2}.\d{4}).*(\d{2}.\d{2}.\d{4}).*(\d{2}.\d{2}.\d{4})',
                                box, re.I)
            termin_I = datetime.strptime(termin(terminy, 1), '%Y-%m-%d') + one_day
            termin_II = datetime.strptime(termin(terminy, 2), '%Y-%m-%d') + one_day
            termin_III = datetime.strptime(termin(terminy, 3), '%Y-%m-%d') + one_day
            return total, termin_I, rata_I, 'P', 3, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif re.search('^(?!.*PEUGEOT).*HDI.*$', page_1, re.I) and not 'PZU' in page_1 and not 'TUZ' in page_1 \
            or re.search('^(?!.*PEUGEOT).*HDI.*$', page_1, re.I) and '„WARTA” S.A. POTWIERDZA':

        box = polisa_box(pdf, 0, 200, 590, 630)
        total = zam_spacji(re.search(r'(ŁĄCZNA SKŁADKA|Składka łączna) (\d*\s?\d+)', box, re.I).group(2))

        if 'JEDNORAZOWO' in box and 'GOTÓWKA' in box:
            rata_I = zam_spacji(re.search(r'kwota: (\d*\s?\d+)', box, re.I).group(1))
            termin_I = re.search(r'(termin płatności:|Termin:) (\d{4}-\d{2}-\d{2})', box, re.I).group(2)
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'JEDNORAZOWO' in box and 'PRZELEW' in box:
            rata_I = zam_spacji(re.search(r'kwota: (\d*\s?\d+)', box, re.I).group(1))
            termin_I = re.search(r'(termin płatności:|Termin:) (\d{4}-\d{2}-\d{2})', box, re.I).group(2)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if '2 RATACH' and 'PRZELEW' in box:
            rata_I = zam_spacji(re.search(r'kwota: (\d*\s?\d+)', box, re.I).group(1))
            rata_II = zam_spacji(re.search(r'kwota: (\d*\s?\d+) (PLN|zł) (\d*\s?\d+)', box, re.I).group(3))
            try:
                termin_I = re.search(r'termin płatności: (\d{4}-\d{2}-\d{2})', box, re.I).group(1)
                termin_II = re.search(r'termin płatności: (\d{4}-\d{2}-\d{2})\s?(\d{4}-\d{2}-\d{2})', box, re.I).group(
                    2)
            except Exception:
                termin = re.search(r'W 2 RATACH Termin: (\d{4}-\d{2}-\d{2})\s*(\d{4}-\d{2}-\d{2})', box, re.I)
                termin_I = termin.group(1)
                termin_II = termin.group(2)
            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'Hestia' in page_1 and not 'MTU' in page_1:
        box = polisa_box(pdf, 0, 120, 590, 600)
        total = re.search(r'DO ZAPŁATY(/ TOTAL PREMIUM)? (\d*\s?\d+)', box, re.I)
        total = int(re.sub(r'([\xa0 ])', '', total.group(2)))

        if not 'II rata' in box and 'gotówka' in box:
            termin_I = re.search(r'płatności (\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif not re.findall(r'(?=.*II\srata)(?=.*przelew|przelewem).*', box, re.I | re.DOTALL):
            try:
                termin_I = re.search(r'płatności (\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            except Exception as e:
                termin_I = re.search(r'(\d{4}[-‑]\d{2}[-‑]\d{2}).*\n.*płatności', box, re.I).group(1)
                print(f'Termin płatności - Hestia {e}')
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*II\srata)(?=.*przelew|przelewem).*', box, re.I | re.DOTALL) and not ' III rata' in box:
            termin_I = re.search(r'płatności I\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            termin_II = re.search(r'II\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            rata_I = re.search(rf'{termin_I},\s*(\d*\s?\d+)', box, re.I).group(1)
            rata_II = re.search(rf'{termin_II},\s*(d*\s?\d+)', box, re.I).group(1)

            return total, zamiana_sep(termin_I), rata_I, 'P', 2, 1, zamiana_sep(termin_II), rata_II, termin_III, \
                   rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*IV\srata)(?=.*przelew|przelewem).*', box, re.I | re.DOTALL):
            termin_I = re.search(r'płatności I\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            termin_II = re.search(r'II\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            termin_III = re.search(r'III\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            termin_IV = re.search(r'IV\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)

            rata_I = re.search(rf'{termin_I},\s*(\d*\s?\d+)', box, re.I).group(1)
            rata_II = re.search(rf'{termin_II},\s*(d*\s?\d+)', box, re.I).group(1)
            rata_III = re.search(rf'{termin_III},\s*(d*\s?\d+)', box, re.I).group(1)
            rata_IV = re.search(rf'{termin_IV},\s*(d*\s?\d+)', box, re.I).group(1)

            return total, zamiana_sep(termin_I), rata_I, 'P', 4, 1, zamiana_sep(termin_II), rata_II, termin_III, \
                   rata_III, termin_IV, rata_IV


    elif 'INTER ' in page_1 and not 'InterRisk' in page_1:
        box = polisa_box(pdf, 0, 220, 590, 490)

        total = re.search(r'kwota składki: (\d*\s?\d+,?\d*)', box, re.I).group(1)
        total = zam_spacji(total)

        if re.findall(r'(?=.*jednorazowo)(?=.*przelewem).*', box, re.I | re.DOTALL):
            (termin := re.search(r'i termin płatności:.*\(do (\d{2}.\d{2}.\d{4})', box, re.I))
            termin_I = re.sub('[^0-9]', '-', termin.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'InterRisk' in page_1:
        pdf_str1 = polisa_str(pdf)[1500:-2600]
        total_match = re.compile(r'(Składka\słączna:\s*|WYSOKOŚĆ\sSKŁADKI\sŁĄCZNEJ:\n)(\d*\s?\d+)')
        (total := re.search(total_match, pdf_str1))
        total = int(re.sub(r'\xa0', '', total.group(2)))

        if re.findall(r'(?=.*jednorazow[o|a])(?=.*płatności:\s*przelewem).*', pdf_str1, re.I | re.DOTALL):
            (termin_I := re.search(r'płatna\sdo\sdnia:\s(\d{4}-\d{2}-\d{2})', pdf_str1, re.I).group(1))

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


        elif re.findall(r'(?=.*Płatność: 2\sraty)(?=.*płatności:\s*gotówką).*', pdf_str1, re.I | re.DOTALL):
            rata_I = re.search(r'1\srata: (\d*\s?\d+,?\d*)', pdf_str1, re.I).group(1).replace(',', '.').replace('\xa0',
                                                                                                                '')
            rata_II = re.search(r'2\srata: (\d*\s?\d+)', pdf_str1, re.I).group(1).replace(',', '.').replace('\xa0', '')

            termin_I = re.search(r'płatna\sdo\sdnia:\s(\d{4}-\d{2}-\d{2})', pdf_str1, re.I).group(1)
            termin_II = re.search(r'2\srata: (.*)(\d{4}-\d{2}-\d{2})', pdf_str1, re.I).group(2)

            return total, termin_I, rata_I, 'G', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


        elif re.findall(r'(?=.*Płatność: 2\sraty)(?=.*płatności:\s*przelewem).*', pdf_str1, re.I | re.DOTALL):
            rata_I = re.search(r'1\srata: (\d*\s?\d+,?\d*)', pdf_str1, re.I).group(1).replace(',', '.').replace('\xa0',
                                                                                                                '')
            rata_II = re.search(r'2\srata: (\d*\s?\d+)', pdf_str1, re.I).group(1).replace(',', '.').replace('\xa0', '')

            termin_I = re.search(r'płatna\sdo\sdnia:\s(\d{4}-\d{2}-\d{2})', pdf_str1, re.I).group(1)
            termin_II = re.search(r'2\srata: (.*)(\d{4}-\d{2}-\d{2})', pdf_str1, re.I).group(2)

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    # Link4
    elif (nr_polisy := re.search('(Numer:?|POLISA DLA PANA)\s?\n?(\w\d+)', page_1, re.I)) and \
            not 'WARTA' in page_1 or 'LINK4' in page_1:
        pdf_str2 = polisa_str(pdf)[400:4800]

        total = re.search(r'(\(w złotych\)|ŁĄCZNIE)\s?(\d*\s?\d+,\d+)', pdf_str2, re.I).group(2)
        total = zam_spacji(total)

        if re.search('(Metoda płatności Karta|rachunku bankowego)', pdf_str2, re.I) or 'Przelew' in pdf_str2 and not \
                'Kolejne raty' in pdf_str2:
            termin = re.search(r'Termin płatności (\d{2}/\d{2}/\d{4})', pdf_str2, re.I)
            try:
                termin_I = re.sub('[^0-9]', '-', termin.group(1))
                termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)
            except Exception:
                termin_I = data_wystawienia() + timedelta(7)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if re.findall(r'(?=.*Przelew)(?=.*Kolejne raty).*', pdf_str2, re.I | re.DOTALL):
            raty = re.search(r'Termin Kwota raty\n.* (\d*\s?\d+,\d{2})\n.* (\d*\s?\d+,\d{2})\n.* (\d*\s?\d+,\d{2})\n.* '
                             r'(\d*\s?\d+,\d{2})', pdf_str2, re.I)
            rata_I = zam_spacji(raty.group(1))
            rata_II = zam_spacji(raty.group(2))
            rata_III = zam_spacji(raty.group(3))
            rata_IV = zam_spacji(raty.group(4))

            terminy = re.search(r'Termin Kwota raty\n(\d{2}/\d{2}/\d{4}).*(\d{2}/\d{2}/\d{4}).*(\d{2}/\d{2}/\d{4}).*'
                                r'(\d{2}/\d{2}/\d{4}).*', pdf_str2, re.I | re.DOTALL)
            termin_I = datetime.strptime(term_pln(terminy, 1), '%Y-%m-%d') + one_day
            termin_II = datetime.strptime(term_pln(terminy, 2), '%Y-%m-%d') + one_day
            termin_III = datetime.strptime(term_pln(terminy, 3), '%Y-%m-%d') + one_day
            termin_IV = datetime.strptime(term_pln(terminy, 4), '%Y-%m-%d') + one_day

            return total, termin_I, rata_I, 'P', 4, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'MTU' in page_1:
        box = polisa_box(pdf, 0, 180, 590, 400)
        total = re.search(r'RAZEM DO ZAPŁATY (\d*\s?\d+)', box, re.I)
        total = int(re.sub(r' ', '', total.group(1)))

        if 'przelew' in box and not 'II rata' in box:
            (termin := re.search(r'i kwoty płatności (\d{4}‑\d{2}‑\d{2})', box, re.I))
            termin_I = re.sub('[^0-9]', '-', termin.group(1))

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if re.findall(r'(?=.*II\srata)(?=.*przelew).*', box, re.I | re.DOTALL) and not re.search(r'IV\srata', box):
            termin_I = re.search(r'płatności I\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            termin_II = re.search(r'II\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)

            rata_I = re.search(rf'{termin_I},\s*(\d*\s?\d+)', box, re.I).group(1)
            rata_II = re.search(rf'{termin_II},\s*(\d*\s?\d+)', box, re.I).group(1)

            return total, zamiana_sep(termin_I), rata_I, 'P', 2, 1, zamiana_sep(termin_II), rata_II, termin_III, \
                   rata_III, termin_IV, rata_IV

        if re.findall(r'(?=.*IV\srata)(?=.*przelew).*', box, re.I | re.DOTALL):
            termin_I = re.search(r'płatności I\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            termin_II = re.search(r'II\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            termin_III = re.search(r'III\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)
            termin_IV = re.search(r'IV\srata\s(\d{4}[-‑]\d{2}[-‑]\d{2})', box, re.I).group(1)

            rata_I = re.search(rf'{termin_I},\s*(\d*\s?\d+)', box, re.I).group(1)
            rata_II = re.search(rf'{termin_II},\s*(\d*\s?\d+)', box, re.I).group(1)
            rata_III = re.search(rf'{termin_III},\s*(\d*\s?\d+)', box, re.I).group(1)
            rata_IV = re.search(rf'{termin_IV},\s*(\d*\s?\d+)', box, re.I).group(1)

            return total, zamiana_sep(termin_I), rata_I, 'P', 4, 1, zamiana_sep(termin_II), rata_II, \
                   zamiana_sep(termin_III), rata_III, zamiana_sep(termin_IV), rata_IV


    elif 'Proama' in page_1:
        box = polisa_box(pdf, 0, 250, 590, 650)
        total = re.search(r'RAZEM: (\d*\s?\d+)', box, re.I)
        total = int(re.sub(r'\xa0', '', total.group(1)))

        if 'przelewem' in box and not 'II rata' in box:
            termin = re.search(r'płatna\s?do (\d{2}.\d{2}.\d{4})', box, re.I)
            termin_I = re.sub('[^0-9]', '-', termin.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*została)(?=.*(pobrana|opłacona)).*', box, re.I):
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.search('(?=.*II\srata)(?!.*III\srata).*', box, re.I):
            rata_I = float(
                re.search(r'I\srata\s(\d*\s?\d+,?\d*)', box, re.I).group(1).replace(' ', '').replace(',', '.'))
            rata_II = float(
                re.search(r'II\srata\s(\d*\s?\d+,?\d*)', box, re.I).group(1).replace(' ', '').replace(',', '.'))

            terminy = re.search(r'I\srata\s.*\spłatna\sdo\s(\d{2}.\d{2}.\d{4}).*(\d{2}.\d{2}.\d{4})', box, re.I)
            termin_I = datetime.strptime(term_pln(terminy, 1), '%Y-%m-%d') + one_day
            termin_II = datetime.strptime(term_pln(terminy, 2), '%Y-%m-%d') + one_day

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'PZU' in page_1:
        pdf_str = polisa_str(pdf)[200:5000]
        total = re.search(r'(Składka łączna:|kwota:) (\d*\s?\d+,?\d{2}?)', pdf_str, re.I).group(2)
        total = zam_spacji(total)

        if re.search('opłacon[ao] w całości', pdf_str, re.I) or 'zapłacono gotówką' in pdf_str:
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if not 'została opłacona w całości.' in pdf_str and 'płatność: półroczna' in pdf_str and \
                'zapłacono gotówką:' in pdf_str:
            rata_I = re.search(r'zapłacono gotówką:\s(\d*\s?\d+,?\d{2}?)', pdf_str, re.I).group(1)
            rata_II = re.search(r'(\d{2}\.\d{2}\.\d{4}.*)\s(\d*\s?\d+,?\d{2}?)', pdf_str, re.I).group(2)

            termin_I = datetime.today()
            termin_II = term_pln(re.search(r'odbiorca: PZU SA (\d{2}\.\d{2}\.\d{4})', pdf_str), 1)

            return total, termin_I, rata_I, 'G', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if not 'została opłacona w całości.' in pdf_str and 'płatność: półroczna' in pdf_str:
            termin_Ia = re.search(r'odbiorca: PZU SA (\d{2}\.\d{2}\.\d{4})', pdf_str)
            termin_IIa = re.search(r'00-133 Warszawa (\d{2}\.\d{2}\.\d{4})', pdf_str)

            termin_I = term_pln(termin_Ia, 1)
            termin_II = term_pln(termin_IIa, 1)

            rata_I = re.search(rf'{termin_Ia.group(1)} r. – ' + '(\d*\s?\d+,?\d{2}?)', pdf_str, re.I).group(1)
            rata_II = re.search(rf'{termin_IIa.group(1)} r. – ' + '(\d*\s?\d+,?\d{2}?)', pdf_str, re.I).group(1)

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if re.search('(jednorazow[o|a])', pdf_str,
                     re.I) or 'tytule przelewu' in pdf_str or 'Osoba wnioskująca o zmiany' in pdf_str:
            # raty = re.search(r'(Kwota w złotych|PLN)\s*(\d*\s?\d+,\d{2})', pdf_str, re.I).group(2)
            # rata_I = zam_spacji(raty)

            termin_I = re.search(r'Termin płatności:? (\d{2}.\d{2}.\d{4})', pdf_str, re.I)
            termin_I = re.sub('[^0-9]', '-', termin_I.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if re.search('(Rata 1 2\n)', pdf_str):
            raty = re.search(r'Kwota w złotych (\d*\s?\d+,\d{2}) (\d*\s?\d+,\d{2})', pdf_str, re.I)
            rata_I = float(raty.group(1).replace(',', '.').replace(' ', ''))
            rata_II = float(raty.group(2).replace(',', '.').replace(' ', ''))

            terminy = re.search(r'Termin płatności (\d{2}.\d{2}.\d{2}) (\d{2}.\d{2}.\d{2})', pdf_str, re.I)
            termin_I = datetime.strptime(term_pln(terminy, 1), '%y-%m-%d') + one_day
            termin_II = datetime.strptime(term_pln(terminy, 2), '%y-%m-%d') + one_day

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'Rata 1 2 3 4' in pdf_str or 'kwartalna' in pdf_str:
            pdf_str = polisa_box(pdf, 190, 50, 360, 770)

            terminy = re.search(r'(Termin płatności|Harmonogram płatności).*\n(\d{2}.\d{2}.\d{2,4}).*\n'
                                r'(\d{2}.\d{2}.\d{2,4}).*\n(\d{2}.\d{2}.\d{2,4}).*\n(\d{2}.\d{2}.\d{2,4})', pdf_str,
                                re.I | re.DOTALL)
            termin_I = datetime.strptime(term_pln(terminy, 2), '%Y-%m-%d') + one_day
            termin_II = datetime.strptime(term_pln(terminy, 3), '%Y-%m-%d') + one_day
            termin_III = datetime.strptime(term_pln(terminy, 4), '%Y-%m-%d') + one_day
            termin_IV = datetime.strptime(term_pln(terminy, 5), '%Y-%m-%d') + one_day

            raty = re.search(r'Kwota w złotych|Harmonogram płatności\n'
                             r'[0-9.\s\w–]+(\d{3}) zł\n'
                             r'[0-9.\s\w–]+(\d{3}) zł\n'
                             r'[0-9.\s\w–]+([0-9]{3}) zł\n'
                             r'[0-9\.\s\w–]+([0-9]{3}) zł\n',
                             pdf_str, re.I)

            rata_I = float(raty.group(1).replace(',', '.').replace(' ', ''))
            rata_II = float(raty.group(2).replace(',', '.').replace(' ', ''))
            rata_III = float(raty.group(3).replace(',', '.').replace(' ', ''))
            rata_IV = float(raty.group(4).replace(',', '.').replace(' ', ''))

            return total, termin_I, rata_I, 'P', 4, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'TUW' in page_1 and not 'TUZ' in page_1:
        pdf_str3 = polisa_str(pdf)[1000:6500]
        total = re.search(r'(Składka łączna:|ubezpieczeniowa razem:) (\d*\s?\d+) PLN', pdf_str3, re.I)
        total = int(re.sub(r'\xa0', '', total.group(2)))

        termin = re.search(r'Termin płatności.*(\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2})', pdf_str3, re.I | re.DOTALL)
        termin_I = re.sub('[^0-9]', '-', termin.group(1))
        termin_I = datetime.strptime(re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I), '%Y-%m-%d') + one_day

        if re.findall(r'(?=.*GOTÓWKA)(?=.*(I raty|1 rata|JEDNORAZOWO)).*', pdf_str3, re.I | re.DOTALL):
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*PRZELEW)(?=.*JEDNORAZOWO).*', pdf_str3, re.I | re.DOTALL):
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*PRZELEW)(?=.*2 rata).*', pdf_str3, re.I | re.DOTALL):
            rata_I = zam_spacji(re.search(r'Kwota (\d*\s?\d+) PLN (\d*\s?\d+)', pdf_str3, re.I).group(1)) - 10
            rata_II = re.search(r'Kwota (\d*\s?\d+) PLN (\d*\s?\d+)', pdf_str3, re.I).group(2)

            terminy = re.search(r'Termin płatności (\d{2}-\d{2}-\d{4}) (\d{2}-\d{2}-\d{4})', pdf_str3, re.I)
            termin_I = term_pln(terminy, 1)
            termin_II = term_pln(terminy, 2)

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'TUZ' in page_1:
        pdf_str = polisa_str(pdf)[1500:6500]
        total = re.search(r'(kwota|Składka) do zapłaty (.*\D) (\d*\s?\d+)', pdf_str, re.I)
        total = int(re.sub(r' ', '', total.group(3)))
        # print(pdf_str)
        if wpisowe := re.search('wysokość wpisowego 10 zł\): (TAK)', pdf_str):
            total = total - 20

        if re.findall(r'(?=.*JEDNORAZOWA)(?=.*Gotówka).*', pdf_str, re.I | re.DOTALL):
            termin_I = re.search(r'płatn[ey] do dnia (\d{4}-\d{2}-\d{2})', pdf_str, re.I).group(1)

            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*JEDNORAZOWA)(?=.*Przelew).*', pdf_str, re.I | re.DOTALL):
            termin_I = re.search(r'płatn[ey] do dnia (\d{4}-\d{2}-\d{2})', pdf_str, re.I).group(1)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif 'PÓŁROCZNA' in pdf_str and 'Przelew' in pdf_str:
            rata_I = re.search(r'Kwota wpłaty w zł (\d+)', pdf_str, re.I).group(1)
            rata_II = re.search(r'Kwota wpłaty w zł (.+) (\d*)', pdf_str, re.I).group(2)
            termin_I = re.search(r'Termin płatności (\d{4}-\d{2}-\d{2})', pdf_str, re.I).group(1)
            termin_II = re.search(r'Termin płatności (.*)(\d{4}-\d{2}-\d{2})', pdf_str, re.I).group(2)
            if wpisowe:
                rata_I = int(rata_I) - 20

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'UNIQA' in page_1 and not 'Sygnatura' in page_1 and not 'Archiwizacja' in page_1:
        pdf_str2 = polisa_str(pdf)[500:-1]
        total = re.search(r'(do\szapłacenia|(?:.*)Składka łączn[ie|a]+:|Składk[a|i]+:|Płatność:) (\d*\s?\d+)',
                          pdf_str2, re.I).group(2)
        # print(pdf_str2)
        if re.findall(r'(?=.*jednorazow[o|a]+)(?=.*gotówka).*', pdf_str2, re.I | re.DOTALL):
            termin_I = re.search(r'(Składka została opłacona dnia:?)\s*'
                                 r'(\d{4}[-\./]\d{2}[-\./]\d{2}|\d{2}[-\./]\d{2}[-\./]\d{4})', pdf_str2,
                                 re.I | re.DOTALL)
            termin_I = term_pln(termin_I, 2)
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if not re.findall(r'(?=.*Rata\s2)(?=.*Nr\skonta).*', pdf_str2, re.I | re.DOTALL):
            termin_I = re.search(r'(Termin płatności:?|opłacona do dnia:)\s*'
                                 r'(\d{4}[-\./]\d{2}[-\./]\d{2}|\d{2}[-\./]\d{2}[-\./]\d{4})', pdf_str2,
                                 re.I | re.DOTALL)
            termin_I = term_pln(termin_I, 2)
            # termin_I = re.sub(r'(\d{4}[-./]\d{2}[-./]\d{2}|\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if re.findall(r'(?=.*Rata\s2)(?=.*Nr\skonta).*', pdf_str2, re.I | re.DOTALL):
            rata_I = re.search(r'(Rata 1:) (\d*\s?\d+,\d*)', pdf_str2).group(2).replace(',', '.')
            rata_II = re.search(r'(Rata 2:) (\d*\s?\d+,\d*)', pdf_str2).group(2).replace(',', '.')

            termin_I = term_pln(
                re.search(r'Rata 1:(.*)zł (\d{4}[-\./]\d{2}[-\./]\d{2}|\d{2}[-\./]\d{2}[-\./]\d{4})', pdf_str2,
                          re.I), 2)
            termin_II = term_pln(
                re.search(r'Rata 2:(.*)zł (\d{4}[-\./]\d{2}[-\./]\d{2}|\d{2}[-\./]\d{2}[-\./]\d{4})', pdf_str2,
                          re.I | re.DOTALL), 2)

            return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        """Uniqa sie z AXA"""
    elif 'UNIQA' in page_1:
        # box = polisa_box(pdf, 0, 300, 590, 700)
        pdf_str2 = polisa_str(pdf)[1000:4500]
        total = re.findall(r'Składka łączna: (\d*\s?\d+)', pdf_str2, re.I)[
            -1]  # re.findall("pattern", "target_text")[-1]
        total = int(re.sub(r'\xa0', '', total))

        if re.findall(r'(?=.*przelewem)(?=.*jednorazowo).*', pdf_str2, re.I | re.DOTALL):
            termin_I = re.search(r'do dnia: (\d{2}.\d{2}.\d{4})', pdf_str2)
            termin_I = re.sub('[^0-9]', '-', termin_I.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if re.findall(r'(?=.*gotówk[aą])(?=.*jednorazowo).*', pdf_str2, re.I | re.DOTALL):
            termin_I = re.search(r'(Składka została opłacona dnia:?)\s*'
                                 r'(\d{4}[-\./]\d{2}[-\./]\d{2}|\d{2}[-\./]\d{2}[-\./]\d{4})', pdf_str2,
                                 re.I | re.DOTALL)
            termin_I = term_pln(termin_I, 2)
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'przelewem' in pdf_str2 and 'w ratach' in pdf_str2:
            if 'II.' in pdf_str2:
                # print(pdf_str2)
                (rata_I := re.search(r'\nI. (\d*\s?\d+)', pdf_str2, re.I).group(1))
                (rata_II := re.search(r'II. (\d*\s?\d+)', pdf_str2, re.I).group(1))

                (termin := re.search(r'/(\d{2}.\d{2}.\d{4})', pdf_str2))
                termin_I = re.sub('[^0-9]', '-', termin.group(1))
                termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

                (termin_II := re.search(r'II. (.*)/(\d{2}.\d{2}.\d{4})', pdf_str2))
                termin_II = re.sub('[^0-9]', '-', termin_II.group(2))
                termin_II = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_II)

                return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        if 'gotówką' in pdf_str2 and 'w ratach' in pdf_str2 and 'II.' in pdf_str2:
            # print(pdf_str2)
            rata_I = re.search(r'\nI. (\d*\s?\d+)', pdf_str2, re.I).group(1)
            rata_II = re.search(r'II. (\d*\s?\d+)', pdf_str2, re.I).group(1)

            termin = re.search(r'/(\d{2}.\d{2}.\d{4})', pdf_str2)
            termin_I = re.sub('[^0-9]', '-', termin.group(1))
            termin_I = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_I)

            termin = re.search(r'II. (.*)/(\d{2}.\d{2}.\d{4})', pdf_str2)
            termin_II = re.sub('[^0-9]', '-', termin.group(2))
            termin_II = re.sub(r'(\d{2})-(\d{2})-(\d{4})', r'\3-\2-\1', termin_II)

            return total, termin_I, rata_I, 'G', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


    elif 'WARTA' in page_1:
        pdf_str = polisa_str(pdf)[100:-300]
        total = re.search(r'(SKŁADKA ŁĄCZNA|Kwota\s?:?|w kwocie|do zapłaty \(zł\):) (\d*.\d+)', pdf_str, re.I)
        total = int(total.group(2).replace('\xa0', '').replace('.', '').replace(' ', ''))

        if re.findall(r'(?=.*JEDNORAZOWO)(?=.*GOTÓWK[A|Ą]).*', pdf_str, re.I | re.DOTALL) and not 'PRZELEW' in pdf_str:
            termin_I = re.search(r'(Termin:|DO DNIA)\s*(\d{4}-\d{2}-\d{2})', pdf_str, re.I).group(2)
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*JEDNORAZOWO)(?=.*PRZELE[W|EM]).*', pdf_str, re.I | re.DOTALL):
            termin_I = re.search(r'(Termin:|DO DNIA)\s*(\d{4}-\d{2}-\d{2})', pdf_str, re.I).group(2)
            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*2 RATACH)(?=.*GOTÓWKA).*', pdf_str, re.I | re.DOTALL):
            termin = re.search(r'2 RATACH Termin: (\d{4}-\d{2}-\d{2}) (\d{4}-\d{2}-\d{2})', pdf_str)
            raty = re.search(r'Kwota: (\d*\s?\d+) zł (\d*\s?\d+)', pdf_str)
            return total, termin.group(1), raty.group(1), 'G', 2, 1, termin.group(2), raty.group(2), \
                   termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*2 RATACH)(?=.*PRZELEW).*', pdf_str, re.I | re.DOTALL):
            termin = re.search(r'2 RATACH Termin: (\d{4}-\d{2}-\d{2}) (\d{4}-\d{2}-\d{2})', pdf_str)
            raty = re.search(r'Kwota: (\d*\s?\d+) zł (\d*\s?\d+)', pdf_str)
            return total, termin.group(1), zam_spacji(raty.group(1)), 'P', 2, 1, termin.group(2), \
                   zam_spacji(raty.group(2)), termin_III, rata_III, termin_IV, rata_IV

        elif re.findall(r'(?=.*3 rat[ach])(?=.*PRZELEW).*', pdf_str, re.I | re.DOTALL) and not '4 rata' in pdf_str:
            print(pdf_str)
            termin = re.search(r'Termin: (\d{4}-\d{2}-\d{2}) (\d{4}-\d{2}-\d{2}) (\d{4}-\d{2}-\d{2})', pdf_str)
            raty = re.search(r'Kwota: (\d+) zł (\d+) zł (\d+)', pdf_str)
            return total, termin.group(1), raty.group(1), 'P', 3, 1, termin.group(2), raty.group(2), \
                   termin.group(3), raty.group(3), termin_IV, rata_IV

        elif (inne := re.findall(r'(?=.*3 rata)(?=.*PRZELEW).*', pdf_str, re.I | re.DOTALL) and '4 rata' in pdf_str) or \
                (transportowe := 'W 4 RATACH' in pdf_str):
            if inne:
                termin = re.search(r'4 RATACH Termin: (\d{4}-\d{2}-\d{2}) (\d{4}-\d{2}-\d{2}) (\d{4}-\d{2}-\d{2}) '
                                   r'(\d{4}-\d{2}-\d{2})', pdf_str)
                raty = re.search(r'Kwota: (\d*\s?\d+) zł (\d*\s?\d+) zł (\d*\s?\d+) zł (\d*\s?\d+)', pdf_str)
                return total, termin.group(1), raty.group(1), 'P', 4, 1, termin.group(2), raty.group(2), \
                       termin.group(3), raty.group(3), termin.group(4), raty.group(4)
            elif transportowe and not 'Pakiet Przedsiębiorca' in pdf_str:
                total = (re.search('(SKŁADKA ŁĄCZNA|zawartej umowy ubezpieczenia):?\s?(\d*.?\d+)', pdf_str).group(2)).replace('.', '')
                termin = re.search(r'Termin: (\d{4}-\d{2}-\d{2}) (\d{4}-\d{2}-\d{2})'
                                   r' (\d{4}-\d{2}-\d{2}) (\d{4}-\d{2}-\d{2})', pdf_str)
                'Termin: 2021-09-01 2021-11-24 2022-02-24 2022-05-24'
                raty = re.search(r'Kwota\s?: (\d*\s?\d+) zł (\d*\s?\d+) zł (\d*\s?\d+) zł (\d*\s?\d+)', pdf_str)
                return zam_spacji(total), \
                       termin.group(1), \
                       zam_spacji(raty.group(1)), \
                       'P', 4, 1, \
                       termin.group(2), \
                       zam_spacji(raty.group(2)), \
                       termin.group(3), \
                       zam_spacji(raty.group(3)), \
                       termin.group(4), \
                       zam_spacji(raty.group(4))

            elif 'Pakiet Przedsiębiorca' in pdf_str:
                raty = re.search(r'Kwota\s?:\s*(\d*.\d+),?\d*\s*(\d*.\d+),?\d*\s*(\d*\.\d+),?\d*\s*(\d*\.?\d+),?\d*',
                                 pdf_str)
                rata_I = raty.group(1).replace(',', '').replace('.', '')
                rata_II = raty.group(2).replace(',', '').replace('.', '')
                rata_III = raty.group(3).replace(',', '').replace('.', '')
                rata_IV = raty.group(4).replace(',', '').replace('.', '')
                termin = re.search(r'4 RATACH Termin płatności\s?: 1. (\d{4}-\d{2}-\d{2}) 2. (\d{4}-\d{2}-\d{2}) 3. ' \
                                   '(\d{4}-\d{2}-\d{2}) 4. (\d{4}-\d{2}-\d{2})', pdf_str)

                return total, termin.group(1), rata_I, 'P', 4, 1, termin.group(2), rata_II, \
                       termin.group(3), rata_III, termin.group(4), rata_IV

    # Wiener
    elif re.search('wiener', page_1, re.I):
        pdf_str = polisa_str(pdf)
        total = re.search(r'(SKŁADKA\sŁĄCZNA|Kwota\s|W\skwocie|оплате):?\s(\d*\s?\.?\d+)', pdf_str, re.I)
        try:
            total = int(total.group(2).replace('\xa0', '').replace('.', '').replace(' ', ''))
        except:
            pass

        if re.search(r'(?=.*gotówka)', pdf_str, re.I | re.DOTALL):
            return total, termin_I, rata_I, 'G', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

        elif re.search(r'(?=.*przelew)(?!.*II\srata).*', pdf_str, re.I | re.DOTALL):

            terminI = re.search(r'(Wysokośćratwzł\n|do\sdnia\s|rata)\s?(\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2})', pdf_str,
                                re.I)
            termin_I = term_pln(terminI, 2)

            return total, termin_I, rata_I, 'P', 1, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


        elif re.findall(r'(?=.*przelew)(?=.*II\srata).*', pdf_str, re.I | re.DOTALL):
            try:
                terminI = re.search(
                    r'(Wysokośćratwzł\n|kwoty płatności|do\sdnia\s|rata)\s?(\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2})', pdf_str,
                    re.I)
                terminII = re.search(r'(II\srata)\s?(\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2})', pdf_str, re.I)

                termin_I = term_pln(terminI, 2)
                termin_II = term_pln(terminII, 2)

                rata_I = re.search(r'I\srata\s(\d{2}­\d{2}­\d{4})?\s–\s(\d*\s?\d+,\d{2})', pdf_str).group(2)
                rata_II = re.search(r'II\srata\s(\d{2}­\d{2}­\d{4})?\s–\s(\d*\s?\d+,\d{2})', pdf_str).group(2)
                rata_I = rata_I.replace(',', '.')
                rata_II = rata_II.replace(',', '.')
            except:
                rata_I = 'BRAK'
                rata_II = 'BRAK'
                print('\nWIENER - Brak informacji o szczegółach płatności!\n')

                return total, termin_I, rata_I, 'P', 2, 1, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV

    return total, termin_I, rata_I, '', '', '', termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


"""Koniec arkusza EXCEL"""
def rozpoznanie_danych(tacka_na_polisy):
    pdf = tacka_na_polisy
    page_ = polisa(pdf)
    page_1, page_1_tok = page_[0], page_[1]
    d = dict(enumerate(page_1_tok))

    if 'TUW' in page_1 and not 'TUZ' in page_1:
        page_123_ = polisa_str(pdf)[:]
        page_123 = words_separately(page_123_)
        d = dict(enumerate(page_123))
    p_lub_r = pesel_regon(d, page_1)

    pr_j = prawo_jazdy(page_1, pdf)

    nazwisko_imie_ = nazwisko_imie(d, page_1, pdf)
    nazwisko = '' if regon_checksum(p_lub_r[1:]) else nazwisko_imie_[0]
    imie = '' if regon_checksum(p_lub_r[1:]) else nazwisko_imie_[1]
    regon_ = regon(p_lub_r[1:])
    nazwa_firmy, ulica_f, nr_ulicy_f, nr_lok, kod_poczt_f, miasto_f, tel, email = regon_
    ulica_f_edit = f'{ulica_f} {nr_ulicy_f}' if not nr_lok else f'{ulica_f} {nr_ulicy_f} m {nr_lok}'
    kod_poczt_f_edit = f'{kod_poczt_f[:2]}-{kod_poczt_f[2:]}' if '-' not in kod_poczt_f else kod_poczt_f
    kod_poczt = kod_pocztowy(page_1, pdf) if kod_pocztowy(page_1, pdf) else kod_poczt_f_edit

    tel_mail_ = tel_mail(page_1, pdf, d, nazwisko)
    tel = tel_mail_[0].replace('\n', '') if tel_mail_[0] and tel_mail_[0] != True else ''
    email = tel_mail_[1]

    przedmiot_ub_ = przedmiot_ub(page_1, pdf)

    marka = przedmiot_ub_[0]
    kod = przedmiot_ub_[1]
    model = przedmiot_ub_[2]
    miasto = przedmiot_ub_[3]
    nr_rej = przedmiot_ub_[4]
    adres = przedmiot_ub_[5]
    rok = przedmiot_ub_[6]

    # przedmiot_ub_ = ['', '', '', '', '', '', '']
    # try:
    #     przedmiot_ub_ = przedmiot_ub(page_1, pdf)
    # except:
    #     pass
    # marka = przedmiot_ub_[0] if przedmiot_ub_[0] else ''
    # kod = przedmiot_ub_[1] if przedmiot_ub_[1] else ''
    # model = przedmiot_ub_[2] if przedmiot_ub_[2] else ''
    # miasto = przedmiot_ub_[3] if przedmiot_ub_[3] else ''
    # nr_rej = przedmiot_ub_[4] if przedmiot_ub_[4] else ''
    # adres = przedmiot_ub_[5] if przedmiot_ub_[5] else ''
    # rok = przedmiot_ub_[6] if przedmiot_ub_[6] else ''

    data_wyst = data_wystawienia()
    data_konca = koniec_ochrony(page_1, pdf)

    numer_polisy_ = numer_polisy(page_1, pdf)
    tow_ub_tor = numer_polisy_[0]
    tow_ub = numer_polisy_[1]
    nr_polisy = numer_polisy_[2]

    przypis_daty_raty_ = przypis_daty_raty(pdf, page_1)

    przypis = przypis_daty_raty_[0]
    termin_I = przypis_daty_raty_[1] if przypis_daty_raty_[1] else data_wyst  # gotówka
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

    return nazwa_firmy, nazwisko, imie, p_lub_r, pr_j, ulica_f_edit, kod_poczt, miasto_f, tel, email, marka, kod, model, \
           miasto, nr_rej, adres, rok, data_wyst, data_konca, tow_ub_tor, tow_ub, nr_polisy, przypis, termin_I, \
           rata_I, f_platnosci, ilosc_rat, nr_raty, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV


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
    wb = ExcelApp.Workbooks("2014 BAZA MAGRO.xlsx")
    ws = wb.Worksheets("BAZA 2014")
    # workbook = ExcelApp.Workbooks("Baza.xlsx")

except:
    ExcelApp = Dispatch("Excel.Application")
    wb = ExcelApp.Workbooks.Open(path + "\\2014 BAZA MAGRO.xlsx")
    ws = wb.Worksheets("BAZA 2014")

ExcelApp.Visible = True

"""Jesienne Bazie"""
# try:
for dane_polisy in tacka_na_polisy(obj):
    nazwa_firmy, nazwisko, imie, p_lub_r, pr_j, ulica_f_edit, kod_poczt, miasto_f, tel, email, marka, kod, model, \
    miasto, nr_rej, adres, rok, data_wyst, data_konca, tow_ub_tor, tow_ub, nr_polisy, przypis, ter_platnosci, rata_I, \
    f_platnosci, ilosc_rat, nr_raty, termin_II, rata_II, termin_III, rata_III, termin_IV, rata_IV = dane_polisy
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
    ExcelApp.Cells(row_to_write, 15).Value = pr_j
    ExcelApp.Cells(row_to_write,
                   16).Value = ulica_f_edit  # f'{ulica_f} {nr_ulicy_f}' if not nr_lok else f'{ulica_f} {nr_ulicy_f} m {nr_lok}'
    ExcelApp.Cells(row_to_write, 17).Value = kod_poczt  # kod_pocztowy(page_1) if not kod_poczt_f else kod_poczt_f_edit
    ExcelApp.Cells(row_to_write, 18).Value = miasto_f
    ExcelApp.Cells(row_to_write, 19).Value = tel
    ExcelApp.Cells(row_to_write, 20).Value = email.lower() if email else ''
    ExcelApp.Cells(row_to_write, 23).Value = marka if marka else kod
    ExcelApp.Cells(row_to_write, 24).Value = model if model else miasto
    ExcelApp.Cells(row_to_write, 25).Value = nr_rej if nr_rej else adres
    ExcelApp.Cells(row_to_write, 26).Value = rok
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
    ExcelApp.Cells(row_to_write, 58).Value = 'aut'
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

        if rata_III:
            ws.Range(f'A{row_to_write + 1}:BH{row_to_write + 1}').Copy()
            ws.Range(f'A{row_to_write + 2}').PasteSpecial()

            ExcelApp.Cells(row_to_write + 2, 49).Value = termin_III
            ExcelApp.Cells(row_to_write + 2, 50).Value = rata_III
            ExcelApp.Cells(row_to_write + 2, 53).Value = 3

            if rata_IV:
                ws.Range(f'A{row_to_write + 2}:BH{row_to_write + 2}').Copy()
                ws.Range(f'A{row_to_write + 3}').PasteSpecial()

                ExcelApp.Cells(row_to_write + 3, 49).Value = termin_IV
                ExcelApp.Cells(row_to_write + 3, 50).Value = rata_IV
                ExcelApp.Cells(row_to_write + 3, 53).Value = 4

"""Opcje zapisania"""
ExcelApp.DisplayAlerts = False
wb.SaveAs(path + "\\2014 BAZA MAGRO.xlsx")

"""Zamknięcie narazie wyłączone..."""
# wb.Close()
ExcelApp.DisplayAlerts = True

# ExcelApp.Application.Quit()

end_time = time.time() - start_time
print('Czas wykonania: {:.2f} sekund'.format(end_time))
