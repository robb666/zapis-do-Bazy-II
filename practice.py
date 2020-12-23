import re
import os


path = os.getcwd()



d = {0: 'wniosek', 1: 'polisa', 2: 'ubezpieczenia', 3: 'nr', 4: '000009550795', 5: 'ubezpieczenie', 6: 'komunikacyjne',
     7: 'euroins', 8: 'auto', 9: 'okres', 10: 'ubezpieczenia', 11: 'od', 12: '2020-11-23', 13: '15', 14: '03', 15: 'do',
     16: '2021-11-22', 17: 'ubezpieczający', 18: 'właściciel', 19: 'grzegorz', 20: 'pesel', 21: '89040401350',
     22: 'telefon', 23: 'diakows4ki', 24: 'adres', 25: '39-460', 26: 'nowa', 27: 'dęba', 28: 'strzelnicza', 29: '27',
     30: 'data', 31: 'urodzenia', 32: '1989-04-04', 33: 'email', 34: 'diakowskigrzegorz', 35: 'gmail', 36: 'com',
     37: 'rok', 38: 'wydania', 39: 'prawa', 40: 'jazdy', 41: '2017-02-23', 42: 'ubezpieczony', 43: 'właściciel',
     44: 'grzegorz', 45: 'pesel', 46: '89040401350', 47: 'telefon', 48: 'diako4wski', 49: 'adres', 50: '39-460',
     51: 'nowa', 52: 'dęba', 53: 'strzelnicza', 54: '27', 55: 'data', 56: 'urodzenia', 57: '1989-04-04', 58: 'email',
     59: 'diakowskigrzegorz', 60: 'gmail', 61: 'com', 62: 'rok', 63: 'wydania'}

with open(path + '\\imiona.txt') as content:
    all_names = content.read().split('\n')
    name = ''
    for k, v in d.items():
        if v.title() in all_names and not re.search('\d', d[k + 4]):
            name = f'{d[k + 4].title()} {v.title()}'
        if v.title() in all_names and re.search('\d', d[k + 4]):
            name = f'{d[k + 5].title()} {v.title()}'


    print(name)





# kod = '1996-01-02 94-004'
#
#
#
# if (re := re.search('\d{2}[-|\xad]\d{3}', kod)):
#     print(re)



# page_1 = """
# InterRisk TU S.A. InterRisk Kontakt   22 575 25 25
# Tu zepnij polisę tVOuelid.le.  Łdn(ąz4nik2aao) ł I w6n w3sa u4 Ł2r-oa95dn,5 z9c-5ie00 -G5, 5frao4xu .Łp (ó4d2ź) 634-55-60 Ubezpieczenie komunikacyjne Pakiet Auto+
# Polisa seria KM-H/P numer 0155239
# Przeglądaj swoje polisy już dzisiaj. Zarejestruj  Chcesz samochód zastępczy na czas naprawy pojazdu z OC lub AC ?
# się w portalu Klienta: https://klient.interrisk.pl Skorzystaj z Sieci Partnerskiej 21200059142244
#  Niniejsza polisa jest jednocześnie wnioskiem ubezpieczeniowym
# 1. DANE UBEZPIECZAJĄCEGO / UBEZPIECZONEGO
# WŁAŚCICIEL, OSOBA FIZYCZNA
# Imie i nazwisko: ANATOLII BIDNYI PESEL: 81072720412
# Adres: 90-349 ŁÓDŹ, KS. BP. WINCENTEGO TYMIENIECKIEGO 16D / 45
# Adres korespondencyjny: UL. KS. BP. WINCENTEGO TYMIENIECKIEGO 16D/45; 90-349 ŁÓDŹ POLSKA; POWIAT ŁÓDŹ
# 2. DANE POJAZDU
# Rodzaj: SAMOCHODY OSOBOWE Rok produkcji: 2001 Pojemność: 1598 cm3
# Marka/typ/model: VOLKSWAGEN GOLF IV 1.6 BASIS Data 1 rejestracji: 2001-10-03 Przebieg: 200000 km
# Nr rejestracyjny: EL78929 Nr VIN: WVWZZZ1JZ2W240562 Liczba miejsc: 5
# Uszkodzenia: BEZ USZKODZEŃ
# Sposób używania: PRYWATNY
# 3. UBEZPIECZENIE ODPOWIEDZIALNOŚCI CYWILNEJ POSIADACZY POJAZDÓW MECHANICZNYCH + ZIELONA KARTA
# Okres ubezpieczenia: od dnia: 2019-12-20 godz.: 00:00 do dnia: 2020-12-19 godz.: 23:59 Rodzaj umowy: nowa
# Przebieg ubezpieczenia zweryfikowany w UFG
# Suma gwarancyjna:  szkody na osobie 5 210 000 euro, szkody w mieniu 1 050 000 euro - w odniesieniu do jednego zdarzenia, którego skutki są objęte
# ubezpieczeniem bez względu na liczbę poszkodowanych
# """


# c = re.compile('(adres\w*|(?<!InterRisk) kontakt\w*|pocztowy|ubezpieczony).+?', re.I)
#
# adres = ''
# if (f := c.search(page_1)):
#     adres = f.group().strip()
#     print(adres)
#
# data = page_1.split()
#
# dystans = [data[data.index(split) - 10: data.index(split) + 200] for split in data if adres in split][0]
#
# print(dystans)
#
# kod_pocztowy = [kod for kod in dystans if re.search('\d{2}-|\xad\d{3}', kod)][0]
# print(kod_pocztowy)















# data = """
# POLISA DLA PANA POJAZDU
# Numer K13857216701
# Ten dokument potwierdza złożenie wniosku o ubezpieczenie
# Na podstawie wniosku o zawarcie ubezpieczenia z dnia 15/05/2019  roku  potwierdzamy zawarcie umowy ubezpieczenia od dnia
# 16/05/2019 w zakresie: Ubezpieczenie Odpowiedzialności Cywilnej Posiadaczy Pojazdów Mechanicznych, Ubezpieczenie Program
# Pomocy z Samochodem Zastępczym i Assistance Opony, Ubezpieczenie Następstw Nieszczęśliwych Wypadków, Ubezpieczenie Smart
# Casco, Ubezpieczenie Auto Assistance w wariancie Auto Asistance
# Okres ubezpieczenia 16/05/2019 (00:00) – 15/05/2020
# Dane pojazdu Marka / Model OPEL Vectra C 2.0 MR'05 Cosmo
# Numer rejestracyjny EL1C276
# Numer VIN W0L0ZCF6871012106
# Rok produkcji 2006
# Liczba miejsc/drzwi 5/4
# Pojemność / Moc silnika 1998 cc/ 129,00 kW
# Rodzaj paliwa Benzyna-gaz
# Sposób użytkowania Wyłącznie prywatny (w tym dojazdy do pracy). Pojazd nie jest użytkowany
# jako: TAXI / Przewóz osób / Usługi kurierskie i transportowe / Nauka jazdy /
# Wynajem / W  wyścigach i rajdach samochodowych
# Okres użytkowania za granicą RP Do 1 miesiąca
# Układ kierowniczy Po lewej stronie
# Rok pierwszego ubezpieczenia w RP
# Imię i nazwisko Radosław  Antosik
# Właściciel pojazdu PESEL 77082209636
# Główny użytkownik Adres Ks. Jerzego Popiełuszki 5m. 89, 94-052 Łódź
# pojazdu Rok uzyskania prawa jazdy 1994 r.
# Młody kierowca Zgodnie z Pana deklaracją przez cały okres ubezpieczenia pojazd nie będzie użytkowany przez osobę, która
# nie ukończyła 24 roku życia.
# Składka i harmonogram Całkowita składka do zapłacenia (w złotych) 1 921,11
# płatności Termin płatności 23/05/2019
# Metoda płatności Przelew
# Termin Kwota raty
# 23/05/2019 480,24
# 16/08/2019 480,29
# 16/11/2019 480,29
# 16/02/2020 480,29
# Kolejne raty powinny zostać opłacone zgodnie z harmonogramem przelewem na nr konta
# 97 1240 2092 9652 0138 5721 6701
# Szczegółowy Ubezpieczenie Odpowiedzialności Cywilnej posiadaczy pojazdów mechanicznych(1)
# zakres ochrony Suma gwarancyjna w odniesieniu do jednego zdarzenia, bez względu na liczbę poszkodowanych:
# ubezpieczeniowej ■ 5 210 000,00 Euro dla wszystkich szkód osobowych, 1 050 000,00 Euro dla wszystkich szkód
# majątkowych
# ■ Składka za OC - 1 487,21 złotych.
# Ubezpieczenie Program pomocy z samochodem zastępczym i assistance opony (2)
# Ubezpieczenie Następstw nieszczęśliwych wypadków (4)
# ■ Suma ubezpieczenia  10 000,00 zł
# Ubezpieczenie Smart Casco (7)
# ■ Suma ubezpieczenia pojazdu  na dzień złożenia wniosku o ubezpieczenie 16 200,00 zł określona wg
# INFO-EKSPERT
# ■ Wariant ubezpieczenia: Szkoda całkowita, Kradzież i Żywioły
# ■ Franszyza redukcyjna: 0 PLN
# Ubezpieczenie Auto Assistance (8)
# ■ Wariant ubezpieczenia: Auto Assistance
# POLISA DLA PANA POJAZDU
# Numer K13857216701
# Ten dokument potwierdza złożenie wniosku o ubezpieczenie
# Na podstawie wniosku o zawarcie ubezpieczenia z dnia 15/05/2019  roku  potwierdzamy zawarcie umowy ubezpieczenia od dnia
# 16/05/2019 w zakresie: Ubezpieczenie Odpowiedzialności Cywilnej Posiadaczy Pojazdów Mechanicznych, Ubezpieczenie Program
# Pomocy z Samochodem Zastępczym i Assistance Opony, Ubezpieczenie Następstw Nieszczęśliwych Wypadków, Ubezpieczenie Smart
# Casco, Ubezpieczenie Auto Assistance w wariancie Auto Asistance
# Okres ubezpieczenia 16/05/2019 (00:00) – 15/05/2020
# Dane pojazdu Marka / Model OPEL Vectra C 2.0 MR'05 Cosmo
# Numer rejestracyjny EL1C276
# Numer VIN W0L0ZCF6871012106
# Rok produkcji 2006
# Liczba miejsc/drzwi 5/4
# Pojemność / Moc silnika 1998 cc/ 129,00 kW
# Rodzaj paliwa Benzyna-gaz
# Sposób użytkowania Wyłącznie prywatny (w tym dojazdy do pracy). Pojazd nie jest użytkowany
# jako: TAXI / Przewóz osób / Usługi kurierskie i transportowe / Nauka jazdy /
# Wynajem / W  wyścigach i rajdach samochodowych
# Okres użytkowania za granicą RP Do 1 miesiąca
# Układ kierowniczy Po lewej stronie
# Rok pierwszego ubezpieczenia w RP
# Imię i nazwisko Radosław  Antosik
# Właściciel pojazdu PESEL 77082209636
# Główny użytkownik Adres Ks. Jerzego Popiełuszki 5m. 89, 94-052 Łódź
# pojazdu Rok uzyskania prawa jazdy 1994 r.
# Młody kierowca Zgodnie z Pana deklaracją przez cały okres ubezpieczenia pojazd nie będzie użytkowany przez osobę, która
# nie ukończyła 24 roku życia.
# Składka i harmonogram Całkowita składka do zapłacenia (w złotych) 1 921,11
# płatności Termin płatności 23/05/2019
# Metoda płatności Przelew
# Termin Kwota raty
# 23/05/2019 480,24
# 16/08/2019 480,29
# 16/11/2019 480,29
# 16/02/2020 480,29
# Kolejne raty powinny zostać opłacone zgodnie z harmonogramem przelewem na nr konta
# 97 1240 2092 9652 0138 5721 6701
# Szczegółowy Ubezpieczenie Odpowiedzialności Cywilnej posiadaczy pojazdów mechanicznych(1)
# zakres ochrony Suma gwarancyjna w odniesieniu do jednego zdarzenia, bez względu na liczbę poszkodowanych:
# ubezpieczeniowej ■ 5 210 000,00 Euro dla wszystkich szkód osobowych, 1 050 000,00 Euro dla wszystkich szkód
# """


# s = 'kod 94-052'


# dystans = [data[data.index(adres) - 10: data.index(adres) + 17] for adres in data
#                    if re.search('adres', adres, re.I) or re.search('kontakt?\w+', adres, re.I)
#                    or adres.lower() == 'pocztowy'][0]


# if (f := re.search('(adres|kontakt?\w+|pocztowy)', data, re.I)):
    # print(data[data.index('Adres'):].split()[6] )
    # print(f)

# for w in data:
#     if re.search('\d{2}-|\xad\d{3}', w, re.I):
#         print(data[data.index(w) - 10: data.index(w) + 17])

