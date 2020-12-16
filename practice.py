import re


# data = ['InterRisk', 'TU', 'S.A.', 'InterRisk', 'Kontakt', '22', '575', '25', '25', 'Tu', 'zepnij', 'polisę',
#         'tVOuelid.le.', 'Łdn(ąz4nik2aao)', 'ł', 'I', 'w6n', 'w3sa', 'u4', 'Ł2r-oa95dn,5', 'z9c-5ie00', '-G5,',
#         '5frao4xu', '.Łp', '(ó4d2ź)', '634-55-60', 'Ubezpieczenie', 'komunikacyjne', 'Pakiet', 'Auto+', 'Polisa',
#         'seria', 'KM-H/P', 'numer', '0155239', 'Przeglądaj', 'swoje', 'polisy', 'już', 'dzisiaj.', 'Zarejestruj',
#         'Chcesz', 'samochód', 'zastępczy', 'na', 'czas', 'naprawy', 'pojazdu', 'z', 'OC', 'lub', 'AC', '?', 'się',
#         'w', 'portalu', 'Klienta:', 'https://klient.interrisk.pl', 'Skorzystaj', 'z', 'Sieci', 'Partnerskiej',
#         '21200059142244', 'Niniejsza', 'polisa', 'jest', 'jednocześnie', 'wnioskiem', 'ubezpieczeniowym', '1.',
#         'DANE', 'UBEZPIECZAJĄCEGO', '/', 'UBEZPIECZONEGO', 'WŁAŚCICIEL,', 'OSOBA', 'FIZYCZNA', 'Imie', 'i', 'nazwisko:',
#         'ANATOLII', 'BIDNYI', 'PESEL:', '81072720412', 'Adres:', '90-349', 'ŁÓDŹ,', 'KS.', 'BP.', 'WINCENTEGO',
#         'TYMIENIECKIEGO', '16D', '/', '45', 'Adres', 'korespondencyjny:', 'UL.', 'KS.', 'BP.', 'WINCENTEGO',
#         'TYMIENIECKIEGO', '16D/45;', '90-349', 'ŁÓDŹ', 'POLSKA;', 'POWIAT', 'ŁÓDŹ', '2.', 'DANE', 'POJAZDU',
#         'Rodzaj:', 'SAMOCHODY', 'OSOBOWE', 'Rok', 'produkcji:', '2001', 'Pojemność:', '1598', 'cm3', 'Marka/typ/model:',
#         'VOLKSWAGEN', 'GOLF', 'IV', '1.6', 'BASIS', 'Data', '1', 'rejestracji:', '2001-10-03', 'Przebieg:', '200000',
#         'km', 'Nr', 'rejestracyjny:', 'EL78929', 'Nr', 'VIN:', 'WVWZZZ1JZ2W240562', 'Liczba', 'miejsc:', '5',
#         'Uszkodzenia:', 'BEZ', 'USZKODZEŃ', 'Sposób', 'używania:', 'PRYWATNY', '3.', 'UBEZPIECZENIE',
#         'ODPOWIEDZIALNOŚCI', 'CYWILNEJ', 'POSIADACZY', 'POJAZDÓW', 'MECHANICZNYCH', '+', 'ZIELONA', 'KARTA',
#         'Okres', 'ubezpieczenia:', 'od', 'dnia:', '2019-12-20', 'godz.:', '00:00', 'do', 'dnia:', '2020-12-19',
#         'godz.:', '23:59', 'Rodzaj', 'umowy:', 'nowa', 'Przebieg', 'ubezpieczenia', 'zweryfikowany', 'w', 'UFG',
#         'Suma', 'gwarancyjna:', 'szkody', 'na', 'osobie', '5', '210', '000', 'euro,', 'szkody', 'w', 'mieniu', '1',
#         '050', '000', 'euro', '-', 'w', 'odniesieniu', 'do', 'jednego', 'zdarzenia,', 'którego', 'skutki', 'są', 'objęte',
#         'ubezpieczeniem', 'bez', 'względu', 'na', 'liczbę', 'poszkodowanych', 'Wydano', '„Potwierdzenie', 'zawarcia',
#         'umowy', 'ubezpieczenia', 'OC”:', 'seria', 'i', 'numer', 'polisy', 'KM-H/P0155239', 'Składka:', '701', 'zł',
#         '4.', 'UBEZPIECZENIE', 'AUTOASSISTANCE', 'Okres', 'ubezpieczenia:', 'od', 'dnia:', '2019-12-20', 'godz.:', '00:00',
#         'do', 'dnia:', '2020-12-19', 'godz.:', '23:59', 'Rodzaj', 'umowy:', 'nowa', 'Wariant:', 'Start', 'Składka:', '0',
#         'zł', '5.', 'SKŁADKA', 'DO', 'ZAPŁATY', 'Składka']


s = """
InterRisk TU S.A. InterRisk Kontakt   22 575 25 25
Tu zepnij polisę tVOuelid.le.  Łdn(ąz4nik2aao) ł I w6n w3sa u4 Ł2r-oa95dn,5 z9c-5ie00 -G5, 5frao4xu .Łp (ó4d2ź) 634-55-60 Ubezpieczenie komunikacyjne Pakiet Auto+
Polisa seria KM-H/P numer 0155239
Przeglądaj swoje polisy już dzisiaj. Zarejestruj  Chcesz samochód zastępczy na czas naprawy pojazdu z OC lub AC ? 
się w portalu Klienta: https://klient.interrisk.pl Skorzystaj z Sieci Partnerskiej 21200059142244
 Niniejsza polisa jest jednocześnie wnioskiem ubezpieczeniowym
1. DANE UBEZPIECZAJĄCEGO / UBEZPIECZONEGO
WŁAŚCICIEL, OSOBA FIZYCZNA
Imie i nazwisko: ANATOLII BIDNYI PESEL: 81072720412
Adres: 90-349 ŁÓDŹ, KS. BP. WINCENTEGO TYMIENIECKIEGO 16D / 45
Adres korespondencyjny: UL. KS. BP. WINCENTEGO TYMIENIECKIEGO 16D/45; 90-349 ŁÓDŹ POLSKA; POWIAT ŁÓDŹ
2. DANE POJAZDU'
"""


data = s.split()


dystans = [data[data.index(adres) - 10: data.index(adres) + 17] for adres in data
                   if re.search('adres?\w+', adres, re.I) or re.search('kontakt?\w+', adres, re.I)
                   or adres.lower() == 'pocztowy'][0]

# print(dystans)

for w in data:
    if re.search('adres?\w+', w, re.I):
        print(data[data.index(w) - 10: data.index(w) + 17])

