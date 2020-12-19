import re



data = 'Adres kontakt pocztowy ubezpieczeony'

c = re.compile('(adres?\w+|kontakt?\w+|pocztowy|ubezpieczony).+?', re.I)
if (f := c.search(data)):

    print(f.group())















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

