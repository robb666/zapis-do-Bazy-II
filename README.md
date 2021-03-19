# zapis-do-Bazy-II
Przenosi dane z plików .pdf do .xlsx.
Rozpoznaje pesel i regon przez sume kontrolną, jeżeli znajdzie regon wtedy pobiera dane podmiotu z Regon API.
Rozpoznaje imiona porównując je z plikiem imiona.txt, w trudniejszych do parsowania przypadkach porównuje marki 
pojazdów z markami zapisanymi w pliku marki.txt.
