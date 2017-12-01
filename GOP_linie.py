#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import win32com.client
Visum = win32com.client.Dispatch("Visum.Visum")
Visum.LoadVersion("E:\setka.ver")


#
# DANE WEJSCIOWE
#
# Sciezka Excela
ExcelPath = "D:/LINIA.xls"
# Arkusz z danymi
Arkusz = "Arkusz1"
# TSys
TSys = "KZK_A"
# numer pierwszej dodanej Vehicle JOurney
# jesli siec jest pusta to moze byc 0, jesli sa dodane Vehicle Journeys, to k niech bedize wieksza niz najwiekszy numer
k = 2000
# kolumna_pierwszego_przystanku=17
# kolumna_pierwszego_przystanku=47
# kolumna_pierwszego_kursu=48


#
# Przygotowanie excela - otwarcie pliku i dostep do arkusza

# otworz plik excela (potrzebny modul xlrd - do pobrania z google, wstawic do python27/lib/site-packages/
plik = xlrd.open_workbook(ExcelPath)
# dostep do konkternego arkusza w ktorym sa dane, uzywajmy zawsze arkusza "DoSkryptu"
arkusz = plik.sheet_by_name(Arkusz)

#
# Przygotowanie obiektow w Visumie potrzebnych przy dodawaniu linii

# to jest obiekt w ktorym zapisane beda kolejne elementy sieci (przystanki, tworzace trase
sciezka = Visum.CreateNetReadRouteSearchTSys()

d = Visum.CreateNetReadRouteSearch()
d.SetForTSys(TSys, sciezka)
# okreslenie parametrow wyszukania najkrotszej sciezki (opis agrumentow w HTML Help: SearchShortestPath
sciezka.SearchShortestPath(1, False, True, 10, 1, 99)

# dostep do obiektow odpowiadajacych za kierunki linii
direction1 = Visum.Net.Directions.ItemByKey(">")
direction2 = Visum.Net.Directions.ItemByKey("<")

# dodawanie UDA dla LineRoutes. Zamkniete w strukturze 'try' (google) dla unikniecia bledow jesli UDAs sa juz wstawione
#try:
# argumenty: Code, Name, Comment, Typ (5 = Int, 3= String)
Visum.Net.LineRoutes.AddUserDefinedAttribute("Przewoznik", "Przewoznik", "Przewoznik", 3)
Visum.Net.LineRoutes.AddUserDefinedAttribute("Tabor", "Tabor", "Tabor", 3)
#except:
# obsluzenie wyjatku - jesli UDAs juz sa - nic nie rob
# pass

# przygotowanie mapy Teryt -> Numer StopPoint, przechowywanej w obiekcie Dict (google)

# pobierz atrybuty podane w formie ["Atr1", "Atr2"] Dla wszystkich obiektów klasy "Zones"

# petla po wszystkich wierszach arkusza, numer kolejnego wiersza to rownum
for rownum in range(1, arkusz.nrows):
    if divmod(rownum, 100)[1] == 0:
        Visum.SaveVersion("D:\setka.ver")
    wartosc = arkusz.row(rownum)  # wszystki komorki koljnego wiersza, wektor

    ln = str(wartosc[1].value)
    print ln
    # dodaj obiekt Line do Visuma o nazwie ln i TSys
    Line = Visum.Net.AddLine(ln, TSys)

    # pobierz nazwe lineroute
    lrn = str(wartosc[2].value)

    # pobierz liste kolejnych przystankow (Teryt)
    przystanki = wartosc[3:6]  # przystanki zaczynaja sie w 18(17+1) kolumnie a koncza na pewno przed 48(47+1)
    route = 0  # inicjalizacja sciezki
    route = Visum.CreateNetElements()  # stworzenieo biektu ktory zbierze kolejne przystanki

    # petla dodajaca wszystkie przystanki
    poprzedni = 0
    petla = 0
    for przystanek in przystanki:
        # warunwk zakonczenia, pusta, albo zerowa wartosc przystanku (albo koniec petli)
        if przystanek.value == "":  # przy porownaniach logicznych dajemy ==, a nie =
            break  # wyjscie z petli
        elif przystanek.value == 0 or przystanek.value == poprzedni:
            pass  # obsluga nieznalezionego przystanku , albo podwojnego przystanku
        else:
            # wx.MessageBox(str(przystanek.value))  # wiadomosc na ekranie o tresci podanej w str (okienko)
            # znajdz konkretny przystanek po jego ItemByKey(No) znalezionym w sloniku na podstawie Teryt

            SP = Visum.Net.StopPoints.ItemByKey(int(przystanek.value))
            route.Add(SP)  # dodaj ten przystanek do trasy
            petla += 1
            poprzedni = przystanek.value

    # stworz line route
    # AddLineRoute(Nazwa,Line (obiekt klasy Line),kierunek,przystanki,sciezka wraz z warunkami szukania trasy)

    if petla > 1:
        lnroute = Visum.Net.AddLineRoute(lrn, Line, direction1, route, sciezka)

        # dodaj wartosci UDAs dla tej linii podane w excelu
        lnroute.SetAttValue("Przewoznik", wartosc[2].value)

        # dodawanie kursow
        kursy = wartosc[7:10]  # kolejne odjazdy zapisane w formie 08:00 zapisane w formie wektora

        # dodaj obiekt TimeProfile o numerze 1 dla danej lnroute
        tp = Visum.Net.AddTimeProfile("1", lnroute)
        ilekursow = 0  # licznik pomocniczy
        # dla kazdego odjazdu dodamy vehicle journey
        for kurs in kursy:
            # warunek wyjscia z petli - pusty kurs
            if kurs.value == "":
                break  # wyjscie
            else:
                vj = Visum.Net.AddVehicleJourney(k, tp)  # dodaj VehicleJourney o numerze dla zadanego TimeProfile
                # zabawa z data/godzina przechwycona z Excela, wbudowane funkjce xlrd
                year, month, day, hour, minute, second = xlrd.xldate_as_tuple(kurs.value, plik.datemode)
                if minute < 10:
                    minute = "0"+str(minute)  # dodanie zera dla czasow ponizej 10 minut
                vj.SetAttValue("Dep", str(hour)+":"+str(minute)+":00")  # przygotowanie do formatu potrzebnego w Visumie
                k += 1  # dodanie licznika petli

        # stworzenie kierunku przeciwnego (Linia, Kierunek przeciwny, LineRoute, Czy kopiowac VehJourneys (True/False)
        #lnroute.InsertOppositeDirection(Line, direction2, lrn, False)
        # lnroute2 = Visum.Net.LineRoutes.ItemByKey(Line, direction2, lrn)
        # lnroute2.SetAttValue("Blad_importu", wartosc[0].value)
        # lnroute2.SetAttValue("Liczba_PolR_OD", str(wartosc[3].value))
        # lnroute2.SetAttValue("Liczba_PolR_DO", str(wartosc[4].value))
        # lnroute2.SetAttValue("Liczba_PolW_OD", str(wartosc[5].value))
        # lnroute2.SetAttValue("Liczba_PolW_DO", str(wartosc[6].value))
        # lnroute2.SetAttValue("Liczba_przyst", str(wartosc[8].value))
        # lnroute2.SetAttValue("Czas_trasy", str(wartosc[9].value))
        # lnroute2.SetAttValue("Nazwa_Linii", wartosc[14].value)
        # lnroute2.SetAttValue("Przewoznik", wartosc[2].value)

        #tp = Visum.Net.AddTimeProfile("2", lnroute2)
        #kursy = wartosc[115:]  # kolejne odjazdy zapisane w formie 08:00  w formie wektora
        # dla kazdego odjazdu dodamy vehicle journey
        #for kurs in kursy:

        #    # warunek wyjscia z petli - pusty kurs
        #    if kurs.value == "":
        #        break  # wyjscie
        #    else:
        #        vj = Visum.Net.AddVehicleJourney(k, tp)  # dodaj VehicleJourney o numerze dla zadanego TimeProfile
        #        # zabawa z data/godzina przechwycona z Excela, wbudowane funkjce xlrd
        #        year, month, day, hour, minute, second = xlrd.xldate_as_tuple(kurs.value, plik.datemode)
        #        if minute < 10:
        #            minute = "0"+str(minute)  # dodanie zera dla czasow ponizej 10 minut
        #        vj.SetAttValue("Dep", str(hour)+":"+str(minute)+":00")  # przygotowanie do formatu potrzebnego w Visumie
        #        k += 1  # dodanie licznika petli
        #Visum.SaveVersion("E:\maz_.ver")

Visum.SaveVersion("D:\koniec.ver")