#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd

plik = xlrd.open_workbook("E:/Mazowsze/pociagi.xlsx")
arkusz = plik.sheet_by_name("Arkusz3")


sciezka = Visum.CreateNetReadRouteSearchTSys()
d=Visum.CreateNetReadRouteSearch()
d.SetForTSys("B", sciezka)
sciezka.SearchShortestPath(3, True, True, 2, 2, 99)

direction1 = Visum.Net.Directions.ItemByKey(">")
direction2 = Visum.Net.Directions.ItemByKey("<")

# petla po wierszach excela
for rownum in range(1, arkusz.nrows):
    wartosc = arkusz.row(rownum)  # zebranie wartosci calego wiersza
    ln = wartosc[0].value  # nazwa linii
    try:
        Visum.Net.AddLine(ln, "B")
    except:
        pass

    lrn = wartosc[1].value  # nazwa line route
    przystanki = wartosc[2:]  # kolejne przystanki
    route = 0 #czysci poprzednia liste elementow
    route = Visum.CreateNetElements()  # tutaj inicjalizujesz liste przystankow
    # #petla po przystankach
    for przystanek in przystanki:
        if przystanek.value == "":
            break
        else:
            SP=Visum.Net.StopPoints.ItemByKey(przystanek.value)
            route.add(SP)
    Visum.Net.AddLineRoute(ln,ln,direction1,route,sciezka)
