#-*- coding: utf-8 -*-

if False:
    import re

    #1. Dodawanie typow przystankow z linii
    SIterator=Visum.Net.StopPoints.Iterator #wykonaj dla kazdego obiektu StopPoint w sieci
    while SIterator.Valid:
        S=SIterator.Item #pobierze kolejny StopPoint (jako obiekt)
        TSys=S.AttValue("Concatenate:LineRoutes\TSysCode").split(",") #sprawdz jakie linie przebiegaja przez ten stop point
        TSys = [str(a) for a in TSys] #Zamien liste na stringi
        TSys=list(set(TSys)) #usun duplikaty
        if TSys==['']:
            S.SetAttValue("TSysSet","") #gdy nie ma zadnego TSys
        elif len(TSys)==1:
            S.SetAttValue("TSysSet",TSys[0]) #gdy jest jeden
        elif len(TSys)==2:
            S.SetAttValue("TSysSet",TSys[0]+","+TSys[1])
        elif len(TSys)==3:
            S.SetAttValue("TSysSet",TSys[0]+","+TSys[1]+","+TSys[2])
        SIterator.Next()


    #2. Przypisanie Stop dla StopArea
    SIterator=Visum.Net.StopAreas.Iterator #wykonaj dla kazdego obiektu StopPoint w sieci
    while SIterator.Valid:
        S=SIterator.Item #pobierz kolejny StopPoint (jako obiekt)
        Numer=int(S.AttValue("StopNo"))
        Koncowka=divmod(Numer,100)[1]
        if Koncowka>0:
            S.SetAttValue("StopNo",Numer-Koncowka)
        SIterator.Next()

    #3. Usuniecie pustych Stops
    SIterator=Visum.Net.Stops.Iterator #wykonaj dla kazdego obiektu StopPoint w sieci
    while SIterator.Valid:
        S=SIterator.Item #pobierz kolejny StopPoint (jako obiekt)
        if S.AttValue("NumStopAreas")==0:
            Visum.Net.RemoveStop(S)
        SIterator.Next()

    SIterator=Visum.Net.StopPoints.Iterator #wykonaj dla kazdego obiektu StopPoint w sieci
    while SIterator.Valid:
        S=SIterator.Item #pobierz kolejny StopPoint (jako obiekt)
        if S.AttValue("Name")[-1]==")":
            S.SetAttValue("Zewn",1)
        SIterator.Next()


        # 1. wstawienie StopPoint i StopArea dla kazdego Stop
    Visum.Graphic.StopDrawing = True #znacznie przyspiesza wykonanie skryptu
    Iterator=Visum.Net.Stops.Iterator #petla po Stops
    print "import przystankow kolejowych"
    print '[Status,Numer, Koncowka,Node.AttValue("No"),Link.AttValue("No"),rel]'

    while Iterator.Valid:
        blad=""
        stop=Iterator.Item
        stopID=stop.AttValue("No")

        stopName=str(stop.AttValue("Name").encode('utf-8'))
        Numer=int(stop.AttValue("No")) #warunkiem bedzie koncowka numer, czy konczy sie na 00

        Koncowka=divmod(Numer,100)[1]
        if Koncowka==0:
            stop.SetAttValue("TypeNo",1) #to jest glowny Stop, nie dodajemy do niego tabliczek

        elif stop.AttValue("Zewn")==None and stop.AttValue("TSYSSET") in ["BUS"]:
            #elif stop.AttValue("TSYSSET") in ["TRAM","BUS,TRAM"]: #tutaj dodajemy odpowiedni filtr
            #elif stop.AttValue("TSYSSET") in ["KM","KM,SKM","SKM","WKD"]: #tutaj dodajemy odpowiedni filtr
            stop.SetAttValue("TypeNo",0)        # TypeNo=0 zostaje

            xS=stop.AttValue("XCoord")
            yS=stop.AttValue("YCoord")

            #snap najblizszego node'a i link'a - >> DOSTOSOWAC PARAMETRY <<
            [Node,res,distS]=Visum.Net.GetNearestNode(xS,yS,500,False)
            if res:
                xN=Node.AttValue("XCoord")
                yN=Node.AttValue("YCoord")
            else:
                blad="nie znalazl wezla"
                Node=Visum.Net.AddNode(10000000+Numer,stop.AttValue("XCoord"),stop.AttValue("YCoord"))

            #ustawienie StopArea w miejscu istniejacego Stop - zmiana wspolrzednych
            StopArea=Visum.Net.AddStopArea(stopID,stop,Node,xS,yS)
            StopArea.SetAttValue("Name",stop.AttValue("Name"))
            StopArea.SetAttValue("StopNo",Numer-Koncowka)

            [Link,res,distL,xL,yL,rel]=Visum.Net.GetNearestLink(xS,yS,400,True) #tylko aktywne
            if res:
                try:
                    StopPoint=Visum.Net.AddStopPointOnLink(stopID,StopArea,Link.AttValue("FromNodeNo"),Link.AttValue("ToNodeNo"),False)
                    try:
                        StopPoint.SetAttValue("RelPos",rel) #ustawienie pozycji StopPoint na linku
                    except:
                        blad="nie przesunal"
                        StopPoint.SetAttValue("Code","nie przesunal")
                    StopPoint.SetAttValue("Name",stop.AttValue("Name"))
                    StopPoint.SetAttValue("TSysSet",stop.AttValue("TSYSSET")) #gdy jest jeden
                except:
                    blad="Nie udalo sie wstawic StopPointa"
            else:
                blad="nie znalazl odcinka"
            if blad == "":
                print str(["OK",Numer, Koncowka,Node.AttValue("No"),Link.AttValue("No"),rel])
            else:
                print str(["BLAD",Numer, Koncowka, blad])

        Iterator.Next()

import win32com.client
Visum = win32com.client.Dispatch("Visum.Visum")
Visum.LoadVersion("E://WBR.ver")


Visum.Graphic.StopDrawing = True
Iterator=Visum.Net.POICategories.ItemByKey(2).POIs.Iterator #tutaj zamiast 2 wpisz numer wartwy POI ze znakami (jest w nawiasie w Visumie)
while Iterator.Valid:
    Znak=Iterator.Item
    if Znak.AttValue("Nazwa").upper()=="D-30":
        #znajdz najblizszy odcinek (skierowany)
        [Link,res,distL,xL,yL,rel]=Visum.Net.GetNearestLink(Znak.AttValue("XCoord"),Znak.AttValue("XCoord"),50,True)  #(X,Y,promien max, tylko aktywne)
        #mozesz skorzystac z 'rel' - to relatywna pozycja na odcinku - z tego mozna odczytac czy to poczatek, czy konie odcinka rel€(0,1)
        if res: #jesli znalazl odcinek w promieniu
            Link.SetAttValue("Ograniczenie",Znak.AttValue("Wartosc")) #przypisanie wartosci do odcinka
    Iterator.Next()
Visum.Graphic.StopDrawing = False



Visum.SaveVersion("E:\\WBR_z_Wszystkimi.ver")




















