"""
 _   _____  
| | /___  \     Intelligent Infrastructure
| |  ___| |     script created by: Rafal Kucharski
| | /  ___/     13/01/2012
| | | |___      info: info@intelligent-infrastructure.eu
|_| |_____|     Copyright (c) Intelligent Infrastructure 2012 

VISUM -> EXCEL tutorial

=====================
Dependencies:
 
1. xlwt
2. win32com
==========================

jan 2011, Krakow Poland
"""
import os

def Get_Files_In_Dir(dir):
    import glob
            
    os.chdir(dir)
    return glob.glob("*.ver")
        
        
        
def Init(path=None):
        #procedura uruchamia Visuma i laduje plik ze sciezki path
        import win32com.client #biblioteka potrzebna do uruchomienia Visuma        
        Visum=win32com.client.Dispatch('Visum.Visum') #uruchomienie Visuma
        if path!=None: Visum.LoadVersion(path) #zaladuj plik
        return Visum #koniec
    
def GetVisumData(Vesrion_File):
    #tutaj pobieram dane z Visuma
    Data=[] #inicjalizacja zmiennej
    List=Visum.Lists.CreateLinkList #tworze liste z Linkam
    List.AddColumn("VehKmTravPrT(AP)") #dodaje do listy kolumny z wybranymiparametrami
    List.AddColumn("VehHourTravtCur(AP)") #jw
    Data.append(Version_File)
    Data.append(List.Sum(0)) #uzyskuje sume kolumny z Visuma i zapisuje do Data wraz z tytulem (pierwszy string to tytul, drugi to wartosc) 
    Data.append(List.Sum(1)) #jw
    #Data.append([]) #pusta linia
    #ponizej sumy macierzy. nie da sie dostac do macierzy PuT, mozna sie do konkrentej macierzy dla DSeg i tak tu robie - jesli masz to zapisane w innych DSeg niz "C" i "PuT", to musisz zmienic ponizej stringi
    #Data.append(["Suma Macierzy 'Car'",Visum.Net.DemandSegments.ItemByKey("Car").ODMatrix.GetODSum()])
    #Data.append(["Suma Macierzy 'PuT'",Visum.Net.DemandSegments.ItemByKey("PuT").ODMatrix.GetODSum()])
    #Data.append([]) #pusta linia
    #Data.append(["Name","PassKmTrav(AP)","PassHourTrav(AP)","PTripsUnlinked(AP)","PTripsTSys(AP)"]) #tytuly tabeli
    #TSysList=Visum.Net.TSystems.GetMultipleAttributes(["Name","PassKmTrav(AP)","PassHourTrav(AP)","PTripsUnlinked(AP)","PTripsTSys(AP)"]) #pobranie wielu atrybutow dla TSystems
    #for TSys in TSysList: #petla dodajaca wszystkie elementy dla TSys
    #    Data.append(TSys)
    #List=Visum.Lists.CreatePuTStatList  #Lista ze statystykami PuTAssignentStats
    #List.AddColumn("MeanJourneyTimePuT") #dodaje do listy kolumny z wybranymi parametrami
    #List.AddColumn("MeanJourneyDistPuT") #jw
    #Stat=List.SaveToArray(1, -1) #zapisuje liste do tabeli
    #Data.append([]) #pusta linia
    #Data.append(["MeanJourneyTimePuT",str(Stat[0])[2:-4]]) #dodaj wartosci, ostatnie dziwne zapisy odcinaja smieci ze stringa (cudzyslow, etc)
    #Data.append(["MeanJourneyDistPuT",str(Stat[1])[1:-2]]) #jw  
    return Data #zwraca cala tabele z pobranymi danymi
   
def dodajdoExcela(Data):    
    ### Ta procedura zapisuje dane z Visuma (zmienna Data) do Excela przy uzyciu obiektu COM
    import win32com.client  #biblioteka potrzebna do uruchomienia Excela
    Excel=win32com.client.Dispatch("Excel.Application")   #uruchomienie Excela
    Excel.Visible = 1  #Excel domyslnie jest niewidoczny, wiec "uwidaczniamy" go
    Excel.Workbooks.Add() #Nowy Zeszyt w Excelu    
    i=0
    for Linia in Data: #petla dla kazdej pobranej danej z Visuma
        i+=1 #liczniki petli
        j=0 
        for Kolumna in Linia:  #petla dla kazdej kolumny w kolejnym wierszu
            j+=1 #zwieksz licznik                 
            Excel.Cells(i,j).Value = str(Kolumna) #zapisz w Excelu w komorce [i,j]
     
    #Excel.ActiveWorkbook.SaveAs(path)  #nie zapisuje,sam sobie musisz zapisac, chyba ze chcesz automatycznie,to usun # na poczatku linii i wstaw sciezke

try: #nieistotne
    Visum
except: 
    Visum=Init() 

A=[]    
dir="D:\\DropBox\\My Dropbox\\JKO\\Projekty\\2011 Wybrzeze Helskie\\Wersje"    
Version_Files=Get_Files_In_Dir(dir)
for Version_File in Version_Files:
    filename=os.path.join(dir,Version_File)
    Visum.LoadVersion(filename)
    A.append(GetVisumData(Version_File))
    
    
    
    


#Data=GetVisumData() #pobierz dane     
dodajdoExcela(A) #zapisz dane do Excela