import xlrd
import win32com

#Visum=win32com.client.Dispatch("Visum.Visum")
#Visum.LoadVersion("E:\\KRK_przystanki.ver")


def odczytaj_xls():
    plik = xlrd.open_workbook("E:\\przystanki_PK.xls")
    arkusz = plik.sheet_by_name('Arkusz1')
    rozklady={}
    StopID=-1
    for rownum in range(1,arkusz.nrows):                    
        ID=int(arkusz.cell(rownum, 1).value)
        Nazwa=arkusz.cell(rownum,2).value
        if arkusz.cell(rownum,3).value.upper()=='AUTOBUSOWY':
            TSys="B"
        else:
            TSys="T"
        Y=float(arkusz.cell(rownum,4).value)
        X=float(arkusz.cell(rownum,5).value)
        ID_Busman=str(arkusz.cell(rownum,12).value)
        
        if ID_Busman.split("-")[0]<>StopID:
            [Node,res,dist]=Visum.Net.GetNearestNode(X,Y,1000,False) #znajdz najblizszy Node (duzy radius)   
            if res:
                
                StopID=ID_Busman.split("-")[0]            
                Stop=Visum.Net.AddStop(StopID,X,Y)
                Stop.SetAttValue("Name",Nazwa)               
                StopArea=Visum.Net.AddStopArea(StopID,Stop,Node,X,Y)
                StopArea.SetAttValue("Name",Nazwa)
            else:
                Visum.WriteToLogFile("Blad przy przystankach: nie znaleziono wezla przy przytsanku (Stop Area), linia excela:" + str(rownum) +" ID_Busman: " + ID_Busman)
                              
                
                
         
        [Link,res,dist,xL,yL,rel] = Visum.Net.GetNearestLink(X,Y,1000,False)
        if res: 
            NumerTabliczki=int(str(ID_Busman.replace("-","0")).replace("t","0"))
            try:
                StopPoint = Visum.Net.AddStopPointOnLink(NumerTabliczki,StopArea,Link.AttValue("FromNodeNo"),Link.AttValue("ToNodeNo"),True) 
                try:
                    StopPoint.SetAttValue("RelPos",rel)
                except:
                    Visum.WriteToLogFile("Blad przy przystankach: nie ustawilem relPos tabliczki przystankowej. , linia excela: " + str(rownum) +" ID_Busman: " + ID_Busman)
                    
                StopPoint.SetAttValue("Code",ID_Busman) 
                StopPoint.SetAttValue("TSysSet",TSys)
            except:
                Visum.WriteToLogFile("Blad przy przystankach: nie udalo sie wstawic tabliczki,  linia excela:  " + str(rownum) +" ID_Busman:" + ID_Busman)
                
                                       
        else:
            Visum.WriteToLogFile("Blad przy przystankach: nie znalazlem odcinka dla tabliczki przystankowej, , linia excela: " + str(rownum) +" ID_Busman:" + ID_Busman)
                
    return rozklady                                    
                    

odczytaj_xls()


#[Node,res,dist]=Visum.Net.GetNearestNode(xS,yS,nsnap,False) #znajdz najblizszy Node (duzy radius)   
        