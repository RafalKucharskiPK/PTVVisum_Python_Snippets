#-*- coding: utf-8 -*-
import xlrd

def VisumInit(path=None):
    """
    ###
    Automatic Plate Number Recognition Support
    (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
    ####
    VISUM INIT
    """
    import win32com.client        
    Visum = win32com.client.Dispatch('Visum.Visum.125')
    if path != None: Visum.LoadVersion(path)
    return Visum


def get_stops():
    stop_points=Visum.Net.StopPoints.GetMultipleAttributes(["No","Name"])
    a={}
    for sp in stop_points:
        a[sp[1].upper()]=int(sp[0])         
    return a
    
def odczytaj_xls():
    plik = xlrd.open_workbook("D:\\linie_przystanki_PK.xls")
    arkusz = plik.sheet_by_name('przystanki_autobusowe_linie_PK')
    rozklady={}
    for rownum in range(1,arkusz.nrows):                    
        wrzut=arkusz.cell(rownum,3).value
        if wrzut !="":
            wrzut=wrzut.upper()
        else:
            wrzut=arkusz.cell(rownum,2).value.upper()
        if not rozklady.has_key(wrzut):
            przystanki=arkusz.cell(rownum,2).value.upper()
            rozklady[wrzut]=arkusz.cell(rownum,2).value.upper() 
            #print wrzut," :::::   ",przystanki                       
    return rozklady                                    
                    
def Add_Line(slownik_linie,slownik_stop_points):
    
    routesearchparameters = Visum.CreateNetReadRouteSearchTSys() 
    d=Visum.CreateNetReadRouteSearch()
    d.SetForTSys("B",routesearchparameters)
    routesearchparameters.SearchShortestPath(3,#    SearchShortestPath ( [in] Enum ShortestPathCriterionT, 
                                             True,#[in] VARIANT_BOOL includeBlockedLinksInRouting, 
                                             True,#[in] VARIANT_BOOL includeBlockedTurnsInRouting, 
                                             2,#[in] double MaxDeviationFactor, 
                                             2,#[in] Enum DoIfNotFound, 
                                             99#[in] VARIANT LinkTypeIfInsert 
                                             )
    try: 
        Visum.Net.Lines.AddUserDefinedAttribute("rozklad_pdf","rozklad_pdf","rozklad_pdf",62)
    except:
        pass   
    try: 
        Visum.Net.Lines.AddUserDefinedAttribute("rozklad_przystanki","rozklad_przystanki","rozklad_przystanki",62)
    except:
        pass
    direction1 = Visum.Net.Directions.ItemByKey(">") 
    direction2 = Visum.Net.Directions.ItemByKey("<") 
    TSys_PKS=Visum.Net.TSystems.ItemByKey("PKS")
    TSys_Mikrobus=Visum.Net.TSystems.ItemByKey("B")
    i=0
    l=len(slownik_linie)
    for linia in slownik_linie:  
        i+=1
        if i>1900 and i not in [500,1146,1922,1987]:                    
            if divmod(i,50)[-1]==0:
                Visum.SaveVersion("C:\stops_bckp_2000.ver") 
            Line=Visum.Net.AddLine(linia,"B")
            Line.SetAttValue("rozklad_pdf",linia)
            przystanki=slownik_linie[linia]
            Line.SetAttValue("rozklad_przystanki",przystanki)
            a=przystanki.split("-")
            a=[e.strip() for e in a]
            #kier 1
            print "=============", linia,"============",i,"/",l
            print "przystanki:  ",a
            sp1=0
            if not (len(a)==2 and a[0]==a[-1]):
                Route=0
                Route=Visum.CreateNetElements()
                flag=False
                for sp in a:
                    
                        try:          
                            print "przystanek ",sp, slownik_stop_points[sp]
                            if slownik_stop_points[sp]!=sp1:
                                sp1=slownik_stop_points[sp]
                                SP=Visum.Net.StopPoints.ItemByKey(slownik_stop_points[sp]) 
                                Route.Add(SP)
                                print "dodalem stop point (znalazlem go!)"
                        except:
                            pass
                    
                        
                    
                print "Count: ", Route.Count   
                if Route.Count>1:
                    #try:
                    #routesearchparameters.SearchShortestPath(1, True, True, 0, 1, 1)
                    Visum.Net.AddLineRoute(linia,Line,direction1,Route,routesearchparameters)
                    
    #                
                
            #kier 2
            
            a=a[::-1]
            if not (len(a)==2 and a[0]==a[-1]): 
                Route=0           
                Route=Visum.CreateNetElements()
                flag=False
                for sp in a:                 
                    try:
                        print "przystanek ",sp, slownik_stop_points[sp]
                        slownik_stop_points[sp]
                        Visum.Net.StopPoints.ItemByKey(slownik_stop_points[sp])  
                        Route.Add(Visum.Net.StopPoints.ItemByKey(slownik_stop_points[sp]))
                        flag=True
                    except:
                        pass                    
                if Route.Count>1:
                    Visum.Net.AddLineRoute(linia,Line,direction2,Route,routesearchparameters)
    
def main_przystanki(nsnap=3000, lsnap=3000):
    i=0
    slownik_stops={}
    Iterator=Visum.Net.Territories.Iterator
    while Iterator.Valid:       
        terytorium=Iterator.Item
        stopID=terytorium.AttValue("No")+100000
        
        nazwa=terytorium.AttValue("Name")
        xS = terytorium.AttValue("XCoord")
        yS=terytorium.AttValue("YCoord")
        przystanek=Visum.Net.AddStop(stopID,xS,yS)
        i+=1
        print "===========",i,"|1908 : ",nazwa,"=============="
        
        [Node,res,dist]=Visum.Net.GetNearestNode(xS,yS,nsnap,False) #znajdz najblizszy Node (duzy radius)   
        if res:
            print "    jest Node"            
            xN=Node.AttValue("XCoord")
            yN=Node.AttValue("YCoord")           
            StopArea=Visum.Net.AddStopArea(stopID,przystanek,Node,xN,yN)  
            StopArea.SetAttValue("Name",nazwa)    
            print "    dodalem StopArea" 
            try:           
                StopPoint=Visum.Net.AddStopPointOnNode(stopID,StopArea,Node)
                print "    dodalem StopPoint"             
                StopPoint.SetAttValue("Name",nazwa)
                StopPoint.SetAttValue("TypeNo",8)
            except:
                print "    nie udalo sie dodaje na link"
                [Link,res,dist,xL,yL,rel] = Visum.Net.GetNearestLink(xS,yS,lsnap,False)
                if res: 
                    StopPoint = Visum.Net.AddStopPointOnLink(stopID,StopArea,Link.AttValue("FromNodeNo"),Link.AttValue("ToNodeNo"),False)              
                    print "    dodalem StopPoint"
                    StopPoint.SetAttValue("TypeNo",8)
                    rel=min(0.99,rel)
                    rel=max(0.01,rel)
                    StopPoint.SetAttValue("RelPos",rel)
                    StopPoint.SetAttValue("Name",nazwa)                
        else:
            Node=Visum.Net.AddNode(stopID,xS,yS)
            print "nie ma node, szukam odcinka"
            [Link,res,dist,xL,yL,rel] = Visum.Net.GetNearestLink(xS,yS,lsnap,False)            
            StopArea=Visum.Net.AddStopArea(stopID,przystanek,Node,xS,yS)
            StopArea.SetAttValue("Name",nazwa) 
            if res:
                print "jest odcinek", res, dist                
                StopPoint = Visum.Net.AddStopPointOnLink(stopID,StopArea,Link.AttValue("FromNodeNo"),Link.AttValue("ToNodeNo"),False)              
                StopPoint.SetAttValue("TypeNo",8)
                rel=min(0.99,rel)
                rel=max(0.01,rel)
                StopPoint.SetAttValue("RelPos",rel)
                StopPoint.SetAttValue("Name",nazwa)
                    
            else:    
                print "Nic nie ma"            
                StopPoint=Visum.Net.AddStopPointOnNode(stopID,StopArea,Node)
                StopPoint.SetAttValue("TypeNo",8)
                print "Robie Stop w Polu"
                StopPoint.SetAttValue("Name",nazwa)
                flag = True 
                
        Iterator.Next()              


#try:
#    Visum 
#except:        
#    Visum=VisumInit("C:\stops_bckp_2000.ver")     
#main_przystanki() 
slownik_linie= odczytaj_xls() 
#slownik_stop_points=get_stops()
#Add_Line(slownik_linie,slownik_stop_points)
i=0
for linia in slownik_linie:     
    i+=1           
    if i in [500,1146,1922,1987]:
        print linia, slownik_linie[linia] 
#    #Line.SetAttValue("rozklad_pdf",linia)
#    przystanki=slownik_linie[linia]
#    a= przystanki.split("-")
#    a=[e.strip() for e in a]
#    
#    for sp in a:
#        try:
#            slownik_stop_points[sp]
#        except:
#            print sp
    

    

    
    
   
#Visum.SaveVersion("C:\stops_sro.ver")