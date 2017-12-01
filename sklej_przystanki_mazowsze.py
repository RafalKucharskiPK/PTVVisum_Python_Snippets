# Procedura do skelajania odklejonej warstwy stops do sieci
#parametry procedury
przetnij_w_przystanku= True
nsnap = 100
lsnap = 100
drugi_kierunek=True


# dwa sposoby wyswietlanaia bledu (wx msg, albo print w konsoli w zaleznosci od zaimprotowanych bibliotek
try:
    wx
    wx_=True
except:
    wx_=False
    

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
    
def errmsg(msg):
    if wx_:
        wx.MessageBox(msg, "Blad skryptu importowania przystnakow", style=wx.OK | wx.ICON_ERROR)
    else:
        print msg
            
def main(przetnij_w_przystanku=True,nsnap=100,lsnap=100,i=100000):
            
    Iterator=Visum.Net.Stops.Iterator
    while Iterator.Valid:       
        przystanek=Iterator.Item
        #definicje
        i+=1
        flag=False   
        if True:                
            stopID=przystanek.AttValue("No")
            xS = przystanek.AttValue("XCoord")
            yS=przystanek.AttValue("YCoord")
            print "========================"
            flag = False         
            
            [Node,res,dist]=Visum.Net.GetNearestNode(xS,yS,10*nsnap,False) #znajdz najblizszy Node (duzy radius)   
            print stopID, dist, res  
            if res:  #jesli znalazles cokolwiek           
                  
                #sprawdz czy StopArea jest juz na Node
                NodeStopAreas=Node.AttValue("CONCATENATE:STOPAREAS\No") 
                xN=Node.AttValue("XCoord")
                yN=Node.AttValue("YCoord")
                print NodeStopAreas         
                if len(NodeStopAreas)==0:
                       print "Nowy StopArea"
                       StopArea=Visum.Net.AddStopArea(stopID,przystanek,Node,xN,yN)
                else:
                       print "Stary StopArea"                       
                       StopArea=Visum.Net.StopAreas.ItemByKey(NodeStopAreas)
                
                
                #jesli Node jest blisko                      
                if dist<=nsnap:  
                    print dist, "<", nsnap
                    #sporobuj dodac na node               
                    try:
                        StopPoint=Visum.Net.AddStopPointOnNode(stopID,StopArea,Node)
                        StopPoint.SetAttValue("TypeNo",5)
                        print "Nowy StopPoint: ", stopID
                        flag= True
                    except: 
                        print "Nie udalo sie"
                        flag = False
                        
                if dist>nsnap or flag == False: 
                    print "Szukam odcinka"            
                    [Link,res,dist,xL,yL,rel] = Visum.Net.GetNearestLink(xS,yS,lsnap,False)
                    print "OdcineK: ", res, dist                
                    if res:                
                       if dist<lsnap: 
                           try:
                               StopPoint = Visum.Net.AddStopPointOnLink(stopID,StopArea,Link.AttValue("FromNodeNo"),Link.AttValue("ToNodeNo"),True)              
                               StopPoint.SetAttValue("TypeNo",6)
                               StopPoint.SetAttValue("RelPos",rel)
                               if drugi_kierunek:
                                   StopPoint = Visum.Net.AddStopPointOnLink(stopID,StopArea,Link.AttValue("ToNodeNo"),Link.AttValue("FromNodeNo"),True)              
                                   StopPoint.SetAttValue("TypeNo",6)
                                   StopPoint.SetAttValue("RelPos",1-rel)
                               flag=True                        
                               print "Dodalem StopPoint OnLink"
                           except:
                               print "Nie da sie na nim dodac"
                               flag=False
            if not flag:
                print "nic nie ma w poblizu"
                [Link,res,dist,xL,yL,rel] = Visum.Net.GetNearestLink(xS,yS,5*lsnap,False)
                Node=Visum.Net.AddNode(stopID*100,xS,yS)
                try:
                    StopArea=Visum.Net.AddStopArea(stopID,przystanek,Node,xS,yS)
                except:
                    pass
                if res:
                    print "O, jednak jest odcinek", res, dist 
                    try:
                        StopPoint = Visum.Net.AddStopPointOnLink(stopID,StopArea,Link.AttValue("FromNodeNo"),Link.AttValue("ToNodeNo"),True)              
                        StopPoint.SetAttValue("TypeNo",7)
                        try:
                            StopPoint.SetAttValue("RelPos",rel)
                        except:
                            pass
                        print "Dodalem StopPoint OnLink"
                        flag=True
                    except:
                        print "Ale nie da sie na nim dodac"
                        flag = False
                else:    
                    print "Nic nie ma"            
                    StopPoint=Visum.Net.AddStopPointOnNode(stopID,StopArea,Node)
                    StopPoint.SetAttValue("TypeNo",8)
                    print "Robie Stop w Polu"
                    flag = True  
            
            StopPoint                                        
        Iterator.Next()  
try:
    Visum 
except:        
    Visum=VisumInit("C:\stops.ver")     
main() 
              
#Visum.SaveVersion("C:\stops2.ver")


"""
Stop point type:
5 - w poblizu wezla, na wezle
6 - w poblizu odcinka na odcinku, stop area na wezle sieci
7 - w poblizu odcinka na odcinku, stop area na nowym nodzie 
8 - kompletnie nie powiazany z siecia
"""
    
    