'''
Created on 15-09-2012

@author: i2
'''
import googlemaps


PKWs=[ "ul.Dluga 40/42 , Konstancin-Jeziorna",
"ul. Wojewodzka 12 , Konstancin-Jeziorna",
"ul. Zeromskiego 15 , Konstancin-Jeziorna",
 "ul.J.Sobieskiego 6 , Konstancin-Jeziorna",
 "ul. Wilanowska 1 , Konstancin-Jeziorna",
 "ul.Jaworskiego 18 , Konstancin-Jeziorna",
 "ul.Bielawska 57 , Konstancin-Jeziorna",
"ul. Szkolna 7 , Konstancin-Jeziorna", 
 "ul. K.Pulaskiego 72 , Konstancin-Jeziorna",
"ul. Wspolna 1/3 , Bielawa",
"ul. Rycerska 13 , Konstancin-Jeziorna"]
g=googlemaps.GoogleMaps()
i=0
coords=[]
for PKW in PKWs:
    i+=1
    a=g.address_to_latlng(PKW)
    coords.append("POINT("+str(a[0])+" "+str(a[1])+")")
    Visum.Net.AddZone(i,str(a[1]),str(a[0]))
    


    
    
