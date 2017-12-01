"""
 _   _____  
| | /___  \     Intelligent Infrastructure
| |  ___| |     script created by: Rafal Kucharski
| | /  ___/     16/09/2011
| | | |___      info: info@intelligent-infrastructure.eu
|_| |_____|     Copyright (c) Intelligent Infrastructure 2011 


=====================
Dependencies: shapely
=====================
 
===========================
End-User License Agreement:
===========================

This software is created and copyrights are owned by Intelligent-Infrastructure - Rafal Kucharski (i2) Krakow Polska, and to by using it you agree with terms stated below:
1.You can use the software only if You bought it from intelligent-infrastructure, or got an written permission of i2 to do so.
2.You can use and modify the software code, as long as you don't sell it's parts commercially.
3.You cannot publish and/or show any parts of the code to third-party users without written permission of i2 
4.If You want to sell the software created by modifying this software, you need to contact with i2 and agree conditions
5.This is one user copy, you cannot use it on multiple computers without written permission to do so
6.You cannot modify this statement
7.You can freely analyse the code, and propose any changes
8. After period defined by special i2 statement this software becomes freeware, so that it can be freely downloaded and/or modified.

"""
try:
    from shapely.wkt import loads    
except:
    import os,wx
    wx.MessageBox(("You need a 'shapely' library (ca.100kB) to use this script closing this box will start downloading process for shapely."), ("Error"), style=wx.ICON_ERROR)
    os.startfile("http://find_it_rafal!!!!!!")
    


OnlyActive=False

POIIterator=Visum.Net.POICategories.Iterator


while POIIterator.Valid:
    POIc=POIIterator.Item
    Polygons=POIc.POIs.GetMultiAttValues("WKTSurface",OnlyActive)    
    Areas=[]
    i=1
    for Polygon in Polygons:
        Polyg=loads(str(Polygon[1]))
        Areas.append((i, Polyg.area))
        i+=1    
      
    try: 
        POIc.POIs.AddUserDefinedAttribute('Area', 'Area', 'Area', 2,10) 
    except: 
        pass  
        
    POIc.POIs.SetMultiAttValues("Area",Areas)    
    
    POIIterator.Next()
POIIterator.Reset()






