"""
 _   _____  
| | /___  \     Intelligent Infrastructure
| |  ___| |     script created by: Rafal Kucharski
| | /  ___/     16/08/2011
| | | |___      info: info@intelligent-infrastructure.eu
|_| |_____|     Copyright (c) Intelligent Infrastructure 2011 



=====================
Dependencies: OpenLayers (www.OpenLayers.org), also script uses WKT format.
=====================
 
==========================
End-User License Agreement:
===========================

This software is created and copyrights are owned by Intelligent-Infrastructure - Rafal Kucharski (i2) Krakow Polska, and to by using it you agree with terms stated below:
1.You can use the software only if You bought it from intelligent-infrastructure, or got an written permission of i2 to do so.
2.You can use and modify the software code, as long as you don't sell it's parts commercially.
3.You cannot publish and/or show any parts of the code to third-party users without written permission of i2 
4.If You want to sell the software created by modifying this software, you need to contact with i2 and agree conditions
5.This is one user copy, you cannot use it on multiple computers without written permission to do so
6.You cannot modify this statement
7.You can freely analyze the code, and propose any changes
8. After period defined by special i2 statement this software becomes freeware, so that it can be freely downloaded and/or modified.

sept 2011, Krakow Poland
"""
import os
import wx


#def Init(path=None):
#        import win32com.client        
#        Visum=win32com.client.Dispatch('Visum.Visum')
#        if path!=None: Visum.LoadVersion(path)
#        return Visum
#try: 
#    Visum
#except: 
#    Visum=Init("D:/zmniejszanie.ver")



def Add_WKT_Elements(WKT):
    for WKTel in WKT: 
        if WKTel[1][-5:]!="EMPTY":   
            nline="addFeature('"+ str(WKTel[1])+"');"
            lines.insert(56,nline+"\n")


wx.MessageBox(("Welcome!\n\nNetPublisher brought to you by intelligent-infrastructure creates html file containing:\n\n1.map background (Google Maps or OpenStreetMap) and \n2.Visum network elements: Links, Nodes, Zones, Zone Boundaries, Territories. \n\nResulting file, which you can publish online is:\nVisum_Path/Exe/AddIns/MapPublisher/Visum_Output.html"))
    
templateHTML=open("template.html","r")
lines=templateHTML.readlines()


Add_WKT_Elements(Visum.Net.Links.GetMultiAttValues("WKTPoly"))
Add_WKT_Elements(Visum.Net.Nodes.GetMultiAttValues("WKTLoc"))

Add_WKT_Elements(Visum.Net.Zones.GetMultiAttValues("WKTLoc"))
Add_WKT_Elements(Visum.Net.Zones.GetMultiAttValues("WKTSurface"))

Add_WKT_Elements(Visum.Net.Territories.GetMultiAttValues("WKTLoc"))
Add_WKT_Elements(Visum.Net.Territories.GetMultiAttValues("WKTSurface"))

filename="Visum_Output"  
filename=filename+".html"
newfile=open(filename,"w")
newfile.writelines(lines)
newfile.close()
os.startfile(filename)
    
    
