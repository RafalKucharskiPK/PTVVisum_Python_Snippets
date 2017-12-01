"""
 _   ______  
| | /___   \     Intelligent Infrastructure
| |  ___|  |     script created by: Rafal Kucharski
| | /  ___/      16/07/2012
| | | |___       info: info@intelligent-infrastructure.eu
|_| |_____|      Copyright (c) Intelligent Infrastructure 2012 

=====================
Dependencies: 
wx
=====================
 
==========================
End-User License Agreement:
===========================
THERE IS NO WARRANTY FOR THE PROGRAM, TO THE EXTENT PERMITTED BY APPLICABLE LAW. 
EXCEPT WHEN OTHERWISE STATED IN WRITING THE COPYRIGHT HOLDERS AND/OR OTHER PARTIES 
PROVIDE THE PROGRAM 'AS IS' WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING, 
BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE. 
THE ENTIRE RISK AS TO THE QUALITY AND PERFORMANCE OF THE PROGRAM IS WITH YOU. 
SHOULD THE PROGRAM PROVE DEFECTIVE, YOU ASSUME THE COST OF ALL NECESSARY SERVICING, REPAIR OR CORRECTION.

This software is created by Intelligent-Infrastructure - Rafal Kucharski (i2) Krakow Polska, who also owns the copyrights. 

By using this software you agree with terms stated below:

1.You can use the software only if You bought it from intelligent-infrastructure, or got written permission of i2 to do so.
2.You can use and modify the software code, as long as you don't sell it's parts commercially.
3.You cannot publish and/or show any parts of the code to third-party users without written permission of i2 
4.If You want to sell the software created by modifying this software, you need to contact with i2 and agree conditions
5.This is one user copy, you cannot use it on multiple computers without written permission to do so
6.You cannot modify this statement
7.You can freely analyze the code, and propose any changes
8.Parts of this code cannot be used to any other software creating without written permission of i2

July 2012, Krakow Poland
"""
import wx
import Aggregator as Agg_module

class Agg_GUI(Agg_module.MyDialog):
    def __init__(self,V):
        Agg_module.MyDialog.__init__(self,V)

stand_alone=True
try:
    Visum
    stand_alone=0
except:
    Visum=Agg_module.Visum_Init("D:/agr.ver")

if __name__ == "__main__":
    if stand_alone:
        app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    APNR = Agg_GUI(Visum)
    app.SetTopWindow(APNR)
    APNR.Show()
    if stand_alone:
        app.MainLoop()
