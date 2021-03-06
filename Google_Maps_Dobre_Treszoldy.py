#!/usr/bin/env python
# -*- coding: iso-8859-15 -*-
# generated by wxGlade 0.6.3 on Mon Jan 24 17:19:11 2011

import sys, os, signal, wx, Image
import win32com.client
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4.QtWebKit import QWebPage

Visum=win32com.client.Dispatch('Visum.Visum')



# begin wxGlade: extracode
# end wxGlade


class WxApp(wx.Frame):
    
    def __init__(self, *args, **kwds):
        
        # begin wxGlade: WxApp.__init__
        
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.label_1 = wx.StaticText(self, -1, "")
        self.label_2 = wx.StaticText(self, -1, "")
        self.label_3 = wx.StaticText(self, -1, "")
        self.label_4 = wx.StaticText(self, -1, "")
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.OnTimer, self.timer)
        self.timer.Start(4500) #### sprawdzanie co 4,5 sekundy (do ustalenia)        
        self.__set_properties()
        self.__do_layout()
        self.Initials()
        
        
        # end wxGlade

    def __set_properties(self):
        # begin wxGlade: WxApp.__set_properties
        self.SetTitle("Google Map Add-On")
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: WxApp.__do_layout
        sizer_1 = wx.BoxSizer(wx.VERTICAL)
        sizer_1.Add(self.label_1, 0, 0, 0)
        sizer_1.Add(self.label_2, 0, 0, 0)
        sizer_1.Add(self.label_3, 0, 0, 0)
        sizer_1.Add(self.label_4, 0, 0, 0)
        self.SetSizer(sizer_1)
        sizer_1.Fit(self)
        self.Layout()
        
                
        # end wxGlade
    def OnTimer(self, event): ### to sie dzieje co 4,5 sekundy        
                        
         
        self.pozycja_stara=self.pozycja ## zapisanie starej pozycji okna
        self.pozycja=Visum.Graphic.GetWindow() ##zapisanie nowej pozycji okna
        self.przesuniecie=abs((self.pozycja_stara[0]-self.pozycja[0])/(self.pozycja_stara[2]-self.pozycja_stara[0])) ## abs(xmin stare - xmin )/(xmax stare - xmin stare)
        self.powiekszenie=abs((self.pozycja_stara[2]-self.pozycja_stara[0])/(self.pozycja[2]-self.pozycja[0])-1) ## szerokosc stara / szerokosc nowa - 1
        
        self.render_flag='No Change'
        if self.Visum_Main_Window_Size!=Visum.Graphic.GetMainWindowPos():
            self.get_Network_window_size()
            self.Net_Window_Size=self.get_Network_window_size()
            self.render_flag=self.Render()
        
        
        if  self.przesuniecie>.05 or self.powiekszenie>.05: ## warunek: wiecej niz 5% przesuniecia/powiekszenia           
            self.render_flag=self.Render()            
                        
            
        else: self.label_4.Label='No Change'
        
        self.label_1.Label='Position: '+str(Visum.Graphic.GetWindow())
        self.label_2.Label='Pan shift: '+str(int(100*self.przesuniecie))+'%'
        self.label_3.Label='Scale shift: '+str(int(100*self.powiekszenie))+'%'
        self.label_4.Label=self.render_flag
                
    def Initials(self):
        self.Katalog_Tla=Visum.GetPath(48)
        self.HTML_URL='D:/ptv/ptv.html'
        self.format=".png"
        self.path="D:/ptv/output.png"
        
        self.Visum_Main_Window_Size=Visum.Graphic.GetMainWindowPos()
        self.pozycja=Visum.Graphic.GetWindow()  ## poczatkowa pozycja okna wyrazona jako: xmin ymin xmax ymax 
        self.Net_Window_Size=self.get_Network_window_size()
        self.web = loadImage(self.path)
        
    
    def get_Network_window_size(self):
        path=self.Katalog_Tla+'__del.jpg'
        Visum.Graphic.Screenshot(path)
        return Image.open(path).size
        self.label_4.Label=str(Image.open(path).size)
        os.remove(self.path)    
        
    def Render(self):
        return 'render'
# end of class WxApp

class loadImage():
    def loading(self):
        self.__loading = True
        print("loading")
        
    def loaded(self, result):
        self.__loading = False
        self.__loaded = result
        print("loaded", result)
        if result == True:
            self.webpage.setViewportSize(self.webpage.mainFrame().contentsSize())
            image = QImage(self.webpage.viewportSize(), QImage.Format_ARGB32)
            painter = QPainter(image)
            self.webpage.mainFrame().render(painter)
            painter.end()
            if os.path.exists(self.path):
                os.remove(self.path)
            image.save(self.path)
            print("saved")
    
    def render(self, url):
        print("render")
        self.webpage.mainFrame().load(QUrl(url))
        while self.__loading == True:
            QCoreApplication.processEvents()
    
    def __init__(self, path):
        print("init")
        self.path = path;
        self.webpage = QWebPage()
        self.webpage.connect(self.webpage, SIGNAL("loadStarted()"), self.loading)
        self.webpage.connect(self.webpage, SIGNAL("loadFinished(bool)"), self.loaded)
        self.__loading = False
        self.__loaded = False




if __name__ == "__main__":
    app = wx.PySimpleApp(0)
    
    wx.InitAllImageHandlers()
    frame_1 = WxApp(None, -1, "")
    app.SetTopWindow(frame_1)
    frame_1.Show()
    app.MainLoop()
    sys.exit(0)
