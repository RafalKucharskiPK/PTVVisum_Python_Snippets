"""
 _   _____  
| | /___  \     Intelligent Infrastructure
| |  ___| |     script created by: Rafal Kucharski
| | /  ___/     16/08/2011
| | | |___      info: info@intelligent-infrastructure.eu
|_| |_____|     Copyright (c) Intelligent Infrastructure 2011 

references: OpenLayers team, PyQt - riverside

=====================
Dependencies:
 
1. OpenLayers (www.openlayers.com)
2. PyQt4 by riverside (i2 has a commercial license)
3. Python Imaging Library
=====================
 
==========================
End-User License Agreement:
===========================
This software uses Google Maps, Yahoo Maps, Bing Maps, and Open Street Maps tiles.
It was not meant to harm any Terms of Service.
It is supposed to work as dynamic background to transport modeling network and nothing more. It is not created to download content from any Map Provider. 
Map provider Copyrights should be visible according to provider's conditions.
 
Therefore, you cannot store files downloaded by means of this script. You need to delete them as soon as you've finished working with Visum.
You cannot print the backgrounds, you cannot publish them, you cannot use it outside PTV Visum.
By doing so you will harm third-party Terms of Service, and intelligent-infrastructure is not responsible for such cases.

This software is created by Intelligent-Infrastructure - Rafal Kucharski (i2) Krakow Polska, who also owns the copyrights. 

By using this software you agree with terms stated below:

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


import sys, os, wx
from time import gmtime, strftime
from urllib2 import Request, urlopen, URLError
import PyQt4.QtCore
from PyQt4.QtCore import SIGNAL, QUrl, QObject, QRect, QMetaObject, Qt, QTimer, QEventLoop, QCoreApplication
from PyQt4.QtGui import QImage, QPainter, QApplication, QWidget, QPushButton, QCheckBox, qApp, QComboBox, QLabel, QDoubleSpinBox, QProgressBar, QSlider, QDesktopWidget, QMessageBox
from PyQt4.QtWebKit import QWebPage, QWebSettings
from PIL import Image
#import win32com.client
#Visum=win32com.client.Dispatch('Visum.Visum')
#Visum.LoadVersion('C:/Users/czaaja/workspace/screen/src/wwa.ver')

_progress = 0
_paramGlobal = 0
_windowSizeGlobal = 0

filename = Visum.GetWorkingFolder() + "\\AddIns\\MapLayer\\log_" + strftime("%d_%m_%Y_%H_%M_%S", gmtime()) + ".txt"
f = open(filename, "w")
f.write("Creating paths...\n")
f.close()
Paths={}
Paths["MainVisum"] = Visum.GetWorkingFolder()
Paths["ScriptFolder"] = Paths["MainVisum"] + "\\AddIns\\MapLayer"
Visum.SetPath(48, Paths["ScriptFolder"])
Paths["Html"] = Paths["ScriptFolder"] + "\\template.html"
Paths["Screenshot"] = Paths["ScriptFolder"] + "\\__del.jpg"
Paths["Images"] = []
Paths["Plugins"] = os.path.dirname(PyQt4.QtCore.__file__) + "/plugins"
f = open(filename, "a")
f.write("Paths creation completed.\n")
f.close()

class WebPage(QWebPage):
    def __init__(self, parent = None):
        global f
        f = open(filename, "a")
        f.write("WebPage init...\n")
        f.close()
        QWebPage.__init__(self, parent)
        f = open(filename, "a")
        f.write("WebPage init completed.\n")
        f.close()

class Main():
    def __init__(self, layer, param, multiply, dimension):
        global _progress, f, _paramGlobal, _windowSizeGlobal
        f = open(filename, "a")
        f.write("Main init...\n")
        f.close()
        os.remove(Paths["Screenshot"])
        self._progress = 0
        self._layer = layer
        self._param = param
        self._dimension = dimension
        self._multiply = multiply
        self._windowSize = Visum.Graphic.GetMainWindowPos()
        self._url = Paths["Html"]
        self._image_path = Paths["ScriptFolder"] + "\\" + str(self._param[0]) + "_" + str(self._param[1]) + "_" + str(self._param[2]) + "_" + str(self._param[3]) + "_" + str(self._dimension[0]) + "_" + str(self._dimension[1]) + "_" + str(self._layer) + "_" + str(self._multiply) + ".png"
        self._loaded = False
        self._loading = False
        self._webpage = WebPage()
        _paramGlobal = self._param
        _windowSizeGlobal = self._windowSize
        _progress = 0
        if os.path.exists(self._image_path) and Paths["Images"].count(self._image_path) != 0: _progress = 99
        f = open(filename, "a")
        f.write("Main init finished.\n")
        f.write("Layer: " + str(self._layer) + ".\n")
        f.write("Parameters: " + str(self._param) + ".\n")
        f.write("Dimension: " + str(self._dimension) + ".\n")
        f.write("Multiplier: " + str(self._multiply) + ".\n")
        f.write("Window size: " + str(self._windowSize) + ".\n")
        f.close()
        self.__draw__()

    def __draw__(self):
        global _progress, f
        f = open(filename, "a")
        f.write("Rendering function...\n")
        f.close()
        for i in range(len(Paths["Images"])):
            Visum.Graphic.Backgrounds.ItemByKey(Paths["Images"][i]).Draw = False
        if not os.path.exists(self._image_path) or Paths["Images"].count(self._image_path) == 0:
            f = open(filename, "a")
            f.write("If not os.path.exists or image not in Paths...\n")
            f.close()
            QObject.connect(self._webpage, SIGNAL("loadStarted()"), self.__loading__)
            QObject.connect(self._webpage, SIGNAL("loadProgress(int)"), self.__progress__)
            QObject.connect(self._webpage, SIGNAL("loadFinished(bool)"), self.__loaded__)
            self._webpage.mainFrame().load(QUrl(self._url))
            self._webpage.mainFrame().evaluateJavaScript("x0 = %f; y0 = %f; x1 = %f; y1 = %f;" % (self._param[0], self._param[1], self._param[2], self._param[3]))
            self._webpage.mainFrame().evaluateJavaScript("width = %d; height = %d;" % (self._multiply*self._dimension[0], self._multiply*self._dimension[1]))
            self._webpage.mainFrame().evaluateJavaScript("value = %d;" % (self._layer))
            while self._loading:
                QCoreApplication.processEvents()
            QObject.disconnect(self._webpage, SIGNAL("loadStarted()"), self.__loading__)
            QObject.disconnect(self._webpage, SIGNAL("loadProgress(int)"), self.__progress__)
            QObject.disconnect(self._webpage, SIGNAL("loadFinished(bool)"), self.__loaded__)
            loop = QEventLoop()
            timer = QTimer()
            timer.setSingleShot(True)
            timer.timeout.connect(loop.quit)
            timerProgress = QTimer()
            QObject.connect(timerProgress, SIGNAL("timeout()"), self.__on_loop__)
            if self._layer == 10: timer.start(0)
            elif self._layer >= 4 and self._layer <= 9:
                timer.start(800*self._multiply*4.0)
                timerProgress.start(800*self._multiply*4.0/16.0)
            else:
                timer.start(2400*self._multiply*4.0)
                timerProgress.start(2400*self._multiply*4.0/49.0)
            loop.exec_()
            QObject.disconnect(timerProgress, SIGNAL("timeout()"), self.__on_loop__)
            background_params = self._webpage.mainFrame().evaluateJavaScript("getRealBounds();").toString()
            background_params = background_params[1:-1]
            background_params = [float(s) for s in background_params.split(", ")]
            self._webpage.setViewportSize(self._webpage.mainFrame().contentsSize())
            f = open(filename, "a")
            f.write("Setting up an image file...\n")
            f.close()
            image = QImage(self._webpage.viewportSize(), QImage.Format_RGB32)
            f = open(filename, "a")
            f.write("Starting a painter...\n")
            f.close()
            painter = QPainter(image)
            f = open(filename, "a")
            f.write("Rendering from frame...\n")
            f.write(str(self._webpage.viewportSize()) + "\n")
            f.close()
            self._webpage.mainFrame().render(painter)
            painter.end()
            image.save(self._image_path)
            f = open(filename, "a")
            f.write("Image saved.\n")
            f.close()
            Visum.Graphic.AddBackgroundOnPos(self._image_path, background_params[0], background_params[1], background_params[2], background_params[3])
            Paths["Images"].append(self._image_path)
            f = open(filename, "a")
            f.write("End of if not os.path.exists or image not in Paths.\n")
            f.close()
        else:
            f = open(filename, "a")
            f.write("If os.path.exists and image in Paths...\n")
            f.close()
            index = Paths["Images"].index(self._image_path)
            Visum.Graphic.Backgrounds.ItemByKey(Paths["Images"][index]).Draw = True
            f = open(filename, "a")
            f.write("End of if os.path.exists and image in Paths...\n")
            f.close()
        _progress = 100
        del self._webpage
        f = open(filename, "a")
        f.write("Rendering completed.\n")
        f.close()

    def __progress__(self, progress):
        global _progress, f
        if self._layer == 10: _progress = progress
        elif self._layer >= 4 and self._layer <= 9: _progress = progress*0.83
        else: _progress = progress*0.5
        f = open(filename, "a")
        f.write("End of progress signal: " + str(_progress) + ".\n")
        f.close()

    def __loading__(self):
        global f
        self._loading = True
        f = open(filename, "a")
        f.write("End of loading signal.\n")
        f.close()
    
    def __on_loop__(self):
        global _progress, f
        _progress += 1
        f = open(filename, "a")
        f.write("End of on loop progress signal: " + str(_progress) + "\n")
        f.close()
    
    def __loaded__(self, ok):
        global f
        self._loading = False
        self._loaded = ok
        f = open(filename, "a")
        f.write("End of if loaded signal: " + str(ok) + ".\n") 
        f.close()

class GUI(QWidget):
    def __init__(self, parent = None):
        global f
        f = open(filename, "a")
        f.write("Widget init.\n")
        f.close()
        QWidget.__init__(self, parent, Qt.WindowStaysOnTopHint)
        self.__setup_gui__(self)
        self._flag = False
        self._change = False
        f = open(filename, "a")
        f.write("End of widget init.\n")
        f.close()

    def closeEvent(self, event):
        reply = QMessageBox.question(self, "Confirm", "Are you sure You want to quit?", QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes: event.accept()
        else: event.ignore()
       
    def __setup_gui__(self, Dialog):
        global f
        f = open(filename, "a")
        f.write("Setup of gui.\n")
        f.close()
        Dialog.setObjectName("Dialog")
        Dialog.resize(270, 145) 
        self.setWindowTitle("Map Layer")       
        screen = QDesktopWidget().screenGeometry()
        size = self.geometry()
        self.move((screen.width()-size.width())/2, (screen.height()-size.height())/2)
        self.Render = QPushButton("Render", Dialog)
        self.Render.setGeometry(QRect(85, 90, 100, 25))
        self.Render.setObjectName("Render")
        self.comboBox = QComboBox(Dialog)
        self.comboBox.setGeometry(QRect(100, 34, 115, 18))
        self.comboBox.setEditable(False)
        self.comboBox.setMaxVisibleItems(11)
        self.comboBox.setInsertPolicy(QComboBox.InsertAtBottom)
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItems(["Google Roadmap", "Google Terrain", "Google Satellite", "Google Hybrid", "Yahoo Roadmap", "Yahoo Satellite", "Yahoo Hybrid", "Bing Roadmap", "Bing Satellite", "Bing Hybrid", "Open Street Maps"])
        self.comboBox.setCurrentIndex(10)
        self.label1 = QLabel("Source:", Dialog)
        self.label1.setGeometry(QRect(55, 35, 35, 16))
        self.label1.setObjectName("label1")
        self.slider = QSlider(Dialog)
        self.slider.setOrientation(Qt.Horizontal)
        self.slider.setMinimum(1)
        self.slider.setMaximum(12)
        self.slider.setValue(4)
        self.slider.setGeometry(QRect(110, 61, 114, 16))
        self.label2 = QLabel("Quality: " + str(self.slider.value()), Dialog)
        self.label2.setGeometry(QRect(47, 61, 54, 16))
        self.label2.setObjectName("label2")
        self.doubleSpinBox = QDoubleSpinBox(Dialog)
        self.doubleSpinBox.setGeometry(QRect(160, 5, 40, 20))
        self.doubleSpinBox.setDecimals(0)
        self.doubleSpinBox.setObjectName("doubleSpinBox")
        self.doubleSpinBox.setMinimum(10.0)
        self.doubleSpinBox.setValue(20.0)
        self.doubleSpinBox.setEnabled(False)
        self.checkBox = QCheckBox("Auto refresh", Dialog)
        self.checkBox.setGeometry(QRect(50, 6, 100, 20))
        self.checkBox.setLayoutDirection(Qt.RightToLeft)
        self.checkBox.setObjectName("checkBox")
        self.progressBar = QProgressBar(Dialog)
        self.progressBar.setGeometry(QRect(5, 130, 260, 10))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setTextVisible(False)
        self.progressBar.setObjectName("progressBar")
        self.progressBar.setVisible(False)
        QObject.connect(self.Render, SIGNAL("clicked()"), Dialog.__repaint__)
        QMetaObject.connectSlotsByName(Dialog)
        QObject.connect(self.slider, SIGNAL("valueChanged(int)"), self.__update_slider_label__)
        QObject.connect(self.comboBox, SIGNAL("activated(int)"), self.__combobox_changed__)
        self.timerRepaint = QTimer()
        QObject.connect(self.checkBox, SIGNAL("clicked()"), self.__activate_timer__)
        QObject.connect(self.timerRepaint, SIGNAL("timeout()"), self.__on_timer__)
        f = open(filename, "a")
        f.write("End of setup of gui.\n") 
        f.close()
    
    def __combobox_changed__(self):
        self._change = True
    
    def __activate_timer__(self):
        self.doubleSpinBox.setEnabled(self.checkBox.isChecked())
        if self.checkBox.isChecked():
            self.timerRepaint.start(self.doubleSpinBox.value()*1000)
            self.Render.setEnabled(False)
            if _progress == 0: self.__repaint__()
        else:
            self.timerRepaint.stop()
            self.Render.setEnabled(True)
    
    def __get_net_size__(self):
        global f
        f = open(filename, "a")
        f.write("Geting net size...\n")
        f.close()
        if not os.path.exists(Paths["Screenshot"]):
            Visum.Graphic.Screenshot(Paths["Screenshot"])
        size = Image.open(Paths["Screenshot"]).size
        f = open(filename, "a")
        f.write("Read net size:" + str(size) + ".\n") 
        f.close()
        return size         

    def __on_timer__(self):
        global _paramGlobal
        self._flag = False
        Visum.Graphic.MaximizeNetWindow()
        param = _paramGlobal
        _paramGlobal = Visum.Graphic.GetWindow()
        shift = abs((param[0]-_paramGlobal[0])/(param[2]-param[0]))
        zoom = abs((param[2]-param[0])/(_paramGlobal[2]-_paramGlobal[0])-1)
        print _windowSizeGlobal
        if _windowSizeGlobal[2:4] != Visum.Graphic.GetMainWindowPos()[2:4]:
            self.__get_net_size__()
            self._flag = True
        elif shift > 0.4 or zoom > 0.2: self._flag = True
        if self._flag or self._change and _progress == 0:
            self.__repaint__()
            self._change = False

    def __update_slider_label__(self, value):
        self.label2.setText("Quality: " + str(value))
        self._change = True
    
    def __update_progress_bar__(self):
        if _progress != 0:
            self.progressBar.setVisible(True)
            self.progressBar.setValue(_progress)
        else: self.progressBar.setVisible(False)
    
    def __rebuild_paths__(self):
        global Paths
        Paths["Images"] = []
        list = os.listdir(Paths["ScriptFolder"])
        imageList = []
        for i in range(len(list)):
            if list[i][-3:] == "png": imageList.append(list[i])
        for i in range(len(imageList)):
            try:
                Visum.Graphic.Backgrounds.ItemByKey(imageList[i])
                Paths["Images"].append(Paths["ScriptFolder"] + "\\" + imageList[i])
            except:
                pass
        
    def __repaint__(self):
        global _progress, f
        if len(Visum.Graphic.Backgrounds.GetAll) != len(Paths["Images"]):
            self.__rebuild_paths__()
        if _progress == 0:
            f = open(filename, "a")
            f.write("Doing repaint...\n")
            f.close()
            QWebSettings.clearMemoryCaches()
            timer = QTimer()
            timer.start(100)
            QObject.connect(timer, SIGNAL("timeout()"), self.__update_progress_bar__)
            Main(self.comboBox.currentIndex(), Visum.Graphic.GetWindow(), self.slider.value()/4.0, self.__get_net_size__())
        Visum.Graphic.Draw()
        self.__update_progress_bar__()
        _progress = 0
        QTimer().singleShot(1500, self.__update_progress_bar__)
        f = open(filename, "a")
        f.write("End of doing repaint.\n") 
        f.close()

class Top(QApplication):    
    def __init__(self, args):
        global f
        f = open(filename, "a")
        f.write("Initializing application...\n")
        f.close()
        QApplication.__init__(self, args)
        qApp.addLibraryPath(Paths["Plugins"])
        window = GUI()
        window.show()
        self.exec_()
        f = open(filename, "a")
        f.write("Application exit code.\n") 
        f.close()
    

       
if __name__ == "__main__":
    if Visum.Net.NetParameters.AttValue("PROJECTIONDEFINITION")[8:20] != "GCS_WGS_1984":
        wx.MessageBox(("Please set a coordinate projection for your VISUM model to GCS_WGS_1984 before using Map Layer.\nYou should find appropriate options via Network->Network parameters->Scale"), ("Error"), style=wx.ICON_ERROR)
    
    else: Top(sys.argv)