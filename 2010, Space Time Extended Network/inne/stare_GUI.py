#import matplotlib
from matplotlib.figure import Figure
import wx
import xlwt
from scipy import sqrt,zeros,Inf,sort,var,mean,exp ## math asset
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
import matplotlib.pyplot as plt ## 2d plotting
from mpl_toolkits.mplot3d import Axes3D    ## 3d plotting
import win32com.client ## to dispatch Visum instance in
import VisumPy.helpers ## to get container
import random ## to randomize JourneyStarts
import time

import cPickle

global Visum, DSeg, TimeZero, PlotSpace, fig, Is_Plot, IsCanvas

IsCanvas=0
TimeZero=time.time()



class DataBase:
    class Arrays: pass
    class Zones: pass
    class Nodes: 
        class CR: pass    
    class Links: 
        class CR: pass
    class Turns:
        class CR: pass    
    class OrigConnectors: pass
    class DestConnectors: pass    
    class Paths: pass
    pass

class Results:
    class Paths: pass
    class Times: pass
    pass
    
Results.Times.Zero=TimeZero
Results.Times.Plot_Cylinder=0
DataBase.tCurCalcMethod=2
DataBase.TimeDistribution=0

class MyFrame(wx.Frame):
        
    def __init__(self, *args, **kwds):        
            self.figure = Figure()        
            # begin wxGlade: MyFrame.__init__
            kwds["style"] = wx.CAPTION|wx.CLOSE_BOX|wx.MINIMIZE_BOX|wx.MAXIMIZE|wx.MAXIMIZE_BOX|wx.SYSTEM_MENU|wx.RESIZE_BORDER|wx.NO_BORDER|wx.CLIP_CHILDREN
            wx.Frame.__init__(self, *args, **kwds)
            self.label_1 = wx.StaticText(self, -1, "Space Time Cylinder")
            self.static_line_1 = wx.StaticLine(self, -1)
            self.button_1 = wx.Button(self, -1, "Script")
            self.Load_Button = wx.Button(self, -1, "Load Version")
            self.label_4 = wx.StaticText(self, -1, "Choose Demand Segment")
            self.label_4_copy_1 = wx.StaticText(self, -1, "Choose time distribution")
            self.DSeg_Combo = wx.ComboBox(self, -1, choices=["DSeg", "-"], style=wx.CB_DROPDOWN)
            self.TimeDistribution_Combo = wx.ComboBox(self, -1, choices=["Random", "Evenly distributed"], style=wx.CB_DROPDOWN)
            self.label_4_copy = wx.StaticText(self, -1, "Choose tCur calculation mode")
            self.tCur_Combo = wx.ComboBox(self, -1, choices=["tCur = conts, visum tCur", "tCur = function( Cumulative_Volume[t] , Capacity Restrain Function Type, Parameters", "tCur = function( Vol[t] , trapezoid diagram for links, queieung theory for turns, t0 for nodes"], style=wx.CB_DROPDOWN)
            self.Database_Button = wx.Button(self, -1, "Prepare Database")
            self.Randomize_Time_Button = wx.Button(self, -1, "Randomize the Time")
            self.Calculate_Paths_Button = wx.Button(self, -1, "Calculate the Paths")
            self.filter_element_type = wx.ComboBox(self, -1, choices=["None", "Zone", "Time", "Node", "Turn"], style=wx.CB_DROPDOWN)
            self.lower_bound = wx.TextCtrl(self, -1, "")
            self.upper_bound = wx.TextCtrl(self, -1, "")
            self.ShowCylinder_Button = wx.Button(self, -1, "Show_Cylinder")
            self.label_3 = wx.StaticText(self, -1, "Node/FromNode/FromNode")
            self.label_3_copy = wx.StaticText(self, -1, "-/ToNode/ViaNode")
            self.label_3_copy_copy = wx.StaticText(self, -1, "-/-/ToNode")
            self.Node1 = wx.TextCtrl(self, -1, "10", style=wx.TE_PROCESS_TAB)
            self.Node2 = wx.TextCtrl(self, -1, "11", style=wx.TE_PROCESS_TAB)
            self.Node3 = wx.TextCtrl(self, -1, "", style=wx.TE_PROCESS_TAB)
            self.Plot_Charactersitics = wx.Button(self, -1, "Plot")
            self.checkbox_1 = wx.CheckBox(self, -1, "Auto DataBase Save")
            self.SaveDatabase_Button = wx.Button(self, -1, "Save Database")
            self.Load_Database_Button = wx.Button(self, -1, "Load Database")
            self.Save_Results_Button = wx.Button(self, -1, "Excel Report")
            self.console = wx.StaticText(self, -1, "")

            self.__set_properties()
            self.__do_layout()

            self.Bind(wx.EVT_BUTTON, self.Script, self.button_1)
            self.Bind(wx.EVT_BUTTON, self.Load_Version_Handler, self.Load_Button)
            self.Bind(wx.EVT_COMBOBOX, self.DSeg_CHange_Handler, self.DSeg_Combo)
            self.Bind(wx.EVT_COMBOBOX, self.TimeDistribution_Handler, self.TimeDistribution_Combo)
            self.Bind(wx.EVT_COMBOBOX, self.tCur_change_Handler, self.tCur_Combo)
            self.Bind(wx.EVT_BUTTON, self.Generate_Database_Handler, self.Database_Button)
            self.Bind(wx.EVT_BUTTON, self.Randomize_The_Time_Handler, self.Randomize_Time_Button)
            self.Bind(wx.EVT_BUTTON, self.Calculate_Main__Handler, self.Calculate_Paths_Button)
            self.Bind(wx.EVT_BUTTON, self.Show_Cylinder_Handler, self.ShowCylinder_Button)
            self.Bind(wx.EVT_BUTTON, self.Plot_Char_Handler, self.Plot_Charactersitics)
            self.Bind(wx.EVT_BUTTON, self.Save_DataBase_Handler, self.SaveDatabase_Button)
            self.Bind(wx.EVT_BUTTON, self.LoadDatabase_Handler, self.Load_Database_Button)
            self.Bind(wx.EVT_BUTTON, self.ExcelExport_Handler, self.Save_Results_Button)
            # end wxGlade

            self.__set_properties()
            self.__do_layout()

            
            self.Bind(wx.EVT_BUTTON, self.Load_Version_Handler, self.Load_Button)
            self.Bind(wx.EVT_COMBOBOX, self.DSeg_CHange_Handler, self.DSeg_Combo)
            self.Bind(wx.EVT_COMBOBOX, self.TimeDistribution_Handler, self.TimeDistribution_Combo)
            self.Bind(wx.EVT_COMBOBOX, self.tCur_change_Handler, self.tCur_Combo)
            self.Bind(wx.EVT_BUTTON, self.Generate_Database_Handler, self.Database_Button)
            self.Bind(wx.EVT_BUTTON, self.Randomize_The_Time_Handler, self.Randomize_Time_Button)
            self.Bind(wx.EVT_BUTTON, self.Calculate_Main__Handler, self.Calculate_Paths_Button)
            self.Bind(wx.EVT_BUTTON, self.Show_Cylinder_Handler, self.ShowCylinder_Button)
            self.Bind(wx.EVT_BUTTON, self.Plot_Char_Handler, self.Plot_Charactersitics)
            self.Bind(wx.EVT_BUTTON, self.Save_DataBase_Handler, self.SaveDatabase_Button)
            self.Bind(wx.EVT_BUTTON, self.LoadDatabase_Handler, self.Load_Database_Button)
            self.Bind(wx.EVT_BUTTON, self.ExcelExport_Handler, self.Save_Results_Button)

    def __set_properties(self):
            # begin wxGlade: MyFrame.__set_properties
            self.SetTitle("Space Time Cylinder")
            self.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
            self.label_1.SetMinSize((-1, -1))
            self.label_1.SetFont(wx.Font(20, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
            self.Load_Button.SetMinSize((95, -1))
            self.DSeg_Combo.SetMinSize((150, 24))
            self.DSeg_Combo.SetSelection(-1)
            self.TimeDistribution_Combo.SetMinSize((150, -1))
            self.TimeDistribution_Combo.SetSelection(0)
            self.tCur_Combo.SetMinSize((450, 24))
            self.tCur_Combo.SetSelection(1)
            self.Database_Button.SetMinSize((150, -1))
            self.Database_Button.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, "MS Shell Dlg 2"))
            self.Randomize_Time_Button.SetMinSize((150, -1))
            self.Randomize_Time_Button.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, "MS Shell Dlg 2"))
            self.Calculate_Paths_Button.SetMinSize((150, -1))
            self.Calculate_Paths_Button.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, "MS Shell Dlg 2"))
            self.filter_element_type.SetMinSize((50, 24))
            self.filter_element_type.SetSelection(0)
            self.lower_bound.SetMinSize((50, 24))
            self.upper_bound.SetMinSize((50, -1))
            self.ShowCylinder_Button.SetMinSize((50, -1))
            self.ShowCylinder_Button.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, "MS Shell Dlg 2"))
            self.label_3.SetMinSize((-1,-1))
            self.label_3_copy.SetMinSize((-1,-1))
            self.label_3_copy_copy.SetMinSize((-1,-1))
            self.console.SetFont(wx.Font(8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, "MS Shell Dlg 2"))
            # end wxGlade
        

    def __do_layout(self):
        global IsCanvas
        if IsCanvas==0:
            
            # begin wxGlade: MyFrame.__do_layout
            sizer_7 = wx.BoxSizer(wx.HORIZONTAL)
            sizer_1 = wx.BoxSizer(wx.VERTICAL)
            sizer_4 = wx.BoxSizer(wx.VERTICAL)
            sizer_8 = wx.BoxSizer(wx.VERTICAL)
            sizer_9 = wx.BoxSizer(wx.HORIZONTAL)
            sizer_6_copy = wx.BoxSizer(wx.HORIZONTAL)
            sizer_5 = wx.BoxSizer(wx.HORIZONTAL)
            sizer_11 = wx.BoxSizer(wx.HORIZONTAL)
            sizer_10 = wx.BoxSizer(wx.VERTICAL)
            sizer_6 = wx.BoxSizer(wx.VERTICAL)
            sizer_12 = wx.BoxSizer(wx.VERTICAL)
            grid_sizer_1 = wx.GridSizer(2, 2, 0, 0)
            sizer_1.Add(self.label_1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 50)
            sizer_1.Add(self.static_line_1, 0, wx.EXPAND, 0)
            sizer_1.Add(self.button_1, 0, 0, 0)
            sizer_1.Add(self.Load_Button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
            grid_sizer_1.Add(self.label_4, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            grid_sizer_1.Add(self.label_4_copy_1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            grid_sizer_1.Add(self.DSeg_Combo, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            grid_sizer_1.Add(self.TimeDistribution_Combo, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            sizer_6.Add(grid_sizer_1, 1, wx.EXPAND, 1)
            sizer_12.Add(self.label_4_copy, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            sizer_12.Add(self.tCur_Combo, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            sizer_6.Add(sizer_12, 1, wx.EXPAND, 0)
            sizer_10.Add(sizer_6, 1, wx.EXPAND, 0)
            sizer_1.Add(sizer_10, 1, wx.EXPAND, 0)
            sizer_4.Add(self.Database_Button, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 8)
            sizer_4.Add(self.Randomize_Time_Button, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 8)
            sizer_4.Add(self.Calculate_Paths_Button, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 8)
            sizer_11.Add(self.filter_element_type, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_11.Add(self.lower_bound, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_11.Add(self.upper_bound, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_11.Add(self.ShowCylinder_Button, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_4.Add(sizer_11, 1, wx.EXPAND, 0)
            sizer_5.Add(self.label_3, 1, wx.ALL, 2)
            sizer_5.Add(self.label_3_copy, 1, wx.ALL, 2)
            sizer_5.Add(self.label_3_copy_copy, 1, wx.ALL, 2)
            sizer_4.Add(sizer_5, 1, wx.EXPAND, 0)
            sizer_6_copy.Add(self.Node1, 1, wx.ALL, 3)
            sizer_6_copy.Add(self.Node2, 1, wx.ALL, 3)
            sizer_6_copy.Add(self.Node3, 1, wx.ALL, 3)
            sizer_6_copy.Add(self.Plot_Charactersitics, 1, wx.ALL, 3)
            sizer_4.Add(sizer_6_copy, 1, wx.EXPAND, 0)
            sizer_9.Add(self.checkbox_1, 0, 0, 0)
            sizer_9.Add(self.SaveDatabase_Button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_9.Add(self.Load_Database_Button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_9.Add(self.Save_Results_Button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_8.Add(sizer_9, 1, wx.EXPAND, 0)
            sizer_4.Add(sizer_8, 1, wx.EXPAND, 0)
            sizer_4.Add(self.console, 0, 0, 0)
            sizer_1.Add(sizer_4, 1, wx.EXPAND, 0)
            sizer_7.Add(sizer_1, 1, wx.ALL|wx.EXPAND, 1)
            self.SetSizer(sizer_7)
            sizer_7.Fit(self)
            self.Layout()
            self.Centre()
            # end wxGlade
        
        else:
            self.canvas = FigureCanvas(self, -1, self.figure)
            sizer_7 = wx.BoxSizer(wx.HORIZONTAL)
            # begin wxGlade: MyFrame.__do_layout
            sizer_7 = wx.BoxSizer(wx.HORIZONTAL)
            sizer_1 = wx.BoxSizer(wx.VERTICAL)
            sizer_4 = wx.BoxSizer(wx.VERTICAL)
            sizer_8 = wx.BoxSizer(wx.VERTICAL)
            sizer_9 = wx.BoxSizer(wx.HORIZONTAL)
            sizer_6_copy = wx.BoxSizer(wx.HORIZONTAL)
            sizer_5 = wx.BoxSizer(wx.HORIZONTAL)
            sizer_11 = wx.BoxSizer(wx.HORIZONTAL)
            sizer_10 = wx.BoxSizer(wx.VERTICAL)
            sizer_6 = wx.BoxSizer(wx.VERTICAL)
            sizer_12 = wx.BoxSizer(wx.VERTICAL)
            grid_sizer_1 = wx.GridSizer(2, 2, 0, 0)
            sizer_1.Add(self.label_1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 50)
            sizer_1.Add(self.static_line_1, 0, wx.EXPAND, 0)
            sizer_1.Add(self.button_1, 0, 0, 0)
            sizer_1.Add(self.Load_Button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
            grid_sizer_1.Add(self.label_4, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            grid_sizer_1.Add(self.label_4_copy_1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            grid_sizer_1.Add(self.DSeg_Combo, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            grid_sizer_1.Add(self.TimeDistribution_Combo, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            sizer_6.Add(grid_sizer_1, 1, wx.EXPAND, 1)
            sizer_12.Add(self.label_4_copy, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            sizer_12.Add(self.tCur_Combo, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 2)
            sizer_6.Add(sizer_12, 1, wx.EXPAND, 0)
            sizer_10.Add(sizer_6, 1, wx.EXPAND, 0)
            sizer_1.Add(sizer_10, 1, wx.EXPAND, 0)
            sizer_4.Add(self.Database_Button, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 8)
            sizer_4.Add(self.Randomize_Time_Button, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 8)
            sizer_4.Add(self.Calculate_Paths_Button, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 8)
            sizer_11.Add(self.filter_element_type, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_11.Add(self.lower_bound, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_11.Add(self.upper_bound, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_11.Add(self.ShowCylinder_Button, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_4.Add(sizer_11, 1, wx.EXPAND, 0)
            sizer_5.Add(self.label_3, 1, wx.ALL, 2)
            sizer_5.Add(self.label_3_copy, 1, wx.ALL, 2)
            sizer_5.Add(self.label_3_copy_copy, 1, wx.ALL, 2)
            sizer_4.Add(sizer_5, 1, wx.EXPAND, 0)
            sizer_6_copy.Add(self.Node1, 1, wx.ALL, 3)
            sizer_6_copy.Add(self.Node2, 1, wx.ALL, 3)
            sizer_6_copy.Add(self.Node3, 1, wx.ALL, 3)
            sizer_6_copy.Add(self.Plot_Charactersitics, 1, wx.ALL, 3)
            sizer_4.Add(sizer_6_copy, 1, wx.EXPAND, 0)
            sizer_9.Add(self.checkbox_1, 0, 0, 0)
            sizer_9.Add(self.SaveDatabase_Button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_9.Add(self.Load_Database_Button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_9.Add(self.Save_Results_Button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 3)
            sizer_8.Add(sizer_9, 1, wx.EXPAND, 0)
            sizer_4.Add(sizer_8, 1, wx.EXPAND, 0)
            sizer_4.Add(self.console, 0, 0, 0)
            sizer_1.Add(sizer_4, 1, wx.EXPAND, 0)
            sizer_7.Add(sizer_1, 1, wx.ALL|wx.EXPAND, 1)
            self.SetSizer(sizer_7)
            sizer_7.Fit(self)
            self.Layout()
            self.Centre()
            # end wxGlade
            sizer_7.Add(self.canvas, 1, wx.RIGHT | wx.TOP | wx.GROW)
                      
            self.Layout()
        
    def add_canvas(self):
        global IsCanvas
        IsCanvas=1
        self.__do_layout()
            

    def update_console(self,label):
        update_console_ext(self,label)
       
    def Open_Visum_Handler(self): # wxGlade: MyFrame.<event_handler>        
       Visum=LoadVisum()
       self.update_console("Visum Instance Dispatched") 
   
    def Load_Version_Handler(self, event):# wxGlade: MyFrame.<event_handler>       
       
       self.dirname = ''
       dlg = wx.FileDialog(self, "Choose a file", self.dirname,"", "*.ver*", wx.OPEN)
       dlg.SetFilename('makenet.ver')
       dlg.SetDirectory('D:/')
       if dlg.ShowModal()==wx.ID_OK:           
            self.path=dlg.GetPath()
       dlg.Destroy()
       self.Open_Visum_Handler()       
       VerLoad(self.path)
       Results.VersionName=self.path            
       label=self.path+" version loaded"
       self.update_console(label)
       self.DSeg_Combo.Clear()       
       Segments=Visum.Net.DemandSegments.GetMultiAttValues("code")
       self.DSeg_Combo.AppendItems([str(Segments[s][1]) for s in range(len(Segments))])
       del Segments
       Results.Times.VersionLoad=time.time()-TimeZero

    
        
        
    def Generate_Database_Handler(self, event): # wxGlade: MyFrame.<event_handler>        
       self.update_console("DataBase generating")
       Get_Assignment_XML_Data()
       self.update_console("XML Assignment Params Specified")     
       error=Get_VDF_Data()
       
       if error!=None:
            print error       
            self.update_console("Unsupported CR function")   
            return 
       Get_Data()   
       self.update_console("Data import from Visum")
       Results.Times.GetData=time.time()-TimeZero
       Get_All_Path_Coords()
       self.update_console("2d Path Coords calculated")
       Results.Times.GetCoords=time.time()-TimeZero
       

    def Randomize_The_Time_Handler(self, event): # wxGlade: MyFrame.<event_handler>
        Set_times_to_Paths()
        self.update_console("Deparature Time for each journey specified")

    def Calculate_Main__Handler(self, event):# wxGlade: MyFrame.<event_handler>
        Cross_The_Time(self)
        Results.Times.CrossTheTime=time.time()-TimeZero
        
    def Show_Cylinder_Handler(self, event): # wxGlade: MyFrame.<event_handler>       
       filter_type=self.filter_element_type.GetSelection()
       lower_bound=self.lower_bound.Label
       upper_bound=self.upper_bound.Label
       Plot_Cylinder(filter_type,lower_bound,upper_bound)
       Results.Times.Plot_Cylinder=time.time()-TimeZero
    
    def ExcelExport_Handler(self, event): # wxGlade: MyFrame.<event_handler>
        #Excel=Generate_Excel_Report()
        Result_Calculation()
        Generate_xlwt_Report(str('D:/results/simulation_'+str(random.randint(0,1000000))+'.xls'))
    
    def Save_DataBase_Handler(self, event,*filename): # wxGlade: MyFrame.<event_handler>
        if self.checkbox_1.IsChecked()==1:
            a=Results.VersionName
            
            #filename='D:/results/'+a[a.find('/')+1:-a.find('/')-2]+Results.AssType[:3]+str(random.randint(0,100))+Results.Distribution+".lft"
        else:
            dlg = wx.FileDialog(self, "Choose a file",self.dirname, "", "*.*", wx.OPEN)            
            dlg.SetDirectory('D:/results')
        if dlg.ShowModal()==wx.ID_OK:
            filename=dlg.GetPath()
            dlg.Destroy()
        Save_DataBase(filename)

    def LoadDatabase_Handler(self, event): # wxGlade: MyFrame.<event_handler>
        dlg = wx.FileDialog(self, "Choose a file", "","D:/results", "*.*", wx.OPEN)        
        dlg.SetDirectory('D:/results')
        if dlg.ShowModal()==wx.ID_OK:
            filename=dlg.GetPath()
            dlg.Destroy()
        Load_DataBase(filename)

    def DSeg_CHange_Handler(self, event): # wxGlade: MyFrame.<event_handler>
        Key=self.DSeg_Combo.GetValue()
        Dseg=DemandSegmentChoice(Key)

    def tCur_change_Handler(self, event): # wxGlade: MyFrame.<event_handler>
        DataBase.tCurCalcMethod=self.tCur_Combo.GetSelection()
        

    def TimeDistribution_Handler(self, event): # wxGlade: MyFrame.<event_handler>
        DataBase.TimeDistribution=self.TimeDistribution_Combo.GetSelection()
    
    def Clear_Plot_Handler(self, event): # wxGlade: MyFrame.<event_handler>
        global IsCanvas
        IsCanvas=0              
        self.__do_layout()
    
    def Script(self, event): # wxGlade: MyFrame.<event_handler>
        global DSeg
        
        def main_calc(noiter): # ile symulacji noiter=1 -> 1 symulacja
            for i in range(1,noiter+1):
                filename="D:/results/file_"+str(no)+"_simno_"+str(i)+"_tCur_"+str(DataBase.tCurCalcMethod)+"timeDist_"+str(DataBase.TimeDistribution)
                Get_Data() # zbierz dane
                Get_All_Path_Coords() # zbierz paths
                Set_times_to_Paths() # ustal paths
                Cross_The_Time(self) # policz
                Results.Times.CrossTheTime=time.time()-TimeZero
                Result_Calculation() # policz raport
                Generate_xlwt_Report(filename+".xls")    # xls            
                Save_DataBase(filename+".lft") # DB file
                               
        for no in range(13,15):
           
           LoadVisum()
           Visum.LoadVersion("D:/results/wersje/"+str(no)+".ver")
           Results.Times.VersionLoad=time.time()-TimeZero
           Results.VersionName="D:/results/wersje/"+str(no)+".ver"
           DSeg=Visum.Net.DemandSegments.ItemByKey("C")
           Results.DSeg="C"
           Get_Assignment_XML_Data()           
           Get_VDF_Data()            
           Results.Times.GetData=time.time()-TimeZero           
           Results.Times.GetCoords=time.time()-TimeZero 
          # DataBase.TimeDistribution=0
          # DataBase.tCurCalcMethod=0
          # main_calc(1)
          # DataBase.tCurCalcMethod=1
          # main_calc(5)
          # DataBase.tCurCalcMethod=2
          # main_calc(5)
           DataBase.TimeDistribution=1           
           DataBase.tCurCalcMethod=0
           main_calc(1)
           DataBase.tCurCalcMethod=1
           main_calc(1)
           DataBase.tCurCalcMethod=2
           main_calc(1)
           
           
           

# end of class MyFrame

    def Plot_Char_Handler(self, event): # wxGlade: MyFrame.<event_handler>
        self.add_canvas()
        if self.Node1.IsEmpty():
            self.update_console("Specify at least one Node")
        elif self.Node2.IsEmpty():
             Node1=self.Node1.Value
             self.DataBasePlot(int(Node1))
             self.update_console("Ploting characteristics for Node: "+str(Node1))
        elif self.Node3.IsEmpty():
             Node1=self.Node1.Value
             Node2=self.Node2.Value
             self.DataBasePlot(int(Node1),int(Node2))
             self.update_console("Ploting characteristics for Link: "+str(Node1)+" to "+str(Node2))
        else:
             Node1=self.Node1.Value
             Node2=self.Node2.Value
             Node3=self.Node3.Value             
             self.DataBasePlot(int(Node1),int(Node2),int(Node3))
             self.update_console("Ploting characteristics for Turn: "+str(Node1)+" via "+str(Node2)+" to "+str(Node3))
              
    def DataBasePlot(self,i,*jk):
        def StepFunction(X,Y):
            Xstep=[]
            Ystep=[]
            Xstep.append(X[0])
            Ystep.append(Y[0])
            
            for i in xrange(1,len(X)-1):
                Xstep.append(X[i])
                Ystep.append(Y[i-1])
                
                Xstep.append(X[i])
                Ystep.append(Y[i])
                
            return Xstep,Ystep
    
        
        if i==-1:
            self.figure.remove()
            
            #self.figure.clf()
        else:
            if len(jk)==2:
                """Plot Turn"""
                PlotDataBase=DataBase.Turns.Data[DataBase.Turns.DictionaryVis2Py[(i,jk[0],jk[1])]]
            elif len(jk)==1:
                """Plot Link"""
                PlotDataBase=DataBase.Links.Data[DataBase.Links.DictionaryVis2Py[(i,jk[0])]]
            elif len(jk)==0:
                """Plot Node"""
                PlotDataBase=DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[i]]
            else: 
                print "error - You need to pass three args"
                return 1
            
            PlotDataBase.append([1.2*PlotDataBase[-1][0],0,0,PlotDataBase[-1][3],PlotDataBase[-1][4],0,PlotDataBase[-1][6]])
            
            prost_1 = [0.05,0.05,0.65,0.3]
            prost_2 = [0.05,0.4,0.65,0.3]
            prost_3 = [0.05,0.75,0.65,0.1]
            prost_4 = [0.75,0.05,0.2,0.9]
            Plot_Vol = self.figure.add_subplot(411)
            Plot_Cum = self.figure.add_subplot(412)
            Plot_InFlow = self.figure.add_subplot(425)
            Plot_OutFlow = self.figure.add_subplot(426)
            Plot_tCur = self.figure.add_subplot(414)
            
            
            
            t=[input[0] for input in PlotDataBase]
            InflowVector=[input[1] for input in PlotDataBase]
            OutflowVector=[input[2] for input in PlotDataBase]
            CumInflowVector=[input[3] for input in PlotDataBase]            
            CumOutflowVector=[input[4] for input in PlotDataBase]
            VolVector=[input[5] for input in PlotDataBase]
            tCurVector=[input[6] for input in PlotDataBase]
            
            Inflow_Interval_del=[]
            Outflow_Interval_del=[]
            
            for i in range(1,len(InflowVector)):
                if InflowVector[i]!=0:
                    Inflow_Interval_del.append(t[i])
                if OutflowVector[i]!=0:
                    Outflow_Interval_del.append(t[i])
            Inflow_Interval_del=sort(Inflow_Interval_del)
            Outflow_Interval_del=sort(Outflow_Interval_del)        
            Inflow_Interval=[Inflow_Interval_del[i]-Inflow_Interval_del[i-1] for i in range(1,len(Inflow_Interval_del))]
            Outflow_Interval=[Outflow_Interval_del[i]-Outflow_Interval_del[i-1] for i in range(1,len(Outflow_Interval_del))]
            max_dens_In=mean(Inflow_Interval)+3*sqrt(var(Inflow_Interval))            
            for i in range(len(Inflow_Interval)):
                if Inflow_Interval[i]>max_dens_In:
                    Inflow_Interval[i]=max_dens_In
            max_dens_Out=mean(Outflow_Interval)+3*sqrt(var(Outflow_Interval))            
            for i in range(len(Outflow_Interval)):
                if Outflow_Interval[i]>max_dens_Out:
                    Outflow_Interval[i]=max_dens_Out
            
            
            Plot_InFlow.set_title('Intflow deltas [t]')
            Plot_OutFlow.set_title('Intflow deltas [t]')            
            Plot_InFlow.hist(Inflow_Interval,30,normed=1)
            xy=[x*0.01 for x in range(0,100*int(max_dens_In))]
            lbd=1/mean(Inflow_Interval)
            yy=[lbd*exp(-lbd*xy[i]) for i in range(len(xy))]                          
            Plot_InFlow.plot(xy,yy,lw=2)
            Plot_InFlow.text(0.05,0.05,"mean"+str(mean(Inflow_Interval)))
            Plot_OutFlow.hist(Outflow_Interval,30,normed=1)
            xy=[x*0.01 for x in range(0,100*int(max_dens_Out))]
            lbd=1/mean(Outflow_Interval)
            yy=[lbd*exp(-lbd*xy[i]) for i in range(len(xy))]                          
            Plot_OutFlow.plot(xy,yy,lw=2)
            Plot_OutFlow.text(0.05,0.05,"mean "+str(mean(Outflow_Interval)))
            
            Plot_Cum.set_title('Cumulative In/Outflow [t]')
            tstep,CumInflowVectorStep=StepFunction(t,CumInflowVector)
            tstep,CumOutflowVectorStep=StepFunction(t,CumOutflowVector)
            Plot_Cum.plot(tstep,CumInflowVectorStep)
            Plot_Cum.plot(tstep,CumOutflowVectorStep)
            
            Plot_Vol.set_title('Volume ( veh on link [t] )')
            Plot_Vol.text(0.5,0.5,mean(VolVector))
            tstep,VolVectorStep=StepFunction(t,VolVector)
            Plot_Vol.plot(tstep,VolVectorStep)
            
            Plot_tCur.set_title('tCur')
            Plot_tCur.text(0.5,0.5,mean(tCurVector))
            tstep,tCurVectorStep=StepFunction(t,tCurVector)
            Plot_tCur.plot(tstep,tCurVectorStep)
            self.canvas.draw()
            
               



def update_console_ext(frame,label):
        Time=time.time()-TimeZero           
        frame.console.Label="SpaceTimeCylinder Console: "+label+" "+str(round(1000*Time)/1000)+"s"
     
def LoadVisum():
    """
    Dispatches the Visum instance and loads the Version file from path
    """
    global Visum
    Visum=win32com.client.Dispatch("Visum.Visum")

def VerLoad(path):
    global DSeg
    DSeg=0
    Visum.LoadVersion(path)
    DSeg=Visum.Net.DemandSegments.ItemByKey("C")
    Results.DSeg="C"
    #Results.AssignmentMethod=Visum.Procedures.
    return DSeg

def DemandSegmentChoice(Key):    
    global DSeg
    Results.DSeg=Key
    DSeg=0    
    DSeg=Visum.Net.DemandSegments.ItemByKey(Key)
    
    
    return DSeg

def Generate_xlwt_Report(filename):    
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Arkusz 1')
        
    ws.write(8,17,Results.VersionName)
    ws.write(9,17,Results.noZones)
    ws.write(10,17,Results.noNodes)
    ws.write(11,17,Results.noLinks)
    ws.write(12,17,Results.noTurns)
    ws.write(13,17,Results.noTrips)
    
    NoSim=1
    ws.write(15,NoSim+6,NoSim+1)
    ws.write(16,NoSim+6,Results.Times.VersionLoad)
    ws.write(17,NoSim+6,Results.Times.GetData)
    ws.write(18,NoSim+6,Results.Times.GetCoords)
    ws.write(19,NoSim+6,Results.Times.CrossTheTime)
    ws.write(23,NoSim+6,Results.Times.Plot_Cylinder)
    ws.write(21,NoSim+6,Results.AssType)
    ws.write(22,NoSim+6,Results.AssParam)
    ws.write(27,NoSim+6,Results.tStart)
    ws.write(28,NoSim+6,Results.tEnd)
    ws.write(29,NoSim+6,Results.Distribution)
    Results.tVector=sort(Results.tVector)
    time_dist=[Results.tVector[i]-Results.tVector[i-1] for i in range(1,len(Results.tVector))]
    
    ws.write(30,NoSim+6,mean(time_dist))
    ws.write(31,NoSim+6,var(time_dist))
    noEl=0
    for i in range(len(Results.Nodes)):
        Row=int(noEl+35+i)            
        ws.write(Row,2,Row-35)
        ws.write(Row,3,Results.VersionName)
        ws.write(Row,4,NoSim+1)        
        for j in range(len(Results.Nodes[i])):
                         
            ws.write(int(Row),int(5+j),str(Results.Nodes[i][j]))
    for i in range(len(Results.Links)):
        Row+=1            
        ws.write(Row,2,Row-35)
        ws.write(Row,3,Results.VersionName)
        ws.write(Row,4,NoSim+1)        
        for j in range(len(Results.Links[i])):
                        
            ws.write(int(Row),int(5+j),Results.Links[i][j])
    for i in range(len(Results.Turns)):
        Row+=1            
        ws.write(Row,2,Row-35)
        ws.write(Row,3,Results.VersionName)
        ws.write(Row,4,NoSim+1)        
        for j in range(len(Results.Turns[i])):                        
            ws.write(int(Row),int(5+j),Results.Turns[i][j])
    
    for i in range(len(Results.PathExcel)):
        Row+=1            
        ws.write(Row,2,Row-35)
        ws.write(Row,3,Results.VersionName)
        ws.write(Row,4,NoSim+1)       
        for j in range(len(Results.PathExcel[i])):                                    
            ws.write(int(Row),int(5+j),Results.PathExcel[i][j])
            
                     
    Results.Times.Generate_Report=time.time()-TimeZero
    ws.write(20,NoSim+6,Results.Times.Generate_Report)
     
    wb.save(filename)



def Generate_Excel_Report():    
    Excel=win32com.client.Dispatch("Excel.Application")
    Excel.Workbooks.Open("D:/results.xls") 
    Excel.Visible = True
    Cells=Excel.ActiveWorkBook.ActiveSheet.Cells
    Cells(8,17).Value=str(Results.VersionName)
    Cells(9,17).Value=str(Results.noZones)
    Cells(10,17).Value=str(Results.noNodes)
    Cells(11,17).Value=str(Results.noLinks)
    Cells(12,17).Value=str(Results.noTurns)
    Cells(13,17).Value=str(Results.noTrips)
    
    NoSim=Cells(9,3).Value
    Cells(15,NoSim+6).Value=str(NoSim+1)
    Cells(16,NoSim+6).Value=str(Results.Times.VersionLoad)
    Cells(17,NoSim+6).Value=str(Results.Times.GetData)
    Cells(18,NoSim+6).Value=str(Results.Times.GetCoords)
    Cells(19,NoSim+6).Value=str(Results.Times.CrossTheTime)
    Cells(21,NoSim+6).Value=str(Results.Times.Plot_Cylinder)
    Cells(23,NoSim+6).Value=Results.AssType
    Cells(22,NoSim+6).Value=Results.AssParam
    Cells(27,NoSim+6).Value=str(Results.tStart)
    Cells(28,NoSim+6).Value=str(Results.tEnd)
    Cells(29,NoSim+6).Value=str(Results.Distribution)
    Results.tVector=sort(Results.tVector)
    time_dist=[Results.tVector[i]-Results.tVector[i-1] for i in range(1,len(Results.tVector))]
    
    Cells(30,NoSim+6).Value=str(mean(time_dist))
    Cells(31,NoSim+6).Value=str(var(time_dist))
    noEl=Cells(8,3).Value
    for i in range(len(Results.Nodes)):
        Row=int(noEl+35+i)            
        Cells(Row,2).Value=str(Row-35)
        Cells(Row,3).Value=str(Results.VersionName)
        Cells(Row,4).Value=str(NoSim+1)        
        for j in range(len(Results.Nodes[i])):
                         
            Cells(int(Row),int(5+j)).Value=str(Results.Nodes[i][j])
    for i in range(len(Results.Links)):
        Row+=1            
        Cells(Row,2).Value=str(Row-35)
        Cells(Row,3).Value=str(Results.VersionName)
        Cells(Row,4).Value=str(NoSim+1)        
        for j in range(len(Results.Links[i])):
                        
            Cells(int(Row),int(5+j)).Value=str(Results.Links[i][j])
    for i in range(len(Results.Turns)):
        Row+=1            
        Cells(Row,2).Value=str(Row-35)
        Cells(Row,3).Value=str(Results.VersionName)
        Cells(Row,4).Value=str(NoSim+1)        
        for j in range(len(Results.Turns[i])):                        
            Cells(int(Row),int(5+j)).Value=str(Results.Turns[i][j])
    
    
    for i in range(len(Results.Paths)): 
        Row+=1            
        Cells(Row,2).Value=Row-35
        Cells(Row,3).Value=Results.VersionName
        Cells(Row,4).Value=NoSim+1
        Cells(Row,5).Value="Path"
        Cells(Row,6).Value=str([Results.Paths[i][0],Results.Paths[i][1]])
        Cells(Row,11).Value=Results.Paths[i][2]
        Cells(Row,12).Value=Results.Paths[i][3]
        Cells(Row,13).Value=Results.Paths[i][3]-Results.Paths[i][2] 
    Results.Times.Generate_Report=time.time()-TimeZero
    Cells(20,NoSim+6).Value=str(Results.Times.Generate_Report)
    Excel.Save
    Excel.Quit()

def Result_Calculation():    
    def generate_excel_line(type,key,length,CrNo,CrName,CrParam,t0,tCur_Visum,Data):     
        def generate_contionuous(t,vect):
            vect_cont=[]
            for i in range(len(t)-1):   
                for j in range(int(t[i]),int(t[1+i])):
                    vect_cont.append(vect[i])
            #ax.plot(t_cont,vect_cont)
            return vect_cont    
    
        """Input""" 
        
        CumInflow=[input[3] for input in Data]
        if max(CumInflow)==0:
            return "Empty"
        else:
            if length==0:
                length_exp=Inf
            else:
                length_exp=length             
            t=[input[0] for input in Data]
            t_cont=range(int(max(t)))
            Inflow=[input[1] for input in Data]
            Outflow=[input[2] for input in Data]                
            CumOutflow=[input[4] for input in Data]
            Vol=[input[5] for input in Data]
            tCur=[input[6] for input in Data]
            
            """Inflow Interval"""
            
            Inflow_Interval_del=[]
            Outflow_Interval_del=[]
            
            for i in range(1,len(Inflow)):
                if Inflow[i]!=0:
                    Inflow_Interval_del.append(t[i])
                if Outflow[i]!=0:
                    Outflow_Interval_del.append(t[i])
                    
            Inflow_Interval=[Inflow_Interval_del[i]-Inflow_Interval_del[i-1] for i in range(1,len(Inflow_Interval_del))]
            Outflow_Interval=[Outflow_Interval_del[i]-Outflow_Interval_del[i-1] for i in range(1,len(Outflow_Interval_del))]
            
            """Quantiles"""
            volsort=sort(Vol)
            Q_Vol=[0,0,0,0]
            for i in range(1,4):    
                Q_Vol[i]=volsort[int(len(volsort)*i/4)]               
            """Continuous functions"""
            Vol_cont=generate_contionuous(t,Vol)
            tCur_cont=generate_contionuous(t,tCur)
            
            #max_dens=mean(Inflow_Interval)+3*sqrt(var(Inflow_Interval))            
            #for i in range(len(Inflow_Interval)):
            #    if Inflow_Interval[i]>max_dens:      
            #        Inflow_Interval[i]=max_dens
            
            return [type,key,length,CrNo,CrName,str(CrParam),t[1],t[-1],t[-1]-t[1],CumInflow[-1],mean(Inflow_Interval),var(Inflow_Interval),mean(Outflow_Interval),var(Outflow_Interval),mean(Vol_cont),mode(Vol_cont)[0][0],var(Vol_cont),Q_Vol[0],Q_Vol[1],Q_Vol[2],Q_Vol[3],max(Vol_cont),max(Vol_cont)/length_exp,mean(Vol_cont)/length_exp,var(Vol_cont)/length_exp,t0,tCur_Visum,max(tCur_cont),mean(tCur_cont),var(tCur_cont)]
            
    Results.Nodes=[]
    Results.Links=[]
    Results.Turns=[]
    for i in range(1,Results.noNodes):
        """type,key,length,CrNo,CrName,CrParam,t0,tCur_Visum,Data"""            
        key=DataBase.Nodes.DictionaryPy2Vis[i]
        VisumData=DataBase.Arrays.Nodes[i]
        CR=DataBase.Nodes.CR.Data[int(DataBase.Nodes.CR.Dict[int(VisumData[4])])]
        line=generate_excel_line("Node",key,0,CR[0],CR[1],CR[2],VisumData[3],VisumData[7],DataBase.Nodes.Data[i])
        if line!="Empty": Results.Nodes.append(line)
    for i in range(1,Results.noLinks):
        """type,key,length,CrNo,CrName,CrParam,t0,tCur_Visum,Data"""            
        key=DataBase.Links.DictionaryPy2Vis[i]
        VisumData=DataBase.Arrays.Links[i]
        CR=DataBase.Links.CR.Data[int(DataBase.Links.CR.Dict[int(VisumData[5])])]
        line=generate_excel_line("Link",key,VisumData[7],CR[0],CR[1],CR[2],VisumData[4],VisumData[8],DataBase.Links.Data[i])            
        if line!="Empty": Results.Nodes.append(line)
    for i in range(1,Results.noTurns):
        """type,key,length,CrNo,CrName,CrParam,t0,tCur_Visum,Data"""            
        key=DataBase.Turns.DictionaryPy2Vis[i]
        VisumData=DataBase.Arrays.Turns[i]
        CR=DataBase.Turns.CR.Data[int(DataBase.Turns.CR.Dict[int(VisumData[4])])]
        line=generate_excel_line("Turn",key,0,CR[0],CR[1],CR[2],VisumData[3],VisumData[6],DataBase.Turns.Data[i])            
        if line!="Empty": Results.Nodes.append(line) 
    noZones=len(DataBase.Arrays.Zones) 
    #print "noZones", noZones   
    Results.AggregatedPaths=[[[[],[],[]] for j in range(noZones)] for i in range(noZones)]
    for i in range(1,len(DataBase.Paths.List)):         
        Results.AggregatedPaths[DataBase.Zones.DictionaryVis2Py[Results.Paths[i][0]]][DataBase.Zones.DictionaryVis2Py[Results.Paths[i][1]]][0].append(Results.Paths[i][2])
        Results.AggregatedPaths[DataBase.Zones.DictionaryVis2Py[Results.Paths[i][0]]][DataBase.Zones.DictionaryVis2Py[Results.Paths[i][1]]][1].append(Results.Paths[i][3])
        Results.AggregatedPaths[DataBase.Zones.DictionaryVis2Py[Results.Paths[i][0]]][DataBase.Zones.DictionaryVis2Py[Results.Paths[i][1]]][2].append(Results.Paths[i][3]-Results.Paths[i][2])
        #print Results.AggregatedPaths[DataBase.Zones.DictionaryVis2Py[Results.Paths[i][0]]][DataBase.Zones.DictionaryVis2Py[Results.Paths[i][1]]]
    Results.PathExcel=[]   
    for i in range(1,noZones):
        for j in range(1,noZones):            
            if len(Results.AggregatedPaths[i][j][2])>0:
                Inflow=sort(Results.AggregatedPaths[i][j][0])
                Outflow=sort(Results.AggregatedPaths[i][j][1])
                tCur=Results.AggregatedPaths[i][j][2]
                               
                #line=["Path",str([i,j]),"-","-","-","-",Results.AggregatedPaths[i][j][0],Results.AggregatedPaths[i][j][1],Results.AggregatedPaths[i][j][1]-Results.AggregatedPaths[i][j][0],len(Results.AggregatedPaths[i][j][2])))
                def generate_contionuous(t,vect):
                    vect_cont=[]
                    for ii in range(len(t)-1):   
                        for jj in range(int(t[ii]),int(t[1+ii])):
                            vect_cont.append(vect[ii])
                    #ax.plot(t_cont,vect_cont)
                    return vect_cont
           
                length=0                            
                t=Inflow
                t_cont=range(int(max(t)))
                Inflow_Interval_del=[]
                Outflow_Interval_del=[]
                for iii in range(1,len(Inflow)):                
                    Inflow_Interval_del.append(t[iii])                
                    Outflow_Interval_del.append(Outflow[iii])                    
                Inflow_Interval=[Inflow_Interval_del[iii]-Inflow_Interval_del[iii-1] for iii in range(1,len(Inflow_Interval_del))]
                Outflow_Interval=[Outflow_Interval_del[iii]-Outflow_Interval_del[iii-1] for iii in range(1,len(Outflow_Interval_del))]
                line=["Path",str((i,j)),length," "," "," ",Inflow[0],max(Outflow),max(Outflow)-Inflow[0],len(Inflow),mean(Inflow_Interval),var(Inflow_Interval),mean(Outflow_Interval),var(Outflow_Interval)," "," "," "," "," ",' ',' ',' ',' ',' '," "," "," ",max(tCur),mean(tCur),var(tCur)]
                Results.PathExcel.append(line)
                print Results.PathExcel[-1]
                
    
        

def Get_Assignment_XML_Data():
    from xml.dom import minidom, Node
    Visum.Procedures.SaveXml('D:/param.par;XML')
    Results.AssType='none'
    def Get_Parameters(node):
        global AssType    
        if node.nodeType == Node.ELEMENT_NODE:
            Napis= ' %s' % node.nodeName
            for (name, value) in node.attributes.items():
                if len(value)>0:
                    
                    Napis=Napis+ ' %s: %s;' % (name, value)
                    
                    if value=='Equilibrium': Results.AssType=value
                    if value=='Incremental': Results.AssType=value 
                    if value=='Cost Equilibrium': Results.AssType=value
                    if value=='Stochastic': Results.AssType=value
                    if value=="Dynamic User Equilibrium" : Results.AssType=value           
            #if node.attributes.get('ID') is not None:
            #    print '    ID: %s' % node.attributes.get('ID').value
        return Napis
    plik = minidom.parse("D:/param.par;xml")
    node = plik.childNodes[0].childNodes[1].childNodes[1]
    Napis=Get_Parameters(node)
    if Results.AssType=='Equilibrium':
        node = plik.childNodes[0].childNodes[1].childNodes[1].childNodes[1].childNodes[1].childNodes[1]
        Results.AssParam=Napis+Get_Parameters(node)
    if Results.AssType=='Incremental':
        node = plik.childNodes[0].childNodes[1].childNodes[1].childNodes[1].childNodes[1]
        Results.AssParam=Napis+Get_Parameters(node)
    else: Results.AssParam=Napis    
    


    
    
      

    
def Get_VDF_Data():
    def Get_Links_VDF_Data():
        DataBase.Links.CR.Dict={}
        DataBase.Links.CR.Data=[[]]
        Results.LinksCR=" "
        maxCR=0
        
        for i in range(len(Visum.Net.LinkTypes.GetMultiAttValues("No"))):
            CrNo=Visum.Procedures.Functions.CrFunctions.AttValue("CrFunctionNo_LinkType(%(typ)s)" %{'typ':i})
            DataBase.Links.CR.Dict[i]=CrNo
            if CrNo+1>maxCR: maxCR=CrNo+1
        
        for i in range(1,int(maxCR)):
            
            DataBase.Links.CR.Data.append([Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("CrFunctionNumber"),str(Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("CrFunctionType")),[],[]])
            
            
            if DataBase.Links.CR.Data[-1][1]=="Constant":
                
                DataBase.Links.CR.Data[-1][2]=[]
                DataBase.Links.CR.Data[-1][3]=[1 for j in range(1000)]
                
            
            if DataBase.Links.CR.Data[-1][1]=="HCM":
                
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("hcm_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("hcm_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor")]
                
                #CrFunctionsData[-1][3]=[(1.0+DataBase.Links.CR.Data[-1][2][0]*((j/(100.0*DataBase.Links.CR.Data[-1][2][2]))**DataBase.Links.CR.Data[-1][2][1])) for j in range(500)]
                a=DataBase.Links.CR.Data[-1][2][0]
                b=DataBase.Links.CR.Data[-1][2][1]
                c=DataBase.Links.CR.Data[-1][2][2]
                DataBase.Links.CR.Data[-1][3]=[1.0+a*(j/(100.0*c))**b for j in range(1000)]       
                
                
                
            if DataBase.Links.CR.Data[-1][1]=="HCM2":
                
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("hcm2_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("hcm2_b1"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("hcm2_b2"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor")]
                a=DataBase.Links.CR.Data[-1][2][0]
                c=DataBase.Links.CR.Data[-1][2][3]
                for j in range(1000):
                    if j<100:
                        b=DataBase.Links.CR.Data[-1][2][1]
                    else:
                        b=DataBase.Links.CR.Data[-1][2][2]
                    DataBase.Links.CR.Data[-1][3].append(1.0+a*(j/(100.0*c))**b)
                
                                                 
            if DataBase.Links.CR.Data[-1][1]=="HCM3":
                
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("hcm3_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("hcm3_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("hcm3_d")]
                a=DataBase.Links.CR.Data[-1][2][0]
                b=DataBase.Links.CR.Data[-1][2][1]
                c=DataBase.Links.CR.Data[-1][2][2]
                DataBase.Links.CR.Data[-1][3]=[1.0+a*(j/(100.0*c))**b for j in range(1000)]
                
            if DataBase.Links.CR.Data[-1][1]=="CONICAL":
                
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("conical_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor")]
                
                a=DataBase.Links.CR.Data[-1][2][0]
                c=DataBase.Links.CR.Data[-1][2][1]
                b=(2.0*a-1)/(2.0*a-2)
                DataBase.Links.CR.Data[-1][3]=[2+sqrt(a**2*(1-j/100.0*c)**2+b**2)-a*(1-j/100.0*c)-b for j in range(1000) ]
                
                
                #DataBase.Links.CR.Data[-1][3]=[(1.0+DataBase.Links.CR.Data[-1][2][0]*((i/(100.0*DataBase.Links.CR.Data[-1][2][2]))**DataBase.Links.CR.Data[-1][2][1])) for i in range(1000)]
            if DataBase.Links.CR.Data[-1][1]=="CONICAL_MARGINAL":        
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("conical_marginal_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor")]
                
                a=DataBase.Links.CR.Data[-1][2][0]
                c=DataBase.Links.CR.Data[-1][2][1]
                b=(2.0*a-1)/(2.0*a-2)
                DataBase.Links.CR.Data[-1][3]=[2+(a**2*(1-j/100.0*c)*(1-2*j/100.0*c)+b**2)/sqrt(a**2*(1-j/100.0*c)**2+b**2)-a*(1-2*j/100.0*c)-b for j in range(1000) ]
               
            if DataBase.Links.CR.Data[-1][1]=="EXPONENTIAL":
                CrFunctionsData[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("exponential_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("exponential_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("exponential_d")]
                a=DataBase.Links.CR.Data[-1][2][0]
                b=DataBase.Links.CR.Data[-1][2][1]
                c=DataBase.Links.CR.Data[-1][2][2]
                d=DataBase.Links.CR.Data[-1][2][3]
                for j in range(1000):
                    if j<100:
                        rest=0
                    else:
                        rest=d*(j/(100.0*c)-1)
                    DataBase.Links.CR.Data[-1][3].append(exp(a*j/(100.0*c))/b+rest)
                
                
                
            if DataBase.Links.CR.Data[-1][1]=="INRETS":
                
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("inrets_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor")]    
                a=DataBase.Links.CR.Data[-1][2][0]
                c=DataBase.Links.CR.Data[-1][2][1]
                for j in range(1000):
                    sat=(j/(100.0*c))
                    if j<100:
                       DataBase.Links.CR.Data[-1][3].append((1.1-a*sat)/(1.1-sat)) 
                    else:
                        DataBase.Links.CR.Data[-1][3].append(((1.1-a)/0.1)*sat**2)        
            
            
            if DataBase.Links.CR.Data[-1][1]=="LOGISTIC":
                
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("logistic_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("logistic_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("logistic_d"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("logistic_f")]
                
                a=DataBase.Links.CR.Data[-1][2][0]
                b=DataBase.Links.CR.Data[-1][2][1]
                c=DataBase.Links.CR.Data[-1][2][2]
                d=DataBase.Links.CR.Data[-1][2][3]
                f=DataBase.Links.CR.Data[-1][2][4]
                DataBase.Links.CR.Data[-1][3]=[ a/(1+f*exp(b-d*(j/(100.0*c)))) for j in range(1000) ]
                        
            if DataBase.Links.CR.Data[-1][1]=="QUADRATIC":
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("quadratic_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("quadratic_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("quadratic_d")]
                
                a=DataBase.Links.CR.Data[-1][2][0]
                b=DataBase.Links.CR.Data[-1][2][1]
                c=DataBase.Links.CR.Data[-1][2][2]
                d=DataBase.Links.CR.Data[-1][2][3]
                DataBase.Links.CR.Data[-1][3]=[ a+b*(j/(100.0*c))+d*(j/(100.0*c))**2 for j in range(1000) ]
                   
                
                
            if DataBase.Links.CR.Data[-1][1]=="SIGMOIDAL_MMF_LINKS":
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("sigmoidal_mmf_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("sigmoidal_mmf_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("sigmoidal_mmf_d"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("sigmoidal_mmf_f")]
                
                a=DataBase.Links.CR.Data[-1][2][0]
                b=DataBase.Links.CR.Data[-1][2][1]
                c=DataBase.Links.CR.Data[-1][2][2]
                d=DataBase.Links.CR.Data[-1][2][3]
                f=DataBase.Links.CR.Data[-1][2][4]
                DataBase.Links.CR.Data[-1][3]=[ (a*b+d*(j/(100.0*c))**f)/(b+(j/(100.0*c))**f) for j in range(1000) ]
                
            
                
            if DataBase.Links.CR.Data[-1][1]=="SIGMOIDAL_MMF_NODES":
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("sigmoidal_mmf_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("sigmoidal_mmf_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("sigmoidal_mmf_d"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("sigmoidal_mmf_f")]
            
                a=DataBase.Links.CR.Data[-1][2][0]
                b=DataBase.Links.CR.Data[-1][2][1]
                c=DataBase.Links.CR.Data[-1][2][2]
                d=DataBase.Links.CR.Data[-1][2][3]
                f=DataBase.Links.CR.Data[-1][2][4]
                DataBase.Links.CR.Data[-1][3]=[ (a*b+d*(j/(100.0*c))**f)/(b+(j/(100.0*c))**f) for j in range(1000) ]
                
            if DataBase.Links.CR.Data[-1][1]=="Akcelik":
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("akcelik_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("akcelik_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("akcelik_d")]
                a=DataBase.Links.CR.Data[-1][2][0]
                b=DataBase.Links.CR.Data[-1][2][1]
                c=DataBase.Links.CR.Data[-1][2][2]
                d=DataBase.Links.CR.Data[-1][2][3]
                
                DataBase.Links.CR.Data[-1][3]=[ 3600/4*a*((j/(100.0*c))-1+sqrt((j/(100.0*c)-1)**2+(8*b*j/(100.0*c))/(d*a))) for j in range(1000) ]
                
            
            if DataBase.Links.CR.Data[-1][1]=="Lohse":
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("lohse_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("lohse_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("lohse_satcrit")]
                a=DataBase.Links.CR.Data[-1][2][0]
                b=DataBase.Links.CR.Data[-1][2][1]
                c=DataBase.Links.CR.Data[-1][2][2]
                satcrit=DataBase.Links.CR.Data[-1][2][3]
                for j in range(1000):
                    sat=(j/(100.0*c))
                    if j<100:
                       DataBase.Links.CR.Data[-1][3].append(1+a*sat**b) 
                    else:
                        DataBase.Links.CR.Data[-1][3].append(1+a*sat**b+a*b*satcrit**(b-1)*(sat-satcrit)) 
    
                
            
            if DataBase.Links.CR.Data[-1][1]=="Linear bottle-neck":
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor")]
                return DataBase.Links.CR.Data[-1][1]
                
            if DataBase.Links.CR.Data[-1][1]=="Akcelik2":
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("akcelik2_a"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("akcelik2_b"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("akcelik2_d")]
                return DataBase.Links.CR.Data[-1][1]
            if DataBase.Links.CR.Data[-1][1]=="TMODEL_LINKS":
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_a1"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_a2"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_b1"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_b2"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_d1"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_d2"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_f1"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_f2"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_satcrit")]
                return DataBase.Links.CR.Data[-1][1]
            if DataBase.Links.CR.Data[-1][1]=="TMODEL_NODES":
                DataBase.Links.CR.Data[-1][2]=[Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_a1"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_a2"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_b1"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_b2"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_d1"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_d2"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_f1"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_f2"),Visum.Procedures.Functions.CrFunctions.CrFunction(i).AttValue("tmodel_satcrit")]
                return DataBase.Links.CR.Data[-1][1]
            Results.LinksCR=Results.LinksCR+"["+DataBase.Links.CR.Data[-1][1]+" with parameters: "+str(DataBase.Links.CR.Data[-1][2])+"], "
            
    def Get_Turns_VDF_Data():
        DataBase.Turns.CR.Dict={}
        DataBase.Turns.CR.Data=[[]]
        Results.TurnsCR=" "
        maxCR=0
        
        for i in range(9):        
            CrNo=Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunctions.AttValue("CrFunctionNo_TurnType(%(typ)s)" %{'typ':i})
            DataBase.Turns.CR.Dict[i]=CrNo
            if CrNo+1>maxCR: maxCR=CrNo+1
        
        for i in range(1,int(maxCR)):
            Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i)
            DataBase.Turns.CR.Data.append([Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("CrFunctionNumber"),str(Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("CrFunctionType")),[],[]])
            
            if DataBase.Links.CR.Data[-1][1]=="Constant":
                
                DataBase.Links.CR.Data[-1][2]=[]
                DataBase.Links.CR.Data[-1][3]=[1 for j in range(1000)]
                
            if DataBase.Turns.CR.Data[-1][1]=="HCM":
                
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("hcm_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("hcm_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor")]
                
                #CrFunctionsData[-1][3]=[(1.0+DataBase.Turns.CR.Data[-1][2][0]*((j/(100.0*DataBase.Turns.CR.Data[-1][2][2]))**DataBase.Turns.CR.Data[-1][2][1])) for j in range(1000)]
                a=DataBase.Turns.CR.Data[-1][2][0]
                b=DataBase.Turns.CR.Data[-1][2][1]
                c=DataBase.Turns.CR.Data[-1][2][2]
                DataBase.Turns.CR.Data[-1][3]=[1.0+a*(j/(100.0*c))**b for j in range(1000)]       
                
                
                
            if DataBase.Turns.CR.Data[-1][1]=="HCM2":
                
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("hcm2_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("hcm2_b1"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("hcm2_b2"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor")]
                a=DataBase.Turns.CR.Data[-1][2][0]
                c=DataBase.Turns.CR.Data[-1][2][3]
                for j in range(1000):
                    if j<100:
                        b=DataBase.Turns.CR.Data[-1][2][1]
                    else:
                        b=DataBase.Turns.CR.Data[-1][2][2]
                    DataBase.Turns.CR.Data[-1][3].append(1.0+a*(j/(100.0*c))**b)
                
                                                 
            if DataBase.Turns.CR.Data[-1][1]=="HCM3":
                
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("hcm3_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("hcm3_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("hcm3_d")]
                a=DataBase.Turns.CR.Data[-1][2][0]
                b=DataBase.Turns.CR.Data[-1][2][1]
                c=DataBase.Turns.CR.Data[-1][2][2]
                DataBase.Turns.CR.Data[-1][3]=[1.0+a*(j/(100.0*c))**b for j in range(1000)]
                
            if DataBase.Turns.CR.Data[-1][1]=="CONICAL":
                
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("conical_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor")]
                
                a=DataBase.Turns.CR.Data[-1][2][0]
                c=DataBase.Turns.CR.Data[-1][2][1]
                b=(2.0*a-1)/(2.0*a-2)
                DataBase.Turns.CR.Data[-1][3]=[2+sqrt(a**2*(1-j/100.0*c)**2+b**2)-a*(1-j/100.0*c)-b for j in range(1000) ]
                
                
                #DataBase.Turns.CR.Data[-1][3]=[(1.0+DataBase.Turns.CR.Data[-1][2][0]*((i/(100.0*DataBase.Turns.CR.Data[-1][2][2]))**DataBase.Turns.CR.Data[-1][2][1])) for i in range(1000)]
            if DataBase.Turns.CR.Data[-1][1]=="CONICAL_MARGINAL":        
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("conical_marginal_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor")]
                
                a=DataBase.Turns.CR.Data[-1][2][0]
                c=DataBase.Turns.CR.Data[-1][2][1]
                b=(2.0*a-1)/(2.0*a-2)
                DataBase.Turns.CR.Data[-1][3]=[2+(a**2*(1-j/100.0*c)*(1-2*j/100.0*c)+b**2)/sqrt(a**2*(1-j/100.0*c)**2+b**2)-a*(1-2*j/100.0*c)-b for j in range(1000) ]
               
            if DataBase.Turns.CR.Data[-1][1]=="EXPONENTIAL":
                CrFunctionsData[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("exponential_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("exponential_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("exponential_d")]
                a=DataBase.Turns.CR.Data[-1][2][0]
                b=DataBase.Turns.CR.Data[-1][2][1]
                c=DataBase.Turns.CR.Data[-1][2][2]
                d=DataBase.Turns.CR.Data[-1][2][3]
                for j in range(1000):
                    if j<100:
                        rest=0
                    else:
                        rest=d*(j/(100.0*c)-1)
                    DataBase.Turns.CR.Data[-1][3].append(exp(a*j/(100.0*c))/b+rest)
                
                
                
            if DataBase.Turns.CR.Data[-1][1]=="INRETS":
                
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("inrets_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor")]    
                a=DataBase.Turns.CR.Data[-1][2][0]
                c=DataBase.Turns.CR.Data[-1][2][1]
                for j in range(1000):
                    sat=(j/(100.0*c))
                    if j<100:
                       DataBase.Turns.CR.Data[-1][3].append((1.1-a*sat)/(1.1-sat)) 
                    else:
                        DataBase.Turns.CR.Data[-1][3].append(((1.1-a)/0.1)*sat**2)        
            
            
            if DataBase.Turns.CR.Data[-1][1]=="LOGISTIC":
                
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("logistic_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("logistic_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("logistic_d"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("logistic_f")]
                
                a=DataBase.Turns.CR.Data[-1][2][0]
                b=DataBase.Turns.CR.Data[-1][2][1]
                c=DataBase.Turns.CR.Data[-1][2][2]
                d=DataBase.Turns.CR.Data[-1][2][3]
                f=DataBase.Turns.CR.Data[-1][2][4]
                DataBase.Turns.CR.Data[-1][3]=[ a/(1+f*exp(b-d*(j/(100.0*c)))) for j in range(1000) ]
                        
            if DataBase.Turns.CR.Data[-1][1]=="QUADRATIC":
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("quadratic_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("quadratic_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("quadratic_d")]
                
                a=DataBase.Turns.CR.Data[-1][2][0]
                b=DataBase.Turns.CR.Data[-1][2][1]
                c=DataBase.Turns.CR.Data[-1][2][2]
                d=DataBase.Turns.CR.Data[-1][2][3]
                DataBase.Turns.CR.Data[-1][3]=[ a+b*(j/(100.0*c))+d*(j/(100.0*c))**2 for j in range(1000) ]
                   
                
                
            if DataBase.Turns.CR.Data[-1][1]=="SIGMOIDAL_MMF_Turns":
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("sigmoidal_mmf_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("sigmoidal_mmf_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("sigmoidal_mmf_d"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("sigmoidal_mmf_f")]
                
                a=DataBase.Turns.CR.Data[-1][2][0]
                b=DataBase.Turns.CR.Data[-1][2][1]
                c=DataBase.Turns.CR.Data[-1][2][2]
                d=DataBase.Turns.CR.Data[-1][2][3]
                f=DataBase.Turns.CR.Data[-1][2][4]
                DataBase.Turns.CR.Data[-1][3]=[ (a*b+d*(j/(100.0*c))**f)/(b+(j/(100.0*c))**f) for j in range(1000) ]
                
            
                
            if DataBase.Turns.CR.Data[-1][1]=="SIGMOIDAL_MMF_NODES":
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("sigmoidal_mmf_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("sigmoidal_mmf_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("sigmoidal_mmf_d"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("sigmoidal_mmf_f")]
            
                a=DataBase.Turns.CR.Data[-1][2][0]
                b=DataBase.Turns.CR.Data[-1][2][1]
                c=DataBase.Turns.CR.Data[-1][2][2]
                d=DataBase.Turns.CR.Data[-1][2][3]
                f=DataBase.Turns.CR.Data[-1][2][4]
                DataBase.Turns.CR.Data[-1][3]=[ (a*b+d*(j/(100.0*c))**f)/(b+(j/(100.0*c))**f) for j in range(1000) ]
                
            if DataBase.Turns.CR.Data[-1][1]=="Akcelik":
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("akcelik_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("akcelik_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("akcelik_d")]
                a=DataBase.Turns.CR.Data[-1][2][0]
                b=DataBase.Turns.CR.Data[-1][2][1]
                c=DataBase.Turns.CR.Data[-1][2][2]
                d=DataBase.Turns.CR.Data[-1][2][3]
                
                DataBase.Turns.CR.Data[-1][3]=[ 3600/4*a*((j/(100.0*c))-1+sqrt((j/(100.0*c)-1)**2+(8*b*j/(100.0*c))/(d*a))) for j in range(1000) ]
                
            
            if DataBase.Turns.CR.Data[-1][1]=="Lohse":
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("lohse_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("lohse_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("lohse_satcrit")]
                a=DataBase.Turns.CR.Data[-1][2][0]
                b=DataBase.Turns.CR.Data[-1][2][1]
                c=DataBase.Turns.CR.Data[-1][2][2]
                satcrit=DataBase.Turns.CR.Data[-1][2][3]
                for j in range(1000):
                    sat=(j/(100.0*c))
                    if j<100:
                       DataBase.Turns.CR.Data[-1][3].append(1+a*sat**b) 
                    else:
                        DataBase.Turns.CR.Data[-1][3].append(1+a*sat**b+a*b*satcrit**(b-1)*(sat-satcrit)) 
    
                
            
            if DataBase.Turns.CR.Data[-1][1]=="Linear bottle-neck":
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor")]
                return DataBase.Turns.CR.Data[-1][1]
                
            if DataBase.Turns.CR.Data[-1][1]=="Akcelik2":
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("akcelik2_a"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("akcelik2_b"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("akcelik2_d")]
                return DataBase.Turns.CR.Data[-1][1]
            if DataBase.Turns.CR.Data[-1][1]=="TMODEL_Turns":
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_a1"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_a2"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_b1"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_b2"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_d1"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_d2"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_f1"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_f2"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_satcrit")]
                return DataBase.Turns.CR.Data[-1][1]
            if DataBase.Turns.CR.Data[-1][1]=="TMODEL_NODES":
                DataBase.Turns.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_a1"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_a2"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_b1"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_b2"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_d1"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_d2"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_f1"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_f2"),Visum.Procedures.Functions.NodeImpedancePara.TurnCrFunction(i).AttValue("tmodel_satcrit")]
                return DataBase.Turns.CR.Data[-1][1]
            Results.TurnsCR=Results.TurnsCR+"["+DataBase.Turns.CR.Data[-1][1]+" with parameters: "+str(DataBase.Turns.CR.Data[-1][2])+"], "
                 
    def Get_Nodes_VDF_Data():
        DataBase.Nodes.CR.Dict={}
        DataBase.Nodes.CR.Data=[[]]
        Results.NodesCR=" "
        maxCR=0
        
        for i in range(100):
            
            CrNo=Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunctions.AttValue("CrFunctionNo_NodeType(%(typ)s)" %{'typ':i})
            DataBase.Nodes.CR.Dict[i]=CrNo
            if CrNo+1>maxCR: maxCR=CrNo+1
        
            
        
        for i in range(1,int(maxCR)):
            
            
            DataBase.Nodes.CR.Data.append([Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("CrFunctionNumber"),str(Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("CrFunctionType")),[],[]])
            
            if DataBase.Links.CR.Data[-1][1]=="Constant":
                
                DataBase.Links.CR.Data[-1][2]=[]
                DataBase.Links.CR.Data[-1][3]=[1 for j in range(1000)]
            
            if DataBase.Nodes.CR.Data[-1][1]=="HCM":
                
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("hcm_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("hcm_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor")]
                
                #CrFunctionsData[-1][3]=[(1.0+DataBase.Nodes.CR.Data[-1][2][0]*((j/(100.0*DataBase.Nodes.CR.Data[-1][2][2]))**DataBase.Nodes.CR.Data[-1][2][1])) for j in range(1000)]
                a=DataBase.Nodes.CR.Data[-1][2][0]
                b=DataBase.Nodes.CR.Data[-1][2][1]
                c=DataBase.Nodes.CR.Data[-1][2][2]
                DataBase.Nodes.CR.Data[-1][3]=[1.0+a*(j/(100.0*c))**b for j in range(1000)]       
                
                
                
            if DataBase.Nodes.CR.Data[-1][1]=="HCM2":
                
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("hcm2_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("hcm2_b1"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("hcm2_b2"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor")]
                a=DataBase.Nodes.CR.Data[-1][2][0]
                c=DataBase.Nodes.CR.Data[-1][2][3]
                for j in range(1000):
                    if j<100:
                        b=DataBase.Nodes.CR.Data[-1][2][1]
                    else:
                        b=DataBase.Nodes.CR.Data[-1][2][2]
                    DataBase.Nodes.CR.Data[-1][3].append(1.0+a*(j/(100.0*c))**b)
                
                                                 
            if DataBase.Nodes.CR.Data[-1][1]=="HCM3":
                
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("hcm3_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("hcm3_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("hcm3_d")]
                a=DataBase.Nodes.CR.Data[-1][2][0]
                b=DataBase.Nodes.CR.Data[-1][2][1]
                c=DataBase.Nodes.CR.Data[-1][2][2]
                DataBase.Nodes.CR.Data[-1][3]=[1.0+a*(j/(100.0*c))**b for j in range(1000)]
                
            if DataBase.Nodes.CR.Data[-1][1]=="CONICAL":
                
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("conical_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor")]
                
                a=DataBase.Nodes.CR.Data[-1][2][0]
                c=DataBase.Nodes.CR.Data[-1][2][1]
                b=(2.0*a-1)/(2.0*a-2)
                DataBase.Nodes.CR.Data[-1][3]=[2+sqrt(a**2*(1-j/100.0*c)**2+b**2)-a*(1-j/100.0*c)-b for j in range(1000) ]
                
                
                #DataBase.Nodes.CR.Data[-1][3]=[(1.0+DataBase.Nodes.CR.Data[-1][2][0]*((i/(100.0*DataBase.Nodes.CR.Data[-1][2][2]))**DataBase.Nodes.CR.Data[-1][2][1])) for i in range(1000)]
            if DataBase.Nodes.CR.Data[-1][1]=="CONICAL_MARGINAL":        
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("conical_marginal_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor")]
                
                a=DataBase.Nodes.CR.Data[-1][2][0]
                c=DataBase.Nodes.CR.Data[-1][2][1]
                b=(2.0*a-1)/(2.0*a-2)
                DataBase.Nodes.CR.Data[-1][3]=[2+(a**2*(1-j/100.0*c)*(1-2*j/100.0*c)+b**2)/sqrt(a**2*(1-j/100.0*c)**2+b**2)-a*(1-2*j/100.0*c)-b for j in range(1000) ]
               
            if DataBase.Nodes.CR.Data[-1][1]=="EXPONENTIAL":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("exponential_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("exponential_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("exponential_d")]
                a=DataBase.Nodes.CR.Data[-1][2][0]
                b=DataBase.Nodes.CR.Data[-1][2][1]
                c=DataBase.Nodes.CR.Data[-1][2][2]
                d=DataBase.Nodes.CR.Data[-1][2][3]
                for j in range(1000):
                    if j<100:
                        rest=0
                    else:
                        rest=d*(j/(100.0*c)-1)
                    DataBase.Nodes.CR.Data[-1][3].append(exp(a*j/(100.0*c))/b+rest)
                
                
                
            if DataBase.Nodes.CR.Data[-1][1]=="INRETS":
                
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("inrets_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor")]    
                a=DataBase.Nodes.CR.Data[-1][2][0]
                c=DataBase.Nodes.CR.Data[-1][2][1]
                for j in range(1000):
                    sat=(j/(100.0*c))
                    if j<100:
                       DataBase.Nodes.CR.Data[-1][3].append((1.1-a*sat)/(1.1-sat)) 
                    else:
                        DataBase.Nodes.CR.Data[-1][3].append(((1.1-a)/0.1)*sat**2)        
            
            
            if DataBase.Nodes.CR.Data[-1][1]=="LOGISTIC":
                
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("logistic_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("logistic_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("logistic_d"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("logistic_f")]
                
                a=DataBase.Nodes.CR.Data[-1][2][0]
                b=DataBase.Nodes.CR.Data[-1][2][1]
                c=DataBase.Nodes.CR.Data[-1][2][2]
                d=DataBase.Nodes.CR.Data[-1][2][3]
                f=DataBase.Nodes.CR.Data[-1][2][4]
                DataBase.Nodes.CR.Data[-1][3]=[ a/(1+f*exp(b-d*(j/(100.0*c)))) for j in range(1000) ]
                        
            if DataBase.Nodes.CR.Data[-1][1]=="QUADRATIC":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("quadratic_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("quadratic_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("quadratic_d")]
                
                a=DataBase.Nodes.CR.Data[-1][2][0]
                b=DataBase.Nodes.CR.Data[-1][2][1]
                c=DataBase.Nodes.CR.Data[-1][2][2]
                d=DataBase.Nodes.CR.Data[-1][2][3]
                DataBase.Nodes.CR.Data[-1][3]=[ a+b*(j/(100.0*c))+d*(j/(100.0*c))**2 for j in range(1000) ]
                   
                
                
            if DataBase.Nodes.CR.Data[-1][1]=="SIGMOIDAL_MMF_Nodes":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("sigmoidal_mmf_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("sigmoidal_mmf_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("sigmoidal_mmf_d"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("sigmoidal_mmf_f")]
                
                a=DataBase.Nodes.CR.Data[-1][2][0]
                b=DataBase.Nodes.CR.Data[-1][2][1]
                c=DataBase.Nodes.CR.Data[-1][2][2]
                d=DataBase.Nodes.CR.Data[-1][2][3]
                f=DataBase.Nodes.CR.Data[-1][2][4]
                DataBase.Nodes.CR.Data[-1][3]=[ (a*b+d*(j/(100.0*c))**f)/(b+(j/(100.0*c))**f) for j in range(1000) ]
                
            
                
            if DataBase.Nodes.CR.Data[-1][1]=="SIGMOIDAL_MMF_NODES":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("sigmoidal_mmf_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("sigmoidal_mmf_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("sigmoidal_mmf_d"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("sigmoidal_mmf_f")]
            
                a=DataBase.Nodes.CR.Data[-1][2][0]
                b=DataBase.Nodes.CR.Data[-1][2][1]
                c=DataBase.Nodes.CR.Data[-1][2][2]
                d=DataBase.Nodes.CR.Data[-1][2][3]
                f=DataBase.Nodes.CR.Data[-1][2][4]
                DataBase.Nodes.CR.Data[-1][3]=[ (a*b+d*(j/(100.0*c))**f)/(b+(j/(100.0*c))**f) for j in range(1000) ]
                
            if DataBase.Nodes.CR.Data[-1][1]=="Akcelik":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("akcelik_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("akcelik_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("akcelik_d")]
                a=DataBase.Nodes.CR.Data[-1][2][0]
                b=DataBase.Nodes.CR.Data[-1][2][1]
                c=DataBase.Nodes.CR.Data[-1][2][2]
                d=DataBase.Nodes.CR.Data[-1][2][3]
                
                DataBase.Nodes.CR.Data[-1][3]=[ 3600/4*a*((j/(100.0*c))-1+sqrt((j/(100.0*c)-1)**2+(8*b*j/(100.0*c))/(d*a))) for j in range(1000) ]
                
            
            if DataBase.Nodes.CR.Data[-1][1]=="Lohse":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("lohse_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("lohse_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("lohse_satcrit")]
                a=DataBase.Nodes.CR.Data[-1][2][0]
                b=DataBase.Nodes.CR.Data[-1][2][1]
                c=DataBase.Nodes.CR.Data[-1][2][2]
                satcrit=DataBase.Nodes.CR.Data[-1][2][3]
                for j in range(1000):
                    sat=(j/(100.0*c))
                    if j<100:
                       DataBase.Nodes.CR.Data[-1][3].append(1+a*sat**b) 
                    else:
                        DataBase.Nodes.CR.Data[-1][3].append(1+a*sat**b+a*b*satcrit**(b-1)*(sat-satcrit)) 
            
            if DataBase.Nodes.CR.Data[-1][1]=="Linear bottle-neck":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor")]
                return DataBase.Nodes.CR.Data[-1][1]
                
            if DataBase.Nodes.CR.Data[-1][1]=="Akcelik2":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("akcelik2_a"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("akcelik2_b"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("akcelik2_d")]
                return DataBase.Nodes.CR.Data[-1][1]
            if DataBase.Nodes.CR.Data[-1][1]=="TMODEL_Nodes":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_a1"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_a2"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_b1"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_b2"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_d1"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_d2"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_f1"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_f2"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_satcrit")]
                return DataBase.Nodes.CR.Data[-1][1]
            if DataBase.Nodes.CR.Data[-1][1]=="TMODEL_NODES":
                DataBase.Nodes.CR.Data[-1][2]=[Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_a1"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_a2"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_b1"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_b2"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("capacityFactor"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_d1"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_d2"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_f1"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_f2"),Visum.Procedures.Functions.NodeImpedancePara.NodeCrFunction(i).AttValue("tmodel_satcrit")]
                return DataBase.Nodes.CR.Data[-1][1]
            Results.NodesCR=Results.NodesCR+"["+DataBase.Nodes.CR.Data[-1][1]+" with parameters: "+str(DataBase.Nodes.CR.Data[-1][2])+"], "
            
    Get_Links_VDF_Data()
    Get_Turns_VDF_Data()
    Get_Nodes_VDF_Data()
           
""" Procedure #4 Get Data"""                            
def Get_Data():
    
    def GetAttributes2Array(Object,Attributes):
        """
        For the given Visum instace it downloads the 
        given Object container with attributes 
        and saves it to the nparray
        """
        Container=VisumPy.helpers.GetContainer(Visum, Object)
        list=Container.GetMultipleAttributes(Attributes)
        
    
        Array=zeros([len(list)+1,len(Attributes)])
        for i in xrange(1,len(list)+1):
            for j in xrange(len(Attributes)):
                Array[i,j]=list[i-1][j]
        del list
        return Array
    
    
    MethodImpAtNodes=Visum.Procedures.Functions.NodeImpedancePara.AttValue("NodeImpedanceMethod")
    
    """  0 - turns
         1 - ICA 
         2 - nodes"""
    # get Needed information from Visum COM into arrays 
    DataBase.Arrays.Nodes=GetAttributes2Array('Nodes',["No","XCoord","YCoord","t0Prt","TypeNo","CapPrT","T0PrT","TurntCurTot"])
    DataBase.Arrays.Links=GetAttributes2Array('Links',["No","FromNodeNo","ToNodeNo","VolVehPrT(AP)","T0_PrTSys(C)","CrNo","CapPrT","Length","TCur_PrTSys(C)"])
    DataBase.Arrays.Turns=GetAttributes2Array('Turns',["FromNodeNo","ViaNodeNo","ToNodeNo","T0_PrTSys(C)","TypeNo","CapPrT","TCur_PrTSys(C)"])
    DataBase.Arrays.Connectors=GetAttributes2Array('Connectors',["ZoneNo","NodeNo","Direction","T0_TSys(C)","TypeNo"])
    DataBase.Arrays.Zones=GetAttributes2Array('Zones',["No","XCoord","YCoord"])
    
    
    # set the dimensions of matrices in database
    noZones=len(DataBase.Arrays.Zones)-1
    noNodes=len(DataBase.Arrays.Nodes)-1
    noLinks=len(DataBase.Arrays.Links)-1
    noTurns=len(DataBase.Arrays.Turns)-1
    noConnectors=len(DataBase.Arrays.Connectors)-1
    Results.noZones=noZones
    Results.noNodes=noNodes
    Results.noLinks=noLinks
    Results.noTurns=noTurns
    #set initial database matrixes of empty DataBase.Arrays
    DataBase.Zones.Data=[[[0,0,0,0,0,0,0,0,0,0,0,0,0]] for d1 in xrange(noZones+1)]
    
    DataBase.Zones.DictionaryPy2Vis={}
    DataBase.Zones.DictionaryVis2Py={}
    for i in xrange(noZones+1):
        DataBase.Zones.DictionaryPy2Vis[i]=int((DataBase.Arrays.Zones[i][0]))
        DataBase.Zones.DictionaryVis2Py[DataBase.Arrays.Zones[i][0]]=i
        
    if DataBase.tCurCalcMethod==0: DataBase.Nodes.Data=[[[0,0,0,0,0,0,DataBase.Arrays.Nodes[i][3],0,0,DataBase.Arrays.Nodes[i][5],0,2,0]] for i in xrange(noNodes+1)]    
    elif DataBase.tCurCalcMethod==1: DataBase.Nodes.Data=[[[0,0,0,0,0,0,DataBase.Arrays.Nodes[i][3],DataBase.Arrays.Nodes[i][3],MethodImpAtNodes/2*DataBase.Nodes.CR.Dict[DataBase.Arrays.Nodes[i][4]],DataBase.Arrays.Nodes[i][5],0,2,0]] for i in xrange(noNodes+1)]    
    elif DataBase.tCurCalcMethod==2: DataBase.Nodes.Data=[[[0,0,0,0,0,0,DataBase.Arrays.Nodes[i][3],DataBase.Arrays.Nodes[i][3],0,DataBase.Arrays.Nodes[i][5],0,2,0]] for i in xrange(noNodes+1)]
        
    DataBase.Nodes.DictionaryPy2Vis={}
    DataBase.Nodes.DictionaryVis2Py={}
    for i in xrange(noNodes+1):
        DataBase.Nodes.DictionaryPy2Vis[i]=(DataBase.Arrays.Nodes[i][0])
        DataBase.Nodes.DictionaryVis2Py[DataBase.Arrays.Nodes[i][0]]=i
        
    
        
    if DataBase.tCurCalcMethod==0: DataBase.Links.Data=[[[0,0,0,0,0,0,DataBase.Arrays.Links[i][4],DataBase.Arrays.Links[i][8],0,DataBase.Arrays.Links[i][6],0,1,DataBase.Arrays.Links[i][7]]] for i in xrange(noLinks+1)]
    elif DataBase.tCurCalcMethod==1: DataBase.Links.Data=[[[0,0,0,0,0,0,DataBase.Arrays.Links[i][4],DataBase.Arrays.Links[i][4],DataBase.Arrays.Links[i][5],DataBase.Arrays.Links[i][6],0,1,DataBase.Arrays.Links[i][7]]] for i in xrange(noLinks+1)]
    elif DataBase.tCurCalcMethod==2: DataBase.Links.Data=[[[0,0,0,0,0,0,DataBase.Arrays.Links[i][4],DataBase.Arrays.Links[i][4],1000,DataBase.Arrays.Links[i][6],0,1,DataBase.Arrays.Links[i][7]]] for i in xrange(noLinks+1)]
    
    
    DataBase.Links.DictionaryPy2Vis={}
    DataBase.Links.DictionaryVis2Py={}
    for i in xrange(noLinks+1):
        DataBase.Links.DictionaryPy2Vis[i]=(DataBase.Arrays.Links[i][1],DataBase.Arrays.Links[i][2])
        DataBase.Links.DictionaryVis2Py[(DataBase.Arrays.Links[i][1],DataBase.Arrays.Links[i][2])]=i
    
    if DataBase.tCurCalcMethod==0: DataBase.Turns.Data=[[[0,0,0,0,0,0,DataBase.Arrays.Turns[i][3],DataBase.Arrays.Turns[i][6],0,DataBase.Arrays.Turns[i][5],0,3,0]] for i in xrange(noTurns+1)]
    elif DataBase.tCurCalcMethod==1: DataBase.Turns.Data=[[[0,0,0,0,0,0,DataBase.Arrays.Turns[i][3],DataBase.Arrays.Turns[i][3],-(MethodImpAtNodes-2)/2*DataBase.Turns.CR.Dict[DataBase.Arrays.Turns[i][4]],DataBase.Arrays.Turns[i][5],0,3,0]] for i in xrange(noTurns+1)]
    elif DataBase.tCurCalcMethod==2: DataBase.Turns.Data=[[[0,0,0,0,0,0,DataBase.Arrays.Turns[i][3],DataBase.Arrays.Turns[i][3],1000,DataBase.Arrays.Turns[i][5],0,3,0]] for i in xrange(noTurns+1)]
    
    
    DataBase.Turns.DictionaryPy2Vis={}
    DataBase.Turns.DictionaryVis2Py={}
    for i in xrange(noTurns+1):
        DataBase.Turns.DictionaryPy2Vis[i]=(DataBase.Arrays.Turns[i][0],DataBase.Arrays.Turns[i][1],DataBase.Arrays.Turns[i][2])
        DataBase.Turns.DictionaryVis2Py[(DataBase.Arrays.Turns[i][0],DataBase.Arrays.Turns[i][1],DataBase.Arrays.Turns[i][2])]=i
   
    
 
    
    DataBase.OrigConnectors.Data=[]
    DataBase.DestConnectors.Data=[]
    DataBase.OrigConnectors.DictionaryPy2Vis={}
    DataBase.OrigConnectors.DictionaryVis2Py={}
    DataBase.DestConnectors.DictionaryPy2Vis={}
    DataBase.DestConnectors.DictionaryVis2Py={}
    k=0
    for i in xrange(noConnectors+1):
        if DataBase.Arrays.Connectors[i][2]==1:
            DataBase.OrigConnectors.Data.append([])
            DataBase.OrigConnectors.Data[-1].append([0,0,0,0,0,0,DataBase.Arrays.Connectors[i][3],DataBase.Arrays.Connectors[i][3],0,0,0,0,0])
            DataBase.OrigConnectors.DictionaryPy2Vis[k]=[(DataBase.Arrays.Connectors[i][0],DataBase.Arrays.Connectors[i][1])]
            DataBase.OrigConnectors.DictionaryVis2Py[(DataBase.Arrays.Connectors[i][0],DataBase.Arrays.Connectors[i][1])]=k
            k+=1
    k=0
    for i in xrange(noConnectors+1):
        if DataBase.Arrays.Connectors[i][2]!=1:
            DataBase.DestConnectors.Data.append([])
            DataBase.DestConnectors.Data[-1].append([0,0,0,0,0,0,DataBase.Arrays.Connectors[i][3],DataBase.Arrays.Connectors[i][3],0,0,0,0,0])
            DataBase.DestConnectors.DictionaryPy2Vis[k]=[(DataBase.Arrays.Connectors[i][1],DataBase.Arrays.Connectors[i][0])]
            DataBase.DestConnectors.DictionaryVis2Py[(DataBase.Arrays.Connectors[i][1],DataBase.Arrays.Connectors[i][0])]=k
            k+=1
   
    
    
    # Matrix of initial paths with initial data t0
    DataBase.Paths.Mtx=[[[] for d1 in xrange(noZones+1)] for d2 in xrange(noZones+1)]
    
    
    """for i in xrange(1,noNodes):
        DataBase.Nodes.Data[i]=[[0,0,0,0,0,0,Arrays.Nodes[i-1][3]]]
    for i in xrange(1,noZones):
        DataBase.Zones.Data[i]=[[0,0,0,0,0,0,0]]
    for i in xrange(1,noLinks):
        DataBase.Links.Data[i-1]=[[0,0,0,0,0,0,Arrays.Links[i-1][4]]]
    for i in xrange(1,len(Arrays.Turns)+1):
        DataBase.Turns.Data[int(Arrays.Turns[i-1][0])][int(Arrays.Turns[i-1][1])][int(Arrays.Turns[i-1][2])]=[[0,0,0,0,0,0,Arrays.Turns[i-1][3]]]
    for i in xrange(1,len(Arrays.Connectors)+1):
        if Arrays.Connectors[i-1][2]==1:
            DataBase.OrigConnectors.Data[int(Arrays.Connectors[i-1][0])][int(Arrays.Connectors[i-1][1])]=[[0,0,0,0,0,0,Arrays.Connectors[i-1][3]]]
        else:
            DataBase.DestConnectors.Data[int(Arrays.Connectors[i-1][1])][int(Arrays.Connectors[i-1][0])]=[[0,0,0,0,0,0,Arrays.Connectors[i-1][3]]]
    del noZones
    del noNodes
"""

""" Procedure #5 GetGeneral Path Coords """
def Get_All_Path_Coords():
    
    def Get_Path_Coords(fromZone,toZone,pathInd):
        
        """
        For the given path it calculates t0 (tCur for the time being) coords (X,Y,t)
        """
        Flow=int(round(DSeg.GetPathFlow(fromZone,toZone,pathInd)-.5))
        TotTime=DSeg.GetPathFlow(fromZone,toZone,pathInd)        
        Nodes_tuple=DSeg.GetPathNodes(fromZone, toZone,pathInd).GetMultiAttValues("No")
        
        Nodes=[]
        X=[]
        Y=[]
        
        
        #first point
        
        X.append(int(DataBase.Arrays.Zones[DataBase.Zones.DictionaryVis2Py[fromZone]][1]))
        Y.append(int(DataBase.Arrays.Zones[DataBase.Zones.DictionaryVis2Py[fromZone]][2]))
        
        #second point
        Nodes.append(int(Nodes_tuple[0][1]))
        if len(Nodes_tuple)>1:
            Nodes.append(int(Nodes_tuple[1][1]))
        
        X.append(int(DataBase.Arrays.Nodes[DataBase.Nodes.DictionaryVis2Py[Nodes_tuple[0][1]]][1]))
        Y.append(int(DataBase.Arrays.Nodes[DataBase.Nodes.DictionaryVis2Py[Nodes_tuple[0][1]]][2]))
        
        #loop
        for t in xrange(1,len(Nodes_tuple)-1):
            #link
            Nodes.append(int(Nodes_tuple[t+1][1]))
            X.append(int(DataBase.Arrays.Nodes[DataBase.Nodes.DictionaryVis2Py[Nodes_tuple[t][1]]][1]))
            Y.append(int(DataBase.Arrays.Nodes[DataBase.Nodes.DictionaryVis2Py[Nodes_tuple[t][1]]][2]))
            
            #turn
            X.append(int(DataBase.Arrays.Nodes[DataBase.Nodes.DictionaryVis2Py[Nodes_tuple[t][1]]][1]))
            Y.append(int(DataBase.Arrays.Nodes[DataBase.Nodes.DictionaryVis2Py[Nodes_tuple[t][1]]][2]))
            
        #last link
        if len(Nodes_tuple)>1:
            X.append(int(DataBase.Arrays.Nodes[DataBase.Nodes.DictionaryVis2Py[Nodes_tuple[-1][1]]][1]))
            Y.append(int(DataBase.Arrays.Nodes[DataBase.Nodes.DictionaryVis2Py[Nodes_tuple[-1][1]]][2]))
            X.append(int(DataBase.Arrays.Zones[DataBase.Zones.DictionaryVis2Py[toZone]][1]))
            Y.append(int(DataBase.Arrays.Zones[DataBase.Zones.DictionaryVis2Py[toZone]][2]))
            
           
        
        return X,Y,Nodes,Flow
    
    noZones=len(DataBase.Arrays.Zones)-1
    
    for i in xrange(1,noZones+1):
        for j in xrange(1,noZones+1):
            if i!=j:
                for k in xrange(1,DSeg.GetNumPaths(DataBase.Zones.DictionaryPy2Vis[i],DataBase.Zones.DictionaryPy2Vis[j])+1):
                    
                    DataBase.Paths.Mtx[i][j].append(Get_Path_Coords(DataBase.Zones.DictionaryPy2Vis[i],DataBase.Zones.DictionaryPy2Vis[j],k))

""" Set times to Paths"""
def Set_times_to_Paths():
    
    noZones=len(DataBase.Arrays.Zones)-1
    tStart=1
    tEnd=3600
    Results.tStart=tStart
    Results.tEnd=tEnd
    Results.tVector=[]
    Results.Distribution=DataBase.TimeDistribution
    no=0
    noTrips=0
    for i in xrange(1,noZones+1):        
        for j in xrange(1,noZones+1):
            for k in xrange(len(DataBase.Paths.Mtx[i][j])):
                
                noTrips+=DataBase.Paths.Mtx[i][j][k][3]
                
    DataBase.Paths.ListDel=[[0,0,0,0,-10] for d1 in xrange(noTrips)]
    
    for i in xrange(1,noZones+1):
        for j in xrange(1,noZones+1):
            for k in xrange(len(DataBase.Paths.Mtx[i][j])):
                path=DataBase.Paths.Mtx[i][j][k]
                
                for l in xrange(path[3]):
                    if DataBase.TimeDistribution==0:
                        t=random.randint(tStart,tEnd)
                    else:                        
                        t=int(l*1.0/path[3]*(tEnd-tStart))                        
                    Results.tVector.append(t)
                    DataBase.Paths.ListDel[no]=[path[0],path[1],path[2],[t],t,1,DataBase.Zones.DictionaryPy2Vis[i],DataBase.Zones.DictionaryPy2Vis[j]]
                    no+=1
    DataBase.Paths.List=[DataBase.Paths.ListDel[i] for i in range(no)]
    Results.Paths=[0 for i in range(no)]
    Results.noTrips=noTrips    
    del DataBase.Paths.ListDel
    

def Cross_The_Time(frame):
    global KJ
    KJ=1/7.0
    
    def tCur_calculate(CrNo,Sat,t0,q,qmax,Vol,type,length):
        
        def calculate_trapez(Vol,length,V0):
            length=1000.0*length 
            Cap=V0*KJ/6
            W=V0/3
            k=Vol/length            
            k1=2*Cap/V0
            if k<=k1: q=-V0/(2*k1)*k*k+V0*k
            else: q=max(0,min(Cap,(KJ-k)*W))    
            v=q/k
            if k==0: tCur=length/V0
            elif v==0: tCur=13#Inf            
            else: tCur=length/v
            #print "Cap",Cap,"V0",V0,"k",k,"v",v,"t",tCur,"Vol",Vol            
            return tCur
        
        def calculate_QT(t0,Queue):
            tCur=max(t0,t0*(Queue+1)*(1+random.random()/10-0.05))                                   
            return tCur
            
        
        CrNo=int(CrNo)
                
        if CrNo==0: tCur=t0
                 
        elif CrNo==1000: ### Model
            if type==1:                
                tCur=calculate_trapez(Vol,length,length*1000.0/t0)                                
            elif type==3:                
                tCur=calculate_QT(t0,Vol)
            else:
                tCur=t0
                          
        elif CrNo<100:
                                  
            if type==1:                        
                fValue=DataBase.Links.CR.Data[CrNo][3][int(100*Sat)]
                type='multi'            
                if DataBase.Links.CR.Data[CrNo][1]=="HCM3": type='HCM3'  
                if DataBase.Links.CR.Data[CrNo][1]=="EXPONENTIAL": type='plus'            
                if DataBase.Links.CR.Data[CrNo][1]=="LOGISTIC": type='plus'
                if DataBase.Links.CR.Data[CrNo][1]=="QUADRATIC": type='plus'
                if DataBase.Links.CR.Data[CrNo][1]=="Akcelik": type='plus'
                if type=='multi': tCur=int(t0*fValue)
                elif type=='plus': tCur=int(t0+fValue)
                elif type=='HCM3': tCur=int(t0*fValue+(q-qmax)*DataBase.Links.CR.Data[CrNo][2][3])
            if type==2:        
                fValue=DataBase.Nodes.CR.Data[CrNo][3][int(100*Sat)]
                type='multi'            
                if DataBase.Nodes.CR.Data[CrNo][1]=="HCM3": type='HCM3'  
                if DataBase.Nodes.CR.Data[CrNo][1]=="EXPONENTIAL": type='plus'            
                if DataBase.Nodes.CR.Data[CrNo][1]=="LOGISTIC": type='plus'
                if DataBase.Nodes.CR.Data[CrNo][1]=="QUADRATIC": type='plus'
                if DataBase.Nodes.CR.Data[CrNo][1]=="Akcelik": type='plus'
                if type=='multi': tCur=(t0*fValue)
                elif type=='plus': tCur=(t0+fValue)
                elif type=='HCM3': tCur=(t0*fValue+(q-qmax)*DataBase.Nodes.CR.Data[CrNo][2][3])
            if type==3:                     
                fValue=DataBase.Turns.CR.Data[CrNo][3][int(100*Sat)]
                type='multi'            
                if DataBase.Turns.CR.Data[CrNo][1]=="HCM3": type='HCM3'  
                if DataBase.Turns.CR.Data[CrNo][1]=="EXPONENTIAL": type='plus'            
                if DataBase.Turns.CR.Data[CrNo][1]=="LOGISTIC": type='plus'
                if DataBase.Turns.CR.Data[CrNo][1]=="QUADRATIC": type='plus'
                if DataBase.Turns.CR.Data[CrNo][1]=="Akcelik": type='plus'
                if type=='multi': tCur=int(t0*fValue)
                elif type=='plus': tCur=int(t0+fValue)
                elif type=='HCM3': tCur=int(t0*fValue+(q-qmax)*DataBase.Turns.CR.Data[CrNo][2][3])
        
        return tCur
               
            
            
    def Inflow_Handler(Last_Column,T):
        
        tCur=tCur_calculate(Last_Column[8],Last_Column[10],Last_Column[7],Last_Column[3],Last_Column[9],Last_Column[5],Last_Column[11],Last_Column[12])
        
        if Last_Column[9]==0:
            VolCapRatio=0
        else:
            VolCapRatio=float(Last_Column[3])/Last_Column[9]            
        New_Column=[T,1,0,Last_Column[3]+1,Last_Column[4],Last_Column[3]+1-Last_Column[4],tCur,Last_Column[7],Last_Column[8],Last_Column[9],VolCapRatio,Last_Column[11],Last_Column[12]]
        
        return New_Column

    def Outflow_Handler(Last_Column,T):
        New_Column=[T,0,1,Last_Column[3],Last_Column[4]+1,Last_Column[3]-Last_Column[4]-1,Last_Column[6],Last_Column[7],Last_Column[8],Last_Column[9],Last_Column[10],Last_Column[11],Last_Column[12]]
        return New_Column
    
    
    
    T=-1
    noCompleted=0
    print "paths ",len(DataBase.Paths.List)
    
    while noCompleted<len(DataBase.Paths.List):
        
        
        T+=1
        
    
        for no in xrange(len(DataBase.Paths.List)):            
                       
            if DataBase.Paths.List[no][4]==T:
                 
                
                noEl=len(DataBase.Paths.List[no][1])
                flag=DataBase.Paths.List[no][5]
                
                DataBase.Paths.List[no][5]+=1
                
                if flag+1<noEl:
                   
                   if flag==1:
                      
                      """Zone Outflow"""
                      Last_Column=DataBase.Zones.Data[DataBase.Zones.DictionaryVis2Py[DataBase.Paths.List[no][6]]][-1]
                      DataBase.Zones.Data[DataBase.Zones.DictionaryVis2Py[DataBase.Paths.List[no][6]]].append(Outflow_Handler(Last_Column,T))
                      """Orig Connector Inflow"""
                      Last_Column=DataBase.OrigConnectors.Data[DataBase.OrigConnectors.DictionaryVis2Py[(DataBase.Paths.List[no][6],DataBase.Paths.List[no][2][0])]][-1]
                      
                      New_Column=Inflow_Handler(Last_Column,T)
                      DataBase.OrigConnectors.Data[DataBase.OrigConnectors.DictionaryVis2Py[(DataBase.Paths.List[no][6],DataBase.Paths.List[no][2][0])]].append(New_Column)
                      """append tCur of Orig Connector to times at place 1 [0 , tCur]"""
                      
                      DataBase.Paths.List[no][3].append(DataBase.Paths.List[no][3][-1]+New_Column[6])
                      
                      """ set new time head for path = tCur (connector) """
                      DataBase.Paths.List[no][4]=int(DataBase.Paths.List[no][3][-1])
                      if  DataBase.Paths.List[no][3][-1]<=(DataBase.Paths.List[no][3][-2]+1): T-=1
                      
                      """ Orig Connector Outflow """
                      Last_Column=DataBase.OrigConnectors.Data[DataBase.OrigConnectors.DictionaryVis2Py[(DataBase.Paths.List[no][6],DataBase.Paths.List[no][2][0])]][-1]
                      New_Column=Outflow_Handler(Last_Column,DataBase.Paths.List[no][3][-1])
                      DataBase.OrigConnectors.Data[DataBase.OrigConnectors.DictionaryVis2Py[(DataBase.Paths.List[no][6],DataBase.Paths.List[no][2][0])]].append(New_Column)
                      
                      """1st Node Inflow """
                      #DataBase.Nodes.Inflow[DataBase.Paths.List[no][2][flag-1]].append(T)
                      Last_Column=DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][flag-1]]][-1]
                      New_Column=Inflow_Handler(Last_Column,DataBase.Paths.List[no][3][-1])
                      DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][flag-1]]].append(New_Column)
                      
                      
                                         
                   elif flag/2< len(DataBase.Paths.List[no][2]):
                        
                       if divmod(flag,2)[1]==0:
                          
                          if (flag>2 and flag/2<len(DataBase.Paths.List[no][2])):
                              """ Turn Outflow without changing the time head, and without appending to times vector"""
                              
                              Last_Column=DataBase.Turns.Data[DataBase.Turns.DictionaryVis2Py[(DataBase.Paths.List[no][2][flag/2-2],DataBase.Paths.List[no][2][flag/2-1],DataBase.Paths.List[no][2][flag/2])]][-1]
                              New_Column=Outflow_Handler(Last_Column,T)
                              DataBase.Turns.Data[DataBase.Turns.DictionaryVis2Py[(DataBase.Paths.List[no][2][flag/2-2],DataBase.Paths.List[no][2][flag/2-1],DataBase.Paths.List[no][2][flag/2])]].append(New_Column)
                              
                              
                          """Node Outflow """                          
                          Last_Column=DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][flag/2-1]]][-1]
                          New_Column=Outflow_Handler(Last_Column,T)
                          DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][flag/2-1]]].append(New_Column)
                          
                          """Link Inflow """                          
                          Last_Column=DataBase.Links.Data[DataBase.Links.DictionaryVis2Py[(DataBase.Paths.List[no][2][flag/2-1],DataBase.Paths.List[no][2][flag/2])]][-1]
                          New_Column=Inflow_Handler(Last_Column,T)
                          DataBase.Links.Data[DataBase.Links.DictionaryVis2Py[(DataBase.Paths.List[no][2][flag/2-1],DataBase.Paths.List[no][2][flag/2])]].append(New_Column)
                          """append tCur of proceding link to times"""
                          DataBase.Paths.List[no][3].append(DataBase.Paths.List[no][3][-1]+New_Column[6])
                          """ set new time head for path = tCur link """
                          DataBase.Paths.List[no][4]=int(DataBase.Paths.List[no][3][-1])
                          if  DataBase.Paths.List[no][3][-1]<=(DataBase.Paths.List[no][3][-2]+1): T-=1
                          
                       elif divmod(flag,2)[1]==1:
                           
                           if (flag>2 and flag/2<len(DataBase.Paths.List[no][2])-1):
                              
                              """Turn Inflow"""
                              
                              Last_Column=DataBase.Turns.Data[DataBase.Turns.DictionaryVis2Py[(DataBase.Paths.List[no][2][(flag-1)/2-1],DataBase.Paths.List[no][2][(flag-1)/2],DataBase.Paths.List[no][2][(flag-1)/2+1])]][-1]
                              New_Column=Inflow_Handler(Last_Column,T)
                              
                              DataBase.Turns.Data[DataBase.Turns.DictionaryVis2Py[(DataBase.Paths.List[no][2][(flag-1)/2-1],DataBase.Paths.List[no][2][(flag-1)/2],DataBase.Paths.List[no][2][(flag-1)/2+1])]].append(New_Column)
                              """append tCur of turn to times"""                              
                              DataBase.Paths.List[no][3].append(DataBase.Paths.List[no][3][-1]+New_Column[6])
                              """ set new time head for path = tCur (turn) """
                              DataBase.Paths.List[no][4]=int(DataBase.Paths.List[no][3][-1])
                              if  DataBase.Paths.List[no][3][-1]<=(DataBase.Paths.List[no][3][-2]+1): T-=1
                              
                           
                           """Link Outflow"""
                           #DataBase.Links.Outflow[DataBase.Paths.List[no][2][(flag-1)/2-1]][DataBase.Paths.List[no][2][(flag-1)/2]].append(T)
                           Last_Column=DataBase.Links.Data[DataBase.Links.DictionaryVis2Py[(DataBase.Paths.List[no][2][(flag-1)/2-1],DataBase.Paths.List[no][2][(flag-1)/2])]][-1]
                           New_Column=Outflow_Handler(Last_Column,T)
                           DataBase.Links.Data[DataBase.Links.DictionaryVis2Py[(DataBase.Paths.List[no][2][(flag-1)/2-1],DataBase.Paths.List[no][2][(flag-1)/2])]].append(New_Column)
                                                      
                           #DataBase.Nodes.Inflow[DataBase.Paths.List[no][2][(flag-1)/2]].append(T)
                           """Node Inflow"""
                           
                           Last_Column=DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][(flag-1)/2]]][-1]
                           New_Column=Inflow_Handler(Last_Column,T)
                           DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][(flag-1)/2]]].append(New_Column)
                     
                elif flag+1==noEl:
                            noCompleted+=1    
                            if divmod(noCompleted,1)[1]==0:                            
                                print noCompleted,len(DataBase.Paths.List),T
                                label="Path no: "+str(noCompleted)+" calculated; "+str(noCompleted/len(DataBase.Paths.List)*100)+"% "
                                MyFrame.update_console(frame,label)                           
                            
                            #print noCompleted,len(DataBase.Paths.List[no][1]),len(DataBase.Paths.List[no][3]),DataBase.Paths.List[no][6],DataBase.Paths.List[no][7]
                            """ flag completed: -1"""
                            DataBase.Paths.List[no][4]=-10
                            #DataBase.Nodes.Outflow[DataBase.Paths.List[no][2][flag/2-1]].append(T)
                            #DataBase.DestConnectors.Inflow[DataBase.Paths.List[no][2][-1]][DataBase.Paths.List[no][7]].append(T)
                            """Last Link Outflow"""
                            if len(DataBase.Paths.List[no][2])>1:
                                Last_Column=DataBase.Links.Data[DataBase.Links.DictionaryVis2Py[(DataBase.Paths.List[no][2][(flag-1)/2-1],DataBase.Paths.List[no][2][(flag-1)/2])]][-1]
                                New_Column=Outflow_Handler(Last_Column,T)
                                DataBase.Links.Data[DataBase.Links.DictionaryVis2Py[(DataBase.Paths.List[no][2][(flag-1)/2-1],DataBase.Paths.List[no][2][(flag-1)/2])]].append(New_Column)
                                """Last Node Inflow"""
                                Last_Column=DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][-1]]][-1]
                                New_Column=Inflow_Handler(Last_Column,T)
                                DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][-1]]].append(New_Column)
                            """ Last Node Outflow"""
                            Last_Column=DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][-1]]][-1]
                            New_Column=Outflow_Handler(Last_Column,T)
                            DataBase.Nodes.Data[DataBase.Nodes.DictionaryVis2Py[DataBase.Paths.List[no][2][-1]]].append(New_Column)
                            """Dest Connector Inflow"""
                            Last_Column=DataBase.DestConnectors.Data[DataBase.DestConnectors.DictionaryVis2Py[(DataBase.Paths.List[no][2][-1],DataBase.Paths.List[no][7])]][-1]
                            New_Column=Outflow_Handler(Last_Column,T)
                            DataBase.DestConnectors.Data[DataBase.DestConnectors.DictionaryVis2Py[(DataBase.Paths.List[no][2][-1],DataBase.Paths.List[no][7])]].append(New_Column)
                            """Dest Connector Outflow - New Time"""
                            #DataBase.DestConnectors.Outflow[DataBase.Paths.List[no][2][-1]][DataBase.Paths.List[no][7]].append(T)
                            Last_Column=DataBase.DestConnectors.Data[DataBase.DestConnectors.DictionaryVis2Py[(DataBase.Paths.List[no][2][-1],DataBase.Paths.List[no][7])]][-1]
                            New_Column=Outflow_Handler(Last_Column,T)
                            DataBase.DestConnectors.Data[DataBase.DestConnectors.DictionaryVis2Py[(DataBase.Paths.List[no][2][-1],DataBase.Paths.List[no][7])]].append(New_Column)
                            """ append last point to time """
                            DataBase.Paths.List[no][3].append(DataBase.Paths.List[no][3][-1]+New_Column[6])
                            DataBase.Paths.List[no][4]=int(DataBase.Paths.List[no][3][-1])
                            
                            
                            Last_Column=DataBase.Zones.Data[DataBase.Zones.DictionaryVis2Py[DataBase.Paths.List[no][7]]][-1]
                            New_Column=Inflow_Handler(Last_Column,T)
                            DataBase.Zones.Data[DataBase.Zones.DictionaryVis2Py[DataBase.Paths.List[no][7]]].append(New_Column)
                            Results.Paths[no]=[DataBase.Paths.List[no][6],DataBase.Paths.List[no][7],DataBase.Paths.List[no][3][0],DataBase.Paths.List[no][3][-1],DataBase.Paths.List[no][3][-1]-DataBase.Paths.List[no][4]]
    """Last line of time for each element - the network goes to sleep :) 
    
    for i in range(1,len(DataBase.Arrays.Nodes+1)):
        if DataBase.Nodes.Data[i][-1][3]>0:
            print DataBase.Nodes.Data[i][-1][3]
            DataBase.Nodes.Data[i].append([T+100,0,0,DataBase.Nodes.Data[i][-1][3],DataBase.Nodes.Data[i][-1][4],0,DataBase.Arrays.Nodes[i][3]])
    for i in range(1,len(DataBase.Arrays.Zones)+1):
        if DataBase.Zones.Data[i][-1][3]>0:
            print DataBase.Zones.Data[i][-1][3]
            DataBase.Zones.Data[i].append([T+100,0,0,DataBase.Zones.Data[i][-1][3],DataBase.Zones.Data[i][-1][4],0,0])
    for i in range(1,len(DataBase.Arrays.Links)+1):
        if DataBase.Links.Data[i]>0:
            if len(DataBase.Links.Data[i][-1])>2:
                print DataBase.Links.Data[i][-1][3]
                DataBase.Links.Data[int(DataBase.Arrays.Links[i-1][1])][int(DataBase.Arrays.Links[i-1][2])].append([T+100,0,0,DataBase.Links.Data[i][-1][3],DataBase.Links.Data[i][-1][4],0,DataBase.Arrays.Links[i][3]])
    for i in range(1,len(DataBase.Arrays.Turns)+1):
        if len(DataBase.Turns.Data[i])>0:
            print DataBase.Turns.Data[i][-1][3]
            DataBase.Turns.Data[int(DataBase.Arrays.Turns[i-1][0])][int(DataBase.Arrays.Turns[i-1][1])][int(DataBase.Arrays.Turns[i-1][2])].append([T+100,0,0,DataBase.Turns.Data[i][-1][3],DataBase.Turns.Data[i][-1][4],0,DataBase.Arrays.Turns[i][3]])
    for i in range(1,len(DataBase.Arrays.Connectors)+1):
        if DataBase.Arrays.Connectors[i-1][2]==1:
            DataBase.OrigConnectors.Data[int(DataBase.Arrays.Connectors[i-1][0])][int(DataBase.Arrays.Connectors[i-1][1])].append([T+100,0,0,DataBase.OrigConnectors.Data[i][-1][3],DataBase.OrigConnectors.Data[i][-1][4],0,DataBase.Arrays.OrigConnectors[i][3]])
        else:
            DataBase.DestConnectors.Data[int(DataBase.Arrays.Connectors[i-1][1])][int(DataBase.Arrays.Connectors[i-1][0])].append([T+100,0,0,DataBase.DestConnectors.Data[i][-1][3],DataBase.DestConnectors.Data[i][-1][4],0,DataBase.Arrays.DestConnectors[i][3]])
    """

    
def Plot_Cylinder(a,b,c):
    fig = plt.figure()
    PlotSpace = Axes3D(fig)
    Plot_Links(PlotSpace)  
    Plot_Dynamic_Paths(PlotSpace,a,b,c)    
    plt.show()
                      
def Plot_Dynamic_Paths(PlotSpace,filter_type,lower_bound,upper_bound):
    
    for i in xrange(len(DataBase.Paths.List)):
        if filter_type==0:            
            PlotSpace.plot(DataBase.Paths.List[i][0],DataBase.Paths.List[i][1],DataBase.Paths.List[i][3],lw=1.5)
        if filter_type==1:
            if DataBase.Paths.List[i][6]==int(lower_bound):
                if DataBase.Paths.List[i][7]==int(upper_bound):
                    PlotSpace.plot(DataBase.Paths.List[i][0],DataBase.Paths.List[i][1],DataBase.Paths.List[i][3],lw=1.5)
        if filter_type==2:
            if DataBase.Paths.List[i][3][0]>int(lower_bound):
                if DataBase.Paths.List[i][3][0]<int(upper_bound):
                    PlotSpace.plot(DataBase.Paths.List[i][0],DataBase.Paths.List[i][1],DataBase.Paths.List[i][3],lw=1.5)
    
def Plot_Links(PlotSpace):
    """
    Plots Links in 2d surface with width reflecting the volume
    """
    for i in xrange(1,len(DataBase.Arrays.Links)):
        Origin=DataBase.Nodes.DictionaryVis2Py[DataBase.Arrays.Links[i,1]]
        Destination=DataBase.Nodes.DictionaryVis2Py[DataBase.Arrays.Links[i,2]]
        X_O=int(DataBase.Arrays.Nodes[Origin,1])
        Y_O=int(DataBase.Arrays.Nodes[Origin,2])
        X_D=int(DataBase.Arrays.Nodes[Destination,1])
        Y_D=int(DataBase.Arrays.Nodes[Destination,2])
        PlotSpace.plot([X_O,X_D],[Y_O,Y_D],[0,0],c='black',zorder=1,lw=5)

def Save_DataBase(filename):
    
    
    
    
        
    plik=open(filename,'w')
    
    
    DataOutput=[[[DataBase.Nodes.Data,DataBase.Links.Data,DataBase.Turns.Data,DataBase.OrigConnectors.Data,DataBase.DestConnectors.Data,DataBase.Zones.Data],
                 [DataBase.Nodes.DictionaryPy2Vis,DataBase.Links.DictionaryPy2Vis,DataBase.Turns.DictionaryPy2Vis,DataBase.OrigConnectors.DictionaryPy2Vis,DataBase.DestConnectors.DictionaryPy2Vis,DataBase.Zones.DictionaryPy2Vis],
                 [DataBase.Nodes.DictionaryVis2Py,DataBase.Links.DictionaryVis2Py,DataBase.Turns.DictionaryVis2Py,DataBase.OrigConnectors.DictionaryVis2Py,DataBase.DestConnectors.DictionaryVis2Py,DataBase.Zones.DictionaryVis2Py],
                 [DataBase.Paths.List]],[Results.VersionName,Results.noZones,Results.noNodes,Results.noLinks,Results.noTurns,Results.noTrips,
                                         Results.Times.VersionLoad,Results.Times.GetData,Results.Times.GetCoords,Results.Times.CrossTheTime,Results.Times.Plot_Cylinder,
                                         Results.AssType,Results.AssParam,Results.tStart,Results.tEnd,Results.Distribution,Results.tVector,Results.PathExcel]]
                 
    
                
    cPickle.dump(DataOutput, plik, protocol=0)
    plik.close()
    
def Load_DataBase(filename):    
    from numpy.core import multiarray as multiarray
    import cPickle #numpy.multiarray  
    #filename= MyFrame.DataBase_Path.Value
    #filename='D:/results/pampara.lft'
    plik= open(filename, 'r')
    
    
    
    [[[DataBase.Nodes.Data,DataBase.Links.Data,DataBase.Turns.Data,DataBase.OrigConnectors.Data,DataBase.DestConnectors.Data,DataBase.Zones.Data],
                 [DataBase.Nodes.DictionaryPy2Vis,DataBase.Links.DictionaryPy2Vis,DataBase.Turns.DictionaryPy2Vis,DataBase.OrigConnectors.DictionaryPy2Vis,DataBase.DestConnectors.DictionaryPy2Vis,DataBase.Zones.DictionaryPy2Vis],
                 [DataBase.Nodes.DictionaryVis2Py,DataBase.Links.DictionaryVis2Py,DataBase.Turns.DictionaryVis2Py,DataBase.OrigConnectors.DictionaryVis2Py,DataBase.DestConnectors.DictionaryVis2Py,DataBase.Zones.DictionaryVis2Py],
                 [DataBase.Paths.List]],[Results.VersionName,Results.noZones,Results.noNodes,Results.noLinks,Results.noTurns,Results.noTrips,
                                         Results.Times.VersionLoad,Results.Times.GetData,Results.Times.GetCoords,Results.Times.CrossTheTime,Results.Times.Plot_Cylinder,
                                         Results.AssType,Results.AssParam,Results.tStart,Results.tEnd,Results.Distribution,Results.tVector,Results.PathExcel]]= cPickle.load(plik)
    
    

def Check_Results():
    
    noZones=len(DataBase.Arrays.Zones)-1
    noNodes=len(DataBase.Arrays.Nodes)-1
    noLinks=len(DataBase.Arrays.Links)-1
    noTurns=len(DataBase.Arrays.Turns)-1
    noConnectors=len(DataBase.Arrays.Connectors)-1
    #nodes
    #for i in xrange(1,noNodes):
        #print i,Visum.Net.Nodes.ItemByKey(DataBase.Nodes.DictionaryPy2Vis[i]).AttValue("VolPrT"),DataBase.Nodes.Data[i][-1][3]
    for i in xrange(1,noTurns):
        print DataBase.Arrays.Turns[i][3],DataBase.Turns.Data[i][-1][3],Visum.Net.Turns.ItemByKey(DataBase.Turns.DictionaryPy2Vis[i][0],DataBase.Turns.DictionaryPy2Vis[i][1],DataBase.Turns.DictionaryPy2Vis[i][2]).AttValue("VolVehPrT(AP)")

if __name__ == "__main__":    
    app = wx.PySimpleApp(0)
    wx.InitAllImageHandlers()
    frame_1 = MyFrame(None, -1, "")
    frame_1.Maximize()
    app.SetTopWindow(frame_1)
    frame_1.Show()
    app.MainLoop()
    


