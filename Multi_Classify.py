"""
 _   _____  
| | /___  \     Intelligent Infrastructure
| |  ___| |     script created by: Rafal Kucharski
| | /  ___/     16/08/2011
| | | |___      info: info@intelligent-infrastructure.eu
|_| |_____|     Copyright (c) Intelligent Infrastructure 2011 

references: ExNumerus.blogspot.com + QHull www.qhull.org

=====================
Dependencies:
 
GUI created with wxGlade
=====================
 
==========================
End-User License Agreement:
===========================
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

try:
    import wx
    import os
    import VisumPy.helpers 
    import win32api
       
except:    
    import win32api
    win32api.MessageBox(0, "Welcome! \n\nThere seems to be a problem with your python package. \n \n To use Multi Classify you probably need to reinstall your python package \n(best way is via Visum install package)" , 'Multi Classify by intelligent infrastructure')
     
try:    
    import VisumPy.helpers
except:
    win32api.MessageBox(0, "Welcome! \n\n You should reinstall Visum to use scripts." , 'Multi Classify by intelligent infrastructure')
 
    



def Init(path=None):
        import win32com.client        
        Visum=win32com.client.Dispatch('Visum.Visum')
        if path!=None: Visum.LoadVersion(path)
        return Visum

try: 
    Visum
except: 
    Visum=Init('D:/makenet.ver')



def Initialize():
    '''
    Create paths to working folder, html and png file
    in standalone version script opens visum file with selected path via COM
    '''
        
    Paths={}
    Paths["MainVisum"] = Visum.GetWorkingFolder()
    Paths["ScriptFolder"] = Paths["MainVisum"] + "\\AddIns\\MultiClassify"
    Visum.SetPath(48, Paths["ScriptFolder"])
    Paths["Logo"]=Paths["ScriptFolder"] + "\\help\\i2_logo.png"    
    Paths["Help"]=Paths["ScriptFolder"] + "\\help.htm"
    Paths["Icon"]=Paths["ScriptFolder"] + "\\help\\i2_icon.png"
    Paths["CODFile"]=Paths["ScriptFolder"]+"\\_delscript.cod"
    Paths["FMA_OD"]=Paths["ScriptFolder"]+"\\_delscriptOD.fma"
    Paths["FMA_SKIM"]=Paths["ScriptFolder"]+"\\_delscriptSkim.fma"
    return Paths

Paths=Initialize()




Version_Files = False
Segmentsdel=Visum.Net.DemandSegments.GetMultiAttValues("code")
Segments=[]
for Segment in Segmentsdel:
    Segments.append(Segment[1])


class MyDialog(wx.Dialog):
    def __init__(self, *args, **kwds):
        self.default=True
        # begin wxGlade: MyDialog.__init__
        kwds["style"] = wx.DEFAULT_DIALOG_STYLE
        wx.Dialog.__init__(self, *args, **kwds)        
        self.srodek_staticbox = wx.StaticBox(self, -1, "for specified intervals")
        self.gora_staticbox = wx.StaticBox(self, -1, "classify matrices")
        self.select_version_txt = wx.StaticText(self, -1, "of selected version files")
        self.button_1_copy = wx.Button(self, -1, ".ver")
        self.select_DSeg_txt = wx.StaticText(self, -1, "for selected DSegs")
        self.list_box_1 = wx.ListBox(self, -1, size=(20,35), choices=Segments, style=wx.LB_MULTIPLE)
        self.list_box_1.SetSelection(0)
        self.default_radiobutton = wx.RadioButton(self, -1, "default", style=wx.RB_GROUP)
        self.user_defined_radiobutton = wx.RadioButton(self, -1, "user defined", style=wx.RB_GROUP)
        self.lbound_txt = wx.StaticText(self, -1, "lower bound")
        self.ubound_txt = wx.StaticText(self, -1, "upper bound")
        self.widthtxt = wx.StaticText(self, -1, "width")
        self.lbound_box = wx.TextCtrl(self, -1, "")
        self.ubound_box = wx.TextCtrl(self, -1, "")
        self.width_box = wx.TextCtrl(self, -1, "")
        self.HelpBtn = wx.Button(self, -1, "Help")
        self.logo = wx.StaticBitmap(self, -1, wx.Bitmap(Paths["Logo"], wx.BITMAP_TYPE_ANY))
        self.ClassifyBtn = wx.Button(self, -1, "Classify")
        self.CancelBtn = wx.Button(self, -1, "Cancel")

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_BUTTON, self.Select_Version, self.button_1_copy)
        self.Bind(wx.EVT_RADIOBUTTON, self.RadioClick, self.default_radiobutton)
        self.Bind(wx.EVT_RADIOBUTTON, self.RadioClick, self.user_defined_radiobutton)
        self.Bind(wx.EVT_BUTTON, self.HelpClick, self.HelpBtn)
        self.Bind(wx.EVT_BUTTON, self.ClassifyClick, self.ClassifyBtn)
        self.Bind(wx.EVT_BUTTON, self.Cancel_Click, self.CancelBtn)
        # end wxGlade

    def __set_properties(self):
        # begin wxGlade: MyDialog.__set_properties
        self.SetTitle("Multi Classify by intelligent infrastructure")
        _icon = wx.EmptyIcon()
        _icon.CopyFromBitmap(wx.Bitmap(Paths["Logo"], wx.BITMAP_TYPE_ANY))
        self.SetIcon(_icon)
        self.SetSize((500, 302))
        self.SetFont(wx.Font(10, wx.DEFAULT, wx.NORMAL, wx.NORMAL, 0, ""))
        self.select_version_txt.SetMinSize((160, -1))             
        self.select_DSeg_txt.SetMinSize((160, -1))
        self.default_radiobutton.SetMinSize((-1,-1))
        self.default_radiobutton.SetValue(1)
        self.user_defined_radiobutton.SetMinSize((-1,-1))
        self.lbound_txt.Enable(False)
        self.ubound_txt.Enable(False)
        self.widthtxt.Enable(False)
        self.lbound_box.Enable(False)
        self.ubound_box.Enable(False)
        self.width_box.Enable(False)
        self.HelpBtn.SetMinSize((87, -1))
        self.button_1_copy.SetMinSize((87, -1))
        self.logo.SetMinSize((200,23))
        self.ClassifyBtn.SetMinSize((87, -1))
        self.CancelBtn.SetMinSize((87, -1))
        self.user_defined_radiobutton.SetValue(False)
        # end wxGlade
        

    def __do_layout(self):
        # begin wxGlade: MyDialog.__do_layout
        glowny = wx.BoxSizer(wx.VERTICAL)
        sizer_5 = wx.BoxSizer(wx.HORIZONTAL)
        srodek = wx.StaticBoxSizer(self.srodek_staticbox, wx.VERTICAL)
        grid_sizer_1 = wx.GridSizer(2, 3, 0, 0)
        sizer_1 = wx.BoxSizer(wx.HORIZONTAL)
        gora = wx.StaticBoxSizer(self.gora_staticbox, wx.VERTICAL)
        sizer_3_copy_copy_copy = wx.BoxSizer(wx.HORIZONTAL)
        sizer_3_copy_copy_1 = wx.BoxSizer(wx.HORIZONTAL)
        sizer_3_copy_copy_1.Add(self.select_version_txt, 3, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        sizer_3_copy_copy_1.Add(self.button_1_copy, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        gora.Add(sizer_3_copy_copy_1, 1, wx.ALL|wx.EXPAND, 0)
        sizer_3_copy_copy_copy.Add(self.select_DSeg_txt, 3, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        sizer_3_copy_copy_copy.Add(self.list_box_1, 3, wx.ALL|wx.ALIGN_BOTTOM, 10)
        gora.Add(sizer_3_copy_copy_copy, 6, wx.ALL|wx.EXPAND, 0)
        glowny.Add(gora, 12, wx.ALL|wx.EXPAND, 5)
        sizer_1.Add(self.default_radiobutton, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        sizer_1.Add(self.user_defined_radiobutton, 1, wx.ALL|wx.ALIGN_BOTTOM, 10)
        srodek.Add(sizer_1, 1, wx.EXPAND, 0)
        grid_sizer_1.Add(self.lbound_txt, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        grid_sizer_1.Add(self.ubound_txt, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        grid_sizer_1.Add(self.widthtxt, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        grid_sizer_1.Add(self.lbound_box, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        grid_sizer_1.Add(self.ubound_box, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        grid_sizer_1.Add(self.width_box, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        srodek.Add(grid_sizer_1, 2, wx.EXPAND, 0)
        glowny.Add(srodek, 12, wx.ALL|wx.EXPAND, 5)
        sizer_5.Add(self.HelpBtn, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 20)
        sizer_5.Add(self.logo, 3, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        sizer_5.Add(self.ClassifyBtn, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        sizer_5.Add(self.CancelBtn, 1, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        glowny.Add(sizer_5, 4, wx.ALL|wx.EXPAND, 10)
        self.SetSizer(glowny)
        self.Layout()
        # end wxGlade
    
    def Select_Version(self, event): # wxGlade: MyDialog.<event_handler>
        global Version_Files    
        dialog = wx.FileDialog ( None, message = 'Select Version Files', wildcard = "*.ver", style = wx.OPEN | wx.MULTIPLE )

        if dialog.ShowModal() == wx.ID_OK:
            Version_Files = dialog.GetPaths()            
        dialog.Destroy()

    def RadioClick(self, event): # wxGlade: MyDialog.<event_handler>
        
        if self.default==False:
            self.lbound_txt.Enable(False)
            self.ubound_txt.Enable(False)
            self.widthtxt.Enable(False)
            self.lbound_box.Enable(False)
            self.ubound_box.Enable(False)
            self.width_box.Enable(False)
            self.user_defined_radiobutton.SetValue(False)
            self.default_radiobutton.SetValue(True)
            
        else:
            self.lbound_txt.Enable(True)
            self.ubound_txt.Enable(True)
            self.widthtxt.Enable(True)
            self.lbound_box.Enable(True)
            self.ubound_box.Enable(True)
            self.width_box.Enable(True)            
            self.user_defined_radiobutton.SetValue(True)
            self.default_radiobutton.SetValue(False)
            
        self.default=not self.default
        
        

    def HelpClick(self, event): # wxGlade: MyDialog.<event_handler>
        os.startfile(Paths["Help"])

    

    def Cancel_Click(self, event): # wxGlade: MyDialog.<event_handler>
        self.Destroy()
        
    
    def ClassifyClick(self, event): # wxGlade: MyDialog.<event_handler>
        
        import win32com.client
        
        def SingleVersion(Version_File,Chosen_Segments,Excel):
    
            #"""load version"""
            if Version_File != False:
                Visum.LoadVersion(Version_File)
            
            #"""deactivate all operations"""
            Operations=Visum.Procedures.Operations.GetAll    
            for Operation in Operations:
                Operation.SetAttValue("Active",0) 
            
            #"""For each chosen demand segment calculate skim matrix   t0!"""
            for Segment in Chosen_Segments:
                Visum.Procedures.Operations.AddOperation(1)
                CalcSkimOper=Visum.Procedures.Operations.ItemByKey(1)
                CalcSkimOper.SetAttValue("OperationType", 103.0)
                CalcSkimOper.SetAttValue("PrTAssignment","PrtSkimMatrix")
                CalcSkimOper.SetAttValue("DSegSet", Segment)
                CalcSkimOper=Visum.Procedures.Operations.ItemByKey(1)        
                Visum.Procedures.Operations.ItemByKey(1).PrTSkimMatrixParameters.SingleSkimMatrixParameters("TRIPDIST").SetAttValue("Calculate",1)
            Visum.Procedures.Execute()
            
            for Segment in Chosen_Segments:
             #   """save matrix"""
                try:
                    Mtx=Visum.Net.DemandSegments.ItemByKey(Segment).ODMatrix
                    Mtx.Save(Paths["FMA_OD"],1)
                      #  """load matrix to MUULI"""                    
                    Visum.MatrixEditor.MLoad(Paths["FMA_OD"])
                                             
                   # """save skim matrix"""
                    no=VisumPy.helpers.skimLookup(Visum,"DIS",Segment)
                    Mtx=Visum.Net.SkimMatrices.ItemByKey(no)
                    Mtx.Save(Paths["FMA_SKIM"],1)
                    
                    #"""set limits"""
                    if self.default_radiobutton.Value==True:                 
                        Low=0
                        Upp=round(max(max(Mtx.GetValues())))
                        Int=round(Upp/10)
                    else:
                        Low=int(self.lbound_box.Value)      
                        Upp=int(self.ubound_box.Value)
                        Int=int(self.width_box.Value) 
                            
                    Visum.MatrixEditor.MClassifyWithMatrixUsingIntervals(Paths["FMA_SKIM"], Low, Upp, Int, Paths["CODFile"])
                    
                  #  """open COD file"""
                    CODFile=open(Paths["CODFile"])
                    """read from COD file"""
                    Linie=CODFile.readlines()
                    Linienowe=[]
                    for i in range(3,len(Linie)):
                        Linienowe.append(Linie[i].split(";")) 
                    if Version_File != False:
                        nazwa_pliku=os.path.splitext(os.path.basename(Version_File))[0]
                        Excel.ActiveWorkBook.Sheets("empty").Copy(Excel.ActiveWorkBook.Sheets("empty"))   
                        Excel.ActiveWorkBook.ActiveSheet.Name = nazwa_pliku+"_"+Segment
                    else: 
                        Excel.ActiveWorkBook.Sheets("empty").Copy(Excel.ActiveWorkBook.Sheets("empty")) 
                        Excel.ActiveWorkBook.ActiveSheet.Name = Segment         
                    
                    
                    Cells=Excel.ActiveWorkBook.ActiveSheet.Cells
                    
                   # """excel paste""
                    for i in range(len(Linienowe)):
                        for j in range(len(Linienowe[i])):
                            if Linienowe[i][j]=="MIN": wartosc=0
                            elif Linienowe[i][j]=="MAX": wartosc=99999999999
                            else: wartosc=Linienowe[i][j]
                            Cells(i+1,j+1).Value=wartosc
                except:
                    win32api.MessageBox(0, "Error \n\nNo matrix assigned to DSeg "+Segment , 'Multi Classify by intelligent infrastructure')
    
                                
                
        
        Selections=self.list_box_1.GetSelections()
        if len(Selections)>0:
            Chosen_Segments=[]
            for Selection in Selections:
                Chosen_Segments.append(Segments[Selection])        
            try: 
                Excel=win32com.client.Dispatch("Excel.Application")                               
            except:
                win32api.MessageBox(0, "Excel COM object not registered - try reinstalling Excel, and this should not happen")
                return
             
            WorkBook=Excel.Workbooks.Open(Paths["ScriptFolder"]+"\\template.xls")     
            
            if Version_Files==False:
                SingleVersion(False,Chosen_Segments,Excel)
            else:
                for Version_File in Version_Files:
                    SingleVersion(Version_File,Chosen_Segments,Excel)
            ##WorkBook.Sheets("empty").Delete() - doesn't work ...
            Excel.Visible=1
            os.remove(Paths["CODFile"])
            os.remove(Paths["FMA_OD"])
            os.remove(Paths["FMA_SKIM"])
        else:
            win32api.MessageBox(0, "Error \n\nSelect at least one DSeg" , 'Multi Classify by intelligent infrastructure')
    
            
        
    

   
    

    



wx.InitAllImageHandlers()
dialog_1 = MyDialog(None, -1, "")
app.SetTopWindow(dialog_1)
dialog_1.Show()

    
