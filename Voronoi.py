"""
 _   _____  
| | /___  \     Intelligent Infrastructure
| |  ___| |     script created by: Rafal Kucharski
| | /  ___/     16/08/2011
| | | |___      info: info@intelligent-infrastructure.eu
|_| |_____|     Copyright (c) Intelligent Infrastructure 2011 



=====================
Dependencies: qvoronoi.exe from qhull.org package
GUI created by means of wxglade (www.wxglade.org)
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


import wx

# begin wxGlade: extracode
# end wxGlade

def Init(path=None):
            import win32com.client        
            Visum=win32com.client.Dispatch('Visum.Visum')
            if path!=None: Visum.LoadVersion(path)
            return Visum

def Calc_Main():
    def voronoi2D(xpt,ypt,cpt=None,threshold=0):
        '''
        This function returns a list of line segments which describe the voronoi
            cells formed by the points in zip(xpt,ypt).
         
        If cpt is provided, it identifies which cells should be returned.
            The boundary of the cell about (xpt[i],ypt[i]) is returned 
                if cpt[i]<=threshold.
                 
        This function requires qvoronoi.exe in the working directory. 
        The working directory must have permissions for read and write access.
        This function will leave 2 files in the working directory:
            data.txt
            results.txt
        This function will overwrite these files if they already exist.
        
        copyrights: ExNumerus http://exnumerus.blogspot.com/2011/02/rough-draft-how-to-generate-voronoi.html
        '''
        os.chdir(cwd)        
        if cpt is None:
            # assign a value to cpt for later use
            cpt = [0 for x in xpt]
         
        # write the data file
        pts_filename = 'data.txt'
        pts_F = open(pts_filename,'w')
        pts_F.write('2 # this is a 2-D input set\n')
        pts_F.write('%i # number of points\n' % len(xpt))
        for i,(x,y) in enumerate(zip(xpt,ypt)):
            pts_F.write('%f %f # data point %i\n' % (x,y,i))
        pts_F.close()
     
        # trigger the shell command
        import subprocess
        p = subprocess.Popen('qvoronoi TI data.txt TO results.txt p FN Fv QJ', shell=True)
        p.wait()
     
        # open the results file and parse results
        results = open('results.txt','r')
        
        
     
        # get 'p' results - the vertices of the voronoi diagram
        
        data = results.readline()
        
        voronoi_x_list = []
        voronoi_y_list = []
        data = results.readline()
        for i in range(0,int(data)):
            data = results.readline()
            xx,yy,dummy = data.split(' ')    
            voronoi_x_list.append(float(xx))
            voronoi_y_list.append(float(yy))
             
        # get 'FN' results - the voronoi edges
        data = results.readline()
        voronoi_idx_list = []
        for i in range(0,int(data)):
            data = results.readline()
            this_list = data.split(' ')[:-1]
            for j in range(len(this_list)):
                this_list[j]=int(this_list[j])-1
            voronoi_idx_list.append(this_list[1:])
             
        # get 'FV' results - pairs of points which define a voronoi edge
        # combine these results to build a complete representation of the 
        data = results.readline()
        voronoi_dict = {}
        for i in range(0,int(data)):
            data = results.readline().split(' ')
     
            pair_idx_1 = int(data[1])
            pair_idx_2 = int(data[2])
     
            vertex_idx_1 = int(data[3])-1
            vertex_idx_2 = int(data[4])-1
     
            try:
                voronoi_dict[pair_idx_1].append({ 'edge_vertices':[vertex_idx_1,vertex_idx_2],
                                          'neighbor': pair_idx_2 })
            except KeyError:
                voronoi_dict[pair_idx_1] = [{ 'edge_vertices':[vertex_idx_1,vertex_idx_2],
                                          'neighbor': pair_idx_2 } ]
     
            try:
                voronoi_dict[pair_idx_2].append({ 'edge_vertices':[vertex_idx_1,vertex_idx_2],
                                          'neighbor': pair_idx_1 })
            except KeyError:
                voronoi_dict[pair_idx_2] = [{ 'edge_vertices':[vertex_idx_1,vertex_idx_2],
                                          'neighbor': pair_idx_1 } ]    
     
                         
        #################
        # generate a collection of voronoi cells
        x_list = []
        y_list = []    
        for point_idx in voronoi_dict.keys():
            
                # display this cell, so add the data to the edge list
                e_list = []
                for edge in voronoi_dict[point_idx]:
                    p1_idx = edge['edge_vertices'][0]
                    p2_idx = edge['edge_vertices'][1]
                    e_list.append((p1_idx,p2_idx))
                 
                # put the vertices points in order so they
                #   walk around the voronoi cells
                p_list = [p1_idx]
                while True:
                    p=p_list[-1]
                    for e in e_list:
                        if p==e[0]:
                            next_p = e[1]
                            break
                        elif p==e[1]:
                            next_p = e[0]
                            break
                    p_list.append(next_p)
                    e_list.remove(e)
                    if p_list[0]==p_list[-1]:
                        # the cell is closed
                        break
                     
                # build point list
                x_moje=[]
                y_moje=[]
                if all([p>=0 for p in p_list]):
                    for p in p_list:
                        if p>=0:
                            x_moje.append(voronoi_x_list[p])
                            y_moje.append(voronoi_y_list[p])  
                                  
                x_list.append(x_moje)
                y_list.append(y_moje)
        
        
        del results
        os.remove('results.txt')
        os.remove('data.txt')                 
        return (x_list,y_list)


    def import_Visum():
        xpt=Visum.Net.Zones.GetMultiAttValues('XCoord')
        xpt=[el[1] for el in xpt]
        ypt=Visum.Net.Zones.GetMultiAttValues('YCoord')
        ypt=[el[1] for el in ypt]
        return [xpt,ypt]
           
        
    
        
    def Create_WKT(WKT_X,WKT_Y):
            Nowe=[]
            for el in range(len(WKT_X)):
                if len(WKT_X[el])>0:
                    if None in WKT_X[el]:
                        Nowy=[]
                    else:
                        WKT_X_male=WKT_X[el]
                        WKT_Y_male=WKT_Y[el]
                        
                        Nowy=[[WKT_X_male[i], WKT_Y_male[i]] for i in range(len(WKT_X_male))]
                        
                        Nowy=', '.join([str(x) for x in Nowy])
                        Nowy=Nowy.replace('],','www')
                        Nowy=Nowy.replace(',','')
                        Nowy=Nowy.replace('www',',')
                        Nowy=Nowy.replace('[','')
                        Nowy=Nowy.replace(']','')
                        Nowy='MULTIPOLYGON((('+Nowy+')))'
                    
                else: Nowy=[]
                Nowe.append(Nowy)
            Nowe_tuple=[]
            Stare=Visum.Net.Zones.GetMultiAttValues('WKTSurface')
            for i in range(len(Stare)):                                  
                Nowe_tuple.append((Stare[i][0],Nowe[i]))                
            Nowe_tuple=tuple(Nowe_tuple)
            Visum.Net.Zones.SetMultiAttValues('WKTSurface',Nowe_tuple)
            
        
    
       
    
    
    [ZoneX,ZoneY]=import_Visum()  
    (WKT_X,WKT_Y) = voronoi2D(ZoneX,ZoneY)  
    Create_WKT(WKT_X,WKT_Y)
    

class MyFrame(wx.Frame):
    def __init__(self, *args, **kwds):
        # begin wxGlade: MyFrame.__init__
        kwds["style"] = wx.DEFAULT_FRAME_STYLE
        wx.Frame.__init__(self, *args, **kwds)
        self.label_1 = wx.StaticText(self, -1, "Script creating Voronoi boundaries for Visum Zones. \n\nActual boundaries for Zones will be replaced. \nBounds for marginal zones will not be calculated. \n\nReferences qHull (www.qhull.org) and ExNumerus (http://exnumerus.blogspot.com) \n\nFor more information see: http://en.wikipedia.org/wiki/Voronoi")
        self.button_1 = wx.Button(self, -1, "Calculate")

        self.__set_properties()
        self.__do_layout()

        self.Bind(wx.EVT_BUTTON, self.go, self.button_1)
        # end wxGlade

    def __set_properties(self):
        # begin wxGlade: MyFrame.__set_properties
        self.SetTitle("frame_1")
        self.SetBackgroundColour(wx.SystemSettings_GetColour(wx.SYS_COLOUR_MENU))
        # end wxGlade

    def __do_layout(self):
        # begin wxGlade: MyFrame.__do_layout
        grid_sizer_1 = wx.GridSizer(2, 1, 0, 0)
        grid_sizer_1.Add(self.label_1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 10)
        grid_sizer_1.Add(self.button_1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.ALIGN_CENTER_VERTICAL, 0)
        self.SetSizer(grid_sizer_1)
        grid_sizer_1.Fit(self)
        self.Layout()
        # end wxGlade

    def go(self, event): # wxGlade: MyFrame.<event_handler>
        Calc_Main()        
        self.Destroy()
        

# end of class MyFrame


if __name__ == "__main__":
    import os
    cwd=os.getcwd()           
    
      
    wx.InitAllImageHandlers()
    frame_1 = MyFrame(None, -1, "")
    app.SetTopWindow(frame_1)
    frame_1.Show()
     
      
    
