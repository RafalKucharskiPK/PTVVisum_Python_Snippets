import win32com.client
Visum=win32com.client.Dispatch("Visum.Visum")
Visum.LoadVersion("C:\grav.ver")




def ExcelInit():
    Excel=win32com.client.Dispatch("Excel.Application")   
    Excel.Visible = 1
    Excel.Workbooks.Add()
    return Excel   


def ZrobHistogramIZapisz():
    Mtx=Visum.Net.Matrices.ItemByKey(2000)
    Mh=Visum.CreateMatrixHistogram(Mtx)    
    Mh.OpenLayout("C:\\intervals.att")
    Mh.Update()
    Mh.SaveDistributionIntervalsForMatrix(2000,"C:\tempres.att")
    
def OtworzWynikiHistogramu():
    Res=open("C:\tempres.att")
    a=[]
    for i in range(1, 37):
        r = Res.readline()
        if i>13:
            a.append(r.split("\t")[2].replace(".",","))
    return a

def WrzucDoExcela(kolumna,wartosci):
    for j,row in enumerate(wartosci):               
        Excel.Cells(kolumna,j).Value = row

def Parametry_Grawitacyjnej(i):
    Proc = Visum.Procedures.Operations.ItemByKey(i)
    for DStrat in ["D-I","D-P","D-N","N-D","P-D","I-D","NZD"]:
        DS = Proc.TripDistributionParameters.TripDistributionDStratParameters(Dstrat)
        
    

def MainLoop():    
    kolumna=2 
    ZrobHistogramIZapisz()    
    wartosci= OtworzWynikiHistogramu()
    WrzucDoExcela(kolumna,wartosci) 
          
        
Excel = ExcelInit() 
MainLoop()

        
        


