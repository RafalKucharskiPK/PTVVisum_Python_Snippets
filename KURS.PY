from scipy import ones
import random
from VisumPy.helpers import *

def Podstawy_Python():
    #liczby
    a=3 #deklarowanie zmiennej, nie trzeba podawac jej typu
    A
    A=3
    
    A==a
    A>a
    a<A
    
    a+a
    a-a
    a*a
    a**2
    a**.5
    a/2
    a=3.0
    a/2
    a=(50-5*6)/4

    
    #tekst
    a='Dzien dobry'
    a+a
    a[0]
    a[:8]
    a[5:8]
    a[-1]
    len(a)
    a.find('Dd')
    a.find('dob')
    
    
    #obliczenia
    dlugosc=30
    szerokosc=40
    wysokosc=50
    pole=dlugosc*szerokosc
    objetosc=pole*wysokosc
    
    # list - bardzo wazny typ w pythonie wektor
    
    Lista=[2,3,4,10,100]
    Lista[0]
    Lista[2]
    Lista[-1]
    Lista[4]=2222
    Lista
    len(Lista)
    Lista.append(3)
    Lista.append([4])
    #lista w liscie
    
    Lista[3]=[300,400,500,'trtrtr']
    Lista
    
    # warunek IF
    
    dlugosc=len(Lista)
    
    if dlugosc==4:
        print 'hej hej hej'  #indent ! ! ! ! ! ! !
    elif dlugosc>5:
        print 'bum bum bum'
    else: print 'trach trach trach'
    
    #range
    
    range(10)
    range(5,10)
    
    #petla for
    
    for element_listy in Lista:
        print element_listy
        
    for i in range(1,10):
        print i
    
    for i in range(len(Lista)):
        print Lista[i]
        
    Lista=[1,2,3,4]
    for element_listy in Lista:
         element_listy=2*element_listy
    
    Lista
    
    for i in range(len(Lista)):
        Lista[i]=2*Lista[i]
        
    Lista

    #definiowanie funkcji
  
  
def Init(path):
        import win32com.client    
        Visum=win32com.client.Dispatch('Visum.Visum')
        Visum.LoadVersion(path)
        return Visum

def Visum_COM(Visum): 
    
    
               
    def distance(Node1,Node2):
        Node1=[Visum.Net.Nodes.ItemByKey(Node1).AttValue('XCoord'),Visum.Net.Nodes.ItemByKey(Node1).AttValue('YCoord')]
        Node2=[Visum.Net.Nodes.ItemByKey(Node2).AttValue('XCoord'),Visum.Net.Nodes.ItemByKey(Node2).AttValue('YCoord')]
        ret=(Node1[0]-Node2[0])**2+(Node1[1]-Node2[1])**2
        ret=ret**.5
        return ret
        
     
    
    def addNodes():
        nodeno=1
        for xcoord in range(9):
            for ycoord in range(9):
                Visum.Net.AddNode(nodeno,xcoord*1000,ycoord*1000)
                nodeno=nodeno+1
    
    def addLinks():
        linkno=1
        for x in range(1,82):
            for y in range(1,82):  
                if x!=y and distance(x,y)<1500 and Visum.Net.Links.LinkExistsByKey(x,y)==0:                
                        Visum.Net.AddLink(linkno,x,y)
                        Visum.Net.Links.ItemByKey(x,y).SetAttValue("V0PrT",30+random.randint(1,30))
                        Visum.Net.Links.ItemByKey(y,x).SetAttValue("V0PrT",Visum.Net.Links.ItemByKey(x,y).AttValue("V0PrT"))
                                        
                        Visum.Net.Links.ItemByKey(x,y).SetAttValue("CapPrT",300+random.randint(1,300))
                        Visum.Net.Links.ItemByKey(y,x).SetAttValue("CapPrT",Visum.Net.Links.ItemByKey(x,y).AttValue("CapPrT"))
                        
                        linkno+=1
    
    def addZones(): 
        zoneno=1             
        for i in range(2):
            for j in range(2):                 
                Visum.Net.AddZone(zoneno,i*12000-2000,j*12000-2000)
                zoneno=zoneno+1
                
    def addConnectors():
       Polaczenia={1:1, 2:9, 4:81, 3:72}
       ZoneIter=Visum.Net.Zones.Iterator        
       while ZoneIter.Valid:
            Zone=ZoneIter.Item            
            ZoneNo=Zone.AttValue('No')
            Visum.Net.AddConnector(ZoneNo,Polaczenia[ZoneNo])            
            ZoneIter.Next()
       ZoneIter.Reset()             
    
    def SetSomeFlows():
        Visum.Net.AddODMatrix(1)
        Matrix=Visum.Net.Matrices.ItemByKey(1)
        
        Mtx=scipy.zeros([4,4])
        Mtx+=30
        Mtx[2,3]=80
        
        
        Matrix.SetValues(Mtx)        
        Visum.Net.DemandSegments.ItemByKey("C").getDemandDescription().SetAttValue("MatrixNo", 1)     
           
    def Assignment_Execution():
        Visum.Procedures.Operations.AddOperation(1)
        Assignment=Visum.Procedures.Operations.ItemByKey(1)
        Assignment.SetAttValue("DSegSet","C")    
        Visum.Procedures.Execute()        
    
    
    def Calculate_Skim_Matrix():
        Operations=Visum.Procedures.Operations.GetAll
        
        for Operation in Operations:
            Operation.SetAttValue("Active",0)          
            
        Visum.Procedures.Operations.AddOperation(1)
        CalcSkimOper=Visum.Procedures.Operations.ItemByKey(1)
        CalcSkimOper.SetAttValue("OperationType", 2.0)
        CalcSkimOper.SetAttValue("PrTAssignment","PrtSkimMatrix")
        CalcSkimOper.SetAttValue("DSegSet", "C")
        CalcSkimOper=Visum.Procedures.Operations.ItemByKey(1)
        Visum.Procedures.Operations.ItemByKey(1).PrTSkimMatrixParameters.SingleSkimMatrixParameters("T0").SetAttValue("Calculate",1)
        Visum.Procedures.Execute()
    
    addNodes()  
    addLinks()
    addZones()
    addConnectors()
    SetSomeFlows()
    Assignment_Execution()
    Calculate_Skim_Matrix()
    
try:
        Visum
except:
        Visum=Init()

Visum_COM(Visum)
