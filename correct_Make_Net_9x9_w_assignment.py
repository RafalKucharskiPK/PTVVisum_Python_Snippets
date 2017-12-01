from scipy import *
from random import randint
import win32com.client
from VisumPy.helpers import *
Visum=win32com.client.Dispatch("Visum.Visum")


def distance(pierwszy,drugi):
    pierwszy=pierwszy-1
    drugi=drugi-1
    return sqrt(square(Nodes_Array[pierwszy,1]-Nodes_Array[drugi,1])+square(Nodes_Array[pierwszy,2]-Nodes_Array[drugi,2]))
 
def clearNet():
    nodeno=len(Visum.Net.Nodes.GetMultiAttValues("No"))
    for i in range(nodeno):
        if Visum.Net.Nodes.NodeExistsByKey(nodeno)==1:
            Visum.Net.RemoveNode(i)
    linkno=len(Visum.Net.Links.GetMultiAttValues("No"))
    for i in range(linkno):
        Visum.Net.RemoveNode(i)

def addNodes():
    nodeno=1
    for xcoord in range(9):
        for ycoord in range(9):
            Visum.Net.AddNode(nodeno,xcoord*1000,ycoord*1000)
            nodeno=nodeno+1
        
    Nodes_list=Visum.Net.Nodes.GetMultipleAttributes(("No","XCoord","YCoord"))
    Nodes_Array=zeros([len(Nodes_list),3])
    
    for i in range(len(Nodes_list)):
        for j in range(3):
            Nodes_Array[i,j]=Nodes_list[i][j]
    Nodes_list=0
    return Nodes_Array
        
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
                    
                    linkno=linkno+1
    #Visum.Net.Links.SetAllAttValues("CapPrT", 1000)
    
    
def addZones(): 
    zoneno=1             
    for i in range(2):
        for j in range(2):                 
            Visum.Net.AddZone(zoneno,i*12000-2000,j*12000-2000)
            zoneno=zoneno+1
            
    Zones_list=Visum.Net.Zones.GetMultipleAttributes(("No","XCoord","YCoord"))
    Zones_Array=zeros([len(Zones_list),3])
    for i in range(len(Zones_list)):
        for j in range(3):
            Zones_Array[i,j]=Zones_list[i][j]
    Zones_list=0
    return Zones_Array
            
            
def addConnectors(Zones_Array,Nodes_Array):
    
    def DistanceChecker(x,Zones_Array,Nodes_Array):
                
        Xz=Zones_Array[x-1,1]
        Yz=Zones_Array[x-1,2]
        s=Inf
        najblizsza=-1
        for j in range(81):
            Xn=Nodes_Array[j,1]
            Yn=Nodes_Array[j,2]
            dist=sqrt(square(Xz-Xn)+square(Yz-Yn))
            print(dist)
            if dist <= s:
                s=dist
                najblizsza=j
        return najblizsza
        
    for i in range(1,len(Zones_Array)+1):
        k=DistanceChecker(i,Zones_Array,Nodes_Array)
        print(i,".....",k)
        Visum.Net.AddConnector(i,k+1)           
                
                   
def SetSomeFlows():
    Visum.Net.AddODMatrix(1)
    Matrix=Visum.Net.Matrices.ItemByKey(1)
    
    Matrix.SetValues(ones([4,4])+300)
    Matrix.SetDiagonal(0)
    Visum.Net.DemandSegments.ItemByKey("C").getDemandDescription().SetAttValue("MatrixNo", 1)   
            
def Assignment_Execution():
    Visum.Procedures.Operations.AddOperation(1)
    Assignment=Visum.Procedures.Operations.ItemByKey(1)
    Assignment.SetAttValue("DSegSet","C")    
    Visum.Procedures.Execute()            

#clearNet()  
Nodes_Array=addNodes()
addLinks()
Zones_Array=addZones()
addConnectors(Zones_Array,Nodes_Array)
SetSomeFlows()
Assignment_Execution()
      
