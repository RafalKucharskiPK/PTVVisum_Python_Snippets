""" 
 _   _____  
| | /___  \     Intelligent Infrastructure
| |  ___| |     script created by: Rafal Kucharski
| | /  ___/     12/01/2012
| | | |___      info: info@intelligent-infrastructure.eu
|_| |_____|     Copyright (c) Intelligent Infrastructure 2012 


=====================
Dependencies:
 
None
=====================
 
==========================
End-User License Agreement:

free script snippet to be used within PTV VISUM package.
Not to be sold, or shared without this preamble.

For more scripts visit: 
www.intelligent-infrastructure.eu 
or contact me at:
info@intelligent-infrastructure.eu

Extended support GIS, .SHP data handling in Visum also available at i2.
===========================

===========================
USAGE:

1. Save your .ver back-up copy, just in case
2. Specify snap radius below
2. Drag and drop onto your Visum window
3. Minimize it (as it modifies network 
"""
def Init(path=None):        
        import win32com.client         
        Visum=win32com.client.Dispatch('Visum.Visum') 
        if path!=None: Visum.LoadVersion(path) 
        return Visum
    
try: #nieistotne
    Visum
except: 
    Visum=Init('D:/A.ver')  
    
    


 
def Dolacz_Slepe():     
    radius=0.05 #specify your snap radius here - in KM   
    Nodes=Visum.Net.Nodes.GetMultipleAttributes(["NumLinks","No","XCoord","YCoord"])    
    leng=len(Nodes)
    i=0
    for Node in Nodes:
        i+=1
        if Node[0]==1:            
                        
            Visum.Filters.InitAll()
            LinkFilter = Visum.Filters.LinkFilter()                  
            LinkFilter.AddCondition("OP_NONE", False, "FromNodeNo",9,Node[1])
            LinkFilter.AddCondition("OP_OR", False, "ToNodeNo",9,Node[1])
            LinkFilter.Complement= True
            LinkFilter.UseFilter = True      
            Result=Visum.Net.GetNearestLink(Node[2],Node[3],radius,True)
            if Result[0]!=None:
                print "Dolacz Slepe: ",i,leng
                try:
                   Result[0].SplitViaNode(Node[1])        
                except:
                    print Result[0].AttValue("FromNodeNo"),Result[0].AttValue("ToNodeNo"),Node[1]
    del Nodes


def Zlacz_Slepe():
    
    def Dist(Node1,Node2):
        if (abs(Node1[2]-Node2[2])<Rkm and abs(Node1[3]-Node2[3])<Rkm):    # tu moze byc blad - czy nie zawezamy pola poszukiwan?     
            return ((Node1[2]-Node2[2])**2+(Node1[3]-Node2[3])**2)**0.5           
        else:
            return 99999999999
    
    
    Rkm=100
    Nodes=Visum.Net.Nodes.GetMultipleAttributes(["NumLinks","No","XCoord","YCoord"])
    l=len(Nodes)
    i=0
    for Node1 in Nodes:
        i+=1
        
        for Node2 in Nodes:
            if Node1[1]!=Node2[1]:
                if (max(Node1[0],Node2[0])<3 and min(Node1[0],Node2[0])<2):
                    print "Zlacz Slepe:", i,len
                    if Dist(Node1,Node2)<Rkm:
                        print Node1[1],Node2[1]                        
                        try:
                            Visum.Net.Nodes.Merge(Node1[1],Node2[1],True,False)
                            print "Merged"
                        except:
                            print "Merge Error"

def Kategoryzacja_Links():
    print 'jestem'
    types=[]
    i=1
    Types=[]
    Links=Visum.Net.Links.GetMultipleAttributes(["FromNodeNo","ToNodeNo","TYPE","TypeNo"])
    l=len(Links)
    for Link in Links:
        
        if Link[2] not in types:
            types.append(Link[2])        
        Types.append((i,types.index(Link[2])))
        i+=1
        print i,l    
    Visum.Net.Links.SetMultiAttValues("TypeNo",Types,False)
    i=0
    print types
    for Ltyp in types:
        Visum.Net.LinkTypes.ItemByKey(i).SetAttValue("Name",Ltyp)
        i+=1
                
def Przetnij_Krzyzujace_Sie():
    
    import shapely.wkt
    maxNodeNo=max([a[1] for a in Visum.Net.Nodes.GetMultiAttValues("No", False)])
    Links=Visum.Net.Links.GetMultipleAttributes(["FromNodeNo","ToNodeNo","WKTPoly","Length"])
    i=0
    l=len(Links)
    for Link1 in Links:
        i+=1
        print "Przeciecia:",i,l
        for Link2 in Links:
            if min(Link1[3],Link2[3])>1:           
                try:
                    
                    L1=shapely.wkt.loads(str(Link1[2]))
                    L2=shapely.wkt.loads(str(Link2[2]))
                    if L1.crosses(L2):
                        Link1_=Visum.Net.Links.ItemByKey(Link1[0],Link1[1])
                        Link2_=Visum.Net.Links.ItemByKey(Link2[0],Link2[1])
                        try:
                            Bridge=Link1_.AttValue("tags")
                            if Bridge.rfind("bridge")<0:                            
                                print "nie most"
                                iPoint=L1.intersection(L2)
                                iPoint=[iPoint.coords.xy[0][0],iPoint.coords.xy[1][0]]
                                try:
                                    maxNodeNo+=1
                                    Visum.Net.AddNode(maxNodeNo,iPoint[0],iPoint[1])                          
                                except:
                                    try:
                                        maxNodeNo+=10
                                        Visum.Net.AddNode(maxNodeNo,iPoint[0],iPoint[1])
                                    except:
                                        print "nie da sie dodac node"  
                                          
                                try:
                                    Link1_.SplitViaNode(maxNodeNo)
                                    Link2_.SplitViaNode(maxNodeNo)
                                    print maxNodeNo
                                except:
                                    print "error nie udal sie split"
                            else:
                                print "most"
                        except:
                            print "error"
                except:
                    print "nie udalo sie przeciecie"
                    
# 
                 
Visum.Graphic.ShowMinimized()
#Dolacz_Slepe()            
#Zlacz_Slepe()    
#Przetnij_Krzyzujace_Sie()
Kategoryzacja_Links()             
V#isum.SaveVersion("D:/A.ver")       
   
    