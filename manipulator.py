import random

def Init(path=None):
        import win32com.client        
        Visum=win32com.client.Dispatch('Visum.Visum.125')
        if path!=None: Visum.LoadVersion(path)
        return Visum

    
def set_Copy():
    Visum.Net.Nodes.AddUserDefinedAttribute('Stary_Node_No_Do_usuniecia', 'Stary_Node_No_Do_usuniecia', 'Stary_Node_No_Do_usuniecia', 1,0) #else create UDAs
        
    Visum.Net.Nodes.SetMultiAttValues("Stary_Node_No_Do_usuniecia", Visum.Net.Nodes.GetMultiAttValues("No"))
    
 
def get_Nodes(Attr):
    return Visum.Net.Nodes.GetMultiAttValues(Attr)

    #return [node[1] for node in Visum.Net.Nodes.GetMultiAttValues(Attr,True)]

def get_Links(Attr):
    return [node for node in Visum.Net.Links.GetMultipleAttributes(Attr)]

def set_Nodes(old_Nodes):    
    z=float(len(old_Nodes))
    for i in range(len(old_Nodes)):
        print "N1", i, i/z
        Visum.Net.Nodes.ItemByKey(old_Nodes[i]).SetAttValue("No",1000000000+old_Nodes[i])
    
    for i in range(len(old_Nodes)):
        print "N2", i, i/z        
        Visum.Net.Nodes.ItemByKey(old_Nodes[i]+1000000000).SetAttValue("No",int(i))


def set_Links(old_Nodes):    
    z=float(len(old_Nodes))
    for i in range(len(old_Nodes)):
        print "L1", i, i/z
        Visum.Net.Links.ItemByKey(old_Nodes[i][1],old_Nodes[i][2]).SetAttValue("No",1000000000+old_Nodes[i][0])
        
def get_StopPoints():
    return [node[1] for node in Visum.Net.StopPoints.GetMultiAttValues("No")]

def set_StopPoints(SP):
    z=float(len(SP))
    for i in range(len(SP)):
        print "SP1", i, i/z        
        Visum.Net.StopPoints.ItemByKey(SP[i]).SetAttValue("No",100000+i)
        
    for i in range(len(SP)):
        print "SP2", i, i/z
        SP=Visum.Net.StopPoints.ItemByKey(100000+i)       
        try:                
            SP.SetAttValue("No",SP.AttValue("NodeNo"))
        except: pass
        
def get_StopAreas():
    return [node[1] for node in Visum.Net.StopAreas.GetMultiAttValues("No")]

def set_StopAreas(SP):
    z=float(len(SP))
    for i in range(len(SP)):
        print "SA1", i, i/z
        Visum.Net.StopAreas.ItemByKey(SP[i]).SetAttValue("No",100000+i)
        
    for i in range(len(SP)):
        print "SA2", i, i/z
        SP=Visum.Net.StopAreas.ItemByKey(100000+i)                
        try: SP.SetAttValue("No",SP.AttValue("NodeNo"))
        except:
            print SP.AttValue("No"), "err" 
        
def get_Zones():
    return [node[1] for node in Visum.Net.Zones.GetMultiAttValues("No")]

def set_Zones(old_Nodes,new_Nodes):
    
    z=float(len(old_Nodes))
    for i in range(len(old_Nodes)):
        print "Z1", i, i/z
        Visum.Net.Zones.ItemByKey(old_Nodes[i]).SetAttValue("No",10000+old_Nodes[i])
        
    for i in range(len(old_Nodes)):
        print "Z2", i, i/z        
        Visum.Net.Zones.ItemByKey(old_Nodes[i]+10000).SetAttValue("No",new_Nodes[i])
        
def get_Stops():
    return [node[1] for node in Visum.Net.Stops.GetMultiAttValues("No")]

def set_Stops(old_Nodes,new_Nodes):
    
    z=float(len(old_Nodes))
    for i in range(len(old_Nodes)):
        print "S1", i, i/z
        Visum.Net.Stops.ItemByKey(old_Nodes[i]).SetAttValue("No",10000+old_Nodes[i])
        
    for i in range(len(old_Nodes)):
        print "S2", i, i/z        
        Visum.Net.Stops.ItemByKey(old_Nodes[i]+10000).SetAttValue("No",new_Nodes[i])
    
try: 
    Visum
except: 
    Visum=Init("D://Dropbox//i2//Prace//___Nie Visumowe//2012, Malopolska//Visum//SHP_KRK_TEST//Wycieta_Czesc.ver")
    
#SA=get_StopAreas()

#i=0
#for Area in SA:
#    
#    try:
#        SArea=Visum.Net.StopAreas.ItemByKey(Area)
#        X=Visum.Net.Nodes.ItemByKey(SArea.AttValue("NodeNo")).AttValue("XCoord")
#        Y=Visum.Net.Nodes.ItemByKey(SArea.AttValue("NodeNo")).AttValue("YCoord")
#        SArea.SetAttValue("XCoord",X)
#        SArea.SetAttValue("YCoord",Y)
#        print i/float(len(SA))
#        i+=1 
#        Stop=Visum.Net.Stops.ItemByKey(SArea.AttValue("StopNO"))
#        Stop.SetAttValue("XCoord",X)
#        Stop.SetAttValue("YCoord",Y)
#    except: pass

#
#
Nodes=get_Nodes("No")
print Nodes
Nodes=list(Nodes)
Nodes=[list(node) for node in Nodes]
for i,node in enumerate(Nodes):
    node[1]=i+100000000
print Nodes

Visum.Net.Nodes.SetMultiAttValues("No",Nodes)

Nodes=get_Nodes("No")
print Nodes
Nodes=list(Nodes)
Nodes=[list(node) for node in Nodes]
for i,node in enumerate(Nodes):
    node[1]=i+1000000
print Nodes
Visum.Net.Nodes.SetMultiAttValues("No",Nodes)





#random.shuffle(Nodes)
#set_Nodes(Nodes)
#Links=get_Links(["No","FromNodeNo","ToNodeNo"])

#set_Links(Links)
Visum.SaveVersion("D:\\po.ver")
#
#
#SP=get_StopPoints()
#set_StopPoints(SP)
#
#Visum.SaveVersion("D:\\GDYNIA_DANE\\rano\\SP.ver")
#
#SA=get_StopAreas()
#set_StopAreas(SA)
#
#Visum.SaveVersion("D:\\GDYNIA_DANE\\rano\\SA.ver")
#Zones=get_Zones()
#print get_Zones()
#
#
#random.shuffle(Zones)
#print Zones
#set_Zones(get_Zones(),Zones)
#
#Visum.SaveVersion("D:\\GDYNIA_DANE\\rano\\Zones.ver")
#
#Stops=get_Stops()
#
#random.shuffle(Stops)
#
#set_Stops(get_Stops(),Stops)
#
#
#
#Visum.SaveVersion("D:\\GDYNIA_DANE\\rano\\All.ver")







    
    

    
