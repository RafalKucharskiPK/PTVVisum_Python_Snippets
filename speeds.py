import win32com.client

import time
TimeZero=time.time()
Visum=win32com.client.Dispatch("Visum.Visum")
print "Get Client: "+ str(time.time()-TimeZero)
TimeZero=time.time()

Visum.LoadVersion("D:/makenet.ver")
print "Load Version: "+ str(time.time()-TimeZero)
TimeZero=time.time()
results=[[],[],[],[],[]]


for a in range(100):
    TimeZero=time.time()
    Visum.Net.Nodes.ItemByKey(1).AttValue("No")    
    results[0].append(time.time()-TimeZero)
    TimeZero=time.time()
    
    for i in range(1,81):
        Visum.Net.Nodes.ItemByKey(1).AttValue("No")
    results[1].append(time.time()-TimeZero)    
    TimeZero=time.time()
    
    Visum.Net.Nodes.GetMultiAttValues('No')    
    results[2].append(time.time()-TimeZero)   
    TimeZero=time.time()
    
    Visum.Net.Nodes.GetMultipleAttributes(["No","Code","Name","AddVal1","AddVal2","AddVal3"])
    results[3].append(time.time()-TimeZero)
    TimeZero=time.time()    
    
    
    List=Visum.Lists.CreateNodeList;
    List.AddColumn('No');
    List.AddColumn('Code');
    List.AddColumn('Name');
    List.AddColumn('AddVal1');
    List.AddColumn('AddVal2');
    List.AddColumn('AddVal3');
    ListCell=List.SaveToArray;    
    results[4].append(time.time()-TimeZero)
    


print "ItemByKey Node 'No' Single: "+str(sum(results[0])/100)
print "ItemByKey Node 'No' 81 nodes :"+str(sum(results[1])/100)
print "ItemByKey Node GetMultiAttValues: "+str(sum(results[2])/100)
print "GetMultipleAttributes :"+str(sum(results[3])/100)
print "Create list: "+str(sum(results[4])/100)



