import wx

def VisumInit(path=None):
    """       
    VISUM INIT
    """

  
    import win32com.client        
    Visum = win32com.client.Dispatch('Visum.Visum.125')
    if path != None: Visum.LoadVersion(path)
    return Visum


def dist(a,b):
    return ((a[0]-b[0])**2+(a[1]-b[1])**2)**0.5
    
def kat(x,y):
    return (b[1]-a[1])/(b[0]-a[0])

def nowa_centroida(KRK,stara,relative):
    noweX=KRK[0]+(stara[0]-KRK[0])*float(relative)
    noweY=KRK[1]+(stara[1]-KRK[1])*float(relative)
    return [noweX,noweY]


row_id={"Krakow": 171, "Nowy Sacz": 172,"Tarnow": 173}
Nodes={"Krakow": 190000+171, "Nowy Sacz": 190000+172,"Tarnow": 190000+173}
KRK_row_id = 171
script=True

try:
    Visum
    
except:
    Visum=VisumInit("C:/domap_pusty.ver")
    script=False

PJTSkim= Visum.Net.Matrices.ItemByKey(10000011).GetValuesDouble()
Krk_row=[]

j=1
for row in PJTSkim:
    
    if row[KRK_row_id]>2000:
        break
    Krk_row.append(row[KRK_row_id])
    j+=1
    
maxdist = float(max(Krk_row))
#maxdist=sum(Krk_row)/float(j)
def get_min(a):
    if a[0]<=a[1]:
        if a[0]<=a[2]:
            return "Krakow"
        else:
            return "Nowy Sacz"
    else:
        if a[1]<=a[2]:
            return "Tarnow"
        else:
            return "Nowy Sacz"
        
        
#for i,dist in enumerate(Krk_row):
#    Krk_row[i]=dist/float(maxdist)

XY={}
Krk=Visum.Net.Zones.ItemByKey(227)
XY["Krakow"]=[Krk.AttValue("XCoord"),Krk.AttValue("YCoord")]
Node=Visum.Net.AddNode(Nodes["Krakow"],XY["Krakow"][0],XY["Krakow"][1])
Node.SetAttValue("TypeNo",7)
Tarnow=Visum.Net.Zones.ItemByKey(228)
XY["Tarnow"]=[Tarnow.AttValue("XCoord"),Tarnow.AttValue("YCoord")]
Node=Visum.Net.AddNode(Nodes["Tarnow"],XY["Tarnow"][0],XY["Tarnow"][1])
Node.SetAttValue("TypeNo",7)
NS=Visum.Net.Zones.ItemByKey(229)
XY["Nowy Sacz"]=[NS.AttValue("XCoord"),NS.AttValue("YCoord")]
Node=Visum.Net.AddNode(Nodes["Nowy Sacz"],XY["Nowy Sacz"][0],XY["Nowy Sacz"][1])
Node.SetAttValue("TypeNo",7)
try:
    Visum.Net.Zones.AddUserDefinedAttribute("MapaOdleglosci_Odleglosc", "MapaOdleglosci_Odleglosc", "MapaOdleglosci_Odleglosc", 2)    
    Visum.Net.Zones.AddUserDefinedAttribute("MapaOdleglosci_Procentowo", "MapaOdleglosci_Procentowo", "MapaOdleglosci_Procentowo", 2)
    Visum.Net.Zones.AddUserDefinedAttribute("MapaOdleglosci_NajblizszeMiasto", "MapaOdleglosci_NajblizszeMiasto", "MapaOdleglosci_NajblizszeMiasto", 5)

except:
    pass
Node = 0    
    
Zones = Visum.Net.Zones.GetMultipleAttributes(["No","XCoord","YCoord"])
i=0
if script:
    dialog = wx.ProgressDialog ('Progress', "Visum skim matrix Calculations", maximum=j)
for Z in Zones:
    if i==j-1:
        break
    Zone = list(Z)
    #najblizszy=get_min([PJTSkim[i][row_id["Krakow"]],PJTSkim[i][row_id["Nowy Sacz"]],PJTSkim[i][row_id["Tarnow"]]])
    najblizszy="Krakow"
    distance=PJTSkim[i][row_id[najblizszy]]
    print najblizszy
    Zone.append(distance) #SKIM VALUE
    Zone.append(Zone[-1]/maxdist) #SKIM VALUE RELATIVE
    #Zone.append(dist(XY[najblizszy],[Zone[1],Zone[2]])) #DISTANCE
    print "1"
    [noweX,noweY]=nowa_centroida(XY[najblizszy],[Zone[1],Zone[2]],Zone[4])
    Visum.Net.AddNode(200001+i,Zone[1],Zone[2])
    print "2"
    Visum.Net.Nodes.ItemByKey(200001+i).SetAttValue("TypeNo",8)
    ItemZone = Visum.Net.Zones.ItemByKey(Zone[0])    
    ItemZone.SetAttValue("XCoord",noweX)
    ItemZone.SetAttValue("YCoord",noweY)
    print "3"
    ItemZone.SetAttValue("MapaOdleglosci_Procentowo",Zone[-1])
    ItemZone.SetAttValue("MapaOdleglosci_Odleglosc",dist(XY[najblizszy],[Zone[1],Zone[2]]))
    ItemZone.SetAttValue("MapaOdleglosci_NajblizszeMiasto",najblizszy)
    
    Visum.Net.AddNode(300001+i,noweX,noweY)
    Visum.Net.Nodes.ItemByKey(300001+i).SetAttValue("TypeNo",9)
    Visum.Net.AddLink(6000000+i,Nodes[najblizszy],300001+i)
    Visum.Net.Links.ItemByKey(Nodes[najblizszy],300001+i).SetAttValue("TypeNo",6)
    print "4"
    Visum.Net.AddLink(5000000+i,Nodes[najblizszy],200001+i)
    Visum.Net.Links.ItemByKey(Nodes[najblizszy],200001+i).SetAttValue("TypeNo",7)
    i+=1
    print "5"
    print Z[0]
    if script:
        dialog.Update(i)
if script:
    dialog.Destroy()

#Visum.SaveVersion("E:/mapaodleglosci.ver")
    










