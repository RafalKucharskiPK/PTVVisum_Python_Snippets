import sqlite3

def VisumInit(path=None,COMAddress='Visum.Visum.125'):
    """
    ###
    Automatic Plate Number Recognition Support
    (c) 2012 Rafal Kucharski info@intelligent-infrastructure.eu
    ####
    VISUM INIT
    """

    import win32com.client        
    Visum = win32com.client.Dispatch(COMAddress)
    if path != None: Visum.LoadVersion(path)
    return Visum

def Init_DB():
    def DB_Architecture(con,cur):
        cur.execute("""create table Stops(Id_S INTEGER PRIMARY KEY)""")
        cur.execute("""create table Lines(Id_L INTEGER PRIMARY KEY AUTOINCREMENT)""")
        cur.execute("""create table Routes(Id_R INTEGER PRIMARY KEY AUTOINCREMENT,
                                            Id_L INTEGER,
                                            Ind INTEGER,
                                            ID_S INTEGER,
                                            FOREIGN KEY (Id_L) REFERENCES Lines(Id_L),
                                            FOREIGN KEY (Id_S) REFERENCES Stops(Id_S)
                                            )""")
        cur.execute("""create table Times(From_Stop Integer, 
                                          To_Stop, Integer,
                                          FOREIGN KEY (From_Stop) REFERENCES Stops(Id_S)
                                          FOREIGN KEY (To_Stop) REFERENCES Stops(Id_S)
                                          )""")
        return con,cur

    con = sqlite3.connect(":memory:")
    cur = con.cursor()
    cur.execute('pragma foreign_keys=ON')  
    con.text_factory = str
    con.commit()
    con,cur=DB_Architecture(con,cur)
    return con,cur        
          
def DB_Input(con,cur):
    StopPoints = Visum.Net.StopPoints.GetMultipleAttributes(['NO'])
    StopPoints=[[int(s[0])] for s in StopPoints]
    cur.executemany('insert into Stops(ID_S) values (?)', StopPoints)
    con.commit()
    LR2DB=[]
    LineRoutes=Visum.Net.LineRoutes.GetMultipleAttributes(["Concatenate:Stoppoints\No","Concatenate:Timeprofiles\Concatenate:Timeprofileitems\Preruntime"])
    Lines=[]
    for i,LR in enumerate(LineRoutes):
        Lines.append([i])
        Stops=LR[0].split(":")
        Stops=[int(s) for s in Stops]
        for j,Stop in enumerate(Stops):
            LR2DB.append([i,j,Stop])
    print Lines
    cur.executemany('insert into Lines(ID_L) values (?)', tuple(Lines))
    cur.executemany('insert into Routes(ID_L,Ind,ID_S) values (?,?,?)', LR2DB)
    con.commit()    
        


Visum=VisumInit("E:\KA.ver")
#
#StopsPoints=[int(s[1]) for s in Visum.Net.Stops.GetMultiAttValues("No")]
#LineRoutes=Visum.Net.LineRoutes.GetMultipleAttributes(["Concatenate:Stoppoints\No","Concatenate:Timeprofiles\Concatenate:Timeprofileitems\Preruntime"])
con,cur=Init_DB()
DB_Input(con,cur)









