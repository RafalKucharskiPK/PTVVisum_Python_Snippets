from dbfpy.dbf import Dbf
import os
folder = "E:\\dropbox\\pk\\phd\\dane"
for plik in os.listdir(folder):
    if plik.endswith(".DBF"):
        baza=Dbf(os.path.join(folder,plik))
        i=0
        for row in baza:
            i+=1
            print row
            if i==10:
                wswsw
            
        
    