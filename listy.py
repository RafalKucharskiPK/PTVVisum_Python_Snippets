import numpy as np


m = np.array([[3, 7, 20], [3, 6, 20], [7, 3, 20]])
o = dict()
first_col = [r[0] for r in m]
for r in m:
    if r[0] in o.keys():
       o[r[0]][-1].append([r[1],r[2]])
    else:
        e=[]
        e.append([r[1],r[2]])
        o[r[0]] = [r[0], first_col.count(r[0]), e] 

print o.keys()
print o.values()

    

