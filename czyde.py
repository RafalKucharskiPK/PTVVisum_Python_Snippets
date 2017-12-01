import os
import sys, PyQt4.QtCore
x=sys.argv
x = os.path.dirname(__file__)
f = open('D:/a.txt', 'w')
f.write(str(x))
f.close()