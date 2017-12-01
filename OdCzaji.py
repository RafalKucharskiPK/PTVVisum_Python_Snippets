import sys
import signal
import os
from PyQt4.QtCore import *
from PyQt4.QtGui import *
from PyQt4.QtWebKit import QWebPage

app = QApplication(sys.argv)
signal.signal(signal.SIGINT, signal.SIG_DFL)
webpage = QWebPage()
def onLoadFinished(result):
    if not result:
        print "Request failed"
        sys.exit(1)

    webpage.setViewportSize(webpage.mainFrame().contentsSize())
    image = QImage(webpage.viewportSize(), QImage.Format_ARGB32)
    painter = QPainter(image)
    webpage.mainFrame().render(painter)
    painter.end()
    if os.path.exists("output.png"):
        os.remove("output.png")
    image.save("output.png")
    sys.exit(0) # quit this application
webpage.mainFrame().load(QUrl("http://google.pl"))
webpage.connect(webpage, SIGNAL("loadFinished(bool)"), onLoadFinished)

sys.exit(app.exec_())