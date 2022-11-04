import os
import shortuuid
from PyQt5.QtCore import QMarginsF
from PyQt5.QtGui import QPageLayout, QPageSize
from xlsx2html import xlsx2html
from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets
import sys
import tempfile

app = QtWidgets.QApplication(sys.argv)
htmlFile = ''

def main(argv, arc):
    filepath = 'C:/GPI/XL2PDF/'

    source_file = filepath+'file.xlsx'
    target_file = filepath+'file.pdf'
    #source_file = sys.argv[1]
    #target_file = sys.argv[2]

    file_name = os.path.basename(target_file)
    #tempHtml = tempfile.NamedTemporaryFile(dir=os.path.dirname(target_file))
    #tempHtml.close()
    #htmlFile = tempHtml.name
    htmlFile = getUniquePath(os.path.dirname(target_file), '.html')
    xlsx2html(source_file, htmlFile)

    loader = QtWebEngineWidgets.QWebEngineView()
    loader.setZoomFactor(1)
    loader.page().pdfPrintingFinished.connect(lambda *args:  finishProcess(htmlFile))
    #loader.page().
    loader.load(QtCore.QUrl.fromLocalFile(htmlFile))

    layout = QPageLayout()
    layout.setPageSize(QPageSize(QPageSize.A4))
    layout.setOrientation(QPageLayout.Portrait)
    layout.setMargins(QMarginsF(15., 15., 15., 15.))
    loader.loadFinished.connect(lambda *args: loader.page().printToPdf(target_file, pageLayout=layout))

    app.exec()


def finishProcess(htmlFile):
    app.quit()
    os.remove(htmlFile)

def getUniquePath(folder, res):
    path = os.path.join(folder, str(shortuuid.uuid())+res)
    while os.path.exists(path):
         path = os.path.join(folder, str(shortuuid.uuid())+res)
    return path

if __name__ == "__main__":
   main(sys.argv, len(sys.argv))
