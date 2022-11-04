import os
import shortuuid
from PyQt5.QtCore import QMarginsF
from PyQt5.QtGui import QPageLayout, QPageSize
from xlsx2html import xlsx2html
from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets
import sys
from configparser import ConfigParser
import traceback

config = ConfigParser()

config.read('config.ini')
config.add_section('main')
config.set('main', 'key1', 'value1')
config.set('main', 'key2', 'value2')
config.set('main', 'key3', 'value3')

app = QtWidgets.QApplication(sys.argv)
htmlFile = ''
layout = QPageLayout()

def main(argv, arc):
        try:
            readConfig()
            filepath = 'C:/GPI/XL2PDF/'
            source_file = filepath+'file.xlsx'
            target_file = filepath+'file.pdf'
            #source_file = sys.argv[1]
            #target_file = sys.argv[2]

            htmlFile = getUniquePath(os.path.dirname(target_file), '.html')
            xlsx2html(source_file, htmlFile)

            loader = QtWebEngineWidgets.QWebEngineView()
            loader.setZoomFactor(1)
            loader.page().pdfPrintingFinished.connect(lambda *args:  finishProcess(htmlFile))
            loader.load(QtCore.QUrl.fromLocalFile(htmlFile))
            loader.loadFinished.connect(lambda *args: loader.page().printToPdf(target_file, pageLayout=layout))

            app.exec()

        except Exception:
            log = open("log.txt", "w")
            traceback.print_exc(file=log)

def readConfig():
    config.read('config.ini')

    left = float(config.get('margins', 'left'))
    right = float(config.get('margins', 'top'))
    top = float(config.get('margins', 'right'))
    bottom = float(config.get('margins', 'bottom'))
    layout.setMargins(QMarginsF(left, top, right, bottom))

    orientation = config.get('page', 'orientation')
    if orientation == 'portrait':
        layout.setOrientation(QPageLayout.Portrait)
    elif orientation == 'landscape':
        layout.setOrientation(QPageLayout.Landscape)

    size = config.get('page', 'size')

    if size == 'A4':
        layout.setPageSize(QPageSize(QPageSize.A4))

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
