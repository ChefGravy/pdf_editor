# -*- coding: utf-8 -*-
"""
Created on Mon Nov  8 12:26:33 2021

@author: Pepper Rodney

"""
from pdf2image import convert_from_path

import sys, os, cv2, win32api, shutil

# UI
from PyQt5.QtWidgets import QApplication, QWidget, QFileDialog, QMessageBox, QMenu, QAction, QMainWindow, QLabel, QVBoxLayout, QHBoxLayout, QStatusBar, QPushButton, QScrollArea, QInputDialog
from PyQt5.QtGui import QPixmap, QMouseEvent, QCursor
from PyQt5 import QtCore

class Window(QMainWindow):
    
    def __init__(self, parent=None):
        super().__init__(parent)
        # main UI window
        self.initUI()
        # mouse listener
        #self.setMouseTracking(True)
        # global var for image zoom
        self._zoomInt = 2500
        # image counter
        self.img_counter = 0
        # directory image list
        self.cwd_images = []

    
    # main window
    def initUI(self):
        
        # run menu
        self.createActions()
        self.createMenus()
        
        # set window title and dimensions
        self.setWindowTitle("PDF Editor")
        
        # window location and dimensions
        self.setGeometry(QtCore.Qt.AlignCenter, QtCore.Qt.AlignCenter, 400, 150)
        
        # main widget object
        self.widget = QWidget()
        # main layout object
        self.layout = QVBoxLayout() 

        # scroll widget
        self.scroll = QScrollArea()
        
        self.widget.setGeometry(QtCore.Qt.AlignCenter, QtCore.Qt.AlignCenter, self.width(), self.height())
        
        #Scroll Area Properties
        self.scroll.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.scroll.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(self.widget)

        self.setCentralWidget(self.scroll)
        
        self.widget.setLayout(self.layout)
        
        # add status bar to window
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        
        
    # mouse clicks
    def mousePressEvent(self, QMouseEvent):
        print(QMouseEvent.pos())
        
    def mouseReleaseEvent(self, QMouseEvent):
        cursor = QCursor()
        print(cursor.pos())
        
  
    #create buttons/actions
    def createActions(self):
        self.main = QAction("&Load...", self, shortcut = "Ctrl + L", enabled = True, triggered = self.main)
        self.about = QAction("&About...", self, shortcut = "Ctrl + A", enabled = True, triggered = self.about)
        self.exitAct = QAction("&Exit...", self, shortcut = "Ctrl + Q", enabled = True, triggered = self.cleanUp)
        
    # create menu
    def createMenus(self):
        fileMenu = QMenu("&File", self)
        fileMenu.addAction(self.main)
        fileMenu.addSeparator()
        fileMenu.addAction(self.exitAct)
        
        helpMenu = QMenu("&Help", self)
        helpMenu.addAction(self.about)
        
        self.menuBar().addMenu(fileMenu)
        self.menuBar().addMenu(helpMenu)
        
    def controller(self):
        # button widget
        self.widget = QWidget()
        self.widget.resize(100,100)
        
        # next image
        self.nextImage = QPushButton("Next Page")
        self.nextImage.clicked.connect(lambda: self.nextPreImg(self.nextImage))
        # previous image
        self.previousImage = QPushButton("Previous Page")
        self.previousImage.clicked.connect(lambda: self.nextPreImg(self.previousImage))
        # zoom in
        self.zmIn = QPushButton("Zoom In")
        self.zmIn.clicked.connect(lambda: self.zoomBtn(self.zmIn))
        # zoom out
        self.zmOut = QPushButton("Zoom Out")
        self.zmOut.clicked.connect(lambda: self.zoomBtn(self.zmOut))
        
        layoutH0 = QHBoxLayout()

        layoutV0 = QVBoxLayout()
        layoutV0.addWidget(self.previousImage)
        layoutV0.addWidget(self.nextImage)
        
        layoutV1 = QVBoxLayout()
        layoutV1.addWidget(self.zmIn)
        layoutV1.addWidget(self.zmOut)
        
        layoutH0.addLayout(layoutV1)
        layoutH0.addLayout(layoutV0)
        
        self.widget.setLayout(layoutH0)
        self.widget.show()

    def nextPreImg(self, b):
        if b.text() == "Next Page":
            if self.img_counter == len(self.cwd_images)-1:
                QMessageBox.about(self,"PDF Editor","Last page!")
            else:
                self.img_counter += 1
                self.label.setPixmap(QPixmap(self.cwd_images[self.img_counter]))
        elif b.text() == "Previous Page":
            if self.img_counter == 0:
                QMessageBox.about(self, "PDF Editor","First page!")
            else:
                self.img_counter -= 1
                self.label.setPixmap(QPixmap(self.cwd_images[self.img_counter]))
            
    def zoomBtn(self, b):
        if b.text() == "Zoom In":
            self._zoomInt += 50
            self.label.setPixmap(self.pixmap.scaled(self._zoomInt, self._zoomInt, QtCore.Qt.KeepAspectRatio))
        elif b.text() == "Zoom Out":
            self._zoomInt -= 50
            self.label.setPixmap(self.pixmap.scaled(self._zoomInt, self._zoomInt, QtCore.Qt.KeepAspectRatio))

    def about(self):
        QMessageBox.about(self, "About Adobe Editor","<p>The <b> Adobe Editor </b> was a project I came up with to "
                         " help me learn more about PyQt. This is an open source project so if you see something that "
                         " needs improvement, https://github.com/ChefGravy is where I keep it. Enjoy!")
        
    # convert pdf pages to images and saves to folder
    # returns directory
    def convertPdf(self): 
        img_counter = 0
        try:
            # open and convert pdf to image
            options = QFileDialog.Options()
            fileName, _ = QFileDialog.getOpenFileName(self, 'QFileDialog.getOpenFileName()', '', 
                                                      'PDF files(*.pdf)', options = options)
            
            images = convert_from_path(fileName, poppler_path = 
                    r"C:\Users\PepperRodney\Documents\Python\Poppler-21.11.0-0\poppler-21.11.0\Library\bin")
            
            # display statusbar and buttons if pdf loaded/converted successfully
            if len(images) > 0:
                self.statusBar.showMessage(".pdf load success!")
                
            # create folder to store images on desktop
            cwd = os.environ['USERPROFILE'] + "\\Desktop\\pdf_editor"
            if not os.path.exists(cwd):
                os.makedirs(cwd)
            os.chdir(cwd)
            
            # store images in newly created directory
            for i in range(len(images)):
                # save pages as images in the pdf
                images[i].save('page' + str(i) + '.jpg', 'JPEG')
                
            # capture image location and store in list
            for root, dirs, files in os.walk(cwd):
                for filename in files:
                    self.cwd_images.append(root + "\\" + filename)
                    img_counter += 1

            return cwd
        
        except Exception as e:
            msgBox = QMessageBox()
            msgBox.warning(self, "Error", "Something went wrong with converting the pdf: {}".format(e))
    
    # Check on the size of the monitor relative to the pdf image.
    # Propose a different size
    def size(self):
        try:
            # first image size
            image = cv2.imread(self.cwd_images[0]).shape
                    
            # monitor dimensions
            monitors = win32api.EnumDisplayMonitors()
            
            # consider multiple monitors for future improvements
            _monitors = win32api.GetMonitorInfo(monitors[0][0])
            monitor = _monitors.get("Monitor")
            
            if image[0] > monitor[2] or image[1] > monitor[3]:
                num, ok = QInputDialog.getInt(self, "Image Size", "The image you're loading is {} by {} pixels "
                                             "while your monitor is {} by {} pixels."
                                             "\nConsider a size (see default value) for your image"
                                             " that you can work with. Aspect ratio is held constant."
                                             .format(image[0], image[1], monitor[2], monitor[3]), monitor[2]-200)
            
            # if user cancels, set size to default
            if ok == False:
                self._zoomInt = image[0]
            else:
                self._zoomInt = num
        
        except Exception as e:
            msgBox = QMessageBox()
            msgBox.warning(self, "Error", "Something went wrong with size calc: {}".format(e))
    
    def loadPdf(self):
        try:
            # enable mouse click events
            self.setEnabled(True)
            
            # create label and pixmap to display image on screen
            self.label = QLabel(self)
            print(self.img_counter)
            self.pixmap = QPixmap(self.cwd_images[self.img_counter])
            self.label.setAlignment(QtCore.Qt.AlignCenter)
        
            # TO DO: allow user to tune size via buttons
            pix = self.pixmap.scaled(self._zoomInt, self._zoomInt, QtCore.Qt.KeepAspectRatio)
            self.label.setPixmap(pix)
            self.layout.addWidget(self.label)
        except Exception as e:
            msgBox = QMessageBox()
            msgBox.warning(self, "Error", "Something went wrong with loading the pdf: {}".format(e)) 

    # close all windows
    def cleanUp(self):
        self.close()
        self.widget.close()
        
    # main run method    
    def main(self):
        # convert pdf to images
        cwd = self.convertPdf()
        if cwd:
            # check pixel size of image
            self.size()
            # load pdf image to screen
            self.loadPdf()
            # load controller
            self.controller()
        else:
            return 0
    
def main():
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
    
    #delete folder with pdf images
    if os.path.isdir(os.environ['USERPROFILE'] + "\\Desktop\\pdf_editor"):
        shutil.rmtree(os.environ['USERPROFILE'] + "\\Desktop\\pdf_editor")
            
if __name__ == "__main__":
    main()
    

