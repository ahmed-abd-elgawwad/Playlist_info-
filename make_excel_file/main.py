from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.uic import loadUiType
import sys
from os import path
from pytube import Playlist
from pytube import YouTube
import pandas as pd
FORM_CLASS,_ =loadUiType(path.join(path.dirname(__file__),"main.ui"))
class MainApp(QMainWindow,FORM_CLASS):
    def __init__(self,parent =None):
        super(MainApp,self).__init__(parent)
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.ui_settings()
        self.buttons()
    # handling the ui settings
    def ui_settings(self):
        self.setFixedSize(500,300)
        self.setWindowTitle("make_excel_file")
    # hindeling button and connecting them
    def buttons(self):
        self.pushButton.clicked.connect(self.make_Sheet)
        self.pushButton_3.clicked.connect(self.get_location)
    def get_location(self):
        self.save_location=QFileDialog.getExistingDirectory(self,"select location")
        self.lineEdit_5.setText(self.save_location)
    def make_Sheet(self):
        # try:
            if self.lineEdit.text() =="" or self.lineEdit_2.text == "":
                QMessageBox.warning(self,"warning","please, put the link first and the file name")
            else:
                url= self.lineEdit.text()
                file_name=self.lineEdit_2.text()
                p=Playlist(url)
                titles=[video.title for video in p.videos]
                urls= p.video_urls
                length=[YouTube(x).length for x in urls]
                file_path=f"{self.save_location}/{file_name}.xlsx" 
                time=[ round(int(x)/60) for x in length]
                df=pd.DataFrame({
                    "Video_title":titles,
                     "Time ( min )":time,
                      "Link":urls
                        })
                with pd.ExcelWriter(file_path) as file:
                    df.to_excel(file,sheet_name="new",index=False)
                QMessageBox.information(self,"sucess","the file is made")
            
        # except:
        #      QMessageBox.information(self,"Error","........")
        
# run on the window
def main():
    app =QApplication(sys.argv)
    window = MainApp()
    window.show()
    app.exec_()
if __name__=="__main__":
    main()
