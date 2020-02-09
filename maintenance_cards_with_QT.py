from PyQt5.QtWidgets import QFileDialog
from PyQt5 import QtWidgets, uic
import win32com.client
import os.path
import sys
from mainGI import Ui_MainWindow  # импорт нашего сгенерированного файла

our_file = None
report = None




#app = QtWidgets.QApplication([])
#win = uic.loadUi(r"main.ui") # расположение вашего файла .ui
 
#win.show()
#sys.exit(app.exec())

class mywindow(QtWidgets.QMainWindow):
    def __init__(self):
        super(mywindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        
        self.ui.pushButton.clicked.connect(self.get_file)
        self.ui.pushButton_2.clicked.connect(self.run)
     
        
    def get_file(self):
        global our_file
        our_file = QFileDialog.getOpenFileName(self, 'Open file')[0]
        our_file = our_file.replace(r'/', '\\') #replace "/" with "\"
        self.ui.lineEdit.setText(our_file)
        
    def run(self):
        run_excell()
        
def run_excell():
    global report
    
    Excel = win32com.client.Dispatch("Excel.Application")
    wb = Excel.Workbooks.Open(our_file) 
    
    z = our_file.split('\\')
    try:
        report = open('\\'.join(z[:len(z)-1]) + r'\report.txt', 'w')
    except:
        pass
    
    for same_sheet in range(wb.Worksheets.Count):
        application.ui.progressBar.setMaximum(wb.Worksheets.Count)
        
        same_sheet += 1
        sheet = Excel.Sheets(same_sheet)
        
        check_row = 45
        continue_checking = True
        key_phrase = False
        
        #while continue_checking:
            #continue_checking = False
            #for z in range(16, 28): #col P to AA (16 - 27)
                #if sheet.Cells(check_row, z).value == 'Администрация оставляет за собой право изменения цен': 
                    #key_phrase = True
                #if sheet.Cells(check_row, z).value != None and key_phrase == False:
                    #continue_checking = True
                    #check_row += 1
                    
        while continue_checking:
            for z in range(16, 28): #col P to AA (16 - 27)
                if sheet.Cells(check_row, z).value == 'Администрация оставляет за собой право изменения цен': 
                    continue_checking = False
                    check_row += 1
                    break
            check_row += 1 
            if check_row > 75: continue_checking = False
                    
        print_area = 'P1:AA' + str(check_row) #'P1:AA70'
        sheet.PageSetup.PrintArea = print_area
        
        path_file = os.path.dirname(our_file)
        #new_file_name = path_file + '\\' + sheet.name.replace(sheet.Name, '_') + '.pdf'
        new_file_name = new_file_name = path_file + '\\' + sheet.name.replace(' ', '_') + '.pdf'
        try:
            print('Лист ' + str(same_sheet) + '---' + sheet.name + '---' + print_area)
            sheet.ExportAsFixedFormat(Type = 0, 
                                      Filename = new_file_name, #r'C:\Users\Dom\Desktop\Лист Microsoft Excel.pdf' 
                                      Quality = 0, 
                                      IncludeDocProperties = True, 
                                      IgnorePrintAreas = False, 
                                      OpenAfterPublish = False)
        except:
            if report != None: report.write('Лист ' + str(same_sheet) + '---' + sheet.name + '\n')
            
        application.ui.progressBar.setValue(same_sheet)
        
    if report != None: report.close() 
    application.ui.progressBar.setValue(0)

app = QtWidgets.QApplication([])
application = mywindow()
application.show()

sys.exit(app.exec())


