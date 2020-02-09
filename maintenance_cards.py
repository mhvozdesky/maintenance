import win32com.client
import os.path
from tkinter.filedialog import askopenfilename

our_file = askopenfilename()
our_file = our_file.replace(r'/', '\\') #replace "/" with "\"

Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(our_file)
#wb.Worksheets.Count

z = our_file.split('\\')
report = open('\\'.join(z[:len(z)-1]) + r'\report.txt', 'w')


for same_sheet in range(wb.Worksheets.Count):
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
        report.write('Лист ' + str(same_sheet) + '---' + sheet.name + '\n')
        
report.close()        



######################################################################################################################################################
#sheet = Excel.Sheets(37)

#wb.Worksheets.Count
#sheet.name

#print sheet.Cells(2,2).value sheet.ExportAsFixedFormat(Type = 0, Filename = r'C:\Users\Dom\Desktop\Microsoft Excel.pdf', IgnorePrintAreas = False)
#sheet.ExportAsFixedFormat(Type = 0, Filename = r'C:\Users\Dom\Desktop\Лист Microsoft Excel.pdf', Quality = 0, IncludeDocProperties = True, IgnorePrintAreas = False, OpenAfterPublish = False, From = 3)
#sheet.Name
#print_area = 'P1:AA70'
#sheet.PageSetup.PrintArea = print_area

#wb.Save()
#wb.Close()
#Excel.Quit()
#Excel.Application.Quit()
#import win32api
#e_msg = win32api.FormatMessage(-2147352567)