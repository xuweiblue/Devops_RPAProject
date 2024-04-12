from win32com import client
import win32com.client as win32
import os
import time
import shutil
from datetime import datetime



#step 1: backup excel file, save file to backup folder
a= os.scandir('C:\\Users\\xuwei\\OneDrive\\Desktop\\AMKH\\testRPA')
b= next(a)
c= next(a)
d= next(a)
print (d)
fileName= str(d) 

strfileName=fileName[fileName.index('M'): fileName.index('>')-1]

strfileNamePDF= strfileName[0: strfileName.index('.')]+'.pdf'

print (strfileName)
file_to_copy='C:\\Users\\xuwei\\OneDrive\\Desktop\\AMKH\\testRPA\\'+strfileName
dest_dir= 'C:\\Users\\xuwei\\OneDrive\\Desktop\\AMKH\\testRPA\\Backup'
arc_dir= 'C:\\Users\\xuwei\\OneDrive\\Desktop\\AMKH\\testRPA\\Archive\\'
fileNamePDFDir=arc_dir+strfileNamePDF

shutil.copy(file_to_copy,dest_dir)
print ('step 1 done')


#step5 to convert excel to pdf and save to archive folder 

def exl2pdf(file_location):
     app= client.Dispatch("Excel.Application")
     app.Interactive = False
     app.Visible = False
     workbook = app.Workbooks.open(file_location)
     output= os.path.splitext(arc_dir+strfileName)[0]
     time.sleep(1)
     workbook.ActiveSheet.ExportAsFixedFormat(0,output)
     time.sleep(1)
     workbook.Close()


     
exl2pdf(file_to_copy)
print ('step 5 done')

#step 7 create pivot table

win32c= win32.constants
#launch excel application
xlApp = win32.Dispatch("Excel.Application")
xlApp.Visible = True #False

#create workbook : wb and ws
#reference workbooks
wb = xlApp.Workbooks.Open(file_to_copy)

#reference worksheets
ws_data = wb.Worksheets("Sheet1")
mywork= wb.sheets.Add()
wb.Worksheets(1).Name ="Pivot table"
ws_report = wb.Worksheets("Pivot table")



#create pivot table cache connection 
pt_cache = wb.PivotCaches().Create(1,ws_data.Range("A1").CurrentRegion)

#create pivot table designer/editor
pt = pt_cache.CreatePivotTable(ws_report.Range("A3"), "myreport_summary")

#row and column grandtotals
pt.ColumnGrand = True
pt.RowGrand = True

#change report layout
pt.RowAxisLayout(1)   #RowAxisLayout(1) for tabular form

#change pivot table style
#Select from Design tab, try out Medium9 or Medium3
pt.TableStyle2 = "pivotStyleMedium21" 



def create_pivot_table(pt):
    #add row
    field_rows = {} 
    field_rows["Movement Type"] = pt.PivotFields("Movement Type")
    field_rows["Movement Type"].Orientation = 1   # 4 for data/value

    #field_rows["Movement Type"].Function = -4112

    #add value
    field_values = {} 
    field_values["Sum of Amount in LC"] = pt.PivotFields("Amount in LC")
    #field_values["Movement Type"] = pt.PivotFields("Movement Type")


    field_values["Sum of Amount in LC"].Orientation =4  # 4 for data/value
    
    field_values["Sum of Amount in LC"].Function = -4157  # -4112 for xlCount


    #field_values["Sum of Amount in LC"].NumberFormat = "#,##0" # "#,##0" for number format
    


#create report
create_pivot_table(pt)

print ('step 7 done')

# step 8 to set formula 
used= ws_data.UsedRange
nrows= used.Row + used.Rows.Count -1

#print (nrows)
formu1 = "=SUMIFS(G:G,C:C,C"
formu2 = ")"


for i in range(2, nrows+1):

    ws_data.Cells(i, "O").Value = formu1+str(i)+formu2
    
print ('step 8 done')

# step 9 save the updated excel file to Archive folder

now= datetime.now()
#date_time=now.strftime("%Y%m%d%H:%M:%S")
date_time=now.strftime("%Y%m%d_%H%M%S")

strfileName1= date_time+'_'+strfileName
strfileName1Dir=arc_dir+strfileName1


wb.SaveAs(strfileName1Dir)

wb.Close()

xlApp.Quit()
print ('step9 done')




#step 11 delete the file:
def DeleteFile(filepath):

     if os.path.exists(filepath):
          os.remove(filepath)
          print("Deleted "+filepath)

     else:
          print ("File doesn't exist")

DeleteFile(file_to_copy)
print ('step 11 done')





     


