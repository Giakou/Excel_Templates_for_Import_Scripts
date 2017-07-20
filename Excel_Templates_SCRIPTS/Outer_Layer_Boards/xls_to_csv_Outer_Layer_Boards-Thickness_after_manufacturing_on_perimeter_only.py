# -*- coding: utf-8 -*-
import xlrd
from datetime import datetime
#1) Excel_Measurements_Playground -> Save the excel file in the Playground dir
#2) Excel_Measurements_Construction -> Save the excel file in the Construction dir
#3) The directories are identical
#4) CSV_files_Playground -> Save the data in csv form in the Playground dir
#5) CSV_files_Construction -> Save the data in csv form in the Construction dir
#6) The directories are identical
#7) Don't forget to change the name of the excel and csv files by adding "_eqentryid" before the .xls or .csv
workbook =  xlrd.open_workbook('/home/atlas-auth/Programs/ATLAS_NSW_AUTH_IMPORT_SCRIPTS/Excel_Measurements_Playground/Outer_Layer_Boards/Outer_Layer_Boards-Thickness_after_manufacturing_on_perimeter_only.xls')
sheet = workbook.sheet_by_name('Sheet1')
csvfile = open('/home/atlas-auth/Programs/ATLAS_NSW_AUTH_IMPORT_SCRIPTS/CSV_files_Playground/Outer_Layer_Boards/Outer_Layer_Boards-Thickness_after_manufacturing_on_perimeter_only.csv', 'wb')
meascomment=sheet.cell(14,0).value
meascommenttype=sheet.cell(17,0).value
eqcomment=sheet.cell(20,0).value
eqcommenttype=sheet.cell(23,0).value
position=sheet.cell(26,0).value
shifter=sheet.cell(11,0).value
EQID=int(sheet.cell(5,0).value)
date=sheet.cell(8,0).value
datetuple=xlrd.xldate_as_tuple(date,workbook.datemode)
fdate=datetime(*datetuple) #convert date in datetime type
fdate=fdate.strftime("%Y-%m-%d %H:%M:%S") #date to str
print "shifter, EQID,date",shifter,EQID,type(shifter),fdate
resstr=""
resstr+=";".join(['MEASSITEHASH','EQENTRYID','CONTEXTNAME','MEASVALUE','PERCENTAGEMEAS','ISVALIDFLAG','INDEXHASH','MEASTIME','SHIFTER','WEBSITEUSERCR','MEASCOMMENT','MEASCOMMENTTYPE','EQCOMMENT','EQCOMMENTTYPE','POSITION']) +"\n"
resstr+=";".join(["",str(EQID),'OLB_THICK_MIN_DT',str(sheet.cell(13,6).value),"","T",'skfh_1',fdate,shifter,'ggiakous',meascomment,meascommenttype,eqcomment,eqcommenttype,position]) +"\n"
resstr+=";".join(["",str(EQID),'OLB_THICK_MAX_DT',str(sheet.cell(13,8).value),"","T",'skfh_1',fdate,shifter,'ggiakous',meascomment,meascommenttype,eqcomment,eqcommenttype,position]) +"\n"
resstr+=";".join(["",str(EQID),'OLB_THICK_AVG_DT',str(sheet.cell(13,10).value),"","T",'skfh_1',fdate,shifter,'ggiakous',meascomment,meascommenttype,eqcomment,eqcommenttype,position]) +"\n"
resstr+=";".join(["",str(EQID),'OLB_THICK_RMS_DT',str(sheet.cell(13,12).value),"","T",'skfh_1',fdate,shifter,'ggiakous',meascomment,meascommenttype,eqcomment,eqcommenttype,position])


csvfile.write(resstr)
csvfile.close()

pattern="%s;%s;"
