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
workbook =  xlrd.open_workbook('/home/atlas-auth/Programs/ATLAS_NSW_AUTH_IMPORT_SCRIPTS/Excel_Measurements_Playground/Drift_Panels/Drift_Panel-Thickness_and_Planarity_Side-2_Vac_OFF/Drift_Panel-Thickness_and_Planarity_Side-2_Vac_OFF.xls')
sheet = workbook.sheet_by_name('Sheet1')
csvfile = open('/home/atlas-auth/Programs/ATLAS_NSW_AUTH_IMPORT_SCRIPTS/CSV_files_Playground/Drift_Panels/Drift_Panel-Thickness_and_Planarity_Side-2_Vac_OFF/Drift_Panel-Thickness_and_Planarity_Side-2_Vac_OFF.csv', 'wb')
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
resstr+=";".join(["",str(EQID),'DT_PAN_PLAN_SIDE_2_MIN',str(sheet.cell(32,2).value),"","T","",fdate,shifter,'ggiakous',meascomment,meascommenttype,eqcomment,eqcommenttype,position]) +"\n"
resstr+=";".join(["",str(EQID),'DT_PAN_PLAN_SIDE_2_MAX',str(sheet.cell(32,4).value),"","T","",fdate,shifter,'ggiakous',meascomment,meascommenttype,eqcomment,eqcommenttype,position]) +"\n"
resstr+=";".join(["",str(EQID),'DT_PAN_PLAN_SIDE_2_AVG',str(sheet.cell(32,6).value),"","T","",fdate,shifter,'ggiakous',meascomment,meascommenttype,eqcomment,eqcommenttype,position]) +"\n"
resstr+=";".join(["",str(EQID),'DT_PAN_PLAN_SIDE_2_RMS',str(sheet.cell(32,8).value),"","T","",fdate,shifter,'ggiakous',meascomment,meascommenttype,eqcomment,eqcommenttype,position])


csvfile.write(resstr)
csvfile.close()

pattern="%s;%s;"
