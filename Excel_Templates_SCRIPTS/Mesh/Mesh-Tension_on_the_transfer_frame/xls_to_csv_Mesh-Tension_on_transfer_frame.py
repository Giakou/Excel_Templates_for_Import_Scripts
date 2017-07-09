import xlrd
from datetime import datetime

workbook =  xlrd.open_workbook('Mesh-Tension_on_transfer_frame_5385.xls')
sheet = workbook.sheet_by_name('Sheet1')
csvfile = open('Mesh-Tension_on_transfer_frame_5385.csv', 'wb')
shifter=sheet.cell(11,0).value
EQID=int(sheet.cell(5,0).value)
date=sheet.cell(8,0).value
datetuple=xlrd.xldate_as_tuple(date,workbook.datemode)
fdate=datetime(*datetuple) #convert date in datetime type
fdate=fdate.strftime("%Y-%m-%d %H:%M:%S") #date to str
print "shifter, EQID,date",shifter,EQID,type(shifter),fdate
resstr=""
resstr+=";".join(['MEASSITEHASH','EQENTRYID','CONTEXTNAME','MEASVALUE','PERCENTAGEMEAS','ISVALIDFLAG','INDEXHASH','MEASTIME','SHIFTER','WEBSITEUSERCR']) +"\n"
resstr+=";".join(["",str(EQID),'MESH_TENS_TF_MIN_DT',str(sheet.cell(22,7).value),"","T",'skfh_1',fdate,shifter,'ggiakous']) +"\n"
resstr+=";".join(["",str(EQID),'MESH_TENS_TF_MAX_DT',str(sheet.cell(22,9).value),"","T",'skfh_1',fdate,shifter,'ggiakous']) +"\n"
resstr+=";".join(["",str(EQID),'MESH_TENS_TF_AVG_DT',str(sheet.cell(22,11).value),"","T",'skfh_1',fdate,shifter,'ggiakous']) +"\n"
resstr+=";".join(["",str(EQID),'MESH_TENS_TF_RMS_DT',str(sheet.cell(22,13).value),"","T",'skfh_1',fdate,shifter,'ggiakous'])


csvfile.write(resstr)
csvfile.close()

pattern="%s;%s;"
