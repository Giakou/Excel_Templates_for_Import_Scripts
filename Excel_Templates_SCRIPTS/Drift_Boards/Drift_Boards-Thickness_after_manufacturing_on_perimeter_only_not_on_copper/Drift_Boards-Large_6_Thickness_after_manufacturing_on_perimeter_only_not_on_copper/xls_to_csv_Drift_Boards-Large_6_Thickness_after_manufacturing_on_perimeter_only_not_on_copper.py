# -*- coding: utf-8 -*-
import xlrd
from datetime import datetime

workbook =  xlrd.open_workbook('Drift_Boards-Large_6_Thickness_after_manufacturing_on_perimeter_only_not_on_copper.xls')
sheet = workbook.sheet_by_name('Sheet1')
csvfile = open('Drift_Boards-Large_6_Thickness_after_manufacturing_on_perimeter_only_not_on_copper.csv', 'wb')
"""
writecsv = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
a=['MEASSITEHASH','EQENTRYID','CONTEXTNAME','MEASVALUE','PERCENTAGEMEAS','ISVALIDFLAG','INDEXHASH','MEASTIME','SHIFTER','WEBSITEUSERCR']
b=['',x1.cell(5,0).value,'DT_PAN_PLAN_MIN',x1.cell(29, 0).value,'','T','skfh_1',x1.cell(5,2).ctype,x1.cell(5,4).value,'ggiakous']
c=['',x1.cell(5,0).value,'DT_PAN_PLAN_MAX',x1.cell(29,2).value,'','T','skfh_1',x1.cell(5,2).value,x1.cell(5,4).value,'ggiakous']
d=['',x1.cell(5,0).value,'DT_PAN_PLAN_AVG',x1.cell(29,4).value,'','T','skfh_1',x1.cell(5,2).value,x1.cell(5,4).value,'ggiakous']
e=['',x1.cell(5,0).value,'DT_PAN_PLAN_RMS',x1.cell(29,6).value,'','T','skfh_1',x1.cell(5,2).value,x1.cell(5,4).value,'ggiakous']
writecsv.writerow(a)
writecsv.writerow(b)
writecsv.writerow(c)
writecsv.writerow(d)
writecsv.writerow(e)
csvfile.close()
"""
#############################################################
shifter=sheet.cell(11,0).value
EQID=int(sheet.cell(5,0).value)
date=sheet.cell(8,0).value
datetuple=xlrd.xldate_as_tuple(date,workbook.datemode)
fdate=datetime(*datetuple) #convert date in datetime type
fdate=fdate.strftime("%Y-%m-%d %H:%M:%S") #date to str
print "shifter, EQID,date",shifter,EQID,type(shifter),fdate
resstr=""
resstr+=";".join(['MEASSITEHASH','EQENTRYID','CONTEXTNAME','MEASVALUE','PERCENTAGEMEAS','ISVALIDFLAG','INDEXHASH','MEASTIME','SHIFTER','WEBSITEUSERCR']) +"\n"
resstr+=";".join(["",str(EQID),'DB_THICK_NOT_ON_COP_MIN_DT',str(sheet.cell(20,1).value),"","T",'skfh_1',fdate,shifter,'ggiakous']) +"\n"
resstr+=";".join(["",str(EQID),'DB_THICK_NOT_ON_COP_MAX_DT',str(sheet.cell(20,3).value),"","T",'skfh_1',fdate,shifter,'ggiakous']) +"\n"
resstr+=";".join(["",str(EQID),'DB_THICK_NOT_ON_COP_AVG_DT',str(sheet.cell(20,5).value),"","T",'skfh_1',fdate,shifter,'ggiakous']) +"\n"
resstr+=";".join(["",str(EQID),'DB_THICK_NOT_ON_COP_RMS_DT',str(sheet.cell(20,7).value),"","T",'skfh_1',fdate,shifter,'ggiakous'])


csvfile.write(resstr)
csvfile.close()

pattern="%s;%s;"



#############################################################

#import xlrd
#import os.path

#wb = xlrd.open_workbook(os.path.join('/home/giakou/Desktop/MICROMEGAS_THESSALONIKI_MEASUREMENTS','Central_panel_planarity.xlsx'))
#wb.sheet_names()
#sh = wb.sheet_by_index(0)
#i = 1
#with open("Central_panel_planarity_and_thickness.txt", "a") as my_file:
#    while sh.cell(i,11).value != 0:
#        Load = sh.cell(i,11).value
 #       all_d = sh.col_values(i, 13, 19)
  #      DB1 = Load + " " + (" ".join(all_d))
   #     my_file.write(DB1 + '\n')
    #    i += 1

############################################################

#def xls_to_csv(Central_panel_planarity.xls, Sheet3, Central_panel_planarity_and_thickness.csv):
 #    import xlrd
  #   import csv
   #  workbook = xlrd.open_workbook(Central_panel_planarity.xls)
    # worksheet = workbook.sheet_by_name(Sheet3)
     #csvfile = open(Central_panel_planarity_and_thickness.csv, 'wb')
#     wr = csv.writer(csvfile, quoting=csv.QUOTE_ALL)

#     for rownum in xrange(worksheet.nrows):
 #        wr.writerow(list(x.encode('utf-8') if type(x) == type(u'') else x
  #                for x in worksheet.row_values(rownum))))

   #  csvfile.close()

#
