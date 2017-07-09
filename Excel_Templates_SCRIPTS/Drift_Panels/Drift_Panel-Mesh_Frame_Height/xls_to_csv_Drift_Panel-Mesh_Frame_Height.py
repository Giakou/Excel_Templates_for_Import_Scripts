# -*- coding: utf-8 -*-
import xlrd
from datetime import datetime

workbook =  xlrd.open_workbook('Drift_Panel-Mesh_Frame_Height_5384_side_1.xls')
sheet = workbook.sheet_by_name('Sheet1')
csvfile = open('Drift_Panel-Mesh_Frame_Height_5384_side_1.csv', 'wb')
shifter=sheet.cell(11,0).value
EQID=int(sheet.cell(5,0).value)
date=sheet.cell(8,0).value
position=sheet.cell(14,0).value
datetuple=xlrd.xldate_as_tuple(date,workbook.datemode)
fdate=datetime(*datetuple) #convert date in datetime type
fdate=fdate.strftime("%Y-%m-%d %H:%M:%S") #date to str
print "shifter, EQID,date",shifter,EQID,type(shifter),fdate
resstr=""
resstr+=";".join(['MEASSITEHASH','EQENTRYID','CONTEXTNAME','POSITION','MEASVALUE','PERCENTAGEMEAS','ISVALIDFLAG','INDEXHASH','MEASTIME','SHIFTER','WEBSITEUSERCR']) +"\n"
resstr+=";".join(["",str(EQID),'DP_MESH_FR_HEIGHT_MIN_DT',position,str(sheet.cell(33,1).value),"","T",'skfh_1',fdate,shifter,'ggiakous']) +"\n"
resstr+=";".join(["",str(EQID),'DP_MESH_FR_HEIGHT_MAX_DT',position,str(sheet.cell(33,3).value),"","T",'skfh_1',fdate,shifter,'ggiakous']) +"\n"
resstr+=";".join(["",str(EQID),'DP_MESH_FR_HEIGHT_AVG_DT',position,str(sheet.cell(33,5).value),"","T",'skfh_1',fdate,shifter,'ggiakous']) +"\n"
resstr+=";".join(["",str(EQID),'DP_MESH_FR_HEIGHT_RMS_DT',position,str(sheet.cell(33,7).value),"","T",'skfh_1',fdate,shifter,'ggiakous'])


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
