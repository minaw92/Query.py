import xlrd
import xlsxwriter
import numpy as np
import time
import cx_Oracle
import os
os.environ['PATH']=("instantclient_18_5")

ip = '69.134.42.124'
port = 1521
SID = 'spldg1'
dsn_tns = cx_Oracle.makedsn(ip, port, SID)
temp=[]
connection= cx_Oracle.connect('DLSEEOPORTCAPACITY', 'D1_3XHyu', dsn_tns)
cur = connection.cursor()
cur.execute("SELECT S5.SITE_HUM_ID as Region,S3.SITE_HUM_ID as KMA,S1.STATE_PROV as State,C1.CIRC_PATH_HUM_ID AS PARENT_ID, C1.STATUS, C1.NBR_CHANNELS, PATH_CHAN_INST.CHAN_NAME, C2.CIRC_PATH_HUM_ID AS MEMBER_ID, C3.CIRC_PATH_HUM_ID AS NEXT_MEMBER_ID, C1.BANDWIDTH, S1.SITE_HUM_ID AS A_SITE, S2.SITE_HUM_ID AS Z_SITE FROM ISE.circ_path_inst C1 LEFT JOIN    ISE.PATH_CHAN_INST ON (PATH_CHAN_INST.PARENT_PATH_INST_ID = C1.CIRC_PATH_INST_ID)LEFT JOIN ISE.SITE_INST S1 ON (C1.A_SIDE_SITE_ID = S1.SITE_INST_ID) LEFT JOIN ISE.SITE_INST S3 ON (S1.PARENT_SITE_INST_ID = S3.SITE_INST_ID) LEFT JOIN ISE.SITE_INST S5 ON (S3.PARENT_SITE_INST_ID = S5.SITE_INST_ID) LEFT JOIN ISE.SITE_INST S2 ON (C1.Z_SIDE_SITE_ID = S2.SITE_INST_ID) LEFT JOIN ISE.SITE_INST S4 ON (S2.PARENT_SITE_INST_ID = S4.SITE_INST_ID) LEFT JOIN ISE.SITE_INST S6 ON (S4.PARENT_SITE_INST_ID = S6.SITE_INST_ID) LEFT JOIN ISE.circ_path_inst C2 ON (C2.CIRC_PATH_INST_ID = PATH_CHAN_INST.MEMBER_PATH_INST_ID) LEFT JOIN ISE.circ_path_inst C3 ON (C3.CIRC_PATH_INST_ID = PATH_CHAN_INST.NEXT_PATH_INST_ID) WHERE( C1.circ_path_hum_id LIKE '%OM10%' OR C1.circ_path_hum_id LIKE '%OFX%' OR C1.circ_path_hum_id LIKE '%OM20%' OR C1.circ_path_hum_id LIKE '%OMC%')ORDER BY PARENT_ID, PATH_CHAN_INST.CHAN_NAME")
rowsinit = cur.fetchall()
for rowinit in rowsinit:
    temp.append(rowinit)

cur.close()
connection.close()
print("SQL read-- success ! ")

intial_list=np.array(temp)
workbookintial = xlsxwriter.Workbook('All_circuits.xlsx')
worksheetintial = workbookintial.add_worksheet()
rowintial2 = 1

worksheetintial.write_string(0, 0, "Region")
worksheetintial.write_string(0, 1, "KMA")
worksheetintial.write_string(0, 2, "State")
worksheetintial.write_string(0, 3, "PARENT_ID")
worksheetintial.write_string(0, 4, "STATUS")
worksheetintial.write_string(0, 5, "NBR_CHANNELS")
worksheetintial.write_string(0, 6, "CHAN_NAME")
worksheetintial.write_string(0, 7, "MEMBER_ID")
worksheetintial.write_string(0, 8, "Next_Path")
worksheetintial.write_string(0, 9, "BANDWIDTH")
worksheetintial.write_string(0, 10, "A_SITE")
worksheetintial.write_string(0, 11, "Z_SITE")
for xintial in (intial_list):

    worksheetintial.write_string(rowintial2, 0,str(xintial[0]))
    worksheetintial.write_string(rowintial2, 1, str(xintial[1]))
    worksheetintial.write_string(rowintial2, 2, str(xintial[2]))
    worksheetintial.write_string(rowintial2, 3, str(xintial[3]))
    worksheetintial.write_string(rowintial2, 4, str(xintial[4]))
    worksheetintial.write_number(rowintial2, 5, int(xintial[5]))
    worksheetintial.write_string(rowintial2, 6, str(xintial[6]))
    worksheetintial.write_string(rowintial2, 7, str(xintial[7]))
    worksheetintial.write_string(rowintial2, 8, str(xintial[8]))
    worksheetintial.write_string(rowintial2, 9, str(xintial[9]))
    worksheetintial.write_string(rowintial2, 10, str(xintial[10]))
    worksheetintial.write_string(rowintial2, 11, str(xintial[11]))
    rowintial2=rowintial2+1

workbookintial.close()

print("All_circuits sheet-- success ! ")

def removeDuplicates(listofElements):
    uniqueList = []

    for elem in listofElements:
        if elem not in uniqueList:
            uniqueList.append(elem)

    return uniqueList


All_src = r'All_circuits.xlsx'
All_wb = xlrd.open_workbook(All_src)
ALL_sheet = All_wb.sheet_by_index(0)

ALL_sheet.cell_value(0, 0)
COM_PARENT_ID = []
for ii in range(ALL_sheet.nrows):
    if ((".TWCC" in ALL_sheet.cell_value(ii, 7)) or (".TWCC" in ALL_sheet.cell_value(ii, 8))):

        if ((ALL_sheet.cell_value(ii, 5) == 4) or (ALL_sheet.cell_value(ii, 5) == 10) or (
                ALL_sheet.cell_value(ii, 5) == 20) or (ALL_sheet.cell_value(ii, 5) == 1)):
            #print("x")


            if ("HUB" in (ALL_sheet.cell_value(ii, 11))) or ("COLO" in (ALL_sheet.cell_value(ii, 11)))or ("HE " in (ALL_sheet.cell_value(ii, 11))):
                #print(ALL_sheet.cell_value(ii, 3))
                aaa = ALL_sheet.cell_value(ii, 3)
                #print(aaa)
                COM_PARENT_ID.append([aaa])

COM_PARENT_ID_R = removeDuplicates(COM_PARENT_ID)
#print(COM_PARENT_ID_R)



Com_R = len(COM_PARENT_ID_R)
# print(Com_R)
x = 0
y = 1
i = 0
empty = 0
arr = []

while i < (ALL_sheet.nrows):

    if (ALL_sheet.cell_value(y, 3) in (COM_PARENT_ID_R[(x)])):

        if (ALL_sheet.cell_value(y, 5) == 1):
            for mm in range(1):
                if ((ALL_sheet.cell_value(y + mm, 7)=="None") and (ALL_sheet.cell_value(y + mm, 8)=="None")):
                    empty = empty + 1
        if (ALL_sheet.cell_value(y, 5) == 4):

            for mm in range(4):
                if ((ALL_sheet.cell_value(y + mm, 7)=="None") and (ALL_sheet.cell_value(y + mm, 8)=="None")):
                    empty = empty + 1

        if (ALL_sheet.cell_value(y, 5) == 10):
            for mm in range(10):
                if ((ALL_sheet.cell_value(y + mm, 7)=="None") and (ALL_sheet.cell_value(y + mm, 8)=="None")):
                    empty = empty + 1

        if (ALL_sheet.cell_value(y, 5) == 20):
            for mm in range(20):
                if ((ALL_sheet.cell_value(y + mm, 7)=="None") and (ALL_sheet.cell_value(y + mm, 8)=="None")):
                    empty = empty + 1
        aa = (COM_PARENT_ID_R[(x)])

        arr.append([ALL_sheet.cell_value(y, 0), ALL_sheet.cell_value(y, 1), ALL_sheet.cell_value(y, 2), aa, empty,
                    ALL_sheet.cell_value(y, 5), ALL_sheet.cell_value(y, 10), ALL_sheet.cell_value(y, 11),
                    ALL_sheet.cell_value(y, 4), ALL_sheet.cell_value(y, 9)])

        empty = 0

        x = x + 1
        y = 1
        i = 0
        distcom = 0
    y = y + 1
    i = i + 1
    if (x == Com_R):
        # print("break")
        break
#print(arr)

final_list=np.array(arr)

#print(final_list)
timestr = time.strftime("%m-%d-%Y")

workbook = xlsxwriter.Workbook('Wave Query  ' +timestr + '.xlsx')

worksheet = workbook.add_worksheet()
row = 0
col = 0
width = len("['32001.OM20.LSAPCAWVOT1.CYPRCABWOT2']       ")
LOCwidth = len("LSANCARC-TWC/COLO-CORESITE-ONE WILSHIRE-LA1 (BACKBONE) (LSANCA3) (LAX00)     ")
CHwidth = len("# Free Channels")
worksheet.set_column(0, 0, width)
worksheet.set_column(0, 1, width)
worksheet.set_column(0, 2, CHwidth)
worksheet.set_column(0, 3, width)
worksheet.set_column(0, 4, CHwidth)
worksheet.set_column(0, 5, CHwidth)
worksheet.set_column(0, 6, LOCwidth)
worksheet.set_column(0, 7, LOCwidth)
worksheet.set_column(0, 8, width)
worksheet.set_column(0, 9, width)

worksheet.write_string(row, 0, "Region")
worksheet.write_string(row, 1, "KMA")
worksheet.write_string(row, 2, "State")
worksheet.write_string(row, 3, "ckt ID")
worksheet.write_string(row, 4, "# Free Channels")
worksheet.write_string(row, 5, "# of Total channels")
worksheet.write_string(row, 6, "ALOC")
worksheet.write_string(row, 7, "ZLOC")
worksheet.write_string(row, 8, "Status")
worksheet.write_string(row, 9, "Bandwidth")


row = 1
col = 0
for Region, KMA, STATE, CKTID, free, TotalCH, ALOC, ZLOC, STATUS, BW in (final_list):
    worksheet.set_column(row, 0, width)
    worksheet.set_column(row, 1, width)
    worksheet.set_column(row, 2, CHwidth)
    worksheet.set_column(row, 3, width)
    worksheet.set_column(row, 4, CHwidth)
    worksheet.set_column(row, 5, CHwidth)
    worksheet.set_column(row, 6, LOCwidth)
    worksheet.set_column(row, 7, LOCwidth)
    worksheet.set_column(row, 8, width)
    worksheet.set_column(row, 9, width)

    worksheet.write_string(row, col, str(Region))
    worksheet.write_string(row, col + 1, str(KMA))
    worksheet.write_string(row, col + 2, str(STATE))
    worksheet.write_string(row, col + 3, str(CKTID)[2:-2])
    worksheet.write_number(row, col + 4, int(free))
    worksheet.write_number(row, col + 5, int(TotalCH))
    worksheet.write_string(row, col + 6, str(ALOC))
    worksheet.write_string(row, col + 7, str(ZLOC))
    worksheet.write_string(row, col + 8, str(STATUS))
    worksheet.write_string(row, col + 9, str(BW))


    row += 1

workbook.close()

print("project-- success ! ")