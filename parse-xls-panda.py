import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import json
import math
import xlwt
import collections

fileName = '03-06-2019_QL_Equip_List.xlsx'
sheetName = 'Eq List'

eqList = pd.read_excel(fileName, sheet_name=sheetName)

found = {}
FOUND_SECTION = ""
for row_index, row in eqList.iterrows():
    colNumber = -1
    section = {}
    while colNumber < (len(row) - 1):
      colNumber += 1
      # print("CELL @ %i/%i: %s" % (row_index, colNumber, row[colNumber]))
      if row[colNumber] == "Client:":
        found["client"] = row[colNumber + 1]
      elif row[colNumber] == "P.O. NO:":
        found["poBox"] = row[colNumber + 1]
      elif row[colNumber] == "JOB NO.:":
        # print("FOUND! JOB NO.: " + row[colNumber + 1])
        found["jobNumber"] = row[colNumber + 1]
      elif row[colNumber] == "DIGITAL SYSTEM LIST":
        FOUND_SECTION = "DIGITAL_SYSTEM_LIST"
        found["sections"] = {}
        found["sections"]["DIGITAL_SYSTEM_LIST"] = []
      elif row[colNumber] == "MECHANICAL SYSTEM LIST":
        FOUND_SECTION = "MECHANICAL_SYSTEM_LIST"
        found["sections"]["MECHANICAL_SYSTEM_LIST"] = []
      elif row[colNumber] == "SENSOR LIST":
        FOUND_SECTION = "SENSOR_LIST"
        found["sections"]["SENSOR_LIST"] = []
      elif row[colNumber] == "Miscellaneous LIST":
        FOUND_SECTION = "Miscellaneous_LIST"
        found["sections"]["Miscellaneous_LIST"] = []
      elif row[colNumber] == "Device Type":
        continue
      else:
        if FOUND_SECTION == "DIGITAL_SYSTEM_LIST" and colNumber == 0:
          if type(row[colNumber]) == float and math.isnan(row[colNumber]): # and not row[colNumber + 1] and not row[colNumber + 2] and not row[colNumber + 3]:
            continue
          section = {
            "DeviceType": row[colNumber],
            "Description": row[colNumber + 1],
            "MakeModel": row[colNumber + 2],
            "Range": row[colNumber + 3],
            "AssetNumber": row[colNumber + 4],
            "SerialNumber": row[colNumber + 5],
            "DueDate": str(row[colNumber + 6])
          }
          found["sections"]["DIGITAL_SYSTEM_LIST"].append(section)
        elif FOUND_SECTION == "MECHANICAL_SYSTEM_LIST" and colNumber == 0:
          if type(row[colNumber]) == float and math.isnan(row[colNumber]): # and not row[colNumber + 1] and not row[colNumber + 2] and not row[colNumber + 3]:
            continue
          section = {
            "DeviceType": row[colNumber],
            "Description": row[colNumber + 1],
            "MakeModel": row[colNumber + 2],
            "Range": row[colNumber + 3],
            "AssetNumber": row[colNumber + 4],
            "SerialNumber": row[colNumber + 5],
            "DueDate": str(row[colNumber + 6])
          }
          found["sections"]["MECHANICAL_SYSTEM_LIST"].append(section)
        elif FOUND_SECTION == "SENSOR_LIST" and colNumber == 0:
          if type(row[colNumber]) == float and math.isnan(row[colNumber]): # and not row[colNumber + 1] and not row[colNumber + 2] and not row[colNumber + 3]:
            continue
          section = {
            "DeviceType": row[colNumber],
            "Description": row[colNumber + 1],
            "MakeModel": row[colNumber + 2],
            "Range": row[colNumber + 3],
            "AssetNumber": row[colNumber + 4],
            "SerialNumber": row[colNumber + 5],
            "DueDate": str(row[colNumber + 6])
          }
          found["sections"]["SENSOR_LIST"].append(section)
        elif FOUND_SECTION == "Miscellaneous_LIST" and colNumber == 0:
          if type(row[colNumber]) == float and math.isnan(row[colNumber]): # and not row[colNumber + 1] and not row[colNumber + 2] and not row[colNumber + 3]:
            continue
          section = {
            "DeviceType": row[colNumber],
            "Description": row[colNumber + 1],
            "MakeModel": row[colNumber + 2],
            "Range": row[colNumber + 3],
            "AssetNumber": row[colNumber + 4],
            "SerialNumber": row[colNumber + 5],
            "DueDate": str(row[colNumber + 6])
          }
          found["sections"]["Miscellaneous_LIST"].append(section)

print(json.dumps(found))

reportFileName = "test-report-output.xls"
sheetName = "SummaryReport"


# JOB: client
# JOB: poBox
# JOB: jobNumber
# JOB: sections
def outputExcelReport(filename, sheet, data):
  book = xlwt.Workbook()
  sh = book.add_sheet(sheet)
  
  i = 0
  sh.write(i,0,data["client"])
  sh.write(i,1,data["jobNumber"])

  for section in data["sections"]:
    i += 1
    print("section: " + str(section))
    # if isinstance(data[key], collections.Mapping):
    #   pass
    # else:
    #   sh.write(i,0,key)
    #   sh.write(i,1,data[key])

  book.save(filename)

outputExcelReport(reportFileName, sheetName, found)


def output(filename, sheet, list1, list2, x, y, z):
    book = xlwt.Workbook()
    sh = book.add_sheet(sheet)

    variables = [x, y, z]
    x_desc = 'Display'
    y_desc = 'Dominance'
    z_desc = 'Test'
    desc = [x_desc, y_desc, z_desc]

    col1_name = 'Stimulus Time'
    col2_name = 'Reaction Time'

    #You may need to group the variables together
    #for n, (v_desc, v) in enumerate(zip(desc, variables)):
    for n, v_desc, v in enumerate(zip(desc, variables)):
        sh.write(n, 0, v_desc)
        sh.write(n, 1, v)

    n+=1

    sh.write(n, 0, col1_name)
    sh.write(n, 1, col2_name)

    for m, e1 in enumerate(list1, n+1):
        sh.write(m, 0, e1)

    for m, e2 in enumerate(list2, n+1):
        sh.write(m, 1, e2)

    book.save(filename)