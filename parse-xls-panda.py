import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import json
import math

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
