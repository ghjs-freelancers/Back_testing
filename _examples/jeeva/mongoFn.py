from typing import OrderedDict
import pymongo
import csv
import openpyxl
from openpyxl import load_workbook
from pprint import pprint

client = pymongo.MongoClient("localhost",27017)
db = client["back_testing"]
excel = db["excel"]

def dbInsert():
    json={
        "location":"C:/Users/Jeevanatham N/Documents/csv&xlsx/sample.csv",
        "filename":"sample",
        "type":"csv"
    }
    excel.insert_one(json)

def getFile():
    files = excel.find({"filename":"sample","type":"csv"})
    for file in files:
        if (file["type"] == "xlsx"):
            excelParse(file["location"],file["filename"])
        elif (file["type"] == "csv"):
            csvParse(file["location"],file["filename"])

def excelParse(filepath,filename):
    excelValue = []
    json = {}
    wb=load_workbook(filepath)
    sheet=wb.active
    max_row=sheet.max_row
    max_column=sheet.max_column
    for i in range(2,max_row+1):
        json = {}
        for j in range(1,max_column+1):
            cell_obj=sheet.cell(row=i,column=j)
            cell_title=sheet.cell(row=1,column=j)
            json[cell_title.value] = cell_obj.value
        excelValue.append(json)
        
def csvParse(filepath,filename):
    with open(filepath, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        csv_reader = list(csv_reader)
    overrideCsv(filepath,filename,csv_reader)

def overrideCsv(filepath,filename,csv_reader):
    with open(filepath, mode='w') as csv_file:
        fieldnames = []
        for i in csv_reader:
            print(i)
        for key in csv_reader[0]:
            fieldnames.append(key)
        writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
        writer.writeheader()
        for i in csv_reader:
            writer.writerow(i)

def overrideExcel():
    pass

if __name__ == "__main__":
    getFile()
