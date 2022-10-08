import os
import re
import xml.etree.ElementTree as ET
from docx_utils.flatten import opc_to_flat_opc

class Vendor(object):
    name = "N/A"
    number = -1

    def __init__(self, name, number):
        self.name = name
        self.number = number


def getListOfFiles(dirName , fileType):
    listOfFile = os.listdir(dirName)
    allFiles = list()
    i = 0
    for entry in listOfFile:
        fullPath = os.path.join(dirName, entry)
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFiles(fullPath , fileType)
        else:
            if fileType in entry:
                allFiles.append(fullPath)
                i = i + 1

    return allFiles


def getListOfFileNames(dirName, fileType):
    listOfFile = os.listdir(dirName)
    allFiles = list()
    i = 0
    for entry in listOfFile:
        fullPath = os.path.join(dirName, entry)
        if os.path.isdir(fullPath):
            allFiles = allFiles + getListOfFileNames(fullPath , fileType)
        else:
            if fileType in entry:
                allFiles.append(os.path.basename(entry))
                i = i + 1

    return allFiles


def docToXml(dirName):
    docs = getListOfFiles(dirName, '.docx')
    for d in docs:
        opc_to_flat_opc(d, d.replace('.docx','.xml'))
        print(d)


def appendAllContent():
    with open('D:\\tt\\word\\t1.txt', "r", encoding='UTF8') as f:
        s = f.readline()
        while s:
            number = re.findall('\t\d+\n',s)
            if number.__len__()>0:
                numberPos = s.find(number[0])
                name = s[0:numberPos]
                print(name + ' -' + number[0])
            s = f.readline()


def getVendorsData():
    with open('D:\\tt\\word\\t1.txt', "r", encoding='UTF8') as f:
        vendors = list()
        s = f.readline()
        while s:
            number = re.findall('\t\d{1,}\n', s)
            if number.__len__() > 0:
                numberPos = s.find(number[0])
                name = s[0:numberPos]
                vendors.append(Vendor(name,number[0].strip()))
            s = f.readline()
        return vendors


def getVendorNumberByName(vendors, fileName):
    try:
        with open(fileName, "r", encoding='UTF8') as f:
            textContains = f.read()
            for v in vendors:
                if v.number!='451' and v.name.upper() in textContains.upper():
                    return v.number


            if 'ФЕДЕРАЛЬНОЕ КАЗНАЧЕЙСТВО' in textContains.upper():
                return '_byhands'
    except FileNotFoundError:
        print('boohoo')

    return 'N/A'


directory = 'D:\\tt\\word\\files'
filetype = '.docx'
vendors = getVendorsData()


for f in getListOfFiles(directory, filetype):
    fileFullName = os.path.basename(f)
    numbers = re.findall('\d{1,}\.docx', f)
    if numbers.__len__() > 0:
        numberPos = fileFullName.find(numbers[0])
        number = numbers[0].replace('.docx', '')
        fileName = fileFullName[0:numberPos].replace('2.6', '').replace('2.2', '').strip()
        correctNumber = getVendorNumberByName(vendors, f.replace('docx','txt'))
        correctness = number == correctNumber.zfill(4)
        if not correctness and correctNumber != '_byhands' and correctNumber != 'N/A':
            try:
                os.rename(f, f.replace(number,correctNumber.zfill(4)))
            except FileExistsError:
                try:
                    os.rename(f, f.replace(number, correctNumber.zfill(4))+'_byhands')
                except FileExistsError:
                    os.rename(f, f.replace(number, correctNumber.zfill(4)) + '_byhands2')
            print (f, f.replace(number,correctNumber.zfill(4)))
       


#docToXml(directory)
#print(getListOfFiles(dir,filetype))