import glob
import sys
import xml.etree.ElementTree as ET
import xlsxwriter
from pathlib import Path
import os

# Read filenames recursively

workbook = xlsxwriter.Workbook('Test.xlsx')
worksheet = workbook.add_worksheet()  # Create new sheet

# Declare row/column of the sheet
i = -1
contentColumn = 0
titleColumn = 0
titleRow = 0


def fileFunction(name):
    row = 1
    column = 0
    files = Path().cwd().glob("**/*.cgxml")
    for filename in files:
        base = os.path.basename(filename)
        if(os.stat(filename).st_size != 0):
            with open(filename, 'r', encoding="utf-8") as content:
                mytree = ET.parse(content)
                myroot = mytree.getroot()

            for x in myroot.findall(wordList[0]):

                if(x.find(name) != None):
                    worksheet.write(titleRow, titleColumn, "Filename")
                    worksheet.write(titleRow, titleColumn +
                                    contentColumn, name)
                    worksheet.write(row, column, base)
                    worksheet.write(row, contentColumn, x.find(name).text)

                row += 1


with open('configFile.txt', 'r') as file:
    wordList = list()
    line = file.readline()

    for word in line.split():
        wordList.append(word)

while(len(wordList) > 1):
    contentColumn += 1
    fileFunction(wordList[i])
    wordList.pop()
    print("-------------------------------")


workbook.close()
print("That's all")
