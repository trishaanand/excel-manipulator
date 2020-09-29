from openpyxl import load_workbook, worksheet, Workbook
from collections import namedtuple

import os
import sys, getopt

def main(argv):
    inputfile = ''
    outputfile = ''
    try:
        opts, args = getopt.getopt(argv,"hi:o:",["ifile=","ofile="])
    except getopt.GetoptError:
        print 'manipulator.py -i <inputfile> -o <outputfile>'
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print 'manipulator.py -i <inputfile> -o <outputfile>'
            sys.exit()
        elif opt in ("-i", "--ifile"):
            inputfile = arg
        elif opt in ("-o", "--ofile"):
            outputfile = arg

    theFile = load_workbook(inputfile)
    allSheetNames = theFile.sheetnames

    excel_file = Workbook()
    excel_sheet = excel_file.create_sheet(title='Result', index=0)

    Influencer = namedtuple("Influencer", "name contact occupation party")
    VillageDetail = namedtuple("VillageDetails", "uid district ac block village")

    for sheet in allSheetNames:
        seen = set()
        result = {}
        villageDetails = {}
        
        currentSheet = theFile[sheet]
        
        for row in range(2, currentSheet.max_row + 1):
            
            villageName = currentSheet.cell(row = row, column = 7).value
            if villageName not in seen :
                seen.add(villageName)
                result[villageName] = []
                villageDetails[villageName] = VillageDetail(currentSheet.cell(row = row, column = 3).value, 
                currentSheet.cell(row = row, column = 4).value, 
                currentSheet.cell(row = row, column = 5).value, 
                currentSheet.cell(row = row, column = 6).value,
                currentSheet.cell(row = row, column = 7).value)

            
            Influencers = []
            Influencers.append(Influencer(currentSheet.cell(row = row, column = 8).value, currentSheet.cell(row = row, column = 9).value, currentSheet.cell(row = row, column = 10).value, currentSheet.cell(row = row, column = 11).value))
            Influencers.append(Influencer(currentSheet.cell(row = row, column = 12).value, currentSheet.cell(row = row, column = 13).value, currentSheet.cell(row = row, column = 14).value, currentSheet.cell(row = row, column = 15).value))
            Influencers.append(Influencer(currentSheet.cell(row = row, column = 16).value, currentSheet.cell(row = row, column = 17).value, currentSheet.cell(row = row, column = 18).value, currentSheet.cell(row = row, column = 19).value))
            Influencers.append(Influencer(currentSheet.cell(row = row, column = 20).value, currentSheet.cell(row = row, column = 21).value, currentSheet.cell(row = row, column = 22).value, currentSheet.cell(row = row, column = 23).value))
            Influencers.append(Influencer(currentSheet.cell(row = row, column = 24).value, currentSheet.cell(row = row, column = 25).value, currentSheet.cell(row = row, column = 26).value, currentSheet.cell(row = row, column = 27).value))

            Influencers = [x for x in Influencers if not x.name is None]

            result[villageName] += Influencers

        # print(str(villageDetails))

    for village in seen :
        rowData = []
        details = villageDetails[village]
        influencers = result[village]
        rowData.append(details.uid)
        rowData.append(details.district)
        rowData.append(details.ac)
        rowData.append(details.block)
        rowData.append(details.village)
        for influencer in influencers :
            rowData.append(influencer.name)
            rowData.append(influencer.contact)
            rowData.append(influencer.occupation)
            rowData.append(influencer.party)
        
        excel_sheet.append(rowData)

    excel_file.save(outputfile)
    print("A new file created for village count " + str(len(villageDetails)))

if __name__ == "__main__":
   main(sys.argv[1:])

    


