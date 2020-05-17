"""
File: CATechoPooler.py
----------------
This program takes in a pooling file with echo transfers calculated and
outputs a echo pooling worksheet and a byHand worksheet by calculating the
destination wells.

call using:
 python3 ./CATechoPooler.py COMET384_Seq7_Echo_Calculations.xlsx testing10
"""
import pandas as pd
import sys
MAX_TRANSFER_VOL_BY_ECHO = 11 #cut-off between echo vs by-hand pooling
MAX_DESTINATION_WELL_VOL = 15 #max volume destination plate well can hold


def main():
    FILENAME = sys.argv[1]
    filebase = sys.argv[2]
    PATH = "./"
    if(len(sys.argv) == 4):
        PATH = str(sys.argv[3])
        if PATH[len(PATH)-1] != "/":
            PATH = str(sys.argv[3]) + "/"
    dataWithDestination = generateDestinationWells(FILENAME, filebase, PATH)
    echoPoolDoc(dataWithDestination, filebase, PATH)
    byHandDoc(dataWithDestination, filebase, PATH)

def generateDestinationWells(FILENAME, filebase, PATH):
    file = pd.read_excel(FILENAME, index_col=None)
    file.loc[:,"byHand"] = file["uL"] > MAX_TRANSFER_VOL_BY_ECHO
    volume = []
    destinationRowList = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    destinationRowIndex = 0
    location = []
    destinationColumn = 1
    currVolumeTotal = 0
    for index, row in file.iterrows():
        if row["uL"] > MAX_TRANSFER_VOL_BY_ECHO:
            volume.append(0)
            location.append("Add by Hand")
            continue
        elif (currVolumeTotal + row["uL"] > MAX_DESTINATION_WELL_VOL):
            destinationColumn += 1
            if destinationColumn >= 25:
                destinationRowIndex += 1
                destinationColumn = 1
            currVolumeTotal = 0
        currVolumeTotal += row["uL"]
        volume.append(row["uL"])
        location.append(destinationRowList[destinationRowIndex]+str(destinationColumn))
    file.loc[:,'volume'] = volume
    file.loc[:,'Destination Well'] = location
    file.to_excel(PATH + filebase +"_Pool.xlsx")
    return file

def echoPoolDoc(DATAFRAME, filebase, PATH):
    echoSubset = DATAFRAME[DATAFRAME['byHand'] != True]
    echoSubset.loc[:,'volume'] = echoSubset.loc[:,'volume'].astype(float)
    echoSubset.loc[:,"SOURCE WELL"] = echoSubset.loc[:,"Source Well"]
    echoSubset.loc[:,"TRANSFER VOLUME"] = echoSubset.loc[:,"volume"] * 1000
    echoSubset.loc[:,"DESTINATION WELL"] = echoSubset.loc[:,"Destination Well"]
    echoSubset.loc[:,"TRANSFER VOLUME"] = echoSubset.loc[:,"TRANSFER VOLUME"].astype(int)
    finalEcho = echoSubset.loc[:,["SOURCE WELL", "TRANSFER VOLUME", "DESTINATION WELL"]]
    finalEcho.to_csv(PATH + filebase + "_CATEchoPoolingDoc.csv", index=False)

def byHandDoc(DATAFRAME, filebase, PATH):
    echoHand = DATAFRAME[DATAFRAME['byHand'] == True]
    echoHand = echoHand[['Source Well', "uL"]]
    echoHand.loc[:,'uL'] = echoHand['uL'].astype(int)
    echoHand.to_csv(PATH + filebase + "_ByHandPoolingDoc.csv", index=False)

if __name__ == '__main__':
    main()
