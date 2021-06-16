import xlsxwriter
import pandas as pd
import numpy as np
import sys
import os
import time
import glob

#####===================================#####
#@TODO: CREATE FUNCTION FOR FILEEXISTS AND PASS FILEPATH AS PARAM
#@TODO: CREATE FUNCTION FOR IO CHECKOUT THAT INCREMENTS EACH CB BY 3
#@TODO: CREATE MONGODB TO USE WITH THIS APPLICATION?
#####===================================#####

def main():
    os.system('cls')
    print("Ignition's Lazy Script")
    print("1.) Create Sequence checkout checklist")
    print("2.) Create HMI checkout checklist")
    print("3.) Create I/O checkout checklist")
    print("4.) Create Safety checkout checklist")
    print("5.) Exit")

    sldOption = input("Enter option: ")

    if(sldOption == "1"):
        seqValidationChecklist()
    if(sldOption == "2"):
        HMI_Checklist()
    if(sldOption == "3"):
        IO_Checklist()
    if(sldOption == "4"):
        Safety_Checklist()
    elif(sldOption == "5"):
        sys.exit()

def HMICheckout_Part1():
    # //************************ STATIC SECTION ************************//
    out_table = []

    #004 SECTION
    out_row = {}
    out_row['action'] = "004 System Overview Help"
    out_row["initial"] = "INT / DATE"
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "HMI# and span of control, system name"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Time correct"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #005 SECTION
    out_row = {}
    out_row['action'] = "005 Controller Status"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "All controller fields filled in correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #006 SECTION
    out_row = {}
    out_row['action'] = "006 Line Faults"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "System, zone, and station banners are shown correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #010 SECTION
    out_row = {}
    out_row['action'] = "010 System Overview"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Orientation matches physical layout"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "GOTO group box functions correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Red fault box around GOTO functions correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Sound Horn"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Lamp Test"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Column locations correct"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #015 SECTION
    out_row = {}
    out_row['action'] = "015 Group Select AZx"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Orientation matches physical layout"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "GOTO group box functions correctly to associated leader drive"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "All CHBs are functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "All E-Stops/Safety mats are functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Column locations correct"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "IFD name correct and status IND. functions correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #025 SECTION
    out_row = {}
    out_row['action'] = "025 Hardwire Status 1 - Process"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "System and all zone banners display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #026 SECTION
    out_row = {}
    out_row['action'] = "026 Hardwire Status 2 - Process"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "System and all zone banners display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #030 SECTION
    out_row = {}
    out_row['action'] = "030 Hardwire Status 3 - Safety"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "System and all zone banners display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #035 SECTION
    out_row = {}
    out_row['action'] = "035 Screen Menu"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Config menu correct"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "All GOTO station buttons go to correct drive screen"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #039 SECTION
    out_row = {}
    out_row['action'] = "039 System Cycle Time"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #041 SECTION
    out_row = {}
    out_row['action'] = "041 Production Counts"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #045 SECTION
    out_row = {}
    out_row['action'] = "045 Carrier Counts"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Orientation matches physical layout"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #090 SECTION
    out_row = {}
    out_row['action'] = "090 Drive Status"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #092 SECTION
    out_row = {}
    out_row['action'] = "092 IO Block Diagnostics"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #093 SECTION
    out_row = {}
    out_row['action'] = "093 Safety Block Diagnostics"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #094 SECTION
    out_row = {}
    out_row['action'] = "094 Station Data"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #097 SECTION
    out_row = {}
    out_row['action'] = "097 Interlocks Produced"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #098 SECTION
    out_row = {}
    out_row['action'] = "098 Interlocks Consumed"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #100 SECTION
    out_row = {}
    out_row['action'] = "100 Alarm History"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Text and indicators display correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Menu buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    dfOutput = pd.DataFrame(out_table)
    dfOutput.to_excel("./hmi_checkout/output_part1.xlsx", sheet_name="Sheet1", header=False, index=False)

    return True

def HMICheckout_Part2(iterationNumber):
    # //************************ DRIVE NUMBERS ************************//

    dfDriveConfig = pd.read_excel("./hmi_checkout/drive_config.xlsx", sheet_name="Sheet1")

    out_table = []

    for index, row in dfDriveConfig.iterrows():
        drive_id = row["drive_id"]
        hmi_control = row["hmi_control"]

        if(hmi_control == iterationNumber):
            out_row = {}
            out_row["header"] = str(drive_id) + " Manual"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "Title display correct"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "All GOTO buttons functional"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "Flow correct to HMI"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "PRX / PE name correct and in correct position"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "JOG functions correct for drive and station"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "Ghostbuster on leader drive functions correctly"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "Next & Previous GOTO functions correctly"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "Drive data functions correctly"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "IO indicator functions correctly"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "Manual select button puts station in manual(not system)"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = "Drive status displays correctly"
            out_row["initial"] = " "
            out_table.append(out_row)

            out_row = {}
            out_row["header"] = " "
            out_row["initial"] = " "
            out_table.append(out_row)

    dfOutput = pd.DataFrame(out_table)
    dfOutput.to_excel("./hmi_checkout/output_part2.xlsx", header=False, index=False)

    return True

def HMICheckout_Part3():
    # //************************ NETWORK ************************//
    out_table = []

    #500 SECTION
    out_row = {}
    out_row['action'] = "500 Network Node Layout AZx"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Orientation matches physical layout"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "GOTO group box functions correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Node name correct"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Node indicators function correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "All E-Stop, HMIs, CHBs, and SCBs are shown correctly"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #510 SECTION
    out_row = {}
    out_row['action'] = "510 ENet Switch Layout 100"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "Title display correct"
    out_row["initial"] = " "
    out_table.append(out_row)
    
    out_row = {}
    out_row['action'] = "All GOTO buttons and fault indicators functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = "Switch status name and status indicators functional"
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row['action'] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    dfOutput = pd.DataFrame(out_table)
    dfOutput.to_excel("./hmi_checkout/output_part3.xlsx", header=False, index=False)

    return True

def HMICheckout_Part4():
    # //************************ ETHERENET SWITCHES ************************//

    dfSwitchConfig = pd.read_excel("./hmi_checkout/switch_config.xlsx", sheet_name="Sheet1")

    out_table = []
    port_list = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]
    x = 520

    for index, row in dfSwitchConfig.iterrows():
        switch_id = row["switch_id"]

        #print("Switch ID: " + switch_id)

        out_row = {}
        out_row["header"] = str(x) + " ENet Switch " + switch_id + " Comm Status"
        out_row["initial"] = " "
        out_table.append(out_row)

        out_row = {}
        out_row["header"] = "Title display correct"
        out_row["initial"] = " "
        out_table.append(out_row)

        out_row = {}
        out_row["header"] = "All GOTO buttons and fault indicators functional"
        out_row["initial"] = " "
        out_table.append(out_row)

        out_row = {}
        out_row["header"] = "Switch status name and status indicator functional"
        out_row["initial"] = " "
        out_table.append(out_row)

        for port in port_list:
            #print("Switch: " + switch_id + " | Port: " + str(port))
            out_row = {}
            out_row["header"] = "Port " + str(port) + " status name and status indicator functional"
            out_row["initial"] = " "
            out_table.append(out_row)

        out_row = {}
        out_row["header"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)
        
        x=x+1

    dfOutput = pd.DataFrame(out_table)
    dfOutput.to_excel("./hmi_checkout/output_part4.xlsx", header=False, index=False)

    return True

def CombineHMISheets(iterationNumber):
    print("Starting combine sheets at iteration number: " + str(iterationNumber))

    dfPart1 = pd.read_excel("./hmi_checkout/output_part1.xlsx", sheet_name="Sheet1", header=None, index=None)
    dfPart2 = pd.read_excel("./hmi_checkout/output_part2.xlsx", sheet_name="Sheet1", header=None, index=None)
    dfPart3 = pd.read_excel("./hmi_checkout/output_part3.xlsx", sheet_name="Sheet1", header=None, index=None)
    dfPart4 = pd.read_excel("./hmi_checkout/output_part4.xlsx", sheet_name="Sheet1", header=None, index=None)

    combined = [dfPart1, dfPart2, dfPart3, dfPart4]

    final = pd.concat(combined)

    final.to_excel("./hmi_checkout/HMI" + str(iterationNumber) + "_Checkout.xlsx", header=False, index=False)

def cleanup():
    print("cleaning up...")
    os.remove("./hmi_checkout/output_part1.xlsx")
    os.remove("./hmi_checkout/output_part2.xlsx")
    os.remove("./hmi_checkout/output_part3.xlsx")
    os.remove("./hmi_checkout/output_part4.xlsx")
    time.sleep(3)
    print("\nfinished. returning to main screen...")
    time.sleep(3)
    main()

def generate_pdpio():
    out_table = []

    #480V
    #CB0107 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0107"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0107 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0201 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0201"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0201 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0204 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0204"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0204 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0207 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0207"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0207 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0210 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0210"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0210 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0213 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0213"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0213 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0216 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0216"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0216 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0219 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0219"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0219 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0231 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0231"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0231 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0234 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0234"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0234 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0237 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0237"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0237 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0240 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0240"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0240 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0243 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0243"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0243 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0246 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0246"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0246 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0120 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0120"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0120 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0118 ROW
    out_row = {}
    out_row["address"] = "DI0730"
    out_row["device_no"] = "CB0118"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "480V CB0118 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #120V
    #CB0408 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0408"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0408 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0410 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0410"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0410 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0412 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0412"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0412 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0414 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0414"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0414 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0416 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0416"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0416 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0418 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0418"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0418 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0432 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0432"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0432 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0434 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0434"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0434 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0436 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0436"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0436 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0508 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0508"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0508 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0510 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0510"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0510 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0512 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0512"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0512 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0514 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0514"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0514 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0516 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0516"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0516 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0518 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0518"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0518 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #CB0532 ROW
    out_row = {}
    out_row["address"] = "DI0800"
    out_row["device_no"] = "CB0532"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "120V CB0532 OK"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    #EMPTY ROW
    out_row = {}
    out_row["address"] = " "
    out_row["device_no"] = " "
    out_row["no"] = " "
    out_row["nc"] = " "
    out_row["actuated_by"] = " "
    out_row["function"] = " "
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    dfOutput = pd.DataFrame(out_table)
    dfOutput.to_excel('./io_checkout/1_pdpio_output.xlsx', header=False, index=False)

def generate_sbkio(sbkNumber, deviceControl, iterationNumber):
    out_table = []

    if("TS" in deviceControl):
        #SBKIO for TrackSwitch
        out_row = {}
        out_row["address"] = sbkNumber + ".I.b00"
        out_row["device_no"] = "PRST"
        out_row["no"] = " - "
        out_row["nc"] = " - "
        out_row["actuated_by"] = " "
        out_row["function"] = "TRACKSWITCH IN THRU POSITION"
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

        out_row = {}
        out_row["address"] = sbkNumber + ".I.b01"
        out_row["device_no"] = "PRSD"
        out_row["no"] = " - "
        out_row["nc"] = " - "
        out_row["actuated_by"] = " "
        out_row["function"] = "TRACKSWITCH IN DIVERT POSITION"
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

        out_row = {}
        out_row["address"] = sbkNumber + ".I.b02"
        out_row["device_no"] = "PE17"
        out_row["no"] = " - "
        out_row["nc"] = " - "
        out_row["actuated_by"] = " "
        out_row["function"] = "TRACKSWITCH PROTECTION"
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

        out_row = {}
        out_row["address"] = sbkNumber + ".O.b00"
        out_row["device_no"] = "SOL1"
        out_row["no"] = " - "
        out_row["nc"] = " - "
        out_row["actuated_by"] = " "
        out_row["function"] = "TRACKSWITCH OPERATE THRU"
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

        out_row = {}
        out_row["address"] = sbkNumber + ".O.b01"
        out_row["device_no"] = "SOL1"
        out_row["no"] = " - "
        out_row["nc"] = " - "
        out_row["actuated_by"] = " "
        out_row["function"] = "TRACKSWITCH OPERATE DIVERT"
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

    #EMPTY ROW
    out_row = {}
    out_row["address"] = " "
    out_row["device_no"] = " "
    out_row["no"] = " "
    out_row["nc"] = " "
    out_row["actuated_by"] = " "
    out_row["function"] = " "
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    dfOutput = pd.DataFrame(out_table)
    dfOutput.to_excel('./io_checkout/2_sbkio_output.xlsx', header=False, index=False)
  
def generate_bkio(device_name, device_control, iterationNumber):
    dfConfigDevices = pd.read_excel("./io_checkout/config_devices.xlsx")

    out_table = []
    
    bk_portlist = [0, 1, 2 , 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]

    if("CHB" in device_control):
        #print("Device: " + device_name + " | Device Control: " + device_control)
        bk_portlist = [0, 1, 2, 3, 8, 9, 10, 11, 12, 13, 14, 15]
    
    for port in bk_portlist: 
        if(port < 10):
            port = str(port).zfill(2)

        if(int(port) <= 3):
            ##BKIO SPARE
            out_row = {}
            out_row["address"] = device_name + ".C.b" + str(port)
            out_row["device_no"] = " - "
            out_row["no"] = " - "
            out_row["nc"] = " - "
            out_row["actuated_by"] = " "
            out_row["function"] = "-"
            out_row["date_tested"] = " "
            out_row["initial"] = " "
            out_table.append(out_row)

    out_row = {}
    out_row["address"] = device_name + ".I.b04"
    out_row["device_no"] = device_control + " PB"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "HOLD RELEASED"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row["address"] = device_name + ".I.b05"
    out_row["device_no"] = device_control + " PB"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "HOLD PRESSED"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row["address"] = device_name + ".O.b06"
    out_row["device_no"] = device_control + " SP"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "SPARE OUTPUT"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    out_row = {}
    out_row["address"] = device_name + ".O.b07"
    out_row["device_no"] = device_control + " LT"
    out_row["no"] = " - "
    out_row["nc"] = " - "
    out_row["actuated_by"] = " "
    out_row["function"] = "HOLD STACK LIGHT"
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    for port in bk_portlist:
        if(port < 10):
            port = str(port).zfill(2)

        ##BKIO SPARE
        out_row = {}
        out_row["address"] = device_name + ".C.b" + str(port)
        out_row["device_no"] = " - "
        out_row["no"] = " - "
        out_row["nc"] = " - "
        out_row["actuated_by"] = " "
        out_row["function"] = "-"
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

    ##EMPTY ROW
    out_row = {}
    out_row["address"] = " "
    out_row["device_no"] = " "
    out_row["no"] = " "
    out_row["nc"] = " "
    out_row["actuated_by"] = " "
    out_row["function"] = " "
    out_row["date_tested"] = " "
    out_row["initial"] = " "
    out_table.append(out_row)

    dfOutput = pd.DataFrame(out_table)
    if(os.path.isdir("./io_checkout/BKIO")):
        dfOutput.to_excel('./io_checkout/BKIO/BK' + str(iterationNumber) + '_output.xlsx', header=False, index=False)
    else:
        os.mkdir("./io_checkout/BKIO")
        dfOutput.to_excel('./io_checkout/BKIO/BK' + str(iterationNumber) + '_output.xlsx', header=False, index=False)
        
def generate_ifdio():

    dfDriveConfig = pd.read_excel("./io_checkout/config_drives.xlsx")


    out_table = []
    ifd_spareports = [3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15] #configure based on number of spare ports for an IFD

    for index, row in dfDriveConfig.iterrows():
        drive_id = row["drive_id"]

        #IFDIO ENTERING
        out_row = {}
        out_row["address"] = "IFD" + str(drive_id) + ".I.b00"
        out_row["device_no"] = "IFD" + str(drive_id) + ".PRS1"
        out_row["no"] = " - "
        out_row["nc"] = " - "
        out_row["actuated_by"] = " "
        out_row["function"] = "ENTERING"
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

        #IFDIO IN POSITION
        out_row = {}
        out_row["address"] = "IFD" + str(drive_id) + ".I.b01"
        out_row["device_no"] = "IFD" + str(drive_id) + ".PE3"
        out_row["no"] = " - "
        out_row["nc"] = " - "
        out_row["actuated_by"] = " "
        out_row["function"] = "IN POSITION"
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

        #IFDIO CHASING
        out_row = {}
        out_row["address"] = "IFD" + str(drive_id) + ".I.b02"
        out_row["device_no"] = "IFD" + str(drive_id) + ".PRS4"
        out_row["no"] = " - "
        out_row["nc"] = " - "
        out_row["actuated_by"] = " "
        out_row["function"] = "CHASING"
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

        for port in ifd_spareports:
            #print(port)
            if(port < 10):
                port = str(port).zfill(2)
                #print(port)

            ##IFDIO SPARE
            out_row = {}
            out_row["address"] = "IFD" + str(drive_id) + ".I.b" + str(port)
            out_row["device_no"] = " - "
            out_row["no"] = " - "
            out_row["nc"] = " - "
            out_row["actuated_by"] = " "
            out_row["function"] = "-"
            out_row["date_tested"] = " "
            out_row["initial"] = " "
            out_table.append(out_row)

        #EMPTY ROW
        out_row = {}
        out_row["address"] = " "
        out_row["device_no"] = " "
        out_row["no"] = " "
        out_row["nc"] = " "
        out_row["actuated_by"] = " "
        out_row["function"] = " "
        out_row["date_tested"] = " "
        out_row["initial"] = " "
        out_table.append(out_row)

    dfOutput = pd.DataFrame(out_table)
    dfOutput.to_excel('./io_checkout/4_ifd_output.xlsx', header=False, index=False)

def combineBKSheets():
    print("Starting combine BK sheets")
    all_data = [] 
    
    for f in glob.glob("io_checkout/BKIO/*.xlsx"):
        all_data.append(pd.read_excel(f, header=None, index=None))

    df = pd.concat(all_data, ignore_index=True)
    df.to_excel("./io_checkout/3_bk_output.xlsx", header=False, index=False)

    print("cleaning up...")

    for f in glob.glob("io_checkout/BKIO/*.xlsx"):
        os.remove(f)
    
    os.rmdir("./io_checkout/BKIO")

def combineIOSheets():
    print("Starting combine IO sheets")
    all_data = []
    
    for f in glob.glob("io_checkout/*_output.xlsx"):
        all_data.append(pd.read_excel(f, header=None, index=None))

    df = pd.concat(all_data, ignore_index=True)
    df.to_excel("./io_checkout/IO_Checkout.xlsx", header=False, index=False)

    print("cleaning up...")

    for f in glob.glob("io_checkout/*_output.xlsx"):
        os.remove(f)

def seqValidationChecklist():
    fileExists = os.path.isfile('./sequence_checkout/config.xlsx')
    if(fileExists):
        print("Config file exist...continuing")
    else:
        print("Config file missing...returning to main")
        time.sleep(3)
        main()

    dfConfig = pd.read_excel("./sequence_checkout/config.xlsx", sheet_name="Sheet1")
    #print(dfConfig)

    out_table = []
    for index, row in dfConfig.iterrows():
        if dfConfig.isnull().values.any() or row["drive_id"] == "nan" or row["role"] == "nan":
            print("Detected invalid configuration in `config.xlsx`...please correct and run again")
            time.sleep(3)
            main()

        drive_id = row["drive_id"]
        role = row["role"]
        #print("Drive ID: " + drive_id + " | " + role)

        if row["role"] == "FOLLOWER":
            #FOLLOWER 1st row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Rotation"
            out_row["actuated_by"] = ""
            out_row["function"] = "Test for correct rotation direction"
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)
            #FOLLOWER 2nd row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Jog Fwd"
            out_row["actuated_by"] = ""
            out_row["function"] = "Jog drive forward from HMI"
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)
            #FOLLOWER 3rd row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Jog Rev"
            out_row["actuated_by"] = ""
            out_row["function"] = "Jog drive reverse from HMI"
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)
            #FOLLOWER EMPTY row seperator
            out_row = {}
            out_row["ifd_number"] = ""
            out_row["tested_action"] = ""
            out_row["actuated_by"] = ""
            out_row["function"] = ""
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)

        if row["role"] == "LEADER":
            #LEADER 1st row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Rotation"
            out_row["actuated_by"] = ""
            out_row["function"] = "Test for correct rotation direction"
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)
            #LEADER 2nd row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Jog Fwd"
            out_row["actuated_by"] = ""
            out_row["function"] = "Jog drive forward from HMI"
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)
            #LEADER 3rd row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Jog Rev"
            out_row["actuated_by"] = ""
            out_row["function"] = "Jog drive reverse from HMI"
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)
            #LEADER 4th row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Xfer In Man"
            out_row["actuated_by"] = ""
            out_row["function"] = "Transfer In in Manual Mode"
            out_row["date_tested"] = "N/A"
            out_row["initial"] = "N/A"
            out_table.append(out_row)
            #LEADER 5th row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Xfer Out Man"
            out_row["actuated_by"] = ""
            out_row["function"] = "Transfer Out in Manual Mode"
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)
            #LEADER 6th row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Xfer In Auto"
            out_row["actuated_by"] = ""
            out_row["function"] = "Transfer In in Auto Mode"
            out_row["date_tested"] = "N/A"
            out_row["initial"] = "N/A"
            out_table.append(out_row)
            #LEADER 7th row
            out_row = {}
            out_row["ifd_number"] = drive_id
            out_row["tested_action"] = "Xfer Out Auto"
            out_row["actuated_by"] = ""
            out_row["function"] = "Transfer Out in Auto Mode"
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)
            #LEADER EMPTY row seperator
            out_row = {}
            out_row["ifd_number"] = ""
            out_row["tested_action"] = ""
            out_row["actuated_by"] = ""
            out_row["function"] = ""
            out_row["date_tested"] = ""
            out_row["initial"] = ""
            out_table.append(out_row)

        
    dfOutput = pd.DataFrame(out_table)
    dfOutput.to_excel("./sequence_checkout/Sequence_Checkout.xlsx", header=False, index=False)

    print("Done!")
    time.sleep(3)
    main()

def HMI_Checklist():
    f1Exists = os.path.isfile('./hmi_checkout/drive_config.xlsx')
    f2Exists = os.path.isfile('./hmi_checkout/switch_config.xlsx')
    if(f1Exists and f2Exists):
        print("Configuration files exzist...continuing")
        time.sleep(2)
    else:
        print("Config file(s) missing...returning to main")
        time.sleep(3)
        main()
    
    os.system('cls')
    dfDriveConfig = pd.read_excel("./hmi_checkout/drive_config.xlsx", sheet_name="Sheet1")

    numHMIs = dfDriveConfig["hmi_control"].max()

    for x in range(1, numHMIs+1):
        if(HMICheckout_Part1()):
            if(HMICheckout_Part2(x)):
                if(HMICheckout_Part3()):
                    if(HMICheckout_Part4()):
                        CombineHMISheets(x)
    
    cleanup()

def IO_Checklist():
    os.system('cls')
    f1Exists = os.path.isfile('./io_checkout/config_devices.xlsx')
    f2Exists = os.path.isfile('./io_checkout/config_drives.xlsx')
    if(f1Exists and f2Exists):
        print("Config files exist...continuing")
    else:
        print("Config file(s) missing...returning to main")
        time.sleep(3)
        main()

    dfDeviceConfig = pd.read_excel('./io_checkout/config_devices.xlsx', sheet_name='Sheet1')
    dfDriveConfig = pd.read_excel('./io_checkout/config_drives.xlsx', sheet_name='Sheet1')

    out_table = []

    print("Generating. Please wait...");

    x,y=1,1
    for index, row in dfDeviceConfig.iterrows():
        #print("Iteration Number: " + str(x))
        device_type = row["device_type"]
        device_control = row["device_control"]
        device_name = row["device_name"]
        
        if(device_type == "PDP"):
            generate_pdpio()
        if(device_type == "SBK"):
            generate_sbkio(device_name, device_control, x)
            x=x+1
        if(device_type == "BK"):
            generate_bkio(device_name, device_control, y)
            y=y+1
            
        

    generate_ifdio()
    combineBKSheets()
    combineIOSheets()

def Safety_Checklist():
    os.system('cls')
    f1Exists = os.path.isfile('./safety_checkout/config.xlsx')
    if(f1Exists):
        print("Config file exists...continuing")
    else:
        print("Config file(s) missing...returning to main")
        time.sleep(3)
        main()


    dfConfig = pd.read_excel('./safety_checkout/config.xlsx', sheet_name="Sheet1")

    out_table = []
    device_list = []
    autozone_list = []

    print("Generating. Please wait...")

    for index, row in dfConfig.iterrows():
        autozone = row['autozone']

        #AUTOZONE
        out_row = {}
        out_row["device"] = ""
        out_row["input"] = "Auto Zone " + str(autozone)
        out_row["initials"] = ""
        out_row["date"] = ""
        out_table.append(out_row)

        #HEADER
        out_row = {}
        out_row["device"] = "UNIT/DEVICE"
        out_row["input"] = " "
        out_row["initials"] = "INITIALS"
        out_row["date"] = "DATE"
        out_table.append(out_row)

        for device in row:
            if device != "NaN":
                if "HMI" in str(device):
                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b00"
                    out_row["input"] = "E-STOP PB CHANNEL 1 INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b01"
                    out_row["input"] = "E-STOP PB CHANNEL 2 INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b02"
                    out_row["input"] = "SAFETY GATE OK CHANNEL 1 INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b03"
                    out_row["input"] = "SAFETY GATE OK CHANNEL 2 INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b04"
                    out_row["input"] = "RESET PB INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b05"
                    out_row["input"] = "AUTO INITIATE PB INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O3.b00"
                    out_row["input"] = "AUTO INITIATE PB INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O3.b01"
                    out_row["input"] = "BLOCKED FLASHING STARVED STEADY LIGHT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O3.b02"
                    out_row["input"] = "MANUAL FLASHING RUN/STOP STEADY LIGHT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O3.b03"
                    out_row["input"] = "HORN MODULE"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O3.b06"
                    out_row["input"] = "SYSTEM RESET PB LIGHT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O3.b07"
                    out_row["input"] = "AUTO INITIATE PB LIGHT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)
                if "SCB" in str(device):
                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b00"
                    out_row["input"] = "SAFETY CONTACTOR 1 OFF"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b01"
                    out_row["input"] = "SAFETY CONTACTOR 2 OFF"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b02"
                    out_row["input"] = "SAFETY CONTACTOR 3 OFF"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b03"
                    out_row["input"] = "SAFETY CONTACTOR 4 OFF"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b04"
                    out_row["input"] = "E-STOP PB CHANNEL 1 INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b05"
                    out_row["input"] = "E-STOP PB CHANNEL 2 INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b06"
                    out_row["input"] = "RESET PB INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_S.I1.b07"
                    out_row["input"] = "AUTO INITIATE PB INPUT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O2.b00"
                    out_row["input"] = "POWER CONTACTOR 1 ENABLE"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O2.b01"
                    out_row["input"] = "POWER CONTACTOR 2 ENABLE"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O2.b02"
                    out_row["input"] = "POWER CONTACTOR 3 ENABLE"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O2.b03"
                    out_row["input"] = "POWER CONTACTOR 4 ENABLE"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.I3.b00"
                    out_row["input"] = "DISCONNECT ON"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.I3.b01"
                    out_row["input"] = "AS1 POWER 102MCP ON"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.I3.b02"
                    out_row["input"] = "AS2 POWER 112MCP ON"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O4.b00"
                    out_row["input"] = "FAULT STACK LT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O4.b01"
                    out_row["input"] = "READY STACK LT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O4.b03"
                    out_row["input"] = "E-STOPPED STACK LT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O4.b04"
                    out_row["input"] = "RESET PB LT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = str(device) + "_P.O4.b05"
                    out_row["input"] = "START PB LT"
                    out_row["initials"] = " "
                    out_row["date"] = " "
                    out_table.append(out_row)
                if "GPB" in str(device):
                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row) #@TODO: UPDATE THIS WHEN WE FIGURE OUT WHAT ITS SUPPOSED TO BE
                if "RPB" in str(device):
                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row) #@TODO: UPDATE THIS WHEN WE FIGURE OUT WHAT ITS SUPPOSED TO BE
                if "EOT" in str(device):
                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row)

                    #DEVICE
                    out_row = {}
                    out_row["device"] = ""
                    out_row["input"] = str(device)
                    out_row["initials"] = ""
                    out_row["date"] = ""
                    out_table.append(out_row) #@TODO: UPDATE THIS WHEN WE FIGURE OUT WHAT ITS SUPPOSED TO BE


    dfOutput = pd.DataFrame(out_table)
    dfOutput.to_excel("./safety_checkout/Safety_Checkout.xlsx", header=False, index=False)    
    print("Done...returning to main")
    time.sleep(3)
    main()

def SparePartsList():
    os.system('cls')
    f1Exists = os.path.isfile('./spare_parts/config.xlsx')
    if(f1Exists):
        print("Config file exists...continuing")
    else:
        print("Config file(s) missing...returning to main")
        time.sleep(3)
        main()

    dfConfig = pd.read_excel('./spare_parts/config.xlsx', sheet_name="Sheet1")

main()