# ================================================
# |  Just Another Certificate Generator (JACG)   |
# |  v1.3 (Specialised for KGSO '24)             |
# |  Made by def-vex                             |
# ================================================

# Code by : def-vex
# Comments by : def-vex
# Optimised by : def-vex (Hopefully at some point)

# Libraries needed are stated in requirements.txt
# To run just change up the code and execute by python3.11 certmaker.py

from PIL import Image # PIL = Pillow (for image-generation and manipulation)
from PIL import ImageDraw
from PIL import ImageFont
import os # Checking and making directories
import datetime # To get time-data
import pandas # To read our spreadsheet
import shutil # Moving files

# Make master folder
if not os.path.exists(r"D:/KGSO Certs"):
    os.mkdir(r"D:/KGSO Certs")

os.chdir(r"D:/KGSO Certs") # Change working directory to our newly-made (or existing) one

# Start time
start = datetime.datetime.utcnow()

# Excel sheet data made global
global excelData

excelData = pandas.read_excel("KGSO'24 Mastersheet"+".xlsx", sheet_name="Externals")

# Get column data from the excel sheet (through excelData using pandas)
def getData(colName:str):
    dataList = excelData[colName].tolist()
    return dataList

# Converting every data variable to global
global modules, names, teams, runNo, winNo, runMods, winMods, overall, overallPos

# Module keywords converted to be converted to their names
modules = {
    "physics":"Particle Paradox (Physics)",
    "crime":"Daedalus' Downfall (Crime)",
    "biochem":"Synaptic Fusion (Bio-Chemistry)",
    "psych":"Freudian Slipknot (Psychology)",
    "fandom":"Kamaji's Realm (Fandom)",
    "gaming":"Brass Battlegrounds (Gaming)",
    "cs":"Lovelace Interface (CS)",
    "astronomy":"Celestia Mechanica (Astronomy)",
    "logic":"Riddler's Realm (Logic)",
    "math":"The Pi Lie (Math)"
}

# All data variables
names = getData("Name")
teams = getData("Team")
runNo = getData("Runner Number")
winNo = getData("Win Number")
runMods = getData("Runner Modules")
winMods = getData("Winner Modules")
overall = getData("Overall")
overallPos = getData("Overall Position")

# Master dictionary for teams
# {
#   "E-001":"Winner_Azzaam_Crime.png", ...,
#   ...
# }
teamDict = {}
for team in teams:
    teamDict[team] = []

global certNum
certNum = 1

# Make master path for all certificates
if os.path.exists(r"D:/KGSO Certs/Certificates/") == False:
    os.mkdir(r"D:/KGSO Certs/Certificates/")

# How long a text is (used for centering purposes)
def chrLen(font, msg:str):
    w = font.getlength(msg)
    return w

# Compress name if too long (Convert parts of names to initials)
# Example : Syed Muhammad ... = S. M. ...
def compress(name:str):
    name = name.split(" ")
    return # UNFINISHED
    
# Master certificate generator function
def certGen(x, runs, wins, part):

    global modules, names, teams, runNo, winNo, runMods, winMods, overall, overallPos, certNum
    
    runYes, winYes = False, False

    done = False # Done will end the loop running in the Master loop for winners (Won't need to with not-winners)
    recentCert = "" # Naming the output png
    msg = "" # What is being printed on the certificate (Initialising)

    fontUse = ImageFont.truetype(os.environ['LOCALAPPDATA'] + "/Microsoft/Windows/Fonts/Fraunces_144pt-Regular.ttf", 25)
    W = 1200 # Width of the certificate template

    if part == False: # Participant certificate or not

        # Module Runner-up Certificate (Order 1)
        if len(runs) != 0:
            templ = Image.open('runnersup.png')
            imageGen = ImageDraw.Draw(templ)
            runYes = True
            modName = modules[runs[0].lower().strip(" ")] # Get module name from keyword
            w = chrLen(fontUse, msg)
            imageGen.text(((W-w)/2, 785/(5/3)), modName, (255, 0, 0), font=fontUse)
            recentCert = runs.pop(0).strip(" ") # Removes one data value from the runs list till it's empty

        # Module Winner Certificate (Order 2)
        if len(wins) != 0 and runYes == False:
            templ = Image.open('winners.png')
            imageGen = ImageDraw.Draw(templ)
            winYes = True
            modName = modules[wins[0].lower().strip(" ")]
            w = chrLen(fontUse, msg)
            imageGen.text(((W-w)/2, 785/(5/3)), modName, (255, 0, 0), font=fontUse)
            recentCert = wins.pop(0).strip(" ") # Removes one data value from the wins list till it's empty

        if overall[x].lower() == "no" and runYes == False and winYes == False:
            templ = Image.open('participation.png')
            imageGen = ImageDraw.Draw(templ)
            done = True

        # Overall Certificate (Order 3)
        if overall[x].lower() == "yes" and winYes == False and runYes == False:
            templ = Image.open('overall.png')
            imageGen = ImageDraw.Draw(templ)
            w = chrLen(fontUse, msg)
            imageGen.text(((W-w)/2, 785/(5/3)), overallPos[x], (255, 0, 0), font=fontUse)
            recentCert = overallPos[x]
            done = True
        
    if recentCert != "" or won == False or part == True: # won variable is redundant (might remove if not lazy)
        if part == True:
            templ = Image.open('participation.png') 
            imageGen = ImageDraw.Draw(templ)
        fontUse = ImageFont.truetype(os.environ['LOCALAPPDATA'] + "/Microsoft/Windows/Fonts/Fraunces_144pt-Regular.ttf", 40)
        
        # Name and team printed on the image
        msg = names[x].title() + "  -  " + teams[x]
        w = chrLen(fontUse, msg)
        imageGen.text(((W-w)/2, 605/(5/3)), msg, (0, 0, 0), font=fontUse)

        msg = "" # Re-init (probably not necessary)
        
        # Making the image and labelling it
        pngName = "{}_{}_{}.png".format(teams[x].strip(" "), names[x].title().strip(" "), recentCert)
        
        # Saving output certificate png
        templ.save(r"D:/KGSO Certs/Certificates/"+pngName)
        certNum += 1 # Counts number of total certificates made (for statistical purposes)
        teamDict[teams[x]].append(pngName)
    else:
        pass

    return done, runs, wins # Returns if done and/or number of runs and wins left to account for (certs wise)

# Iterating over each name (Master Loop)
for x in range(len(names)):
    
    print(names[x].title().strip(" "), certNum)
    
    done = False
    won = False

    try:
        runs = runMods[x].split(",") # Splitting runMods (Example : (Crime,CS) = ["Crime", "CS"])
        won = True # Again, this variable is confusing me (probably will remove)
    except:
        pass
    try:
        wins = winMods[x].split(",") # Splitting winMods (Example : (Crime,CS) = ["Crime", "CS"])
        won = True # REDUNDANT
    except: 
        pass
    if overall[x] == "Yes":
        won = True # REDUNDANT
    if won == False:
        done = certGen(x, runs, wins, True) # Participant Certificate for not-winners
    else:
        while done != True:
            done, runs, wins = certGen(x, runs, wins, False) # Winner certificates
        done, runs, wins = certGen(x, [], [], True) # Participant Certificate for winners
        

# Iterating over each team in teams dictionary
for x in teamDict:
    if x.startswith("E"):
        teamType = "External"
    else:
        teamType = "Internal"

    # Make parent folders (External/Internal)
    if os.path.exists(r"D:/KGSO Certs/Certificates/{}".format(teamType)) == False:
        os.mkdir(r"D:/KGSO Certs/Certificates/{}".format(teamType))

    # Make folder for each team
    if os.path.exists(r"D:/KGSO Certs/Certificates/{}/{}".format(teamType, x)) == False:
        os.mkdir(r"D:/KGSO Certs/Certificates/{}/{}".format(teamType, x))
    
    for y in teamDict[x]:
        
        name = y.split("_")[1] # Get name from dict value

        # Make folder for each person in the team
        if os.path.exists(r"D:/KGSO Certs/Certificates/{}/{}/{}".format(teamType, x, name)) == False:
            os.mkdir(r"D:/KGSO Certs/Certificates/{}/{}/{}".format(teamType, x, name))

        shutil.move(r"D:/KGSO Certs/Certificates/{}".format(y), r"D:/KGSO Certs/Certificates/{}/{}/{}/{}".format(teamType, x, name, y))

end = datetime.datetime.utcnow()
takenTime = end - start # Total cert-generating time taken
perCert = takenTime / certNum

file = open("Stats.txt", "a") # Writing to stats.txt for statistical purposes (obv)

file.write(f"""\n\nTime taken: {takenTime}
Per Generation: {perCert}
Number of Generations: {certNum}""")

file.close()