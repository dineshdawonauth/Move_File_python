
# ----------------------------------------------------------------------------------------------------
# Imports

import os
import sys
import shutil
import pandas as pd
from termcolor import colored
import time

# ----------------------------------------------------------------------------------------------------
# Source, Destination Directories & Global Variables

Folder_Name = input("Enter Destination Folder > ")

Source = "/Users/dineshdawonauth/Desktop/" + Folder_Name + "/"
Destination = "/Users/dineshdawonauth/Desktop/Marked"

print(colored('\n' + 'Your source folder is {}'.format(Folder_Name) + '\n', 'blue'))

try:
    Files = sorted(os.listdir(Source))
except Exception as err:
    print("[-] Folder doens't exist!")

# Create the destination folder
try:
    os.mkdir(Destination)
except OSError:
    print(colored('\n' + "[-] Destination Directory creation FAILED ", 'red'))
    print("-" * 100 + "\n")
else:
    print(
        colored('\n' + "[+] Destination Directory creation SUCCESS ", 'blue'))
    print("-" * 100 + "\n")

# ----------------------------------------------------------------------------------------------------
# Get name list from excel
try:
    names_col = pd.read_excel("/Users/dineshdawonauth/Desktop/names.xlsx")

    Firstname_pd = (names_col['First name'])
    Surnames_pd = (names_col['Surname'])

    FirstName = Firstname_pd.to_numpy()
    Surnames = Surnames_pd.to_numpy()

    # ----------------------------------------------------------------------------------------------------
    # Combine names

    jointName = []

    for indiv in FirstName:
        jointName.append("-" + indiv)

    FullName = [b + a for a, b in zip(jointName, Surnames)]

    # ----------------------------------------------------------------------------------------------------
    # Algorithm to return matched students

    namestripped = []
    miscstripped = []
    matches = []
    finalMatch = []

    for file in Files:
        index = file.index("_")
        namestripped.append(file[:index])
        miscstripped.append(file[index:])

    for i in range(len(FullName)):
        for name in namestripped:
            if FullName[i] == name:
                matches.append(name)

    for j in range(len(Files)):
        for name in matches:
            if name in Files[j]:
                finalMatch.append(Files[j])

    # ----------------------------------------------------------------------------------------------------
    # Move files

    def movefile():

        for i in range(len(Files)):
            for match in finalMatch:
                if match == Files[i]:
                    index = match.index("_")
                    sourceN = Source + str(match)
                    destinationN = Destination + "/" + str(match)
                    print("Source for {} is {} ".format(
                        match[:index], sourceN))
                    print("Destination for {} is {} ".format(
                        match[:index], destinationN))

                    try:
                        shutil.move(sourceN, Destination)
                        print(
                            colored('\n' + "[+] {} moved Successfully".format(match), 'blue'))
                        print("-" * 100 + "\n")
                    except OSError as err:
                        print(
                            colored('\n' + "[-] Operation FAILED. Error is {}".format(err), 'red'))
                        print("-" * 100 + "\n")
                    time.sleep(1)


# ----------------------------------------------------------------------------------------------------
# Exception

except FileNotFoundError:
    print(colored("[-] File not found", 'red'))

# ----------------------------------------------------------------------------------------------------
# Run
try:
    movefile()
    os.system("figlet \"Move Successful  ! \" ")
except NameError:
    print(colored("[-] Move Failed", 'red'))

# ----------------------------------------------------------------------------------------------------
