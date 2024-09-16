import os  # to check and rename files
import pandas
import openpyxl


def setup():
    # default name of the file downloaded from https://nwis.waterdata.usgs.gov/usa/nwis/peak.
    if os.path.isfile("./peak"):
        os.rename('peak', 'peak.txt')  # We add the .txt extension to tell the program it is a text file.
        print("Changed file type to text file.")

    if os.path.isfile("./peak.txt"):
        print("File Found.")
        return 0
    else:
        print("File not found. Please make sure the file is named 'peak.txt' and is in the correct directory.")
        exit()


def removeComments(fileName, n):
    with open(fileName, 'r') as file:
        lines = file.readlines()

    with open(fileName, 'w') as file:
        file.writelines(lines[n:])

    return 0


def main():
    # Setup
    fileFound = setup()
    if fileFound == 0:
        print("Setup Complete.")
    else:
        print("Error 1")
        return 0

    # Find Number of Lines to Remove
    fileName = "peak.txt"
    n = 0

    with open(fileName, 'r') as file:
        for currentLine in file:
            if currentLine.startswith("#"):
                n += 1

    # Remove the lines that start with '#'
    if removeComments(fileName, n) == 0:
        print("Comments Removed.")
    else:
        print("Error 2")
        exit()

    print("Ready to convert .txt file to .xlsx")

    # Convert .txt to .xlsx
    df = pandas.read_csv("peak.txt", delimiter='\t')

    df.to_excel("peak.xlsx", index=False)

    workbook = openpyxl.load_workbook("peak.xlsx")
    worksheet = workbook["Sheet1"]
    worksheet.delete_rows(2)
    workbook.save("peak.xlsx")
    print("Removed Header 2.")

    return 0


main()
