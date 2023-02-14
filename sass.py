#!/usr/bin/python3
import os
import sys
termwidth = os.get_terminal_size()[0]
print("Welcome to Sukriti's Attendance Synchronization System (SASS)!".center(termwidth, ' '))

infile = ""
outfile = ""
interactive = False
verbosity = 0

for i in range (0,len(sys.argv)):
    match sys.argv[i]:
        case "-i":
            infile = sys.argv[i+1]
        case "-o":
            outfile = sys.argv[i+1]
        case "-v":
            verbosity = 1
        case "--interactive" | "-I":
            interactive = True

if not interactive and (not len(infile) or not len(outfile)):
    print("\n Bad command invocation. Printing Help text...\n")
    print("Usage:\npython3 sass.py -i [input.xlsx] -o [output.xlsx]\n-or-")
    print("python3 sass.py --interactive\n")
    print("Options:")
    print("-i                   specify input excel/google sheets file")
    print("-o                   specify target / output excel file")
    print("-I, --interactive    start in interactive mode")
    print("-v                   increase verbosity")

try:
    import openpyxl as op
except(ModuleNotFoundError):
    print("Excel support library not installed")
    if input("Install it now? (y/N)").lower() == "y":
        os.system("pip3 install openpyxl")
    else:
        print("Fatal error! Please install openpyxl manually.")
        exit()

class Attendee:
    fname: str
    lname: str
    year: int
    email: str
    time: str

def read_file():
    attendees = []
    workbookin = op.load_workbook(filename=infile)
    sheet = wb.active
    labels = sheet.rows[0]
    # assign default positions of data
    tsindex = 0
    fnindex = 1
    lnindex = 2
    yrindex = 3
    emindex = 4
    # auto-assign indexes based on first row
    for lbi in range(0,len(labels)):
        match labels[lbi].lower():
            case "timestamp":
                tsindex = lbi
            case "first name":
                fnindex = lbi
            case "last name":
                lnindex = lbi
            case "grad year":
                yrindex = lbi
            case "wpi email":
                emindex = lbi

    for row in sheet.iter_rows(min_row=2, values_only=True):
        attendees.append(Attendee())
        attendees[-1].fname = row[fnindex]
        attendees[-1].lname = row[lnindex]
        attendees[-1].time = row[tsindex]
        attendees[-1].year = int(row[yrindex])
        attendees[-1].email = row[emindex]

