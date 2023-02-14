#!/usr/bin/python3
import os
import sys
termwidth = os.get_terminal_size()[0]
print("Welcome to Sukriti's Attendance Synchronization System (SASS)!".center(termwidth, ' '))

infile = ""
outfile = ""

for i in range (0,len(sys.argv)):
    match sys.argv[i]:
        case "-i":
            infile = sys.argv[i+1]
        case "-o":
            outfile = sys.argv[i+1]


if not interactive and (not len(infile) or not len(outfile)):
    print("\n Bad command invocation. Printing Help text...\n")
    print("Usage:\npython3 sass.py -i [input.xlsx] -o [output.xlsx]")
    print("Options:")
    print("-i                   specify input excel/google sheets file")
    print("-o                   specify target / output excel file")
    exit()
    
try:
    import openpyxl as op
except(ModuleNotFoundError):
    print("Excel support library not installed")
    if input("Install it now? (y/N)").lower() == "y":
        os.system("pip3 install openpyxl")
    else:
        print("Fatal error! Please install openpyxl manually.")
        exit()

print("The output sheet may have formulas that do not account for this program's changes, although I'll try my best to update them. To get around this, I can remove the formulas and write in my own calulated values.")
internalCalc = (input("Override Excel Formulas? (y/N) ").lower() == "y")

class Attendee:
    fname: str
    lname: str
    year: int
    email: str
    time: str

def read_file():
    attendees = []
    wb = op.load_workbook(filename=infile)
    sheet = wb.active
    labels = tuple(sheet.rows)[0]
    # assign default positions of data
    tsindex = 0
    fnindex = 1
    lnindex = 2
    yrindex = 3
    emindex = 4
    # auto-assign indexes based on first row
    for lbi in range(0,len(labels)):
        match labels[lbi].value.lower():
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
    return attendees

class Event:
    def __init__(self, name, col):
        self.name = name
        self.col = col

# i'm writing this at 11 PM, it needs to be done by tomorrow, get ready for the worst search function you've ever seen in your life
def write_file():
    import shutil
    shutil.copyfile(outfile, outfile+".bak")
    print("Created backup of Excel File\n")
    wb = op.load_workbook(filename=outfile)
    sheet = wb.active
    labels = tuple(sheet.rows)[2]
    events = []
    people = read_file()
    event = Event("default",0)
    yrindex = 0
    memberindex = 1
    termattendanceindex = 14
    activityindex = 15
    for i in range(0, len(labels)):
        match labels[i].value:
            case "Grad Year":
                yrindex = i
            case "Member":
                memberindex = i
            case "":
                pass
            case "Term attendance":
                termattendanceindex = i
            case "Activity": # may not be correct
                activityindex = i
            case _:
                events.append(Event(labels[i].value, i))
    if len(events):
        print("I found the following Events:".center(termwidth))
        for i in range(0, len(events)):
            print("{}) {}".format(i+1,events[i].name))
        print("\n0) Add new Event")
        try:
            choice = int(input("\n Please select an option: "))
            if choice:
                event = events[choice-1]
            else:
                sheet.insert_cols(events[-1].col+1)
                event = Event(input("Please enter a new event name: "),events[-1].col+1)
        except(ValueError, IndexError):
            print("Bad input! Please enter a number!")
            write_file()
            return
    else:
        "No existing events were detected."
        event = Event(input("Please enter a new event name: "),memberindex+1)
        sheet.insert_cols(memberindex+1)
    for row in sheet.iter_rows():
        updated = False
        name = row[memberindex].value.lower()
        year = int(row[yearindex].value)
        for person in people:
            if name == person.fname + " " + person.lname and year == person.year:
                if row[event.col] is None:
                    row[event.col] = 1
                else:
                    row[event.col] += 1
                people.remove(person)
                updated = True
                break
        if updated:
            # TODO - add formula updating / internal calc system
            pass
    wb.save(filename)
            
                
