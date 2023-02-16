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


if (not len(infile) or not len(outfile)):
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
    print("read {} attendees from input sheet:".format(len(attendees)))
    for i in attendees:
        print("fname: ", i.fname)
        print("lname: ", i.lname)
        print("year: ", i.year, type(i.year))
    return attendees

class Event:
    def __init__(self, name, col):
        self.name = name
        self.col = col

# i'm writing this at 11 PM, it needs to be done by tomorrow, get ready for the worst search function you've ever seen in your life
def write_file():
    import shutil
    shutil.copyfile(outfile, outfile+".bak")
    print("Created backup of Excel File as {}.bak\n".format(outfile))
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
                if labels[i].value is not None:
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
                sheet.insert_cols(events[-1].col+2)
                event = Event(input("Please enter a new event name: "),events[-1].col+1)
                labels = tuple(sheet.rows)[2]
                labels[events[-1].col].value = events[-1].name
        except(ValueError, IndexError):
            print("Bad input! Please enter a number!")
            write_file()
            return
    else:
        "No existing events were detected."
        event = Event(input("Please enter a new event name: "),memberindex+1)
        sheet.insert_cols(memberindex+1)
        labels = tuple(sheet.rows)[2]
        labels[memberindex+1].value = event[-1].name
    for row in sheet.iter_rows(min_row=4, values_only=False):
        name = row[memberindex].value.lower()
        year = int(row[yrindex].value)
        for person in people:
            if name == (person.fname.lower() + " " + person.lname.lower()) and year == person.year:
                print("debug: processing person",person.fname)
                if row[event.col].value is None:
                    row[event.col].value = 1
                else:
                    row[event.col].value += 1
                people.remove(person)
                # Recalculate column indexes
                labels = tuple(sheet.rows)[2]
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
                if internalCalc:
                    tally = 0
                    for ev in events+[event]:
                        try:
                            tally += int(row[ev.col].value)
                        except(ValueError):
                            print("Internal Calulation failed. Not updating counts for {}".format(person.fname))
                            break
                    row[termattendanceindex].value = tally
                else:
                    formulae = []
                    for cell in row:
                        if cell.value.count("="):
                            formulae.append(cell.column)
                    if len(formulae) == 0:
                        print("did not detect any formulas, skipping update process")
                        break
                    oldcol = ""
                    newcol = ""
                    for chara in row[event.col-1].coordinate:
                        if chara.isalpha():
                            oldcol += chara
                    for chara in row[event.col].coordinate:
                        if chara.isalpha():
                            newcol += chara
                    for formula in formulae:
                        formula = formula.replace(oldcol, newcol)
                break
            else:
                print("debug: {} did not match {}".format((person.fname + " " + person.lname), name))
    if len(people):
        print("The following {} attendees were listed on the form but not present in the main sheet:")
        for i in people:
            print("* ", i.fname, i.lname)
        if input("Add attendees to main sheet? (Y/n)").lower() != 'n':
            for i in people:
                personName = i.fname + " " + i.lname
                sheet.append({yrindex+1: i.year, memberindex+1: personName, termattendanceindex+1: 1, events[-1].col+1: 1})
    wb.save(outfile)
                
write_file()
print("done")
