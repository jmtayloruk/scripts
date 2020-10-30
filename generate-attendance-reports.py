'''
    ####################################################################################
    ######## Script to generate attendance report from Zoom 'participants' logs ########
    ####################################################################################
    Do please drop me an email to let me know if you find this script useful (or have any additions to contribute):
    jonathan.taylor at glasgow.ac.uk
        -- Jonathan Taylor, Glasgow University, October 2020

    Version 2.0 updates: Semi-automated feature to merge personal and university email addresses. Generate report tables by day and by week.
    Version 3.0 update:  Command-line parameters

    Command line usage:
                One or more parameters giving paths to directory containing .csv files.
                    Each directory will be processed independently. The directory will be scanned for all files of the form
                    "participants*.csv" (downloaded from zoom - see below), and attendance reports will be generated in that same directory.
                    If no directories specified, script runs on current directory
                    (If you include *unquoted* wildcards in your command line paths, all matches will be processed independently)
                Can also include an optional parameter (e.g. -m2) specifying that students must attend a minimum of 2 sessions.
                    Warnings will then be generated for any students not meeting that threshold
                    (but watch out for students who may have signed in under a personal and a university email address,
                     if these have not been successfully paired together)

    Outputs:
    - a file 'meeting-report.csv' listing all the participants, and giving a chronological account of which meetings
      they have attended (and for how long). For now, this is just a list of all dates that the student has attended
      (and the total time they attended for).
    - a file 'meeting-report-by-day.csv' giving a table of participants vs dates attended.
    - a file 'meeting-report-by-week.csv' giving a table of participants vs weeks attended
      (helpful for labs where different students attend once each, on different days of the week)
    - warnings about low-attending students

    Limitations:
    - At the moment the output data is sorted by email address (which is not a very helpful ordering
      given that the students have numerical email addresses). However, email address seemed like the most
      reliable thing to use as a key, given that some students change their display name on Zoom.
    - Script currently assumes only one meeting per day, and will fuse into one entry if there was
      more than one separate 'participants' file referring to the same calendar date.
    - Signing in to zoom is no proof that the student is actually present or engaged at the computer at the other end!

    Customization (Glasgow University):
    - To merge records between students' personal and university email addresses,
      manually edit the code that initializes the dictionary 'emailMapping'.
      If you just run the script without this, the code will try and make suggestions about possible pairings it has guessed for you.
      You would then need to manually edit the emailMapping dictionary to add those pairings.
    - If it is hard to distinguish demonstrator emails from undergraduate student emails,
      you may need to edit knownDemonstratorEmails to manually identify demonstrator emails
    Additional customization (other universities):
    - Edit the function "StaffEmail()" to identify staff/demonstrator email addresses that should be excluded from the reports
    - Edit the function "UniversityStudentEmail()" to identify email addresses that are official university stuent email addresses
    
    To download the participants files:
    The meeting owner should go to "edit meeting" then Reports -> Usage and search for the relevant date range.
    This brings up a table of data about meetings. Click on the number (hyperlink) in the "Participants" column of the table,
    and that pops up a list of participants. Click "Export", then hit close/escape to go back to the main table.
    Repeat for the other meetings of interest, and then gather all the downloaded files into a suitable directory for processing.
'''

import numpy as np
import csv, datetime, glob, sys

def StaffEmail(email):
    # Returns True if this looks like a staff email, so this can be excluded from the attendance report.
    # If staff emails are indistinguishable in format from student emails, or you don't want to bother with this,
    # just return False from this function. In that case, the worst that will happen is you'll get attendance reports
    # for staff/demonstrators as well.
    # If you can't distinguish some or all demonstrator emails, but want to enter them manually, then
    # knownDemonstratorEmails can contain a manually-curated list of demonstrator emails 
    # that would be otherwise indistinguishable from undergraduate email addresses.
    knownDemonstratorEmails = []
    return ("@glasgow.ac.uk" in email) or           \
           ("@research.glasgow.ac.uk" in email) or  \
           ("@gla.ac.uk" in email) or               \
           ("@research.gla.ac.uk" in email) or      \
           (email in knownDemonstratorEmails)

def UniversityStudentEmail(email):
    # Returns True if this looks like an official university student email, as opposed to a personal email address.
    # If you can't specifically distinguish student email addresses from staff, just return True for all university email addresses.
    return "@student.gla.ac.uk" in email

def date_from_isoweek(iso_year, iso_weeknumber, iso_weekday):
    # From stackoverflow (https://stackoverflow.com/questions/304256/whats-the-best-way-to-find-the-inverse-of-datetime-isocalendar)
    return datetime.datetime.strptime(
        '{:04d} {:02d} {:d}'.format(iso_year, iso_weeknumber, iso_weekday),
        '%G %V %u').date()

# Manually curated list of email pairs for students who I have noticed switched from personal to GU emails.
# The cell at the end can be useful in suggesting matches to add here, but that relies on the student 
# giving themselves a correct and clear name to accompany their non-GU email address.
# e.g. add entries like:
#     "easyrider2001@hotmail.com": "1234567a@student.gla.ac.uk",
emailMapping = { }
reverseEmailMapping = {v: k for k, v in emailMapping.items()}

directoriesToProcess = []
warningThreshold = 0
processDefaultDirectory = True
for arg in sys.argv[1:]:
    if arg.startswith("-m"):
        if (len(arg) == 2):
            print("Usage: \"-m4\" to warn for students who have attended <=4 sessions")
        else:
            print("Will warn for students who have attended <={0} sessions".format(arg[2:]))
            warningThreshold = int(arg[2:])
    else:
        processDefaultDirectory = False
        if "*" in arg:
            print("Warning: ignoring quoted wildcard was passed in as a command line parameter \"{0}\".".format(arg))
            print("  That approach is not supported - use an unquoted wildcard if you want to process a batch of directories independently")
        else:
            directoriesToProcess.append(arg)

if processDefaultDirectory:
    print("No directories specified in command line arguments specified - processing current directory")
    directoriesToProcess = ["."]


for sourcePath in directoriesToProcess:
    print("\n===== Processing directory \"{0}\" =====".format(sourcePath))
    
    ##########################################
    ### Load all available meeting records ###
    ##########################################

    # Dictionary to accumulate all our data
    emails = dict()

    # List of input files to process
    filenames = glob.glob("{0}/participants*.csv".format(sourcePath))

    for fn in filenames:
        # First a quick peek to check the date in the file
        with open(fn, newline='') as csvfile:
            csvreader = csv.reader(csvfile, delimiter=',')
            row1 = next(csvreader)
            row2 = next(csvreader)
            print('Processing file {0} (date {1})'.format(fn, row2[2]))
        # Now process the file properly
        with open(fn, newline='') as csvfile:
            csvreader = csv.reader(csvfile, delimiter=',')
            for row in csvreader:
                if row[0].endswith('Name (Original Name)'):
                    # Skip header row.
                    # Note that the use of 'endswidth' avoids problems caused by a weird
                    # unicode character that Zoom puts at the very start of the .csv files it generates.
                    continue
                    
                # Parse data row
                name = row[0]
                email = row[1].lower()  # Convert to lowercase because some students seemed to change that mid-semester
                if email in emailMapping:
                    emailKey = emailMapping[email]
                else:
                    emailKey = email
                start = row[2]
                mins = int(row[4])
                date_time_obj = datetime.datetime.strptime(start, '%d/%m/%Y %H:%M:%S %p')
                date = date_time_obj.date()

                # Create an entry if we have not encountered this student before
                if not emailKey in emails:
                    emails[emailKey] = dict()

                if date in emails[emailKey]:
                    # We already have an entry for this student on this date.
                    # Add the number of minutes from the current data line we have just read
                    emails[emailKey][date][3] += mins
                else:
                    # Create a new entry for this student on this date
                    emails[emailKey][date] = [name, email, start, mins]


    ################################################################################################
    ### Useful utility routine to spot who we have failed to match up to a Glasgow email address ###
    ################################################################################################
    for email in sorted(emails):
        entry = emails[email]

        if (not StaffEmail(email)) and (not UniversityStudentEmail(email)):
            firstEntry = next(iter(entry.values()))
            print("NOTE: student {0}, {1} not matched to university email address".format(firstEntry[0], firstEntry[1]))
            # See if we can find a match for the surname in an entry that *does* have a GU email address
            possibleSurname = firstEntry[0].split(' ')[-1]
            for email2 in sorted(emails):
                if UniversityStudentEmail(email2):
                    for date in emails[email2]:
                        thisEntry = emails[email2][date]
                        thisName = thisEntry[0]
                        if possibleSurname in thisName:
                            print(" Might match to {0}, {1}?".format(thisEntry[0], thisEntry[1]))
                            print(" If so, manually add table row \"{0}\": \"{1}\",".format(firstEntry[1], thisEntry[1]))
                            break

    #####################################################
    ### Generate the attendance list for all students ###
    #####################################################
    # Also generates warnings about low-attending students
    # who have attended <= the specified minimum number of sessions.
    outputStudentAttendanceOnly = True

    with open("{0}/meeting-report.csv".format(sourcePath), mode='w') as csvOutput:
        csvwriter = csv.writer(csvOutput, delimiter=',')
        for email in sorted(emails):
            if outputStudentAttendanceOnly and StaffEmail(email):
                continue
                
            entry = emails[email]

            # Write out data to meeting report
            for date in sorted(entry):
                row = entry[date]
                #print('{0}, {1}, {2}, {3}'.format(row[0], row[1], row[2], row[3]))
                csvwriter.writerow(row)
            # Monitor low-attending students
            if (not StaffEmail(email)) and (len(entry) <= warningThreshold):
                print("WARNING: student {0} {1} only attended {2} sessions".format(row[0], row[1], len(entry)))

    ########################################################
    ### Generate table of student name vs dates attended ###
    ########################################################

    # First identify all the meeting dates
    dateCatalogue = dict()
    for email in sorted(emails):
        entry = emails[email]
        for date in entry:
            if not date in dateCatalogue:
                dateCatalogue[date] = date

    # Now build up our table
    with open("{0}/meeting-report-by-date.csv".format(sourcePath), mode='w') as csvOutput:
        csvwriter = csv.writer(csvOutput, delimiter=',')
        outputRow = ["Name", "Email"]
        for date in sorted(dateCatalogue):
            outputRow.append(date)
        csvwriter.writerow(outputRow)
        
        for email in sorted(emails):
            if outputStudentAttendanceOnly and StaffEmail(email):
                continue
                
            studentRecord = emails[email]
            firstEntry = next(iter(studentRecord.values()))
            name = firstEntry[0]

            # Write out row to meeting report
            outputRow = [name, email]
            for date in sorted(dateCatalogue):
                if date in studentRecord:
                    row = studentRecord[date]
                    outputRow.append(row[3])
                else:
                    outputRow.append("")
            csvwriter.writerow(outputRow)

    ########################################################
    ### Generate table of student name vs weeks attended ###
    ########################################################

    # First identify all the meeting weeks
    weekCatalogue = dict()
    for email in sorted(emails):
        entry = emails[email]
        for date in entry:
            week = date.isocalendar()[1]
            if not week in weekCatalogue:
                weekStart = date_from_isoweek(date.isocalendar()[0],
                                              date.isocalendar()[1],
                                              1)
                weekCatalogue[week] = weekStart

    # Now build up our table
    with open("{0}/meeting-report-by-week.csv".format(sourcePath), mode='w') as csvOutput:
        csvwriter = csv.writer(csvOutput, delimiter=',')
        outputRow = ["Name", "Email"]
        for date in sorted(weekCatalogue):
            outputRow.append(weekCatalogue[date])
        csvwriter.writerow(outputRow)
        
        for email in sorted(emails):
            if outputStudentAttendanceOnly and StaffEmail(email):
                continue
                
            studentRecord = emails[email]
            firstEntry = next(iter(studentRecord.values()))
            name = firstEntry[0]

            # Write out row to meeting report
            outputRow = [name, email]
            for week in sorted(weekCatalogue):
                attendanceSum = 0
                for date in studentRecord:
                    studentEntryWeek = date.isocalendar()[1]
                    if (studentEntryWeek == week):
                        attendanceSum += studentRecord[date][3]
                if (attendanceSum > 0):
                    outputRow.append(attendanceSum)
                else:
                    outputRow.append("")
            csvwriter.writerow(outputRow)
