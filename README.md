# PeerEvalScript
This is a python script that calculates an average peer evaluation score for each student in a course, using student survey results from Qualtrics.  This script was created for CIS 115 at Kansas State University.

### To use:
You will need two Excel workbooks to give to the script as input.  The first contains a list of all the students in the class and the second contains the results from the Qualtrics peer eval survey.  Make sure that both of these workbooks have the data stored on Sheet 0 of the workbook. Save these Excel files in the same directory as the script.  If you have the results as a CSV, you will need to open the file in Excel and save as an Excel workbook.  Then, open a command prompt or terminal and navigate to the directory holding the script and the input files.  Then, issue the command:
`python peer_eval_script.py`

You first be prompted for the name of the Excel file holding the survey results.  Enter the name of the file, including the extension:

```
Enter the file name of the Excel sheet where the survey results are located (including the file extension): TopicResearchPeerReview.xlsx
```

Next, you will be prompted for the name of the Excel file holding the student names and emails.  Again, enter the name with the entension:

```
Enter the file name of the Excel sheet where the emails and names of all students are located (including the file extension: students.xlsx
```

If your files are read successfully, you will see the message:

`Peer eval results successfully written to results.xls file in this directory`

Open the "results.xls" file in this directory to see your peer evalution results.

### Expected format of the all students Excel sheet:
The script requires a sheet containing the names and emails of all the student in the class.  This is necessary because it is
possible that a student in the course will neither complete the survey nor be evaluated by another student.  If this happens,
that student's email address will not be located in the survey Excel sheet and no score will be calculated for this student.  This would result in a results sheet containing less students than there are student in the course, forcing the instructor/TA to manually determine which students are missing.

The columns in the all students spreadsheet should have the following format:
```
Student email
Student last name
Student first name
```

There should be no header rows in the spreadsheet, only data.

An example would be:
```
johndoe@ksu.edu,Doe,John
janedoe@ksu.edu,Doe,Jane
```

### Expected format of Qualtrics survey results Excel sheet:
The columns of the Qualtrics survey spreadsheet should have the following format:
```
Ignored column 0
Ignored column 1
...
Ignored column 19  
This student's email  
Teammate email 1  
Teammate email 2  
...
Teammate email [MAX_GROUP_SIZE - 1]
This student's score for question 1
This student's score for question 2
...
This student's score for question [NUMBER_OF_QUESTIONS]
Teammate 1's score for question 1
Teammate 1's score for question 2
...
Teammate 1's score for question [NUMBER_OF_QUESTIONS]
Teammate 2's score for question 1
Teammate 2's score for question 2
...
Teammate 2's score for question [NUMBER_OF_QUESTIONS]
...
Teammate [MAX_GROUP_SIZE - 1]'s score for question 1
Teammate [MAX_GROUP_SIZE - 1]'s score for question 2
...
Teammate [MAX_GROUP_SIZE - 1]'s score for question [NUMBER_OF_QUESTIONS]
Comments written by this student
Ignored column
Ignored column
Ignored column
```

The rows of the spreadsheet should have the following format:
```
Question name (ignored)
Question prompt (ignored)
Ignored row
Student data entry 1
Student data entry 2
...
Student data entry [NUMBER_OF_COMPLETED_SURVEYS]
```

### Format of results.xls
After successful completion of the script, results.xls will contain the peer evalutation results.  The spreadsheet will contain,
for each student in the course, whether or not that student completed the survey, that student's self evaluation scores for
each question, that student's average peer evaluation score for each question, the comments written by that student, and the comments written by that student's teammates.

The columns are in the following format:
```
Student email
Student name
Whether or not the student completed the survey
Self evaluation score for question 1
Self evaluation score for question 2
...
Self evaluation score for question [NUMBER_OF_QUESTIONS]
Average peer evaluation score for question 1
Average peer evaluation score for question 2
...
Average peer evaluation score for [NUMBER_OF_QUESTIONS]
Comments written by this student
Comments written by this student's teammates in the following format: <teammate 1 name>: <comment>,<teammate 2 name>: <comment>,...
```

### Encountering errors reading in student data
The students enter their email address and the email address of their teammates when completing the survey.  This provides some degree
of normalization, as the email address format is checked with a regular expression.  However, students are still apt to enter
incorrect emails.  When you first run the script, you will likely encounter several error message saying that an email address
was not found in the all students worksheet.

If a student entered their own email address incorrectly, you will see this error message:
```
ERROR: Student name <student email address> not in students dictionary
```

When this occurs, the rest of that student's data will be skipped, so more errors may be uncovered once you fix this student's faulty
email address.

If a student entered a teammate's email incorrectly, you will see this error message:
```
ERROR: Teammate name <teammate email address> not in students dictionary
```

You will likely see several of these error message print out the first time you run the script.  If so, search for the faulty
email addresses in the survey results spreadsheet and fix it to the correct email.  You may need to look at the names of students in a group in Canvas and search for their email address in https://search.k-state.edu (fun, I know).

### Transferring the peer evaluation results to the grades spreadsheet
The easiest way to transfer the scores from the results spreadsheet to the grades spreadsheet is to sort both spreadsheets by
student email so that the row numbers in each spreadsheet correspond to the same student (you will need to first lock the header
in the results spreadsheet so that it will not be sorted in).  Then, copy and paste the columns (excluding the header) for completed survey, self evaluation scores, and average peer evaluation scores from the results spreadsheet into the grades spreadsheet.

### Updating the peer eval script
There are four constant variables at the beginning of the script:
```
MAX_GROUP_SIZE = 6
NUMBER_OF_QUESTIONS = 6
BEGINNING_COLS_TO_IGNORE = 20
BEGINNING_ROWS_TO_IGNORE = 3
```
These variables may need to be updated to reflect changes in the course or Qualtrics survey.
