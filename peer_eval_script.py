import xlrd
import xlwt
from xlwt import Workbook

# The maximum number of students in a group (the number of times each question is asked in the survey).
MAX_GROUP_SIZE = 6

# The number of questions asked on the survey per group member.
NUMBER_OF_QUESTIONS = 6

# The number of columns in the survey sheet before the column that asks the for the student's own email
BEGINNING_COLS_TO_IGNORE = 20

#The number of header rows before the actual student answers begin
BEGINNING_ROWS_TO_IGNORE = 3

# The name of the file where the survey results are stored
SURVEY_FILE_NAME = input("Enter the file name of the Excel sheet where the survey results are located (including the file extension): ")

# The name of the file holding all the student name and emails
STUDENTS_FILE_NAME = input("Enter the file name of the Excel sheet where the emails and names of all students are located (including the file extension: ")

class Student:
	def __init__(self, name):
		# A list of the scores that the student gives themselves. This is a list of ints: the first element is the student's
		# self score for Q1, the second element is the self score for Q2...
		self.self_scores = []
	
		# The scores that other students have given this student. This is a list of lists - the first element is a list
		# of all scores for Q1, the second a list of all scores for Q2...
		self.team_scores = [[] for x in range(NUMBER_OF_QUESTIONS)]
		
		# Whether or not this student filled out the survey
		self.completed_survey = False
		
		# The student's name in the format "Firstname Lastname"
		self.name = name
		
		# The comments written by this student
		self.comments_by_student = ""
		
		# The comments written by teammates in this student's group
		self.comments_by_teammates = ""

# Read in the survey excel sheet
survey_wb = xlrd.open_workbook(SURVEY_FILE_NAME)
survey_sheet = survey_wb.sheet_by_index(0)

# Read in the excel sheet that contains all the students' names. We need this sheet because not all students will complete the survey
# and it's theoretically possible that a student could not be named on the survey at all.
students_wb = xlrd.open_workbook(STUDENTS_FILE_NAME)
students_sheet = students_wb.sheet_by_index(0)

# Open a workbook to write the results to
results_wb = Workbook()
results_sheet = results_wb.add_sheet("Results")

# A dictionary that holds all the students - the key is the student's email and the value is a Student object
# that holds all the student's eval scores and whether or not the student completed the survey.
students = {}

# Iterate through each entry in the sheet of all student names and add an entry for each student to the students dictionary.
for i in range(0, students_sheet.nrows):
	email = students_sheet.cell_value(i,0)
	last_name = students_sheet.cell_value(i,1)
	first_name = students_sheet.cell_value(i,2)
	students[email] = Student("{} {}".format(first_name,last_name))

# Iterate through each entry in the survey sheet (skipping the header rows) and add the scores to the appropriate Student object
for i in range(BEGINNING_ROWS_TO_IGNORE,survey_sheet.nrows):
	email = survey_sheet.cell_value(i,BEGINNING_COLS_TO_IGNORE)
	if email == "blank" or email == "":
		continue
	if email not in students:
		print("ERROR: Student name {} not in students dictionary".format(email))
		continue
		
	students[email].completed_survey = True
	
	# Add the self scores
	for j in range(BEGINNING_COLS_TO_IGNORE + MAX_GROUP_SIZE, BEGINNING_COLS_TO_IGNORE + MAX_GROUP_SIZE + NUMBER_OF_QUESTIONS):
		students[email].self_scores.append(survey_sheet.cell_value(i,j))
		
	# Add in this student's comments about the group
	comments = survey_sheet.cell_value(i,BEGINNING_COLS_TO_IGNORE + MAX_GROUP_SIZE + (NUMBER_OF_QUESTIONS * MAX_GROUP_SIZE))
	students[email].comments_by_student += comments
		
	# Add the teammate scores
	for k in range(1, MAX_GROUP_SIZE):
		teammate_email = survey_sheet.cell_value(i,BEGINNING_COLS_TO_IGNORE + k)
		if teammate_email == "blank" or teammate_email == "":
			continue
		if teammate_email not in students:
			print("ERROR: Teammate name {} not in students dictionary".format(teammate_email))
			continue
		
		# Add the student's teammate scores to the teammate's entry
		for question_number in range(NUMBER_OF_QUESTIONS):
			score = survey_sheet.cell_value(i, BEGINNING_COLS_TO_IGNORE + MAX_GROUP_SIZE + NUMBER_OF_QUESTIONS*(k) + question_number)
			if score != "" and score != "blank":
				students[teammate_email].team_scores[question_number].append(score)
				
		# Add the student's comments to the teammates's comments_by_teammates
		students[teammate_email].comments_by_teammates += "{}: {},".format(students[email].name,comments)
		
# Write the headers to the results sheet
results_sheet.write(0,0,"Student Email")
results_sheet.write(0,1,"Student Name")
results_sheet.write(0,2,"Completed Survey?")
number_begin_cols = 3

# Write a header for each question for the self evaluation
for i in range(NUMBER_OF_QUESTIONS):
	results_sheet.write(0,number_begin_cols + i,"Self Eval: Q{}".format(i+1))
	
# Write a header for each question for the average peer eval score
for i in range(NUMBER_OF_QUESTIONS):
	results_sheet.write(0,number_begin_cols + NUMBER_OF_QUESTIONS + i,"Average Peer Eval: Q{}".format(i+1))
	
# Write a header for comments by this student
results_sheet.write(0,number_begin_cols + (NUMBER_OF_QUESTIONS * 2),"Comments Written by this Student")

# Write a header for comments by teammates in this student's group
results_sheet.write(0,number_begin_cols + (NUMBER_OF_QUESTIONS * 2) + 1,"Comments Written by Teammates of this Student")

# Calculate the student's average peer eval scores and write them to the results sheet
row_number = 1
for email in students.keys():
	# Write the student's email
	results_sheet.write(row_number,0,email)
	
	# Write the student's name
	results_sheet.write(row_number,1,students[email].name)
	
	# Write whether or not the student completed the survey
	results_sheet.write(row_number,2,students[email].completed_survey)
	
	# Write self eval scores
	if students[email].completed_survey:
		for i in range(NUMBER_OF_QUESTIONS):
			results_sheet.write(row_number,number_begin_cols + i,students[email].self_scores[i])
		
	# Write average peer eval scores
	for i in range(NUMBER_OF_QUESTIONS):
		peer_scores = students[email].team_scores[i]
		# Avoiding a divide by zero error if no one evaluated the student
		if (len(peer_scores) > 0):
			score = sum(peer_scores)/len(peer_scores)
			results_sheet.write(row_number,number_begin_cols + NUMBER_OF_QUESTIONS + i,score)
			
	# Write this student's comments
	results_sheet.write(row_number,number_begin_cols + NUMBER_OF_QUESTIONS * 2,students[email].comments_by_student)
	
	# Write this student's comments
	results_sheet.write(row_number,number_begin_cols + NUMBER_OF_QUESTIONS * 2 + 1,students[email].comments_by_teammates)
	
	row_number += 1

# Save the results to a file
results_wb.save("results.xls")

print("Peer eval results successfully written to results.xls file in this directory")



