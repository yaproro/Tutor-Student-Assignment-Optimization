import sys
import cplex
import docplex.mp
import pandas as pd
from docplex.mp.model import Model

debug_mode = False #Toggle if you wish to view model variables, constraints and objective

#Set up classes to hold information regarding new students, exisiting students and tutor
class NewStudent():
	def __init__(self, studentid, needs, tuitioncentre):
		self.studentid = studentid
		self.needs = needs
		self.tuitioncentre = tuitioncentre

	def __str__(self):
		return f"New student with id: {self.studentid} and needs: {self.needs} at tution center: {self.tuitioncentre}"

class ExistingStudent():
	def __init__(self, studentid, needs, tuitioncentre, tutorid):
		self.studentid = studentid
		self.needs = needs
		self.tuitioncentre = tuitioncentre
		self.tutorid = tutorid

	def __str__(self):
		return f"Exisiting student with id: {self.studentid} and needs: {self.needs} at tution center: {self.tuitioncentre} assgined to tutor: {self.tutorid}"

class Tutor():
	def __init__(self, tutorid, skills, prefcenter1, prefcenter2, capacity):
		self.tutorid = tutorid
		self.skills = skills
		self.prefcenter1 = prefcenter1
		self.prefcenter2 = prefcenter2
		self.capacity = capacity 

	def __str__(self):
		return f"Tutor with id: {self.tutorid} and skills: {self.skills}, prefer teaching at: {self.prefcenter1} and {self.prefcenter2} with capacity: {self.capacity}"
 
#---------------------------------------------------------------------------------------------------------

#------------
#Import data 
#------------

#Get data file path from user input then try to read the file
#Terminate program if file not found or there were issues reading file 
data_path = input("Enter the path to the Excel data file: ")

try:
	sheets = pd.read_excel(data_path, sheet_name=None)
except FileNotFoundError:
    print(f"Error: File at {data_path} not found.")
    sys.exit(1)
except Exception as e:
    print(f"Error: {e} when reading excel file.")
    sys.exit(1)

#Assumes excel workbook contains three sheets "New Students", "Tutor Information" and "Existing Students". If not, exit.
for req_sheet in ["New Students", "Tutor Information", "Existing Students"]:
    if req_sheet not in sheets:
        print(f"Error: Missing '{req_sheet}' sheet in Excel workbook.")
        sys.exit(1)

#Extract data from the three sheets required
new_students_df = sheets["New Students"]
existing_students_df = sheets["Existing Students"]
tutors_df = sheets["Tutor Information"] 

#-----------------
#Populate the data
#-----------------

new_students = []
existing_students = []
tutors = []

#Checks each sheet for their required columns. If any column missing, exit. 
for req_column in ["studentId", "tutoringNeed", "tuitionCentre"]:
	if req_column not in new_students_df.columns: 
		print(f"Error: Missing required column '{req_column}' in 'New Students' sheet.")
		sys.exit(1) 

for req_column in ["studentId", "tutoringNeed", "tuitionCentre", "tutorId"]:
	if req_column not in existing_students_df.columns: 
		print(f"Error: Missing required column '{req_column}' in 'Existing Students' sheet.")
		sys.exit(1) 

for req_column in ["tutorId", "tutoringSkills", "preferredCentre1", "preferredCentre2", "maxOverallCapacity"]:
	if req_column not in tutors_df.columns: 
		print(f"Error: Missing required column '{req_column}' in 'Tutor Information' sheet.")
		sys.exit(1) 

#For each sheet, iterate the rows and extract data of each new student, active exisiting student and tutors
for _, row in new_students_df.iterrows():
	new_students.append(NewStudent(row["studentId"], row["tutoringNeed"], row["tuitionCentre"]))

for _, row in existing_students_df.iterrows():
	if row["active"] :
		existing_students.append(ExistingStudent(row["studentId"], row["tutoringNeed"], row["tuitionCentre"], row["tutorId"]))

for _, row in tutors_df.iterrows():
	tutors.append(Tutor(row["tutorId"], row["tutoringSkills"], row["preferredCentre1"], row["preferredCentre2"], row["maxOverallCapacity"]))

#For sanity check of what data got loaded
print(f"Data loaded consist of {len(new_students)} new students, {len(existing_students)} active existing students and {len(tutors)} tutors.")

#---------------------------------------------------------------------------------------------------------

#Creating the model 
m = Model(name='tutor_allocation')

#-----------------
#Define variables
#-----------------

#Define the decision variables
#xij is a binary var, it is 1 if and only if student i is assigned to tutor j  
x = {(i.studentid, jj.tutorid): m.binary_var(name=f"x_{i.studentid}_{j.tutorid}") for i in new_students for j in tutors}

#Define the helper variables 
#yj is a binary var, it is 1 if and only if tutor j is assigned at least one student 
y = {j.tutorid: m.binary_var(name=f"y_{j.tutorid}") for j in tutors}
#wj represents the workload of the tutor. It is an integer var with value the number of students assigned to the tutor 
w = {j.tutorid: m.integer_var(name=f"w_{j.tutorid}") for j in tutors}
#U and L are integer variables that represent the upper and lowerbounds of all assigned tutor's workload 
U = m.integer_var(name="workload_upperbound") 
L = m.integer_var(name="workload_lowerbound") 

#Set upperbound for interger var, U and L.
#Since we will be minimizing U - L, U and L should not take on a value more than the max capacity across all tutors 
#Did not reduce L's upperbound further as it is not necessarily that all tutors are selected 
maxcapacity = max(j.capacity for j in tutors)
U.ub = maxcapacity
L.ub = maxcapacity 

#-------------------
#Define constraints 
#-------------------

#1) Ensure every student only has one tutor 
#2) Ensure every student's extensive tutoring needs are met 
for i in new_students:
	if i.needs == "Extensive":
		#Ensure student with extensive needs has to be assigned exactly one tutor with extensive skills
		m.add_constraint(sum(x[i.studentid, j.tutorid] for j in tutors if j.skills == "Extensive") == 1) 
		for j in tutors:
			if j.skills != "Extensive":
				#Ensure student with extensive needs will not be assign a tutor that doesnt have extensive skills
				x[i.studentid, j.tutorid].ub = 0  
	else:
		#For other students with normal needs, they can be assigned any tutor and exactly one tutor
		m.add_constraint(sum(x[i.studentid, j.tutorid] for j in tutors) == 1)

#3) Ensure tutors are not assigned students beyond their maximum capacity 
#4) Ensure helper variable y is correctly set, if its 0 the tutor cannot be assigned any student, if its 1 the tutor has to respect its max capacity 
# Set lowerbound for binary var, yj. If there is exisiting student, the tutor is used  
# Set upperbound for integer var, wj. The workload shouldnt exceed capacity 
#5) Ensure helper variable w value is correctly set and takes into account exisiting students 
#6) Ensure helper variable U is upperbound across all tutor workload 
#7) Ensure helper variable L is lowerbound across all tutor workload 
for j in tutors:
	existingworkload = sum(1 for i in existing_students if i.tutorid == j.tutorid)
	#(j.capacity - existingworkload) is the remaining number of students we can assign to a tutor 
	m.add_constraint(sum(x[i.studentid, j.tutorid] for i in new_students) <= (j.capacity - existingworkload)*y[j.tutorid] ) #Constaint 3 and 4
	if existingworkload > 0:
		#If tutor has exisiting student, he is already "assigned"
		y[j.tutorid].lb = 1   

	#Since wj are integer variables, just to set an upperbound on their values to reduce search space 
	w[j.tutorid].ub = j.capacity  
	#wj represents the number of students assigned to tutor j, it should be the sum over all students xij and the existing workload
	m.add_constraint(sum(x[i.studentid, j.tutorid] for i in new_students) + existingworkload == w[j.tutorid] ) #Constraint 5

	#We want U and L to be the upperbound and lowerbound of all assigned tutor's workload 
	#If a tutor is not assigned (yj=0 and wj=0), the value of U and L should not be affected by these equations too 
	#For constraint 6, if yj = 0, wj = 0 <= U, this constraint doesnt affect the value of U  
	#For constraint 7, if yj = 1, we get wj >= L as intended. 
	#                  if yj = 0, we get wj = 0 >= L - (max capacity) and thus (max capcity) >= L, this constraint doesnt affect the value of L
	m.add_constraint( w[j.tutorid] <= U ) #Constraint 6
	m.add_constraint( w[j.tutorid] >= L - (maxcapacity)*(1-y[j.tutorid]) ) #Constraint 7

#-----------------
#Define objective
#-----------------

c1 = 10   #Penalty for not balancing workload
c2 = 0    #Penalty for each tutor used 
c3 = 4    #Penalty for not giving tutors any preferred location 
c4 = 2    #Penalty for not giving tutors 1st preferred location 

#Want to minimize total sum of yj which corresponds to the number of tutors assigned
tutors_used = sum(y[j.tutorid] for j in tutors)  

not_preferred_location_terms = []
second_choice_location_terms = []
for j in tutors:
	for i in new_students:   
		if i.tuitioncentre not in [j.prefcenter1, j.prefcenter2]:
			#Want to minimize total sum of xij where student i goes to tuition centre that is not in tutor j's preference
			#which corresponds to minimizing assigning a student from a tuition centre outside tutor's preference
			not_preferred_location_terms.append(x[i.studentid, j.tutorid]) 
		elif i.tuitioncentre == j.prefcenter2:
			#Similarly, there is a smaller penalty if a tutor is assigned a student from a tuition centre that is their second choice
			#Only no penalty given for student assignment from a tuition centre that is their first choice
			second_choice_location_terms.append(x[i.studentid, j.tutorid]) 

#Minimizing the first term (U-L) is to balance workload among tutors assigned
m.minimize( c1*(U-L) + c2*tutors_used + c3*(m.sum(not_preferred_location_terms)) + c4*(m.sum(second_choice_location_terms)))

#----------------------
#For debugging purposes
#----------------------

if debug_mode: 
	m.print_information() 
	print(m.objective_expr)               #To check each of the objective's terms 
	for ct in m.iter_constraints():
		print(ct)                         #To check every constraint  
	for var in m.iter_variables():
		print(var.name, var.lb, var.ub)   #To check every variable's upper and lower bound 

s = m.solve() 
m.print_solution()

#Print objective penalty breakdown and the solution
if s: 
	#Get breakdown of each term in the objective
    workload_balancing_penalty = c1 * (s.get_value(U) - s.get_value(L))
    tutors_used_penalty = c2 * sum(s.get_value(y[j.tutorid]) for j in tutors)
    not_preferred_location_penalty = c3 * sum(s.get_value(var) for var in not_preferred_location_terms)
    second_choice_location_penalty = c4 * sum(s.get_value(var) for var in second_choice_location_terms)
    
    #Print breakdown of each term in the objective 
    print("-------------------------------------------------------------")
    print("Objective breakdown:")
    print(f"Total penalty: {s.objective_value}")
    print(f"Balancing workload : weight = {c1}, U = {int(s.get_value(U))}, L = {int(s.get_value(L))}. Total = {int(workload_balancing_penalty)}")
    print(f"Tutors used : weight = {c2}, tutors used = {sum(s.get_value(y[j.tutorid]) for j in tutors)}. Total = {int(tutors_used_penalty)}")
    print(f"Not preferred location penalty : weight = {c3}, Total = {int(not_preferred_location_penalty)}")
    print(f"Second choice location penalty : weight = {c4}, Total = {int(second_choice_location_penalty)}") 
    print("-------------------------------------------------------------")

    print("Solution:")
    for j in tutors: 
        if s.get_value(w[j.tutorid]):
            print(f"Tutor {j.tutorid} has {int(s.get_value(w[j.tutorid]))} student(s) assigned.")
    for i in new_students:
        for j in tutors:
            if s.get_value(x[i.studentid, j.tutorid]):  
                print(f"Student {i.studentid} is assigned to Tutor {j.tutorid}.")