# Tutor-Student Assignment Optimization using CPLEX Mixed Integer Programming 

This repository contains a Python notebook that formulates and solves a **tutor–student assignment optimization problem** using **IBM CPLEX** and python library **docplex**. 

Background
A tutoring organisation with multiple tuition centres admits a fresh batch of students every month. New students are each assigned to a tutor and attend classes at one of the centres, as chosen by the student. Tutors on the other hand are free to teach at any centre and have indicated their top 2 preference. Before the start of a new month, the organization assigns a tutor to each new incoming student, meeting specific requirements.

Requirements
1) Each student can only be assigned one tutor.
2) Students that require extensive tutoring may only be assigned to tutors with extensive skills. Students who require normal tutoring can be assigned to tutors with normal or extensive skills.
3) The total number of new students assigned and existing students from previous assignments cannot exceed a tutor's maximum overall capacity.

The goal is to optimally assign tutors to students based on their tutoring needs while considering the following two scenarios. 
1) Minimize total number of tutors assigned while maximizing tutor’s preference on tuition centre.
2) Balance tutor’s workload while maximizing tutor’s preference on tuition centre.

---

## Input File Requirements

The python script will prompt for an **Excel file path** as input. 
The following Excel file extensions are supported: .xls, .xlsx, .xlsm, .xlsb, .odf, .ods and .odt 

Additionally, the input Excel file **must** contain the following three sheets: 
1. **New Students**  
2. **Tutor Information**  
3. **Existing Students**

Each sheet must include the following columns:
### 1. New Students
| Column Name | Description |
|--------------|-------------|
| `studentId` | Unique identifier for each student |
| `tutoringNeed` | 'Normal' or 'Extensive' need required |
| `tuitionCentre` | Tuition centre location |

### 2. Tutor Information
| Column Name | Description |
|--------------|-------------|
| `tutorId` | Unique identifier for each tutor |
| `tutoringSkills` | 'Normal' or 'Extensive' skills acquired |
| `preferredCentre1` | First preferred tuition centre |
| `preferredCentre2` | Second preferred tuition centre |
| `maxOverallCapacity` | Maximum number of students a tutor can be assigned |

### 3. Existing Students
| Column Name | Description |
|--------------|-------------|
| `studentId` | Unique identifier for each student |
| `tutoringNeed` | 'Normal' or 'Extensive' need required |
| `tuitionCentre` | Tuition centre location |
| `tutorId` | Tutor id of tutor currently assigned to the student |

Note that missing columns or incorrect sheet names will cause the script to terminate early.  
---

## Dependencies

Refer to `requirements.txt` for full list of packages and versions required to run python notebook
If using only python script, refer to `requirements_py_script_only.txt`

Install dependencies via pip:
```bash
pip install cplex docplex pandas openpyxl jupyterlab
```
 
---

## Modifiable values within script 
 
**Debug Mode**
For detailed optimization model information that includes all variables, constraints and objective components, set debug_mode to True on line 7 of the python script or 2nd cell in the python notebook.  
 
**Objective term weights**
Depending on how the tutoring organization values each set of objectives, the weights can be modified to quantitatively reflect its importance. 
On lines 187–190 of the python script or 5th last cell in the python notebook, the model defines four objective weight variables (`c1`, `c2`, `c3`, `c4`).  
The higher the weights, the more important it is to achieve the following goals. 

c1 - Importance of balancing all assigned tutor's workload  
c2 - Importance of reducing number of tutors hired
c3 - Importance of tutors having 1st or 2nd choice preferred location
c4 - Importance of tutors having 1st choice preferred location 
Generally c4 < c3. 
 
---

## Running the Script

1. Ensure dependencies are installed and input data follows the format stated in Input File Requirements. 
2. Toggle debug mode if required and modify weights c1 - c4 as required.  
3. Run the script:
   ```bash
   python tutor_assignment.py
   ```
4. Enter Excel file path when prompted "Enter the path to the Excel data file". 

---

## Expected Output

When the data is successfully loaded from file, the script will print : Data loaded consist of X new students, Y active existing students and Z tutors.
If the solver successfully gets a solution, it will print variable values followed by a breakdown of the objective score.
Lastly, we expect to see a list of students and their respective assigned tutors.   

Example : 
```
-------------------------------------------------------------
Objective breakdown:
Total penalty: 12.0
Balancing workload : weight = 10, U = 5, L = 4. Total = 10
Tutors used : weight = 0, tutors used = 7.0. Total = 0
Not preferred location penalty : weight = 4, Total = 0
Second choice location penalty : weight = 2, Total = 2
-------------------------------------------------------------
Solution:
Student S0001 is assigned to Tutor A001
Student S0002 is assigned to Tutor A003
Student S0003 is assigned to Tutor A009
Student S0004 is assigned to Tutor A001
Student S0005 is assigned to Tutor A009
... 
```