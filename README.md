# School-Report-Sheet
Student Report Sheet README
Project Title
Automated Student Report Sheet in Excel

Overview
This project is an automated Excel-based student report sheet designed for teachers and schools to manage student academic records efficiently.
The sheet automatically:
•	Calculates subject grades 
•	Generates comments/remarks 
•	Computes average scores 
•	Assigns overall grades 
•	Determines class position 
•	Displays pass/fail status 
using Excel formulas such as:
•	VLOOKUP 
•	AVERAGE 
•	SUM 
•	RANK 
•	IF 

Features
Student Information Section
Stores:
•	Student ID 
•	Student Name 
•	Department/Class 
•	Session 
•	Term 

Automatic Grade Generation
Grades are generated automatically from student scores using VLOOKUP.
Example Grade Scale
Score Range	Grade
80–100	    A
70–79	    B
50–69	    C
40–49	    D
0–39	    F

Automatic Remarks
Remarks are generated automatically based on average score.
Example Remarks
Score Range	Remark
80–100	    Excellent
70–79	    Very Good
60–69	    Good
50–59	    Fair
0–49	    Fail

Performance Summary
The report sheet automatically calculates:
•	Total Score 
•	Average Score 
•	Highest Score 
•	Lowest Score 
•	Overall Grade 
•	Pass/Fail Status 

Class Positioning
Student positions are generated automatically using the RANK function.
Example:
•	1st 
•	2nd 
•	3rd 

Excel Functions Used
Function	Purpose
VLOOKUP	    Generate grades and remarks
AVERAGE	    Calculate average score
SUM	        Calculate total score
MAX	        Highest score
MIN	        Lowest score
RANK	    Position in class
IF	        Pass/Fail logic

Images/Screenshoots
![alt text](<Screenshot (21)-1.png>)
![alt text](<Screenshot (22)-1.png>)

Sample Formulas
Average Score
=AVERAGE(B8:B14)
Grade Generation
=VLOOKUP(B8,Sheet3!$A$2:$B$6,2,TRUE)
Remark Generation
=VLOOKUP(B16,Sheet3!$D$2:$E$6,2,TRUE)
Position in Class
=RANK(E2,$E$2:$E$21,0)

Folder Structure
Workbook

Sheet1 → Student Database
Sheet2 → Report Card
Sheet3 → Reference Tables

Requirements
•	Microsoft Excel 2016 or later 
•	Basic understanding of Excel formulas 

Future Improvements
Possible upgrades:
•	Dashboard analytics 
•	Charts and performance graphs 
•	Attendance tracking 
•	Multiple term automation 
•	Printable report card layout 
•	Automatic PDF export 

Author
Created by: [David Chimezie]

License
This project is free to use for educational purposes.
