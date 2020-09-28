# Attendance-Management-System
An fully autonomous attendance management system using face recognition [Opencv].

A script that recognises faces and mark attendance for the recognised faces in an excel sheet. 
You need following libraries pre-installed on your system: 
1.face_recognition 
2.Opencv 
3.openpyxl 
4.datetime

HOW TO USE:

1. Save images of people as '1.jpg','2.jpg'...... numbers being the roll numbers of the person!(here in valid_images folder) 
2. Run the program.
3. An excel file will be created, marked with the attendance for the faces it recognised.
Keep all the images and the python script in the same folder as well as run the python script for the same folder.
4. Excel will be created with the number of the month as name. For eg. if current month is january, then the name of the excel file created will be '1.xlsx'.
Suppose today is 4th of january then the attendance will be marked in the 'D' column of the '1.xlsx' file.
