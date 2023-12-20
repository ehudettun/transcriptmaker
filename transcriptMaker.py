import gspread

header = """<html>

<head>   <!-- Required meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">

    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous">

    <title>Transcript</title>
</head>
<body>
    <div class="container">
        <div class="row flex-nowrap">
            <div class = "col-md-4">
        <img  src='iclogo.png' class = 'img-fluid'>
            </div>
            <div class = "col-md-4"></div>
            <div class = "col-md-4">
  <p class="mt-5">Internal Compass Music <br>
    "Matanel" Creative Music Program <br>
    PO #362, Mitzpe Ramon 8060000 <br>
    Phone: +972-(0)8-376-0064
    internalcompassmusic com
  </h3>
</div>
</div>

<br>
<br>
<center><h4>Student Details</h4></center>
<br>"""

footer = """</div>
<script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/popper.js@1.12.9/dist/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@4.0.0/dist/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
</body>
</html>"""

gc = gspread.service_account(filename='studious-karma-323716-94a7ae1ccc79.json')
gradesfile = gc.open("grades")
coursesSheet = gc.open("Master Course List").worksheet('master')
courseDetails = ["Course Code", "Course Title","Teacher", "Hours/Semester", "Credits", "Yearly/Semester", "Grade"]
studentsfile = gc.open("Master Student List").worksheet("master")
studentID = '318250149'
studentCell = studentsfile.find(studentID)
studentRow = studentsfile.row_values(studentCell.row)
studentGrades = {}
for gradesheet in gradesfile.worksheets():
    cell = gradesheet.find(studentID)
    if cell:
        yeargrades = []
        row = gradesheet.row_values(cell.row)
        for index,grade in enumerate(row):
           if index>4:
                if grade:
                    yeargrades.append([gradesheet.cell(4,index+1).value,grade])
        studentGrades[gradesheet.title] = yeargrades








file=open('transcript.html', 'w', encoding="utf-8")
file.write(header)




file.write("<table class='table'>")
file.write("<th>")
file.write("<td>Name</td><td>ID</td><td>Date of Birth</td><td>Starting Date</td><td> Graduation Date</td><td>Instruent/Department")
file.write("</tr>")
file.write("<th>")
for i in [3,0,5,6,7,10]:
    file.write("<td>"+studentRow[i]+"</td>")
file.write("</tr>")
file.write("</table>")
file.write("<br><center><h4>Academic Records</h4></center><br>")

for year in studentGrades:
    file.write("<h4>"+str(year)+"</h4>")    
    file.write("<table class='table'>")
    file.write("<tr>")
    for detail in courseDetails:
        file.write("<td>"+detail+"</td>")
    file.write("</tr>")
    for grade in studentGrades[year]:
        print(grade)
        courseCode=grade[0]
        gradeValue=grade[1]
        courseRow = coursesSheet.row_values(coursesSheet.find(courseCode).row)
        courseName = courseRow[2]
        teacher = courseRow[coursesSheet.find(year).col-1]
        hours = courseRow[3]
        credits = courseRow[4]
        yearly =courseRow[5]
        file.write("<tr><td>"+courseCode+"</td><td>"+courseName+"</td><td>"+teacher+"</td><td>"+hours+"</td><td>"+credits+"</td><td>"+yearly+"</td><td>"+gradeValue+"</td></tr>")
    file.write("</table>")





file.write(footer)






