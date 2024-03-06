
# Run this the first time you run the code
# !pip install python-docx
# LIBRARY TO GENERATE WORD DOC
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
# LIBRARY TO GET DATE/TIME
from datetime import datetime, timedelta, timezone, date
# LIBRARY FOR TIMEZONE
import pytz
# LIBRARY TO CONVERT NUMBER TO WEEKDAY NAME
import calendar

# DEFINE THE CURRENT DATE & TIME
current_date = date.today().strftime("%Y-%m-%d")
sydney_tz = pytz.timezone('Australia/Sydney')
current_datetime = datetime.now(sydney_tz).strftime("%Y-%m-%d-%H-%M-%S")
sydney_date = datetime.now(pytz.timezone('Australia/Sydney'))

# BASED ON A LIST, ASK THE USER QUESTIONS TO GET INFORMATION 
def getQuestions():
  questions = {
      "Enter the heart rate": 0
      , "Enter the client's Name (First & Last Name)": ""
      , "Enter the client's appointment time (24 HOUR TIME HHMM)": ""
      , "Enter the client's appointment date (YYYY-MM-DD)": ""
  }
  
  # LOOP THROUGH THE LIST AND ADD EACH ANSWER INTO AN ARRAY/LIST
  for i in questions:
    question = i
    temp_var = input(f"{question}: ")
    questions[i] = temp_var

    # DEPENDING ON THE HEART RATE, THEN PROVIDE THE USER WITH A DIFFERENT TEMPLATE - 90 IS HARD CODED 
  if int(questions['Enter the heart rate']) >= 90:
      template_version = "A"
  elif int(questions['Enter the heart rate']) < 90:
      template_version = "B"

    # CALL FUNCTIONS TO FORMAT DATE VALUES
  week_day_num = formatDate(questions["Enter the client's appointment date (YYYY-MM-DD)"], "week_day_num")
  week_day = formatDate(questions["Enter the client's appointment date (YYYY-MM-DD)"], "week_day")
    
    # CALL FUNCTION TO BEGIN CREATING THE DOCUMENT
  createDocument(questions, template_version, week_day, week_day_num)

# BEGINING OF FUNCTION TO FORMAT THE DATE VALUES
def formatDate(date_string, requested_value):
  # INITIALISING VARIABLE
  return_value = ""
  
  # DATE CONVERSION INTO PROPER FORMAT
  date_formatted = datetime.strptime(date_string, '%Y-%m-%d')
  
  # USING THE FORMATTED DATE TO GET WEEK DAY NUMBER
  week_day_num = date_formatted.weekday()

  # USING THE WEEKDAY NUMBER TO GET THE WEEK DAY NAME
  week_day = calendar.day_name[date_formatted.weekday()]
  
  # DEPENDING ON WHAT VARIABLE IS PASSED IN, THEN RETURN A DIFFERENT VALUE (NUMBER OR NAME)
  if requested_value == "week_day":
    return_value = week_day
    print(f"Formatted Date: {date_formatted}, Weekday: {week_day}")
  elif requested_value == "week_day_num":
    return_value = week_day_num
    print(f"Formatted Date: {date_formatted},  Week day num: {week_day_num}")
  return return_value



        
# CREATE THE DOCUMENT
def createDocument(questions, template_version, week_day, week_day_num):
  
  minus_days = calculateDates(week_day_num)
  calculated_time = calculateTime(questions["Enter the client's appointment time (24 HOUR TIME HHMM)"])
  # Create a document
  doc = docx.Document()

  # Add a paragraph to the document
  p = doc.add_paragraph()

  # Add some formatting to the paragraph
  p.paragraph_format.line_spacing = 1
  p.paragraph_format.space_after = 0

  # Add a run to the paragraph
  # run = p.add_run("python-docx")
  run = p.add_run("")

  # Add some formatting to the run
  run.bold = True
  run.italic = True
  run.font.name = 'Arial'
  run.underline
  run.font.size = docx.shared.Pt(16)

  # Add more text to the same paragraph
  run = p.add_run("CARDIAC CT PREPARATION")

  # paragraph.alignment = 0 # for left, 1 for center, 2 right, 3
  # Format the run
  run.alignment = 1
  run.bold = True
  run.font.name = 'Arial'
  run.font.size = docx.shared.Pt(16)

  # Add another paragraph (left blank for an empty line)
  doc.add_paragraph()

  # Add another paragraph
  p = doc.add_paragraph()
  run.font.size = docx.shared.Pt(12)

  # Add a run and format it
  var_name = questions["Enter the client's Name (First & Last Name)"]
  run = p.add_run(f"Name: {var_name}")
  run.font.name = 'Arial'
  run.font.size = docx.shared.Pt(12)

  p = doc.add_paragraph()
  var_date = questions["Enter the client's appointment date (YYYY-MM-DD)"]
  var_time = questions["Enter the client's appointment time (24 HOUR TIME HHMM)"]
  run = p.add_run(f"Appointment: {week_day} {var_date} at {calculated_time['appointment_time']}. Arrive 15 minutes before at {calculated_time['arrival_time']}")
  run.font.name = 'Arial'
  run.font.size = docx.shared.Pt(12)

  doc.add_paragraph()

  if template_version == "A":
    p = doc.add_paragraph()
    run = p.add_run(f"""
Take one tablet (Metoprolol 50mg) {minus_days['1_days_before']} night before bed.

Take one tablet (Metoprolol 50mg) {week_day} Morning at {calculated_time['2_hours_before_appointment_time']}, 2 hours before appointment time.
    """)
    run.font.name = 'Arial'
    run.font.size = docx.shared.Pt(12)
  elif template_version == "B":
    p = doc.add_paragraph()
    run = p.add_run(f"YOU'RE HEALTHY, DON'T NEED TO TAKE ANY MEDICINE.")
    run.font.name = 'Arial'
    run.font.size = docx.shared.Pt(12)

  doc.add_paragraph()
  p = doc.add_paragraph()
  run = p.add_run(f"{minus_days['1_days_before']}")
  run.font.name = 'Arial'
  run.underline = True
  run.font.size = docx.shared.Pt(12)

  p = doc.add_paragraph()
  run = p.add_run(f"""
    No caffeine or stimulants (Coffee, tea, diet pills, sports drinks, chocolate)
    No Viagra""")
  run.font.name = 'Arial'
  run.font.size = docx.shared.Pt(12)

  doc.add_paragraph()
  p = doc.add_paragraph()
  run = p.add_run(f"{week_day} Morning")
  run.font.name = 'Arial'
  run.underline = True
  run.font.size = docx.shared.Pt(12)

  p = doc.add_paragraph()
  run = p.add_run("""
    No exercise
    No smoking or anti smoking medicines
    Fasting for four hours prior to your appointment
    Keep up water intake
    """)
  run.font.name = 'Arial'
  run.underline
  run.font.size = docx.shared.Pt(12)

  p = doc.add_paragraph()
  run = p.add_run("Take all your normal medications")
  run.font.name = 'Arial'
  run.underline
  run.font.size = docx.shared.Pt(12)

  # Save the document
  client_name = questions["Enter the client's Name (First & Last Name)"]
  doc_name = (f"South-East-Radiology-Instructions-{client_name}-{current_datetime}.docx").replace(" ","-")
  doc.save(doc_name)

# START OF DATE CALCULATION FUNCTION
# DEPENDING ON THE CALCULATED VALUE APPLY A TRANSFORMATION TO GET THE CORRECT DATE
def calculateDates(week_day_num):
    returnDayNums ={
        "2_days_before": ""
        ,"1_days_before": ""
    }

    returnDayNums["2_days_before"] = week_day_num -2
    returnDayNums["1_days_before"] = week_day_num -1

    # IF THE VALUE FALLS UNDER 0, THEN YOU NEED TO ADD 7 TO BRING IT BACK INTO THE DATE RANGE
    if returnDayNums["2_days_before"] < 0:
        returnDayNums["2_days_before"] = calendar.day_name[returnDayNums["2_days_before"]+7]
    else: 
        returnDayNums["2_days_before"] =calendar.day_name[returnDayNums["2_days_before"]]

    if returnDayNums["1_days_before"] < 0:
        returnDayNums["1_days_before"] = calendar.day_name[returnDayNums["1_days_before"]+7]
    else:
        returnDayNums["1_days_before"] = calendar.day_name[returnDayNums["1_days_before"]]
    
    print(returnDayNums)
    return returnDayNums

# CALCULATING THE APPOINTMENT TIME AND ARRIVAL TIME
def calculateTime(time):
   formatted_time = {
      "appointment_time": ""
      ,"arrival_time": ""
      , "2_hours_before_appointment_time": ""
   }
   converted_appointment_time = datetime.strptime(time, "%H%M")
   converted_arrival_time = datetime.strptime(time, "%H%M") - timedelta(minutes = 15)
   converted_2_hours_before_appointment_time = datetime.strptime(time, "%H%M") - timedelta(hours = 2)
   formatted_time["appointment_time"] = converted_appointment_time.strftime("%I:%M %p")
   formatted_time["arrival_time"] = converted_arrival_time.strftime("%I:%M %p")
   formatted_time["2_hours_before_appointment_time"] = converted_2_hours_before_appointment_time.strftime("%I:%M %p")
   print(formatted_time)
   return formatted_time

def main():
  getQuestions()

if __name__ == '__main__':
    main()