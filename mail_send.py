import smtplib
import time
import datetime
import csv
from datetime import datetime,date,timedelta
from jinja2 import Environment, PackageLoader, select_autoescape, FileSystemLoader
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import private
import html

#jinja2 library generate the HTML code dynamically using HTML Templates while
# email.mime takes the generated HTML code and render it.
Check = False

def extract_data():
    tutor_names = []
    time = []
    tutor_email = []
    tutees = []
    dates = []
    with open("C:\\Users\\kero6\Desktop\\daily report.csv", "r") as f:
        reader = csv.DictReader(f)
        columns = ["Tutor First Name", "Tutor Last Name", "Session Time", "Tutor Email", "Tutee Name", "Session Date"]
        today = date.today()
        #%d for the day of the month, %b for the abbreviated month name, and %y for the year as two digits.
        #%-d for the day of the month without the leading zero.
        today_str = today.strftime("%#d-%b-%y")
        today_str_space = today.strftime("%#d %b %y")
        tomorrow = (date.today() + timedelta(1))
        tomorrow_str = tomorrow.strftime("%d-%b-%y")
        tomorrow_str_space = tomorrow.strftime("%d %b %y")
        for row in reader:
            if row["Session Date"] == today_str or row["Session Date"] == today_str_space:
                first_name = row["Tutor First Name"] if "Tutor First Name" in row else "N/A"
                last_name = row["Tutor Last Name"] if "Tutor Last Name" in row else "N/A"
                tutor_names.append(f"{first_name} {last_name}")
                tmp = row["Session Time"].replace(" (GMT+2)", "")
                time.append(tmp)
                tutor_email.append(row["Tutor Email"])
                tutees.append(row["Tutee Name"])
                dates.append(row["Session Date"])
            else:
                #print(row["Session Date"])
                #print(today_str)
                continue
    return tutor_names,time,tutor_email,tutees,dates

tutors_list,time,tutors_email,tutees,dates = extract_data()
# print(tutors_list)
# print(time)
# print(tutors_email)
def convert_time():
    time_num = []
    for time_string in time:
        # Convert the time string to a datetime object
        time_obj = datetime.strptime(time_string, "%I:%M%p")
        # Extract the hour and minute as integers
        # time is converted in 24-hour format 
        hour = time_obj.hour
        minute = time_obj.minute
        #print(str(hour) + " " + str(minute))
        tmp = hour + (minute / 60)
        time_num.append(tmp)
    # print(time_num)
    return time_num
time_num = convert_time()
tutors_schedule = [{"name": k, "time": v, "Tutee": z, "date": x} for k, v, z, x in zip(tutors_list, time_num, tutees, dates)]
#print(tutors_schedule)
conflicts = []
Appointments = []
# 6 arrays each one corresponds to a room. each array is of size 96 which is number of quarter hours in a day.
room_arrays = [[0] * 96 for i in range(6)]

def assign_room(tutor_obj):
    #print(len(tutors_schedule))
    rooms = ["P001", "P002", "G014", "G015", "G010", "G011"]
    #the algorithm:
    start_index = tutor_obj['time'] / 0.25
    
    end_index = start_index + 3
    start_index = int(start_index)
    
    end_index = int(end_index)
    # i is room index
    i = 0
    minutes = '{:02d}'.format(int((start_index % 4)*15))
    if start_index/4 < 12:
        hours = '{:02d}'.format(int(start_index/4))
        TIME = f"{hours}:{minutes} AM"
    else:
        hours = '{:02d}'.format(int((start_index/4)-12))
        TIME = f"{hours}:{minutes} PM"
    while room_arrays[i][start_index] == 1 or room_arrays[i][end_index] == 1:
        i+=1
        if(i == 6):
            #conflicts.append({'name' : tutor_obj['name'] , 'time' : TIME})
            app = {'tutor': tutor_obj['name'] , 'Room': "" , 'time': TIME, 'start index': start_index ,'date': tutor_obj['date'] , 'tutee': tutor_obj['Tutee']}
            Appointments.append(app)
            break
    if i < 6:
        for j in range(start_index,end_index):
            room_arrays[i][j] = 1
        app = {'tutor': tutor_obj['name'] , 'Room': rooms[i] , 'time': TIME, 'start index': start_index ,'date': tutor_obj['date'] , 'tutee': tutor_obj['Tutee']}
        Appointments.append(app)
        #print(app)
        
def sort_Appointments():
    Appointments.sort(key=lambda item:item['start index'], reverse=False)

def pick_tutors():
    for tutor in tutors_schedule:
        cur_name = tutor['name']
        assign_room(tutor)
        tutors_schedule.remove(tutor)
        for Tutor in tutors_schedule:
            if cur_name == Tutor['name']:
                assign_room(Tutor)
                tutors_schedule.remove(Tutor)


def print_Appointments():
    print("Appointments: ")
    for app in Appointments:
        print(app)
    print("Conflicts: ")
    for conflict in conflicts:
        print(conflict)
while tutors_schedule:
    pick_tutors()
#print_Appointments()


def check(html_content , to):
    if(to == 'lib'):
        with open("mail to library.html", "w") as file:
            file.write(html_content)
    else:
        with open("mail to tutors.html", "w") as file:
                file.write(html_content)  

def make_table(to):
    # Set up Jinja2 environment and load HTML template
    env = Environment(loader=FileSystemLoader('.'))
    if(to == 'lib'):
        template = env.get_template('table_template_lib.html')
    else:
        template = env.get_template('table_template.html')
    # Render HTML template with appointment data
    html = template.render(appointments=Appointments)  
    return html

def send_email_tutors(Check):
    global Appointments
    sort_Appointments()
    # Define email credentials and recipient
    sender_email = private.email
    #receiver_email = ["kero678.kk@gmail.com","Amira.kamal@aucegypt.edu"]
    receiver_email = "kero678.kk@gmail.com"
    #receiver_email = "Amira.kamal@aucegypt.edu"
    #tutors_email_mock = ["kero678.kk@gmail.com","norhan_soliman@aucegypt.edu"]
    password = private.password
    html = make_table('tutors')
    if Check:
        check(html,'tutors')
        return
    message =  MIMEMultipart()
    message['Subject'] = "Assigned Rooms for Today's Tutoring Sessions"
    message['From'] = sender_email
    #message['To'] = ", ".join(tutors_email_mock)
    message['To'] = receiver_email

    html = MIMEText(html, 'html')
    
    message.attach(html)
    # Set up the SMTP server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, password)
    #server.sendmail(sender_email, tutors_email_mock, message.as_string())
    server.sendmail(sender_email, receiver_email, message.as_string())
    server.quit()

def send_email_lib(Check):
    sender_email = private.email
    password = private.password

    #to = "bservices@aucegypt.edu"
    to = "kero678.kk@gmail.com"
    #cc = ["dinah@aucegypt.edu",  "fady.michel@aucegypt.edu"]
    cc = "kero678.kk@gmail.com"
    Msg = MIMEMultipart()
    Msg['From'] = sender_email
    Msg['To'] = to
    Msg['Subject'] = "Today’s Peer Tutoring Sessions"
    Msg['Cc'] = cc
    html_lib = html = make_table('lib')
    if Check:
        check(html_lib,'lib')
        return
    html_lib = MIMEText(html_lib, 'html')
    Msg.attach(html_lib)
    # Set up the SMTP server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, password)
    #server.sendmail(sender_email, tutors_email_mock, Msg.as_string())
    server.sendmail(sender_email, to, Msg.as_string())
    server.quit()

send_email_tutors(Check)
send_email_lib(Check)



