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
import os
import re


#jinja2 library generate the HTML code dynamically using HTML Templates while
# email.mime takes the generated HTML code and render it.
#print(Check)
#Check = True

def rename():
    ## Delete the old file first:
    # file path
    file_path = "C:\\Users\\kero6\Desktop\\daily report.csv"
    # delete file
    os.remove(file_path)
    ## rename the new one:
    # directory containing the files
    dir_path = "C:\\Users\\kero6\\Desktop"
    # the first part of the name
    name_prefix  = 'admin_download_meetings_detailed_AUC Peer Tutoring'
    new_name = 'daily report.csv'
    for file_name in os.listdir(dir_path):
        # check if file name starts with the given prefix
        if file_name.startswith(name_prefix):
            # rename file
            os.rename(os.path.join(dir_path, file_name), os.path.join(dir_path, new_name))

def dbg(i):
    print(f"here {i}")
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
            #if row["Session Date"] == today_str or row["Session Date"] == today_str_space:
                first_name = row["Tutor First Name"] if "Tutor First Name" in row else "N/A"
                last_name = row["Tutor Last Name"] if "Tutor Last Name" in row else "N/A"
                tutor_names.append(f"{first_name} {last_name}")
                tmp = row["Session Time"].replace(" (GMT+2)", "")
                time.append(tmp)
                tutor_email.append(row["Tutor Email"])
                tutees.append(row["Tutee Name"])
                dates.append(row["Session Date"])
            #else:
                #print(row["Session Date"])
                #print(today_str)
                continue
    return tutor_names,time,tutor_email,tutees,dates
rename()
tutors_list,time_list,tutors_email,tutees,dates = extract_data()
# print(tutors_list)
# print(time_list)
# print(tutors_email)
# print(dates)
special_Appointments = []
Special = False
def custom_sessions():
    global time_list, Special
    val = input("Please Type:\n 1 to add a session to the list\n 2 to to add a session(s) to be sent in separate mail\n 3 to proceed\n")
    while(val != '3'):
        if val == '1' or val == '2':
            #dbg(0)
            Tutor_name = input("Please enter Tutor Name: ")
            Tutee_name = input("Please enter Tutee Name: ")
            Smail = input("Please enter tutor mail: ")
            Sdate = date.today().strftime("%#d %b %y")
            if val == '1':
                # add the session to the lists.
                #dbg(1)
                Stime = input("Please enter time as 1:30pm (please use the exact same format): ")
                tutors_list.append(Tutor_name)
                tutees.append(Tutee_name)
                time_list.append(Stime)
                dates.append(Sdate)
                tutors_email.append(Smail)
            elif val == '2':
                #dbg(2)
                Special = True
                Stime = input("Please enter time as 01:30 PM: ")
                Sroom = input("Please enter the room: ")
                app = {'tutor': Tutor_name , 'Room': Sroom , 'time': Stime, 'start index': 0 ,'date': Sdate , 'tutee': Tutee_name}
                special_Appointments.append(app)
        val = input("Please Type:\n 1 to add a session to the list\n 2 to to add a session(s) to be sent in separate mail\n 3 to proceed\n")

#TODO
#def generate_date(day,month,year): 


def find_index(tutor): #called inside
    indices = []
    for i in range(len(tutors_list)):
        if tutors_list[i] == tutor:
            indices.append(i)
    return indices

custom_sessions()
#print(special_Appointments)
Check = True
def Check_decision():
    global Check
    Check_txt = input("Please type:\n 1 if you want to generate mock email.\n 0 if you want to proceed the email directly.\n")
    if Check_txt == '1':
        Check = True
    else:
        Check = False


def remove_sessions():
    val = input("Please Type:\n 0 to remove a session form the list\n 1 to proceed\n")
    while(val != '1'):
            print_Appointments(Appointments)
            i = input("Please Type the number of session you want to delete or type x to cancel\n")
            i = i.lower()
            if i != 'x':
                i = int(i)
                i= i-1
                del Appointments[i]
            val = input("Please Type:\n 0 to remove a session form the list\n 1 to proceed\n")



def convert_time():
    time_num = []
    for time_string in time_list:
        # Convert the time string to a datetime object
        time_obj = datetime.strptime(time_string, "%I:%M%p")
        # Extract the hour and minute as integers
        # time is converted in 24-hour format 
        hour = time_obj.hour
        minute = time_obj.minute
        #print(str(hour) + " " + str(minute))
        tmp = hour + (minute / 60)
        time_num.append(tmp)
    #print(time_num)
    return time_num

def making_schedule(): #using the direct lists
    time_num = convert_time()
    tutors_schedule = [{"name": k, "time": v, "Tutee": z, "date": x} for k, v, z, x in zip(tutors_list, time_num, tutees, dates)]
    return tutors_schedule

tutors_schedule = making_schedule()
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
    if start_index == 0:
        TIME = "12:00 AM"
    elif start_index/4 < 12:
        hours = '{:02d}'.format(int(start_index/4))
        TIME = f"{hours}:{minutes} AM"
    elif start_index/4 > 12 :
        hours = '{:02d}'.format(int((start_index/4)-12))
        TIME = f"{hours}:{minutes} PM"
    elif start_index/4 == 12:
        TIME = "12:00 PM"
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


def print_Appointments(Appointments):
    print("Appointments: ")
    i = 1
    for app in Appointments:
        print(f"{i}- tutor: {app['tutor']}, tutee: {app['tutee']}, time: {app['time']}")
        i = i+1

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

def make_table(to, Special):
    # Set up Jinja2 environment and load HTML template
    env = Environment(loader=FileSystemLoader('.'))
    if(to == 'lib'):
        template = env.get_template('table_template_lib.html')
    else:
        template = env.get_template('table_template.html')
    # Render HTML template with appointment data
    if Special:
        #dbg(1)
        html = template.render(appointments=special_Appointments)  
    else:
        html = template.render(appointments=Appointments)
    return html

sort_Appointments()
remove_sessions()
Check_decision()

def send_email_tutors(Check):
    global Appointments, special_Appointments
    global Special
    # Define email credentials and recipient
    
    sender_email = private.email
    #receiver_email = "Amira.kamal@aucegypt.edu"
    tutors_email_mock = ["kero678.kk@gmail.com","norhan_soliman@aucegypt.edu"]
    to = "kero678.kk@gmail.com"
    password = private.password
    html = make_table('tutors', Special)
    if Check:
        check(html,'tutors')
        return
    message =  MIMEMultipart()
    message['Subject'] = "Assigned Rooms for Today's Tutoring Sessions"
    message['From'] = sender_email
    #message['To'] = ", ".join(tutors_email_mock)
    message['To'] = to
    html = MIMEText(html, 'html')
    message.attach(html)
    # Set up the SMTP server
    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(sender_email, password)
    server.sendmail(sender_email, to, message.as_string())
    #server.sendmail(sender_email, tutors_email_mock, message.as_string())
    server.quit()

def send_email_lib(Check):
    sender_email = private.email
    password = private.password
    
    #to = "bservices@aucegypt.edu"
    to = "kero678.kk@gmail.com"
    #to = "nm.hashem@aucegypt.edu"
    #cc = ["dinah@aucegypt.edu",  "fady.michel@aucegypt.edu"]
    cc = ["kero678.kk@gmail.com" , "kero678.kk@gmail.com"]
    Msg = MIMEMultipart()
    Msg['From'] = sender_email
    Msg['To'] = to
    Msg['Subject'] = "Today’s Peer Tutoring Sessions"
    Msg['Cc'] = ", ".join(cc)
    html_lib = make_table('lib', Special)
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

#TODO: enable user to add or remove a session after the HTML file is generated.
#TODO: the program stores the added sessions in a text file so that if the user checks the email format he should not reenter the added sessions.