import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from credentials import MY_ADDRESS, PASSWORD
import csv
import glob


# Returns a Template object comprising the contents of the
# file specified by filename.
# taken from https://www.freecodecamp.org/news/send-emails-using-code-4fcea9df63f/
def read_template(filename):
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def main():
    message_template1 = read_template('messages/message_not_late.txt')
    message_template2 = read_template('messages/message_late.txt')
    message_template3 = read_template('messages/message_absent.txt')
    # set up the SMTP server
    s = smtplib.SMTP(host='smtp-auth.iitb.ac.in', port=587)
    s.starttls()
    s.login(MY_ADDRESS, PASSWORD)
    # Loop over all the students
    with open('/home/ishank/Desktop/EE229_student_marks/EE229-StudentList-29Sep.csv') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        for row in csv_reader:
            # Check if it is an empty row without name or roll number (ignore)
            if not (any(c.isalpha() for c in row['Name']) and row['Roll No.'].isalnum()):
                continue
            # if(row['Name'] in glob.glob('*,*')):
            #     continue
            # Create a new message object for every row of the CSV file
            msg = MIMEMultipart()
            # Add in the actual person data to the appropriate message template
            if row['Time Pen.'] == "0":
                message = message_template1.substitute(PERSON_NAME=row['Name'], TOTAL_SCORE=row['Exam 01'])
            elif row['Exam 01'] == "Ab":
                message = message_template3.substitute(PERSON_NAME=row['Name'])
            else:
                message = message_template2.substitute(PERSON_NAME=row['Name'], TOTAL_SCORE=row['Exam 01'],
                                                       LATE_PENALTY=row['Time Pen.'])
            # Prints out the message body for our sake
            # print(message)
            # setup the parameters of the message
            msg['From'] = MY_ADDRESS
            msg['To'] = row['Roll No.'] + "@iitb.ac.in"
            msg['Subject'] = "EE 229 Quiz 01 Marks"
            # msg['CC'] = "manmohan@iitb.ac.in"
            # add in the message body
            msg.attach(MIMEText(message, 'plain'))
            # Collection of PDF file names
            pdf_names = glob.glob('/home/ishank/Desktop/EE229_student_marks/Exam01_Q1_2_3_6_7_8_9_compressed/{0}*'.format(row['Roll No.'])) + \
            glob.glob('/home/ishank/Desktop/EE229_student_marks/Exam01_Q4_5_compressed/{0}*'.format(row['Roll No.']))
            # Add attachment files if any
            for pdf_file_name in pdf_names:
                # Attach the pdf to the msg going by e-mail
                with open(pdf_file_name, "rb") as f:
                    # attach = email.mime.application.MIMEApplication(f.read(),_subtype="pdf")
                    attach = MIMEApplication(f.read(), _subtype="pdf")
                attach.add_header('Content-Disposition', 'attachment', filename=str(pdf_file_name.split('/')[-1]))
                msg.attach(attach)
            # Send the message via the SMTP server set up earlier
            try:
                s.send_message(msg)
            # Sending the message may fail for various reasons
            # All the reasons have been encapsulated in this exception
            except Exception:
                print("An exception occurred for the message corresponding to Roll Number {0}".format(row['Roll No.']))
                continue
            del msg
    # Terminate the SMTP session and close the connection
    s.quit()


if __name__ == '__main__':
    main()
