import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from credentials import MY_ADDRESS, PASSWORD
import csv
import glob


def read_template(filename):
    """
    Returns a Template object comprising the contents of the
    file specified by filename.
    taken from https://www.freecodecamp.org/news/send-emails-using-code-4fcea9df63f/
    """
    with open(filename, 'r', encoding='utf-8') as template_file:
        template_file_content = template_file.read()
    return Template(template_file_content)


def main():
    message_template = read_template('message.txt')
    # set up the SMTP server
    s = smtplib.SMTP(host='smtp-auth.iitb.ac.in', port=587)
    s.starttls()
    s.login(MY_ADDRESS, PASSWORD)
    # Loop over all the students
    with open('EE229-StudentList-29Sep.csv') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        for row in csv_reader:
            # Check if it is an empty row without name or roll number
            if not (any(c.isalpha() for c in row['Name']) and row['Roll No.'].isalnum()):
                continue
            # if(row['Name'] in glob.glob('*,*')):
            #     continue
            # Create a new message object for every row of the CSV file
            msg = MIMEMultipart()
            # Add in the actual person data to the message template
            message = message_template.substitute(PERSON_NAME=row['Name'],
                                                  TOTAL_SCORE=row['Exam 01'], LATE_PENALTY=row['Time Pen.'],
                                                  NET_SCORE=row['Net 01'])
            # Prints out the message body for our sake
            print(message)
            # setup the parameters of the message
            msg['From'] = MY_ADDRESS
            msg['To'] = row['Roll No.'] + "@iitb.ac.in"
            msg['Subject'] = "EE 229 Quiz 01 Marks"
            # msg['CC'] = "animesh@ee.iitb.ac.in"
            # add in the message body
            msg.attach(MIMEText(message, 'plain'))
            # Add attachment files if any
            for pdf_file_name in glob.glob('{0}*'.format(row['Roll No.'])):
                # Attach the pdf to the msg going by e-mail
                with open(pdf_file_name, "rb") as f:
                    # attach = email.mime.application.MIMEApplication(f.read(),_subtype="pdf")
                    attach = MIMEApplication(f.read(), _subtype="pdf")
                attach.add_header('Content-Disposition', 'attachment', filename=str(pdf_file_name))
                msg.attach(attach)
            # send the message via the server set up earlier.
            # s.send_message(msg)
            del msg
    # Terminate the SMTP session and close the connection
    s.quit()


if __name__ == '__main__':
    main()
