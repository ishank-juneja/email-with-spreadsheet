## Send customized emails with a spreadsheet
Send customized emails to a collection of contacts drawn from a spread sheet 
using a simple python script.

## Features

- Send customized emails using easy to use template files as opposed to a long, cumbersome and difficult to read string that is 
part of the code itself. 
- Add PDF customized attachments with file names matching a simpe template
- Very little setup required - no local email clients or software needed  
- Ideal if the user prefers customizable stand alone scripts as opposed 
to finished browser extension software

## Dependencies
The python script relies on the following dependencies

    smtplib
    string
    email
    csv
    glob

## Usage
Create a file called `credentials.py` in the same directory as `mail-students.py`. Add 2 lines to the file stating the the 
email address you want to use and its password.  
    
    MY_ADDRESS = 'manmohan@iitb.ac.in'
    PASSWORD = 'isThisStrongEnough?'

Create message template files for the message body to be shared with recipients. 
For instance, to email every student in a class informing them of their test scores, the template message would look like,

    Dear ${PERSON_NAME},

    Your marks in Quiz 01 are ${TOTAL_SCORE} out of 14 (Without Penalty).
    Late penalty will correspond to ${LATE_PENALTY} minutes.
    Please find attached your graded answer script(s).
    
    Regards,
    Teaching Assistant Team
    Email: narendra_modi@gmail.com with Prof Manmohan Singh (manmohan@iitb.ac.in) on CC if there are any discrepancies

Use your message template at the top of the main function

    message_template1 = read_template('messages/message_example.txt')
    
- Set up a local SMTP server using the command, the port number for the outgoing SMTP server is usually 587. 
For the details visit your email providers webpage. For instance, for my university [IIT Bombay](http://www.iitb.ac.in/), the iformatio
was [available here](https://www.cc.iitb.ac.in/page/configurewebmail). 

     s = smtplib.SMTP(host='email-host-SMTP-server-name', port=PORT_NUMBER)

Save your spreadsheet containing email-addresses and other information as a `csv` file. For example the below CSV file `example.csv`. 
Keep the fields of the spread-sheet you want to use at the back of your head, these will be needed next.  

    Sl. No.,Roll No.,Name,Exam 01,Time Pen.
    1,160110RF5,Narendra Modi,7,0
    2,214S700R9,Indira Gandhi,5,12
    3,16D070012,Ishank Juneja,6,6
    4,16D070QT2,Jwaharlal Nehru,8,0

Specify the spread sheet file location in-

    # Loop over all the students
    with open('/home/ishank/Desktop/EE229_student_marks/EE229-StudentList-29Sep.csv') as csv_file:
    
Now we must provide the script information about the various placeholders - {PERSON_NAME} etc. - used in the template message file.
Do so by matching spreadsheet column headers with the place-holder names.

    message = message_template2.substitute(PERSON_NAME=row['Name'], TOTAL_SCORE=row['Exam 01'],
                                                       LATE_PENALTY=row['Time Pen.'])
                                                       
Specify From, to, CC etc,

    msg['From'] = MY_ADDRESS
    msg['To'] = row['Roll No.'] + "@iitb.ac.in"
    msg['Subject'] = "EE XYZ course Quiz 01 Marks"
    msg['CC'] = "manmohan@iitb.ac.in"

If PDF files, for instance graded answer scripts are to be attached, specify their folder locations instead of-

    pdf_names = glob.glob('/home/ishank/Desktop/EE229_student_marks/Exam01_Q1_2_3_6_7_8_9_compressed/{0}*'.format(row['Roll No.'])) + \
            glob.glob('/home/ishank/Desktop/EE229_student_marks/Exam01_Q4_5_compressed/{0}*'.format(row['Roll No.']))
The `glob` simply matches the file names to a template with the `*` acting like a placeholder for 0 or more characters.

A `try-except` block checks if sending an email was successful from the SMTP servers end, if not, it prints a message. For instance, the email might 
fail if the PDF (or other) attachments are too large for the service. This does not guarantee the intended recipient gets the message since their email address might still be incorrect. 

Another repository of interest might be: [PDF Tools](https://github.com/ishank-juneja/pdf-tools) which has tools for annotating, merging, splitting and compressing PDF files.
 
### References:
[1] https://www.freecodecamp.org/news/send-emails-using-code-4fcea9df63f/
