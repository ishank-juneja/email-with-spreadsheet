import csv


def main():
    # Loop over all the students
    with open('EE229-StudentList-29Sep.csv') as csv_file:
        csv_reader = csv.DictReader(csv_file, delimiter=',')
        for row in csv_reader:
            # Add in the actual person data to the message template
            print(row['Name'], row['Exam 01'], row['Time Pen.'], row['Net 01'])
            
if __name__ == '__main__':
    main()
