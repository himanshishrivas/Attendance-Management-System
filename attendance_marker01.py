import face_recognition
import cv2
from openpyxl import Workbook
import datetime


# Get a reference to webcam #0 (the default one)
video_capture = cv2.VideoCapture(0)

# Create a woorksheet
book = Workbook()
sheet = book.active

# Load images.

image_1 = face_recognition.load_image_file('valid_images/1.jpg')
image_1_face_encoding = face_recognition.face_encodings(image_1)[0]

image_2 = face_recognition.load_image_file('valid_images/2.jpg')
image_2_face_encoding = face_recognition.face_encodings(image_2)[0]

image_3 = face_recognition.load_image_file('valid_images/3.jpg')
image_3_face_encoding = face_recognition.face_encodings(image_3)[0]

image_4 = face_recognition.load_image_file('valid_images/4.jpg')
image_4_face_encoding = face_recognition.face_encodings(image_4)[0]

image_5 = face_recognition.load_image_file('valid_images/5.jpg')
image_5_face_encoding = face_recognition.face_encodings(image_5)[0]

# Create arrays of known face encodings and their names
known_face_encodings = [

    image_1_face_encoding,
    image_2_face_encoding,
    image_3_face_encoding,
    image_4_face_encoding,
    image_5_face_encoding

]
known_face_names = [

    "1",
    "2",
    "3",
    "4",
    "5"

]


# Initialize some variables
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True

# Load present date and time
now = datetime.datetime.now()
today = now.day
month = now.month

while True:
    # Grab a single frame of video
    ret, frame = video_capture.read()

    # Convert the image from BGR color (which OpenCV uses) to RGB color (which face_recognition uses)
    rgb_frame = frame[:, :, ::-1]

    # Find all the faces and face enqcodings in the frame of video
    face_locations = face_recognition.face_locations(rgb_frame)
    face_encodings = face_recognition.face_encodings(rgb_frame, face_locations)

    # Loop through each face in this frame of video
    for (top, right, bottom, left), face_encoding in zip(face_locations, face_encodings):
        # See if the face is a match for the known face(s)
        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)

        name = "Allien"

        # If a match was found in known_face_encodings, just use the first one.
        # if True in matches:
        #     first_match_index = matches.index(True)
        #     name = known_face_names[first_match_index]



        '''def fill_attendance():
            ts = time.time()
            Date = datetime.datetime.fromtimestamp(ts).strftime('%Y_%m_%d')
            timeStamp = datetime.datetime.fromtimestamp(ts).strftime('%H:%M:%S')
            Time = datetime.datetime.fromtimestamp(ts).strftime('%H:%M:%S')
            Hour, Minute, Second = timeStamp.split(":")
            ####Creatting csv of attendance

            ##Create table for Attendance
            date_for_DB = datetime.datetime.fromtimestamp(ts).strftime('%Y_%m_%d')
            global subb
            subb = SUB_ENTRY.get()
            DB_table_name = str(subb + "_" + Date + "_Time_" + Hour + "_" + Minute + "_" + Second)

            import pymysql.connections

            ###Connect to the database
            try:
                global cursor
                connection = pymysql.connect(host='localhost', user='root', password='', db='Face_reco_fill')
                cursor = connection.cursor()
            except Exception as e:
                print(e)

            sql = "CREATE TABLE " + DB_Table_name + """
                            (ID INT NOT NULL AUTO_INCREMENT,
                             ENROLLMENT varchar(100) NOT NULL,
                             NAME VARCHAR(50) NOT NULL,
                             DATE VARCHAR(20) NOT NULL,
                             TIME VARCHAR(20) NOT NULL,
                                 PRIMARY KEY (ID)
                                 );
                            """
            ####Now enter attendance in Database
            insert_data = "INSERT INTO " + DB_Table_name + " (ID,ENROLLMENT,NAME,DATE,TIME) VALUES (0, %s, %s, %s,%s)"
            VALUES = (str(Id), str(aa), str(date), str(timeStamp))
            try:
                cursor.execute(sql)  ##for create a table
                cursor.execute(insert_data, VALUES)  ##For insert data into table
            except Exception as ex:
                print(ex)  #

            M = 'Attendance filled Successfully'
            Notifica.configure(text=M, bg="Green", fg="white", width=33, font=('times', 15, 'bold'))
            Notifica.place(x=20, y=250)




        def create_csv():
            import csv
            cursor.execute("select * from " + DB_table_name + ";")
            csv_name = 'C:/Users/kusha/PycharmProjects/Attendace managemnt system/Attendance/Manually Attendance/' + DB_table_name + '.csv'
            with open(csv_name, "w") as csv_file:
                csv_writer = csv.writer(csv_file)
                csv_writer.writerow([i[0] for i in cursor.description])  # write headers
                csv_writer.writerows(cursor)
                O = "CSV created Successfully"
                Notifi.configure(text=O, bg="Green", fg="white", width=33, font=('times', 19, 'bold'))
                Notifi.place(x=180, y=380)'''

        # If a match was found in known_face_encodings, just use the first one.

        if True in matches:
            first_match_index = matches.index(True)
            name = known_face_names[first_match_index]
            # Assign attendance
            if int(name) in range(1, 61):
                sheet.cell(row=int(name), column=int(today)).value = "Present"

            else:
                pass

            face_names.append(name)

        process_this_frame = not process_this_frame

        # Or instead, use the known face with the smallest distance to the new face
        '''face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
        best_match_index = np.argmin(face_distances)
        if matches[best_match_index]:
            name = known_face_names[best_match_index] 
            '''

        # Draw a box around the face
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)

        # Draw a label with a name below the face
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

    # Display the resulting image
    cv2.imshow('Video', frame)

    # Save Woorksheet as present month
    book.save(str(month) + '.xls')

    # Hit 'q' on the keyboard to quit!
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

# Release handle to the webcam
video_capture.release()
cv2.destroyAllWindows()