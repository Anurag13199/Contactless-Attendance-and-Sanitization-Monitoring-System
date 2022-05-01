import PySimpleGUI as sg
import cv2
import numpy as np
import face_recognition
import pickle
from openpyxl import Workbook
import datetime
import yagmail


def main():

    sg.theme('dark grey 9')

    # define the window layout
    layout = [[sg.Text('Employee Attendance and Sanitization Management', size=(50, 1), justification='center', font='Helvetica 20')],
              [sg.Text('Enter Name', size=(20, 1), font='Helvetica 18') ],
              [sg.Input(key='name', size=(20, 1))],
              [sg.Text('Enter Email ID', size=(20, 1), font='Helvetica 18') ],
              [sg.Input(key='em_id', size=(20, 1))],
              [sg.Text('Enter Employee ID', size=(20, 1), font='Helvetica 18') ],
              [sg.Input(key='ref_id', size=(20, 1))],
              [sg.Button('Add Employee', size=(13, 1), font='Helvetica 14'),
               sg.Button('Capture', size=(10,1), font='Helvetica 14', visible=False),
               sg.Button('Start', size=(10, 1), font='Any 14'),
               sg.Button('Stop', size=(10, 1), font='Any 14'),
               sg.Button('Exit', size=(10, 1), font='Helvetica 14'), ]]

    # create the window and show it without the plot
    window = sg.Window('Demo Application - OpenCV Integration',
                       layout, location=(800, 400))

    while True:
        event, values = window.read(timeout=20)
        if event == 'Exit' or event == sg.WIN_CLOSED:
            return

        elif event == 'Add Employee':
            if values['name'] and values['ref_id'] and values['em_id']:
                window['Capture'].update(visible=True)
                name=values['name']
                ref_id=values['ref_id']
                em_id=values['em_id']

                try:
                    f=open("ref_name.pkl","rb")
                    ref_dictt=pickle.load(f)
                    f.close()
                except:
                    ref_dictt={}
                ref_dictt[ref_id]= {'name': name, 'email': em_id }
                f=open("ref_name.pkl","wb")
                pickle.dump(ref_dictt,f)
                f.close()
                try:
                    f=open("ref_embed.pkl","rb")
                    embed_dictt=pickle.load(f)
                    f.close()
                except:
                    embed_dictt={}


                for i in range(5):
                    # key = cv2. waitKey(1)
                    webcam = cv2.VideoCapture(0, cv2.CAP_DSHOW)
                    while True:

                        check, frame = webcam.read()
                        cv2.imshow("Capturing", frame)
                        small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
                        rgb_small_frame = small_frame[:, :, ::-1]
                        # cv2.imshow("Capturing1", rgb_small_frame)


                        event, values = window.read(timeout=20)
                        if event == 'Capture' :
                            face_locations = face_recognition.face_locations(rgb_small_frame)
                            if face_locations != []:
                                face_encoding = face_recognition.face_encodings(frame)[0]
                                if ref_id in embed_dictt:
                                    embed_dictt[ref_id]+=[face_encoding]
                                else:
                                    embed_dictt[ref_id]=[face_encoding]
                                webcam.release()
                                # cv2.waitKey(1)
                                cv2.destroyAllWindows()
                                break
                        elif event == 'Stop':
                            print("Turning off camera.")
                            webcam.release()
                            print("Camera off.")
                            print("Program ended.")
                            cv2.destroyAllWindows()
                            window['Capture'].update(visible=False)
                            break
                    if event == 'Stop':
                        break



                f=open("ref_embed.pkl","wb")
                pickle.dump(embed_dictt,f)
                f.close()
            else:
                sg.popup('Enter Details')




        elif event == 'Start':
            f=open("ref_name.pkl","rb")
            ref_dictt=pickle.load(f)
            f.close()
            f=open("ref_embed.pkl","rb")
            embed_dictt=pickle.load(f)
            f.close()

            yag = yagmail.SMTP('your_email', '****your_password****')

            known_face_encodings = []
            known_face_names = []
            for ref_id , embed_list in embed_dictt.items():
                for my_embed in embed_list:
                    known_face_encodings +=[my_embed]
                    known_face_names += [ref_id]

            video_capture = cv2.VideoCapture(0)
            face_locations = []
            face_encodings = []
            face_names = []
            process_this_frame = True
            wb = Workbook()
            attendance = wb.active
            attendance.title = "Attendance "
            deets = ["Name", "ID"]
            attendance.append(deets)
            while True  :

                ret, frame = video_capture.read()
                small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
                rgb_small_frame = small_frame[:, :, ::-1]
                if process_this_frame:
                    face_locations = face_recognition.face_locations(rgb_small_frame)
                    face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)
                    face_names = []
                    user_attd = []
                    for face_encoding in face_encodings:
                        matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
                        name = "Unknown"
                        face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
                        best_match_index = np.argmin(face_distances)
                        if matches[best_match_index]:
                            name = known_face_names[best_match_index]
                        face_names.append(name)
                process_this_frame = not process_this_frame
                for (top_s, right, bottom, left), name in zip(face_locations, face_names):
                    top_s *= 4
                    right *= 4
                    bottom *= 4
                    left *= 4
                    cv2.rectangle(frame, (left, top_s), (right, bottom), (0, 0, 255), 2)
                    cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
                    font = cv2.FONT_HERSHEY_DUPLEX
                    cv2.putText(frame, ref_dictt[name]['name'] + '  ' + ref_dictt[name]['email'], (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)
                    user_attd = ( ref_dictt[name]['name'], ref_dictt[name]['email'] )
                    att_list = []
                    for row in attendance.iter_rows(min_row = 1, max_col = 2, values_only=True):
                        att_list.append(row)
                    if user_attd not in att_list:
                        # print(type(user_attd))
                        attendance.append(user_attd)
                        contents = [
                            "Your attendance for the day "+str(datetime.datetime.now())+" has been recorded"
                            ]
                        yag.send(ref_dictt[name]['email'], 'Attendance '+str(datetime.datetime.now()), contents)
                        print("true")


                font = cv2.FONT_HERSHEY_DUPLEX
                # service = get_service()
                # msg = create_message('to_email', 'from_email', 'Attendance','Your Attendance is recorded. ',None)
                # send_message(service,'me', msg)
                cv2.imshow('Video', frame)
                event, values = window.read(timeout=20)
                if event == 'Stop':
                    break
            # wb.save('attendance'+ str(datetime.datetime.now()) + '.xlsx')
            wb.save('attendance.xlsx')
            video_capture.release()
            cv2.destroyAllWindows()

        # if recording:
        #     ret, frame = cap.read()
        #     imgbytes = cv2.imencode('.png', frame)[1].tobytes()  # ditto
        #     window['image'].update(data=imgbytes)


main()
