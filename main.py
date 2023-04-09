'''
Final. Everything works correctly.

1. Works with only one PowerPoint and Word application at a time.
2. Prefer to be on the workspace in either app.
3. Prefer closing the application with gesture instead of Close button. 
'''

import cv2
import numpy as np
import time as t
import mediapipe as mp
from tensorflow.keras.models import load_model
from pywinauto.application import Application
from pywinauto.keyboard import send_keys
from pywinauto.mouse import click
import pygetwindow as gw
from math import hypot
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

def scan_gesture(controls_list , aos , app_type = None):

    try:

        if(aos == 'appli'):
            feed = 12*['a']
        elif(aos == 'speci'):
            feed = 10*['a']   
        elif(aos == 'slide'):
            feed = 15*['a']   

    except:

        feed = 13*['a']


    while True:
    # Read each frame from the webcam

        _, frame = cap.read()

        x, y, c = frame.shape

        # Flip the frame vertically
        frame = cv2.flip(frame, 1)
        framergb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        # Get hand landmark prediction
        result = hands.process(framergb)
        
        control = '_'

        # post process the result
        if result.multi_hand_landmarks:
            landmarks = []
            for handslms in result.multi_hand_landmarks:
                for lm in handslms.landmark:
                    # print(id, lm)
                    lmx = int(lm.x * x)
                    lmy = int(lm.y * y)

                    landmarks.append([lmx, lmy])

                # Drawing landmarks on frames
                mpDraw.draw_landmarks(frame, handslms, mpHands.HAND_CONNECTIONS)

                # Predict gesture
                prediction = model.predict([landmarks])
                classID = np.argmax(prediction)
                control = controls_list[classID]


        # show the prediction on the frame
        if(control == '_' and aos == 'appli'):
            cam_screen_show(frame, 'Main Menu')

        elif(control == '_' and aos == 'speci'):
            if(app_type == 'PPT'):
                cam_screen_show(frame, 'PPT Sub Menu')  
            elif(app_type == 'DOC'):  
                cam_screen_show(frame, 'DOC Sub Menu')
            elif(app_type == 'VOL'):
                cam_screen_show(frame, 'Volume Control')

        elif(control == '_' and aos == 'slide'):
            cam_screen_show(frame , 'Presenting...')
        
        else:
            cam_screen_show(frame , control)


        feed.pop(0)
        feed.append(control)

        feed_set = set(feed)

        # return the hand class if feed_set==1
        if len(feed_set)==1 and feed[9] in available_controls:
            return feed[9]



def cam_screen_show(frame , text_var):
        
    cv2.putText(frame, text_var, (10, 70), cv2.FONT_HERSHEY_SIMPLEX, 2, (0,0,255), 4, cv2.LINE_AA)

    cv2.namedWindow("Output", cv2.WINDOW_NORMAL)
    cv2.resizeWindow("Output", 300, 200)
    cv2.imshow("Output", frame)
    cv2.setWindowProperty("Output", cv2.WND_PROP_TOPMOST, 1)
    cv2.moveWindow("Output", -15, 810)
    cv2.waitKey(1)



def extract_app_name(app_type):

    windows_list = gw.getAllTitles()

    for app_name in windows_list:
    
        if app_type == 'PPT' and '- PowerPoint' in app_name:
            return app_name, True

        elif app_type == 'DOC' and '- Word' in app_name:
            return app_name, True

    return '', False



def clear_stuff():

    try:
        t.sleep(1)
        send_keys('{VK_ESCAPE}')

        t.sleep(0.25)
        click(button='left', coords=(330, 10))

        t.sleep(0.3)
        send_keys('{VK_ESCAPE}')

        t.sleep(0.2)
        send_keys('{VK_ESCAPE}')        

    except:
        pass



def app_controls(app , app_name, app_active , app_window , app_type , alt_app_active):

    try:

        if(app_type == 'PPT'):
            specific_controls = ppt_controls
        elif(app_type == 'DOC'):
            specific_controls = doc_controls

        current_app_name = app_name


        while True:

            t.sleep(2)

            in_app_control = scan_gesture(specific_controls, 'speci', app_type)

            if(app.is_process_running() == True):

                if current_app_name not in gw.getAllTitles():
                    app_name , app_active = extract_app_name(app_type) 
                    app = Application(backend="uia").connect(title = app_name, timeout = 20)   
                    current_app_name = app_name  
                    app_window = gw.getWindowsWithTitle(app_name)[0]                    

            else:

                return  None, '' , False , False , None


            if(in_app_control not in ['Back to Main Menu','Switch to PowerPoint','Switch to Word']):
                if not(app_window.isMaximized):
                    app_window.maximize()
                app_window.activate()



            if(in_app_control == 'Save As'):
                save_as(app , app_name)
                continue


            elif(in_app_control == 'Print'):
                print_(app , app_name)
                continue


            elif(in_app_control == 'Open Pinned Presentation' or in_app_control == 'Open Pinned Document'):
                
                pinned_file_name = open_pinned_file(app , app_name , app_type)

                if(pinned_file_name == current_app_name):                    
                    continue
                    
                else:
                    if(app_type == 'PPT'):                    
                        return None , pinned_file_name , False , False , 'Open/Connect PowerPoint'
                        
                    elif(app_type == 'DOC'):                        
                        return None , pinned_file_name , False , False , 'Open/Connect Word'
                    
                    else:
                        continue


            elif(in_app_control == 'Switch to PowerPoint'):

                if(app_type == 'PPT'):

                    if not(app_window.isMaximized):
                        app_window.maximize()
                    app_window.activate()

                elif(app_type == 'DOC'):

                    if alt_app_active == True:
                        return app , app_name , True , True , 'Open/Connect PowerPoint'
        
                else:
                    return None, '', False , False , None


            elif(in_app_control == 'Switch to Word'):

                if(app_type == 'DOC'):

                    if not(app_window.isMaximized):
                        app_window.maximize()
                    app_window.activate()

                elif(app_type == 'PPT'):

                    if alt_app_active == True:
                        return app, app_name , True , True , 'Open/Connect Word'
        
                else:
                    return None, '', False , False , None


            elif(in_app_control == 'Begin Slideshow'):
                slideshow(app , app_name)
                continue


            elif(in_app_control == 'Close PowerPoint' or in_app_control == 'Close Word'):
                
                close(app , app_name)

                t.sleep(1)

                if(app.is_process_running() == True):                    
                    continue

                elif(app.is_process_running() == False):                    
                    return None , '' , False, False , None

                else:
                    continue


            elif(in_app_control == 'Back to Main Menu'):
                break


        return app , app_name, True, True , None


    except Exception as e: 

        if app.is_process_running():
            return app , app_name ,True, True , None
        
        else:
            return  None , '', False , False , None



def save_as(app , app_name):

    try:

        clear_stuff()

        t.sleep(0.3)
        send_keys('%f')

        t.sleep(0.3)
        send_keys('{a}')

        t.sleep(0.3)
        send_keys('{o}')
    
    except:
        clear_stuff()           
        t.sleep(1)     
        return 



def print_(app , app_name):

    try:

        clear_stuff()    

        t.sleep(0.3)
        send_keys('^p')        

    except:
        clear_stuff()    
        t.sleep(1)
        return 



def slideshow(app , app_name):

    try:

        clear_stuff()

        status_bar_wrapper = app[app_name].child_window(title="Status Bar", class_name = 'NetUInetpane').wrapper_object()

        slide_status = status_bar_wrapper.children_texts()[0]

        x = slide_status.find('f')

        total_slides = int(slide_status[x+2:])

        current_slide = 1

        move = ''        

        slideshow_button = app[app_name].child_window(title="From Beginning", control_type="Button").wrapper_object()

        slideshow_button.click_input()

        t.sleep(0.3)
    
        while(move != 'End Slideshow'):

            move = scan_gesture(slideshow_controls , 'slide', 'PPT')

            if(move == 'Next Slide'):     
                send_keys('{VK_RIGHT}')
                current_slide += 1 
                if(current_slide > total_slides):
                    break

            elif(move == 'Previous Slide'):   
                send_keys('{VK_LEFT}')  
                current_slide -= 1         
                if(current_slide == 0):
                    break                   

        t.sleep(0.5)
        send_keys('{VK_ESCAPE}')

    except:
        send_keys('{VK_ESCAPE}')

    t.sleep(2)



def open_pinned_file(app , app_name , app_type):

    try:

        clear_stuff()

        t.sleep(0.2)
        send_keys('%f')
        t.sleep(0.3)
        send_keys('{h}')
        t.sleep(0.3)
        send_keys('{y}')
        t.sleep(0.3)
        send_keys('{d}')
        t.sleep(0.3)
        
        pinned_list_wrap = app[app_name].child_window(title="Pinned", class_name = 'NetUIListView').wrapper_object()

        pinned_list = pinned_list_wrap.children_texts()

        #if pinned items don't exist, exception occurs
        if(app_type == 'PPT'):
            pinned_file_name = pinned_list[0] + ' - PowerPoint'
        else:
            pinned_file_name = pinned_list[0] + ' - Word'

        if(pinned_file_name == app_name):
            return app_name 

        close(app , app_name)

        t.sleep(1)

        if(app.is_process_running() == True):                
            return app_name
            
        elif(app.is_process_running() == False):                
            return pinned_file_name

        else:
            return app_name            

    except:
        send_keys('{VK_ESCAPE}')
        return app_name



def close(app , app_name):

    try:

        clear_stuff()        
        
        t.sleep(0.3)
        send_keys('^s')        

        t.sleep(0.5)

        close_button = app[app_name].child_window(title="Close", control_type="Button", class_name = 'NetUIAppFrameHelper').wrapper_object()

        close_button.click_input()

        t.sleep(1)        

    except:
        clear_stuff()
        return 


def volume_control():

    count = 0

    t.sleep(3)

    while True:
        # Read each frame from the webcam

        _, frame = cap.read()

        x, y, c = frame.shape

        # Flip the frame vertically
        frame = cv2.flip(frame, 1)
        framergb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        # Get hand landmark prediction
        result = hands.process(framergb)

        currentVolume_100 = 'Volume Control'

        currentVolume = volume.GetMasterVolumeLevel()

        # post process the result
        if result.multi_hand_landmarks:
            landmarks = []
            for handslms in result.multi_hand_landmarks:
                for id , lm in enumerate(handslms.landmark):
                    
                    lmx = int(lm.x * y)
                    lmy = int(lm.y * x)

                    landmarks.append([id , lmx, lmy])

                # Drawing landmarks on frames
                mpDraw.draw_landmarks(frame, handslms, mpHands.HAND_CONNECTIONS)

            if landmarks != []:

                x1,y1 = landmarks[4][1],landmarks[4][2]  #thumb
                x2,y2 = landmarks[8][1],landmarks[8][2]  #index finger
                #creating circle at the tips of thumb and index finger
                cv2.circle(frame,(x1,y1),13,(255,0,0),cv2.FILLED) #image #fingers #radius #rgb
                cv2.circle(frame,(x2,y2),13,(255,0,0),cv2.FILLED) 
                cv2.line(frame,(x1,y1),(x2,y2),(255,0,0),3)  #create a line b/w tips of index finger and thumb
        
                length = hypot(x2-x1,y2-y1)

                if length > 140:
                    volume.SetMasterVolumeLevel(currentVolume + 0.14, None)                    
                    currentVolume_100 = '      ' + str(round(volume.GetMasterVolumeLevelScalar()*100)) + '%'
                    if count> -30:
                        count -= 1
                    
                elif length>60:                  
                    volume.SetMasterVolumeLevel(currentVolume - 0.18, None)
                    currentVolume_100 = '      ' + str(round(volume.GetMasterVolumeLevelScalar()*100)) + '%'
                    if count> -30:
                        count -= 1

                elif length<40:
                    currentVolume_100 = 'Exit Volume Control'
                    count += 1
                        
        cam_screen_show(frame, currentVolume_100)

        if(count == 80):
            t.sleep(3)
            return None 



# initialize mediapipe
mpHands = mp.solutions.hands
hands = mpHands.Hands(max_num_hands=1, min_detection_confidence=0.5)
mpDraw = mp.solutions.drawing_utils

# Load the gesture recognizer model.
model = load_model('mp_hand_gesture')

# Load class names
f = open('application_controls.names', 'r')
application_controls = f.read().split('\n')
f.close()


f = open('ppt_controls.names', 'r')
ppt_controls = f.read().split('\n')
f.close()


f = open('doc_controls.names', 'r')
doc_controls = f.read().split('\n')
f.close()


f = open('slideshow_controls.names', 'r')
slideshow_controls = f.read().split('\n')
f.close()


available_controls = ['Open/Connect PowerPoint' , 'Open/Connect Word' , 'Stop Program' , 'Begin Slideshow' , 'Print' , 'Save As' , 'Close PowerPoint' , 'Close Word' , 'Back to Main Menu' , 'Open Pinned Presentation' , 'Open Pinned Document' , 'Switch to PowerPoint' , 'Switch to Word', 'End Slideshow' , 'Next Slide' , 'Previous Slide']

ppt_app_active = False
doc_app_active = False
ppt_app_name = ''
doc_app_name = ''
ppt_connect_run = False
doc_connect_run = False

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# Initialize the webcam
cap = cv2.VideoCapture(0)

program_control = None


while True:

    try:

        if(program_control == None):
            program_control = scan_gesture(application_controls , 'appli')


        if(program_control == 'Open/Connect PowerPoint'):

            if(ppt_app_name == '' and ppt_app_active == False):            
                ppt_app_name, ppt_app_active = extract_app_name('PPT')                        

            if(ppt_app_active == False):

                ppt_app = Application(backend="uia").start(r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE", timeout = 20) # open powerpoint app
                t.sleep(1)
                
                if(ppt_app_name != ''):
                    
                    pinned_tab = ppt_app['PowerPoint'].child_window(title="Pinned", class_name = 'NetUITabHeader').wrapper_object()

                    pinned_tab.click_input()

                    send_keys('{TAB}')
                    t.sleep(0.5)
                    send_keys('{ENTER}')

                else:
                    send_keys('{ENTER}')

                t.sleep(1)
                ppt_app_name , ppt_app_active = extract_app_name('PPT')            


            if(ppt_app_active == True):

                if ppt_connect_run == False:
                    ppt_app = Application(backend="uia").connect(title = ppt_app_name, timeout = 20)     
                    ppt_app_window = gw.getWindowsWithTitle(ppt_app_name)[0]
                    ppt_connect_run = True

                if not(ppt_app_window.isMaximized):
                    ppt_app_window.maximize()

                ppt_app_window.activate()
                ppt_app , ppt_app_name , ppt_app_active , ppt_connect_run , program_control = app_controls(ppt_app, ppt_app_name , ppt_app_active, ppt_app_window , 'PPT' , doc_app_active) 


        elif(program_control == 'Open/Connect Word'):            

            if(doc_app_name == '' and doc_app_active == False):                
                doc_app_name, doc_app_active = extract_app_name('DOC')                

            if(doc_app_active == False):

                doc_app = Application(backend="uia").start(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE", timeout = 20) # open word app
                t.sleep(1)

                if(doc_app_name != ''):
                    
                    pinned_tab = doc_app['Word'].child_window(title="Pinned", class_name = 'NetUITabHeader').wrapper_object()

                    pinned_tab.click_input()

                    send_keys('{TAB}')
                    t.sleep(0.5)
                    send_keys('{ENTER}')

                else:
                    send_keys('{ENTER}')

                t.sleep(1)
                doc_app_name , doc_app_active = extract_app_name('DOC')         
        

            if(doc_app_active == True):

                if doc_connect_run == False:
                    doc_app = Application(backend="uia").connect(title = doc_app_name, timeout = 20)     
                    doc_app_window = gw.getWindowsWithTitle(doc_app_name)[0]
                    doc_connect_run = True

                if not(doc_app_window.isMaximized):
                    doc_app_window.maximize()

                doc_app_window.activate()
                doc_app , doc_app_name , doc_app_active , doc_connect_run , program_control= app_controls(doc_app, doc_app_name , doc_app_active , doc_app_window , 'DOC' , ppt_app_active) 


        elif(program_control == 'Open Volume Control'):
            program_control = volume_control()


        elif(program_control == 'Stop Program'):            
            break


    except Exception as e: 
        
        try:
            if(e == 'Error code from Windows: 1400 - Invalid window handle.'):        

                if(program_control == 'Open/Connect PowerPoint'):
                    ppt_app = None 
                    ppt_app_name = '' 
                    ppt_app_active = False 
                    ppt_connect_run = False
                    program_control = None
                
                elif (program_control == 'Open/Connect Word'):
                    doc_app = None 
                    doc_app_name = '' 
                    doc_app_active = False 
                    doc_connect_run = False
                    program_control = None

            t.sleep(2)
                    
        except:

            ppt_app = None 
            ppt_app_name = '' 
            ppt_app_active = False 
            ppt_connect_run = False
            doc_app = None 
            doc_app_name = '' 
            doc_app_active = False 
            doc_connect_run = False
            program_control = None

            t.sleep(2)


cap.release()

cv2.destroyAllWindows()
