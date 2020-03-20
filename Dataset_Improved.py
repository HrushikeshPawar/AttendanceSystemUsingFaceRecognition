#Import Important Libraries
import cv2 #To do all the hard work of Face Detection and Recognition
import os #To read and store the necessary input and output data
import imutils #For using Webcam and playing with images
from imutils.video import VideoStream
import time

def Data(name,roll,path):
    #Load the Webcam
    cam = VideoStream(scr=0).start()

    #Import the required Cassade File, for Face Detection
    face_detector = cv2.CascadeClassifier('Required_Files/haarcascade_frontalface_default.xml')

    print("\n [INFO] Initializing face capture. Look at the camera and wait ...\n")

    # Initialize individual sampling face count
    count = 0 #Will count the number of images
    while(True):
        img = cam.read() #Take a single frame from the VideoStream
        orig = img.copy() #Make a copy of this image frame
        img = cv2.flip(img, 1) # flip video image vertically
        img = imutils.resize(img, width=600) #Set the size of video window


        #Convert a coloured image into Black-and-White and then detect all the faces
        #in a frame
        faces = face_detector.detectMultiScale(
            cv2.cvtColor(img, cv2.COLOR_BGR2GRAY), scaleFactor=1.1,
            minNeighbors=5, minSize=(30, 30))

        for (x,y,w,h) in faces:
            #Draw a rectangle around the face
            cv2.rectangle(img, (x,y), (x+w,y+h), (255,0,0), 2)
            cv2.imshow('image', img) #Stream the video with detected faces

        k = cv2.waitKey(1) & 0xff #Check if any key is pressed
        if k == 107:  #If "k" is pressed, capture your image
            count += 1  #Increase the count of images by 1
            #Save the image in specified folder
            cv2.imwrite(path +"/"+ str(count) + ".jpg", orig)
            print(" [INFO] Photos Saved - {}/10".format(count))
        elif count >= 10 : # Take 10 face sample and stop video
             break


    # Do a bit of cleanup and print the details
    print("\n [INFO] Exiting Program and cleanup stuff")

    #Turn of th camera
    cv2.destroyAllWindows()
    cam.stop()
    return
