# Zoom-Auto-Joiner


Automatically join and leave Zoom meetings. 

As easy as it sounds, this program allows you to join and leave Zoom meetings automatically and hands free.



This program makes your school / college life much easier, it allows you to sleep for literally a whole week without worrying about joining your classes.



## 🏁 Getting Started

in order for this program to work properly, you are going to need to make some changes to it.

### 🔨 Changes
#### ➡️ Changes performed in the Zoom app:

First of all, you have to go to Zoom > Settings > General > then uncheck "Ask me to confirm when I leave a meeting"

Secondly, go to > Settings > Video > then check "Turn off my video when joining meeting"

Third of all, go to > Settings > Audio > then check "Automaatically join audio by computer when joining a meeting" > and check "Mute my microphone when joining a meeting"

Last but not least, go to > Settings > Keyboard Shortcuts > Scroll down until you see "End meeting" and change the shortcut to "Alt+X" (for some it is set by default)


#### ➡️ Changes performed in the provided Excel files :

This is the most important change you are going to perform, and you will see why in a bit:

after downloading the script from github, open up the "Daily Program" folder, go to the Monday file and add your Zoom meeting information accordingly (same with all the other days of the week).

Add the Starting Time of your meeting > Ending Time of your meeting > your Meeting's ID > your Meeting's Password

Starting and Ending time MUST be in 24hr format (check the provided "example.xlsx" file)

*Check the given example excel file if needed*

P.S: 

If a specific meeting doesn't have a password,leave the password cell empty.


### 💻 Requirements:
And this is a very important step.

#### 🌐General Requirments

Zoom app 

Web Browser (Firefox, Chrome or any other browser)

#### 🐍 Python 3.0
If you don't already have python, you are going to have to install it from [this link](https://www.python.org/downloads/).

#### 📚 Libraries
You are going to have to install a couple libraries in order for this program to run.
This won't take much time nor much space in your system.


Used libraries:
* [webbrowser](https://docs.python.org/3/library/webbrowser.html)
* [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
* [pyautogui](https://pyautogui.readthedocs.io/en/latest/install.html)
* [datetime](https://docs.python.org/3/library/datetime.html)
* [time](https://docs.python.org/3/library/time.html?highlight=time#module-time)



#### 📥 Using "pip install"

>py -m pip install -r requirements.txt (in cmd)
>OR
>pip install -r requirements.txt (in python terminal)


You might get an error like so in your CMD terminal:

>*ERROR: Could not find a version that satisfies the requirement webbrowser (from versions: none)*
>
>*ERROR: No matching distribution found for webbrowser*

If so...do the following, write each line in your cmd respectively:
>py -m pip install webbrowser
>py -m pip install openpyxl
>py -m pip install pyautogui
>py -m pip install datetime
>py -m pip install time

## 🏃‍♂️ Running the program
In order to finally run the program, all you have to do is run the "main.py" file, and in order to access this file go to the project's folder then double click the "main.py" file:

> \Zoom-Auto-Joiner\main.py


Whenever you want to join your classes, run the "main.py" file and you'll be good to go.

