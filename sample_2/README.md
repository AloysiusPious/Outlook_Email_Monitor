Bugs to Fixed:
1. If any csv or png file not found then it should not give any error.(Refer: Its already fixed for abc Monitoring, check Node Down.png file)
2. DO the Step 1 for all other file attachment
########################Flask Virtual Env ####################
This is what worked for me when I got a similar error in Windows:

Install virtualenv:

pip install virtualenv
Create a virtualenv:
cd C:\Users\pious.aloysius\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra
8p0\LocalCache\local-packages\Python312\Scripts

./virtualenv flask
Navigate to Scripts and activate the virtualenv:
cd C:\Users\pious.aloysius\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\Scr
ipts\flask\Scripts>

activate
Install Flask:

python -m pip install flask
Check if flask is installed:

python -m pip list
############################ Flask End #######################
########### Make it self executable
pip install pyinstaller
pyinstaller --onefile detasad_web_gui.py
