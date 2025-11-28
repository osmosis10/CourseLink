# CourseLink
#### Author: Moses Lemma - Business Systems Analyst Student

## Description
CourseLink is a simple UI that automates certain tasks that are  
needlessly time consuming when preparing Employee Learn Course  
data to be uploaded to SAMI (Sales and Marcom Intelligence). CourseLink uses the following libraries
- Tkinter (For GUI)
- Openpyxl (For excel functions)
- PIL (For applying the icon.ico file for the app icon)
- Pyinstaller (To compile the executable)   

<br>
<br>

<br>
<br>

##  Update Providers
The user can click the 'Update Providers' button to update all  
employees with 'Non-PCL ILT' (in report #2) with the correct providers.  
from report #3 Additionally, the word 'Archive' will be removed from  
any cells containing it in the Training Provider column.  
<br>
<br>

## Update Approval
The user can click the 'Update Approval' button to update the  
approval status for all LINKEDIN LEARNING courses (in report #1)  
from 'Approved' to 'Completed'.  (Copy those records and paste in report #2 manually).
<br>
<br>

## Check For Duplicates
The user can click the 'Check For Duplicates' button to check if  
there are any duplicate records in report #2.
<br>
<br>

## Return Instances of Archive
The user can return any instance of 'Archive' occuring in the   
training title column in report #2.
<br>
<br>

## Output scrollbox
The user can view any output produced by the updates to validate or confirm. 

<br>

## Toggle warning
The user can toggle the 'Toggle warnings' switch to turn off warnings for the following:
- Warning to complete manual steps 3, 6 and 7 after closing.
- Warning when there are 1 or more of necessary excel files missing.
- Warning when no actions were performed because the files were already  
  updated.

## Output files ##
The results of the CourseLink will be exported to a courseUpdated.txt file, with a new file generated if any already exists.

Critical exceptions will be logged and outputed to a debug_log.txt.

## How to create executable ##
In order to generate the executable ensure you have [pyinstaller](https://pypi.org/project/pyinstaller/) and [openpyxl](https://pypi.org/project/openpyxl/) installed on your machine. Afterwards:  
1. Navigate to working directory and open command line
2. attempt to run ```pyinstaller CourseLinkSRC.py``` note: If this fails you may have to use the path to pyinstaller itself ex. ```C:\Users\YourUserName\AppData\Local\Packages\PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0\LocalCache\local-packages\Python312\site-packages\PyInstaller\__main__.py CourseLinkSRC.py```

