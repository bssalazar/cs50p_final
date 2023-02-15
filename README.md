# BeMyStudent - College Students List Generator ###
## Created by: Brian Salazar for CS50P Final ##
## GitHub profile: https://github.com/bssalazar ##
### Video Demo: [Youtube](https://youtu.be/UjKzuwCSijc) ###
#### Description: ####
**BeMyStudent (BMS)** is a simple Python program that uses the [tkinter](https://docs.python.org/3/library/tkinter.html) package to generate a basic interface to create a list of students in a course/class.
The goal of this program is to let teachers run a simple Python program to collect personal information
of students which then may be used in announcements,
e-mails, class requirements, and all other information that need to be disseminated to the people on the class list.
This can be an alternative to on-the-spot inputs to a spreadsheet program (e.g. MS Excel)
or Google Forms which requires internet connection. 


## How it works: ##
The **BMS** tkinter GUI asks the user (i.e. student) the following information:
+ First name
+ Last name
+ Year level
  - Using *combobox* (or drop-down menu), type the input or choose from the list of options
+ College where the user belongs
  - Using *combobox*, type the input or choose from the list of options
+ Enrollment status
+ Contact number
+ Active and valid email address

*Note: User may choose to tick the enrollment status box if he has finished
the enrollment process or leave it unchecked if not. Checked: "Enrolled", Unchecked: "Pending" status*


Before submitting the form, there is another checkbox for Data Sharing Consent (DSC) which 
is required to be ticked in order to successfully add the data input for the fields above. 
If left unchecked, even if all data above are valid, an error message will appear.
When this is ticked but there are empty fields or invalid input above, different error
message pop-ups will appear depending on which field has error. *Add student* button
calls the `get()` method of the input widgets to capture the inputs. These will be passed as 
parameters to the [lambda](https://docs.python.org/3/reference/expressions.html#lambda) expression 
`enter_data()`. 

The parameters will be passed to multiple functions inside `enter_data()` which validates them. When
invalid, the validating function will return `False` and the error pop-up appears. Otherwise, the values will just
be returned back. 

Assuming that all inputs are valid and the *DSC* checkbox is ticked:
`filename = "student_data" + ".xlsx"` sets the filename of the list.
This will be the file that will be created in the local directory where this script is found.
Using the [os.path module](https://docs.python.org/3/library/os.path.html#module-os.path), `if not os.path.exists(filename):` checks the directory if a file already exists. 
If it doesn't find any, a new workbook will be generated using the [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
library. Column headings will be set by `heading = ["First name", "Last name", "Year", "College", "Status", "Contact no.", "Email add"]`
and the entries will be appended to the next row below and saved. Lastly, using *messagebox* a pop-up message showing 
successful addition of data to the workbook will be shown to the user. 

However, if the `student_data.xlsx` already exists, the function `duplicate_checker()` will be called.
The parameter of this function is a dictionary with column headings as keys and input values as the corresponding values (input_dict).
Using [pandas.read_excel](https://pandas.pydata.org/docs/reference/api/pandas.read_excel.html),
the file will be returned as Dataframe. The method `df.to_dict(orient="records")` will transform the dataframe rows to a list of dictionaries
where each dictionary keys are the column names(heading) and the values are the values in their corresponding rows.
Once it finds that input_dict already exists in the list of dictionaries from `.to_dict()`, it will return `False`. 
An error message will appear as duplicate entries are not allowed. Otherwise, `True` is the return value. 
The existing file will be loaded and the data will be appended below the last data row. A messagebox shows again which indicates successful
addition to the sheet.

The last button in the 'Options' frame is a 'Clear data' button. Once clicked, the 
function `clear_all()` is called. This function sets the entry fields into their default. 
For entry and combobox/dropdown fields, these will be blank. Checkboxes will be unchecked again. 
Doing this will enable the next user to enter their personal information. Clicking this button is a 
necessary step to save time. Instead of manually deleting the entries of the previous user, the next one can just 
start typing right away.


Upon the entry of the last user, the GUI window can just be closed to end the execution of the program. 
The file can now be opened and viewed using a spreadsheet app like (MS Excel) where the person has the option
to resize column width and keep it as a spreadsheet file, or export it as PDF and print it out. 

However, when the script is executed again, the incoming data entries will be appended to the existing file.
If the user wishes to create a fresh new list, the existing file must be deleted in the directory first before running the program again.

### Areas for improvement/other features to add: ###
- Create a button which opens a dialog box that enables user to select an existing student list file and load it. This will be the file where new inputs will be appended.
- A button in the GUI window that allows user to set a desired file name and create a new file. This way, the file name will not be hard-coded in the program and will just come from the user input.







