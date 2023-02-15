import openpyxl
import os
import pandas as pd
import re
import tkinter
from tkinter import ttk
from tkinter import messagebox
from validator_collection import checkers


colleges = ["Engineering", "Science", "Arts and Letters", "Mass Communication",
            "Economics", "Music", "Public Administration", "Business",
            "Statistics", "Architecture", "Education", "Law"]

year_level = ["Frosh", "Soph", "Junior", "Senior"]


def main():
    window_main()


def name_check(first, last):
    # checks validity of 'first' and 'last' parameter
    if first == '' or last == '':
        return False, False
    else:
        return first.title().strip(), last.title().strip()


def level_check(year):
    # checks validity of 'year' input
    if year not in year_level:
        return False
    else:
        return year


def col_check(home):
    # checks validity of 'home' input
    if home not in colleges:
        return False
    else:
        return home


def contact_check(num):
    # checks if contact number contains numbers only and follows the correct format.
    if num.isnumeric():
        num_pattern = re.match(r'^6309[0-9]{9}$', num)
        if num_pattern:
            return num
        else:
            return False
    else:
        return False


def email_check(address):
    # checks if 'address' is a valid email address
    if checkers.is_email(address):
        return address
    else:
        return False


def enter_data(fname, surname, level, col, enrol, contact, email, consent):
    """
    enter_data shows most of the logic behind the functionalities of this program
    Parameters are data obtained when the 'Add student' button widget command is executed.
    This will get the data from all Entry, Checkbutton/StringVar, and Combobox widgets.

    How it works:

    1. Check if any of first or last name fields is empty. If one or both is empty, show error message pop-up.
    2. For year level and college: check if user input belongs to list of options. If not, show error message pop-up.
    3. Enrollment status: if checked, 'Enrolled' is stored. Otherwise, 'Pending' is stored. Default state is unchecked ('Pending')
    4. For contact number: check if input contains numbers only. If not, error pop-up appears. Then,
        regex detects if the pattern 6309XXXXXXXXX is there. Otherwise, error pop-up appears.
    5. Using validator-collection, email address input by user is validated. If it's invalid, error message appears.
    6. Data sharing consent: when CHECKED and ALL entries above are valid, the 'Add student' button proceeds.
        A message box indicating successful addition of student to list will appear and details will
        be added to the Excel sheet. Otherwise, error messages from invalid input in widgets above will show
        and error message for not ticking the consent box will appear.
    7. When user details are successfully added to the Excel file and user clicks 'Add student' again,
        an error message will appear as duplicate entries are not allowed.
    8. Clear data button, when clicked, calls clear_all() that resets the value of input fields
        to empty or their default values.
    """
    response_check = []
    first, last = name_check(fname, surname)
    # for displaying error message.
    if (first, last) == (False, False):
        tkinter.messagebox.showerror(title="Error", message="First and last name is required")
    age = level_check(level.title())
    if age is False:
        tkinter.messagebox.showerror(title="Error", message="Please put proper year level.")
    org = col_check(col.title())
    if org is False:
        tkinter.messagebox.showerror(title="Error", message="College selected not in the choices.")
    number = contact_check(contact)
    if number is False:
        tkinter.messagebox.showerror(title="Error", message="Contact number should only contain numbers and in this format: 6309XXXXXXXXX")
    mail = email_check(email)
    if mail is False:
        tkinter.messagebox.showerror(title="Error", message="Invalid email address!")
    # collects the values returned from all _check() functions called.
    response_check.extend([first, last, age, org, number, mail])
    # checks if all other values are valid/functions above did not return any False.
    if False not in response_check:
        # validate if Data Sharing Consent checkbutton in ticked.
        if consent == "agree":

            filename = "student_data" + ".xlsx"
            input_dict = {"First name": first, "Last name": last, "Year": age, "College": org, "Status": enrol,
                          "Contact no.": int(number), "Email add": mail}
            # if the Excel database 'student_data.xlsx' does not exist, it proceeds to create the file and
            # append the values entered by the user.
            if not os.path.exists(filename):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First name", "Last name", "Year", "College", "Status", "Contact no.", "Email add"]
                sheet.append(heading)
                sheet.append([first, last, age, org, enrol, number, mail])
                workbook.save(filename)
                tkinter.messagebox.showinfo(title="Information", message="Student successfully added to list.")
            # if file already exists, duplicate_checker is called to check if user attempts
            # to add duplicate entries to the Excel database. Either a confirmation or error message will show.
            else:
                exists = duplicate_checker(input_dict)
                if exists is True:
                    workbook = openpyxl.load_workbook(filename)
                    sheet = workbook.active
                    sheet.append([first, last, age, org, enrol, number, mail])
                    workbook.save(filename)
                    tkinter.messagebox.showinfo(title="Information", message="Student successfully added to list.")
                else:
                    tkinter.messagebox.showerror(title="Error", message="Duplicates not allowed.")
        else:
            tkinter.messagebox.showerror(title="Error", message="Check the consent box to proceed.")
    else:
        tkinter.messagebox.showerror(title="Error", message="Check the consent box to proceed.")


def duplicate_checker(data):
    # opens the existing Excel database named student_data.xlsx
    # Makes every row in the file a dictionary where column name are keys and corresponding rows/cells are values.
    df = pd.read_excel("student_data.xlsx", engine='openpyxl')
    input_data = df.to_dict(orient="records")
    # check if duplicate entries exist
    if data in input_data:
        return False
    else:
        return True


def window_main():
    # set window label/main title
    window.title("BeMyStudent - College Students List Generator")
    # create frame and set background to black
    frame = tkinter.Frame(window)
    frame.pack()
    frame.configure(background="black")

    # Create frames for student info widgets and put them into their corresponding place using grid.
    student_info_frame = tkinter.LabelFrame(frame, text="Student Information", bg="dark gray")
    student_info_frame.grid(row=0, column=0, padx=20, pady=10)

    # create first and last name labels.
    first_name_label = tkinter.Label(student_info_frame, text="First name", bg="dark gray")
    last_name_label = tkinter.Label(student_info_frame, text="Last name", bg="dark gray")
    first_name_label.grid(row=0, column=0, padx=10, sticky="w")
    last_name_label.grid(row=0, column=1, padx=10, sticky="w")

    # create first and last name entry fields.
    first_name_entry = tkinter.Entry(student_info_frame)
    last_name_entry = tkinter.Entry(student_info_frame)
    first_name_entry.grid(row=1, column=0, pady=8)
    last_name_entry.grid(row=1, column=1, pady=8)

    # create year level and college labels & entry fields.
    year_label = tkinter.Label(student_info_frame, text="Year Level", bg="dark gray")
    year_entry = ttk.Combobox(student_info_frame, values=year_level)
    year_label.grid(row=2, column=0, padx=10, sticky="w")
    year_entry.grid(row=3, column=0, padx=10, pady=5)

    college_label = tkinter.Label(student_info_frame, text="College", bg="dark gray")
    college_combobox = ttk.Combobox(student_info_frame, values=sorted(colleges))
    college_label.grid(row=2, column=1, padx=10, sticky="w")
    college_combobox.grid(row=3, column=1, padx=10, pady=5)

    # Create frame for status and contact info widgets
    courses_frame = tkinter.LabelFrame(frame, text="Status and Contact Information", bg="light gray")
    courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)
    # this widget checks if student is fully enrolled or not.
    registered_label = tkinter.Label(courses_frame, text="Check, if applicable:", bg="light gray")
    registration = tkinter.StringVar(value="Pending")
    registered_check = tkinter.Checkbutton(courses_frame, text="Enrollment complete", variable=registration,
                                           offvalue="Pending", onvalue="Enrolled", bg="light gray")
    registered_label.grid(row=0, column=0, sticky="w")
    registered_check.grid(row=1, column=0, sticky="w")

    # widgets for getting student's contact info
    contact_label = tkinter.Label(courses_frame, text="Contact Number", bg="light gray")
    contact_entry = tkinter.Entry(courses_frame)
    contact_label.grid(row=0, column=1, sticky="w")
    contact_entry.grid(row=1, column=1, padx=10, pady=5)

    email_label = tkinter.Label(courses_frame, text="Active email address", bg="light gray")
    email_entry = tkinter.Entry(courses_frame)
    email_label.grid(row=0, column=2, sticky="w")
    email_entry.grid(row=1, column=2, padx=10, pady=5)

    # Checkbox to see if student agrees to share personal info. It won't proceed if student disagrees.
    consent_frame = tkinter.LabelFrame(frame, text="Data Sharing Consent")
    consent_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)
    required = tkinter.Label(consent_frame, text="All fields, except enrollment status, are required.")
    required.grid(row=0, column=0, sticky="w")
    consent_response = tkinter.StringVar(value="disagree")
    consent_check = tkinter.Checkbutton(consent_frame,
                                        text="I agree to share my personal information for the purposes of this class.",
                                        variable=consent_response, onvalue="agree",
                                        offvalue="disagree")
    consent_check.grid(row=1, column=0)

    # when conditions are satisfied, details input by student will be added to the Excel database/list.
    button_frame = tkinter.LabelFrame(frame, text="Options")
    button_frame.grid(row=3, column=0, padx=20, pady=10, sticky="e")
    button = tkinter.Button(button_frame, text="Add student", command=lambda: enter_data(first_name_entry.get(),
                                                                                         last_name_entry.get(),
                                                                                         year_entry.get(),
                                                                                         college_combobox.get(),
                                                                                         registration.get(),
                                                                                         contact_entry.get(),
                                                                                         email_entry.get(),
                                                                                         consent_response.get()),
                            bg="white")
    button.grid(row=0, column=0, sticky="w")
    clear_button = tkinter.Button(button_frame, text="Clear data", command=lambda: clear_all(), bg="white")
    clear_button.grid(row=0, column=1, sticky="e")

    # when called, resets the input fields into their default values.
    def clear_all():
        first_name_entry.delete(0, 'end')
        last_name_entry.delete(0, 'end')
        year_entry.delete(0, 'end')
        college_combobox.delete(0, 'end')
        registered_check.deselect()
        contact_entry.delete(0, 'end')
        email_entry.delete(0, 'end')
        consent_check.deselect()


if __name__ == "__main__":
    window = tkinter.Tk()
    main()
    window.mainloop()
