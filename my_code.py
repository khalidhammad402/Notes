from openpyxl import Workbook
import imaplib
import email
from tkinter import *
from tkinter.filedialog import askdirectory, asksaveasfilename
from tkcalendar import DateEntry, Calendar
import os
from tkinter import messagebox
import re
from datetime import date, datetime
from tkinter import ttk


def call_email_window(con):
    def logout(e=None):
        if email_window.focus_get() == logout_button or e is None:
            try:
                con.logout()
                email_window.quit()
                email_window.destroy()
                open_login_window()
            except:
                messagebox.showerror("Logout Failed", "Please Retry Logging Out!")
                return


    def toggle_edit_save(button, entry):
        if entry["state"] == "normal":
            entry.configure(state="readonly")
            button.configure(text="Edit")
        elif entry["state"] == "readonly":
            entry.configure(state="normal")
            button.configure(text="Save")

    def handle_subject_title(event):
        if subject_title_combobox.current() == 0:
            subject_title.set("Shortname_Internship_Application_MobileNo.")
            subject_title_entry['state'] = 'disabled'

        elif subject_title_combobox.current() == 1:
            subject_title.set("Shortname_Internship_Application_MobileNo.")
            subject_title_entry['state'] = 'readonly'
        else:
            subject_title.set("")
            subject_title_entry['state'] = 'disabled'

    def set_student_folder_path():
        path = askdirectory(title="Location for Student Files")
        print(path)
        if not isinstance(path, tuple):
            student_files_folder_path.set(path.strip())

    def is_date_valid(date_string):
        try:
            datetime.strptime(date_string, "%d-%m-%Y")
            return True
        except:
            return False

    def is_to_gte_from(from_string, to_string):

        if is_date_valid(from_string) and is_date_valid(to_string):
            from_day, from_month, from_year = list(map(int,from_string.split("-")))
            to_day, to_month, to_year = list(map(int,to_string.split("-")))
            if to_year > from_year:
                return True
            elif to_year==from_year and to_month > from_month:
                return True
            elif to_year==from_year and to_month==from_month and to_day >= from_day:
                return True
            else:
                return False
        else:
            return False

    def is_date_lte_today(date_string):
        today = date.today()
        if is_to_gte_from(date_string, today.strftime("%d-%m-%Y")):
            return True
        else:
            return False

    def is_student_folder_path_valid():
        try:
            if student_files_folder_path.get():
                files = os.listdir(path=student_files_folder_path.get())
            else:
                foldername = f"Downloaded_Application_{datetime.now().strftime('%d-%m-%Y-%T')}"
                foldername = foldername.replace(":", ".")

                path = os.path.join(os.path.join(os.path.expanduser('~'), 'Desktop'), foldername)
                defaultpath.set(path)

                if not os.path.exists(defaultpath.get()):
                    os.makedirs(defaultpath.get())
                student_files_folder_path.set(defaultpath.get())

            return True
        except:
            return False

    def validate_before_student_files():
        if not is_date_valid(from_date.get()):
            messagebox.showerror("Invalid Date", "Please Enter 'From' date in specified format")
            return False
        elif not is_date_valid(to_date.get()):
            messagebox.showerror("Invalid Date", "Please Enter 'To' date in specified format")
            return False
        elif not is_date_lte_today(from_date.get()):
            messagebox.showerror("Date Incompatible", "Please Enter 'From Date' less than or equal to today")
            return False
        elif not is_to_gte_from("01-01-1900", from_date.get()):
            messagebox.showerror("Date Incompatible",
                                 "Please Enter 'From Date' greater than or equal to 01-01-1900")
            return False
        elif not is_date_lte_today(to_date.get()):
            messagebox.showerror("Date Incompatible", "Please Enter 'To Date' less than or equal to today")
            return False
        elif not is_to_gte_from(from_date.get(), to_date.get()):
            messagebox.showerror("Date Incompatible",
                                 "Please Enter 'To Date' greater than or equal to 'From Date' ")
            return False
        elif file_title.get() == "":
            messagebox.showerror("Empty File Title", "Please Enter File Title")
            return False
        elif not is_student_folder_path_valid():
            messagebox.showerror("Destination Path Invalid", "Please Enter valid folder path")
            return False

        return True

    def count_files(path):
        try:
            return len(os.listdir(path))
        except:
            return 0

    def is_pattern_matched(mail_subject):
        mail_subject = mail_subject.strip().upper().replace("\n", "").replace("\r", "").replace(" ", "")\
            .replace("/", "-")
        input_pattern = "[^_]+_INTERNSHIP_APPLICATION_[0-9]{10}$"
        regex = re.fullmatch(input_pattern, mail_subject)
        if regex is None:
            return False
        return True

    def is_file_title_valid(file_name):
        pattern = "[^_]+_"*(len(file_title.get().split("_")))
        pattern = pattern[:-1]
        pattern = pattern + filetype_combobox["values"][filetype_combobox.current()]
        regex = re.fullmatch(pattern, file_name)
        if regex is None:
            return False
        return True

    def download_attachments(msgs):
        count = 0
        for part in msgs.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()

            if bool(filename):
                filename = filename.replace("\n", "").replace("\r", "").replace(" ", "").replace("/", "-")
                if is_file_title_valid(filename):
                    date = "-".join(msgs['date'].split(" ")[1:4])
                    mobile_no = "-"
                    if subject_title_combobox.current() == 0:
                        mobile_no = msgs['subject'].strip().split("_")[-1]

                    else:
                        if is_pattern_matched(msgs['subject']):
                            mobile_no = msgs['subject'].strip().split("_")[-1]

                    filename = ".".join(filename.split(".")[:-1])
                    filename = f"{msgs['from']}_{mobile_no}_" + filename + f"_{date}{filetype_combobox['values'][filetype_combobox.current()]}"
                    filename = filename.replace("\n", "").replace("\r", "").replace(" ", "").replace("/", "-").replace(
                        "<", "(").replace(">", ")")
                    print(filename)
                    filePath = os.path.join(student_files_folder_path.get(), filename)
                    filePath = filePath.replace("\\", "/")

                    with open(filePath, 'wb') as f:
                        f.write(part.get_payload(decode=True))
                    count += 1
        return count

    def download_student_files():
        if validate_before_student_files():
            from_dt_obj = datetime.strptime(from_date.get(), "%d-%m-%Y")
            to_dt_obj = datetime.strptime(to_date.get(), "%d-%m-%Y")
            search_string = f'(SINCE "{from_dt_obj.strftime("%d-%b-%Y")}" (OR ON "{to_dt_obj.strftime("%d-%b-%Y")}" BEFORE "{to_dt_obj.strftime("%d-%b-%Y")}"))'
            if subject_title_combobox.current() == 1:
                search_string = f'(SINCE "{from_dt_obj.strftime("%d-%b-%Y")}" (OR ON "{to_dt_obj.strftime("%d-%b-%Y")}" BEFORE "{to_dt_obj.strftime("%d-%b-%Y")}") SUBJECT "{subject_title.get().strip()}")'
            result, data_list = con.search(None, search_string)
            ids = data_list[0]
            ids_list = ids.split()
            count = count_files(student_files_folder_path)
            init_count = count
            for m_ids in ids_list:
                result, data = con.fetch(m_ids, '(RFC822)')
                raw = email.message_from_bytes(data[0][1])
                if subject_title_combobox.current() == 0:
                    if is_pattern_matched(raw['subject']):
                        count += download_attachments(raw)
                elif subject_title_combobox.current() == 1:
                    if raw['subject'] == subject_title.get().strip():
                        count += download_attachments(raw)
                else:
                    count += download_attachments(raw)
            if count == 0:
                if defaultpath.get() == student_files_folder_path.get():
                    try:

                        if os.path.exists(defaultpath.get()):
                            os.rmdir(defaultpath.get())
                            student_files_folder_path.set("")
                            defaultpath.set("")
                    except:
                        print("Could not remove folder")

            if count == init_count:
                messagebox.showinfo("No Files Found", "No match Found With The Given Format")
            else:
                messagebox.showinfo("Downloaded",
                                    f"{count - init_count} Files of specified format downloaded successfully")

    def validate_before_excel():
        if not os.path.exists(student_files_folder_path.get()):
            messagebox.showerror("Source Path Invalid", "Please check the student files folder path")
            return False
        return True

    def download_excel_to_folder():
        if validate_before_excel():
            files = os.listdir(path=student_files_folder_path.get())
            is_file = False
            for f_title in files:
                if is_file_title_valid("_".join(
                        f_title.split("_")[2:-1]) + f"{filetype_combobox['values'][filetype_combobox.current()]}"):
                    is_file = True

            if is_file:
                workbook = Workbook()
                sheet = workbook.active
                sheet_headings = ['Email', 'MobileNo.']

                sheet_headings.extend(file_title.get().split("_"))
                sheet_headings.append('Date')

                for i in range(len(sheet_headings)):
                    sheet[f"{chr(65 + i)}1"] = sheet_headings[i]

                row = 2

                for f_title in files:
                    if is_file_title_valid("_".join(
                            f_title.split("_")[2:-1]) + f"{filetype_combobox['values'][filetype_combobox.current()]}"):
                        data_list = f_title.split("_")
                        data_list[-1] = data_list[-1].split(".")[0]
                        for j in range(len(data_list)):
                            sheet[f"{chr(j + 65)}{row}"] = data_list[j]
                        row = row + 1

                initialfilename = f"Downloaded_Applications_{datetime.now().strftime('%d-%m-%Y-%T')}_Excel"
                initialfilename = initialfilename.replace(":", ".")

                excel_file_path = asksaveasfilename(initialdir=student_files_folder_path.get(),
                                                    initialfile=initialfilename, filetypes=[('Excel File', '*.xlsx')],
                                                    defaultextension=".xlsx")

                if excel_file_path:
                    try:
                        workbook.save(excel_file_path)
                        messagebox.showinfo("Downloaded", "Information Saved From " + student_files_folder_path.get())
                    except:
                        messagebox.showerror("Error", "File could not be saved. Please Try Again")
            else:
                messagebox.showinfo("No such File", "Files with specified format in the given folder doesn't exist.")

    email_window = Tk()
    email_window.title("Email Window!")
    screen_width = email_window.winfo_screenwidth()
    screen_height = email_window.winfo_screenheight()
    x = int((screen_width / 2) - (900 / 2))
    y = int((screen_height / 2) - (800 / 2))
    email_window.geometry(f'900x800+{x}+{y}')

    bgimg = PhotoImage(file="images/emailbg.png")
    bglabel = Label(email_window, image=bgimg)
    bglabel.place(x=0, y=0, relwidth=1, relheight=1)

    buttonframe = Frame(email_window)
    buttonframe.pack(side=BOTTOM)

    framebg = PhotoImage(file="images/whitebg.png")
    fbglabel = Label(buttonframe, image=framebg)
    fbglabel.place(x=0, y=0, relwidth=1, relheight=1)

    from_date = StringVar("")
    from_label = Label(buttonframe, text="From: ", font=('Times New Roman', 17), bg='white')
    from_label.grid(column=1, row=1, padx=10, pady=10)
    DateEntry(buttonframe, textvariable=from_date, selectmode='day', date_pattern="dd-mm-yyyy",
                                cursor='mouse', font=('Times New Roman', 17)).grid(column=2, row=1, padx=10, pady=10)


    to_date = StringVar("")
    to_label = Label(buttonframe, text="To: ", font=('Times New Roman', 17), bg='white')
    to_label.grid(column=1, row=2, padx=10, pady=10)
    DateEntry(buttonframe, textvariable=to_date, selectmode='day', date_pattern="dd-mm-yyyy",
                              cursor='mouse', font=('Times New Roman', 17)).grid(column=2, row=2, padx=10, pady=10)

    Label(buttonframe, text="Subject Title: ", font=('Times New Roman', 17), bg='white').\
        grid(column=1, row=3, padx=10, pady=10)

    subject_title = StringVar("")
    subject_title.set("Shortname_Internship_Application_MobileNo.")
    subject_title_entry = Entry(buttonframe, textvariable=subject_title, state="readonly", readonlybackground='#fffff0',
                                disabledbackground='#fffff0', disabledforeground='black', bd=3, bg='white',
                                cursor='mouse', font=('Times New Roman', 17))
    subject_title_entry.grid(column=2, row=3, padx=10, pady=10)
    subject_title_button = Button(buttonframe, text="Edit",
                                  command=lambda: toggle_edit_save(subject_title_button, subject_title_entry),
                                  activebackground='grey', activeforeground='white', bg='white',
                                  font=('Times New Roman', 17, 'bold'), height=1, cursor='mouse')
    subject_title_button.grid(column=3, row=3, padx=10, pady=10)
    #     subject_title_button.configure({"disabledbackground":"white"})

    subject_title_combobox = ttk.Combobox(buttonframe, values=["Pattern", "Specific", "All"], width=10,
                                          state="readonly", justify="center", cursor='mouse',
                                          font=('Times New Roman', 17))
    subject_title_combobox.current(0)
    subject_title_combobox.bind("<<ComboboxSelected>>", handle_subject_title)
    subject_title_combobox.grid(column=4, row=3, padx=10, pady=10)

    file_title_label = Label(buttonframe, text="File Title: ", font=('Times New Roman', 17), bg='white')
    file_title_label.grid(column=1, row=4, padx=10, pady=10)
    file_title = StringVar("")
    file_title.set("Shortname_Degree_Subject_Duration_AppDate_City")
    file_title_entry = Entry(buttonframe, textvariable=file_title, readonlybackground='#fffff0', state="readonly", bd=3,
                             bg='white', cursor='mouse', font=('Times New Roman', 17))
    file_title_entry.grid(column=2, row=4, padx=10, pady=10)

    file_title_button = Button(buttonframe, text="Edit",
                               command=lambda: toggle_edit_save(file_title_button, file_title_entry),
                               activebackground='grey', activeforeground='white', bg='white',
                               font=('Times New Roman', 17, 'bold'), height=1, cursor='mouse')
    file_title_button.grid(column=3, row=4, padx=10, pady=10)

    filetype_label = Label(buttonframe, text="File Type: ", font=('Times New Roman', 17), bg='white')
    filetype_label.grid(column=1, row=5, padx=10, pady=10)

    filetype_combobox = ttk.Combobox(buttonframe, values=[".pdf", ".docx"], width=15, state="readonly",
                                     justify="center", cursor='mouse', font=('Times New Roman', 17))
    filetype_combobox.current(0)
    filetype_combobox.grid(column=2, row=5, padx=10, pady=10)

    student_files_location_label = Label(buttonframe, text="Enter Location to Store Student Files",
                                         font=('Times New Roman', 17), bg='white')
    student_files_location_label.grid(column=1, row=7, padx=5, pady=10)

    student_files_folder_path = StringVar("")
    defaultpath = StringVar("")
    student_files_location_entry = Entry(buttonframe, textvariable=student_files_folder_path, bd=3, bg='white',
                                         cursor='mouse', font=('Times New Roman', 17))
    student_files_location_entry.grid(column=2, row=7, padx=10, pady=10)
    browse_files_button = Button(buttonframe, text="Browse Files", command=set_student_folder_path,
                                 activebackground='grey', activeforeground='white', bg='white',
                                 font=('Times New Roman', 17, 'bold'), height=1, cursor='mouse')
    browse_files_button.grid(column=4, row=7, padx=10, pady=10)

    download_student_file_button = Button(buttonframe, text="Download Student Files", command=download_student_files,
                                          activebackground='grey', activeforeground='white', bg='white',
                                          font=('Times New Roman', 17, 'bold'), height=1, cursor='mouse')
    download_student_file_button.grid(column=2, row=8, padx=10, pady=10)

    download_excel_file_button = Button(buttonframe, text="Generate Excel File", command=download_excel_to_folder,
                                        activebackground='grey', activeforeground='white', bg='white',
                                        font=('Times New Roman', 17, 'bold'), height=1, cursor='mouse')
    download_excel_file_button.grid(column=2, row=10, padx=10, pady=10)

    logout_button = Button(buttonframe, text="LOG OUT", command=logout, activebackground='grey',
                           activeforeground='white', bg='white', font=('Times New Roman', 17, 'bold'), height=1,
                           cursor='mouse')
    logout_button.grid(column=2, row=12, padx=10, pady=10)

    style = ttk.Style()
    style.theme_create('combostyle', parent='alt',
                       settings={'TCombobox':
                                     {'configure':
                                          {'selectbackground': 'white',
                                           'selectforeground': 'black',
                                           'fieldbackground': 'white',
                                           'background': 'white'

                                           }}})
    style.theme_use('combostyle')

    email_window.resizable(0, 0)
    email_window.bind("<Return>", logout)

    email_window.mainloop()


def open_login_window():
    def login(e=None):
        try:
            imap_host_server = domain_name.get().strip()
            con = imaplib.IMAP4_SSL("imap.gmail.com")
            con.login("khalidhammad402@gmail.com", "uzhisopnyphkjhdp")
            #             print("Successfully logged in")
            con.select("INBOX")
            login_window.quit()
            login_window.destroy()
            call_email_window(con)
            return
        except:
            #             print("Authentication Failed!")
            domain_name.set("")
            username.set("")
            password.set("")
            messagebox.showerror("Login Failed", "Invalid Crendentials. Please Retry!")
            return

    login_window = Tk()
    login_window.title("Main-Window!")
    screen_width = login_window.winfo_screenwidth()
    screen_height = login_window.winfo_screenheight()
    x = int((screen_width / 2) - (450 / 2))
    y = int((screen_height / 2) - (450 / 2))
    login_window.geometry(f'450x450+{x}+{y}')
    bgimg = PhotoImage(file="images/emailbg.png")
    bglabel = Label(login_window, image=bgimg)
    bglabel.place(x=0, y=0, relwidth=1, relheight=1)

    buttonframe = Frame(login_window)
    buttonframe.pack(side=BOTTOM)

    framebg = PhotoImage(file="images/whitebg.png")
    fbglabel = Label(buttonframe, image=framebg)
    fbglabel.place(x=0, y=0, relwidth=1, relheight=1)

    domain_name = StringVar("")
    Label(buttonframe, text="Enter Domain Name: ", font=('Times New Roman', 15), bg='white').grid(
        column=1, row=1, padx=10, pady=10)
    Entry(buttonframe, textvariable=domain_name, bd=3, bg='white', cursor='mouse',
                         font=('Times New Roman', 15)).grid(column=2, row=1, padx=10, pady=10)

    username = StringVar("")
    Label(buttonframe, text="Enter Email Id: ", font=('Times New Roman', 15), bg='white').grid(
        column=1, row=2, padx=10, pady=10)
    Entry(buttonframe, textvariable=username, bd=3, bg='white', cursor='mouse',
                           font=('Times New Roman', 15)).grid(column=2, row=2, padx=10, pady=10)

    password = StringVar("")
    Label(buttonframe, text="Enter password: ", font=('Times New Roman', 15), bg='white').grid(
        column=1, row=3, padx=10, pady=10)
    Entry(buttonframe, textvariable=password, show="*", bd=3, bg='white', cursor='mouse',
                           font=('Times New Roman', 15)).grid(column=2, row=3, padx=10, pady=10)

    login_button = Button(buttonframe, text="LOGIN", command=login, activebackground='grey', activeforeground='white',
                          bg='white', font=('Times New Roman', 13, 'bold'), height=1, cursor='mouse')
    login_button.grid(column=1, row=4, padx=10, pady=20, columnspan=2)

    login_window.resizable(0, 0)
    login_window.bind("<Return>", login)

    login_window.mainloop()


open_login_window()