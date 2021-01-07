#!/usr/local/bin/python3.7

# Import Modules
import numpy as np
import pandas as pd
import pymongo
from bson import ObjectId
from gridfs import GridFS
from tkcalendar import DateEntry
import csv
import glob
import os
import platform
import shutil
import subprocess
import tempfile
import threading
from datetime import datetime, timedelta
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

root = Tk()

# Global Variables
app_version = "0.0.1"
app_author = "Chris Rowlands"
myFont = ("Tahoma", "-12")
myFontBold = ("Tahoma", "-12", "bold")
myFontBoldUnderline = ("Tahoma", "-12", "bold", "underline")
os_type = platform.system()
use_database = "Test"

# OS Dependant Variables
if os_type == "Windows":
    button_height = 1
    dash_height = 500
    dash_width = 580
    from ctypes import windll

    windll.shcore.SetProcessDpiAwareness(2)

elif os_type == "Darwin":
    button_height = 2
    dash_height = 490
    dash_width = 570

# Change Working Directory to same directory as script file
os.chdir(os.path.dirname(os.path.abspath(__file__)))


# Database Class for Connection and CRUD Operations
# noinspection PyUnresolvedReferences
class Database:
    def autoreconnect_retry(func, retries=3):
        def db_op_wrapper(*args, **kwargs):

            tries = 0
            while tries < retries:
                try:
                    return func(*args, **kwargs)

                except pymongo.errors.AutoReconnect:
                    tries += 1

            messagebox.showerror("Exception",
                                 "Couldn't connect to the database, after %d retries. Please Log In again." % retries)

        return db_op_wrapper

    def make_connection(self, username_val, password_val):
        # Make Database Connection Object
        try:
            # Attempt To Make Connection
            Database.username_val = username_val
            Database.password_val = password_val

            Database.database_client = pymongo.MongoClient(
                "mongodb+srv://" + Database.username_val + ":" + Database.password_val + "@cluster0.xhsaj.mongodb.net/<dbname>?retryWrites=true&w=majority",
                serverSelectionTimeoutMS=3000)
            Database.database_client.server_info()

            db_name = "CrudApp"
            collection_name = "CrudAppCollection"
            log_collection_name = "Log"
            cert_collection_name = "TestFiles"

            Database.crud_app_db = Database.database_client[db_name]
            Database.crud_app_collection = Database.crud_app_db[collection_name]
            Database.log_collection = Database.crud_app_db[log_collection_name]
            Database.certificates_db = Database.database_client[db_name]
            Database.certificate_storage = GridFS(Database.certificates_db, cert_collection_name)



        except pymongo.errors.PyMongoError as e:
            return messagebox.showerror("Exception", e, parent=self.login_container)

        else:
            Database.log_action("Logged In", "")
            return True

    def log_action(action, log_filter):

        try:
            Database.log_collection.insert_one(
                {"Timestamp": datetime.now(),
                 "User": Database.username_val,
                 "Action": action,
                 "Filter": log_filter
                 })

        except pymongo.errors.PyMongoError:
            pass


class Login:
    def __init__(self, master):
        self.login_container = Toplevel(master)
        self.login_container.grid_columnconfigure((0, 6), weight=1)
        self.login_container.grid_rowconfigure((0, 1, 2, 3, 4, 5), weight=1)
        self.login_container.geometry("380x235")
        self.login_container.resizable(False, False)
        self.login_container.config(background="white")
        self.login_container.title("Boiler Service App")

        # Gets the requested values of the height and width.
        self.windowWidth = 380
        self.windowHeight = 235
        # Gets both half the screen width/height and window width/height
        self.positionRight = int(root.winfo_screenwidth() / 2 - self.windowWidth / 2)
        self.positionDown = int(root.winfo_screenheight() / 2 - self.windowHeight / 2)
        # Positions the window in the center of the page.
        self.login_container.geometry("+{}+{}".format(self.positionRight, self.positionDown))

        # Declare variables
        self.username_val = StringVar()
        self.password_val = StringVar()
        self.username_val.set('Admin')
        self.password_val.set('CRUDappTest123')

        # Create Logo Canvas
        self.img_logo = PhotoImage(file="graphics/logo_small.gif")
        self.logo_canvas = Canvas(self.login_container, width=280, height=70, background='white', highlightthickness=0,
                                  borderwidth=0)
        self.logo_canvas.grid_propagate(0)
        self.logo_canvas.grid(row=0, column=1, columnspan=4, rowspan=1, pady=10)
        self.logo_canvas.create_image(5, 5, anchor='nw', image=self.img_logo)

        self.login_frame = LabelFrame(self.login_container, text="Log In", width=360, height=100, background="white",
                                      font=myFont)
        self.login_frame.grid_propagate(0)
        self.login_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.login_frame.grid_rowconfigure((1, 2), weight=1)
        self.login_frame.grid(row=1, column=1, columnspan=4, rowspan=4)

        self.username_label = Label(self.login_frame, text="Enter Username:", anchor='e', background='white',
                                    font=myFont)
        self.username_label.grid(row=1, column=0, columnspan=1, rowspan=1, padx=5, pady=5, sticky="we")
        self.username_entry = Entry(self.login_frame, textvariable=self.username_val, font=myFont)
        self.username_entry.focus()
        self.username_entry.grid(row=1, column=1, columnspan=2, rowspan=1, sticky='ew')
        self.password_label = Label(self.login_frame, text="Enter Password:", anchor='e', background='white',
                                    font=myFont)
        self.password_label.grid(row=2, column=0, columnspan=1, rowspan=1, padx=5, pady=5, sticky="we")
        self.password_entry = Entry(self.login_frame, show="*", textvariable=self.password_val, font=myFont)
        self.password_entry.grid(row=2, column=1, columnspan=2, rowspan=1, sticky='ew')
        self.login_button = Button(self.login_container, text="Log In", background="cornflower blue", font=myFont,
                                   height=button_height, command=self.check_credentials)
        self.login_button.grid(row=5, column=3, columnspan=2, rowspan=1, pady=(5, 10), sticky="nsew", padx=(2, 0))
        self.quit_button = Button(self.login_container, text="Quit", background="cornflower blue", font=myFont,
                                  height=button_height, command=root.quit)
        self.quit_button.grid(row=5, column=1, columnspan=2, rowspan=1, pady=(5, 10), sticky="nsew", padx=(0, 2))

        self.password_entry.bind('<Return>', self.check_credentials)
        self.username_entry.bind('<Return>', self.check_credentials)
        self.login_container.protocol("WM_DELETE_WINDOW", root.quit)

    def check_credentials(self, event=None):

        self.login_button.config(state="disabled")
        self.login_button.update()

        if Database.make_connection(self, self.username_val.get(), self.password_val.get()):
            self.login_container.destroy()
            app.populate_combo_box()
            root.deiconify()
            app.refresh_dashboard()
        else:
            self.login_button.update()
            self.login_button.config(state="normal")


# noinspection PyUnresolvedReferences
class Mainwindow:
    month_report_court_selected_val = StringVar()
    month_report_month_selected_val = StringVar()
    custom_report_selected = StringVar()
    month_report_court_selected_val.set("_Show All")
    month_report_month_selected_val.set("_Show All")
    custom_report_selected.set("_Show All Units")

    def __init__(self, master):
        # Open Login Window at startup.
        Login(master)

        # Declare logo Image
        self.img_logo = PhotoImage(file="graphics/logo_small.gif")

        # Create App Container
        self.container = Frame(master, background="white")
        self.container.pack()

        # Create Logo Canvas
        self.logo_canvas = Canvas(self.container, width=280, height=70, background="white", highlightthickness=0,
                                  borderwidth=0)
        self.logo_canvas.grid(row=0, column=12, columnspan=6, rowspan=1, pady=10)
        self.logo_canvas.create_image(5, 5, anchor='nw', image=self.img_logo)

        # Create Report Frame & Label
        self.report_buttons = LabelFrame(self.container, text='Options', height=420, width=300, background="white",
                                         font=myFont)
        self.report_buttons.grid_columnconfigure((0, 4), weight=1)
        self.report_buttons.grid_propagate(0)
        self.report_buttons.grid(row=1, column=12, columnspan=6, rowspan=3, padx=5, pady=(0, 10))

        # Create Report Options & Buttons
        self.month_select_frame = LabelFrame(self.report_buttons, text="", width=290, height=145, background="white",
                                             highlightthickness=0, borderwidth=0, font=myFont)
        self.month_select_frame.grid_columnconfigure((0, 1), weight=1, uniform="half")
        self.month_select_frame.grid_propagate(0)
        self.month_select_frame.grid(row=1, column=1, columnspan=2, rowspan=1, sticky="we")

        self.custom_select_frame = LabelFrame(self.report_buttons, text="", width=290, height=90, background="white",
                                              highlightthickness=0, borderwidth=0, font=myFont)
        self.custom_select_frame.grid_columnconfigure(0, weight=1)
        self.custom_select_frame.grid_propagate(0)
        self.custom_select_frame.grid(row=2, column=1, columnspan=2, rowspan=1, sticky="we")

        self.report_select_unit_court_label = Label(self.month_select_frame, text="Unit Court:", anchor="e",
                                                    background="white", font=myFont)
        self.report_select_unit_court_label.grid(row=0, column=0, columnspan=1, rowspan=1, padx=5, sticky="we")

        self.report_select_unit_court_box = ttk.Combobox(self.month_select_frame, state="readonly",
                                                         textvariable=Mainwindow.month_report_court_selected_val,
                                                         background="white", font=myFont)
        self.report_select_unit_court_box.grid(row=0, column=1, columnspan=1, rowspan=1, padx=(5, 10), sticky="we")
        self.month_select_label = Label(self.month_select_frame, text="Month:", anchor="e", background="white",
                                        font=myFont)
        self.month_select_label.grid(row=1, column=0, columnspan=1, rowspan=1, pady=5, padx=5, sticky="we")
        self.month_select = ttk.Combobox(self.month_select_frame, state='readonly',
                                         textvariable=Mainwindow.month_report_month_selected_val, background="white",
                                         font=myFont)
        self.month_select["values"] = ['_Show All', 'January', 'February', 'March', 'April', 'May', 'June', 'July',
                                       'August', 'September', 'October', 'November', 'December']
        self.month_select.grid(row=1, column=1, columnspan=1, rowspan=1, padx=(5, 10), sticky="we")
        self.generate_month_button = Button(self.month_select_frame, text="Generate Monthly Report",
                                            height=button_height, background="cornflower blue", font=myFont,
                                            command=lambda: ReportWindow(master, "Monthly"))
        self.generate_month_button.grid(row=2, column=0, columnspan=2, rowspan=1, pady=2, padx=7, sticky="we")
        self.generate_expired_button = Button(self.month_select_frame, text="Generate Overdue Report",
                                              height=button_height, background="cornflower blue", fg='red', font=myFont,
                                              command=lambda: ReportWindow(master, "Overdue"))
        self.generate_expired_button.grid(row=3, column=0, columnspan=2, rowspan=1, pady=2, padx=7, sticky="we")
        self.custom_select_label = Label(self.custom_select_frame, text="Custom Report Criteria:", anchor='w',
                                         background="white", font=myFont)
        self.custom_select_label.grid(row=0, column=0, columnspan=1, rowspan=1, padx=10, sticky="we")
        self.custom_select = ttk.Combobox(self.custom_select_frame, state='readonly', background="white", font=myFont,
                                          textvariable=Mainwindow.custom_report_selected)
        self.custom_select.grid(row=1, column=0, columnspan=1, rowspan=1, padx=10, sticky="we")
        self.generate_custom_button = Button(self.custom_select_frame, text="Generate Custom Report",
                                             height=button_height, background="cornflower blue", font=myFont,
                                             command=lambda: ReportWindow(master, "Custom"))
        self.generate_custom_button.grid(row=2, column=0, columnspan=1, rowspan=1, pady=2, padx=7, sticky="we")

        # Create Record Frame & Label
        self.modify_record = LabelFrame(self.report_buttons, text="", height=150, background="white", font=myFont,
                                        highlightthickness=0, borderwidth=0)
        self.modify_record.grid_columnconfigure((1, 2), weight=1)
        self.modify_record.grid_propagate(0)
        self.modify_record.grid(row=3, column=1, columnspan=2, rowspan=1, sticky="we")

        # Create Dashboard Frame
        self.dashboard_style = ttk.Style()
        self.dashboard_style.configure("TNotebook", background="white", borderwidth=0, border="black",
                                       highlightthickness=0, padding=0, highlightcolor="white",
                                       highlightbackground="white", takefocus="")
        self.dashboard_style.configure("TNotebook.Tab", borderwidth=1, highlightthickness=0, font=myFont,
                                       background="white", highlightbackground="white", highlightcolor="white",
                                       takefocus="")
        self.dashboard_style.configure("TProgressbar", background="white", highlightbackground="white",
                                       troughcolor="white", takefocus=False)
        self.dashboard_style.configure("Treeview", takefocus=False, font=myFont)
        self.dashboard_style.configure("Treeview.Heading", font=myFont)
        self.dashboard_frame = ttk.Notebook(self.container, height=dash_height, width=dash_width, style="TNotebook",
                                            takefocus=False)
        self.dashboard_frame.grid_propagate(0)
        self.dashboard_frame.grid_columnconfigure((1, 2, 3), weight=2)
        self.dashboard_frame.grid_columnconfigure((0, 4), weight=1)
        self.dashboard_summary_tab = Frame(self.dashboard_frame, borderwidth=0, highlightthickness=0,
                                           background="white", highlightbackground="white", highlightcolor="white",
                                           relief="flat")
        self.dashboard_summary_tab.grid_columnconfigure((1, 2, 3), weight=2, uniform="active")
        self.dashboard_summary_tab.grid_columnconfigure((0, 4), weight=1, uniform="space")
        self.dashboard_summary_tab.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9), weight=1, uniform="rows")
        self.dashboard_frame.add(self.dashboard_summary_tab, text="Summary")
        self.dashboard_house_tab = Frame(self.dashboard_frame, borderwidth=0, highlightthickness=0, background="white",
                                         highlightbackground="white", highlightcolor="white", relief="flat")
        self.dashboard_house_tab.grid_columnconfigure((1, 2, 3), weight=2, uniform="active")
        self.dashboard_house_tab.grid_columnconfigure((0, 4), weight=1, uniform="space")
        self.dashboard_house_tab.grid_rowconfigure((1, 2, 3, 4, 5, 6, 7, 8, 9, 10), weight=1, uniform="rows")
        self.dashboard_frame.add(self.dashboard_house_tab, text="Unit House")
        self.dashboard_frame.grid(row=0, column=0, columnspan=12, rowspan=4, pady=10, padx=5, sticky="NESW")

        # Create Record Buttons
        self.view_edit_tenant = Button(self.modify_record, text="View / Edit Details", background="cornflower blue",
                                       height=button_height, font=myFont,
                                       command=lambda: ViewEditTenant(master, "Main", self))
        self.view_edit_tenant.grid(row=0, column=1, columnspan=2, rowspan=1, pady=(5, 2), padx=7, sticky="we")
        self.search_tenant_button = Button(self.modify_record, text="Search Units", background="cornflower blue",
                                           height=button_height, font=myFont, command=lambda: SearchTenant(master))
        self.search_tenant_button.grid(row=1, column=1, columnspan=2, rowspan=1, pady=2, padx=7, sticky="ew")
        self.view_edit_unit = Button(self.modify_record, text="View / Edit Unit", font=myFont, state="disabled",
                                     height=button_height, background="cornflower blue",
                                     command=lambda: ViewEditUnit(master))
        self.view_edit_unit.grid(row=2, column=1, columnspan=2, rowspan=1, pady=2, padx=7, sticky="we")
        self.add_new_unit = Button(self.modify_record, text="Add New Unit", state="disabled", height=button_height,
                                   background="cornflower blue", font=myFont, command=lambda: AddNewUnit(master))
        self.add_new_unit.grid(row=3, column=1, columnspan=2, rowspan=1, pady=2, padx=7, sticky="we")

        # Create Quit Button
        self.log_out_button = Button(self.container, text="Log Out", height=button_height, background="cornflower blue",
                                     font=myFont, command=lambda: self.log_out(master))
        self.log_out_button.grid(row=7, column=12, columnspan=3, rowspan=1, padx=10, sticky="we")
        self.exit_button = Button(self.container, text="Exit", height=button_height, font=myFont,
                                  background="cornflower blue", command=root.quit)
        self.exit_button.grid(row=7, column=15, columnspan=3, rowspan=1, padx=10, sticky="we")

        # Create dashboard
        self.summary_table_frame = LabelFrame(self.dashboard_summary_tab, text="", background="white", font=myFont,
                                              highlightthickness=0, borderwidth=0)
        self.summary_table_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1)
        self.summary_table_frame.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9), weight=1, uniform="row")
        self.summary_table_frame.grid(row=0, column=1, columnspan=3, rowspan=2, padx=11, pady=10, sticky="nswe")

        # User account display
        self.user_account_label = Label(self.container, background="white", font=myFont)
        self.user_account_label.grid(row=7, column=11)
        self.user_settings = Button(self.container, background="cornflower blue", font=myFont, height=button_height,
                                    state="disabled", text="User Settings", command=lambda: UserSettings(master))
        self.user_settings.grid(row=7, column=10, sticky="nsew")
        self.refresh_button = Button(self.container, text="Refresh", height=button_height, font=myFont, takefocus=False,
                                     background="cornflower blue", command=lambda: self.refresh_dashboard())
        self.refresh_button.grid(row=7, column=0, sticky="nsew", padx=(10, 0))
        self.last_refreshed_time = Label(self.container, text="Last Refresh:", background="white", font=myFont,
                                         anchor="w")
        self.last_refreshed_time.grid(row=7, column=1, rowspan=1, columnspan=9, sticky="w")

        # dashboard House Tab
        self.summary_table_house_frame = LabelFrame(self.dashboard_house_tab, text="", height=180, background="white",
                                                    font=myFont, highlightthickness=1, borderwidth=1)
        self.summary_table_house_frame.grid_columnconfigure(0, weight=1)
        self.summary_table_house_frame.grid_rowconfigure(0, weight=1)
        self.summary_table_house_frame.grid(row=1, column=0, columnspan=5, rowspan=10, sticky="nswe")

        self.summary_house_columns = ["Unit Court / House", "In Compliance", "Overdue", "<30 Days", "30-60 Days"]

        self.summary_house_treeview = ttk.Treeview(self.summary_table_house_frame, columns=self.summary_house_columns,
                                                   show=["headings", "tree"])
        self.minwidth = self.summary_house_treeview.column('#0', option='minwidth')
        self.summary_house_treeview.column('#0', width=self.minwidth)
        self.summary_house_treeview.tag_configure("Top", font=myFontBold)
        self.summary_house_treeview.grid(row=0, column=0, sticky="nsew")
        self.house_vert_scroll_bar = ttk.Scrollbar(self.summary_table_house_frame, orient="vertical",
                                                   command=self.summary_house_treeview.yview)
        self.house_vert_scroll_bar.grid(row=0, column=1, sticky='ns')
        self.summary_house_treeview.configure(yscrollcommand=self.house_vert_scroll_bar.set)

        for col in self.summary_house_columns:
            self.summary_house_treeview.heading(col, text=col)
            if col == "Unit Court / House":
                self.summary_house_treeview.column(col, width=150)
            else:
                self.summary_house_treeview.column(col, width=80)

    @Database.autoreconnect_retry
    def create_house_summary(self):

        self.summary_house_treeview.delete(*self.summary_house_treeview.get_children())
        self.summary_unit_court_list = Database.crud_app_collection.distinct("Unit_Court")

        # Declare variables for 30 and days days from current date
        self.thirty_days = datetime.now() + timedelta(days=30)
        self.sixty_days = datetime.now() + timedelta(days=60)

        for i in range(len(self.summary_unit_court_list)):
            self.court_values_list = [Database.crud_app_collection.count_documents(
                {'Unit_Court': self.summary_unit_court_list[i], 'Service_Due_ISO': {'$gt': datetime.now()}}),
                Database.crud_app_collection.count_documents(
                    {'Unit_Court': self.summary_unit_court_list[i],
                     'Service_Due_ISO': {'$lt': datetime.now()}}),
                Database.crud_app_collection.count_documents(
                    {'Unit_Court': self.summary_unit_court_list[i],
                     'Service_Due_ISO': {'$gt': datetime.now(), '$lt': self.thirty_days}}),
                Database.crud_app_collection.count_documents(
                    {'Unit_Court': self.summary_unit_court_list[i],
                     'Service_Due_ISO': {'$gte': self.thirty_days, '$lt': self.sixty_days}})]
            self.summary_house_treeview.insert("", i, self.summary_unit_court_list[i],
                                               values=[self.summary_unit_court_list[i], self.court_values_list[0],
                                                       self.court_values_list[1], self.court_values_list[2],
                                                       self.court_values_list[3]], tags='Top')

            self.summary_house_list = Database.crud_app_collection.distinct("Unit_House", {
                "Unit_Court": self.summary_unit_court_list[i]})

            for h in range(len(self.summary_house_list)):
                self.house_values_list = [Database.crud_app_collection.count_documents(
                    {'Unit_House': self.summary_house_list[h], 'Service_Due_ISO': {'$gt': datetime.now()}}),
                    Database.crud_app_collection.count_documents(
                        {'Unit_House': self.summary_house_list[h],
                         'Service_Due_ISO': {'$lt': datetime.now()}}),
                    Database.crud_app_collection.count_documents(
                        {'Unit_House': self.summary_house_list[h],
                         'Service_Due_ISO': {'$gt': datetime.now(), '$lt': self.thirty_days}}),
                    Database.crud_app_collection.count_documents(
                        {'Unit_House': self.summary_house_list[h],
                         'Service_Due_ISO': {'$gte': self.thirty_days, '$lt': self.sixty_days}})]
                self.summary_house_treeview.insert(self.summary_unit_court_list[i], h,
                                                   values=[self.summary_house_list[h], self.house_values_list[0],
                                                           self.house_values_list[1], self.house_values_list[2],
                                                           self.house_values_list[3]])
                self.dashboard_progress.step(2)

    def refresh_dashboard(self):
        def refresh_process():
            self.dashboard_progress.step(5)
            self.create_dashboard()
            self.dashboard_progress.step(5)
            self.create_house_summary()
            self.dashboard_progress.step(10)
            self.refresh_button.update()
            self.refresh_button.config(state="normal")
            self.log_out_button.config(state="normal")
            self.exit_button.config(state="normal")
            self.dashboard_progress.grid_forget()
            self.last_refreshed_time["text"] = "Last Refresh: " + datetime.now().strftime('%H:%M:%S')
            self.last_refreshed_time.grid(row=7, column=1, rowspan=1, columnspan=9, sticky="we", padx=(5, 0))

        self.last_refreshed_time.grid_forget()
        self.dashboard_progress = ttk.Progressbar(self.container, length=100, orient=HORIZONTAL, mode='determinate')
        self.dashboard_progress.grid(row=7, column=1, rowspan=1, columnspan=9, sticky="NSWE", pady=2)
        self.refresh_button.config(state="disabled")
        self.log_out_button.config(state="disabled")
        self.exit_button.config(state="disabled")
        self.refresh_button.update()
        self.extra_thread = threading.Thread(target=refresh_process)
        self.extra_thread.daemon = True
        self.extra_thread.start()

    @staticmethod
    def log_out(master):

        # Handle log out from Mainwindow
        # Hide root Window
        root.withdraw()
        # Close Database connection
        Database.database_client.close()
        # Reset Username and Password Variables
        Database.username_val = ""
        Database.password_val = ""
        # Display Login Window
        Login(master)

    @Database.autoreconnect_retry
    def create_dashboard(self):
        # Code for Creating Table

        # Declare variables for 30 and days days from current date
        self.thirty_days = datetime.now() + timedelta(days=30)
        self.sixty_days = datetime.now() + timedelta(days=60)
        self.dash_unit_court_list = Database.crud_app_collection.distinct("Unit_Court")

        self.summary_data = [("Unit Court", "In Compliance", "Overdue", "<30 Days", "30-60 Days")]
        for court in self.dash_unit_court_list:
            self.summary_data.append((court, Database.crud_app_collection.count_documents(
                                 {'Unit_Court': court, 'Service_Due_ISO': {'$gt': datetime.now()}}),
                              Database.crud_app_collection.count_documents(
                                  {'Unit_Court': court, 'Service_Due_ISO': {'$lt': datetime.now()}}),
                              Database.crud_app_collection.count_documents({'Unit_Court': court,
                                                                               'Service_Due_ISO': {
                                                                                   '$gt': datetime.now(),
                                                                                   '$lt': self.thirty_days}}),
                              Database.crud_app_collection.count_documents({'Unit_Court': court,
                                                                               'Service_Due_ISO': {
                                                                                   '$gte': self.thirty_days,
                                                                                   '$lt': self.sixty_days}})))

        self.summary_data.append(("Totals", Database.crud_app_collection.count_documents(
                                 {'Service_Due_ISO': {'$gt': datetime.now()}}),
                              Database.crud_app_collection.count_documents(
                                  {'Service_Due_ISO': {'$lt': datetime.now()}}),
                              Database.crud_app_collection.count_documents(
                                  {'Service_Due_ISO': {'$gt': datetime.now(), '$lt': self.thirty_days}}),
                              Database.crud_app_collection.count_documents(
                                  {'Service_Due_ISO': {'$gt': self.thirty_days, '$lt': self.sixty_days}})))



        # Create table using Tkinter Entry widgets
        for r in range(10):
            if r == 1 or r == 8:
                self.head_sep = ttk.Separator(self.summary_table_frame)
                self.head_sep.grid(row=r, column=0, columnspan=5, sticky="ew")
            else:
                for c in range(5):
                    if r == 0 or r == 9:
                        # Set Font options for Header row of table
                        self.e = Entry(self.summary_table_frame, width=12, justify="center",
                                       disabledforeground="cornflower blue", disabledbackground="white",
                                       font=myFontBold, borderwidth=0)
                    elif r > 0 and c > 0:
                        # Set font options for general table cells
                        self.e = Entry(self.summary_table_frame, width=12, justify="center", disabledforeground="black",
                                       disabledbackground="white", font=myFont, borderwidth=0)
                    else:
                        self.e = Entry(self.summary_table_frame, width=12, disabledforeground="black",
                                       disabledbackground="white", font=myFont, borderwidth=0)
                    if c == 2:
                        self.e.config(disabledforeground="red")
                    self.e.grid(row=r, column=c, sticky="nesw")
                    if r in {2, 3, 4, 5, 6, 7, 8}:
                        self.e.insert(END, self.summary_data[r - 1][c])
                    elif r == 9:
                        self.e.insert(END, self.summary_data[r - 2][c])
                    else:
                        self.e.insert(END, self.summary_data[r][c])
                    self.e["state"] = "disabled"

    @Database.autoreconnect_retry
    def populate_combo_box(self):
        # Populate combobox with distinct values from database.
        self.unit_court_select_list = ["_Show All"] + Database.crud_app_collection.distinct("Unit_Court")
        self.report_select_unit_court_box["values"] = self.unit_court_select_list

        self.custom_select_unit_courts = []
        for court in Database.crud_app_collection.distinct("Unit_Court"):
            self.custom_select_unit_courts.append("Show Units: " + court)

        self.custom_select["values"] = ['_Show All Units', 'No Valid Email', 'No Valid Phone',
                                        'With Valid Email'] + self.custom_select_unit_courts

        # Show which user is currently logged in
        self.user_account_label["text"] = "User:   " + Database.username_val


# noinspection PyUnresolvedReferences
class ReportWindow:
    @Database.autoreconnect_retry
    def __init__(self, master, report_type):

        self.report_container = Toplevel(master)
        self.report_container.title("Report Viewer")
        self.report_container.geometry("1200x620")
        self.report_container.resizable(False, False)
        self.report_container.config(background="white")
        self.report_container.grid_columnconfigure(
            (0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19), weight=1, uniform="col")
        self.report_container.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12), weight=1, uniform="row")
        self.report_container.focus_force()
        self.master = master

        # Gets the requested values of the height and width.
        self.windowWidth = 1200
        self.windowHeight = 620
        # Gets both half the screen width/height and window width/height
        self.positionRight = int(root.winfo_screenwidth() / 2 - self.windowWidth / 2)
        self.positionDown = int(root.winfo_screenheight() / 2 - self.windowHeight / 2)
        # Positions the window in the center of the page.
        self.report_container.geometry("+{}+{}".format(self.positionRight, self.positionDown))

        self.report_name = StringVar()
        self.report_name.set("")
        self.report_label = Label(self.report_container, textvariable=self.report_name, background="white", font=myFont,
                                  width=40, anchor="w")
        self.report_label.config(font=("Arial", -14, 'bold'))
        self.report_label.grid(row=0, column=0, columnspan=10, rowspan=1, padx=20, pady=5)

        self.report_count_val = StringVar()
        self.report_count_val.set("Results:")
        self.report_count_label = Label(self.report_container, textvariable=self.report_count_val, background="white",
                                        font=myFont, width=40, anchor="e")
        self.report_count_label.config(font=("Arial", -14, 'bold'))
        self.report_count_label.grid(row=0, column=10, columnspan=10, rowspan=1, pady=5)

        self.report_frame = Frame(self.report_container, highlightthickness=1, borderwidth=1)
        self.report_frame.grid_propagate(0)
        self.report_frame.grid_columnconfigure(0, weight=1)
        self.report_frame.grid_rowconfigure(0, weight=1)
        self.report_frame.grid(row=1, column=0, columnspan=20, rowspan=11, padx=10, pady=5, sticky="nsew")

        self.email_all_results_button = Button(self.report_container, text="Copy All Email Addresses", font=myFont,
                                               height=button_height, background="cornflower blue",
                                               command=lambda: self.email_result_tenants())
        self.email_all_results_button.grid(row=12, column=0, columnspan=4, padx=(10, 0), sticky="we")

        self.copy_selected_emails = Button(self.report_container, text="Copy Selected Email Addresses", font=myFont,
                                           height=button_height, background="cornflower blue",
                                           command=lambda: self.copy_selected())
        self.copy_selected_emails.grid(row=12, column=4, columnspan=4, rowspan=1, sticky="ew", padx=(10, 0))

        self.email_button = Button(self.report_container, text="Email Report", height=button_height,
                                   background="cornflower blue", state="disabled", font=myFont)
        self.email_button.grid(row=12, column=14, columnspan=2, rowspan=1, pady=5, padx=(0, 10), sticky="we")

        self.save_button = Button(self.report_container, text="Export", height=button_height,
                                  background="cornflower blue", font=myFont, command=lambda: self.save_report())
        self.save_button.grid(row=12, column=16, columnspan=2, rowspan=1, pady=5, padx=(0, 10), sticky="we")

        self.close_button = Button(self.report_container, text="Close", height=button_height,
                                   background="cornflower blue", font=myFont, command=lambda: self.close_window())
        self.close_button.grid(row=12, column=18, columnspan=2, rowspan=1, pady=5, padx=(0, 10), sticky="we")

        # Hide root window while Toplevel open
        root.withdraw()
        # Handle "x" on Toplevel
        self.report_container.protocol("WM_DELETE_WINDOW", self.close_window)

        self.query_filter = {}
        self.months_values = {"January": 1, "February": 2, "March": 3, "April": 4, "May": 5, "June": 6, "July": 7,
                              "August": 8, "September": 9, "October": 10, "November": 11, "December": 12}
        # Declare variables for 30 and days days from current date
        self.thirty_days = datetime.now() + timedelta(days=30)
        self.sixty_days = datetime.now() + timedelta(days=60)
        self.recent_overdue = datetime.now() - timedelta(days=30)

        # Create the report query filter
        if report_type == "Monthly":

            self.report_contain = ""
            self.report_month = ""
            self.report_year = ""

            if Mainwindow.month_report_court_selected_val.get() != "_Show All":
                self.query_filter["Unit_Court"] = Mainwindow.month_report_court_selected_val.get()
                self.report_contain = Mainwindow.month_report_court_selected_val.get()
            else:
                self.report_contain = "All Courts"

            if Mainwindow.month_report_month_selected_val.get() != "_Show All":
                self.current_month = datetime.now().month
                self.current_year = datetime.now().year

                if self.months_values[Mainwindow.month_report_month_selected_val.get()] < self.current_month:
                    self.query_filter["Service_Due_Month"] = self.months_values[
                        Mainwindow.month_report_month_selected_val.get()]
                    self.query_filter["Service_Due_Year"] = self.current_year + 1
                    self.report_month = Mainwindow.month_report_month_selected_val.get()
                    self.report_year = str(self.current_year + 1)
                else:
                    self.query_filter["Service_Due_Month"] = self.months_values[
                        Mainwindow.month_report_month_selected_val.get()]
                    self.query_filter["Service_Due_Year"] = self.current_year
                    self.query_filter["Service_Due_ISO"] = {"$gt": datetime.now()}
                    self.report_month = Mainwindow.month_report_month_selected_val.get()
                    self.report_year = str(self.current_year)
            else:
                self.query_filter["Service_Due_ISO"] = {"$gt": datetime.now()}
                self.report_month = "All Future Dates"

            # Name report based on query type
            self.report_name.set(
                "Report for Units, " + self.report_contain + ", " + self.report_month + " " + self.report_year)

        elif report_type == "Overdue":
            self.report_contain = ""
            self.report_month = ""
            self.report_year = ""

            if Mainwindow.month_report_court_selected_val.get() != "_Show All":
                self.query_filter["Unit_Court"] = Mainwindow.month_report_court_selected_val.get()
                self.report_contain = Mainwindow.month_report_court_selected_val.get()
            else:
                self.report_contain = "All Courts"

            self.query_filter["Service_Due_ISO"] = {"$lt": datetime.now()}
            self.report_name.set("Report for Overdue Units, " + self.report_contain)

        elif report_type == "Custom":
            if Mainwindow.custom_report_selected.get() == "No Valid Email":
                self.query_filter["Email_1"] = ""
                self.report_name.set("Report for, Units without Emails")

            elif Mainwindow.custom_report_selected.get() == "No Valid Phone":
                self.query_filter["Phone"] = ""
                self.report_name.set("Report for, Units without Phone Numbers")

            elif Mainwindow.custom_report_selected.get() == "_Show All Units":
                self.report_name.set("Report for, Show All Units")

            elif Mainwindow.custom_report_selected.get() == "With Valid Email":
                self.query_filter["Email_1"] = {"$exists": "true", "$ne": ""}
                self.report_name.set("Report for, Units with Emails")

            elif Mainwindow.custom_report_selected.get().startswith("Show Units:"):
                self.court = Mainwindow.custom_report_selected.get().split(": ")
                self.query_filter["Unit_Court"] = self.court[1]
                self.report_name.set("Report for, " + self.court[1])

        elif report_type == "Reminder Long Overdue":
            self.query_filter["Service_Due_ISO"] = {'$lt': self.recent_overdue}
            self.report_name.set("Reminders for Long Overdue")
        elif report_type == "Reminder Overdue":
            self.query_filter["Service_Due_ISO"] = {'$lt': datetime.now(), '$gte': self.recent_overdue}
            self.report_name.set("Reminders for Recently Overdue")
        elif report_type == "Reminder Thirty":
            self.query_filter["Service_Due_ISO"] = {'$gt': datetime.now(), '$lt': self.thirty_days}
            self.report_name.set("Reminders for Thirty Days")
        elif report_type == "Reminder Sixty":
            self.query_filter["Service_Due_ISO"] = {'$gte': self.thirty_days, '$lt': self.sixty_days}
            self.report_name.set("Reminders for Sixty Days")

        Database.log_action("Report: " + report_type, str(self.query_filter))

        # Get Collection Columns from database
        self.db_columns = ["Unit", "Unit_Court", "Unit_House", "Name", "Phone", "Email_1", "Last_Serviced",
                           "Service_Due"]

        self.populate_results_table()

    def populate_results_table(self):

        try:
            self.query_results = list(Database.crud_app_collection.find(self.query_filter).sort(
                [("Unit_Court", pymongo.ASCENDING), ("Unit", pymongo.ASCENDING)]))
        except pymongo.errors.PyMongoError as e:
            messagebox.showerror("Exception", e, parent=self.report_container)
            self.close_window()

        else:
            # Populate Treeview with filtered data
            self.results_table_view = ttk.Treeview(self.report_frame, columns=self.db_columns, show="headings",
                                                   height=32)
            for col in self.db_columns:
                self.results_table_view.heading(col, text=col)
            self.results_table_view.column("Unit", width=35, stretch=NO)
            self.results_table_view.column("Unit_Court", width=100)
            self.results_table_view.column("Unit_House", width=100)
            self.results_table_view.column("Name", width=280)
            self.results_table_view.column("Phone", width=90)
            self.results_table_view.column("Email_1", width=200)
            self.results_table_view.column("Last_Serviced", width=90)
            self.results_table_view.column("Service_Due", width=90)
            self.results_table_view.grid(row=0, column=0, sticky="nsew")
            self.report_frame.grid_columnconfigure(0, weight=1)
            self.vert_scroll_bar = ttk.Scrollbar(self.report_frame, orient="vertical",
                                                 command=self.results_table_view.yview)
            self.vert_scroll_bar.grid(row=0, column=10, sticky='ns')
            self.results_table_view.configure(yscrollcommand=self.vert_scroll_bar.set)

            self.results_count = 0
            for i in range(len(self.query_results)):
                self.results_count = self.results_count + 1
                self.serviced_last = "-"
                self.serviced_due = "-"
                if datetime.strftime(self.query_results[i]["Last_Serviced_ISO"], "%d/%m/%Y") != "01/01/1970":
                    self.serviced_last = datetime.strftime(self.query_results[i]["Last_Serviced_ISO"], "%d/%m/%Y")
                    self.serviced_due = datetime.strftime(self.query_results[i]["Service_Due_ISO"], "%d/%m/%Y")
                self.results_table_view.insert("", i, values=[self.query_results[i]["Unit"],
                                                              self.query_results[i]["Unit_Court"],
                                                              self.query_results[i]["Unit_House"],
                                                              self.query_results[i]["Name"],
                                                              self.query_results[i]["Phone"],
                                                              self.query_results[i]["Email_1"], self.serviced_last,
                                                              self.serviced_due])

            self.report_count_val.set("Results: " + str(self.results_count))
            self.results_table_view.bind('<Double-Button-1>', self.select_item)

    def copy_selected(self):

        selected_items = self.results_table_view.selection()
        emails_selected = []
        if len(selected_items) != 0:
            for item in selected_items:
                if self.results_table_view.item(item)["values"][5] != "":
                    emails_selected.append(self.results_table_view.item(item)["values"][5])
            if len(emails_selected) != 0:
                selected_results = pd.DataFrame(emails_selected)
                selected_results.to_clipboard(excel=True, sep=',', index=False, header=False)
                if len(emails_selected) == 1:
                    messagebox.showinfo("Clipboard", str(len(emails_selected)) + " Email Address Copied to Clipboard",
                                        parent=self.report_container)
                elif len(emails_selected) > 1:
                    messagebox.showinfo("Clipboard", str(len(emails_selected)) + " Email Addresses Copied to Clipboard",
                                        parent=self.report_container)
                print(selected_results)

    def select_item(self, *args):
        cur_item = self.results_table_view.focus()
        ReportWindow.unit_num_selected = self.results_table_view.item(cur_item)["values"][0]
        ReportWindow.unit_court_selected = self.results_table_view.item(cur_item)["values"][1]

        ViewEditTenant(self.master, "Report", self)
        self.report_container.withdraw()

    def close_window(self):
        # Handle closing Toplevel event.
        self.report_container.destroy()
        # Make root window visible on close
        root.deiconify()

    def save_report(self):

        save_directory = filedialog.asksaveasfilename(initialfile=self.report_name.get(), defaultextension=".csv")

        if save_directory != '':
            self.report_list = self.query_results
            for row in self.report_list:
                del row["_id"]
                del row["Unit_Full"]
                del row["RTM_group"]
                del row["Ownership"]
                del row["Director"]
                del row["Service_Due_Month"]
                del row["Service_Due_Year"]
                row["Notes"] = row["Notes"].replace('\n', ' ')
                row["Notes"] = row["Notes"].replace('\t', ' ')
                if datetime.strftime(row["Last_Serviced_ISO"], "%d/%m/%Y") != "01/01/1970":
                    row["Service_Due_ISO"] = datetime.strftime(row["Service_Due_ISO"], "%d/%m/%Y")
                    row["Last_Serviced_ISO"] = datetime.strftime(row["Last_Serviced_ISO"], "%d/%m/%Y")
                else:
                    row["Service_Due_ISO"] = ""
                    row["Last_Serviced_ISO"] = ""

            self.csv_keys = self.report_list[0].keys()
            with open(save_directory, 'w', newline='') as output_file:
                dict_writer = csv.DictWriter(output_file, fieldnames=self.csv_keys)
                dict_writer.writeheader()
                dict_writer.writerows(self.report_list)

    def email_result_tenants(self):

        if len(self.query_results) != 0:
            email_to_copy = pd.DataFrame(self.query_results)
            email_to_copy['Email_1'].replace("", np.nan, inplace=True)
            email_to_copy.dropna(subset=['Email_1'], inplace=True)
            email_to_copy['Email_1'].to_clipboard(excel=True, sep=',', index=False, header=False)
            if len(email_to_copy) == 1:
                messagebox.showinfo("Clipboard", str(len(email_to_copy)) + " Email Address Copied to Clipboard",
                                    parent=self.report_container)
            elif len(email_to_copy) > 1:
                messagebox.showinfo("Clipboard", str(len(email_to_copy)) + " Email Addresses Copied to Clipboard",
                                    parent=self.report_container)


# noinspection PyUnresolvedReferences
class ViewEditTenant:
    def __init__(self, master, came_from, source_window):

        # Populate Combobox from distinct "Unit_Court" Values
        self.unit_court_list = Database.crud_app_collection.distinct("Unit_Court")
        self.unit_house_list = Database.crud_app_collection.distinct("Unit_House")
        self.came_from = came_from
        self.source_window = source_window

        # Declare Variables and Trace changes
        self.unit_court_selected = StringVar()
        self.unit_selected = StringVar()
        self.unit_house_selected = StringVar()

        self.unit_house_selected.set("")

        self.unit_court_trace = self.unit_court_selected.trace_add("write", lambda *args,
                                                                                   unit_filter="court": self.update_unit_list(
            unit_filter))
        self.unit_house_trace = self.unit_house_selected.trace_add("write", lambda *args,
                                                                                   unit_filter="house": self.update_unit_list(
            unit_filter))

        # Create Tkinter Toplevel widget
        self.view_edit_container = Toplevel(master)
        self.view_edit_container.title("Century Wharf - View / Edit Details")
        self.view_edit_container.geometry("1000x620")
        self.view_edit_container.resizable(False, False)
        self.view_edit_container.config(background="white")
        self.view_edit_container.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight=1, uniform="cols")
        self.view_edit_container.grid_rowconfigure(
            (0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20), weight=1, uniform="rows")
        self.view_edit_container.focus_force()

        # Gets the requested values of the height and width.
        self.windowWidth = 1000
        self.windowHeight = 620
        # Gets both half the screen width/height and window width/height
        self.positionRight = int(root.winfo_screenwidth() / 2 - self.windowWidth / 2)
        self.positionDown = int(root.winfo_screenheight() / 2 - self.windowHeight / 2)
        # Positions the window in the center of the page.
        self.view_edit_container.geometry("+{}+{}".format(self.positionRight, self.positionDown))

        # Initiate widget with edit mode inactive
        self.edit_mode = False
        self.cert_edit_mode = False

        # Create Toplevel frames
        self.select_unit_frame = LabelFrame(self.view_edit_container, text="Unit Details", background="white",
                                            font=myFont)
        self.select_unit_frame.grid_propagate(0)
        self.select_unit_frame.grid_rowconfigure((0, 1, 2, 3, 4), weight=1)
        self.select_unit_frame.grid(row=0, column=0, columnspan=4, rowspan=7, padx=(10, 5), pady=(10, 0), sticky="NSEW")

        self.show_tenant_frame = LabelFrame(self.view_edit_container, text="Owner Details", background="white",
                                            font=myFont)
        self.show_tenant_frame.grid_propagate(0)
        self.show_tenant_frame.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11), weight=1)
        self.show_tenant_frame.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.show_tenant_frame.grid(row=7, column=0, columnspan=4, rowspan=14, padx=(10, 5), pady=(10, 0),
                                    sticky="NSEW")

        self.show_certificate_frame = LabelFrame(self.view_edit_container, text="Boiler Service Details",
                                                 background="white", font=myFont)
        self.show_certificate_frame.grid_propagate(0)
        self.show_certificate_frame.grid_rowconfigure((0, 1, 2, 3), weight=1, uniform="rows")
        self.show_certificate_frame.grid_columnconfigure((0, 1, 2), weight=1, uniform="cols")
        self.show_certificate_frame.grid(row=0, column=4, columnspan=4, rowspan=5, padx=(5, 10), pady=(10, 0),
                                         sticky="NSEW")

        self.rtm_member_frame = LabelFrame(self.view_edit_container, text="RTM Details", background="white",
                                           font=myFont)
        self.rtm_member_frame.grid_propagate(0)
        self.rtm_member_frame.grid_columnconfigure((0, 1, 2, 3, 4), weight=1, uniform='cols')
        self.rtm_member_frame.grid_rowconfigure((0, 1), weight=1, uniform="rows")
        self.rtm_member_frame.grid(row=5, column=4, columnspan=4, rowspan=4, padx=(5, 10), pady=(10, 0), sticky="NSEW")

        self.unit_files_frame = LabelFrame(self.view_edit_container, text="Files (10 Max)", background="white",
                                           font=myFont)
        self.unit_files_frame.grid_propagate(0)
        self.unit_files_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10), weight=1, uniform="cols")
        self.unit_files_frame.grid_rowconfigure((0, 1, 2), weight=1, uniform="rows")
        self.unit_files_frame.grid(row=9, column=4, columnspan=4, rowspan=5, sticky="NSEW", padx=(5, 10), pady=(10, 0))

        self.show_notes_frame = LabelFrame(self.view_edit_container, text="Notes", background="white", font=myFont)
        self.show_notes_frame.grid_propagate(0)
        self.show_notes_frame.grid_columnconfigure(0, weight=1)
        self.show_notes_frame.grid_rowconfigure(0, weight=1)
        self.show_notes_frame.grid(row=14, column=4, columnspan=4, rowspan=7, padx=(5, 10), pady=(10, 0), sticky="NSEW")

        self.show_notes_input = Text(self.show_notes_frame, width=52, height=18, state="disabled", background="white",
                                     font=myFont, highlightthickness=0, borderwidth=0)
        self.show_notes_input.grid(row=0, column=0, columnspan=1, rowspan=1, sticky="nsew")

        self.notes_scroll_bar = ttk.Scrollbar(self.show_notes_frame, orient="vertical",
                                              command=self.show_notes_input.yview)
        self.notes_scroll_bar.grid(row=0, column=1, sticky='ns')
        self.show_notes_input.configure(yscrollcommand=self.notes_scroll_bar.set)

        # Create Buttons
        self.edit_button = Button(self.view_edit_container, text="Edit", height=button_height, state="disabled",
                                  font=myFont, background="cornflower blue", command=self.toggle_edit_mode)
        self.edit_button.grid(row=21, column=0, columnspan=1, rowspan=1, padx=(10, 2), sticky="we", pady=10)

        self.update_button = Button(self.view_edit_container, text="Update", height=button_height, font=myFont,
                                    background="cornflower blue", state="disabled",
                                    command=lambda: self.update_unit(self.update_filter))
        self.update_button.grid(row=21, column=1, columnspan=1, rowspan=1, padx=2, sticky="we", pady=10)

        self.cancel_update = Button(self.view_edit_container, text="Cancel Update", height=button_height, font=myFont,
                                    background="cornflower blue", state="disabled", command=self.cancel_update_unit)
        self.cancel_update.grid(row=21, column=2, columnspan=1, rowspan=1, padx=2, sticky="we", pady=10)

        self.close_button = Button(self.view_edit_container, text="Close", height=button_height, font=myFont,
                                   background="cornflower blue", command=lambda: self.close_window())
        self.close_button.grid(row=21, column=3, columnspan=1, rowspan=1, padx=(2, 5), sticky="we", pady=10)

        # Create Unit Widgets
        self.select_unit_house_labelframe = LabelFrame(self.select_unit_frame, width=275, height=140,
                                                       background="white", font=myFont, highlightthickness=0,
                                                       borderwidth=0)
        self.select_unit_house_labelframe.grid_propagate(0)
        self.select_unit_house_labelframe.grid_rowconfigure((0, 1, 2, 3), weight=1)
        self.select_unit_house_labelframe.grid(row=0, column=0, columnspan=1, rowspan=5, padx=5, pady=5)

        self.select_unit_court_label = Label(self.select_unit_house_labelframe, text="Unit Court:", background="white",
                                             font=myFont)
        self.select_unit_court_label.grid(row=0, column=0, columnspan=1, rowspan=1, padx=5)

        self.select_unit_house_label = Label(self.select_unit_house_labelframe, text="Unit House:", background="white",
                                             font=myFont)
        self.select_unit_house_label.grid(row=1, column=0, columnspan=1, rowspan=1, padx=5, )

        self.select_unit_court_box = ttk.Combobox(self.select_unit_house_labelframe, state='readonly',
                                                  background="white", takefocus=False, font=myFont, width=15,
                                                  textvariable=self.unit_court_selected)
        self.select_unit_court_box["values"] = self.unit_court_list
        self.select_unit_court_box.grid(row=0, column=1, columnspan=1, rowspan=1, pady=1)

        self.select_unit_house_box = ttk.Combobox(self.select_unit_house_labelframe, state='readonly',
                                                  background="white", takefocus=False, font=myFont, width=15,
                                                  textvariable=self.unit_house_selected)
        self.select_unit_house_box["values"] = self.unit_house_list
        self.select_unit_house_box.grid(row=1, column=1, columnspan=1, rowspan=1)

        self.select_unit_label = Label(self.select_unit_house_labelframe, text="Unit:", font=myFont, background="white")
        self.select_unit_label.grid(row=2, column=0, columnspan=1, rowspan=1)

        self.select_unit_box = ttk.Combobox(self.select_unit_house_labelframe, state='readonly', takefocus=False,
                                            background="white", font=myFont, width=15, textvariable=self.unit_selected)
        self.select_unit_box.grid(row=2, column=1, columnspan=1, rowspan=1, pady=1)

        self.search_unit_button = Button(self.select_unit_house_labelframe, text="View Unit", height=button_height,
                                         background="cornflower blue", width=15, font=myFont,
                                         command=lambda: self.populate_unit_fields())
        self.search_unit_button.grid(row=3, column=1, columnspan=1, rowspan=1, pady=10)

        self.details_unit_address = Label(self.select_unit_frame, text="Unit Address:", anchor="w", width=15,
                                          background="white", font=myFont)
        self.details_unit_address.grid(row=0, column=1, columnspan=1, rowspan=1)

        self.unit_full = Text(self.select_unit_frame, height=7, width=20, state="disabled", background="white",
                              font=("Arial", -12), highlightthickness=0, borderwidth=0)
        self.unit_full.grid(row=1, column=1, columnspan=1, rowspan=3)

        # Create Tenant widgets
        self.tenant_name_label = Label(self.show_tenant_frame, text="Name:", anchor="e", font=myFont)
        self.tenant_name_label.grid(row=0, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_address_line_one_label = Label(self.show_tenant_frame, text="Address Line 1:", anchor="e",
                                                   font=myFont)
        self.tenant_address_line_one_label.grid(row=1, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_address_line_two_label = Label(self.show_tenant_frame, text="Address Line 2:", anchor="e",
                                                   font=myFont)
        self.tenant_address_line_two_label.grid(row=2, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_address_line_three_label = Label(self.show_tenant_frame, text="Address Line 3:", anchor="e",
                                                     font=myFont)
        self.tenant_address_line_three_label.grid(row=3, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_city_label = Label(self.show_tenant_frame, text="City:", anchor="e", font=myFont)
        self.tenant_city_label.grid(row=4, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_county_label = Label(self.show_tenant_frame, text="County:", anchor="e", font=myFont)
        self.tenant_county_label.grid(row=5, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_postcode_label = Label(self.show_tenant_frame, text="PostCode:", anchor="e", font=myFont)
        self.tenant_postcode_label.grid(row=6, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_country_label = Label(self.show_tenant_frame, text="Country:", anchor="e", font=myFont)
        self.tenant_country_label.grid(row=7, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_phone_label = Label(self.show_tenant_frame, text="Phone:", anchor="e", font=myFont)
        self.tenant_phone_label.grid(row=8, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_email_1_label = Label(self.show_tenant_frame, text="E-mail:", anchor="e", font=myFont)
        self.tenant_email_1_label.grid(row=9, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_email_2_label = Label(self.show_tenant_frame, text="E-mail 2:", anchor="e", font=myFont)
        self.tenant_email_2_label.grid(row=10, column=0, columnspan=1, rowspan=1, sticky="we")

        self.tenant_director_label = Label(self.show_tenant_frame, text="Director:", anchor="e", font=myFont)
        self.tenant_director_label.grid(row=11, column=0, columnspan=1, rowspan=1, sticky="we")

        # Create Certificate Widgets
        self.last_serviced_label = Label(self.show_certificate_frame, text="Last Serviced:", anchor="e", width=20,
                                         background="white", font=myFont)
        self.last_serviced_label.grid(row=0, column=0, columnspan=1, rowspan=1)

        self.service_due_label = Label(self.show_certificate_frame, text="Service Due:", anchor="e", width=20,
                                       background="white", font=myFont)
        self.service_due_label.grid(row=1, column=0, columnspan=1, rowspan=1)

        self.days_until_service_due_label = Label(self.show_certificate_frame, text="Days Until Service Due:",
                                                  anchor="e", width=20, background="white", font=myFont)
        self.days_until_service_due_label.grid(row=2, column=0, columnspan=1, rowspan=1)

        self.valid_cert_label = Label(self.show_certificate_frame, text="In Compliance:", anchor="e", width=20,
                                      background="white", font=myFont)
        self.valid_cert_label.grid(row=3, column=0, columnspan=1, rowspan=1)

        self.update_cert_view = Button(self.show_certificate_frame, text="Edit Service Date",
                                       background="cornflower blue", height=button_height, font=myFont,
                                       state="disabled", command=self.edit_service_date)
        self.update_cert_view.grid(row=0, column=2, columnspan=1, rowspan=1, sticky="ew", padx=5)

        # RTM Frame Widgets

        self.rtm_group = IntVar()
        self.rtm_member = StringVar()
        self.rtm_email = StringVar()

        self.rtm_group_label = Label(self.rtm_member_frame, text="RTM Group:", font=myFont, background="white",
                                     anchor="e")
        self.rtm_group_label.grid(row=0, column=0, sticky="NSEW")

        self.rtm_group_val = Label(self.rtm_member_frame, textvariable=self.rtm_group, font=myFont, background="white")
        self.rtm_group_val.grid(row=0, column=1, sticky="NSEW")

        self.rtm_member_status_label = Label(self.rtm_member_frame, text="RTM Member:", font=myFont, background="white",
                                             anchor="e")
        self.rtm_member_status_label.grid(row=0, column=2, sticky="NSEW")

        self.rtm_email_label = Label(self.rtm_member_frame, text="RTM E-mail:", font=myFont, background="white",
                                     anchor="e")
        self.rtm_email_label.grid(row=1, column=0, columnspan=1, sticky="EW", pady=(0, 5))

        self.rtm_member_radio_yes = Radiobutton(self.rtm_member_frame, text="Yes", value="Yes", font=myFont,
                                                background="white", variable=self.rtm_member, state="disabled")
        self.rtm_member_radio_yes.grid(row=0, column=3, columnspan=1, rowspan=1, pady=2)
        self.rtm_member_radio_no = Radiobutton(self.rtm_member_frame, text="No", value="No", font=myFont,
                                               background="white", variable=self.rtm_member, state="disabled")
        self.rtm_member_radio_no.grid(row=0, column=4, columnspan=1, rowspan=1, pady=2)

        self.rtm_email_entry = Entry(self.rtm_member_frame, font=myFont, state="disabled", textvariable=self.rtm_email)
        self.rtm_email_entry.grid(row=1, column=1, columnspan=4, sticky="EW", padx=(0, 10), pady=(0, 5))

        # Files frames
        self.files_treeview = ttk.Treeview(self.unit_files_frame, show="", columns=["#0"], selectmode="browse")

        # self.files_treeview.column("#0", width=0)
        # self.files_treeview.column("#0", minwidth=0)
        self.files_treeview.grid(row=0, column=0, rowspan=3, columnspan=6, sticky="nsew")

        self.files_scroll_bar = ttk.Scrollbar(self.unit_files_frame, orient="vertical",
                                              command=self.files_treeview.yview)
        self.files_scroll_bar.grid(row=0, column=6, sticky='ns', rowspan=3, columnspan=1)
        self.files_treeview.configure(yscrollcommand=self.files_scroll_bar.set)

        self.add_file_button = Button(self.unit_files_frame, font=myFont, background="cornflower blue", text="Add File",
                                      height=button_height, command=self.upload_certificate)
        self.add_file_button.grid(row=0, column=7, rowspan=1, columnspan=4, sticky="nsew", padx=4, pady=(4, 2))

        self.view_file_button = Button(self.unit_files_frame, font=myFont, background="cornflower blue",
                                       text="View File", height=button_height, command=self.view_certificate)
        self.view_file_button.grid(row=1, column=7, rowspan=1, columnspan=4, sticky="nsew", padx=4, pady=(2, 2))

        self.del_file_button = Button(self.unit_files_frame, font=myFont, background="cornflower blue",
                                      text="Delete File", height=button_height, command=self.delete_file)
        self.del_file_button.grid(row=2, column=7, rowspan=1, columnspan=4, sticky="nsew", padx=4, pady=(2, 4))

        self.del_file_button["state"] = "disabled"
        self.add_file_button["state"] = "disabled"
        self.view_file_button["state"] = "disabled"

        # Variables
        self.consent = StringVar()
        self.director = StringVar()
        self.tenant_name_val = StringVar()
        self.address_line_one_val = StringVar()
        self.address_line_two_val = StringVar()
        self.address_line_three_val = StringVar()
        self.city_val = StringVar()
        self.county_val = StringVar()
        self.postcode_val = StringVar()
        self.country_val = StringVar()
        self.phone_val = StringVar()
        self.email_val = StringVar()
        self.email_2_val = StringVar()
        self.last_serviced_val = StringVar()
        self.service_due_val = StringVar()
        self.days_until_service_due_val = StringVar()
        self.valid_cert_val = StringVar()

        # Servicing labels
        self.last_serviced_display = DateEntry(self.show_certificate_frame, textvariable=self.last_serviced_val,
                                               width=20, background="cornflower blue", foreground="black",
                                               date_pattern='dd/MM/yyyy', font=myFont, state="disabled")
        if os_type != "Windows":
            self.last_serviced_display.top_cal.overrideredirect(False)
        self.last_serviced_display.grid(row=0, column=1, columnspan=1, rowspan=1)

        self.service_due_display = Label(self.show_certificate_frame, textvariable=self.service_due_val, width=20,
                                         background="white", font=myFont, anchor="w")
        self.service_due_display.grid(row=1, column=1, columnspan=1, rowspan=1)

        self.days_until_service_due_display = Label(self.show_certificate_frame,
                                                    textvariable=self.days_until_service_due_val, background="white",
                                                    font=myFont, width=20, anchor="w")
        self.days_until_service_due_display.grid(row=2, column=1, columnspan=1, rowspan=1)

        self.valid_cert_display = Label(self.show_certificate_frame, textvariable=self.valid_cert_val, width=20,
                                        background="white", font=myFont, anchor="w")
        self.valid_cert_display.grid(row=3, column=1, columnspan=1, rowspan=1)

        # Tenant Info Inputs
        self.tenant_name_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                       textvariable=self.tenant_name_val)
        self.tenant_name_entry.grid(row=0, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_address_line_one_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                                   textvariable=self.address_line_one_val)
        self.tenant_address_line_one_entry.grid(row=1, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_address_line_two_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                                   textvariable=self.address_line_two_val)
        self.tenant_address_line_two_entry.grid(row=2, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_address_line_three_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                                     textvariable=self.address_line_three_val)
        self.tenant_address_line_three_entry.grid(row=3, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_city_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                       textvariable=self.city_val)
        self.tenant_city_entry.grid(row=4, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_county_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                         textvariable=self.county_val)
        self.tenant_county_entry.grid(row=5, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_postcode_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                           textvariable=self.postcode_val)
        self.tenant_postcode_entry.grid(row=6, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_country_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                          textvariable=self.country_val)
        self.tenant_country_entry.grid(row=7, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_phone_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                        textvariable=self.phone_val)
        self.tenant_phone_entry.grid(row=8, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_email_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                        textvariable=self.email_val)
        self.tenant_email_entry.grid(row=9, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_email_2_entry = Entry(self.show_tenant_frame, state="disabled", font=myFont,
                                          textvariable=self.email_2_val)
        self.tenant_email_2_entry.grid(row=10, column=1, columnspan=2, rowspan=1, sticky="we")

        self.tenant_director_radio_yes = Radiobutton(self.show_tenant_frame, text="Yes", value="Yes", font=myFont,
                                                     variable=self.director, state="disabled")
        self.tenant_director_radio_yes.grid(row=11, column=1, columnspan=1, rowspan=1, pady=2)
        self.tenant_director_radio_no = Radiobutton(self.show_tenant_frame, text="No", value="No", font=myFont,
                                                    variable=self.director, state="disabled")
        self.tenant_director_radio_no.grid(row=11, column=2, columnspan=1, rowspan=1, pady=2)

        self.update_filter = {}

        if self.came_from == "Search":
            self.unit_court_selected.set(SearchTenant.unit_court_selected)
            self.unit_selected.set(SearchTenant.unit_num_selected)
            self.populate_unit_fields()

        elif self.came_from == "Main":
            self.unit_court_selected.set("")
            self.unit_selected.set("")

        elif self.came_from == "Report":
            self.unit_court_selected.set(ReportWindow.unit_court_selected)
            self.unit_selected.set(ReportWindow.unit_num_selected)
            self.populate_unit_fields()

        # Set background / disabledbackground for children of frame
        for child in self.show_tenant_frame.children.values():
            child["disabledforeground"] = "black"
            child["background"] = "white"
        for child in self.rtm_member_frame.children.values():
            child["disabledforeground"] = "black"
            child["background"] = "white"

        self.cert_temp_direct = tempfile.mkdtemp(prefix="BoilerApp_", suffix="_Certs")
        # Hide root window while Toplevel open
        root.withdraw()
        # Handle "x" on Toplevel
        self.view_edit_container.protocol("WM_DELETE_WINDOW", self.close_window)
        self.files_treeview.bind('<Double-Button-1>', self.view_certificate)
        self.files_treeview.bind('<<TreeviewSelect>>', self.files_on_select)

    @Database.autoreconnect_retry
    def delete_file(self):
        self.file_selected = self.files_treeview.focus()
        self.file_id_to_delete = self.files_treeview.item(self.file_selected)["values"][1]

        if messagebox.askyesno(title="File Delete", message="Are you sure you want to delete: " +
                                                            self.files_treeview.item(self.file_selected)["values"][
                                                                0] + " ?"):
            try:
                self.delete_old_file = Database.certificate_storage.delete(ObjectId(self.file_id_to_delete))

                for i in range(1, 11):
                    if self.unit_document["File_" + str(i)] == ObjectId(self.file_id_to_delete):
                        self.cert_url_update = Database.crud_app_collection.update_one(
                            self.update_filter,
                            {"$set":
                                 {"File_" + str(i): ""}
                             })
                        break
                    else:
                        continue

            except pymongo.errors.PyMongoError as e:
                return messagebox.showerror("Exception", e, parent=self.view_edit_container)

            else:
                Database.log_action("File Deleted: " + str(self.files_treeview.item(self.file_selected)["values"][0]),
                                    str(self.update_filter))
                messagebox.showinfo("File deleted successfully",
                                    "Successfully deleted file for: \n\n" + self.unit_full.get("1.0", "end-1c"),
                                    parent=self.view_edit_container)
                self.populate_unit_fields()
                self.del_file_button["state"] = "disabled"

    def files_on_select(self, *args):
        self.del_file_button["state"] = "normal"
        self.view_file_button["state"] = "normal"

    @Database.autoreconnect_retry
    def view_certificate(self, *args):

        cur_item = self.files_treeview.focus()

        if cur_item != "":
            self.file_selected = self.files_treeview.item(cur_item)["values"][1]

            try:
                self.cert_download = Database.certificate_storage.get(ObjectId(self.file_selected)).read()

                with tempfile.NamedTemporaryFile(
                        prefix=self.files_treeview.item(cur_item)["values"][0].replace(" ", "_") + "_", suffix=".pdf",
                        delete=False, dir=self.cert_temp_direct) as self.tempfilename:
                    self.saved_file = open(self.tempfilename.name, "wb")
                    self.saved_file.write(self.cert_download)
                    self.saved_file.close()
                if os_type == "Darwin":
                    subprocess.run(['open', self.tempfilename.name], check=True)
                elif os_type == "Windows":
                    os.system("start " + self.tempfilename.name)

            except pymongo.errors.PyMongoError as e:
                return messagebox.showerror("Exception", e, parent=self.view_edit_container)

    @Database.autoreconnect_retry
    def upload_certificate(self):

        self.upload_file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        self.file_field = ""
        if self.upload_file_path != "":
            for i in range(1, 11):
                if self.unit_document["File_" + str(i)] == "":
                    self.file_field = "File_" + str(i)
                    break
                else:
                    continue

            if self.file_field == "":
                messagebox.showerror("File limited reached!", "Please delete a file. Maximum files per Unit is 10.",
                                     parent=self.view_edit_container)
            else:
                try:
                    with open(self.upload_file_path, "rb") as input_file:
                        upload_object = Database.certificate_storage.put(input_file, content_type="application/pdf",
                                                                         filename=os.path.basename(
                                                                             self.upload_file_path))

                    self.cert_url_update = Database.crud_app_collection.update_one(
                        self.update_filter,
                        {"$set": {self.file_field: upload_object}})
                except pymongo.errors.PyMongoError as e:
                    return messagebox.showerror("Exception", e, parent=self.view_edit_container)

                else:
                    Database.log_action("File Uploaded: " + str(os.path.basename(self.upload_file_path)),
                                        str(self.update_filter))
                    messagebox.showinfo("Upload Successful",
                                        "Successfully uploaded new file for: \n\n" + self.unit_full.get("1.0",
                                                                                                        "end-1c"),
                                        parent=self.view_edit_container)

        self.populate_unit_fields()

    @Database.autoreconnect_retry
    def edit_service_date(self):

        if self.cert_edit_mode is False:
            self.update_cert_view["text"] = "Update Service Date"

            self.cert_edit_mode = True
            self.last_serviced_display["state"] = "normal"
            self.select_unit_box["state"] = "disabled"
            self.select_unit_house_box["state"] = "disabled"
            self.select_unit_court_box["state"] = "disabled"
            self.search_unit_button["state"] = "disabled"
            self.edit_button["state"] = "disabled"
            self.close_button["state"] = "disabled"

        else:
            if self.last_serviced_val.get() != "":
                try:
                    add_year = datetime.strptime(self.last_serviced_val.get(), "%d/%m/%Y")
                    add_year = add_year.replace(year=add_year.year + 1)
                    add_year = add_year.strftime("%d/%m/%Y")
                    self.service_due_val.set(add_year)
                except ValueError:
                    return messagebox.showerror("Update Failed", "Date, " + str(
                        self.last_serviced_val.get()) + " does not match format DD/MM/YYYY",
                                                parent=self.view_edit_container)

            else:
                self.service_due_val.set("")

            # Write Update to Database
            try:
                if self.last_serviced_val.get() != "":
                    Database.crud_app_collection.update_one(
                        self.update_filter,
                        {"$set":
                             {"Last_Serviced_ISO": datetime.strptime(self.last_serviced_val.get(), "%d/%m/%Y"),
                              "Service_Due_ISO": datetime.strptime(self.service_due_val.get(), "%d/%m/%Y"),
                              "Service_Due_Month": datetime.strptime(self.service_due_val.get(), "%d/%m/%Y").month,
                              "Service_Due_Year": datetime.strptime(self.service_due_val.get(), "%d/%m/%Y").year
                              }})

                else:
                    Database.crud_app_collection.update_one(
                        self.update_filter,
                        {"$set":
                             {"Last_Serviced_ISO": datetime.strptime("01/01/1970", "%d/%m/%Y"),
                              "Service_Due_ISO": datetime.strptime("01/01/1970", "%d/%m/%Y"),
                              "Service_Due_Month": datetime.strptime("01/01/1970", "%d/%m/%Y").month,
                              "Service_Due_Year": datetime.strptime("01/01/1970", "%d/%m/%Y").year
                              }})

            except pymongo.errors.PyMongoError as e:
                return messagebox.showerror("Exception", e, parent=self.view_edit_container)

            else:

                self.calculate_dates()
                self.update_cert_view["text"] = "Edit Service Date"

                self.cert_edit_mode = False
                self.last_serviced_display["state"] = "disabled"
                self.select_unit_box["state"] = "readonly"
                self.select_unit_house_box["state"] = "readonly"
                self.select_unit_court_box["state"] = "readonly"
                self.search_unit_button["state"] = "normal"
                self.edit_button["state"] = "normal"
                self.close_button["state"] = "normal"

                if self.last_serviced_val.get() != "":
                    Database.log_action("Service Date Updated: " + str(self.last_serviced_val.get()),
                                        str(self.update_filter))
                else:
                    Database.log_action("Service Date Updated: ", str(self.update_filter))
                messagebox.showinfo("Update Successful",
                                    "Successfully Updated Service Date for: \n\n" + self.unit_full.get("1.0", "end-1c"),
                                    parent=self.view_edit_container)

    @Database.autoreconnect_retry
    def update_unit_list(self, unit_filter):
        # Populate Unit List once unit house selected
        try:
            if unit_filter == "house":
                self.select_unit_box["values"] = Database.crud_app_collection.distinct("Unit", {
                    "Unit_House": self.unit_house_selected.get()})
                self.unit_selected.set("")
                self.unit_court_selected.trace_remove("write", self.unit_court_trace)
                self.unit_court_selected.set("")
                self.unit_court_trace = self.unit_court_selected.trace_add("write", lambda *args,
                                                                                           unit_filter="court": self.update_unit_list(
                    unit_filter))
            elif unit_filter == "court":
                self.select_unit_box["values"] = Database.crud_app_collection.distinct("Unit", {
                    "Unit_Court": self.unit_court_selected.get()})
                self.unit_selected.set("")
                self.unit_house_selected.trace_remove("write", self.unit_house_trace)
                self.unit_house_selected.set("")
                self.unit_house_trace = self.unit_house_selected.trace_add("write", lambda *args,
                                                                                           unit_filter="house": self.update_unit_list(
                    unit_filter))

        except pymongo.errors.PyMongoError as e:
            return messagebox.showerror("Exception", e, parent=self.view_edit_container)

    @Database.autoreconnect_retry
    def populate_unit_fields(self):

        # Filtered Documents
        if self.unit_selected.get() == "":
            raise ValueError(
                messagebox.showerror("Unit Not Valid", "Please Select Unit Number", parent=self.view_edit_container))
        else:
            self.update_filter = {}
            if self.unit_court_selected.get() == "":
                self.update_filter["Unit_House"] = self.unit_house_selected.get()
                self.update_filter["Unit"] = int(self.unit_selected.get())
            elif self.unit_house_selected.get() == "":
                self.update_filter["Unit_Court"] = self.unit_court_selected.get()
                self.update_filter["Unit"] = int(self.unit_selected.get())
            else:
                self.update_filter["Unit_Court"] = self.unit_court_selected.get()
                self.update_filter["Unit_House"] = self.unit_house_selected.get()
                self.update_filter["Unit"] = int(self.unit_selected.get())

        try:

            self.unit_document = Database.crud_app_collection.find_one(self.update_filter)

        except pymongo.errors.PyMongoError as e:
            return messagebox.showerror("Exception", e, parent=self.view_edit_container)

        else:
            # Populate entry fields from Database Document
            self.edit_button['state'] = 'normal'
            self.tenant_name_val.set(self.unit_document["Name"])
            self.address_line_one_val.set(self.unit_document["Address_Line_1"])
            self.address_line_two_val.set(self.unit_document["Address_Line_2"])
            self.address_line_three_val.set(self.unit_document["Address_Line_3"])
            self.city_val.set(self.unit_document["City"])
            self.county_val.set(self.unit_document["County"])
            self.postcode_val.set(self.unit_document["Postcode"])
            self.country_val.set(self.unit_document["Country"])
            self.phone_val.set(self.unit_document["Phone"])
            self.email_val.set(self.unit_document["Email_1"])
            self.email_2_val.set(self.unit_document["Email_2"])
            self.director.set(self.unit_document["Director"])
            self.rtm_group.set(self.unit_document["RTM_Group"])
            self.rtm_member.set(self.unit_document["RTM_Member"])
            self.rtm_email.set(self.unit_document["RTM_Email"])

            if datetime.strftime(self.unit_document["Last_Serviced_ISO"], "%d/%m/%Y") == "01/01/1970":
                self.last_serviced_val.set("")
                self.service_due_val.set("")

            else:
                self.last_serviced_val.set(datetime.strftime(self.unit_document["Last_Serviced_ISO"], "%d/%m/%Y"))
                self.service_due_val.set(datetime.strftime(self.unit_document["Service_Due_ISO"], "%d/%m/%Y"))

            self.calculate_dates()

            # Populate Full Unit Address
            self.unit_full.config(state=NORMAL)
            self.unit_full.delete(1.0, END)
            self.unit_full.insert(1.0,
                                  self.unit_document["Unit_Full"] + '\n' + self.unit_document["Unit_House"] + '\n' +
                                  self.unit_document["Unit_Road"] + '\n' + self.unit_document["Unit_City"] + '\n' +
                                  self.unit_document["Unit_Postcode"])
            self.unit_full.config(state=DISABLED)

            # Populate Notes Input
            self.show_notes_input.config(state=NORMAL)
            self.show_notes_input.delete(1.0, END)
            self.show_notes_input.insert(1.0, self.unit_document["Notes"])
            self.show_notes_input.config(state=DISABLED)

            self.update_cert_view["state"] = "normal"

            self.files_treeview.delete(*self.files_treeview.get_children())

            self.file_list = []
            for i in range(1, 11):
                if self.unit_document["File_" + str(i)] != "":
                    self.file_query_result = Database.certificate_storage.find_one(
                        {"_id": self.unit_document["File_" + str(i)]})
                    self.file_list.append({
                        "filename": self.file_query_result.name,
                        "_id": self.unit_document["File_" + str(i)]
                    })

                    self.files_treeview.insert("", i - 1, values=[self.file_query_result.name,
                                                                  self.unit_document["File_" + str(i)]])

            self.add_file_button["state"] = "normal"
            self.del_file_button["state"] = "disabled"
            self.view_file_button["state"] = "disabled"

    def calculate_dates(self):

        if self.last_serviced_val.get() != "":
            if self.service_due_val.get() != "":
                self.days_until = datetime.strptime(self.service_due_val.get(),
                                                    "%d/%m/%Y").date() - datetime.now().date()
                self.days_until_service_due_val.set(self.days_until.days)
                if self.days_until.days >= 0:
                    self.valid_cert_val.set("Yes")
                else:
                    self.valid_cert_val.set("No")
            else:
                self.valid_cert_val.set("No")
                self.days_until_service_due_val.set("-")
        else:
            self.service_due_val.set("")
            self.valid_cert_val.set("No")
            self.days_until_service_due_val.set("-")

    def close_window(self):
        # Handle closing Toplevel event.

        self.direct_delete = glob.glob(os.path.join(tempfile.gettempdir(), "CenturyWharf_*"))
        for path in self.direct_delete:
            try:
                shutil.rmtree(path)
            except PermissionError:
                continue

        if self.came_from == "Search":
            self.view_edit_container.destroy()
            self.source_window.search_container.deiconify()
            self.source_window.populate_treeview()
        elif self.came_from == "Main":
            self.view_edit_container.destroy()
            root.deiconify()
        elif self.came_from == "Report":
            self.view_edit_container.destroy()
            self.source_window.report_container.deiconify()
            self.source_window.populate_results_table()

    def toggle_edit_mode(self):
        # Toggle between edit and view mode
        if not self.edit_mode:
            # Enable entry boxes when edit mode entered
            for child in self.show_tenant_frame.children.values():
                child["state"] = "normal"
            for child in self.rtm_member_frame.children.values():
                child["state"] = "normal"
            # Toggle edit mode to true
            self.edit_mode = True
            # Disable/enable widgets
            self.edit_button["state"] = "disabled"
            self.close_button["state"] = "disabled"
            self.update_button["state"] = "normal"
            self.cancel_update["state"] = "normal"
            self.select_unit_box["state"] = "disabled"
            self.select_unit_house_box["state"] = "disabled"
            self.select_unit_court_box["state"] = "disabled"
            self.update_cert_view["state"] = "disabled"

            self.show_notes_input["state"] = "normal"
            self.search_unit_button["state"] = "disabled"
        else:

            # Disable entry boxes when update / cancelupdate triggered
            for child in self.show_tenant_frame.children.values():
                child["state"] = "disabled"
            for child in self.rtm_member_frame.children.values():
                child["state"] = "disabled"
            # Toggle edit mode to False
            self.edit_mode = False
            # Disable/enable widgets
            self.close_button["state"] = "normal"
            self.edit_button["state"] = "normal"
            self.update_button["state"] = "disabled"
            self.cancel_update["state"] = "disabled"
            self.select_unit_box["state"] = "readonly"
            self.select_unit_house_box["state"] = "readonly"
            self.select_unit_court_box["state"] = "readonly"
            self.update_cert_view["state"] = "normal"

            self.show_notes_input["state"] = "disabled"
            self.search_unit_button["state"] = "normal"

    @Database.autoreconnect_retry
    def update_unit(self, update_filter):

        # Write Update to Database
        self.update_filter = update_filter

        try:

            Database.crud_app_collection.update_one(
                self.update_filter,
                {"$set":
                     {"Name": self.tenant_name_val.get(),
                      "Address_Line_1": self.address_line_one_val.get(),
                      "Address_Line_2": self.address_line_two_val.get(),
                      "Address_Line_3": self.address_line_three_val.get(),
                      "City": self.city_val.get(),
                      "County": self.county_val.get(),
                      "Postcode": self.postcode_val.get(),
                      "Country": self.country_val.get(),
                      "Phone": self.phone_val.get(),
                      "Email_1": self.email_val.get(),
                      "Email_2": self.email_2_val.get(),
                      "Director": self.director.get(),
                      "Notes": self.show_notes_input.get("1.0", "end-1c"),
                      "RTM_Email": self.rtm_email.get(),
                      "RTM_Member": self.rtm_member.get()

                      }})

        except pymongo.errors.PyMongoError as e:
            return messagebox.showerror("Exception", e, parent=self.view_edit_container)

        else:
            # Once updated, deactivate edit mode
            self.toggle_edit_mode()

            Database.log_action("Owner Details Updated", str(self.update_filter))
            messagebox.showinfo("Update Successful",
                                "Successfully Updated Tenant Details for: \n\n" + self.unit_full.get("1.0", "end-1c"),
                                parent=self.view_edit_container)

    def cancel_update_unit(self):
        # deactivate edit mode
        self.toggle_edit_mode()
        # Repopulate entry fields from database data values
        self.populate_unit_fields()


class ViewEditUnit:
    def __init__(self, master):

        # Populate Combobox from distinct "Unit_Court" Values
        self.unit_court_list = Database.crud_app_collection.distinct("Unit_Court")
        self.unit_house_list = Database.crud_app_collection.distinct("Unit_House")

        # Set Variables and handle Trace method
        self.unit_court_selected = StringVar()
        self.unit_court_selected.set("")
        self.unit_house_selected = StringVar()
        self.unit_house_selected.set("")
        self.unit_selected = StringVar()
        self.unit_selected.set("")
        self.unit_court_selected.trace("w",
                                       lambda *args, passed=self.unit_court_selected: self.update_unit_house(passed))
        self.unit_house_selected.trace("w",
                                       lambda *args, passed=self.unit_house_selected: self.update_unit_list(passed))

        # Build Tkinter Toplevel Widget
        self.view_edit_unit_container = Toplevel(master)
        self.view_edit_unit_container.title("View / Edit Unit")
        self.view_edit_unit_container.geometry("500x450+50+50")
        self.view_edit_unit_container.resizable(False, False)

        # Set edit mode to False on widget initalise.
        self.edit_mode = False

        # Build Tkinter LabelFrames
        self.view_edit_select_unit = LabelFrame(self.view_edit_unit_container, text="Select Unit to View / Edit",
                                                width=480, height=170)
        self.view_edit_select_unit.grid_propagate(0)
        self.view_edit_select_unit.grid(row=0, column=0, columnspan=4, rowspan=4, padx=10, pady=10)
        self.view_edit_unit_details = LabelFrame(self.view_edit_unit_container, text="Unit Details", width=480,
                                                 height=190)
        self.view_edit_unit_details.grid_propagate(0)
        self.view_edit_unit_details.grid(row=5, column=0, columnspan=4, rowspan=8, padx=10, pady=10)

        # Tkinter Select Labels
        self.edit_select_unit_court_label = Label(self.view_edit_select_unit, text="Select Unit Court:", anchor="e",
                                                  width=15)
        self.edit_select_unit_court_label.grid(row=0, column=0, columnspan=1, rowspan=1, pady=5, padx=10)
        self.edit_select_unit_house_label = Label(self.view_edit_select_unit, text="Select Unit House:", anchor="e",
                                                  width=15)
        self.edit_select_unit_house_label.grid(row=1, column=0, columnspan=1, rowspan=1, padx=10)
        self.edit_select_unit_label = Label(self.view_edit_select_unit, text="Select Unit:", anchor="e", width=15)
        self.edit_select_unit_label.grid(row=2, column=0, columnspan=1, rowspan=1, padx=10, pady=5)

        # Tkinter select Combobox
        self.edit_select_unit_court_box = ttk.Combobox(self.view_edit_select_unit, state="readonly", takefocus=False,
                                                       width=30, textvariable=self.unit_court_selected)
        self.edit_select_unit_court_box["values"] = self.unit_court_list
        self.edit_select_unit_court_box.grid(row=0, column=1, columnspan=2, rowspan=1, pady=5)
        self.edit_select_unit_house_box = ttk.Combobox(self.view_edit_select_unit, state="readonly", takefocus=False,
                                                       width=30, textvariable=self.unit_house_selected)
        self.edit_select_unit_house_box.grid(row=1, column=1, columnspan=2, rowspan=1)
        self.edit_select_unit_num_box = ttk.Combobox(self.view_edit_select_unit, state="readonly", takefocus=False,
                                                     width=30, textvariable=self.unit_selected)
        self.edit_select_unit_num_box.grid(row=2, column=1, columnspan=2, rowspan=1)
        self.search_unit_button = Button(self.view_edit_select_unit, text="View Unit", height=2, width=30,
                                         command=lambda: self.populate_unit_fields(self.unit_selected,
                                                                                   self.unit_court_selected,
                                                                                   self.unit_house_selected))
        self.search_unit_button.grid(row=3, column=1, columnspan=1, rowspan=1)

        # Tkinter edit, update, cancel update, close buttons
        self.unit_edit_button = Button(self.view_edit_unit_container, text="Edit", width=11, height=2,
                                       command=lambda: self.toggle_edit_mode())
        self.unit_edit_button.grid(row=13, column=0, columnspan=1, rowspan=1)
        self.unit_update_button = Button(self.view_edit_unit_container, text="Update", width=11, height=2,
                                         state="disabled",
                                         command=lambda: self.update_unit(self.unit_selected, self.unit_court_selected,
                                                                          self.unit_house_selected))
        self.unit_update_button.grid(row=13, column=1, columnspan=1, rowspan=1)
        self.unit_cancel_update_button = Button(self.view_edit_unit_container, text="Cancel Update", width=11, height=2,
                                                state="disabled", command=lambda: self.cancel_update_unit())
        self.unit_cancel_update_button.grid(row=13, column=2, columnspan=1, rowspan=1)
        self.unit_container_close_button = Button(self.view_edit_unit_container, text="Close", width=11, height=2,
                                                  command=lambda: self.close_window())
        self.unit_container_close_button.grid(row=13, column=3, columnspan=1, rowspan=1)

        # Tkinter Unit Labels
        self.unit_label = Label(self.view_edit_unit_details, text="Unit Number:", width=15, anchor="e")
        self.unit_label.grid(row=0, column=0, columnspan=1, rowspan=1)
        self.unit_court_label = Label(self.view_edit_unit_details, text="Unit Court:", width=15, anchor="e")
        self.unit_court_label.grid(row=1, column=0, columnspan=1, rowspan=1)
        self.unit_house_label = Label(self.view_edit_unit_details, text="Unit House:", width=15, anchor="e")
        self.unit_house_label.grid(row=2, column=0, columnspan=1, rowspan=1)
        self.unit_road_label = Label(self.view_edit_unit_details, text="Unit Road:", width=15, anchor="e")
        self.unit_road_label.grid(row=3, column=0, columnspan=1, rowspan=1)
        self.unit_city_label = Label(self.view_edit_unit_details, text="Unit City:", width=15, anchor="e")
        self.unit_city_label.grid(row=4, column=0, columnspan=1, rowspan=1)
        self.unit_postcode_label = Label(self.view_edit_unit_details, text="Unit Postcode:", width=15, anchor="e")
        self.unit_postcode_label.grid(row=5, column=0, columnspan=1, rowspan=1)

        # Declare StringVar
        self.unit_num_val = StringVar()
        self.unit_court_val = StringVar()
        self.unit_house_val = StringVar()
        self.unit_road_val = StringVar()
        self.unit_city_val = StringVar()
        self.unit_postcode_val = StringVar()

        # Build Tkinter Entry widgets
        self.unit_entry = Entry(self.view_edit_unit_details, width=30, state="disabled", textvariable=self.unit_num_val)
        self.unit_entry.grid(row=0, column=1, columnspan=2, rowspan=1)
        self.unit_court_entry = Entry(self.view_edit_unit_details, width=30, state="disabled",
                                      textvariable=self.unit_court_val)
        self.unit_court_entry.grid(row=1, column=1, columnspan=2, rowspan=1)
        self.unit_house_entry = Entry(self.view_edit_unit_details, width=30, state="disabled",
                                      textvariable=self.unit_house_val)
        self.unit_house_entry.grid(row=2, column=1, columnspan=2, rowspan=1)
        self.unit_road_entry = Entry(self.view_edit_unit_details, width=30, state="disabled",
                                     textvariable=self.unit_road_val)
        self.unit_road_entry.grid(row=3, column=1, columnspan=2, rowspan=1)
        self.unit_city_entry = Entry(self.view_edit_unit_details, width=30, state="disabled",
                                     textvariable=self.unit_city_val)
        self.unit_city_entry.grid(row=4, column=1, columnspan=2, rowspan=1)
        self.unit_postcode_entry = Entry(self.view_edit_unit_details, width=30, state="disabled",
                                         textvariable=self.unit_postcode_val)
        self.unit_postcode_entry.grid(row=5, column=1, columnspan=2, rowspan=1)

        # Hide root window while Toplevel open
        root.withdraw()
        # Handle "x" on Toplevel
        self.view_edit_unit_container.protocol("WM_DELETE_WINDOW", self.close_window)

    def close_window(self):
        # Handle closing Toplevel event.
        self.view_edit_unit_container.destroy()
        # Make root window visible on close
        root.deiconify()

    def toggle_edit_mode(self):
        if self.edit_mode is False:
            # Enable entry boxes when edit mode entered
            for child in self.view_edit_unit_details.children.values():
                child["state"] = "normal"
            self.unit_update_button["state"] = "normal"
            self.unit_cancel_update_button["state"] = "normal"
            self.unit_edit_button["state"] = "disabled"
            self.unit_container_close_button["state"] = "disabled"
            self.edit_select_unit_court_box["state"] = "disabled"
            self.edit_select_unit_house_box["state"] = "disabled"
            self.edit_select_unit_num_box["state"] = "disabled"
            self.edit_mode = True
        else:
            # disable entry boxes when edit mode entered
            for child in self.view_edit_unit_details.children.values():
                child["state"] = "disabled"
            self.unit_update_button["state"] = "disabled"
            self.unit_cancel_update_button["state"] = "disabled"
            self.unit_edit_button["state"] = "normal"
            self.unit_container_close_button["state"] = "normal"
            self.edit_select_unit_court_box["state"] = "readonly"
            self.edit_select_unit_house_box["state"] = "readonly"
            self.edit_select_unit_num_box["state"] = "readonly"
            self.edit_mode = False

    def cancel_update_unit(self):
        # Toggle edit mode to False
        self.toggle_edit_mode()
        # Populate Entry fields from values in database
        self.populate_unit_fields(self.unit_selected, self.unit_court_selected, self.unit_house_selected)

    def update_unit_list(self, unit_house_selected):
        # Populate Unit List once unit house selected
        self.edit_select_unit_num_box["values"] = Database.crud_app_collection.distinct("Unit", {
            "Unit_House": unit_house_selected.get()})
        self.unit_selected.set("")

    def update_unit_house(self, unit_court_selected):
        # Populate Unit House once Unit Court Selected
        self.edit_select_unit_house_box["values"] = Database.crud_app_collection.distinct("Unit_House", {
            "Unit_Court": unit_court_selected.get()})
        self.unit_house_selected.set("")
        self.unit_selected.set("")

    def populate_unit_fields(self, unit_selected, unit_court_selected, unit_house_selected):

        # Get filtered data from database
        unit_document = Database.crud_app_collection.find_one(
            {"Unit_House": unit_house_selected.get(), "Unit_Court": unit_court_selected.get(),
             "Unit": int(unit_selected.get())})

        # Set input variables to values from database
        self.unit_num_val.set(unit_document["Unit"])
        self.unit_court_val.set(unit_document["Unit_Court"])
        self.unit_house_val.set(unit_document["Unit_House"])
        self.unit_road_val.set(unit_document["Unit_Road"])
        self.unit_city_val.set(unit_document["Unit_City"])
        self.unit_postcode_val.set(unit_document["Unit_Postcode"])

    def update_unit(self, unit_selected, unit_court_selected, unit_house_selected):

        # Toggle edit mode on Update
        self.toggle_edit_mode()
        # Write Updated values to database

        Database.crud_app_collection.update_one(
            {"Unit_House": unit_house_selected.get(), "Unit_Court": unit_court_selected.get(),
             "Unit": int(unit_selected.get())},
            {"$set":
                 {"Unit": int(self.unit_num_val.get()),
                  "Unit_Court": self.unit_court_val.get(),
                  "Unit_House": self.unit_house_val.get(),
                  "Unit_Road": self.unit_road_val.get(),
                  "Unit_City": self.unit_city_val.get(),
                  "Unit_Postcode": self.unit_postcode_val.get()
                  }})


# noinspection PyUnresolvedReferences
class UserSettings:
    def __init__(self, master):
        self.user_settings_container = Toplevel(master)
        self.user_settings_container.title("User Account Settings")
        self.user_settings_container.geometry("390x210")
        self.user_settings_container.resizable(False, False)
        self.user_settings_container.config(background="white")
        self.user_settings_container.grid_columnconfigure((0, 1), weight=1, uniform="c")
        self.user_settings_container.grid_rowconfigure((0, 1), weight=1, uniform="r")

        # Gets the requested values of the height and width.
        self.windowWidth = 390
        self.windowHeight = 210
        # Gets both half the screen width/height and window width/height
        self.positionRight = int(root.winfo_screenwidth() / 2 - self.windowWidth / 2)
        self.positionDown = int(root.winfo_screenheight() / 2 - self.windowHeight / 2)
        # Positions the window in the center of the page.
        self.user_settings_container.geometry("+{}+{}".format(self.positionRight, self.positionDown))

        self.detail_frame = LabelFrame(self.user_settings_container, text="Change Password", highlightthickness=1,
                                       highlightbackground="black", borderwidth=1, background="white")
        self.detail_frame.grid_rowconfigure((0, 1, 2, 3), weight=1, uniform="r")
        self.detail_frame.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="c")
        self.detail_frame.grid(row=0, column=0, columnspan=2, rowspan=2, sticky="NSEW", padx=10, pady=5)

        self.settings_close_window = Button(self.user_settings_container, text="Close", height=button_height,
                                            font=myFont, command=lambda: self.close_window())
        self.settings_close_window.grid(row=2, column=0, columnspan=1, rowspan=1, sticky="EW", padx=(10, 5),
                                        pady=(5, 10))

        self.settings_update_password = Button(self.user_settings_container, text="Update", height=button_height,
                                               font=myFont, command=lambda: self.update_password())
        self.settings_update_password.grid(row=2, column=1, columnspan=1, rowspan=1, sticky="ew", padx=(5, 10),
                                           pady=(5, 10))

        self.user_settings_label = Label(self.detail_frame, text="Username:", font=myFontBold, anchor="e")
        self.user_settings_label.grid(row=0, column=0, columnspan=2, sticky="ew", pady=5, padx=5)

        self.user_name_label = Label(self.detail_frame, text=Database.username_val, font=myFontBold, anchor="w")
        self.user_name_label.grid(row=0, column=2, columnspan=2, sticky="ew", pady=5, padx=(0, 10))

        self.current_password_label = Label(self.detail_frame, text="Enter Current Password:", font=myFont, anchor="e")
        self.current_password_label.grid(row=1, column=0, columnspan=2, sticky="ew", pady=5, padx=5)

        self.current_password_val = StringVar()
        self.current_password_entry = Entry(self.detail_frame, font=myFont, show="*",
                                            textvariable=self.current_password_val)
        self.current_password_entry.grid(row=1, column=2, columnspan=2, sticky="ew", padx=(0, 10))

        self.new_password_label = Label(self.detail_frame, text="Enter New Password:", font=myFont, anchor="e")
        self.new_password_label.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5, padx=5)

        self.new_password_val = StringVar()
        self.new_password_entry = Entry(self.detail_frame, font=myFont, show="*", textvariable=self.new_password_val)
        self.new_password_entry.grid(row=2, column=2, columnspan=2, sticky="ew", padx=(0, 10))

        self.confirm_password_label = Label(self.detail_frame, text="Confirm New Password:", font=myFont, anchor="e")
        self.confirm_password_label.grid(row=3, column=0, columnspan=2, sticky="ew", pady=5, padx=5)

        self.confirm_password_val = StringVar()
        self.confirm_password_entry = Entry(self.detail_frame, font=myFont, show="*",
                                            textvariable=self.confirm_password_val)
        self.confirm_password_entry.grid(row=3, column=2, columnspan=2, sticky="ew", padx=(0, 10))

        # Hide root window while Toplevel open
        root.withdraw()
        # Handle "x" on Toplevel
        self.user_settings_container.protocol("WM_DELETE_WINDOW", self.close_window)

    def update_password(self):

        if self.current_password_val.get() != Database.password_val:
            return messagebox.showerror("Error", "Current password is not correct.",
                                        parent=self.user_settings_container)
        elif self.new_password_val.get() != self.confirm_password_val.get():
            return messagebox.showerror("Error", "New passwords do not match.", parent=self.user_settings_container)
        elif self.new_password_val.get() == "":
            return messagebox.showerror("Error", "Please enter new password.", parent=self.user_settings_container)
        else:
            try:
                self.password_update = Database.century_wharf_db.command("updateUser", Database.username_val,
                                                                         pwd=self.new_password_val.get())
            except pymongo.errors.PyMongoError as e:
                return messagebox.showerror("Exception", e, parent=self.user_settings_container)
            else:
                Database.password_val = self.new_password_val.get()
                messagebox.showinfo("Update Successful",
                                    "Password for User: " + Database.username_val + ", has successfully been changed.",
                                    parent=self.user_settings_container)
                self.close_window()

    def close_window(self):
        # Handle closing Toplevel event.
        self.user_settings_container.destroy()
        # Make root window visible on close
        root.deiconify()


class AddNewUnit:
    def __init__(self, master):
        # Create Tkinter Toplevel Widget
        self.new_unit_container = Toplevel(master)
        self.new_unit_container.title("Add New Unit")
        self.new_unit_container.geometry("500x260+50+50")
        self.new_unit_container.resizable(False, False)

        # Create Tkinter Labelframe
        self.new_unit_details = LabelFrame(self.new_unit_container, text="Add New Unit", width=480, height=190)
        self.new_unit_details.grid_propagate(0)
        self.new_unit_details.grid(row=5, column=0, columnspan=4, rowspan=8, padx=10, pady=10)

        # Create Tkinter Buttons
        self.unit_save_button = Button(self.new_unit_container, text="Save", width=15, height=2)
        self.unit_save_button.grid(row=13, column=0, columnspan=1, rowspan=1, padx=12)
        self.unit_clear_button = Button(self.new_unit_container, text="Clear", width=15, height=2)
        self.unit_clear_button.grid(row=13, column=1, columnspan=1, rowspan=1, padx=12)
        self.unit_container_close_button = Button(self.new_unit_container, text="Close", width=15, height=2,
                                                  command=lambda: self.close_window())
        self.unit_container_close_button.grid(row=13, column=2, columnspan=1, rowspan=1, padx=12)

        # Create Tkinter Label Widgets
        self.unit_id_label = Label(self.new_unit_details, text="Unit ID:", width=15, anchor="e")
        self.unit_id_label.grid(row=0, column=0, columnspan=1, rowspan=1)
        self.unit_name_label = Label(self.new_unit_details, text="Unit:", width=15, anchor="e")
        self.unit_name_label.grid(row=1, column=0, columnspan=1, rowspan=1)
        self.unit_house_label = Label(self.new_unit_details, text="Unit House:", width=15, anchor="e")
        self.unit_house_label.grid(row=2, column=0, columnspan=1, rowspan=1)
        self.unit_road_label = Label(self.new_unit_details, text="Unit Road:", width=15, anchor="e")
        self.unit_road_label.grid(row=3, column=0, columnspan=1, rowspan=1)
        self.unit_postcode_label = Label(self.new_unit_details, text="Unit Postcode:", width=15, anchor="e")
        self.unit_postcode_label.grid(row=4, column=0, columnspan=1, rowspan=1)
        self.unit_ownership_label = Label(self.new_unit_details, text="Ownership:", width=15, anchor="e")
        self.unit_ownership_label.grid(row=5, column=0, columnspan=1, rowspan=1)

        # Create Tkinter Entry Widgets
        self.unit_name_entry = Entry(self.new_unit_details, width=30)
        self.unit_name_entry.grid(row=1, column=1, columnspan=2, rowspan=1)
        self.unit_house_entry = Entry(self.new_unit_details, width=30)
        self.unit_house_entry.grid(row=2, column=1, columnspan=2, rowspan=1)
        self.unit_road_entry = Entry(self.new_unit_details, width=30)
        self.unit_road_entry.grid(row=3, column=1, columnspan=2, rowspan=1)
        self.unit_postcode_entry = Entry(self.new_unit_details, width=30)
        self.unit_postcode_entry.grid(row=4, column=1, columnspan=2, rowspan=1)
        self.unit_ownership_entry = Entry(self.new_unit_details, width=30)
        self.unit_ownership_entry.grid(row=5, column=1, columnspan=2, rowspan=1)

        # Hide root window while Toplevel open
        root.withdraw()
        # Handle "x" on Toplevel
        self.new_unit_container.protocol("WM_DELETE_WINDOW", self.close_window)

    def close_window(self):
        # Handle closing Toplevel event.
        self.new_unit_container.destroy()
        # Make root window visible on close
        root.deiconify()


class SearchTenant:
    def __init__(self, master):

        self.search_container = Toplevel(master)
        self.search_container.title("Century Wharf - Search Tenants")
        self.search_container.geometry("1200x620")
        self.search_container.resizable(False, False)
        self.search_container.config(background="white")
        self.search_container.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8, 9), weight=1, uniform="cols")
        self.search_container.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7, 8), weight=1, uniform="rows")
        self.search_container.focus_force()
        self.master = master

        # Gets the requested values of the height and width.
        self.windowWidth = 1200
        self.windowHeight = 620
        # Gets both half the screen width/height and window width/height
        self.positionRight = int(root.winfo_screenwidth() / 2 - self.windowWidth / 2)
        self.positionDown = int(root.winfo_screenheight() / 2 - self.windowHeight / 2)
        # Positions the window in the center of the page.
        self.search_container.geometry("+{}+{}".format(self.positionRight, self.positionDown))

        # Hide root window while Toplevel open
        root.withdraw()
        # Handle "x" on Toplevel
        self.search_container.protocol("WM_DELETE_WINDOW", self.close_window)

        self.search_critera_frame = LabelFrame(self.search_container, text="Search Critera", font=myFont,
                                               background="white")
        self.search_critera_frame.grid_columnconfigure((0, 1, 2, 3, 4, 5, 7, 8), weight=1, uniform="cols")
        self.search_critera_frame.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6), weight=1, uniform="rows")
        self.search_critera_frame.grid_propagate(0)
        self.search_critera_frame.grid(row=0, column=0, columnspan=10, rowspan=3, sticky="nswe", padx=10, pady=(5, 10))

        self.search_results_frame = LabelFrame(self.search_container, text="Search Results", font=myFont,
                                               background="white")
        self.search_results_frame.grid_columnconfigure(0, weight=1)
        self.search_results_frame.grid_rowconfigure(0, weight=1)
        self.search_results_frame.grid_propagate(0)
        self.search_results_frame.grid(row=3, column=0, columnspan=10, rowspan=6, sticky="nswe", padx=10)

        self.close_button = Button(self.search_container, text="Close", height=button_height, font=myFont,
                                   background="cornflower blue", command=lambda: self.close_window())
        self.close_button.grid(row=9, column=9, columnspan=1, rowspan=1, sticky="we", padx=(0, 10), pady=10)

        self.copy_email_addresses = Button(self.search_container, text="Copy All Email Addresses", font=myFont,
                                           height=button_height, state="disabled", background="cornflower blue",
                                           command=lambda: self.email_result_tenants())
        self.copy_email_addresses.grid(row=9, column=0, columnspan=2, rowspan=1, sticky="ew", padx=(10, 0))
        self.copy_selected_emails = Button(self.search_container, text="Copy Selected Email Addresses", font=myFont,
                                           height=button_height, state="disabled", background="cornflower blue",
                                           command=lambda: self.copy_selected())
        self.copy_selected_emails.grid(row=9, column=2, columnspan=2, rowspan=1, sticky="ew", padx=10)
        self.save_button = Button(self.search_container, text="Export", height=button_height,
                                  background="cornflower blue", state="disabled", font=myFont,
                                  command=lambda: self.save_report())
        self.save_button.grid(row=9, column=8, columnspan=1, rowspan=1, padx=(0, 10), sticky="we")

        self.date_label_before = Label(self.search_critera_frame, text="Before:", font=myFont, background="white")
        self.date_label_before.grid(row=0, column=5, sticky="ew")
        self.date_label_after = Label(self.search_critera_frame, text="After:", font=myFont, background="white")
        self.date_label_after.grid(row=0, column=7, sticky="ew")

        self.db_field_values = ["Unit", "Unit_Court", "Unit_House", "Name", "Email_1", "Last_Serviced_ISO",
                                "Service_Due_ISO"]
        self.field_logic_values = ["Contains", "Equal", "Not Equal", "Greater Than", "Less Than"]
        self.row_logic_values = ["And", "Or"]
        self.dict_search_strings = {}

        for i in range(len(self.db_field_values)):

            if i <= 4:
                self.dict_search_strings[self.db_field_values[i]] = StringVar()
                self.dict_search_strings[self.db_field_values[i]].set("")

                self.unit_field_label = Label(self.search_critera_frame, text=self.db_field_values[i] + ":",
                                              font=myFont, background="white", anchor="e")
                self.unit_field_label.grid(row=i, column=0, sticky="nsew")

                self.unit_field_search_string = Entry(self.search_critera_frame, font=myFont,
                                                      textvariable=self.dict_search_strings[self.db_field_values[i]],
                                                      background="white")
                self.unit_field_search_string.grid(row=i, column=1, columnspan=3, sticky="ew")
                self.unit_field_search_string.bind('<Return>', self.search_db)
            else:
                self.dict_search_strings[self.db_field_values[i]] = {}
                self.dict_search_strings[self.db_field_values[i]]["before"] = StringVar()
                self.dict_search_strings[self.db_field_values[i]]["before"].set("")

                self.unit_field_label = Label(self.search_critera_frame, text=self.db_field_values[i] + ":",
                                              font=myFont, background="white", anchor="e")
                self.unit_field_label.grid(row=i - 4, column=4, sticky="nsew")

                self.date_picker_before = DateEntry(self.search_critera_frame, background="cornflower blue",
                                                    textvariable=self.dict_search_strings[self.db_field_values[i]][
                                                        "before"], foreground="black", date_pattern='dd/MM/yyyy')
                if os_type != "Windows":
                    self.date_picker_before.top_cal.overrideredirect(False)
                self.date_picker_before.configure(validate='none')
                self.date_picker_before.grid(row=i - 4, column=5, sticky="ew")
                self.dict_search_strings[self.db_field_values[i]]["before"].set("")
                self.date_picker_before.bind('<Return>', self.search_db)

                self.and_label = Label(self.search_critera_frame, background="white", text="And", font=myFont)
                self.and_label.grid(row=i - 4, column=6, sticky="nsew")

                self.dict_search_strings[self.db_field_values[i]]["after"] = StringVar()
                self.dict_search_strings[self.db_field_values[i]]["after"].set("")

                self.date_picker_after = DateEntry(self.search_critera_frame, background="cornflower blue",
                                                   textvariable=self.dict_search_strings[self.db_field_values[i]][
                                                       "after"], foreground="black", date_pattern='dd/MM/yyyy')
                if os_type != "Windows":
                    self.date_picker_after.top_cal.overrideredirect(False)
                self.date_picker_after.configure(validate='none')
                self.date_picker_after.grid(row=i - 4, column=7, sticky="ew", padx=(0, 5))
                self.date_picker_after.bind('<Return>', self.search_db)
                self.dict_search_strings[self.db_field_values[i]]["after"].set("")

        self.no_checkbox_val = IntVar()
        self.no_checkbox_val.set(1)

        self.no_date_checkbox = Checkbutton(self.search_critera_frame, background="white", text="Include No Date",
                                            variable=self.no_checkbox_val, font=myFont)
        self.no_date_checkbox.grid(row=1, column=8, columnspan=1, rowspan=2)

        self.rtm_status_label = Label(self.search_critera_frame, background="white", text="RTM Member:", font=myFont,
                                      anchor="e")
        self.rtm_status_label.grid(row=5, column=0, sticky="nsew")

        self.rtm_search_val = StringVar()
        self.rtm_search_val.set("Both")

        self.rtm_member_search_yes = Radiobutton(self.search_critera_frame, text="Yes", value="Yes", font=myFont,
                                                 background="white", variable=self.rtm_search_val)
        self.rtm_member_search_yes.grid(row=5, column=1, columnspan=1, rowspan=1, pady=2)
        self.rtm_member_search_no = Radiobutton(self.search_critera_frame, text="No", value="No", font=myFont,
                                                background="white", variable=self.rtm_search_val)
        self.rtm_member_search_no.grid(row=5, column=2, columnspan=1, rowspan=1, pady=2)
        self.rtm_member_search_both = Radiobutton(self.search_critera_frame, text="Both", value="Both",
                                                  background="white", font=myFont, variable=self.rtm_search_val)
        self.rtm_member_search_both.grid(row=5, column=3, columnspan=1, rowspan=1, pady=2)

        self.which_email_label = Label(self.search_critera_frame, text="Display Email:", background="white",
                                       font=myFont, anchor="e")
        self.which_email_label.grid(row=4, column=4, sticky="nsew")

        self.rtm_email_select = StringVar()
        self.rtm_email_select.set("Owner")

        self.rtm_email_yes = Radiobutton(self.search_critera_frame, text="Owners E-Mail", value="Owner",
                                         background="white", font=myFont, variable=self.rtm_email_select)
        self.rtm_email_yes.grid(row=4, column=5, columnspan=1, rowspan=1)
        self.rtm_email_no = Radiobutton(self.search_critera_frame, text="RTM E-Mail (If Available)", value="RTM",
                                        background="white", font=myFont, variable=self.rtm_email_select)
        self.rtm_email_no.grid(row=4, column=6, columnspan=2, rowspan=1)

        self.rtm_group_val_one = IntVar()
        self.rtm_group_val_two = IntVar()
        self.rtm_group_val_three = IntVar()

        self.rtm_group_val_one.set(1)
        self.rtm_group_val_two.set(1)
        self.rtm_group_val_three.set(1)

        self.rtm_group_search_label = Label(self.search_critera_frame, text="RTM Group:", font=myFont,
                                            background="white", anchor="e")
        self.rtm_group_search_label.grid(row=6, column=0, columnspan=1, rowspan=1, sticky="ew")

        self.rtm_group_one_check = Checkbutton(self.search_critera_frame, background="white", text="1",
                                               variable=self.rtm_group_val_one, font=myFont)
        self.rtm_group_one_check.grid(row=6, column=1, columnspan=1, rowspan=1)
        self.rtm_group_two_check = Checkbutton(self.search_critera_frame, background="white", text="2",
                                               variable=self.rtm_group_val_two, font=myFont)
        self.rtm_group_two_check.grid(row=6, column=2, columnspan=1, rowspan=1)
        self.rtm_group_three_check = Checkbutton(self.search_critera_frame, background="white", text="3",
                                                 variable=self.rtm_group_val_three, font=myFont)
        self.rtm_group_three_check.grid(row=6, column=3, columnspan=1, rowspan=1)

        self.num_results_label = Label(self.search_critera_frame, text="", background="white", font=myFontBold)
        self.num_results_label.grid(row=4, column=8, rowspan=2, sticky="nsew")

        self.search_button = Button(self.search_critera_frame, text="Search", font=myFont, background="cornflower blue",
                                    height=button_height, command=lambda: self.search_db())
        self.search_button.grid(row=5, column=7, rowspan=2, sticky="ew", pady=(0, 5), padx=(2, 15))
        self.clear_search_button = Button(self.search_critera_frame, text="Clear Search", font=myFont,
                                          background="cornflower blue", height=button_height,
                                          command=lambda: self.clear_search())
        self.clear_search_button.grid(row=5, column=5, rowspan=2, sticky="ew", pady=(0, 5), padx=2)

        self.search_filter = {}
        self.db_columns = ["Unit", "Unit_Court", "Unit_House", "Name", "Phone", "Email_1", "Last_Serviced",
                           "Service_Due"]

        self.search_results_treeview = ttk.Treeview(self.search_results_frame, columns=self.db_columns, show="headings")
        for col in self.db_columns:
            self.search_results_treeview.heading(col, text=col)
        self.search_results_treeview.column("Unit", width=35, stretch=NO)
        self.search_results_treeview.column("Unit_Court", width=100)
        self.search_results_treeview.column("Unit_House", width=100)
        self.search_results_treeview.column("Name", width=250)
        self.search_results_treeview.column("Phone", width=100)
        self.search_results_treeview.column("Email_1", width=200)
        self.search_results_treeview.column("Last_Serviced", width=100)
        self.search_results_treeview.column("Service_Due", width=100)
        self.search_results_treeview.grid(row=0, column=0, sticky="nsew")
        self.search_results_treeview.heading("Email_1", text="E-Mail")

        self.vert_scroll_bar = ttk.Scrollbar(self.search_results_frame, orient="vertical",
                                             command=self.search_results_treeview.yview)
        self.vert_scroll_bar.grid(row=0, column=10, sticky='ns')
        self.search_results_treeview.configure(yscrollcommand=self.vert_scroll_bar.set)
        self.search_results_treeview.bind('<Double-Button-1>', self.select_item)

    def copy_selected(self):

        selected_items = self.search_results_treeview.selection()
        emails_selected = []
        if len(selected_items) != 0:
            for item in selected_items:
                if self.search_results_treeview.item(item)["values"][5] != "":
                    emails_selected.append(self.search_results_treeview.item(item)["values"][5])
            if len(emails_selected) != 0:
                selected_results = pd.DataFrame(emails_selected)
                selected_results.to_clipboard(excel=True, sep=',', index=False, header=False)
                if len(emails_selected) == 1:
                    messagebox.showinfo("Clipboard", str(len(emails_selected)) + " Email Address Copied to Clipboard",
                                        parent=self.search_container)
                elif len(emails_selected) > 1:
                    messagebox.showinfo("Clipboard", str(len(emails_selected)) + " Email Addresses Copied to Clipboard",
                                        parent=self.search_container)

    def save_report(self):

        if len(self.search_results) != 0:

            self.save_directory = filedialog.asksaveasfilename(initialfile="Search Results", defaultextension=".csv")
            if self.save_directory != '':
                self.report_list = self.search_results
                for row in self.report_list:
                    del row["_id"]
                    del row["Unit_Full"]
                    del row["RTM_group"]
                    del row["Ownership"]
                    del row["Director"]
                    del row["Service_Due_Month"]
                    del row["Service_Due_Year"]
                    row["Notes"] = row["Notes"].replace('\n', ' ')
                    row["Notes"] = row["Notes"].replace('\t', ' ')
                    if datetime.strftime(row["Last_Serviced_ISO"], "%d/%m/%Y") != "01/01/1970":
                        row["Service_Due_ISO"] = datetime.strftime(row["Service_Due_ISO"], "%d/%m/%Y")
                        row["Last_Serviced_ISO"] = datetime.strftime(row["Last_Serviced_ISO"], "%d/%m/%Y")
                    else:
                        row["Service_Due_ISO"] = ""
                        row["Last_Serviced_ISO"] = ""

                self.csv_keys = self.report_list[0].keys()
                with open(self.save_directory, 'w', newline='') as output_file:
                    dict_writer = csv.DictWriter(output_file, fieldnames=self.csv_keys)
                    dict_writer.writeheader()
                    dict_writer.writerows(self.report_list)

    def email_result_tenants(self):

        emails_selected = []
        if len(self.search_results) != 0:
            for item in self.search_results_treeview.get_children():
                if self.search_results_treeview.item(item)["values"][5] != "":
                    emails_selected.append(self.search_results_treeview.item(item)["values"][5])
            if len(emails_selected) != 0:
                selected_results = pd.DataFrame(emails_selected)
                selected_results.to_clipboard(excel=True, sep=',', index=False, header=False)
                if len(emails_selected) == 1:
                    messagebox.showinfo("Clipboard", str(len(emails_selected)) + " Email Address Copied to Clipboard",
                                        parent=self.search_container)
                elif len(emails_selected) > 1:
                    messagebox.showinfo("Clipboard", str(len(emails_selected)) + " Email Addresses Copied to Clipboard",
                                        parent=self.search_container)

    def select_item(self, *args):
        cur_item = self.search_results_treeview.focus()
        SearchTenant.unit_num_selected = self.search_results_treeview.item(cur_item)["values"][0]
        SearchTenant.unit_court_selected = self.search_results_treeview.item(cur_item)["values"][1]

        ViewEditTenant(self.master, "Search", self)
        self.search_container.withdraw()

    def populate_treeview(self):

        self.search_results_treeview.delete(*self.search_results_treeview.get_children())
        # noinspection PyUnresolvedReferences
        try:
            self.search_results = list(Database.crud_app_collection.find(self.search_filter).sort(
                [("Unit_Court", pymongo.ASCENDING), ("Unit", pymongo.ASCENDING)]))
        except pymongo.errors.PyMongoError as e:
            return messagebox.showerror("Exception", e, parent=self.search_container)

        else:
            Database.log_action("Search", str(self.search_filter))
            for i in range(len(self.search_results)):
                self.serviced_last = "-"
                self.serviced_due = "-"
                if datetime.strftime(self.search_results[i]["Last_Serviced_ISO"], "%d/%m/%Y") != "01/01/1970":
                    self.serviced_last = datetime.strftime(self.search_results[i]["Last_Serviced_ISO"], "%d/%m/%Y")
                    self.serviced_due = datetime.strftime(self.search_results[i]["Service_Due_ISO"], "%d/%m/%Y")

                if self.rtm_email_select.get() == "Owner":
                    self.search_results_treeview.insert("", i, values=[self.search_results[i]["Unit"],
                                                                       self.search_results[i]["Unit_Court"],
                                                                       self.search_results[i]["Unit_House"],
                                                                       self.search_results[i]["Name"],
                                                                       self.search_results[i]["Phone"],
                                                                       self.search_results[i]["Email_1"],
                                                                       self.serviced_last, self.serviced_due])

                elif self.rtm_email_select.get() == "RTM":
                    if self.search_results[i]["RTM_Email"] == "":
                        display_email = self.search_results[i]["Email_1"]
                    else:
                        display_email = self.search_results[i]["RTM_Email"]
                    self.search_results_treeview.insert("", i, values=[self.search_results[i]["Unit"],
                                                                       self.search_results[i]["Unit_Court"],
                                                                       self.search_results[i]["Unit_House"],
                                                                       self.search_results[i]["Name"],
                                                                       self.search_results[i]["Phone"],
                                                                       display_email, self.serviced_last,
                                                                       self.serviced_due])

            self.num_results_label.configure(text="Results: " + str(len(self.search_results)))
            if len(self.search_results) != 0:
                self.copy_email_addresses.configure(state="normal")
                self.save_button.configure(state="normal")
                self.copy_selected_emails.configure(state="normal")
            elif len(self.search_results) == 0:
                self.copy_email_addresses.configure(state="disabled")
                self.save_button.configure(state="disabled")
                self.copy_selected_emails.configure(state="disabled")

    def clear_search(self):

        self.search_results_treeview.delete(*self.search_results_treeview.get_children())

        for key, value in self.dict_search_strings.items():
            if isinstance(value, dict):
                self.dict_search_strings[key]["before"].set("")
                self.dict_search_strings[key]["after"].set("")
            else:
                self.dict_search_strings[key].set("")

        self.num_results_label.configure(text="")
        self.copy_email_addresses.configure(state="disabled")
        self.save_button.configure(state="disabled")
        self.copy_selected_emails.configure(state="disabled")

        self.rtm_search_val.set("Both")
        self.rtm_group_val_one.set(1)
        self.rtm_group_val_two.set(1)
        self.rtm_group_val_three.set(1)
        self.rtm_email_select.set("Owner")
        self.no_checkbox_val.set(1)

    def search_db(self, *args):

        self.search_filter = {"$and": []}

        if self.no_checkbox_val.get() == 0:
            self.search_filter["$and"].append(
                {"Last_Serviced_ISO": {"$ne": datetime.strptime("01/01/1970", "%d/%m/%Y")}})

        for key, value in self.dict_search_strings.items():

            if isinstance(value, dict):
                if self.dict_search_strings[key]["before"].get() != "":
                    self.search_filter["$and"].append(
                        {key: {"$lte": datetime.strptime(self.dict_search_strings[key]["before"].get(), "%d/%m/%Y")}})
                if self.dict_search_strings[key]["after"].get() != "":
                    self.search_filter["$and"].append(
                        {key: {"$gte": datetime.strptime(self.dict_search_strings[key]["after"].get(), "%d/%m/%Y")}})

            elif value.get() != "":
                if key == "Unit":
                    self.search_filter["$and"].append({key: {"$eq": int(value.get().strip())}})
                else:
                    search_terms = value.get().strip().split()
                    for i in range(len(search_terms)):
                        self.search_filter["$and"].append({key: {"$regex": search_terms[i], "$options": "i"}})

        if self.rtm_search_val.get() != "Both":
            self.search_filter["$and"].append({"RTM_Member": {"$eq": self.rtm_search_val.get()}})

        if self.rtm_group_val_one.get() != 1:
            self.search_filter["$and"].append({"RTM_group": {"$ne": 1}})
        if self.rtm_group_val_two.get() != 1:
            self.search_filter["$and"].append({"RTM_group": {"$ne": 2}})
        if self.rtm_group_val_three.get() != 1:
            self.search_filter["$and"].append({"RTM_group": {"$ne": 3}})

        if self.search_filter["$and"]:
            self.populate_treeview()
        else:
            self.search_filter = {}
            self.populate_treeview()

    def close_window(self):
        # Handle closing Toplevel event.
        self.search_container.destroy()
        # Make root window visible on close
        root.deiconify()


app = Mainwindow(root)
dpi = root.winfo_fpixels('1i')
scale_factor = dpi / 72

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
# root.call('tk', 'scaling', scale_factor)
root.geometry("920x590")
root.resizable(False, False)
root.config(background="white")
# Gets the requested values of the height and width.
windowWidth = 920
windowHeight = 590
# Gets both half the screen width/height and window width/height
positionRight = int(root.winfo_screenwidth() / 2 - windowWidth / 2)
positionDown = int(root.winfo_screenheight() / 2 - windowHeight / 2)
# Positions the window in the center of the page.
root.geometry("+{}+{}".format(positionRight, positionDown))
root.iconphoto(True, PhotoImage(file='graphics/icon.png'))
root.title("Boiler Service App                   Version: " + app_version)
root.withdraw()
root.mainloop()
