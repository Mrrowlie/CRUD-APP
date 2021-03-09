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


class Database:

    def make_connection():
        # Make Database Connection Object
        try:
            # Attempt To Make Connection
            Database.username_val = ""
            Database.password_val = ""

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
            return True

    def upload_certificate():

        currentFile = __file__
        realPath = os.path.realpath(currentFile)
        dirPath = os.path.dirname(realPath)
        print(dirPath)

        for entry in os.scandir(dirPath):


            if (entry.path.endswith(".pdf")) and entry.is_file():
                print(entry)

                for document in Database.crud_app_collection.find():

                    file_field = ""
                    for i in range(1, 11):
                        if document["File_" + str(i)] == "":
                            file_field = "File_" + str(i)
                            break
                        else:
                            continue

                    if file_field == "":
                        print("File limited reached!", "Please delete a file. Maximum files per Unit is 10.")
                    else:
                        try:
                            with open(entry.path, "rb") as input_file:
                                upload_object = Database.certificate_storage.put(input_file, content_type="application/pdf",
                                                                                 filename=os.path.basename(
                                                                                     entry))

                            cert_url_update = Database.crud_app_collection.update_one(
                                document,
                                {"$set": {file_field: upload_object}})
                        except pymongo.errors.PyMongoError as e:
                            return print("Exception " + e)

                        else:
                            pass


Database.make_connection()
Database.upload_certificate()
