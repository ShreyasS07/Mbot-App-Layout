import sys
from datetime import datetime
import traceback
import mBotEnc

try:
    import gc
    import tkinter as tk
    from tkinter import *
    from tkinter import ttk
    from tkinter import messagebox
    from tkinter.messagebox import askyesno
    from tkinter import filedialog as fd
    from PIL import Image, ImageTk
    import App_Version as ver
    from mBotEnc import encData, decData, NewList       # Importing Enc & dec functions
    import base64

    import json
    import time
    import shutil
    import wmi
    import os
    import re
    import copy
    import psutil

    loopCount = 0
    runTimer = 10  # in sec
    treadLoop = None
    try:
        os.remove("superError.txt")
    except:
        pp = 0

    class Tk_int:

        def __init__(self, root, style):
            already_running = self.appRunning()
            self.error_txt_file = "errorLog.txt"

            self.XLfont = ("arial", 13)
            self.XLfontB = ("arial", 13, "bold")

            self.Lfont = ("arial", 12)
            self.LfontB = ("arial", 12, "bold")

            self.Mfont = ("arial", 11)
            self.MfontB = ("arial", 11, "bold")

            self.Sfont = ("arial", 10)
            self.SfontB = ("arial", 10, "bold")

            self.XSfont = ("arial", 9)
            self.XSfontB = ("arial", 9, "bold")

            expireDate = ver.expireDate
            now = datetime.now()
            d1 = datetime.strptime(now.strftime("%Y/%m/%d"), "%Y/%m/%d")
            d2 = datetime.strptime(expireDate, "%Y/%m/%d")
            delta = d2 - d1
            status = delta.days

            self.root = root
            self.style = style
            title = ver.title
            self.root.configure(background='white')


            self.processJsonPath = "ProcessList.json"
            self.taskJsonPath = "TaskList.json"
            # self.rawprocessJsonPath = "RawProcessList.json"
            # self.rawtaskJsonPath = "RawTaskList.json"


            if (delta.days <= 0):
                self.root.title(title + "(" + str(status) + ")")
                status = 0
            else:
                self.root.title(title)
            self.w = 300
            self.h = 60
            self.x = 0
            self.y = 0
            self.root.resizable(0, 0)
            self.root.configure(background='white')
            # self.root.iconphoto(False, tk.PhotoImage(file='icons/icon.png'))
            self.root.iconphoto(False, tk.PhotoImage(file="icons\\icon.png"))
            self.root.geometry('%dx%d+%d+%d' % (self.w, self.h, self.x, self.y))
            if (status == 0):
                msg = "APP EXPIRED"
                l2 = tk.Label(self.root, anchor="c", width=30, background='white', height=2, text=msg, wraplength=300, font=self.XSfont)
                l2.grid(row=1, column=0)
            elif (already_running > 1):
                msg = "APP ALREADY RUNNING.\n INSTANCE : " + str(already_running)
                l2 = tk.Label(self.root, anchor="c", width=30, background='white', height=2, text=msg, wraplength=300, font=self.XSfont)
                l2.grid(row=1, column=0)
            else:
                try:
                    os.remove(self.error_txt_file)
                except:
                    pp = 0
                f = open(self.error_txt_file, "w")
                f.write("")
                f.close()

                taskbar_height = 80
                screen_tolerance = 10
                # screen_width = 1000       #self.root.winfo_screenwidth()
                # screen_height = 600       #self.root.winfo_screenheight()
                # self.w = int(screen_width) - screen_tolerance
                # self.h = int(screen_height) - taskbar_height
                self.w = int(self.root.winfo_screenwidth()) - screen_tolerance
                self.h = int(self.root.winfo_screenheight()) - taskbar_height
                self.root.geometry('%dx%d+%d+%d' % (self.w, self.h, self.x, self.y))

                self.Mframe = Frame(self.root, relief=SOLID)
                self.Mframe.place(x=0, y=0, width=self.w, height=self.h)
                self.Mframe.configure(background='white')
                self.Tframe_height = 60
                self.scroll_width = 16
                self.scroll_tolerance = self.scroll_width + 13

                self.number_of_portions = 4
                self.portion_of_width = int(self.w / self.number_of_portions)
                self.Lframe_width = self.portion_of_width

                # FRAME TOP
                self.Tframe = Frame(self.Mframe, relief=SUNKEN, borderwidth=1, background="white")
                self.Tframe.place(x=0, y=0, width=self.Lframe_width, height=self.Tframe_height)
                self.Tframe.configure(background='white')
                # im = Image.open("icons/logo-full.png")
                im = Image.open("icons\\logo-large.png")
                img = im.resize((175, 53), Image.Resampling.LANCZOS)
                img = ImageTk.PhotoImage(img)
                im_button = tk.Button(self.Tframe, bg="white", image=img, borderwidth=1, relief="flat")
                im_button.grid(row=1, column=0)
                im_button.image = img

                # self.message_display = tk.StringVar()
                # self.message_display.set("WELCOME")
                # self.message_label = Label(self.Tframe, fg='black', bg='white', height=2, font=self.Sfont, textvariable=self.message_display)
                # self.message_label.grid(row=1, column=1, sticky=W, padx=5, pady=(3, 0))

                # FRAME LEFT
                self.Lframe = Frame(self.Mframe, relief=SUNKEN, borderwidth=0)
                self.Lframe.place(x=0, y=self.Tframe_height, width=self.Lframe_width, height=self.h - self.Tframe_height)
                # self.Lframe_cnv = tk.Canvas(self.Lframe, width=self.scroll_width )
                # self.Lframe_cnv.configure(background='white')
                # self.Lframe_frm = tk.Frame(self.Lframe_cnv)
                # self.Lframe_frm.configure(background='white')
                # self.Lframe_vsrl = tk.Scrollbar(self.Lframe)
                # self.Lframe_cnv.config(yscrollcommand=self.Lframe_vsrl.set, highlightthickness=1)
                # self.Lframe_vsrl.config(orient=tk.VERTICAL, command=self.Lframe_cnv.yview)
                # self.Lframe_vsrl.pack(fill=tk.Y, side=tk.RIGHT, expand=tk.FALSE)
                # self.Lframe_vsrl.config(width= self.scroll_width)
                # self.Lframe_hsrl = tk.Scrollbar(self.Lframe)
                # self.Lframe_cnv.config(xscrollcommand=self.Lframe_hsrl.set, highlightthickness=1)
                # self.Lframe_hsrl.config(orient=tk.HORIZONTAL, command=self.Lframe_cnv.xview)
                # self.Lframe_hsrl.pack(fill=tk.X, side=tk.BOTTOM, expand=tk.FALSE)
                # self.Lframe_hsrl.config(width= self.scroll_width)
                # self.Lframe_cnv.pack(fill=tk.BOTH, side=tk.LEFT, expand=tk.TRUE)
                # self.Lframe_cnv.create_window(0, 0, window=self.Lframe_frm, anchor=tk.NW)

                # Button(self.Lframe, fg='black', bg='white', height=2, font=self.Sfont, text="CREATE PROCESS",
                #        command=lambda: self.create_process_form()).grid(row=0, column=0, sticky=W, padx=5, pady=(3, 0))


                Process_options = ["CREATE PROCESS", "CREATE TASK","CREATE PARENT BOT & START TIME", "CREATE MINI BOT", "ASSIGN PROCESS & SCHEDULE"]
                self.event_name = ttk.Combobox(self.Lframe, background='white', values=Process_options, width=35,   font=self.Sfont)
                self.event_name['state'] = 'readonly'
                self.event_name.grid(row=1, column=0, columnspan=2, sticky="we", )
                self.event_name.bind("<<ComboboxSelected>>", self.choose_event)



                # FRAME 1 LEFT
                strt_pos = 50

                remainig_height = (self.h - self.Tframe_height - strt_pos)
                main_form_height = int(remainig_height / 2)
                self.L1frame = Frame(self.Lframe, relief=SUNKEN, borderwidth=0)
                self.L1frame.place(x=0, y=strt_pos, width=self.Lframe_width, height=main_form_height)
                self.L1frame_cnv = tk.Canvas(self.L1frame, width=self.scroll_width)
                # self.L1frame_cnv.configure(background='red')
                self.L1frame_cnv.configure(background='white')
                self.L1frame_frm = tk.Frame(self.L1frame_cnv)
                self.L1frame_frm.configure(background='white')
                self.L1frame_vsrl = tk.Scrollbar(self.L1frame)
                self.L1frame_cnv.config(yscrollcommand=self.L1frame_vsrl.set, highlightthickness=1)
                self.L1frame_vsrl.config(orient=tk.VERTICAL, command=self.L1frame_cnv.yview)
                self.L1frame_vsrl.pack(fill=tk.Y, side=tk.RIGHT, expand=tk.FALSE)
                self.L1frame_vsrl.config(width=self.scroll_width)
                self.L1frame_hsrl = tk.Scrollbar(self.L1frame)
                self.L1frame_cnv.config(xscrollcommand=self.L1frame_hsrl.set, highlightthickness=1)
                self.L1frame_hsrl.config(orient=tk.HORIZONTAL, command=self.L1frame_cnv.xview)
                self.L1frame_hsrl.pack(fill=tk.X, side=tk.BOTTOM, expand=tk.FALSE)
                self.L1frame_hsrl.config(width=self.scroll_width)
                self.L1frame_cnv.pack(fill=tk.BOTH, side=tk.LEFT, expand=tk.TRUE)
                self.L1frame_cnv.create_window(0, 0, window=self.L1frame_frm, anchor=tk.NW)

                # for ii in range(1, 50):
                #     for i2 in range(10):
                #         Label(self.L1frame_frm, fg='black', bg='white', height=2, font=self.Sfont,
                #               text="" + str(ii) + "," + str(i2) + " L1").grid(row=ii, column=i2, sticky=W, padx=5, pady=(3, 0))
                self.updateScrollRegion(self.L1frame_cnv, self.L1frame_frm)

                self.L2frame = Frame(self.Lframe, relief=SUNKEN, borderwidth=0)
                self.L2frame.place(x=0, y=strt_pos + main_form_height, width=self.Lframe_width, height=abs(remainig_height - main_form_height))
                self.L2frame_cnv = tk.Canvas(self.L2frame, width=self.scroll_width)
                # self.L2frame_cnv.configure(background='red')
                self.L2frame_cnv.configure(background='white')
                self.L2frame_frm = tk.Frame(self.L2frame_cnv)
                self.L2frame_frm.configure(background='white')
                self.L2frame_vsrl = tk.Scrollbar(self.L2frame)
                self.L2frame_cnv.config(yscrollcommand=self.L2frame_vsrl.set, highlightthickness=1)
                self.L2frame_vsrl.config(orient=tk.VERTICAL, command=self.L2frame_cnv.yview)
                self.L2frame_vsrl.pack(fill=tk.Y, side=tk.RIGHT, expand=tk.FALSE)
                self.L2frame_vsrl.config(width=self.scroll_width)
                self.L2frame_hsrl = tk.Scrollbar(self.L2frame)
                self.L2frame_cnv.config(xscrollcommand=self.L2frame_hsrl.set, highlightthickness=1)
                self.L2frame_hsrl.config(orient=tk.HORIZONTAL, command=self.L2frame_cnv.xview)
                self.L2frame_hsrl.pack(fill=tk.X, side=tk.BOTTOM, expand=tk.FALSE)
                self.L2frame_hsrl.config(width=self.scroll_width)
                self.L2frame_cnv.pack(fill=tk.BOTH, side=tk.LEFT, expand=tk.TRUE)
                self.L2frame_cnv.create_window(0, 0, window=self.L2frame_frm, anchor=tk.NW)

                # for ii in range(1, 50):
                #     for i2 in range(10):
                #         Label(self.L2frame_frm, fg='black', bg='white', height=2, font=self.Sfont,
                #               text=str(ii) + "," + str(i2) + " L2  ").grid(row=ii, column=i2, sticky=W, padx=5, pady=(3, 0))
                self.updateScrollRegion(self.L2frame_cnv, self.L2frame_frm)

                # RIGHT FRAME
                self.Rframe_width = self.w - self.Lframe_width - self.scroll_width
                self.Rframe = Frame(self.Mframe, relief=SUNKEN, borderwidth=0)
                self.Rframe.place(x=self.Lframe_width, y=0, width=self.Rframe_width, height=self.h)
                self.Rframe_cnv = tk.Canvas(self.Rframe, width=self.scroll_width)
                self.Rframe_cnv.configure(background='white')
                self.Rframe_frm = tk.Frame(self.Rframe_cnv)
                self.Rframe_frm.configure(background='white')
                self.Rframe_vsrl = tk.Scrollbar(self.Rframe)
                self.Rframe_cnv.config(yscrollcommand=self.Rframe_vsrl.set, highlightthickness=1)
                self.Rframe_vsrl.config(orient=tk.VERTICAL, command=self.Rframe_cnv.yview)
                self.Rframe_vsrl.pack(fill=tk.Y, side=tk.RIGHT, expand=tk.FALSE)
                self.Rframe_vsrl.config(width=self.scroll_width)
                self.Rframe_hsrl = tk.Scrollbar(self.Rframe)
                self.Rframe_cnv.config(xscrollcommand=self.Rframe_hsrl.set, highlightthickness=1)
                self.Rframe_hsrl.config(orient=tk.HORIZONTAL, command=self.Rframe_cnv.xview)
                self.Rframe_hsrl.pack(fill=tk.X, side=tk.BOTTOM, expand=tk.FALSE)
                self.Rframe_hsrl.config(width=self.scroll_width)
                self.Rframe_cnv.pack(fill=tk.BOTH, side=tk.LEFT, expand=tk.TRUE)
                self.Rframe_cnv.create_window(0, 0, window=self.Rframe_frm, anchor=tk.NW)

                # for ii in range(1, 50):
                #     for i2 in range(50):
                #         Label(self.Rframe_frm, fg='black', bg='white', height=2, font=self.Sfont,
                #               text=str(ii) + "," + str(i2) + " R  ").grid(row=ii, column=i2, sticky=W, padx=5,
                #                                                           pady=(3, 0))
                self.updateScrollRegion(self.Rframe_cnv, self.Rframe_frm)

            self.root.protocol('WM_DELETE_WINDOW', self.destroy_me)
            self.root.update_idletasks()
            self.root.update()


        def choose_event(self, event):
            self.selected_event = self.event_name.get()
            # print(event)
            for wdg in self.Rframe_frm.winfo_children():
                wdg.destroy()
            for wdg in self.L1frame_frm.winfo_children():
                wdg.destroy()
            for wdg in self.L2frame_frm.winfo_children():
                wdg.destroy()
            if self.selected_event == 'CREATE PROCESS':
                # print("Creating Process from")
                self.create_process_form()                          # Create process form
            elif self.selected_event == 'CREATE TASK':
                self.create_task_form()                             # Create task form
            elif self.selected_event == 'CREATE PARENT BOT & START TIME':
                print("No function Created")                          # Create Parent BOT
            elif self.selected_event == 'CREATE MINI BOT':
                print("No function Created")                         # Creating Mini bot
            elif self.selected_event == 'ASSIGN PROCESS & SCHEDULE':
                print("No function Created")                          # Assigning Process


        """ Process form """

        def readprocessJson(self):
            try:
                # Raw Process Json
                # with open(f'{self.rawprocessJsonPath}', 'r') as file:
                #     rawProcessJson = json.load(file)

                # Enc Process JSON
                with open(f'{self.processJsonPath}', 'r') as file:
                    encProcessJson = json.load(file)
                ProcessJson = mBotEnc.decJsonData(encProcessJson)
            except:
                rawProcessJson = []
                ProcessJson = []

            return ProcessJson
            # return rawProcessJson

        def writeprocessJson(self, ProcessJsonData):
            # RAW process Json
            # with open(self.rawprocessJsonPath, "w") as outfile:
            #     outfile.write(json.dumps(ProcessJsonData, indent=4))

            # Enc Process JSON
            encJson = mBotEnc.encJsonData(ProcessJsonData)
            with open(self.processJsonPath, "w") as outfile:
                outfile.write(json.dumps(encJson, indent=4))
            # print("ENCRYPT PROCESS", encJson)

 # -------------------------------------------------------------
        # Process Json without Encryption

        # def readprocessJson(self):
        #     try:
        #         with open(f'{self.rawprocessJsonPath}', 'r') as file:
        #             ProcessJson = json.load(file)
        #     except:
        #         ProcessJson = []
        #     return ProcessJson
        #
        # def writeprocessJson(self, ProcessJsonData):
        #     with open(self.rawprocessJsonPath, "w") as outfile:
        #         outfile.write(json.dumps(ProcessJsonData, indent=4))

        def submit_process_form(self):
            global proc_name
            proc_name = self.process_name.get().strip()
            if(proc_name):
                ProcessJsonData = self.readprocessJson()
                formated = list(filter(lambda x: str(x[1]).lower() == str(proc_name).lower(), enumerate(ProcessJsonData)))
                if (formated):
                    print("ERROR: PROCESS NAME ALREADY EXIST")
                else:
                    # ProcessJsonData.append(f"Process Name: {proc_name}")
                    ProcessJsonData.append(proc_name)
                    self.writeprocessJson(ProcessJsonData)
                    print("SUCCESS: PROCES NAME CREATED")
            else:
                print("ERROR: PROCES NAME CANNOT BE EMPTY")
            self.show_process_data()
            # self.process_name.set("")
            self.process_name.delete(0, END)

        def show_process_data(self):
            r2ID = 0
            ProcessJsonData = self.readprocessJson()
            for pj in ProcessJsonData:
                Button(self.Rframe_frm, fg='black', bg='white', height=2, font=self.Sfont, text=pj).grid(row=r2ID, column=0, sticky=W, padx=20, pady=(3, 0))
                r2ID += 1

        def create_process_form(self):
            self.show_process_data()
            rId = 0
            tk.Label(self.L1frame_frm, text="PROCESS NAME", font=self.Sfont).grid(row=rId, column=0, padx=5, pady=5, sticky=W)
            rId += 1
            self.process_name = tk.Entry(self.L1frame_frm, width=35, font=self.Sfont)
            self.process_name.grid(row=rId, column=0, padx=5, pady=(0, 10))
            rId += 1
            tk.Button(self.L1frame_frm, text="Submit", font=self.Sfont, command=lambda:self.submit_process_form(),
                      width=30).grid(row=rId, column=0, padx=5, pady=10)


        """ TASK form """

        def readtaskJson(self):
            try:
                # RAW task json
                # with open(f'{self.rawtaskJsonPath}', 'r') as file:
                #     rawTaskJson = json.load(file)

                # ENC Task JSON
                with open(f'{self.taskJsonPath}', 'r') as file:
                    encTaskJson = json.load(file)
                TaskJson = mBotEnc.decJsonData(encTaskJson)
                # Saving Decoded data to Json
                # with open(self.taskJsonPath, 'w') as outfile:
                #     json.dump(decJson, outfile, indent=4)
            except:
                rawTaskJson = []
                TaskJson = []
            return TaskJson
            # return rawTaskJson

        def writetaskJson(self, TaskJsonData):
            #Raw task json
            # with open(self.rawtaskJsonPath, "w") as outfile:
            #     outfile.write(json.dumps(TaskJsonData, indent=4))

            # ENC TASK JSON
            encJson = mBotEnc.encJsonData(TaskJsonData)
            with open(self.taskJsonPath, "w") as outfile:
                outfile.write(json.dumps(encJson, indent=4))

# --------------------------------------------------
#         # Raw Task Json without Enc
#         def readtaskJson(self):
#             try:
#                 with open(f'{self.rawtaskJsonPath}', 'r') as file:
#                     TaskJson = json.load(file)
#             except:
#                 TaskJson = []
#             return TaskJson
#
#         def writetaskJson(self, TaskJsonData):
#             with open(self.rawtaskJsonPath, "w") as outfile:
#                 outfile.write(json.dumps(TaskJsonData, indent=4))


        def submit_task_form(self):
            tasks_name = self.task_name.get().strip()
            task_option = self.task_option.get().strip()
            choose_option = self.choose_option.get().strip()
            if tasks_name:
                TaskJsonData = self.readtaskJson()

                formated = list(filter(lambda x: isinstance(x, dict) and x.get('Task Name', '').lower() == str(tasks_name).lower(), TaskJsonData))
                if formated:
                    print("ERROR: TASK NAME ALREADY EXISTS")
                else:
                    task_data = {"Process Name": choose_option, "Task Name": tasks_name, "Task Type": task_option}
                    TaskJsonData.append(task_data)
                    self.writetaskJson(TaskJsonData)
                    print("SUCCESS: TASK NAME CREATED")

                    # Call the display_task_data function to display the data on l2 frame
                    self.display_task_data(task_option)
            else:
                print("ERROR: TASK NAME CANNOT BE EMPTY")
            self.show_task_data()
            # self.task_name.delete(0, END)         #Uncommenting this the TaskName will be saved to JSON

        def show_task_data(self):
            r2ID = 0
            TaskJsonData = self.readtaskJson()
            for pj in TaskJsonData:
                Button(self.Rframe_frm, fg='black', bg='white', height=2, font=self.Sfont, text=pj).grid(row=r2ID, column=0, sticky=W, padx=20, pady=(3, 0))
                # Button(self.Rframe_frm, fg='black', bg='white', height=2, font=self.Sfont, text=pj["Task Name"]).grid(row=r2ID, column=0, sticky=W, padx=20, pady=(3, 0))
                r2ID += 1


        def create_task_form(self):
            self.show_task_data()
            rId = 0
            tk.Label(self.L1frame_frm, text="CHOOSE PROCESS", font=self.Sfont).grid(row=rId, column=0, padx=5, pady=5, sticky=W)

            rId += 1
            process_names = self.readprocessJson()  # Getting process names from JSON
            self.choose_option = ttk.Combobox(self.L1frame_frm, values=process_names, font=self.Sfont, width=30)        # TAking process names as options
            self.choose_option.grid(row=rId, column=0, padx=5, pady=(0, 10))

            rId += 1
            tk.Label(self.L1frame_frm, text="TASK NAME", font=self.Sfont).grid(row=rId, column=0, padx=5, pady=5, sticky=W)
            rId += 1
            self.task_name = tk.Entry(self.L1frame_frm, width=35, font=self.Sfont)
            self.task_name.grid(row=rId, column=0, padx=5, pady=(0, 10))

            # Create dropdown menu to select taksk
            rId += 1
            tk.Label(self.L1frame_frm, text="TASK TYPE", font=self.Sfont).grid(row=rId, column=0, padx=5, pady=5, sticky=W)
            rId += 1
            self.task_option = ttk.Combobox(self.L1frame_frm, values=["Image Location", "Screen Location", "Excel Configuration",
                                                                      "Calender Location", "Others"], font=self.Sfont, width=30)
            self.task_option.grid(row=rId, column=0, padx=5, pady=(0, 10))
            # self.task_option.current(0)

            rId += 1
            tk.Button(self.L1frame_frm, text="Submit", font=self.Sfont, command=lambda: self.submit_task_form(),
                      width=30).grid(row=rId, column=0, padx=5, pady=10)



        def display_task_data(self, task_option):
            task_type = task_option
            for widget in self.L2frame_frm.winfo_children():
                widget.destroy()

            if task_type == "Image Location":
                try:
                    with open("JSON Files\\image_location.json", "r") as file:
                        data = json.load(file)
                    print("Image JSON file taken")
                except FileNotFoundError:
                    print("No file found.")
                    message = "No file found."
                    label = tk.Label(self.L2frame_frm, text=message, font=self.Sfont, width=20)
                    label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
                    return
                # Dropdown menu
                names = [item["name"] for item in data]
                self.dropdown = ttk.Combobox(self.L2frame_frm, values=names, font=self.Sfont, width=35)
                self.dropdown.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                self.dropdown.bind("<<ComboboxSelected>>", self.image_options)

            elif task_type == "Screen Location":
                try:
                    with open("JSON Files\\screen_location.json", "r") as file:
                        data = json.load(file)
                        print("Screen JSON file taken")
                except FileNotFoundError:
                    print("No file found.")
                    message = "No file found."
                    label = tk.Label(self.L2frame_frm, text=message, font=self.Sfont, width=20)
                    label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
                    return
                # Dropdown menu
                names = [item["name"] for item in data]
                self.dropdown = ttk.Combobox(self.L2frame_frm, values=names, font=self.Sfont, width=35)
                self.dropdown.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                self.dropdown.bind("<<ComboboxSelected>>", self.screen_options)

            elif task_type == "Excel Configuration":
                def select_image():
                    try:
                        with open("JSON Files\\image_location.json", "r") as file:
                            data = json.load(file)
                            print("Excel JSON taken")
                    except FileNotFoundError:
                        print("No file found.")
                        message = "No file found."
                        label = tk.Label(self.L2frame_frm, text=message, font=self.Sfont, width=20)
                        label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
                        return
                    # Dropdown menu
                    names = [item["name"] for item in data]
                    self.dropdown = ttk.Combobox(self.L2frame_frm, values=names, font=self.Sfont, width=30)
                    self.dropdown.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                    self.dropdown.bind("<<ComboboxSelected>>", self.image_excel)

                def select_screen():
                    try:
                        with open("JSON Files\\screen_location.json", "r") as file:
                            data = json.load(file)
                            print("Excel JSON taken")
                    except FileNotFoundError:
                        print("No file found.")
                        message = "No file found."
                        label = tk.Label(self.L2frame_frm, text=message, font=self.Sfont, width=20)
                        label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
                        return
                    # Dropdown menu
                    names = [item["name"] for item in data]
                    self.dropdown = ttk.Combobox(self.L2frame_frm, values=names, font=self.Sfont, width=30)
                    self.dropdown.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                    self.dropdown.bind("<<ComboboxSelected>>", self.screen_excel)

                def button_clicked(button_name):
                    if button_name == "Image":
                        select_image()
                    elif button_name == "Screen":
                        select_screen()
                    text_label.grid_remove()
                    image_button.grid_remove()
                    screen_button.grid_remove()

                text_label = ttk.Label(self.L2frame_frm, text="Select the Form")
                text_label.grid(row=0, column=0, pady=5)
                image_button = ttk.Button(self.L2frame_frm, text="Image", command=lambda: button_clicked("Image"))
                image_button.grid(row=1, column=0, padx=5, pady=5)
                screen_button = ttk.Button(self.L2frame_frm, text="Screen", command=lambda: button_clicked("Screen"))
                screen_button.grid(row=1, column=1, padx=5, pady=5)


            elif task_type == "Calender Location":
                try:
                    with open("JSON Files\\calendar_location.json", "r") as file:
                        data = json.load(file)
                        print("Calender JSON taken")
                except FileNotFoundError:
                    print("No file found.")
                    message = "No file found."
                    label = tk.Label(self.L2frame_frm, text=message, font=self.Sfont, width=20)
                    label.grid(row=0, column=0, padx=5, pady=5, sticky=W)
                    return
                # Dropdown menu
                names = [item["week"] for item in data]
                self.dropdown = ttk.Combobox(self.L2frame_frm, values=names, font=self.Sfont, width=35)
                self.dropdown.grid(row=1, column=0, padx=5, pady=5, sticky=W)

            elif task_type == "Others":
                # with open("others.json", "r") as file:
                #     data = json.load(file)
                options = ["Switch browser tab", "Switch windows tab", "Open Application", "Start Excel Application", "Create End process"]
                self.dropdown = ttk.Combobox(self.L2frame_frm, values=options, font=self.Sfont, width=35)
                self.dropdown.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                self.dropdown.bind("<<ComboboxSelected>>", self.others_options)




        # Saving others data to JSONS
        def others_json_1(self):
            data = [
                {
                    "act_id": 1,
                    "task_title": self.task_name.get(),
                    "task_type": self.task_option.get(),
                    "task_subtype": "browser_tab",
                    "act_name": self.other_option.get(),
                    "act_type": "Browser",
                    "clicker": [],
                    "calender_matrix": [],
                    "cv": [],
                    "cv_calender": "[]",
                    "act_value": "",
                    "act_input": "",
                    "act_delay": int(self.delay_entry.get()),
                    "act_newtab": self.tab_dropdown.get(),
                    "other_process": "",
                    "move": "nowhere"
                }
            ]
            print("\nJSON DATA", data)
            filename = self.task_name.get() + ".json"
            print("JSON FILE NAME", filename)
            with open(filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("JSON SAVED")

        def others_json_2(self):
            data = [
                {
                    "act_id": 1,
                    "task_title": self.task_name.get(),
                    "task_type": self.task_option.get(),
                    "task_subtype": "windows_tab",
                    "act_name": self.other_option.get(),
                    "act_type": "Computer",
                    "clicker": [],
                    "calender_matrix": [],
                    "cv": [],
                    "cv_calender": "[]",
                    "act_value": "",
                    "act_input": "",
                    "act_delay": int(self.delay_entry.get()),
                    "act_newtab": self.tab_dropdown.get(),
                    "other_process": "",
                    "move": "nowhere"
                }
            ]
            print("\nJSON DATA", data)
            filename = self.task_name.get() + ".json"
            print("JSON FILE NAME", filename)
            with open(filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("JSON SAVED")

        def others_json_3(self):
            data = [
                {
                    "act_id": 1,
                    "task_title": self.task_name.get(),
                    "task_type": self.task_option.get(),
                    "task_subtype": "open",
                    "act_name":  self.other_option.get(),
                    "act_type": "Application",
                    "clicker": [],
                    "calender_matrix": [],
                    "cv": [],
                    "cv_calender": "[]",
                    "act_value": self.name_entry.get(),
                    "act_input": self.url_entry.get(),
                    "act_delay": int(self.delay_entry.get()),
                    "act_newtab": "",
                    "other_process": "",
                    "move": "nowhere"
                }
            ]
            print("\nJSON DATA", data)
            filename = self.task_name.get() + ".json"
            print("JSON FILE NAME", filename)
            with open(filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("JSON SAVED")

        def others_json_4(self):
            data = [
                {
                    "act_id": 1,
                    "task_title": self.task_name.get(),
                    "task_type": self.task_option.get(),
                    "task_subtype": "excel_start",
                    "act_name": self.other_option.get(),
                    "act_type": "Excel",
                    "clicker": [],
                    "calender_matrix": [],
                    "cv": [],
                    "cv_calender": "[]",
                    "act_value": "",
                    "act_input": "",
                    "act_delay": "",
                    "act_newtab": "",
                    "other_process": "",
                    "excel_path": self.path_entry.get(),
                    "excel_search_data": self.data_entry.get(),
                    "excel_search_column": self.col_entry.get(),
                    "move": "nowhere",
                    "is_dynamic":self.checkbox_var.get()
                }
            ]
            print("\nJSON DATA", data)
            filename = self.task_name.get() + ".json"
            print("JSON FILE NAME", filename)
            with open(filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("JSON SAVED")

        def others_json_5(self):
            data = [
                {
                    "act_id": 1,
                    "task_title": self.task_name.get(),
                    "task_type": self.task_option.get(),
                    "task_subtype": "end",
                    "act_name": self.other_option.get(),
                    "act_type": "",
                    "clicker": [],
                    "calender_matrix": [],
                    "cv": [],
                    "cv_calender": "[]",
                    "act_value": self.loop_entry.get(),
                    "act_input": self.img_entry.get(),
                    "act_delay": self.delay_entry.get(),
                    "act_newtab": "",
                    "other_process": "",
                    "move": "nowhere"
                }
            ]
            print("\nJSON DATA", data)
            filename = self.task_name.get() + ".json"
            print("JSON FILE NAME", filename)
            with open(filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("JSON SAVED")

        def others_options(self, event):
            other_option = self.dropdown.get()
            self.dropdown.grid_remove()
            if other_option == "Switch browser tab":
                #Dropdown menu
                self.other_option = Entry(self.L2frame_frm, font=self.Sfont, width=20)
                self.other_option.insert(0, other_option)
                self.other_option.configure(state='readonly')
                self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                # self.other_option = Label(self.L2frame_frm, text=other_option, font=self.Sfont)
                # self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                tab_options = ["Default Tab "] + ["New Tab"] + [str(i) for i in range(1, 10)]      # tab_options = ["Default Tab", "New Tab", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
                self.tab_dropdown = ttk.Combobox(self.L2frame_frm, values=tab_options, font=self.Sfont, width=15)
                self.tab_dropdown.grid(row=2, column=1, padx=5, pady=5, sticky=W)
                # Entry for delay
                self.delay_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.delay_entry.grid(row=3, column=1, padx=5, pady=5, sticky=W)
                # Tab & delay label
                self.selected_label = Label(self.L2frame_frm, text="TASK:", font=self.Sfont)
                self.selected_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                tab_label = Label(self.L2frame_frm, text="TAB", font=self.Sfont)
                tab_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
                delay_label = Label(self.L2frame_frm, text="DELAY", font=self.Sfont)
                delay_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
                # Save button
                save_button = Button(self.L2frame_frm, text="Save", command=self.others_json_1, font=self.Sfont, width=15)
                save_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)
            elif other_option == "Switch windows tab":
                self.other_option = Entry(self.L2frame_frm, font=self.Sfont, width=20)
                self.other_option.insert(0, other_option)
                self.other_option.configure(state='readonly')
                self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                # self.other_option = Label(self.L2frame_frm, text=other_option, font=self.Sfont)
                # self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                tab_options = ["Tab " + str(i) for i in range(1, 10)]
                self.tab_dropdown = ttk.Combobox(self.L2frame_frm, values=tab_options, font=self.Sfont, width=15)
                self.tab_dropdown.grid(row=2, column=1, padx=5, pady=5, sticky=W)
                self.delay_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.delay_entry.grid(row=3, column=1, padx=5, pady=5, sticky=W)
                self.selected_label = Label(self.L2frame_frm, text="TASK:", font=self.Sfont)
                self.selected_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                tab_label = Label(self.L2frame_frm, text="TAB", font=self.Sfont)
                tab_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
                delay_label = Label(self.L2frame_frm, text="DELAY", font=self.Sfont)
                delay_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
                # Save button
                save_button = Button(self.L2frame_frm, text="Save", command=self.others_json_1, font=self.Sfont, width=15)
                save_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)
            elif other_option == "Open Application":
                # Entry for URL DELAY NAME
                self.other_option = Entry(self.L2frame_frm, font=self.Sfont, width=20)
                self.other_option.insert(0, other_option)
                self.other_option.configure(state='readonly')
                self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                # self.other_option = Label(self.L2frame_frm, text=other_option, font=self.Sfont)
                # self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                self.url_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.url_entry.grid(row=2, column=1, padx=5, pady=5, sticky=W)
                self.delay_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.delay_entry.grid(row=3, column=1, padx=5, pady=5, sticky=W)
                self.name_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.name_entry.grid(row=4, column=1, padx=5, pady=5, sticky=W)
                # URL & delay label
                self.selected_label = Label(self.L2frame_frm, text="TASK:", font=self.Sfont)
                self.selected_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                url_label = Label(self.L2frame_frm, text="URL", font=self.Sfont)
                url_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
                delay_label = Label(self.L2frame_frm, text="DELAY", font=self.Sfont)
                delay_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
                name_label = Label(self.L2frame_frm, text="NAME:", font=self.Sfont)
                name_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)
                # Save button
                save_button = Button(self.L2frame_frm, text="Save", command=self.others_json_3, font=self.Sfont, width=15)
                save_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)
            elif other_option == "Start Excel Application":
                # Entry for PATH DELAY NAME
                self.other_option = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.other_option.insert(0, other_option)
                self.other_option.configure(state='readonly')
                self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                # self.other_option = Label(self.L2frame_frm, text=other_option, font=self.Sfont)
                # self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                self.path_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.path_entry.grid(row=2, column=1, padx=5, pady=5, sticky=W)
                self.col_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.col_entry.grid(row=3, column=1, padx=5, pady=5, sticky=W)
                self.data_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.data_entry.grid(row=4, column=1, padx=5, pady=5, sticky=W)
                # PATH label
                self.selected_label = Label(self.L2frame_frm, text="TASK:")
                self.selected_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                path_label = Label(self.L2frame_frm, text="EXCEL PATH")
                path_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
                name_label = Label(self.L2frame_frm, text="EXCEL SEARCH COLUMN NAME")
                name_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
                data_label = Label(self.L2frame_frm, text="EXCEL SEARCH COLUMN DATA:")
                data_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)
                # Check box
                self.checkbox_var = IntVar()
                label = Label(self.L2frame_frm, text="Is Dynamic", font=self.Sfont)
                label.grid(row=5, column=0, padx=5, pady=5, sticky=W)
                checkbox = Checkbutton(self.L2frame_frm, variable=self.checkbox_var)
                checkbox.grid(row=5, column=1, padx=5, pady=5, sticky=W)
                # Save button
                save_button = Button(self.L2frame_frm, text="Save", command=self.others_json_4, font=self.Sfont, width=10)
                save_button.grid(row=6, column=0, columnspan=2, padx=5, pady=5)
            elif other_option == "Create End process":
                # Entry for IMG DELAY NAME
                self.other_option = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.other_option.insert(0, other_option)
                self.other_option.configure(state='readonly')
                self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                # self.other_option = Label(self.L2frame_frm, text=other_option, font=self.Sfont)
                # self.other_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                self.img_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.img_entry.grid(row=2, column=1, padx=5, pady=5, sticky=W)
                self.delay_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.delay_entry.grid(row=3, column=1, padx=5, pady=5, sticky=W)
                self.loop_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.loop_entry.grid(row=4, column=1, padx=5, pady=5, sticky=W)
                # IMG & delay label
                self.selected_label = Label(self.L2frame_frm, text="TASK:", font=self.Sfont)
                self.selected_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                img_label = Label(self.L2frame_frm, text="IMAGE", font=self.Sfont)
                img_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
                delay_label = Label(self.L2frame_frm, text="DELAY", font=self.Sfont)
                delay_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
                loop_label = Label(self.L2frame_frm, text="LOOP", font=self.Sfont)
                loop_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)
                # Save button
                save_button = Button(self.L2frame_frm, text="Save", command=self.others_json_5, font=self.Sfont, width=15)
                save_button.grid(row=5, column=0, columnspan=2, padx=5, pady=5)


        # Saving screen data JSON
        def screen_json(self):
            data = [
                {
                    "act_id": 1,
                    "task_title": self.task_name.get(),
                    "task_type": self.task_option.get(),
                    "task_subtype": "",
                    "act_name": self.screen_option.get(),
                    "act_type": "",
                    "clicker": [
                        {
                            "name": self.task_name.get(),
                            "x": "119",
                            "y": "285",
                            "status": "True",
                            "delay": int(self.delay_entry.get()),
                            "operation": "click",
                            "calender_type": self.event_dropdown.get(),
                            "move": self.move_dropdown.get(),
                            "times": self.time_entry.get(),
                            "wait": self.wait_entry.get(),
                            "other_process": self.proc_entry.get(),
                            "copy": "",
                            "type": "clicker",
                            "Do Not Repeat Column": self.checkbox_var1.get(),
                            "Repeat Column": self.checkbox_var2.get(),
                            "Skip Matching": self.checkbox_var3.get()
                        }
                    ],
                    "calender_matrix": [],
                    "cv": [],
                    "cv_calender": "[]",
                    "act_value": "",
                    "act_input": "",
                    "act_delay": "1",
                    "act_newtab": "",
                    "other_process": "",
                    "move": "nowhere"
                }
            ]
            print("\nJSON DATA", data)
            filename = self.task_name.get() + ".json"
            print("JSON FILE NAME", filename)
            with open(filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("JSON SAVED")

        def screen_options(self, event):
            screen_option = self.dropdown.get()
            self.dropdown.grid_remove()
            if (screen_option == "click1" or screen_option == "click2" or screen_option == "click3"
                    or screen_option == "click4" or screen_option == "click5"):
                self.screen_option = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.screen_option.insert(0, screen_option)
                self.screen_option.configure(state='readonly')
                self.screen_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                # self.screen_option = Label(self.L2frame_frm, text=screen_option, font=self.Sfont)
                # self.screen_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                self.delay_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.delay_entry.grid(row=2, column=1, padx=5, pady=5, sticky=W)
                events = ["Click", "Right Click", "Double Click", "Paste", "Paste & Enter", "Delete & Enter", "Move To",
                          "Left Arrow", "Right Arrow", "Up Arrow", "Down Arrow", "Type & Enter", "Type", "Shift Ctrl Down Arrow",
                          "Type & Tab", "Type & Down Arrow", "Enter & Down Arrow", "Delete", "Filter Function", "Ctrl+T",
                          "Down Arrow & Space Bar", "Check"]
                self.event_dropdown = ttk.Combobox(self.L2frame_frm, values=events, font=self.Sfont, width=15)
                self.event_dropdown.grid(row=3, column=1, padx=5, pady=5, sticky=W)
                Calender = ["None", "Today", "Yesterday", "Day before Yesterday", "Last week", "Last month"]
                self.event_dropdown = ttk.Combobox(self.L2frame_frm, values=Calender, font=self.Sfont, width=15)
                self.event_dropdown.grid(row=4, column=1, padx=5, pady=5, sticky=W)
                MOVE_TO = ["now here", "Previous Task", "Previous Process", "Next Process", "Next Task"]
                self.move_dropdown = ttk.Combobox(self.L2frame_frm, values=MOVE_TO, font=self.Sfont, width=15)
                self.move_dropdown.grid(row=5, column=1, padx=5, pady=5, sticky=W)
                self.time_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.time_entry.grid(row=6, column=1, padx=5, pady=5, sticky=W)
                self.wait_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.wait_entry.grid(row=7, column=1, padx=5, pady=5, sticky=W)
                self.proc_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.proc_entry.grid(row=8, column=1, padx=5, pady=5, sticky=W)
                self.paste_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.paste_entry.grid(row=9, column=1, padx=5, pady=5, sticky=W)

                # labelS
                self.selected_label = Label(self.L2frame_frm, text="TASK:", font=self.Sfont)
                self.selected_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                delay_label = Label(self.L2frame_frm, text="DELAY", font=self.Sfont)
                delay_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
                event_label = Label(self.L2frame_frm, text="EVENT", font=self.Sfont)
                event_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
                cal_label = Label(self.L2frame_frm, text="CALENDER TYPE", font=self.Sfont)
                cal_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)
                move_label = Label(self.L2frame_frm, text="MOVE TO", font=self.Sfont)
                move_label.grid(row=5, column=0, padx=5, pady=5, sticky=W)
                time_label = Label(self.L2frame_frm, text="TIMES", font=self.Sfont)
                time_label.grid(row=6, column=0, padx=5, pady=5, sticky=W)
                wait_label = Label(self.L2frame_frm, text="WAIT", font=self.Sfont)
                wait_label.grid(row=7, column=0, padx=5, pady=5, sticky=W)
                proc_label = Label(self.L2frame_frm, text="OTHER PROCESS", font=self.Sfont)
                proc_label.grid(row=8, column=0, padx=5, pady=5, sticky=W)
                paste_label = Label(self.L2frame_frm, text="PASTE", font=self.Sfont)
                paste_label.grid(row=9, column=0, padx=5, pady=5, sticky=W)

                # Creating the check box
                self.checkbox_var1 = IntVar()
                label = Label(self.L2frame_frm, text="Do Not Repeat Column", font=self.Sfont)
                label.grid(row=10, column=0, padx=5, pady=5, sticky=W)
                checkbox = Checkbutton(self.L2frame_frm, variable=self.checkbox_var1)
                checkbox.grid(row=10, column=1, padx=5, pady=5, sticky=W)
                self.checkbox_var2 = IntVar()
                label = Label(self.L2frame_frm, text="Repeat Column", font=self.Sfont)
                label.grid(row=11, column=0, padx=5, pady=5, sticky=W)
                checkbox = Checkbutton(self.L2frame_frm, variable=self.checkbox_var2)
                checkbox.grid(row=11, column=1, padx=5, pady=5, sticky=W)
                self.checkbox_var3 = IntVar()
                label = Label(self.L2frame_frm, text="Skip Matching", font=self.Sfont)
                label.grid(row=12, column=0, padx=5, pady=5, sticky=W)
                checkbox = Checkbutton(self.L2frame_frm, variable=self.checkbox_var3)
                checkbox.grid(row=12, column=1, padx=5, pady=5, sticky=W)
                # Save button
                save_button = Button(self.L2frame_frm, text="Save", command=self.screen_json, font=self.Sfont, width=15)
                save_button.grid(row=13, column=0, columnspan=2, padx=5, pady=5)


        # Saving the Image data to JSON
        def image_json(self):
            data = [
                {
                    "act_id": 1,
                    "task_title": self.task_name.get(),
                    "task_type": self.task_option.get(),
                    "task_subtype": "",
                    "act_name": self.image_option.get(),
                    "act_type": "",
                    "clicker": [],
                    "calender_matrix": [],
                    "cv": [
                        {
                            "operation": "template matching",
                            "delay": int(self.delay_entry.get()),
                            "loop_delay": str(self.loop_entry.get()),
                            "image": self.image_option.get() + ".jpg",
                            "move": self.move_dropdown.get(),
                            "other_process": self.proc_entry.get(),
                            "wait": self.wait_entry.get()
                        },
                        {
                            "name": self.task_name.get(),
                            "x": "196",
                            "y": "17",
                            "image": self.image_option.get() + ".jpg",
                            # "image": self.image_option + ".jpg",
                            "status": "True",
                            "delay": str(self.delay_entry.get()),
                            "loop_delay": str(self.loop_entry.get()),
                            "operation": self.event_dropdown.get(),
                            "calender_type": self.cal_dropdown.get(),
                            "move": self.move_dropdown.get(),
                            "times": self.time_entry.get(),
                            "wait": self.wait_entry.get(),
                            "other_process": self.proc_entry.get(),
                            "copy": "",
                            "screenshot": self.ss_dropdown.get(),
                            "repeat until": self.checkbox_var1.get()
                        }
                    ],
                    "cv_calender": "[]",
                    "act_value": "",
                    "act_input": "",
                    "act_delay": str(self.delay_entry.get()),
                    "act_newtab": "",
                    "other_process": "",
                    "move": self.move_dropdown.get()
                }
            ]
            print("\nJSON DATA", data)
            filename = self.task_name.get() + ".json"
            print("JSON FILE NAME", filename)
            with open(filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("JSON SAVED")

        def image_options(self, event):
            image_option = self.dropdown.get()
            self.dropdown.grid_remove()
            if (image_option == "open_erp" or image_option == "user_name" or image_option == "pswrd" or
                    image_option == "enter" or image_option == "click_start" or image_option == "search" or
                    image_option == "select_first" or image_option == "prodcode_1" or image_option == "select_first2" or
                    image_option == "bom_component" or image_option == "maximise" or image_option == "select_first3" or
                    image_option == "click_line" or image_option == "search" or image_option == "accessories_serial" or
                    image_option == "click_accessories_serial" or image_option == "prd_code6" or image_option == "select_first4" or
                    image_option == "new" or image_option == "prodcode1" or image_option == "search2" or image_option == "view" or
                    image_option == "click_test" or image_option == "next_line"):

                self.image_option = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.image_option.insert(0, image_option)
                self.image_option.configure(state='readonly')
                self.image_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                # self.image_option = Label(self.L2frame_frm, text=image_option, font=self.Sfont)
                # self.image_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                self.delay_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.delay_entry.grid(row=2, column=1, padx=5, pady=5, sticky=W)
                self.loop_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.loop_entry.grid(row=3, column=1, padx=5, pady=5, sticky=W)
                events = ["Click", "Right Click", "Double Click", "Paste", "Paste & Enter", "Delete & Enter", "Move To",
                          "Left Arrow", "Right Arrow", "Up Arrow", "Down Arrow", "Type & Enter", "Type", "Shift Ctrl Down Arrow",
                          "Type & Tab", "Type & Down Arrow", "Enter & Down Arrow", "Delete", "Filter Function", "Ctrl+T",
                          "Down Arrow & Space Bar", "Check"]
                self.event_dropdown = ttk.Combobox(self.L2frame_frm, values=events, font=self.Sfont, width=15)
                self.event_dropdown.grid(row=4, column=1, padx=5, pady=5, sticky=W)
                Calender = ["None", "Today", "Yesterday", "Day before Yesterday", "Last week", "Last month"]
                self.cal_dropdown = ttk.Combobox(self.L2frame_frm, values=Calender, font=self.Sfont, width=15)
                self.cal_dropdown.grid(row=5, column=1, padx=5, pady=5, sticky=W)
                MOVE_TO = ["now here", "Previous Task", "Previous Process", "Next Process", "Next Task"]
                self.move_dropdown = ttk.Combobox(self.L2frame_frm, values=MOVE_TO, font=self.Sfont, width=15)
                self.move_dropdown.grid(row=6, column=1, padx=5, pady=5, sticky=W)
                self.time_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.time_entry.grid(row=7, column=1, padx=5, pady=5, sticky=W)
                self.wait_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.wait_entry.grid(row=8, column=1, padx=5, pady=5, sticky=W)
                self.proc_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.proc_entry.grid(row=9, column=1, padx=5, pady=5, sticky=W)
                self.paste_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.paste_entry.grid(row=10, column=1, padx=5, pady=5, sticky=W)
                ss = ["Once", "Multiple"]
                self.ss_dropdown = ttk.Combobox(self.L2frame_frm, values=ss, font=self.Sfont, width=15)
                self.ss_dropdown.grid(row=11, column=1, padx=5, pady=5, sticky=W)

                # labelS
                self.selected_label = Label(self.L2frame_frm, text="TASK:", font=self.Sfont)
                self.selected_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                delay_label = Label(self.L2frame_frm, text="DELAY", font=self.Sfont)
                delay_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
                loop_label = Label(self.L2frame_frm, text="LOOP", font=self.Sfont)
                loop_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
                event_label = Label(self.L2frame_frm, text="EVENT", font=self.Sfont)
                event_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)
                cal_label = Label(self.L2frame_frm, text="CALENDER TYPE", font=self.Sfont)
                cal_label.grid(row=5, column=0, padx=5, pady=5, sticky=W)
                move_label = Label(self.L2frame_frm, text="MOVE TO", font=self.Sfont)
                move_label.grid(row=6, column=0, padx=5, pady=5, sticky=W)
                time_label = Label(self.L2frame_frm, text="TIMES", font=self.Sfont)
                time_label.grid(row=7, column=0, padx=5, pady=5, sticky=W)
                wait_label = Label(self.L2frame_frm, text="WAIT", font=self.Sfont)
                wait_label.grid(row=8, column=0, padx=5, pady=5, sticky=W)
                proc_label = Label(self.L2frame_frm, text="OTHER PROCESS", font=self.Sfont)
                proc_label.grid(row=9, column=0, padx=5, pady=5, sticky=W)
                paste_label = Label(self.L2frame_frm, text="PASTE", font=self.Sfont)
                paste_label.grid(row=10, column=0, padx=5, pady=5, sticky=W)
                ss_label = Label(self.L2frame_frm, text="SCREENSHOT", font=self.Sfont)
                ss_label.grid(row=11, column=0, padx=5, pady=5, sticky=W)
                # Check box
                label = Label(self.L2frame_frm, text="REPEAT UNTIL", font=self.Sfont)
                label.grid(row=12, column=0, padx=5, pady=5, sticky=W)
                self.checkbox_var1 = IntVar()
                checkbox = Checkbutton(self.L2frame_frm, variable=self.checkbox_var1)
                checkbox.grid(row=12, column=1, padx=5, pady=5, sticky=W)
                # Save button
                save_button = Button(self.L2frame_frm, text="Save", command=self.image_json, font=self.Sfont, width=15)
                save_button.grid(row=13, column=0, columnspan=2, padx=5, pady=5)


        def excel_screen_json(self):
            data = [
                {
                    "act_id": 1,
                    "task_title": self.task_name.get(),
                    "task_type": self.task_option.get(),
                    "task_subtype": "",
                    "act_name": self.screen_option.get(),
                    "act_type": "",
                    "clicker": [
                        {
                            "name": self.task_name.get(),
                            "x": "119",
                            "y": "285",
                            "status": "True",
                            "delay": str(self.delay_entry.get()),
                            "operation":  self.event_dropdown.get(),
                            "move": self.move_dropdown.get(),
                            "times": self.time_entry.get(),
                            "wait": self.wait_entry.get(),
                            "other_process": self.process_entry.get(),
                            "copy": self.paste_entry.get(),
                            "type": "excel",
                            "column": "",
                            "image": "",
                            "loop_delay": "",
                            "excel_path": ""
                        }
                    ],
                    "calender_matrix": [],
                    "cv": [],
                    "cv_calender": "[]",
                    "act_value": "",
                    "act_input": "",
                    "act_delay": "",
                    "act_newtab": "",
                    "other_process": self.process_entry.get(),
                    "excel_path": "",
                    "move": self.move_dropdown.get(),
                }
            ]
            print("\nJSON DATA", data)
            filename = self.task_name.get() + ".json"
            print("JSON FILE NAME", filename)
            with open(filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("JSON SAVED")

        def screen_excel(self, event):
            screen_excel = self.dropdown.get()
            self.dropdown.grid_remove()
            if (screen_excel == "click1" or screen_excel == "click2" or screen_excel == "click3"
                    or screen_excel == "click4" or screen_excel == "click5"):
                self.screen_option = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.screen_option.insert(0, screen_excel)
                self.screen_option.configure(state='readonly')
                self.screen_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                # self.image_option = Label(self.L2frame_frm, text=image_option, font=self.Sfont)
                # self.image_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                self.selected_label = Label(self.L2frame_frm, text="TASK:", font=self.Sfont)
                self.selected_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                self.delay_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.delay_entry.grid(row=2, column=1, padx=5, pady=5, sticky=W)
                delay_label = Label(self.L2frame_frm, text="DELAY", font=self.Sfont)
                delay_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
                events = ["Click", "Double Click", "Paste", "Delete & Paste", "Enter", "Move To", "Left Arrow",
                          "Right Arrow", "Up Arrow", "Down Arrow", "Type & Enter", "Type", "ShiftCtrl DownArrow",
                          "Tab Enter", "Type & Tab", "Type & DownArrow",  "Enter & DownArrow",  "Delete",  "Filter Function",
                          "Ctrl+T",  "DownArrow & SpaceBar", "Check" ]
                self.event_dropdown = ttk.Combobox(self.L2frame_frm, values=events, font=self.Sfont, width=15)
                self.event_dropdown.grid(row=3, column=1, padx=5, pady=5, sticky=W)
                event_label = Label(self.L2frame_frm, text="EVENT", font=self.Sfont)
                event_label.grid(row=3, column=0, padx=5, pady=5, sticky=W)
                Calender = ["None", "Today", "Yesterday", "Day before Yesterday", "Last week", "Last month"]
                self.cal_dropdown = ttk.Combobox(self.L2frame_frm, values=Calender, font=self.Sfont, width=15)
                self.cal_dropdown.grid(row=4, column=1, padx=5, pady=5, sticky=W)
                cal_label = Label(self.L2frame_frm, text="CALENDER TYPE", font=self.Sfont)
                cal_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)
                move_option = ["now here", "Previous Task", "Previous Process", "Next Process", "Next Task"]
                self.move_dropdown = ttk.Combobox(self.L2frame_frm, values=move_option, font=self.Sfont, width=15)
                self.move_dropdown.grid(row=5, column=1, padx=5, pady=5, sticky=W)
                move_label = Label(self.L2frame_frm, text="MOVE TO", font=self.Sfont)
                move_label.grid(row=5, column=0, padx=5, pady=5, sticky=W)
                self.time_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.time_entry.grid(row=6, column=1, padx=5, pady=5, sticky=W)
                time_label = Label(self.L2frame_frm, text="TIMES", font=self.Sfont)
                time_label.grid(row=6, column=0, padx=5, pady=5, sticky=W)
                self.process_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.process_entry.grid(row=7, column=1, padx=5, pady=5, sticky=W)
                process_label = Label(self.L2frame_frm, text="OTHER PROCESS", font=self.Sfont)
                process_label.grid(row=7, column=0, padx=5, pady=5, sticky=W)
                self.paste_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.paste_entry.grid(row=8, column=1, padx=5, pady=5, sticky=W)
                paste_label = Label(self.L2frame_frm, text="PASTE", font=self.Sfont)
                paste_label.grid(row=8, column=0, padx=5, pady=5, sticky=W)
                # Save button
                save_button = Button(self.L2frame_frm, text="Save", command=self.excel_screen_json, font=self.Sfont, width=15)
                save_button.grid(row=9, column=0, columnspan=2, padx=5, pady=5)


        def excel_image_json(self):
            data = [
                {
                    "act_id": 1,
                    "task_title": self.task_name.get(),
                    "task_type": self.task_option.get(),
                    "task_subtype": "",
                    "act_name": "Click",
                    "act_type": self.excel_option.get(),
                    "clicker": [],
                    "calender_matrix": [],
                    "cv": [
                        {
                            "operation": "templatematching",
                            # "delay": self.wait_entry.get(),
                            "delay": "",
                            # "loop_delay": int(self.delay_entry.get()),
                            "loop_delay": "",
                            "image": self.excel_option.get() + ".jpg",
                            "move": self.move_dropdown.get(),
                            "other_process": "",
                            "wait": self.wait_entry.get(),
                            "excel_path": self.excel_path_entry.get(),
                            "compare": self.compare_col_entry.get(),
                            "type": "excel",
                            "check": self.checkbox_var1.get(),
                            "repeat": self.checkbox_var2.get(),
                            "do_not_repeat": self.checkbox_var3.get(),
                            "skip_matching": self.checkbox_var4.get()
                        },
                        {
                            "name": self.task_name.get(),
                            "x": "196",
                            "y": "17",
                            "image": self.excel_option.get() + ".jpg",
                            "status": "True",
                            "copy": self.paste_entry.get(),
                            "column": self.col_name_entry.get(),
                            "delay": int(self.delay_entry.get()),
                            "loop_delay": self.loop_entry.get(),
                            "operation": self.event_dropdown.get(),
                            "move": self.move_dropdown.get(),
                            "times": self.time_entry.get(),
                            "wait": self.wait_entry.get(),
                            "excel_path": self.excel_path_entry.get(),
                            "compare": self.compare_col_entry.get(),
                            "type": "excel",
                            "amount_check": self.amt_entry.get(),
                            "check": self.checkbox_var1.get(),
                            "repeat": self.checkbox_var2.get(),
                            "do_not_repeat": self.checkbox_var3.get(),
                            "0": [],
                            "1": [],
                            "calender_type": self.cal_dropdown.get(),
                            "skip_matching": self.checkbox_var4.get(),
                            "path": self.app_path_entry.get()
                        }
                    ],
                    "cv_calender": "[]",
                    "act_value": "",
                    "act_input": "",
                    "act_delay": "",
                    "act_newtab": "",
                    "other_process": "",
                    "excel_path": "",
                    "move": self.move_dropdown.get()
                }
            ]
            print("\nJSON DATA", data)
            filename = self.task_name.get() + ".json"
            print("JSON FILE NAME", filename)
            with open(filename, "w") as json_file:
                json.dump(data, json_file, indent=4)
            print("JSON SAVED")

        def image_excel(self, event):
            excel_option = self.dropdown.get()
            self.dropdown.grid_remove()
            if (excel_option == "open_erp" or excel_option == "user_name" or excel_option == "pswrd" or
                    excel_option == "enter" or excel_option == "click_start" or excel_option == "search" or
                    excel_option == "select_first" or excel_option == "prodcode_1" or excel_option == "select_first2" or
                    excel_option == "bom_component" or excel_option == "maximise" or excel_option == "select_first3" or
                    excel_option == "click_line" or excel_option == "search" or excel_option == "accessories_serial" or
                    excel_option == "click_accessories_serial" or excel_option == "prd_code6" or excel_option == "select_first4" or
                    excel_option == "new" or excel_option == "prodcode1" or excel_option == "search2" or excel_option == "view" or
                    excel_option == "click_test" or excel_option == "next_line"):
                self.excel_option = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.excel_option.insert(0, excel_option)
                self.excel_option.configure(state='readonly')
                self.excel_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                # self.image_option = Label(self.L2frame_frm, text=image_option, font=self.Sfont)
                # self.image_option.grid(row=1, column=1, padx=5, pady=5, sticky=W)
                self.selected_label = Label(self.L2frame_frm, text="TASK:", font=self.Sfont)
                self.selected_label.grid(row=1, column=0, padx=5, pady=5, sticky=W)
                self.col_name_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.col_name_entry.grid(row=2, column=1, padx=5, pady=5, sticky=W)
                col_name_label = Label(self.L2frame_frm, text="Column Name", font=self.Sfont)
                col_name_label.grid(row=2, column=0, padx=5, pady=5, sticky=W)
                self.excel_path_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.excel_path_entry.grid(row=3, column=1, padx=5, pady=5, sticky=W)
                excel_path = Label(self.L2frame_frm, text="EXCEL PATH", font=self.Sfont)
                excel_path.grid(row=3, column=0, padx=5, pady=5, sticky=W)
                self.delay_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.delay_entry.grid(row=4, column=1, padx=5, pady=5, sticky=W)
                delay_label = Label(self.L2frame_frm, text="DELAY", font=self.Sfont)
                delay_label.grid(row=4, column=0, padx=5, pady=5, sticky=W)
                self.app_path_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.app_path_entry.grid(row=5, column=1, padx=5, pady=5, sticky=W)
                app_path = Label(self.L2frame_frm, text="Open application path", font=self.Sfont)
                app_path.grid(row=5, column=0, padx=5, pady=5, sticky=W)
                self.compare_col_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.compare_col_entry.grid(row=6, column=1, padx=5, pady=5, sticky=W)
                compare_col = Label(self.L2frame_frm, text="Compare column", font=self.Sfont)
                compare_col.grid(row=6, column=0, padx=5, pady=5, sticky=W)
                self.loop_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.loop_entry.grid(row=7, column=1, padx=5, pady=5, sticky=W)
                loop_label = Label(self.L2frame_frm, text="LOOP", font=self.Sfont)
                loop_label.grid(row=7, column=0, padx=5, pady=5, sticky=W)
                events = ["Paste", "Click", "Double Click", "Delete & Paste", "Enter", "Move To", "Left Arrow",
                          "Right Arrow", "Up Arrow", "Down Arrow", "CtrlA & Copy",  "CtrlA & Paste", "SelectAll", "Sap Post",
                          "SelectAll & Copy", "Tab", "ShiftTab", "Paste & Enter", "Type & Enter", "Type", "Right Click", "ShiftCtrl DownArrow",
                          "Tab Enter", "Type & Tab", "Type & DownArrow", "Enter & DownArrow", "Delete", "Filter Function", "Ctrl+T",
                          "DownArrow & SpaceBar", "Open Application", "Check"]
                self.event_dropdown = ttk.Combobox(self.L2frame_frm, values=events, font=self.Sfont, width=15)
                self.event_dropdown.grid(row=8, column=1, padx=5, pady=5, sticky=W)
                event_label = Label(self.L2frame_frm, text="EVENT", font=self.Sfont)
                event_label.grid(row=8, column=0, padx=5, pady=5, sticky=W)
                Calender = ["None", "Today", "Yesterday", "Day before Yesterday", "Last week", "Last month"]
                self.cal_dropdown = ttk.Combobox(self.L2frame_frm, values=Calender, font=self.Sfont, width=15)
                self.cal_dropdown.grid(row=9, column=1, padx=5, pady=5, sticky=W)
                cal_label = Label(self.L2frame_frm, text="CALENDER TYPE", font=self.Sfont)
                cal_label.grid(row=9, column=0, padx=5, pady=5, sticky=W)
                self.time_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.time_entry.grid(row=10, column=1, padx=5, pady=5, sticky=W)
                time_label = Label(self.L2frame_frm, text="TIMES", font=self.Sfont)
                time_label.grid(row=10, column=0, padx=5, pady=5, sticky=W)
                self.wait_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.wait_entry.grid(row=11, column=1, padx=5, pady=5, sticky=W)
                wait_label = Label(self.L2frame_frm, text="WAIT", font=self.Sfont)
                wait_label.grid(row=11, column=0, padx=5, pady=5, sticky=W)
                move_option = ["now here", "Previous Task", "Previous Process", "Next Process", "Next Task"]
                self.move_dropdown = ttk.Combobox(self.L2frame_frm, values=move_option, font=self.Sfont, width=15)
                self.move_dropdown.grid(row=12, column=1, padx=5, pady=5, sticky=W)
                move_label = Label(self.L2frame_frm, text="MOVE TO", font=self.Sfont)
                move_label.grid(row=12, column=0, padx=5, pady=5, sticky=W)
                self.paste_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.paste_entry.grid(row=13, column=1, padx=5, pady=5, sticky=W)
                paste_label = Label(self.L2frame_frm, text="PASTE", font=self.Sfont)
                paste_label.grid(row=13, column=0, padx=5, pady=5, sticky=W)
                self.amt_entry = Entry(self.L2frame_frm, font=self.Sfont, width=15)
                self.amt_entry.grid(row=14, column=1, padx=5, pady=5, sticky=W)
                amt_label = Label(self.L2frame_frm, text="CHECK AMOUNT", font=self.Sfont)
                amt_label.grid(row=14, column=0, padx=5, pady=5, sticky=W)
                # Check box
                self.checkbox_var1 = IntVar()
                label = Label(self.L2frame_frm, text="CHECK COLUMN", font=self.Sfont)
                label.grid(row=15, column=0, padx=5, pady=5, sticky=W)
                checkbox = Checkbutton(self.L2frame_frm, variable=self.checkbox_var1)
                checkbox.grid(row=15, column=1, padx=5, pady=5, sticky=W)
                self.checkbox_var2 = IntVar()
                label = Label(self.L2frame_frm, text="REPEAT COLUMN", font=self.Sfont)
                label.grid(row=16, column=0, padx=5, pady=5, sticky=W)
                checkbox = Checkbutton(self.L2frame_frm, variable=self.checkbox_var2)
                checkbox.grid(row=16, column=1, padx=5, pady=5, sticky=W)
                self.checkbox_var3 = IntVar()
                label = Label(self.L2frame_frm, text="Do Not Repeat Column", font=self.Sfont)
                label.grid(row=17, column=0, padx=5, pady=5, sticky=W)
                checkbox = Checkbutton(self.L2frame_frm, variable=self.checkbox_var3)
                checkbox.grid(row=17, column=1, padx=5, pady=5, sticky=W)
                self.checkbox_var4 = IntVar()
                label = Label(self.L2frame_frm, text="SKIP MATCHING", font=self.Sfont)
                label.grid(row=18, column=0, padx=5, pady=5, sticky=W)
                checkbox = Checkbutton(self.L2frame_frm, variable=self.checkbox_var4)
                checkbox.grid(row=18, column=1, padx=5, pady=5, sticky=W)
                # Save button
                save_button = Button(self.L2frame_frm, text="Save", command=self.excel_image_json, font=self.Sfont, width=15)
                save_button.grid(row=19, column=0, columnspan=2, padx=5, pady=5)






            else:
                data = []
                # Dropdown menu
                names = [item["name"] for item in data]
                self.dropdown = ttk.Combobox(self.L2frame_frm, values=names, font=self.Sfont, width=30)
                self.dropdown.grid(row=1, column=0, padx=5, pady=5, sticky=W)


            self.updateScrollRegion(self.L1frame_cnv, self.L1frame_frm)
            self.updateScrollRegion(self.Rframe_cnv, self.Rframe_frm)
            self.updateScrollRegion(self.L2frame_cnv, self.L2frame_frm)












        def run_gc(self, text='general'):
            collected = gc.collect()
            if (collected):
                print("CLEANED > " + text + " > %d objects." % (collected))

        def appRunning(self):
            c = wmi.WMI()
            already_running = 0
            appName = str(ver.appName).lower().replace(" ", "")
            for process in c.Win32_Process():
                ppname = str(process.Name).lower().replace(" ", "")
                # print(ppname)
                if (ppname == appName):
                    already_running += 1
            print("RUNNING INSTANCE >> ", already_running)
            return already_running

        def destroy_me(self):
            print("destroy_me")
            global window, treadLoop
            answer = askyesno(title='Quit ' + ver.title, message='Are you sure that you want to quit?')
            if (answer):
                try:
                    treadLoop.cancel()
                except:
                    pp = 0
                current_system_pid = os.getpid()
                print("current_system_pid", current_system_pid)
                ThisSystem = psutil.Process(current_system_pid)
                ThisSystem.terminate()
                window.destroy()

        def error_log(self, ev, page_type=""):
            if (ev):
                exception_type, exception_object, exception_traceback = sys.exc_info()
                exception_type = str(exception_type).replace("<", "").replace(">", "")
                filename = exception_traceback.tb_frame.f_code.co_filename
                only_filename = filename.split("\\")[-1].replace(".py", "")
                line_number = exception_traceback.tb_lineno
                ev = str(ev).replace(".py", ' (FILE) ')
                time_ticker = str(datetime.now().strftime("%d/%m/%Y %H:%M:%S")).upper()
                print(traceback.format_exc())
                error_text = ""
                print("*" * 100)

                trace_data = str(traceback.format_exc()).splitlines()
                trace_error = []
                exclude_files = ["generic", "pool", "multiprocess"]
                for td in trace_data:
                    if ", line " in td:
                        try:
                            error_val = str(td.replace('"', "").split('\\')[-1]).replace(".py", "")
                            start_val = error_val.split(",")[0]
                            if start_val not in exclude_files:
                                std = "   " + error_val.replace(",", " |").replace(" line ", " ").replace(" in ", " ")
                                trace_error.append(std)
                        except:
                            pp = 0

                trace_str = "\n".join(trace_error)
                if (trace_str):
                    trace_str = "\n" + trace_str
                # jsonstr1 = json.dumps(exception_traceback.__dict__)
                # print(trace_data)
                # print(exception_traceback)
                print("+++  PAGE : ", page_type)
                print("+++  TIME : ", time_ticker)
                print("+++  TYPE : ", exception_type)
                print("+++  FILE : ", only_filename)
                print("+++  LINE : ", line_number)
                print("+++ ERROR :", ev)
                print("+++ TRACE : ", trace_str)
                print("*" * 100)
                error_text += "*****************************************"
                error_text += "\n  PAGE : " + str(page_type)
                error_text += "\n  TIME : " + str(time_ticker)
                error_text += "\n  TYPE : " + str(exception_type)
                error_text += "\n  FILE : " + str(only_filename)
                error_text += "\n  LINE : " + str(line_number)
                error_text += "\n ERROR : " + str(ev)
                error_text += "\n TRACE : " + trace_str
                error_text += "\n"
                error_text += "*****************************************"
                error_text += "\n"

                f = open(self.error_txt_file, "a")
                f.write(error_text)
                f.close()

        def updateScrollRegion(self, cnv, frm):
            # print("updateScrollRegion TRIGGERED")
            cnv.update_idletasks()
            cnv.config(scrollregion=frm.bbox())


    root = tk.Tk()
    style = ttk.Style()
    clss = Tk_int(root, style)
    root.mainloop()
except Exception as e:
    page_type = "SUPER"
    if (e):
        exception_type, exception_object, exception_traceback = sys.exc_info()
        exception_type = str(exception_type).replace("<", "").replace(">", "")
        filename = exception_traceback.tb_frame.f_code.co_filename
        only_filename = filename.split("\\")[-1].replace(".py", "")
        line_number = exception_traceback.tb_lineno
        ev = str(e).replace(".py", ' (FILE) ')
        time_ticker = str(datetime.now().strftime("%d/%m/%Y %H:%M:%S")).upper()
        print(traceback.format_exc())
        error_text = ""
        print("*" * 100)

        trace_data = str(traceback.format_exc()).splitlines()
        trace_error = []
        exclude_files = ["generic", "pool", "multiprocess"]
        for td in trace_data:
            if ", line " in td:
                try:
                    error_val = str(td.replace('"', "").split('\\')[-1]).replace(".py", "")
                    start_val = error_val.split(",")[0]
                    if start_val not in exclude_files:
                        std = "   " + error_val.replace(",", " |").replace(" line ", " ").replace(" in ", " ")
                        trace_error.append(std)
                except:
                    pp = 0

        trace_str = "\n".join(trace_error)
        if (trace_str):
            trace_str = "\n" + trace_str
        # jsonstr1 = json.dumps(exception_traceback.__dict__)
        # print(trace_data)
        # print(exception_traceback)
        print("+++  PAGE : ", page_type)
        print("+++  TIME : ", time_ticker)
        print("+++  TYPE : ", exception_type)
        print("+++  FILE : ", only_filename)
        print("+++  LINE : ", line_number)
        print("+++ ERROR :", ev)
        print("+++ TRACE : ", trace_str)
        print("*" * 100)
        error_text += "*****************************************"
        error_text += "\n  PAGE : " + str(page_type)
        error_text += "\n  TIME : " + str(time_ticker)
        error_text += "\n  TYPE : " + str(exception_type)
        error_text += "\n  FILE : " + str(only_filename)
        error_text += "\n  LINE : " + str(line_number)
        error_text += "\n ERROR : " + str(ev)
        error_text += "\n TRACE : " + trace_str
        error_text += "\n"
        error_text += "*****************************************"
        error_text += "\n"

        f = open("superError.txt", "w")
        f.write(error_text)
        f.close()