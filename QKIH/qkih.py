import random
from tkinter.ttk import *
from tkinter import *
from win32gui import *
import pyautogui as pg
import tkinter as tk
from tkinter import ttk
import sv_ttk
from ttkthemes import ThemedTk
import ctypes
import json
import pywinauto as pw
from pywinauto import Desktop
from pywinauto import keyboard
import re
import time

kb = pw.keyboard

windows = Desktop(backend="uia").windows()
myappid = 'mycompany.myproduct.subproduct.version'
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

window = ThemedTk(theme='winxpblue')
window.title('QKIH')
window.geometry("900x500")

window.iconphoto(False, PhotoImage(file='assets/qkih.ico'))
window.configure(background = "black")
window.resizable(width=False, height=False)

bg = PhotoImage(file="assets/qkih.png")
label1 = Label(window, image=bg)
label1.place(x=450, y=-80)
label1.configure(background="black")


class QKIH(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        sv_ttk.use_dark_theme()
        sv_ttk.set_theme("dark")

        self.in_perfix = False
        self.minms = 0
        self.maxms = 0
        self.option_var = StringVar(window)
        QKIH.define_accespoints(self)
        QKIH.load_last_edit(self)
        window.protocol("WM_DELETE_WINDOW", lambda:QKIH.save_last_edit(self))

    def load_last_edit(self):
        try:
            with open("storage\last_script.json", "r") as f:
                last_edit = json.load(f)
                self.script_entry.insert(INSERT, last_edit["script"])
                self.scale_min.set(last_edit["min_ms"])
                self.scale_max.set(last_edit["max_ms"])
                self.key_delay_label_max.configure(
                    text=f"max(ms):"+str(last_edit["max_ms"]))
                self.key_delay_label_min.configure(
                    text=f"min(ms):"+str(last_edit["min_ms"]))
                if last_edit["looped"] == True:
                    self.repeat_button.state(["alternate"])
                else:
                    self.repeat_button.state(["!alternate"])
        except:
            print("last edit was corupted")

    def define_accespoints(self):
        # Access point selector.
        self.title_label = Label(window, text="Quick Keystroke Injection Handler", font=("Arial", 20), bg="black", fg="white")
        self.title_label.place(x=30, y=10)

        self.title_label = Label(window, text="Current entry point:", font=(
            "Arial", 12), bg="black", fg="gray")
        self.title_label.place(x=30, y=60)
        
        # update tabs button
        self.update_tabs_button = ttk.Button(
            window, text="🔃", command=self.update_tabs)
        self.update_tabs_button.place(x=165, y=57, height=30)

        self.update_tabs()

        self.title_label = Label(window, text="Script:", font=(
            "Arial", 12), bg="black", fg="gray")
        self.title_label.place(x=30, y=100)

        # make text area with background image
        self.script_entry = Text(window)
        self.script_entry.configure(background="black", fg="cyan")
        self.script_entry.place(x=30, y=130, width=600, height=300, bordermode=OUTSIDE)

        self.inject_button = ttk.Button(window, text="Quick inject", command=self.inject_script)
        self.inject_button.place(x=30, y=450, width=150, height=30)

        # save to json
        self.save_button = ttk.Button(window, text="Save New", command=self.save_script)
        self.save_button.place(x=200, y=450, width=150, height=30)
        
        # loop toggle button
        self.repeat_button = ttk.Checkbutton(
            window, text="Loop", command=lambda: self.update_loop(self.repeat_button.state()))
        self.repeat_button.place(x=564, y=399, height=30)

        # key delay slider
        self.scale_max = ttk.Scale(
            window, from_=0, to=1000, orient=HORIZONTAL, length=100, command=lambda x: self.update_ms(x, "max"))
        self.scale_max.place(x=464, y=399, height=30)

        # max key delay label text allign right to left
        self.key_delay_label_max = ttk.Label(window, text="max(ms):0", font=(
            "Arial", 12))
        self.key_delay_label_max.place(x=391, y=399, height=30)

        # min key delay slider
        self.scale_min = ttk.Scale(
            window, from_=0, to=1000, orient=HORIZONTAL, length=100, command=lambda x: self.update_ms(x, "min"))
        self.scale_min.place(x=284, y=399, height=30)

        # min key delay label text allign right to left
        self.key_delay_label_min = ttk.Label(window, text="min(ms):0", font=(
            "Arial", 12))
        self.key_delay_label_min.place(x=215, y=399, height=30)

        self.script_entry.bind('<KeyPress>', self.update_colors)
    
    def update_loop(self, state):
        if len(state) == 3:
            self.loop = False
        elif len(state) == 4:
            self.loop = True
        self.loop = state
    
    def update_colors(self, event):
        if event.char == "<":
            self.in_perfix = True
        elif event.char == ">":
            self.in_perfix = False
        if self.in_perfix:
            self.script_entry.tag_add("prefix", "insert linestart", "insert lineend")
            self.script_entry.tag_config("prefix", foreground="red")
        else:
            self.script_entry.tag_remove("prefix", "insert linestart", "insert lineend")
            self.script_entry.tag_config("prefix", foreground="cyan")
        
        # highlight all symbol characters
        self.script_entry.tag_add("symbol", "insert linestart", "insert lineend")
        self.script_entry.tag_config("symbol", foreground="green")
        self.script_entry.tag_remove("symbol", "insert linestart", "insert lineend-1c")
        self.script_entry.tag_remove("symbol", "insert linestart+1c", "insert lineend")






    def update_ms(self, x, type):
        if type == "min":
            self.key_delay_label_min.configure(text=f"{type}(ms):"+str(round(float(x))))
            self.minms = round(float(x))
            if self.minms > self.maxms:
                self.maxms = self.minms
                self.key_delay_label_max.configure(text=f"max(ms):"+str(self.maxms))
                self.scale_max.set(self.maxms)
        else:
            self.key_delay_label_max.configure(text=f"{type}(ms):"+str(round(float(x))))
            self.maxms = round(float(x))
            if self.maxms < self.minms:
                self.minms = self.maxms
                self.key_delay_label_min.configure(text=f"min(ms):"+str(self.minms))
                self.scale_min.set(self.minms)

    def update_tabs(self):
        try:
            self.ap_selector.destroy()
        except:
            pass

        self.open_tabs = []
        for w in windows:
            if w.window_text() != "":
                self.open_tabs.append(w.window_text())
        self.ap_selector = ttk.OptionMenu(
            window, self.option_var, 'no entry point set', *self.open_tabs)
        self.ap_selector.place(x=210, y=57, height=30)
        self.ap_selector.bind("<<ComboboxSelected>>",lambda: print('a'))
        return self.open_tabs

    def refresh(self):
        self.destroy()
        self.__init__()

    def save_script(self):
        # get script
        script = self.script_entry.get("1.0", "end-1c")

        # get access point
        ap = self.option_var.get()

        # add to json
        with open('storage\\scripts.json', 'w') as f:
            json.dump({
                ap: {
                    "id": random.randint(500, 9999999999999),
                    "meanth_entry": ap,
                    "min_ms": self.minms,
                    "max_ms": self.maxms,
                    "looped": self.repeat_button.instate(['selected']),
                    "script": script
                }
            }, f, indent=4)
            
    def save_last_edit(self):
        # get script
        script = self.script_entry.get("1.0", "end-1c")

        # get access point
        ap = self.option_var.get()

        # add to json
        with open('storage\\last_script.json', 'w') as f:
            json.dump(
            {
                    "id": random.randint(500, 9999999999999),
                    "meanth_entry": ap,
                    "min_ms": self.minms,
                    "max_ms": self.maxms,
                    "looped": self.repeat_button.instate(['selected']),
                    "script": script
                
            }, f, indent=4)
            quit('saved last edit')

    def inject_script(self):
        # get script
        script = self.script_entry.get("1.0", "end-1c")
        # get access point
        ap = self.option_var.get()
        # inject script
        self.slept = False
        for w in windows:
            if w.window_text() == ap:
                w.set_focus()
                w.set_focus()
                key_num = -1
                keys_to_skip = 0
                key_delay = random.randint(self.minms, self.maxms)/1000
                print("key delay: "+str(self.minms)+"/"+str(self.maxms))
                for key in script:
                    if keys_to_skip > 0:
                        keys_to_skip -= 1
                        key_num += 1
                        continue
                    if self.slept == False:
                        time.sleep(key_delay)
                    key_num += 1
                    print("|"+key+"| -- "+str(key_num))
                    match key:
                        case " ":
                            kb.send_keys("{SPACE}")
                        case "<":
                            end_arrow_pos = script[key_num:].find(">")
                            raw_function = script[key_num:end_arrow_pos+key_num+1]
                            print("start: "+str(key_num)+" end: "+str(end_arrow_pos+key_num+1)+" raw: "+raw_function+" len: "+str(len(raw_function)))
                            key_function = re.findall(
                                r'\<[a-z-0-9_\s]*\>', raw_function)[0]
                            keys_to_skip = len(key_function)-1
                            sleep_time = 0
                            repeat_time = 1
                            stroke_input = None
                            if "sleep" in key_function:
                                sleep_time = float(re.findall(r'\d+', key_function)[0])
                                key_function = "<sleep>"
                            elif "click" in key_function:
                                xpos = int(re.findall(r'\d+', key_function)[0])
                                ypos = int(re.findall(r'\d+', key_function)[1])
                                click_side = key_function[1]
                                key_function = f"<{click_side}click>"
                            elif re.findall(r'\d+', key_function):
                                repeat_time = int(re.findall(r'\d+', key_function)[0])
                                key_function = re.findall(r'\<.*\d+', key_function)[0]
                                key_function = key_function[:-1]+">"
                            elif "down" in key_function or "up" in key_function:
                                manual = ['ctrl', 'shift', 'alt', 'win']
                                driven = ['VK_CONTROL', 'VK_SHIFT', 'VK_MENU', 'VK_LWIN']

                                for manual_key, driven_key in zip(manual, driven):
                                    if manual_key in key_function:
                                        key_function = key_function.replace(manual_key, driven_key)

                                stroke_input = "{"+key_function.replace(
                                    "<", "").replace(">", "").replace("_u", " u").replace("_d", " d")+"}"
                                print("stroke_input:",
                                    stroke_input, "key_function", key_function)
                                kb.send_keys(stroke_input)
                                

                            for k in range(repeat_time):
                                print("|"+str(key_function).lower()+"|")
                                match str(key_function).lower():
                                    case "<enter>":
                                        kb.send_keys("{ENTER}")
                                    case "<tab>":
                                        kb.send_keys("{TAB}")
                                    case "<backspace>":
                                        kb.send_keys("{BACKSPACE}")
                                    case "<space>":
                                        kb.send_keys("{SPACE}")
                                    case "<sleep>":
                                        time.sleep(sleep_time)
                                    case "<win>":
                                        kb.send_keys("{LWIN}")
                                    case "<up>":
                                        kb.send_keys("{UP}")
                                    case "<down>":
                                        kb.send_keys("{DOWN}")
                                    case "<left>":
                                        kb.send_keys("{LEFT}")
                                    case "<right>":
                                        kb.send_keys("{RIGHT}")
                                if k < repeat_time-1:
                                    time.sleep(key_delay)
                                    self.slept = True
                            if self.slept == True:
                                self.slept = False
                        case _:
                            kb.send_keys(key)

main = QKIH()

mainloop()
