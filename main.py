import tkinter as tk
from tkinter import ttk
import openpyxl
from openpyxl import load_workbook
import customtkinter as ctk
import ttkbootstrap as tb
from functools import partial


win = ctk.CTk()
win.title("Data Entry Form")
win.geometry("1060x430")
win.resizable(False, False)


# Set up main frame
frame = ttk.Frame(win)
frame.pack()

# SETUP DATA MANAGEMENT
path = "Excel_file/Employee Sample Data.xlsx" # SETUP EXCELL FILE HERE
workbook = openpyxl.load_workbook(path)
sheet = workbook.active
list_values = list(sheet.values)
filter_search = []

tree_frame = ttk.Frame(win, width=1060, height=240)
tree_frame.pack()
tree_frame.pack_propagate(False)


# SETUP Treeview functions
def full_data_list(win):
    cols = list_values[0]

    my_tree = ttk.Treeview(tree_frame, columns=cols, show="headings")

    tree_scroll = tb.Scrollbar(tree_frame, orient='horizontal', bootstyle="danger round", command=my_tree.xview)
    tree_scroll.pack(side='bottom', fill="y")

    for col_name in cols:
        my_tree.heading(col_name, text = col_name)


    my_tree.pack(pady=15, padx=15)

    for value_tuple in list_values[1:]:
        my_tree.insert("", tk.END, values=value_tuple)

    tree_scroll.place(relx=0.03, rely=0.88, relwidth=0.9, relheight=0.06)
    my_tree.config(xscrollcommand=tree_scroll.set)

    return my_tree


def filtered_data(win, filtered_search, my_tree):
    for value_tuple in list_values[1:]:
        if (value_tuple[0].startswith(filtered_search[0]) and
                value_tuple[1].startswith(filtered_search[1]) and
                value_tuple[2].startswith(filtered_search[2]) and
                value_tuple[3].startswith(filtered_search[3]) and
                value_tuple[4].startswith(filtered_search[4])):

            my_tree.insert("", tk.END, values=value_tuple)

def get_name_for_forms(idx):
    return list_values[0][idx]

def get_info(idx):
    info = set([i[idx] for i in list_values[1:]])
    return list(info)


# Set Up Frames and app
class FrameBox:
    def __init__(self, text: str):
        self.text = text
        self.curr_frame = tk.Label(frame)


class InsideFrameBox:
    FORMS = {}
    ENTRIES = {}
    def __init__(self, parent_frame, name, bar):
        self.parent_frame = parent_frame
        self.name = name
        self.bar = bar

    def create_current_form_and_entries(self):
        # Create Form Name
        curr_element = tk.Label(self.parent_frame, text=self.name)
        InsideFrameBox.FORMS[self.name] = curr_element

        #Create Form Entry and Add them in collections
        if self.bar:
            idx = len(self.ENTRIES.keys())
            curr_entry = ttk.Combobox(self.parent_frame, values=get_info(idx))
        else:
            curr_entry = ttk.Entry(self.parent_frame)

        InsideFrameBox.ENTRIES[self.name] = curr_entry

    @staticmethod
    def display_forms():
        row, col = 0, 0
        # Get Forms and display them in the Parent Frame
        idx = 0
        for text in InsideFrameBox.FORMS.keys():
            idx += 1
            if idx >= len(list_values[0]) / 2:
                idx = 1
                row, col = row + 2,  0
            InsideFrameBox.FORMS[text].grid(row=row, column=col,padx=30, pady=10)

            InsideFrameBox.ENTRIES[text].grid(row=row+1, column=col, padx=30, pady=(0, 10))
            col += 1


# BUILD CORE
UserInformation = FrameBox("Search Information")
UserInformation.curr_frame.grid(row=0, column=0, padx=30, pady=20)

#Create each label text according to excel first 5 columns
def create_elements():
    for idx, heading in enumerate(list_values[0][0:5]):
        curr_bar = False
        if idx >= 3:
            curr_bar = True
        element = InsideFrameBox(UserInformation.curr_frame, name=get_name_for_forms(idx), bar=curr_bar).create_current_form_and_entries()

create_elements()

# Buttons and Button command functions
def search_data():
    filter_search = []
    for key in InsideFrameBox.ENTRIES:
        item = InsideFrameBox.ENTRIES[key].get()
        filter_search.append(item)

    if [i for i in filter_search if i.strip() != ""]:
        my_tree.delete(*my_tree.get_children())
        filtered_data(win, filter_search, my_tree)


def reset_button():
    my_tree.delete(*my_tree.get_children())
    filtered_data(win, [''] * 5, my_tree)


button1 = ctk.CTkButton(UserInformation.curr_frame,text="Enter Data", command=search_data)
button1.grid(row=15, column=2, pady=(25, 10))


button3 = ctk.CTkButton(UserInformation.curr_frame, text="Reset", command= reset_button)
button3.grid(row=15, column=1, pady=(25, 10))

# Show Everything
my_tree = full_data_list(win)
InsideFrameBox.display_forms()
win.mainloop()