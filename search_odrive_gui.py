import os
import re
import time
import tkinter
import tkinter as tk
from sys import exit as sexit
from tkinter import ttk
import ctypes
import win32console
import win32gui

import customtkinter as ctk
from collections import namedtuple

#ic.configureOutput(prefix=f'Debug | ', includeContext=True)

loading_message = None
element_list = None
og_element_list = None
entered_search = None
base_directories = None
search_text = None
dot_search = None
driver_search = None
docuware_query = None
display_to_company = None
company_to_location = None # This maps each company to its location. Ex. Company LLC -> TCS
root = None
search_hint_element = None
geometry = None
started_program = False
line_frame = None
full_selection = None
full_selection_name = None
full_selection_dot = None
full_selection_type = None
version = "V5.5"
tcs_blue = "#1E2D60"
tcs_red = "#BF1E2E"
TypeDirectory = namedtuple("TypeDirectory", ["name", "path"])

es_active = "ES - Active"
tcs_active = "TCS - Active"
es_inactive = "ES - Inactive"
tcs_inactive = "TCS - Inactive"
tcs_main = "TCS"
es_main = "ES"

dot_label = None
name_label = None
type_label = None



es_active_dir = TypeDirectory(es_active, f"O:\\Docuware Backup\\{es_active}")
tcs_active_dir = TypeDirectory(tcs_active, f"O:\\Docuware Backup\\{tcs_active}")

es_inactive_dir = TypeDirectory(es_inactive, f"O:\\Docuware Backup\\{es_inactive}")
tcs_inactive_dir = TypeDirectory(tcs_inactive, f"O:\\Docuware Backup\\{tcs_inactive}")

tcs_main_dir = TypeDirectory(tcs_main, r"O:\1 TCS Clients")
#tcs_main_dir = TypeDirectory(tcs_main, r"O:\Chase Minert") # this was just to test the program without having to wait 5 minutes
es_main_dir = TypeDirectory(es_main, r"O:\1 ES Clients")
driver_to_path = dict()
location_to_path = {
    es_active: es_active_dir,
    tcs_active: tcs_active_dir,
    es_inactive: es_inactive_dir,
    tcs_inactive: tcs_inactive_dir,
    es_main: es_main_dir,
    tcs_main: tcs_main_dir
}

location_to_display_location = {
    es_active: "ES",
    tcs_active: "TCS",
    es_inactive: "ES",
    tcs_inactive: "TCS",
    tcs_main: "TCS",
    es_main: "ES"
}

search_entry = None
dimentions = "710x493"

def custom_sort_key(s):
    # Check if the first character is alphabetic
    return (0, s) if s[0].isalpha() else (1, s)
def get_all_companies(base_dirs):
    global dot_search
    global display_to_company
    global company_to_location
    all_comps = []
    for location, base_dir in base_dirs:
        temp_list = [(location, element) for element in os.listdir(base_dir)]
        all_comps.extend(temp_list)
    highlight_top()
    all_comps_display = []
    if is_dot_search():
        for location, element in all_comps:
            company_to_location[element] = location
            reversed = get_display_name(element)
            if reversed:
                all_comps_display.append(reversed)
                display_to_company[reversed] = element
    elif is_driver_search():
        for location, element in all_comps:
            company_to_location[element] = location
            company_path = get_path_from_company(element, full=True) # gets the full path to the company
            drivers = get_drivers(company_path)

            all_comps_display.extend(drivers)
            for driver in drivers:
                display_to_company[driver] = element
        all_comps_display.sort(key=custom_sort_key) # put numbers and special characters at the end
    else:
        for location, element in all_comps:
            company_to_location[element] = location
            display_name = get_display_name(element, swap=False)
            all_comps_display.append(display_name)
            display_to_company[display_name] = element
    return all_comps_display

def get_drivers(location):
    global driver_to_path

    dqf_path = os.path.join(location, "DQF")
    drivers = []
    if os.path.exists(dqf_path) and os.path.isdir(dqf_path):
        for driver in os.listdir(dqf_path):
            full_path = os.path.join(dqf_path, driver)
            if os.path.isdir(full_path):
                driver_to_path[driver] = full_path
                drivers.append(driver)
    return drivers
def update_label(loading=False):
    global search_hint_element, docuware_query

    if loading:
        search_hint_element.configure(text="Loading...")
    update_full_selection(initialize=True)
    if is_dot_search() and docuware_query:
        search_hint_element.configure(text="Enter DOT Number")
        line_frame.grid_forget()
        line_frame.grid(padx=(93, 445), sticky="NS")
    elif is_driver_search():
        search_hint_element.configure(text="Enter Driver Name")
    elif docuware_query:
        search_hint_element.configure(text="Enter Company Name")
        line_frame.grid_forget()
        line_frame.grid(padx=(417, 121), sticky='NS')
    else:
        search_hint_element.configure(text="Enter Company Name")


def get_display_name(text, swap=True):
    global docuware_query

    if not docuware_query:
        max_len = 60 # for regular search o drive program
        return shorten_name(text, 60)

    dot_number_length = 7
    max_dot_length = 10
    character = " - "
    new_character = "|"
    last_index = text.rfind(character)
    # Check if the character is found
    if last_index != -1:
        # Split the string into two parts
        company = text[:last_index].strip()
        dot_number = text[last_index + 3:].strip().replace("NODOT", "NO DOT").replace("NoDOT", "NO DOT")

        company = shorten_name(company)

        '''
        if len(dot_number) <= dot_number_length:
            dot_number = dot_number.ljust(max_dot_length)
        else:
            dot_number = dot_number[:dot_number_length] + "..."
        '''
        dot_number = shorten_name(dot_number, max_dot_length, dot_number=True)

        # Concatenate in reversed order
        if swap:
            return dot_number + new_character + " " + company
        else:
            return company + new_character + " " + dot_number
    else:
        return text


def extract_info(text):
    global docuware_query
    if not docuware_query:
        return text, None
    character = " - "
    last_index = text.rfind(character)
    # Check if the character is found
    if last_index != -1:
        # Split the string into two parts
        company = text[:last_index].strip()
        dot_number = text[last_index + 3:].strip()

        return company, dot_number
    else:
        return text


def shorten_name(name: str, max_len=46, dot_number=False): # dot_number parameter just means if name is a dot number
    if is_dot_search() and not dot_number:
        max_len += 1

    if len(name) < max_len:
        return name.ljust(max_len)
    else:
        return name[:max_len - 3] + "..."


def string_is_int(string):
    try:
        int(string)
        return True
    except ValueError:
        return False


def get_window_position(geometry_str):
    match = re.search(r'\+(\d+)\+(\d+)', geometry_str)
    if match:
        return match.groups()  # Returns (x_offset, y_offset) as a tuple
    return None, None


def load_company_list():  # (display name maps to actual name)
    global element_list
    global og_element_list
    global display_to_company
    global suggestions_listbox
    global base_directories
    global search_entry
    global line_frame, root

    update_label()
    search_entry.delete(0, tk.END)

    element_list = get_all_companies(base_directories)
    og_element_list = element_list[:]

    suggestions_listbox.delete(0, tk.END)

    for element in og_element_list:
        suggestions_listbox.insert(tk.END, element)
    time.sleep(2)
    highlight_top()



def on_search_change(event):
    global entered_search
    event: tkinter.Event
    if event and (event.keycode == 40 or event.keycode == 38):
        return
    if event is None:
        global element_list
        global og_element_list
        for element in element_list:
            suggestions_listbox.insert(tk.END, element)
    event_str = ""
    if event is not None:
        event_str = event.keysym
    if event_str == "Return" and not entered_search:
        on_search_enter(event)
    search_term = search_entry.get().lower()
    search_term_len = len(search_term)
    if search_term_len == 1:
        return
    matching_elements = [element for element in element_list if element.lower().startswith(search_term)]
    suggestions_listbox.delete(0, tk.END)

    # Display matching elements as suggestions
    for element in matching_elements:
        suggestions_listbox.insert(tk.END, element)
    highlight_top()


def select_wrapper(event, window):
    global all_dots_test

    on_suggestion_select()
    window.iconify()


def on_suggestion_select():
    global dot_search
    global display_to_company
    global base_directories
    global driver_to_path

    selected_index = suggestions_listbox.curselection()
    if selected_index:
        displayed_name = suggestions_listbox.get(selected_index[0])
        search_entry.delete(0, tk.END)
        search_entry.insert(0, displayed_name)
        suggestions_listbox.place_forget()
        global entered_search
        entered_search = True
        if is_driver_search():
            direct_path = driver_to_path[displayed_name]
            load_direct_path(direct_path)
        else:
            actual_name = display_to_company[displayed_name]
            load_path(base_directories, actual_name)

        clear_program()


def do_nothing(event):
    return

def update_full_selection(index=0, initialize=False):
    global full_selection, suggestions_listbox, loading_message, docuware_query, location_to_display_location
    global full_selection_name, full_selection_dot, full_selection_type
    if initialize:
        full_selection_name.configure(text="")
        full_selection_dot.configure(text="")
        full_selection_type.configure(text="")
        return


    display_str = suggestions_listbox.get(index)

    if display_str == loading_message:
        return

    full_str = display_to_company[display_str] # "Really Long Company Name..." will map to "Really Long Company Name LLC" for example
    full_location = company_to_location[full_str] # "Company LLC" will map to "TCS" for example
    display_location = location_to_display_location[full_location] # "TCS - Active" will display as just "TCS" for example
    company, dot = extract_info(full_str)
    if is_dot_search() and docuware_query:
        all_info_str = f"{dot} - {company}"
    elif docuware_query:
        all_info_str = f"{company} - {dot}"
    else:
        all_info_str = company
    if not dot:
        dot = "Not Indexed"
    full_selection_name.configure(text=company)
    full_selection_dot.configure(text=dot)
    full_selection_type.configure(text=full_location)





def on_motion(event):
    index = suggestions_listbox.nearest(event.y)
    suggestions_listbox.selection_clear(0, tk.END)
    suggestions_listbox.select_set(index)
    suggestions_listbox.activate(index)
    update_full_selection(index)


def on_arrow_down(event):
    max_index = suggestions_listbox.size() - 1

    index = suggestions_listbox.curselection()[0] + 1

    index = max_index if index == max_index + 1 else index

    suggestions_listbox.see(index)
    suggestions_listbox.selection_clear(0, tk.END)
    suggestions_listbox.select_set(index)
    suggestions_listbox.activate(index)
    update_full_selection(index)


def on_arrow_up(event):
    index = suggestions_listbox.curselection()[0] - 1

    index = 0 if index == -1 else index

    suggestions_listbox.see(index)
    suggestions_listbox.selection_clear(0, tk.END)
    suggestions_listbox.select_set(index)
    suggestions_listbox.activate(index)
    update_full_selection(index)
    # suggestions_listbox.yview_scroll(1, tkinter.UNITS)


def highlight_top():
    suggestions_listbox.selection_clear(0, tk.END)
    suggestions_listbox.select_set(0)
    suggestions_listbox.activate(0)


def on_leave(event):
    suggestions_listbox.selection_clear(0, tk.END)
    highlight_top()


def on_search_enter(event):
    global entered_search
    global driver_to_path
    top_element = suggestions_listbox.get(0)
    search_entry.delete(0, tk.END)
    search_entry.insert(0, top_element)
    suggestions_listbox.place_forget()
    if is_driver_search():
        direct_path = driver_to_path[top_element]
        load_direct_path(direct_path)
    else:
        actual_name = display_to_company[top_element]
        load_path(base_directories, top_element)
    entered_search = False
    clear_program()


def clear_program():
    global entered_search
    search_entry.delete(0, tk.END)
    suggestions_listbox.delete(0, tk.END)
    for element in og_element_list:
        suggestions_listbox.insert(tk.END, element)
    set_focus()
    highlight_top()


def set_focus():
    search_entry.focus_set()


def set_docuware_query(search_docuware): # 0 = TCS/ES, 1 = Docuware Active, 2 = Docuware Inactive
    global docuware_query
    global root
    global search_text
    global base_directories
    global dot_search
    global geometry
    global tcs_active_dir, es_active_dir, tcs_inactive_dir, es_inactive_dir, tcs_main_dir, es_main_dir

    geometry = root.geometry()

    docuware_query = search_docuware != 0

    if search_docuware == 0:
        search_text = "Search TCS/ES"
        base_directories = [es_main_dir, tcs_main_dir]
    elif search_docuware == 1:
        search_text = "Search Docuware Active"
        base_directories = [es_active_dir, tcs_active_dir]
    elif search_docuware == 2:
        search_text = "Search Docuware Inactive"
        base_directories = [es_inactive_dir, tcs_inactive_dir]

    root.withdraw()
    main_program_display()
    # root.destroy()

def get_path_from_company(company_name, full=False): # full means get the path to the company, otherwise it just gets the base directory
    global company_to_location, location_to_path

    location = company_to_location[company_name]
    _, base_dir = location_to_path[location]

    if full:
        return os.path.join(base_dir, company_name)
    else:
        return base_dir
def load_path(base_dirs, company_name, direct_path=None):

    # List the contents of the directory
    global loading_message
    global company_to_location
    global location_to_path

    if company_name == loading_message:
        return

    base_dir = get_path_from_company(company_name) # we don't care about the name at this point, thus the _

    path = os.path.join(base_dir, company_name)
    load_direct_path(path)

def load_direct_path(path):
    if os.path.exists(path):
        os.startfile(path)



def get_exe_name():
    current_directory = os.getcwd()
    for file in os.listdir(current_directory):
        if file.endswith(".exe"):
            return file[0:-4]


def is_dot_search():
    global dot_search
    return dot_search.get() == "1"

def is_driver_search():
    global driver_search
    return driver_search.get() == "1"




def set_dot_search(val: bool):
    global dot_search
    # dot_search: tkinter.StringVar
    dot_search.set("0")


def check_frame_size(frame):
    global root
    root_size = (root.winfo_width(), root.winfo_height())

    current_size = (frame.winfo_width(), frame.winfo_height())
    print(f"{current_size} | {root_size}")

    root.after(100, lambda: check_frame_size(frame))


def on_close(*args):
    for window in args:
        window.destroy()
    sexit()


def set_program():
    global loading_message, element_list, og_element_list, entered_search
    global base_directories, search_text, dot_search, driver_search, docuware_query
    global display_to_company, root, search_hint_element, started_program, type_label
    global search_entry, all_dots_test, line_frame, full_selection, company_to_location
    global full_selection_name, full_selection_dot, full_selection_type
    global name_label, dot_label, type_label, driver_to_path

    loading_message = "Loading From O Drive..."
    element_list = [loading_message]
    og_element_list = []
    entered_search = False
    base_directories = []
    search_text = ""
    dot_search = ""
    driver_search = ""
    docuware_query = ""
    display_to_company = dict()
    company_to_location = dict()
    search_entry = None
    line_frame = None
    full_selection = None
    type_label = None
    name_label = None
    dot_label = None
    full_selection_name = None
    full_selection_dot = None
    full_selection_type = None
    driver_to_path = dict()

def show_root():
    global root
    root.deiconify()
    #root: customtkinter.windows.ctk_tk.CTk



def selection_screen_display():
    global root
    global version
    global geometry
    global started_program
    global tcs_blue, tcs_red
    global dimentions

    file_name = get_exe_name()
    if started_program:
        x, y = get_window_position(geometry)
        if x and y:
            root.geometry(f'+{x}+{y}')
            root.mainloop()
    else:
        root = ctk.CTk()

    root.geometry(dimentions)

    root.protocol("WM_DELETE_WINDOW", lambda: on_close(root))

    started_program = True

    root.title(file_name)

    # Layout configuration

    root.iconbitmap("logo.ico")

    # Create and place the buttons
    intro_font = ctk.CTkFont(family="Arial", size=35, weight='bold', underline=1)
    question = ctk.CTkLabel(root, text=f"Search O Drive {version}", font=intro_font, text_color="white")
    question.grid(pady=(60, 0))

    root.columnconfigure(0, weight=1)

    intro_hint = ctk.CTkLabel(root, text="Select A Search", font=("Arial", 15), text_color="white")
    intro_hint.grid(pady=(45, 15))

    buttons_frame = ctk.CTkFrame(root, width=1, height=1, fg_color="transparent")
    buttons_frame.grid()

    # 0 = TCS/ES, 1 = Docuware Active, 2 = Docuware Inactive

    button_font = ("Verdana", 20)
    o_drive_button = ctk.CTkButton(buttons_frame, text="TCS/ES", command=lambda: set_docuware_query(0),
                                   font=button_font, border_width=2, border_color=tcs_red, fg_color=tcs_blue,
                                   text_color="white", hover_color="#000040", corner_radius=10)
    o_drive_button.grid(pady=(11, 25), ipadx=38, ipady=5)

    docuware_active_button = ctk.CTkButton(buttons_frame, text="Docuware Active", command=lambda: set_docuware_query(1),
                                    font=button_font, border_width=2, border_color=tcs_red, fg_color=tcs_blue,
                                    text_color="white", hover_color="#000040", corner_radius=10)
    docuware_active_button.grid(pady=(0, 25), ipadx=14, ipady=5)

    docuware_inactive_button = ctk.CTkButton(buttons_frame, text="Docuware Inactive", command=lambda: set_docuware_query(2),
                                    font=button_font, border_width=2, border_color=tcs_red, fg_color=tcs_blue,
                                    text_color="white", hover_color="#000040", corner_radius=10)
    docuware_inactive_button.grid(pady=(0, 25), padx=20, ipady=5, ipadx=5)

    # Start the GUI event loop
    root.mainloop()


def main_program_display():
    file_name = get_exe_name()
    global dimentions
    global root
    global search_text
    global dot_search, driver_search
    global docuware_query
    global search_hint_element, type_label
    global tcs_blue, tcs_red, all_dots_test, line_frame, full_selection
    global dot_label, name_label, type_label
    global full_selection_name, full_selection_dot, full_selection_type
    main_window = ctk.CTkToplevel()
    main_window.protocol("WM_DELETE_WINDOW", lambda: on_close(main_window, root))

    main_window.iconbitmap("logo.ico")

    x, y = get_window_position(geometry)
    if x and y:
        main_window.geometry(f'+{x}+{y}')

    dot_search = ctk.StringVar()
    driver_search = ctk.StringVar()

    main_window.geometry(dimentions)
    main_window.title(file_name)

    style = ttk.Style()
    style.configure("Custom.TListbox", padding=(5, 5))

    # Create a label and an entry widget for the search

    header_font = ctk.CTkFont(family="Arial", size=20, weight='bold', underline=3)

    main_window.columnconfigure(0, weight=1)

    very_top_frame = ctk.CTkFrame(main_window, fg_color="transparent", width=670, height=30)
    very_top_frame.grid_propagate(False)
    very_top_frame.grid(row=0, column=0, pady=(20, 0))

    #very_top_frame.columnconfigure(0, weight=1)
    #very_top_frame.columnconfigure(1, weight=1)
    #very_top_frame.columnconfigure(2, weight=1)

    empty_frame = ctk.CTkFrame(very_top_frame, width=180, height=10, fg_color="transparent")
    empty_frame.grid_propagate(False)
    empty_frame.grid(row=0, column=0)

    program_label = ctk.CTkLabel(very_top_frame, text=search_text, font=header_font, text_color="white", width=300)
    program_label.grid(row=0, column=1)

    credit_label = ctk.CTkLabel(very_top_frame, text="Created by Chase Minert", font=("Arial", 12), text_color="grey40", anchor='w')
    credit_label.grid(row=0, column=2, padx=(30, 0))



    center_frame = ctk.CTkFrame(main_window, border_color="#525252", border_width=1)
    center_frame.grid(row=1, column=0, columnspan=100, pady=(20, 21))

    # instr_label = ctk.CTkLabel(center_frame, text="Click the company name\nor press enter:", font=("Arial", 12), text_color="grey50")
    # instr_label.pack(padx=10, pady=20, side=tk.LEFT)

    button_font = ("Arial", 20, "bold")
    switch_button = ctk.CTkButton(center_frame, text="Change Search", command=lambda: back_to_selection(main_window),
                                  font=button_font, border_width=2, border_color=tcs_red, fg_color=tcs_blue,
                                  hover_color="#000040", corner_radius=10, text_color="white")
    switch_button.grid(row=0, column=0, padx=(10, 2), pady=10, ipady=5)

    search_frame = ctk.CTkFrame(center_frame)
    search_frame.grid(row=0, column=1, pady=10, padx=10)

    search_hint_element = ctk.CTkLabel(search_frame, text="Enter Company Name", font=("Arial", 12, 'bold'), text_color="white")
    search_hint_element.grid(row=0, column=0)

    # credit_label.place(relx=0, anchor='w')
    # credit_label.pack()
    global search_entry
    search_entry = ctk.CTkEntry(search_frame, width=250, height=25)  # Set dark blue color theme
    # search_entry.grid(row=3, column=3)
    search_entry.grid(row=1, column=0, pady=(0, 10), padx=20)

    checkbox_frame = ctk.CTkFrame(center_frame, width=90, height=55)
    checkbox_frame.grid(row=0, column=2, pady=10, padx=38)



    if docuware_query:
        checkbox_text = ctk.CTkLabel(checkbox_frame, text="Search by DOT#", font=("Arial", 12, "bold"), text_color="white")
        checkbox_text.grid(row=0, column=0, padx=10, ipady=0)
        checkbox = ctk.CTkCheckBox(checkbox_frame, text="", font=("Arial", 12, "bold"), variable=dot_search,
                                   onvalue=True, offvalue=False, command=load_company_list, width=24,
                                   hover_color=tcs_blue, fg_color=tcs_blue, border_width=2)
    else:
        checkbox_text = ctk.CTkLabel(checkbox_frame, text="Search by DQF", font=("Arial", 12, "bold"), text_color="white")
        checkbox_text.grid(row=0, column=0, padx=10, ipady=0)
        checkbox = ctk.CTkCheckBox(checkbox_frame, text="", font=("Arial", 12, "bold"), variable=driver_search,
                                   onvalue=True, offvalue=False, command=load_company_list, width=24,
                                   hover_color=tcs_blue, fg_color=tcs_blue, border_width=2)
    checkbox.grid(row=1, column=0, padx=(13, 7), pady=(0, 10))





    selection_frame = ctk.CTkFrame(main_window, width=553, height=100, fg_color="transparent")
    selection_frame.grid(row=2, column=0)

    selection_frame.columnconfigure(0, weight=1)


    name_label = ctk.CTkLabel(selection_frame, text="Company Name:", font=("Arial", 14, "bold"), anchor='nw', width=115)
    name_label.grid(row=0, column=0, padx=2, pady=(2, 9), sticky='NW')

    dot_label = ctk.CTkLabel(selection_frame, text="Dot Number:", font=("Arial", 14, "bold"), anchor='nw')
    dot_label.grid(row=1, column=0, padx=2, pady=(0, 9), sticky='NW')

    type_label = ctk.CTkLabel(selection_frame, text="Company Type:", font=("Arial", 14, "bold"), anchor='nw')
    type_label.grid(row=2, column=0, padx=2, pady=(0, 2), sticky="NW")


    max_len = 433

    full_selection_name = ctk.CTkLabel(selection_frame, text="", font=("Arial", 14), wraplength=max_len, anchor="nw", justify="left", width=max_len)
    full_selection_name.grid(row=0, column=1, sticky='NW', pady=(2, 5))

    full_selection_dot = ctk.CTkLabel(selection_frame, text="", font=("Arial", 14), wraplength=max_len, anchor="nw", justify='left')
    full_selection_dot.grid(row=1, column=1, sticky='NW', pady=(0, 5))

    full_selection_type = ctk.CTkLabel(selection_frame, text="", font=("Arial", 14), wraplength=max_len, anchor="nw", justify='left')
    full_selection_type.grid(row=2, column=1, sticky='NW', pady=(0, 2))


    main_window.rowconfigure(3, weight=1)
    # Create a Listbox widget for displaying suggestions
    global suggestions_listbox
    suggestions_listbox = tk.Listbox(main_window, border=True, background="grey15", foreground="white", justify="left",
                                     width=60, font=("Courier New", 11, "bold"), borderwidth=3, height=25)

    suggestions_listbox.grid(row=3, column=0, pady=(13, 30), sticky="N")

    # suggestions_listbox.place_forget()

    line_frame = ctk.CTkFrame(suggestions_listbox, width=3, height=290, fg_color="white", border_width=1,
                              corner_radius=0, border_color="white")

    spacer_frame = ctk.CTkFrame(main_window, width=200, height=2000, fg_color="black")
    #spacer_frame.grid(row=4, column=0)

    main_window.after(500, load_company_list)
    main_window.after(150, lambda: main_window.iconbitmap("logo.ico"))
    # root.after(5000, lambda: check_frame_size(checkbox_frame))

    main_window.bind("<Return>", lambda event, window=main_window: select_wrapper(event, window))
    # Bind events
    search_entry.bind("<KeyRelease>", on_search_change)
    suggestions_listbox.bind("<ButtonRelease-1>", lambda event, window=main_window: select_wrapper(event, window))

    suggestions_listbox.bind("<Motion>", on_motion)

    main_window.bind("<Down>", on_arrow_down)
    main_window.bind("<Up>", on_arrow_up)
    # suggestions_listbox.bind("<Leave>", on_leave)

    # root.bind("<Visibility>", highlight_top_wrapper)

    # root.after(100, set_focus)

    on_search_change(None)
    # Start the GUI main loop
    # root.mainloop()


def back_to_selection(window):
    global geometry

    geometry = window.geometry()
    show_root()
    window.focus_set()
    window.destroy()
    main()


def main():
    ctk.set_appearance_mode("dark")
    global geometry
    global root

    set_program()
    selection_screen_display()
    main_program_display()

def close_console():
    # Get the console window handle
    SW_MINIMIZE = 6
    # Get the console window handle
    console_window_handle = ctypes.windll.kernel32.GetConsoleWindow()

    # Minimize the console window
    ctypes.windll.user32.ShowWindow(console_window_handle, SW_MINIMIZE)
    window = win32console.GetConsoleWindow()

    # Hide the console window
    win32gui.ShowWindow(window, 0)

if __name__ == "__main__":
    close_console()
    main()
