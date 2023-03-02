import os
import re
import customtkinter
import docx
import extract_msg

from sys import platform
import tkinter
from tkinter import filedialog
from tkinter import messagebox

# Initializing main window
root = tkinter.Tk()
root.title("Simple File Search by RS")
wide = 700
root.geometry(str(wide) + 'x820')
customtkinter.set_appearance_mode("System")

# Supported extensions
extensions = {'TXT': 0, 'DOCX': 0, 'MSG': 0}


def clear():
    """Clears path widget"""
    path_entry.delete(0, tkinter.END)


def browse_files():
    """Opens file dialog box"""
    clear()

    if platform == "win32":
        path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents')

    elif platform == "darwin":
        path = os.path.expanduser("~/Documents/")

    else:
        messagebox.showwarning("Platform not supported", "Support for Windows & MacOS only")
        return

    dir_name = filedialog.askdirectory(initialdir=path)

    path_entry.insert(0, dir_name)


def get_path():
    """Gets path from path widget"""
    path = path_entry.get()
    return path


def form_extensions_list():
    """Creates list of file extensions to search files
    of respective filetypes.  Variable extensions is defined
    at the checkbuttons section and represents a dict of
    doctype(:str) as key and Tkinter StringVar as value
    """
    form_list = []
    for doc_type in extensions:
        checkb_var = extensions[doc_type].get()
        if checkb_var != '0':
            form_list.append(checkb_var)
        else:
            pass
    return form_list


def convert_to_text(folder, file_name, file_extension):
    """Converts file of certain filetype into iterable object
    or list of paragraphs
    """
    full_path = os.path.join(folder, file_name + file_extension)

    if file_extension == '.txt':
        with open(full_path) as text:
            text = text.readlines()

    elif file_extension == '.msg':
        msg = extract_msg.Message(full_path)
        msg_body = msg.body
        msg_body = re.sub(r'\n\s*\n', '\n', msg_body)
        msg_body = msg_body.splitlines()
        text = msg_body

    elif file_extension == '.docx':
        doc = docx.Document(full_path)
        text = doc.paragraphs

    else:
        pass

    return text


def search():
    """Performs search in each paragraph of each file in directory
    if its filetype is supported.  Available: docx, msg, txt.
    Linked to button"""
    quary = search_entry.get()
    folder = get_path()
    formats = form_extensions_list()

    if len(quary) <= 3:
        messagebox.showwarning("Short query", "Search query should be "
                                              "4+ characters long")
        return

    elif len(quary) >= 50:
        messagebox.showwarning("Long query", "Search query should be "
                                             "less than 50 characters long")
        return

    elif not os.path.exists(folder):
        messagebox.showwarning("No directory", "There is no such directory")
        return

    elif not formats:  # empty formats list returns False

        messagebox.showwarning("No filetypes selected", "Choose at least "
                                                        "one filetype")
        return

    output_text.delete("1.0", tkinter.END)
    output_text.tag_remove('found', '1.0', tkinter.END)
    output_text.tag_remove('filename', '1.0', tkinter.END)

    for file in os.listdir(folder):
        file_name, file_extension = os.path.splitext(file)
        file_extension = file_extension.lower()

        if file_extension not in formats:
            continue

        try:
            text_by_paras = convert_to_text(folder, file_name, file_extension)
        except UnboundLocalError:
            continue
        except TypeError:
            if file_extension == '.msg':
                print(file_name + '\n--- is encrypted')
                continue

        search_results = {}

        for para in text_by_paras:

            para = para.text + '\n' * 0 \
                if isinstance(para, docx.text.paragraph.Paragraph) else para

            result = re.search(quary, para, flags=re.I)

            if result:
                search_results.update({para: result.start()})

        if not search_results:
            continue

        output_text.insert(tkinter.END, f"\tResults for file \"{file}\"\n\n",
                           'filename_tag')

        for key_para, start_value in search_results.items():
            idx = 'end-1 chars'  # may be 'end-1 chars' /
            idx = output_text.index(idx)
            output_text.insert(tkinter.END, key_para)
            len_quary = len(quary)
            idx = f"{idx}+{start_value}c"  # offset may be +1 char
            lastidx = f"{idx}+{len_quary}c"
            output_text.tag_add('found', idx, lastidx)
            output_text.insert(tkinter.END, '\n' * 2)

    if output_text.compare("end-1c", "==", "1.0"):
        output_text.insert(tkinter.END, f"\tNo results found",
                           'filename_tag')

    output_text.tag_config('filename_tag', font=font_heading_text)
    output_text.tag_config('found', background='yellow', foreground='red')


# -------------- COLOURS, FONTS AND THEMES

color_background = '#f0f8ff'
color_text_field = color_background  # May be changed for readability
color_buttons_entry = '#566D7E'
color_main_text = '#330000'
color_entry_text = 'white'
root.configure(bg=color_background)
font_buttons = ('Calibri', 18, "bold")
font_labels = ('Calibri', 16, "bold")
font_entries = ('Calibri', 14)
font_main_text = ('Times', 14)
font_heading_text = ('Times', 14, "bold")

# -------------- CREATE WIDGETS

# Scrollbar
scrollbar = tkinter.Scrollbar(root, orient=tkinter.VERTICAL)

# Filepath entry frame
frame_path = tkinter.Frame(root, bg=color_background)
path_label = tkinter.Label(frame_path, text='Insert path:', font=font_labels,
                           bg=color_background)
path_entry = customtkinter.CTkEntry(master=frame_path, width=350,
                                    fg_color=color_buttons_entry, font=font_entries,
                                    text_color=color_entry_text)
clear_button = customtkinter.CTkButton(master=frame_path, width=130,
                                       text='Clear', command=clear,
                                       font=font_buttons, fg_color=color_buttons_entry)
button_explore = customtkinter.CTkButton(master=frame_path, text="Browse Folder",
                                         command=browse_files, font=font_buttons,
                                         fg_color=color_buttons_entry)

# Search entry frame
frame_search = tkinter.Frame(root, bg=color_background)
search_label = tkinter.Label(frame_search, text='Search text:',
                             font=font_labels, bg=color_background)
search_entry = customtkinter.CTkEntry(master=frame_search, width=350,
                                      fg_color=color_buttons_entry, font=font_entries,
                                      text_color=color_entry_text)
search_button = customtkinter.CTkButton(master=frame_search, width=140,
                                        text='Search', command=search,
                                        font=font_buttons, fg_color=color_buttons_entry)

# Checkbuttons frame
frame_checkb = tkinter.Frame(root, bg=color_background)

output_text = tkinter.Text(root, bg=color_text_field, font=font_main_text,
                           wrap='word', width=100, borderwidth=30,
                           relief=tkinter.FLAT)

# -------------- PACK WIDGETS

scrollbar.pack(side=tkinter.RIGHT, fill=tkinter.Y)

frame_path.pack()
button_explore.pack(side=tkinter.BOTTOM, anchor='center', pady=10)
path_label.pack(side=tkinter.LEFT)
path_entry.pack(side=tkinter.LEFT, pady=10, padx=10)
clear_button.pack(side=tkinter.LEFT, ipadx=5)

frame_search.pack()
search_label.pack(side=tkinter.LEFT)
search_entry.pack(side=tkinter.LEFT, pady=10, padx=10)
search_entry.focus_set()
search_button.pack(side=tkinter.LEFT)

# -------------- CREATE AND PACK CHECKBUTTONS

frame_checkb.pack()

for doctype in extensions:  # defined at the beginning
    extensions[doctype] = tkinter.StringVar()
    button = tkinter.Checkbutton(frame_checkb, text=doctype, variable=extensions[doctype],
                                 onvalue='.' + doctype.lower(), font=font_labels,
                                 bg=color_background, highlightbackground=color_background)
    button.deselect()
    button.pack(side=tkinter.LEFT)

# -------------- PACK MAIN TEXT FIELD

output_text.pack(side=tkinter.TOP, fill=tkinter.Y, anchor=tkinter.N,
                 expand=1, padx=10, ipady=50)

output_text.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=output_text.yview)

# -------------- MAINLOOP
if __name__ == "__main__":
    root.mainloop()
