
"""
This program is created with the purpose of searching for PII (Personal Identifiable Information).
Specifically email addresses, names of persons, personal id numbers, monetary card numbers.
Location data found in image files is also extracted.
"""

__author__ = "Andre Sele"

import os
import re
# --- IMPORT SECTION ---
# Dependencies
import sqlite3
import threading
import time
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import asksaveasfile, askdirectory

import docx
import magic
import plum.exceptions
import spacy
from exif import Image
from openpyxl import load_workbook
from pdfminer.high_level import extract_text
from spacy.language import Language
from spacy_language_detection import LanguageDetector

# Spacy models (These are downloaded by the following example: python -m spacy download en_core_web_md)
import en_core_web_md
import nb_core_news_lg

localtime = time.asctime(time.localtime(time.time()))
utctime = time.asctime(time.gmtime(time.time()))

class Hits:
    """
    | Class variables for hits.
    """

    def __init__(self):
        self.Hits_li_email = []
        self.Hits_li_idNum = []
        self.Hits_li_cardNum = []
        self.Hits_li_gps = []
        self.Hits_li_names = []
        self.Hits_li_num = ''
        self.Time_used = ""
        self.Error_li = []


# Variable for accessing class Hits
Hits_ = Hits()
exit_event = threading.Event()  # For killing progress bar Thread (p2)


def select_dir() -> str:
    """
    | Tkinter widget for selecting directory to process.
    """
    win = Tk()
    win.withdraw()
    selected_dir = askdirectory(title="****** SELECT DIRECTORY TO PROCESS ******")
    win.destroy()
    return selected_dir


def progress_widget():
    """
    | Tkinter widget for progress bar.
    | Indicate app is processing (indeterminate mode).
    """
    while True:
        global win
        win = Tk()
        label = Label(win, text="Searching for PII", font="50")
        label.pack(pady=5)

        progbar = ttk.Progressbar(win, orient=HORIZONTAL, length=220, mode="indeterminate")
        progbar.pack(pady=20)
        win.geometry('300x150')
        win.title("PII_Finder")

        progbar.start()
        if exit_event.is_set():
            break

        win.mainloop()


def hits_to_file():
    """
    | Tkinter window
    | creates Tkinter window, button, label.
    """
    # Create an instance of tkinter window
    win = Tk()

    # Set the geometry of tkinter window
    win.geometry("400x250")
    T = Text(win, height=5, width=52)
    l = Label(win, text="Specify filename\n and path\n to store output")

    def save_file():
        """
        | Provide widget for saving hits to file.
        | Filetype options are TXT and CSV
        """
        f = asksaveasfile(initialfile='Untitled.txt',
                          defaultextension=".txt", filetypes=[("All Files", "*.*"), ("Text Documents", "*.txt")])
        sf = str(f)
        sf = sf.split()[1].replace('name=', '')
        sf = sf.replace("'", '')
        raw_sf = r'{}'.format(sf)
        win.destroy()

        # Write to file
        with open(raw_sf, 'a') as file:
            file.write('\n' + "Local time of creation: " + localtime + '\n' + "UTC time of creation: " + utctime + '\n')
            file.write('\n' + Hits_.Time_used + '\n')
            file.write('\n' + 'EMAIL ADDRESSES:' + '\n')
            for res in set(Hits_.Hits_li_email):
                file.write(res + '\n')
            file.write('\n' + 'ID NUMBERS:' + '\n')
            for res in set(Hits_.Hits_li_idNum):
                file.write(res + '\n')
            file.write('\n' + 'MONETARY CARD NUMBERS:' + '\n')
            for res in set(Hits_.Hits_li_cardNum):
                file.write(res + '\n')
            file.write('\n' + 'PERSON NAMES:' + '\n')
            for res in set(Hits_.Hits_li_names):
                file.write(res + '\n')
            file.write('\n' + 'GPS COORDINATES:' + '\n')
            for res in set(Hits_.Hits_li_gps):
                file.write(res + '\n')

        with open(raw_sf[:-3] + '_error_log.txt', 'a') as efile:
            for err in set(Hits_.Error_li):
                efile.write('\n' + str(err) + '\n')

    T.pack()
    l.pack()

    # Create a button
    btn = Button(win, text="Save", command=lambda: save_file())
    btn.pack(pady=10)

    win.mainloop()


def get_lang_detector(nlp, name) -> classmethod:
    """
    | Spacy LanguageDetector class
    """
    return LanguageDetector(seed=42)


def state_language(text: str) -> str:
    """
    | Use get_lang_detector function to detect language of text strings.
    | Returns appropriate spaCy model based on the language that is detected.
    | This program includes support for Norwegian and English.
    """
    nlp_model = spacy.load("en_core_web_sm")
    Language.factory("language_detector", func=get_lang_detector)
    nlp_model.add_pipe('language_detector', last=True)

    doc = nlp_model(text)
    language = doc._.language
    id_lang = language.get('language')
    if id_lang == 'no':  # Norwegian
        mod = 'nb_core_news_lg'
    else:
        mod = 'en_core_web_md'
    return mod


def convert_to_bytes(x: str) -> bytes:
    """
    | Convert input string to bytes
    """
    x = x.encode('utf-8')
    return x


def re_mail_matcher() -> str:
    """
    | For email address search.
    | Regular expression for email addresses.
    """
    re_mail = [r'[æøåÆØÅa-zA-Z0-9+._-]+@[æøåÆØÅa-zA-Z0-9._-]+\.[æøåÆØÅa-zA-Z0-9_-]+']
    return re_mail


def re_idNum_matcher() -> list:
    """
    | List of regular expressions for personal id numbers.
    | Match standard format for Nordic countries, Poland, UK, US.
    """
    re_idNum = [r'\b\d{11}\b',
                r'\b[a-ceghj-npr-tw-zA-CEGHJ-PR-TW-Z]{2}(?:\d){6}[a-dA-D]?\b',
                r'\b\d{3}\-\d{2}\-\d{4}\b', r'\b\d{11}\d', r'\b\d{6}\-\d{4}\b', r'\b\d{6}\-\d{3}[a-zA-Z]\b']
    return re_idNum


def re_cardNum_matcher() -> list:
    """
    | Regex for standard monetary card number format.
    """
    re_cardNum = [r'\b\d{4}\-\d{4}\-\d{4}\-\d{4}\b']  # include more!
    return re_cardNum


def name_finder(text, path):
    """
    | Use spaCy to find human names.
    """
    # Spacy model is loaded according to what language is detected by the state_language function.
    nlp = spacy.load(state_language(text))
    doc = nlp(text)
    per_li = []
    # Entities labeled as person names are added to the list per_li.
    for ent in doc.ents:
        if ent.label_ == "PERSON" or ent.label_ == 'PER':
            per_li.append(ent)
    # Spacy entities are converted to string and added to class Hits.
    per_li = [str(item) for item in per_li]
    per_li = list(set(per_li))
    per_li = sorted(per_li)
    for i in per_li:
        Hits_.Hits_li_names.append(i + ', ' + path)


def gps_coord(File_Name):
    """
    | Check for gps coordinates in image files.
    """
    file_name = File_Name
    pathpath = os.path.normpath(file_name)
    lat = 'gps_latitude'
    lng = 'gps_longitude'
    # Errors are raised when using exif.Image on png files. See exception list.
    try:
        with open(file_name, 'rb') as img_file:
            img = Image(img_file)

            if img.has_exif:
                if lat and lng in img.list_all():
                    h_lat = img.gps_latitude
                    h_lng = img.gps_longitude
                    hit = "Lat:" + str(h_lat), "Long:" + str(h_lng), pathpath
                    Hits_.Hits_li_gps.append(str(hit))
            else:
                pass
    except (OSError, ValueError, plum.exceptions.UnpackError):
        pass


# Extract text etc from xlsx file to search for given values.
def xlsx_reader(File_Name):
    """
    | Extract text from excel files for search/match process.
    """
    info_li = []
    file_name = File_Name
    pathpath = os.path.normpath(file_name)
    # Open xlsx file
    wb = load_workbook(file_name)
    # Read sheet
    ws = wb.active
    # Extract values from cells
    cells = (list(ws.rows))
    for cell in cells:
        for info in cell:
            if info.value != None:
                i = str(info.value)
                info_li.append(i)
    text = ' '.join(info_li)
    name_finder(text, pathpath)
    # Find email addresses.
    for i in re_mail_matcher():
        res = re.findall(i, text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_email.append(hit)
        else:
            continue
    # Find id numbers.
    for i in re_idNum_matcher():
        res = re.findall(i, text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_idNum.append(hit)
        else:
            continue
    # Find monetary card numbers.
    for i in re_cardNum_matcher():
        res = re.findall(i, text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_cardNum.append(hit)
        else:
            continue


def pdf_reader(file_name):
    """
    | Extract text from pdf files for search/match process.
    | Use pdfminer to extract text.
    """
    pathpath = os.path.normpath(file_name)
    pdf = file_name
    Text = extract_text(pdf)
    name_finder(Text, pathpath)
    # Find email addresses.
    for i in re_mail_matcher():
        ResSearch = re.findall(i.casefold(), Text.casefold())  # make case insensitive
        if ResSearch:
            for i in ResSearch:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_email.append(hit)
        else:
            continue
    # Find id numbers.
    for i in re_idNum_matcher():
        res = re.findall(i, Text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_idNum.append(hit)
        else:
            continue
    # Find monetary card numbers.
    for i in re_cardNum_matcher():
        res = re.findall(i, Text)
        if res:
            for i in res:
                hit = i + ', ' + pathpath
                Hits_.Hits_li_cardNum.append(hit)
        else:
            continue


def docx_reader(File_Name):
    """
    | Extract text from docx files for search/match process.
    """
    try:
        file_name = File_Name
        pathpath = os.path.normpath(file_name)
        doc = docx.Document(file_name)

        Text = []
        for para in doc.paragraphs:
            Text.append(para.text)
        Text = '\n'.join(Text)
        name_finder(Text, pathpath)
        Text = Text.casefold()
        # Find email addresses.
        for i in re_mail_matcher():
            res = re.findall(i, Text)
            if res:
                for i in res:
                    Hits_.Hits_li_email.append(i + ", " + pathpath)
            else:
                continue
        # Find id numbers.
        for i in re_idNum_matcher():
            res = re.findall(i, Text)
            if res:
                for i in res:
                    hit = i + ', ' + pathpath
                    Hits_.Hits_li_idNum.append(hit)
            else:
                continue
        # Find monetary card numbers.
        for i in re_cardNum_matcher():
            res = re.findall(i, Text)
            if res:
                for i in res:
                    hit = i + ', ' + pathpath
                    Hits_.Hits_li_cardNum.append(hit)
            else:
                continue
    except PermissionError:
        pass
    except docx.opc.exceptions.PackageNotFoundError:
        pass


def db_reader(File_Name):
    """
    | Connect to and read rows of database tables.
    """
    file_name = File_Name
    pathpath = os.path.normpath(file_name)

    # connect to db
    with sqlite3.connect(file_name) as connection:
        c = connection.cursor()
        for tables in c.execute("SELECT name FROM sqlite_master WHERE type='table';"):
            for table in tables:
                c.execute(f"SELECT * FROM {table}")
                Text = c.fetchall()
                Text = str(Text)
                name_finder(Text, pathpath)
                # Find email addresses.
                for i in re_mail_matcher():
                    res = re.findall(i, Text)
                    if res:
                        for i in res:
                            Hits_.Hits_li_email.append(i + ", " + pathpath)
                    else:
                        continue
                # Find id numbers.
                for i in re_idNum_matcher():
                    res = re.findall(i, Text)
                    if res:
                        for i in res:
                            hit = i + ', ' + pathpath
                            Hits_.Hits_li_idNum.append(hit)
                    else:
                        continue
                # Find monetary card numbers.
                for i in re_cardNum_matcher():
                    res = re.findall(i, Text)
                    if res:
                        for i in res:
                            hit = i + ', ' + pathpath
                            Hits_.Hits_li_cardNum.append(hit)
                    else:
                        continue


def read_file(File_Name):
    """
    | Standard file opener and reader.
    | Open and read files in byte-form (mode=rb).
    """
    file_name = File_Name
    pathpath = os.path.normpath(file_name)
    match_li = []
    fn = open(file_name, mode='r')
    tn = fn.read()
    name_finder(tn, pathpath)
    fn.close()
    f = open(file_name, mode='rb')
    t = f.read()

    # Find email addresses.
    for i in re_mail_matcher():
        i = i.encode()
        # Add to match_li if match is found
        res = re.findall(i, t, re.IGNORECASE)  # case sensitivity!!!
        # print(i, "is match for:", str(res), pathpath)
        if res:
            # print(res)
            for i in res:
                try:
                    i = i.decode('utf-8', 'backslashreplace')
                except Exception:
                    pass
                Hits_.Hits_li_email.append(str(i) + ", " + pathpath)
                # add_hit_to_li(str(i) + " !!! " + pathpath)
        else:
            continue
    # Find id numbers.
    for i in re_idNum_matcher():
        i = i.encode()
        res = re.findall(i, t, re.IGNORECASE)

        if res:
            for i in res:
                try:
                    i = i.decode('utf-8', 'backslashreplace')
                except Exception:
                    pass
                hit = str(i) + ', ' + pathpath
                Hits_.Hits_li_idNum.append(hit)
    # Find monetary card numbers.
    for i in re_cardNum_matcher():
        i = i.encode()
        res = re.findall(i, t, re.IGNORECASE)
        if res:
            for i in res:
                try:
                    i = i.decode('utf-8', 'backslashreplace')
                except Exception:
                    pass
                hit = i.decode() + ', ' + pathpath
                Hits_.Hits_li_cardNum.append(hit)

    f.close()


p1 = threading.Thread(target=progress_widget)  # Progress bar widget.


def walker():
    """
    | Walks directories and subdirectories to execute search.
    | Identify file type and extract text with appropriate function.
    """
    directory = select_dir()  # see select_dir() function.
    p1.daemon = True  # Allows python to exit even if thread is still running.
    p1.start()  # Starting Progress bar.

    for subdir, dirs, files in os.walk(directory):
        for file in files:
            # File-path from os
            paths = os.path.join(subdir, file)
            try:
                ftype = magic.from_file(paths, mime=True)
            except Exception as e:
                Hits_.Error_li.append(str(e) + ', ' + str(paths))
                pass

            # PDF files.
            if "pdf" in ftype:
                try:
                    pdf_reader(paths)
                except Exception as e:
                    Hits_.Error_li.append(str(e) + ', ' + str(paths))
                    pass

            # DOC and DOCX files.
            elif ftype == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                try:
                    docx_reader(paths)
                except Exception as e:
                    Hits_.Error_li.append(str(e) + ', ' + str(paths))
                    pass
                except docx.opc.exceptions.PackageNotFoundError:
                    pass
            # XLSX files.
            elif ftype == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                try:
                    xlsx_reader(paths)
                except Exception as e:
                    Hits_.Error_li.append(str(e) + ', ' + str(paths))
                    pass
            # PNG, JPG, JPEG files.
            elif ftype[:5] == 'image':
                try:
                    gps_coord(paths)
                except Exception as e:
                    Hits_.Error_li.append(str(e) + ', ' + str(paths))
                    pass
            # SQLITE DATABASE files.
            elif ftype == 'application/x-sqlite3':
                try:
                    db_reader(paths)
                except Exception as e:
                    Hits_.Error_li.append(str(e) + ', ' + str(paths))
                    pass

            # IF NONE OF THE TYPES ABOVE ARE IDENTIFIED, A GENERAL FILE OPENER IS USED (read_file function).
            else:
                try:
                    read_file(paths)
                except Exception as e:
                    Hits_.Error_li.append(str(e) + ', ' + str(paths))
                    pass


p2 = threading.Thread(target=walker)  # Directory walker (main function).


def main():
    p2.start()  # Starting main function.
    p2.join()  # Allows main function to finish before continuing to "exit_event.set()".
    exit_event.set()  # Sets exit_event to True in attempt at breaking loop of Progress bar.
    win.withdraw()  # Hides progress bar widget window.
    global stop
    stop = time.time()
    hits_to_file()  # Select path and save output.


# Starts script/program
if __name__ == '__main__':
    try:
        start = time.time()
        main()

        print(round(stop - start, 2))
    except Exception as e:
        Hits_.Error_li.append(str(e))
        pass
