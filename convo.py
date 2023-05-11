from tkinter import filedialog, HORIZONTAL, scrolledtext
from os import listdir, makedirs
from os.path import isfile, join, abspath, splitext, exists
import win32com.client as win32
from win32com.client import constants
from tkinter.ttk import Progressbar
import tkinter as tk
import logging
import re
import shutil
from tkinter.messagebox import showinfo

# logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)
ch.setFormatter(formatter)
fileHandler = logging.FileHandler("convo.log")
fileHandler.setFormatter(formatter)
fileHandler.setLevel(logging.DEBUG)
logger.addHandler(ch)
logger.addHandler(fileHandler)


def convert_files(config, ui):
    word = None
    excel = None
    powerpoint = None

    dir_files = listdir(config.dir_path)
    # count max files
    count = len(list(filter(lambda f: splitext(f)[1] in (".doc",".ppt", ".xls"), dir_files)))
    step = 100/count
    ui.update()

    old_path = join(abspath(config.dir_path),"OLD FORMAT")
    if not exists(old_path):
        makedirs(old_path)

    for file in dir_files:
        filepath = join(abspath(config.dir_path),file)
        ext = splitext(file)[1]
        if isfile(filepath):
            try:
                message = ""
                if ext == ".doc":
                    if word is None:
                        word = win32.gencache.EnsureDispatch('Word.Application')
                        word.Visible = False
                    # https://learn.microsoft.com/en-us/office/vba/api/overview/word
                    message = "convert file {0} to docx".format(file, ext)
                    logger.info(message)
                    ui.write_line(message)
                    doc = word.Documents.Open(filepath)
                    doc.Activate()
                    new_file_abs = re.sub(r'\.\w+$', '.docx', abspath(filepath))
                    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
                    doc.Close(False)
                    shutil.move(filepath, old_path)
                elif ext == ".xls":
                    if excel is None: 
                        excel = win32.gencache.EnsureDispatch('Excel.Application')
                        excel.Visible = False
                    # https://learn.microsoft.com/en-us/office/vba/api/overview/excel
                    message = "convert file {0} to xlsx".format(file, ext)
                    logger.info(message)
                    ui.write_line(message)
                    book = excel.Workbooks.Open(filepath)
                    book.Activate()
                    excel.DisplayAlerts = False
                    new_file_abs = re.sub(r'\.\w+$', '.xlsx', abspath(filepath))
                    book.SaveAs(new_file_abs, FileFormat=constants.xlOpenXMLWorkbook)
                    book.Close()
                    shutil.move(filepath, old_path)
                elif ext == ".ppt":
                    if powerpoint is None:
                        powerpoint = win32.gencache.EnsureDispatch('PowerPoint.Application')
                    # https://learn.microsoft.com/en-us/office/vba/api/overview/powerpoint
                    message = "convert file {0} to pptx".format(file, ext)
                    logger.info(message)
                    ui.write_line(message)
                    pres = powerpoint.Presentations.Open(filepath)
                    new_file_abs = re.sub(r'\.\w+$', '.pptx', abspath(filepath))
                    pres.SaveAs(new_file_abs, FileFormat=constants.ppSaveAsOpenXMLPresentation)
                    pres.Close()
                    shutil.move(filepath, old_path)
            except:
                message = "conversion of file {0} failed".format(file)
                logger.info(message)
                ui.write_line(message)
                continue
            ui.register_progress(step)
    if word is not None:
        word.Quit()
    if excel is not None:
        excel.Quit()
    if powerpoint is not None:
        powerpoint.Quit()
    ui.show_info("Conversion process completed")

# class UiLoggingHandler(logging.StreamHandler):
#     def __init__(self, ui):
#         logging.StreamHandler.__init__(self)
#         self.ui = ui

#     def emit(self, record):
#         msg = self.format(record)
#         self.ui.write_line(msg)

class Config:
    def __init__(self):
        self.dir_path = "Folder not yet selected"
        self.progress = 0
        
class UI:
    def __init__(self, config):
        self.config = config

    def build(self):
        logger.info("building ui")
        self.r = tk.Tk()
        self.r.grid()
        self.r.title('Counting Seconds')
        quit_button = tk.Button(self.r, text='Exit', width=25, command=self.r.destroy)
        quit_button.grid(column=3, row=4, padx=10, pady=2)
        start_button = tk.Button(self.r, text='Start', width=25, command=self.convert)
        start_button.grid(column=3, row=2, padx=10, pady=2)
        select_button = tk.Button(self.r, text='Select Folder', width=25, command=self.select_dir)
        select_button.grid(column=3, row=1, padx=10, pady=2)

        self.dir_label = tk.Label(self.r, relief=tk.RAISED)
        self.dir_label.grid(column=0, row=1, padx=10, pady=10)
        self.dir_label.config(text=config.dir_path)

        self.progress_text = scrolledtext.ScrolledText(self.r, height = 5, width = 50)
        self.progress_text.grid(column=0, row=2, rowspan=4, padx=10, pady=10)

        self.bar = Progressbar(self.r, orient=HORIZONTAL, length=300, mode='determinate')

        self.r.mainloop()

    def select_dir(self):
        config.dir_path = filedialog.askdirectory()
        logger.info("directory select: {0}".format(config.dir_path))
        self.dir_label.config(text=config.dir_path)
        self.reset()

    def convert(self):
        self.bar.grid(column=0, row=0, columnspan=4,  padx=2, pady=2)
        convert_files(config, self)

    def update(self): 
        self.r.update()

    def register_progress(self, step):
        if self.bar['value'] < 100:
            self.bar['value'] += step
        self.bar.update()

    def show_info(self, message):
        showinfo(message=message)

    def write_line(self, message):
        self.progress_text.config(state=tk.NORMAL)
        self.progress_text.insert(tk.END, message + "\n")
        self.progress_text.config(state=tk.DISABLED)
        self.progress_text.update()

    def reset(self):
        self.progress_text.config(state=tk.NORMAL)
        self.progress_text.delete('1.0', tk.END)
        self.progress_text.config(state=tk.DISABLED)
        self.bar['value'] = 0
        self.bar.grid_remove()
        

if __name__ == '__main__':
    logger.debug("convo started")
    config = Config()
    ui = UI(config)
    ui.build()



