from tkinter import *
from tkinter import filedialog
from pdf2docx import Converter
from pathlib import Path
from tkinter import messagebox
from datetime import datetime


# Funciones
def convertToDocx():
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Text files", "*.pdf*"),
                                                     ("all files", "*.*")))
    label_file_explorer.configure(text="File Opened: " + filename)
    global pdf_file
    pdf_file = Path(filename)


def executeConversion():
    if 'pdf_file' in globals():
        current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        docx_filename = pdf_file.stem + "_" + current_time + ".docx"
        docx_file = pdf_file.with_name(docx_filename)
        cv = Converter(str(pdf_file))
        cv.convert(str(docx_file))
        cv.close()
        messagebox.showinfo("Success", "Conversion completed successfully.")
    else:
        messagebox.showinfo("Error", "No PDF file selected.")

def version():
    messagebox.showinfo("Version","V 1.0.0.1 \\ Created by Yair Carvajal")

def do_nothing():
    pass

# Ventana principal
main = Tk()
main.title("PDF ConvertDoc")
main.geometry("430x270")
main.config(bg="#314252")
main.resizable(0, 0)

# Crear una etiqueta para mostrar el explorador de archivos
label_file_explorer = Label(main,
                            text="File Explorer using Tkinter",
                            width=50, height=2,
                            fg="blue")
label_file_explorer.place(x=40, y=50)

# Botones
button_browse = Button(main, text="Browse Files", command=convertToDocx)
button_browse.place(x=40, y=20)

button_execute = Button(main, text="Execute", command=executeConversion)
button_execute.place(x=40, y=90)

# Menú
menubar = Menu(main)
main.config(menu=menubar)

filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Version", command=version)
filemenu.add_command(label="Exit", command=main.destroy)

menubar.add_cascade(label="Archivo", menu=filemenu)

main.mainloop()