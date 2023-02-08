from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
from pdf2docx import Converter
import docx2pdf
import img2pdf
# from pdf2image import convert_from_path


def pdfprotector():
    def browse():
        global filename
        filename = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              title="Select image file",
                                              filetype=(('PDF File', '*.pdf'), ('all files', '*.*')))
        entry1.insert(0, filename)

    def protect():

        mainfile = source.get()
        protect_file = target.get()
        code = password.get()
        if mainfile == "" and protect_file == "" and code == "":
            messagebox.showerror("Invalid", " All entries are empty !!!")
        elif mainfile == "":
            messagebox.showerror("Invalid", " mainfile is empty !!!")
        elif protect_file == "":
            messagebox.showerror("Invalid", " protectfile is empty !!!")
        elif code == "":
            messagebox.showerror("Invalid", " code is empty !!!")
        else:
            try:
                out = PdfFileWriter()
                file = PdfFileReader(filename)
                num = file.numPages

                for idx in range(num):
                    page = file.getPage(idx)
                    out.addPage(page)
                out.encrypt(code)
                with open(protect_file, "wb") as f:
                    out.write(f)

                messagebox.showinfo("valid", " SUCCESSFULLY PROTECT YOUR FILE !!!")
                pdf_protector.destroy()

            except:
                source.set("")
                target.set("")
                password.set("")
                messagebox.showerror("Invalid", "Invalid entry !!!")
    messagebox.showinfo("welcome", "Now you can Protect PDF")
    pdf_protector = Toplevel(root)
    pdf_protector.title("PDF PROTECTOR")
    pdf_protector.wm_iconbitmap("pdf.ico")
    pdf_protector.geometry("800x500")
    font_tuple = ("Comic Sans MS", 35, "bold")
    pdf_protector_label = Label(pdf_protector, text="PDF PROTECTOR", fg="purple")
    pdf_protector_label.place(x=50, y=40)
    pdf_protector_label.configure(font=font_tuple)
    source = StringVar()
    label1 = Label(pdf_protector, text="PDF  File Name :-")
    label1.place(x=100, y=150)
    entry1 = Entry(pdf_protector, textvariable=source, font="Helvetica",width=50)
    entry1.place(x=200, y=150)
    button = Button(pdf_protector, text="Browse", command=browse, bg="blue", fg="white", font="Helvetica 10 bold")
    button.place(x=700, y=150)

    target = StringVar()
    label2 = Label(pdf_protector, text="Target File :-", font="Helvetica 10")
    label2.place(x=100, y=200)
    entry2 = Entry(pdf_protector, textvariable=target, font="Helvetica ")
    entry2.place(x=200, y=200)

    password = StringVar()
    label3 = Label(pdf_protector, text="Password :-", font="Helvetica ")
    label3.place(x=100, y=250)
    entry3 = Entry(pdf_protector, textvariable=password, font="Helvetica ")
    entry3.place(x=200, y=250)

    button1 = Button(pdf_protector, text="PROTECT", command=protect, bg="blue", fg="white", font="Helvetica 10 bold")
    button1.place(x=150, y=300)


def pdf2word():
    def browse2():
        global pdf_file2
        pdf_file2 = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              title="Select image file",
                                              filetype=(('PDF File', '*.pdf'), ('all files', '*.*')))
        nameentry.insert(0, pdf_file2)
    def submit2():

        word_file2 = wordname2value.get()
        print(pdf_file2)
        if word_file2 == "":
            messagebox.showerror("Information", "Something Wents Wrong")
            word_file2.destroy()
        else:
            try:
                cv = Converter(pdf_file2)
                cv.convert(word_file2, start=0, end=None)
                cv.close()
                messagebox.showinfo("Information", "PDF TO WORD CONVERSION IS SUCCESSFULLY DONE")
                pdf_word.destroy()

            except:
                messagebox.showerror("Info", "Somethings went wrong")


    messagebox.showinfo("welcome", "Now you can convert PDF TO WORD")
    pdf_word = Toplevel(root)
    pdf_word.title("PDF TO WORD")
    pdf_word.wm_iconbitmap("pdf.ico")
    pdf_word.geometry("800x500")
    font_tuple = ("Comic Sans MS", 35, "bold")
    pdf_word_label = Label(pdf_word, text="PDF TO WORD CONVERTOR", fg="purple", font="Helvetica 30 bold ")
    pdf_word_label.place(x=50, y=70)
    pdf_word_label.configure(font=font_tuple)
    pdfname2 = Label(pdf_word, text="PDF File Name :-", font="Helvetica 15 ")
    pdfname2.place(x=50, y=180)
    wordname2 = Label(pdf_word, text="WORD File Name :-", font="Helvetica 15 ")
    wordname2.place(x=50, y=260)

    pdfname2value = StringVar()
    wordname2value = StringVar()

    nameentry = Entry(pdf_word, textvariable=pdfname2value, width=50, font='Helvetica 12')
    nameentry.place(x=240, y=180)
    aadharentry = Entry(pdf_word, textvariable=wordname2value, width=25, font='Helvetica 12')
    aadharentry.place(x=240, y=260)

    button = Button(pdf_word, text="Browse", command=browse2, bg="blue", fg="white", font="Helvetica 10 bold")
    button.place(x=720, y=180)

    submitbutton = Button(pdf_word, text="UPLOAD", command=submit2,
                          width=13, bg="blue", fg="white", font="Helvetica 15 bold")
    submitbutton.place(x=300, y=350)


def word2pdf():
    def browse3():
        global word_file3
        word_file3 = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              title="Select image file",
                                              filetype=(('WORD File', '*.docx'), ('all files', '*.*')))
        wordname3.insert(0, word_file3)
    def submit3():
        pdf_file3 = pdfname3value.get()
        if pdf_file3 =="":
            messagebox.showerror("Information", "Something Wents Wrong")
            word_pdf.destroy()
        else:
            try:
                docx2pdf.convert(word_file3, pdf_file3)
                messagebox.showinfo("Information", "WORD TO PDF CONVERSION IS SUCCESSFULLY DONE")
                word_pdf.destroy()
            except:
                messagebox.showerror("Information", "Something Wents Wrong")

    messagebox.showinfo("welcome", "Now you can convert WORD TO PDF")
    word_pdf = Toplevel(root)
    word_pdf.title("WORD TO PDF")
    word_pdf.wm_iconbitmap("pdf.ico")
    word_pdf.geometry("800x500")
    font_tuple = ("Comic Sans MS", 35, "bold")
    word_pdf_label = Label(word_pdf, text="WORD TO PDF CONVERTOR", fg="purple")
    word_pdf_label.place(x=50, y=70)
    word_pdf_label.configure(font=font_tuple)
    wordname3 = Label(word_pdf, text="WORD File Name :-", font="Helvetica 15 ")
    wordname3.place(x=50, y=180)
    pdfname3 = Label(word_pdf, text="PDF File Name :-", font="Helvetica 15 ")
    pdfname3.place(x=50, y=260)

    pdfname3value = StringVar()
    wordname3value = StringVar()

    wordname3 = Entry(word_pdf, textvariable=wordname3value, width=50, font='Helvetica 12')
    wordname3.place(x=240, y=180)
    pdfname3 = Entry(word_pdf, textvariable=pdfname3value, width=25, font='Helvetica 12')
    pdfname3.place(x=240, y=260)

    button = Button(word_pdf, text="Browse", command=browse3, bg="blue", fg="white", font="Helvetica 10 bold")
    button.place(x=720, y=180)

    submitbutton = Button(word_pdf, text="UPLOAD", command=submit3,
                          width=15, bg="blue", fg="white", font="Helvetica 15 bold")
    submitbutton.place(x=300, y=350)


def img_pdf():
    def browse4():
        global word_file3
        img_file4 = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              title="Select image file",
                                              filetype=(('JPG File', '*.jpg'), ('all files', '*.*')))
        imgname4.insert(0, img_file4)
    def submit4():
        img_file4 = imgname4value.get()
        pdf_file4 = pdfname4value.get()
        try:
            with open(pdf_file4, 'wb') as f:
                f.write(img2pdf.convert(img_file4))
                messagebox.showinfo("Information", "IMG TO PDF CONVERSION IS SUCCESSFULLY DONE")
                imgpdf.destroy()
        except:
            messagebox.showerror("Information", "Somethings wents wrong")

    messagebox.showinfo("welcome", "Now you can convert IMG TO PDF")
    imgpdf = Toplevel(root)
    imgpdf.title("IMG TO PDF")
    imgpdf.wm_iconbitmap("pdf.ico")
    imgpdf.geometry("800x500")
    font_tuple = ("Comic Sans MS", 35, "bold")
    imgpdf_label = Label(imgpdf, text="IMG TO PDF CONVERTOR", fg="purple", font="Helvetica 30 bold ")
    imgpdf_label.place(x=50, y=70)
    imgpdf_label.configure(font=font_tuple)
    imgname4 = Label(imgpdf, text="IMG File Name :-", font="Helvetica 15 ")
    imgname4.place(x=50, y=180)
    pdfname4 = Label(imgpdf, text="PDF File Name :-", font="Helvetica 15 ")
    pdfname4.place(x=50, y=260)

    imgname4value = StringVar()
    pdfname4value = StringVar()

    imgname4 = Entry(imgpdf, textvariable=imgname4value, width=50, font='Helvetica 12')
    imgname4.place(x=240, y=180)
    pdfname4 = Entry(imgpdf, textvariable=pdfname4value, width=25, font='Helvetica 12')
    pdfname4.place(x=240, y=260)

    button = Button(imgpdf, text="Browse", command=browse4, bg="blue", fg="white", font="Helvetica 10 bold")
    button.place(x=720, y=180)

    submitbutton = Button(imgpdf, text="UPLOAD", command=submit4,
                          width=15, bg="blue", fg="white", font="Helvetica 15 bold")
    submitbutton.place(x=300, y=350)


def pdf2img():
    def browse5():
        global pdf_file5
        pdf_file5 = filedialog.askopenfilename(initialdir=os.getcwd(),
                                              title="Select image file",
                                              filetype=(('PDF File', '*.pdf'), ('all files', '*.*')))
        pdfname5.insert(0, pdf_file5)

    def submit5():
        img_file5 = imgname5value.get()
        pdf_file5 = pdfname5value.get()
        print(pdf_file5)

        if img_file5 == "":
            messagebox.showerror("Information", "Please fill img name")
            pdf_img.destroy()
        else:
            image = convert_from_path(pdf_file5)
            for i in image:
                i.save(f"{img_file5}.jpg", "JPEG")
            messagebox.showinfo("Information", "PDF TO IMG CONVERSION IS SUCCESSFULLY DONE")
            pdf_img.destroy()
            # try:
            #     image = convert_from_path(pdf_file5)
            #     for i in image:
            #         i.save(f"{img_file5}.jpg", "JPEG")
            #     messagebox.showinfo("Information", "PDF TO IMG CONVERSION IS SUCCESSFULLY DONE")
            #     pdf_img.destroy()
            #
            #
            # except:
            #     messagebox.showerror("Information", "Somethings wents wrong")

    messagebox.showinfo("welcome", "Now you can convert PDF TO IMG")
    pdf_img = Toplevel(root)
    pdf_img.title("PDF TO IMG")
    root.wm_iconbitmap("pdf.ico")
    pdf_img.geometry("800x500")
    font_tuple = ("Comic Sans MS", 35, "bold")
    pdf_img_label = Label(pdf_img, text="PDF TO IMG CONVERTOR", fg="purple", font="Helvetica 30 bold ")
    pdf_img_label.place(x=50, y=70)
    pdf_img_label.configure(font=font_tuple)
    pdfname5 = Label(pdf_img, text="PDF File Name :-", font="Helvetica 15 ")
    pdfname5.place(x=50, y=180)
    wordname5 = Label(pdf_img, text="IMG File Name :-", font="Helvetica 15 ")
    wordname5.place(x=50, y=260)

    imgname5value = StringVar()
    pdfname5value = StringVar()

    pdfname5 = Entry(pdf_img, textvariable=pdfname5value, width=50, font='Helvetica 12')
    pdfname5.place(x=240, y=180)
    imgname5 = Entry(pdf_img, textvariable=imgname5value, width=25, font='Helvetica 12')
    imgname5.place(x=240, y=260)
    warning5 = Label(pdf_img, text="Name without jpg ", font="Helvetica 10 ")
    warning5.place(x=300, y=290)

    button = Button(pdf_img, text="Browse", command=browse5, bg="blue", fg="white", font="Helvetica 10 bold")
    button.place(x=720, y=180)

    submitbutton = Button(pdf_img, text="UPLOAD", command=submit5,
                          width=15, bg="blue", fg="white", font="Helvetica 15 bold")
    submitbutton.place(x=300, y=350)


root = Tk()
root.geometry("1310x700")
root.title("PDF MANAGEMENT")
root.wm_iconbitmap("pdf.ico")
photo = PhotoImage(file="frontpage.png")
root.config(highlightbackground="black", highlightthickness=5)
label = Label(image=photo)
label.pack(side=TOP, fill=BOTH, expand=1)
root.resizable(FALSE,FALSE)

protectorbutton = Button(root, text="PDF PROTECTOR", command=pdfprotector,
                         width=15, bg="blue", fg="white", font="Helvetica 15 bold")
protectorbutton.place(x=30, y=500)

pdf2wordbutton = Button(root, text="PDF TO WORD", command=pdf2word,
                        width=15, bg="blue", fg="white", font="Helvetica 15 bold")
pdf2wordbutton.place(x=260, y=500)

word2pdfbutton = Button(root, text="WORD TO PDF", command=word2pdf,
                        width=15, bg="blue", fg="white", font="Helvetica 15 bold")
word2pdfbutton.place(x=510, y=500)

img2pdfbutton = Button(root, text="IMG TO PDF", command=img_pdf,
                       width=15, bg="blue", fg="white", font="Helvetica 15 bold")
img2pdfbutton.place(x=780, y=500)

pdf2imgbutton = Button(root, text="PDF TO IMG", command=pdf2img,
                       width=15, bg="blue", fg="white", font="Helvetica 15 bold")
pdf2imgbutton.place(x=1050, y=500)

root.mainloop()
