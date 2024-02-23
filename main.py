from customtkinter import CTk,CTkButton
from tkinter import filedialog,Menu
from  sheet2dict import Worksheet
from time import sleep
from pywhatkit import sendwhatmsg_instantly

main_widget=CTk()
main_widget.geometry("500x500")



def sendmessage():
    file = filedialog.askopenfile()
    fpath = file.name
    w = Worksheet()

    student_data = w.xlsx_to_dict(path=fpath)

    lst = student_data.sheet_items

    for row in lst:
        phone = "+" + row["Phone Numbers"]
        row.pop("Phone Numbers")
        sendwhatmsg_instantly(phone,    f"{row}",tab_close=True)
        sleep(5)



send_button=CTkButton(main_widget,text="Send",command=sendmessage)



send_button.pack(expand=True)
main_widget.mainloop()
