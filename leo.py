import tkinter as tk
from tkinter import ttk
from tkinter import Scrollbar
import speech_recognition as sr
import pyttsx3
import pyodbc
import re
from PIL import Image,ImageTk
import os

# تنظیمات گفتار
engine = pyttsx3.init()
engine.say("سلام، حال شما چطور است؟")
engine.runAndWait()
def speak(text):
    
    engine.say(text)
    engine.runAndWait()

    
# اتصال به دیتابیس Access
def get_db_connection():
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=C:\Users\Lenovo\Desktop\OnlineCalc\OC.accdb;'
    )
    return pyodbc.connect(conn_str)

# تشخیص فرمان از متن
def parse_command(text):
    text = text.lower()
    if match := re.search(r'قیمت\s+کد\s+(\d+)', text):
        return ('get_price', int(match.group(1)))
    elif 'گزارش فروش امروز' in text:
        return ('report_today',)
    if match := re.search(r'حساب\s+(.*)', text):
        return ('get_invc', match.group(1).strip())
    if match := re.search(r'فاکتورهای\s+(.*)', text):
        return ('get_invcdet', match.group(1).strip())
    return ('unknown',)

# اجرای فرمان
def handle_command(cmd):
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        #موجودی و قیمت کالا
        if cmd[0] == 'get_price':
            cursor.execute("SELECT Ourprice FROM Sale WHERE IDs=?", (cmd[1],))
            row = cursor.fetchone()
            msg = f"قیمت کالا {cmd[1]}، {row[0]} تومان است."
            output_label.config(text=msg)
            #نمایش تصویر کالا 
            try:
                pn=str(cmd[1])
                Image_path = os.path.join(r"C:/Users/Lenovo/Desktop/Temp/T1", (pn+".jpg"))
                if os.path.exists(Image_path):
                    img=Image.open(Image_path)
                    img=img.resize((500,500))
                    img_tk=ImageTk.PhotoImage(img)
                    image_label.config(image=img_tk)
                    image_label.image=img_tk
                else:
                    image_label.config(image='')
                    print("Image Not Found !")
            except Exception as e:
                print(f"Cant Open File:{e}")
        #گزارش ها
        elif cmd[0] == 'report_today':
            cursor.execute("SELECT SUM(Ourprice) FROM Sale WHERE odate = DATE()")
            total = cursor.fetchone()[0] or 0
            msg = f"مجموع فروش امروز {total} تومان است."
        #حساب اشخاص 
        elif cmd[0] == 'get_invc':
            cstmr_nam=cmd[1]

            cursor.execute("SELECT SUM(TotalP) FROM Sale WHERE byr LIKE ?", (f'%{cstmr_nam}%',))
            total=cursor.fetchone()[0] or 0
            row = cursor.fetchone()
            msg = f"حساب {cstmr_nam}: {total} تومان است."
            output_label.config(text=msg)
        #ریزحسابها 
        elif cmd[0] == 'get_invcdet':
            cstmr_nam=cmd[1]
            
            cursor.execute("SELECT IDs,Poname,Ourprice,TotalP,odate,byr FROM Sale WHERE byr LIKE ?", (f'%{cstmr_nam}%',))
            rows=cursor.fetchall()
            for row in tree.get_children():
                tree.delete(row)
            if not rows:
                output_label.config(text=f"هیچ فاکتوری یافت نشد")
            else:
                for row in rows:
                    tree.insert("","end",values=row)
    except Exception as e:
        output_label.config(text="خطا :{e}")
       
        speak(msg)
        output_label.config(text=msg)

    except Exception as e:
        speak("خطا در اجرای فرمان")
        output_label.config(text=f"خطا: {e}")

# ضبط و پردازش صدا
def listen_and_process():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        output_label.config(text="در حال گوش دادن...")
        audio = recognizer.listen(source)

    try:
        text = recognizer.recognize_google(audio, language="fa-IR")
        command_label.config(text="دستور: " + text)
        cmd = parse_command(text)
        handle_command(cmd)
    except:
        speak("متوجه نشدم، دوباره بگو.")
        output_label.config(text="متوجه نشدم.")

# رابط گرافیکی
root = tk.Tk()
root.title("هوش مصنوعی لـئـو")
screen_width=root.winfo_screenwidth()
screen_height=root.winfo_screenheight()
root.geometry(f"{screen_width}x{screen_height}")

#نمایش اطلاعات بر اساس لیست مرتب
tree=ttk.Treeview(root,columns=("کدکالا" ,"نام کالا" ,"فی" ,"جمع" ,"تاریخ" ,"خریدار"),show="headings",height=8)
tree.heading("کدکالا" ,text="کدکالا")
tree.heading("نام کالا" ,text="نام کالا")
tree.heading("فی" ,text="فی")
tree.heading("جمع" ,text="جمع")
tree.heading("تاریخ" ,text="تاریخ")
tree.heading("خریدار" ,text="خریدار")

tree.column("کدکالا",anchor="center",width=100,stretch=False)
tree.column("نام کالا",anchor="w",width=100 ,stretch=False)
tree.column("فی",anchor="center",width=100,stretch=False)
tree.column("جمع",anchor="center",width=100,stretch=False)
tree.column("تاریخ",anchor="center",width=100,stretch=False)
tree.column("خریدار",anchor="w",width=100,stretch=False)
tree.pack(pady=10)

#تعریف اسکرول بار
Scrollbar=Scrollbar(root,orient="vertical",command=tree.yview)
tree.configure(yscrollcommand=Scrollbar.set)
Scrollbar.pack(side="right",fill="y")

title_label = tk.Label(root, text=":) سلام من لـئـو هستم ", font=("B Nazanin", 16))
title_label.pack(pady=20)

record_button = tk.Button(root, text="ازم بپرس", command=listen_and_process, font=("B Nazanin", 14))
record_button.pack(pady=10)

command_label = tk.Label(root, text="", font=("B Nazanin", 12))
command_label.pack(pady=10)
#خروجی لیبل
output_label = tk.Label(root, text="", font=("B Nazanin", 12))
output_label.pack(pady=10)

image_label=tk.Label(root)
image_label.pack(pady=10)
root.mainloop()