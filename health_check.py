from tkinter import *
import tkinter.messagebox as msgbox
from openpyxl import load_workbook
import datetime
from openpyxl.styles import Font

wb = load_workbook("변화기록표.xlsx") # sample.xlsx 파일에서 wb 를 불러옴


root = Tk()
root.title("운동기록 프로그램")
root.geometry("640x480+500+200") # 가로 *세로+ x좌표 + y좌표
root.resizable(False, False) # x,y 값 변경 불가 (창 크기 변경 불가)

now = datetime.datetime.now()
now_day = (now.strftime('%Y-%m-%d'))

def start_btn():
    
    if pysical_value.get() == 1:
        ws1 = wb["가슴"]
        
        #[날짜] 입력
        col = "B"
        row = 4
        
        while ws1[f"{col}{row}"].value is not None:
            row += 1
                
        ws1[f"{col}{row}"] = now_day
        ws1[f"{col}{row}"].font = Font(name="Calibri", size=11)
        
        #[운동종류] 입력
        col = "C"
        pysical_name = name_input.get()
        ws1[f"{col}{row}"] = pysical_name
        ws1[f"{col}{row}"].font = Font(name="Calibri", size=11)
        name_input.delete(0, END)
        
        #[세트 수] 입력
        col = "E"
        pysical_setset = pysical_set_input.get()
        ws1[f"{col}{row}"] = pysical_setset
        ws1[f"{col}{row}"].font = Font(name="Calibri", size=11)
        pysical_set_input.delete(0, END)
        
        #[세트 휴식시간] 입력
        col = "F"
        pysical_set_time = set_time_input.get()
        ws1[f"{col}{row}"] = pysical_set_time
        ws1[f"{col}{row}"].font = Font(name="Calibri", size=11)
        set_time_input.delete(0, END)
        
        #[총 세트 수]
        col = "D"
        number = pysical_setset.count("+") + 1
        ws1[f"{col}{row}"] = number
        ws1[f"{col}{row}"].font = Font(name="Calibri", size=11)
        
    elif pysical_value.get() == 2:
        ws2 = wb["등"]
        
        #[날짜] 입력
        col = "B"
        row = 4
        
        while ws2[f"{col}{row}"].value is not None:
            row += 1
                
        ws2[f"{col}{row}"] = now_day
        ws2[f"{col}{row}"].font = Font(name="Calibri", size=11)
        
        #[운동종류] 입력
        col = "C"
        pysical_name = name_input.get()
        ws2[f"{col}{row}"] = pysical_name
        ws2[f"{col}{row}"].font = Font(name="Calibri", size=11)
        name_input.delete(0, END)
        
        #[세트 수] 입력
        col = "E"
        pysical_setset = pysical_set_input.get()
        ws2[f"{col}{row}"] = pysical_setset
        ws2[f"{col}{row}"].font = Font(name="Calibri", size=11)
        pysical_set_input.delete(0, END)
        
        #[세트 휴식시간] 입력
        col = "F"
        pysical_set_time = set_time_input.get()
        ws2[f"{col}{row}"] = pysical_set_time
        ws2[f"{col}{row}"].font = Font(name="Calibri", size=11)
        set_time_input.delete(0, END)
        
        #[총 세트 수]
        col = "D"
        number = pysical_setset.count('+') + 1
        ws2[f"{col}{row}"] = number
        ws2[f"{col}{row}"].font = Font(name="Calibri", size=11)
        
    elif pysical_value.get() == 3:
        ws3 = wb["하체"]
        
        #[날짜] 입력
        col = "B"
        row = 4
        
        while ws3[f"{col}{row}"].value is not None:
            row += 1
                
        ws3[f"{col}{row}"] = now_day
        ws3[f"{col}{row}"].font = Font(name="Calibri", size=11)
        
        #[운동종류] 입력
        col = "C"
        pysical_name = name_input.get()
        ws3[f"{col}{row}"] = pysical_name
        ws3[f"{col}{row}"].font = Font(name="Calibri", size=11)
        name_input.delete(0, END)
        
        #[세트 수] 입력
        col = "E"
        pysical_setset = pysical_set_input.get()
        ws3[f"{col}{row}"] = pysical_setset
        ws3[f"{col}{row}"].font = Font(name="Calibri", size=11)
        pysical_set_input.delete(0, END)
        
        #[세트 휴식시간] 입력
        col = "F"
        pysical_set_time = set_time_input.get()
        ws3[f"{col}{row}"] = pysical_set_time
        ws3[f"{col}{row}"].font = Font(name="Calibri", size=11)
        set_time_input.delete(0, END)

        #[총 세트 수]
        col = "D"
        number = pysical_setset.count("+") + 1
        ws3[f"{col}{row}"] = number
        ws3[f"{col}{row}"].font = Font(name="Calibri", size=11)
        
    wb.save("변화기록표.xlsx")
    msgbox.showinfo("알림", "정상적으로 작성 완료되었습니다.")
    


# 운동 종류
pysical_frame = Frame(root, relief="solid", bd=1)
pysical_frame_title = Label(pysical_frame, text="[운동 종류]")
pysical_frame.pack(side="top")
pysical_frame_title.pack()
pysical_value = IntVar()
pysical0 = Radiobutton(pysical_frame, text="가슴", value=1, variable=pysical_value)
pysical1 = Radiobutton(pysical_frame, text="등", value=2, variable=pysical_value)
pysical2 = Radiobutton(pysical_frame, text="하체", value=3, variable=pysical_value)
pysical0.pack()
pysical1.pack()
pysical2.pack()

# 운동 이름
name = Frame(root)
name_title = Label(name, text="[운동 종류]")
name_input = Entry(root, width=30)
name.pack(side="top")
name_title.pack(pady=3)
name_input.pack(pady=3)

# 세트당 횟수
pysical_set = Frame(root)
pysical_set_title = Label(pysical_set, text="[세트당 횟수]")
pysical_set_input = Entry(root, width=30)
pysical_set.pack(side="top")
pysical_set_title.pack(pady=3)
pysical_set_input.pack(pady=3)

# 세트 쉬는시간
set_time = Frame(root)
set_time_title = Label(set_time, text="[세트 쉬는시간]")
set_time_input = Entry(root, width=30)
set_time.pack(side="top")
set_time_title.pack(pady=3)
set_time_input.pack(pady=3)

# 버튼 프레임
button_frame = Frame(root)
button1 = Button(button_frame, text="실행", font=("", 15, "bold"), height=2, width=7, command=start_btn)
button_frame.pack(side="bottom", pady=3, padx=3, ipady=25)
button1.pack()

# my name
my_name_frame = Frame(root)
my_name_frame1 = Label(my_name_frame, text="만든이 : I_enable")
my_name_frame.place(relx=0.80, rely=0.93)
my_name_frame1.pack()


root.mainloop()