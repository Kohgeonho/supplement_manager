import pandas as pd
import tkinter.ttk as ttk
import tkinter
from tkinter import *
from tkinter import filedialog
from openpyxl import load_workbook

# Tool 창 설정
tool = Tk()
tool.title("보충보급 관리 프로그램")
tool.geometry("450x350+600+300")    # 가로 * 세로 + x좌표 + y좌표
tool.resizable(False, False)    # 높이, 너비 값 변경 불가
tool.iconphoto(False, PhotoImage(file='icon.gif'))


# 정보 입력
dataframe = LabelFrame(tool, text="정보 입력", font="맑은고딕 18 bold")
dataframe.pack(side="top", pady=10, ipady=5)


# 결과 조회
resultframe = LabelFrame(tool, text="결과 조회", font="맑은고딕 18 bold")
resultframe.pack(side="bottom", pady=30, ipady=5)


# 입대 년 월
dateframe = LabelFrame(dataframe, text="입대 년 월", font="맑은고딕 12")
dateframe.grid(row=0, column=0, padx=10, pady=5)

years = [str(i) + "년" for i in range(2020, 2030)]
yearbox = ttk.Combobox(dateframe, width=8, height=10, values=years, state="readonly", font="맑은고딕 12")
yearbox.current(0)    # 0번째 인덱스 값 선택
yearbox.pack(side="left", padx=3, pady=5)

months = [str(i) + "월" for i in range(1, 13)]
monthbox = ttk.Combobox(dateframe, width=8, height=12, values=months, state="readonly", font="맑은고딕 12")
monthbox.current(0)
monthbox.pack(side="right", padx=3, pady=5)

def msgbox(title, info):
    tkinter.messagebox.showinfo(title, info)
def errorbox(title, info):
    tkinter.messagebox.showerror(title, info)

def HR_result():

    # 파일 선택
    def select_file():
        files = filedialog.askopenfilename(title="파일을 선택하세요", \
                                           filetypes=(("Excel 파일", "*.xlsx"), \
                                                      ("모든 파일", "*.*")), \
                                           initialdir="./")    # 현재 경로를 보여줌
        df = pd.read_excel(files)
        df2 = pd.DataFrame()
        df2['부대'] = df['Unnamed: 0']
        df2['군번'] = df['Unnamed: 26']
        df2['이름'] = df['Unnamed: 33']
        df3 = df2.dropna(how='any')
        return df3

    def save_data(df):
        year = yearbox.get()
        month = monthbox.get()

        writer = pd.ExcelWriter('member_info.xlsx', mode='a')
        df.to_excel(writer, sheet_name=year+month, index=False)
        writer.save()

    try:
        save_data(select_file())
        msgbox("인사 명령 결과", "완료되었습니다.")
    except PermissionError:
        errorbox("Permission Error", "접근 권한이 없습니다. \n 파일을 닫고 다시 진행해주십시오.")
    except AssertionError:
        errorbox("Assertion Error", "파일 선택이 취소되었습니다.")


# 최초 인사 명령 결과 버튼
btn1 = Button(dataframe, width=20, height=2, text="최초 인사 명령 결과", font="맑은고딕 12", command=HR_result)
btn1.grid(row=0, column=1, padx=10, pady=5)


# 피복 사이즈 정보 버튼
btn2 = Button(dataframe, width=20, height=2, text="피복 사이즈 정보", font="맑은고딕 12")
btn2.grid(row=1, column=0, padx=10, pady=5)


# 최종 부대 분류 결과 버튼
btn3 = Button(dataframe, width=20, height=2, text="최종 부대 분류 결과", font="맑은고딕 12")
btn3.grid(row=1, column=1, padx=10, pady=5)


# 결과 조회
def show_data():
    data = Tk()
    data.title("결과 조회")
    data.mainloop()


# 품목별 조회 버튼
btn4 = Button(resultframe, width=20, height=2, text="품목별 조회", font="맑은고딕 12", command=show_data)
btn4.grid(row=0, column=0, padx=10, pady=5)


# 부대별 조회 버튼
btn5 = Button(resultframe, width=20, height=2, text="부대별 조회", font="맑은고딕 12", command=show_data)
btn5.grid(row=0, column=1, padx=10, pady=5)

tool.mainloop()
