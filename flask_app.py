# app.py (Flask 애플리케이션)
from flask import Flask, render_template, request, send_file
import xlrd
import pandas as pd
import re
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)

df = None  # DataFrame을 저장할 변수
test = ""

def save_uploaded_file(file):
    if file:
        filename = secure_filename(file.filename)
        file_path = filename
        file.save(file_path)
        return file_path
    return None

@app.route("/")
def index():
    global df
    if df is None:
        df = pd.DataFrame()
    return render_template("index.html", df=df)


@app.route("/", methods=["POST"])
def filter_data():
    global df
    global test

    # 엑셀 파일 업로드 처리
    uploaded_file = request.files["netlist_file"]
    if uploaded_file.filename != "":
        file_path = save_uploaded_file(uploaded_file)  # 파일을 저장하고 경로를 얻음
        test = file_path
        if file_path:
            data = open(file_path, "rb").read()  # 파일 내용 읽음
            workbook = xlrd.open_workbook(file_contents=data)
            workbook = xlrd.open_workbook(file_contents=data)
            worksheet = workbook.sheet_by_index(0)

            # 입력 필터 조건 받기
            des1 = request.form.get("date1")
            des6 = request.form.get("date2")
            des2 = request.form.get("time")
            des3 = request.form.get("group")
            des4 = request.form.get("installer")
            des5 = request.form.get("str")
            des7 = request.form.get("apart")
            # 데이터 필터링을 위한 리스트
            data_list = []

            for row in range(worksheet.nrows):
                row_data = worksheet.row_values(row)

                # 날짜 필터링
                date_value = row_data[5]
                if des1 and des6:
                    if isinstance(date_value, float):
                        date_value = xlrd.xldate_as_tuple(date_value, workbook.datemode)
                        year, month, day, _, _, _ = date_value
                        date_str = f"{year}-{month:02d}-{day:02d}"
                    else:
                        date_parts = date_value.split()
                        if date_parts:
                            date_str = date_parts[0]
                        else:
                            # 날짜 형식이 아니면 continue
                            continue

                    if not (des1 <= date_str <= des6):
                        continue

                # 타임 필터링
                if des2 and des2 not in row_data[9]:
                    continue

                # 그룹 필터링 
                if des3:
                    if des3 == "영남" and "영남" not in row_data[3]:
                        continue
                    elif des3 == "서울" and "영남" in row_data[3]:
                        continue

                # 시공자 필터링
                if des4 and des4 not in row_data[7]:
                    continue
                # 상태 필터링
                if des5 and des5 not in row_data[7]:
                    continue
                # 그룹 필터링
                if des7 and des7 not in row_data[3]:
                    continue
                # 데이터 추가
                data_list.append(row_data)
            

            # 결과 출력
            df = pd.DataFrame(data_list)
            df = df.iloc[:, 2:] 
            df = df[df.iloc[:, 0] != '이름']
            df.iloc[:, 5] = df.iloc[:, 5].apply(lambda x: re.sub(r'^(.*?)-(.*)$', r'(\1)\2', x))

            
            new_row = ["고객명", "아파트명", "계약조건", "시공일", "비고", "시공자", "연락처","추가사항",
                    "비고","주소","동/호수","입력시간","기타사항"]
            df.loc[-1] = new_row
            df.index = df.index + 1 
            df = df.sort_index()

            FileName = test.replace("uploads\\", '')

            return render_template("index.html", df=df, FileName=FileName)
    else:
        if test:
            data = open(test, "rb").read()  # 파일 내용 읽음
            workbook = xlrd.open_workbook(file_contents=data)
            workbook = xlrd.open_workbook(file_contents=data)
            worksheet = workbook.sheet_by_index(0)
        
                    # 입력 필터 조건 받기
            des1 = request.form.get("date1")
            des6 = request.form.get("date2")
            des2 = request.form.get("time")
            des3 = request.form.get("group")
            des4 = request.form.get("installer")
            des5 = request.form.get("str")
            des7 = request.form.get("apart")
            # 데이터 필터링을 위한 리스트
            data_list = []

            for row in range(worksheet.nrows):
                row_data = worksheet.row_values(row)

                # 날짜 필터링
                date_value = row_data[5]
                if des1 and des6:
                    if isinstance(date_value, float):
                        date_value = xlrd.xldate_as_tuple(date_value, workbook.datemode)
                        year, month, day, _, _, _ = date_value
                        date_str = f"{year}-{month:02d}-{day:02d}"
                    else:
                        date_parts = date_value.split()
                        if date_parts:
                            date_str = date_parts[0]
                        else:
                            # 날짜 형식이 아니면 continue
                            continue

                    if not (des1 <= date_str <= des6):
                        continue

                # 타임 필터링
                if des2 and des2 not in row_data[9]:
                    continue

                # 그룹 필터링 
                if des3:
                    if des3 == "영남" and "영남" not in row_data[3]:
                        continue
                    elif des3 == "서울" and "영남" in row_data[3]:
                        continue

                # 시공자 필터링
                if des4 and des4 not in row_data[7]:
                    continue
                # 상태 필터링
                if des5 and des5 not in row_data[7]:
                    continue
                # 그룹 필터링
                if des7 and des7 not in row_data[3]:
                    continue
                # 데이터 추가
                data_list.append(row_data)
            

            # 결과 출력
            df = pd.DataFrame(data_list)
            df = df.iloc[:, 2:] 
            df = df[df.iloc[:, 0] != '이름']
            df.iloc[:, 5] = df.iloc[:, 5].apply(lambda x: re.sub(r'^(.*?)-(.*)$', r'(\1)\2', x))

            
            new_row = ["고객명", "아파트명", "계약조건", "시공일", "비고", "시공자", "연락처","추가사항",
                    "비고","주소","동/호수","입력시간","기타사항"]
            df.loc[-1] = new_row
            df.index = df.index + 1 
            df = df.sort_index()
            FileName = test.replace("uploads\\", '')

            return render_template("index.html", df=df, FileName=FileName)

@app.route("/download_excel")
def download_excel():
    global df

    excel_filename = "output.xlsx"
    df.to_excel(excel_filename, index=False)

    return send_file(excel_filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
