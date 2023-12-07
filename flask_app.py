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
testlist = ["태훈","용균","김경","이인","박찬", "임효", "용준","정병","안심","인효","권병","광주","조근","오수","홍진","이천","시흥","평택"]
def save_uploaded_file(file):
    if file:

        filename = secure_filename(file.filename)
        file_path = filename
        file.save(file_path)
        return file_path
    return None

# def extract_unique_prefix(data_list, column_index):
#     prefixes = set()
#     for row_data in data_list:
#         value = row_data[column_index]
#         if "-" in value:
#             prefix = value.split("-")[0]
#             prefixes.add(prefix)
#     sorted_prefixes = sorted(prefixes)
#     return sorted_prefixes

def extract_unique_prefix_once(worksheet, column_index):
    prefixes = set()
    for row in range(worksheet.nrows):
        row_data = worksheet.row_values(row)
        value = row_data[column_index]
        if "-" in value:
            prefix = value.split("-")[0]
            prefixes.add(prefix)
    sorted_prefixes = sorted(prefixes)
    return sorted_prefixes
def extract_unique_prefix_two(worksheet, column_index):
    prefixes = set()
    for row in range(worksheet.nrows):
        row_data = worksheet.row_values(row)
        value = row_data[column_index]
        if "-" in value:
            prefix = value.split("-")[1]
            prefixes.add(prefix)
    sorted_prefixes = sorted(prefixes)
    return sorted_prefixes

def extract_unique_prefix_three(worksheet, column_index):
    prefixes = set()
    for row in range(worksheet.nrows):
        row_data = worksheet.row_values(row)
        value = row_data[column_index]
        if "-" not in value:
            prefixes.add(value)
    sorted_prefixes = sorted(prefixes)
    return sorted_prefixes

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
            des8 = request.form.get("search")

            # 데이터 필터링을 위한 리스트
            data_list = []

            for row in range(worksheet.nrows):
                row_data = worksheet.row_values(row)
                available_str_options = extract_unique_prefix_once(worksheet, 7)
                available_str_options2 = extract_unique_prefix_two(worksheet, 7)
                available_str_options3 = extract_unique_prefix_three(worksheet, 7)

                str_options_html = "\n".join(
                    f'<option value="{option}">{option}</option>' for option in available_str_options
                )
                str_options_html2 = "\n".join(
                    f'<option value="{option}">{option}</option>' for option in available_str_options2
                )
                str_options_html3 = "\n".join(
                    f'<option value="{option}">{option}</option>' for option in available_str_options3
                )
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
                    if des3 == "영남" and "★" not in row_data[3]:
                        continue
                    elif des3 == "서울" and ("★" in row_data[3] or "■" in row_data[3]):
                        continue
                    elif des3 == "특별관리" and ("★" in row_data[3] or "☆" in row_data[3]):
                        continue


                # 시공자 필터링
                if des4 and des4 not in row_data[7]:
                    continue
                # 상태 필터링
                if des5:
                    if des5 in testlist and not any(row_data[7].startswith(des5 + "-") for des5 in testlist):
                        continue
                    elif des5 not in row_data[7]:
                        continue

                # 그룹 필터링
                if des7 and des7 not in row_data[3]:
                    continue

                # 검색 필터링
                if des8 and des8 not in row_data[2] and des8 not in row_data[8]:
                    continue



                # 데이터 추가
                data_list.append(row_data)

            if data_list:
                # 결과 출력
                df = pd.DataFrame(data_list)
                df = df.iloc[:, 2:]
                df = df[df.iloc[:, 0] != '이름']
                df.iloc[:, 1] = df.iloc[:, 1].apply(lambda x: x.replace('=', '') if '=' in x else x)

                if not df.empty:
                    new_row = ["고객명", "아파트명", "계약조건", "시공일", "비고", "시공자", "연락처", "시간",
                            "비고", "주소", "동/호수", "입력시간", "기타사항"]
                    df.loc[-1] = new_row
                    df.index = df.index + 1
                    df = df.sort_index()
                else:
                    # 데이터가 없는 경우 빈 DataFrame 생성
                    df = pd.DataFrame(columns=["고객명", "아파트명", "계약조건", "시공일", "비고", "시공자", "연락처", "시간",
                                            "비고", "주소", "동/호수", "입력시간", "기타사항"])

                FileName = test.replace("uploads\\", '')


                return render_template("index.html", df=df, FileName=FileName, test=test, filter_conditions={
                    "date1": des1, "date2": des6,
                    "time": des2, "group": des3,
                    "installer": des4, "str": des5,
                    "apart": des7
                }, str_options_html=str_options_html, str_options_html2=str_options_html2,str_options_html3=str_options_html3)
            else:
                alert_message = "조건에 맞는 데이터가 없습니다"  # 알림 메시지 설정
                return render_template("index.html", df=None, alert_message=alert_message)

    else:
        if test:
            data = open(test, "rb").read()  # 파일 내용 읽음
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
            des8 = request.form.get("search")
            # 데이터 필터링을 위한 리스트
            data_list = []
            for row in range(worksheet.nrows):
                row_data = worksheet.row_values(row)
                available_str_options = extract_unique_prefix_once(worksheet, 7)
                available_str_options2 = extract_unique_prefix_two(worksheet, 7)
                available_str_options3 = extract_unique_prefix_three(worksheet, 7)

                str_options_html = "\n".join(
                    f'<option value="{option}">{option}</option>' for option in available_str_options
                )
                str_options_html2 = "\n".join(
                    f'<option value="{option}">{option}</option>' for option in available_str_options2
                )
                str_options_html3 = "\n".join(
                    f'<option value="{option}">{option}</option>' for option in available_str_options3
                )
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
                    if des3 == "영남" and "★" not in row_data[3]:
                        continue
                    elif des3 == "서울" and ("★" in row_data[3] or "■" in row_data[3]):
                        continue
                    elif des3 == "특별관리" and ("★" in row_data[3] or "☆" in row_data[3]):
                        continue


                # 시공자 필터링
                if des4 and des4 not in row_data[7]:
                    continue
                # 상태 필터링
                if des5:
                    if des5 in testlist and not any(row_data[7].startswith(des5 + "-") for des5 in testlist):
                        continue
                    elif des5 not in row_data[7]:
                        continue


                # 그룹 필터링
                if des7 and des7 not in row_data[3]:
                    continue

                # 검색 필터링
                # 검색 필터링
                if des8 and des8 not in row_data[2] and des8 not in row_data[8]:
                    continue


                # 데이터 추가
                data_list.append(row_data)

            
            if data_list:
                # 결과 출력
                df = pd.DataFrame(data_list)
                df = df.iloc[:, 2:]
                df = df[df.iloc[:, 0] != '이름']
                df.iloc[:, 1] = df.iloc[:, 1].apply(lambda x: x.replace('=', '') if '=' in x else x)

                if not df.empty:
                    new_row = ["고객명", "아파트명", "계약조건", "시공일", "비고", "시공자", "연락처", "시간",
                            "비고", "주소", "동/호수", "입력시간", "기타사항"]
                    df.loc[-1] = new_row
                    df.index = df.index + 1
                    df = df.sort_index()
                else:
                    # 데이터가 없는 경우 빈 DataFrame 생성
                    df = pd.DataFrame(columns=["고객명", "아파트명", "계약조건", "시공일", "비고", "시공자", "연락처", "시간",
                                            "비고", "주소", "동/호수", "입력시간", "기타사항"])

                FileName = test.replace("uploads\\", '')

                return render_template("index.html", df=df, FileName=FileName, test=test, filter_conditions={
                    "date1": des1, "date2": des6,
                    "time": des2, "group": des3,
                    "installer": des4, "str": des5,
                    "apart": des7
                }, str_options_html=str_options_html, str_options_html2=str_options_html2,str_options_html3=str_options_html3)
            else:
                alert_message = "조건에 맞는 데이터가 없습니다"  # 알림 메시지 설정
                return render_template("index.html", df=None, alert_message=alert_message)
        else:
            alert_message = "파일을 입력해주세요"  # 알림 메시지 설정
            return render_template("index.html", df=None, alert_message=alert_message)

@app.route("/download_excel")
def download_excel():
    global df

    excel_filename = "output.xlsx"
    df.to_excel(excel_filename, index=False)

    return send_file(excel_filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)