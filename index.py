import subprocess
import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta
import os
from openpyxl import load_workbook

# 저장소 목록 정의
repositories = {
    'CIVIL NX Master': 'C:\\Users\\LEEGEONWOO\\Dev\\CIVIL_NX\\genw_new',
    'CIVIL NX 955': 'C:\\Users\\LEEGEONWOO\\Dev\\CIVIL_NX_v955',
    'eGen': 'C:\\Users\\LEEGEONWOO\\Dev\\eGen\\egen_jp_2017_ODA',
    'plug in': 'C:\\Users\\LEEGEONWOO\\Dev\\API\\PUBLIC-plugins',
    # 추가 저장소 경로
}

TODAY = datetime.date.today()
def main():
  for project, path in repositories.items():
      print(f"Commit History for {project}:")
      with subprocess.Popen(
            ["git", "-C", path, "log", "--pretty=format:%h - %an, %ar : %s"],
            stdout=subprocess.PIPE,
            text=True,
            encoding="utf-8"
        ) as proc:
        data = inputData(proc)
        if data :
          df = pd.DataFrame(data)
          folder_path =  f"./logExcel/{str(TODAY).split("-")[0]}/"
          if not os.path.exists(folder_path):
            os.makedirs(folder_path, exist_ok=True)

          file_path = os.path.join(folder_path, f"{TODAY}.xlsx")
          mode = 'a' if os.path.exists(file_path) else 'w'
          existSheet = 'replace' if os.path.exists(file_path) else None
          with pd.ExcelWriter(file_path, engine='openpyxl', mode=mode, if_sheet_exists=existSheet) as writer:
            df.to_excel(writer, index=False, sheet_name=f"{project}")
            # 열 너비 설정
            workbook = writer.book
            worksheet = writer.sheets[project]
            worksheet.column_dimensions['A'].width = 20
            worksheet.column_dimensions['B'].width = 20
            worksheet.column_dimensions['C'].width = 20
            worksheet.column_dimensions['D'].width = 100
          # 필터 설정
          setFilter(project, file_path)

def setFilter(project, file_path):
    wb = load_workbook(file_path)
    ws = wb[project]
    ws.auto_filter.ref = "B:C"
    wb.save(file_path)

def inputData(proc):
  newData = {
  '코드': [],
  '이름': [],
  '날짜': [],
  'commit': [],
  }
  for line in proc.stdout:
    sLine = line.split(" : ")[0].split(",")
    if len(sLine) > 2 : continue
    bYearCheck = sLine[1].split()[1]
    if bYearCheck == "months" or bYearCheck == "years" or  bYearCheck == "year": continue

    code = sLine[0].split(" - ")[0]
    user = sLine[0].split(" - ")[1]
    commitMsg = line.split(" : ")[1]
    if user == "chlim" or user == "gw.lee":
      bWeekCheck = sLine[1].split()[1]
      if bWeekCheck == "weeks" or bWeekCheck == "week":
        week = sLine[1].split()[0]
        if int(week) < 9:
          newData["코드"].append(code)
          newData["이름"].append(user)
          newData["날짜"].append(getDate("week", int(week)))
          newData["commit"].append(commitMsg)
      elif bWeekCheck == "hours" or bWeekCheck == "hour":
        hour = sLine[1].split()[0]
        newData["코드"].append(code)
        newData["이름"].append(user)
        newData["날짜"].append(getDate("hour", int(hour)))
        newData["commit"].append(commitMsg)
      else :
        day = sLine[1].split()[0]
        if int(day) < 32:
          newData["코드"].append(code)
          newData["이름"].append(user)
          newData["날짜"].append(getDate("day", int(day)))
          newData["commit"].append(commitMsg)
  return newData

def getDate(strDay, nCnt):
  if strDay == "week":
    # week
    ago = TODAY - relativedelta(weeks=nCnt)
  elif strDay == "hour":
    # hour
    ago = TODAY - datetime.timedelta(hours=nCnt)
  else :
    # day
    ago = TODAY - relativedelta(days=nCnt)
  return ago
  
if __name__ == "__main__":
  main()