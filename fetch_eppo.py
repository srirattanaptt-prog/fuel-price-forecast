import datetime
import json
import os
import urllib.request
from bs4 import BeautifulSoup
import holidays
import pandas as pd
import requests


def get_current_sgd_thb():
  """ดึงอัตราแลกเปลี่ยน SGD/THB สดจาก Google Finance"""
  url = "https://www.google.com/finance/quote/SGD-THB"
  headers = {"User-Agent": "Mozilla/5.0"}
  try:
    resp = requests.get(url, headers=headers, timeout=10)
    soup = BeautifulSoup(resp.text, "html.parser")
    rate_div = soup.find("div", {"class": "YMlKec fxKbKc"})
    if rate_div:
      return float(rate_div.text.replace(",", ""))
  except Exception as e:
    print(f"FX Fetch Error: {e}")
  return 27.6134  # ค่าสำรองกรณีเชื่อมต่อขัดข้อง


def is_thai_holiday_or_weekend(dt):
  """ตรวจสอบว่าเป็นวันเสาร์-อาทิตย์ หรือวันหยุดราชการไทยหรือไม่"""
  # 1. เช็กวันเสาร์ (5) หรือวันอาทิตย์ (6)
  if dt.weekday() >= 5:
    return True, "วันหยุดสุดสัปดาห์ (เสาร์-อาทิตย์)"

  # 2. เช็กวันหยุดราชการไทย (Thai Official Holidays)
  thai_holidays = holidays.Thailand(years=dt.year)
  if dt.date() in thai_holidays:
    holiday_name = thai_holidays.get(dt.date())
    return True, f"วันหยุดราชการ ({holiday_name})"

  return False, ""


def main():
  # ตั้งเวลาอ้างอิงเวลาประเทศไทย (UTC+7)
  tz_th = datetime.timezone(datetime.timedelta(hours=7))
  now_th = datetime.datetime.now(tz_th)
  today_str = now_th.strftime("%Y-%m-%d")

  # 1. ตรวจสอบเงื่อนไขวันหยุด
  is_holiday, reason = is_thai_holiday_or_weekend(now_th)
  if is_holiday:
    print(f"[{now_th.strftime('%Y-%m-%d %H:%M:%S')}] วันนี้เป็น{reason} -> ระบบข้ามการทำงาน")
    return

  # 2. ตรวจสอบว่าวันนี้ sync สำเร็จไปแล้วหรือยัง
  if os.path.exists("data.json"):
    try:
      with open("data.json", "r", encoding="utf-8") as f:
        old_data = json.load(f)
        if (
            old_data.get("date") == today_str
            and old_data.get("syncStatus") is True
        ):
          print(
              f"[{now_th.strftime('%H:%M:%S')}] ข้อมูลของวันที่ {today_str}"
              " อัปเดตสมบูรณ์เรียบร้อยแล้ว ข้ามการทำงาน"
          )
          return
    except Exception:
      pass

  # 3. สร้าง URL ตามรูปแบบการโพสต์ของ สนพ.
  year = now_th.year
  month = now_th.month
  day = now_th.day
  excel_url = f"https://www.eppo.go.th/wp-content/uploads/{year}/{month:02d}/pt-price-st-{year}-{month}-{day}.xlsx"
  print(f"กำลังตรวจสอบไฟล์ที่: {excel_url}")

  headers = {"User-Agent": "Mozilla/5.0"}
  req = urllib.request.Request(excel_url, headers=headers)

  try:
    with urllib.request.urlopen(req, timeout=15) as resp:
      content = resp.read()
      with open("temp_eppo.xlsx", "wb") as f:
        f.write(content)
    print("พบไฟล์ Excel ใหม่ประจำวัน กำลังประมวลผล...")

    # 4. อ่านข้อมูลแท็บ MARKETING MARGIN
    df = pd.read_excel("temp_eppo.xlsx", sheet_name="MARKETING MARGIN")

    diesel_col = "B7" if "B7" in df.columns else "H-DIESEL"
    g95_col = "GASOHOL95E10"

    diesel_margin = float(df[diesel_col].dropna().iloc[-1])
    g95_margin = float(df[g95_col].dropna().iloc[-1])

    # 5. ดึงอัตราแลกเปลี่ยน SGD/THB
    fx_rate = get_current_sgd_thb()

    # 6. บันทึกผลลง data.json
    payload = {
        "date": today_str,
        "updatedAt": now_th.strftime("%Y-%m-%d %H:%M:%S"),
        "dieselMargin": diesel_margin,
        "g95Margin": g95_margin,
        "fxRate": fx_rate,
        "syncStatus": True,
    }

    with open("data.json", "w", encoding="utf-8") as f:
      json.dump(payload, f, ensure_ascii=False, indent=2)

    print(
        f"บันทึกข้อมูลสำเร็จ! ดีเซล: {diesel_margin:.4f}, G95:"
        f" {g95_margin:.4f}, FX: {fx_rate:.4f}"
    )

    if os.path.exists("temp_eppo.xlsx"):
      os.remove("temp_eppo.xlsx")

  except urllib.error.HTTPError as e:
    if e.code == 404:
      print(
          f"[{now_th.strftime('%H:%M:%S')}] ยังไม่พบไฟล์ประจำวัน (404 Not"
          " Found) - ระบบจะรอรอบถัดไปอีก 15 นาที"
      )
    else:
      print(f"HTTP Error: {e.code}")
  except Exception as e:
    print(f"Error: {e}")


if __name__ == "__main__":
  main()