import os
import json
import urllib.request
import pandas as pd

# 內政部預售屋備查三縣市檔案下載網址代碼
# A: 台北市, F: 新北市, H: 桃園市
CITIES = [
    ("台北市", "A", "https://plvr.land.moi.gov.tw//Download?type=zip&fileName=a_lvr_buildcase.xls"),
    ("新北市", "F", "https://plvr.land.moi.gov.tw//Download?type=zip&fileName=f_lvr_buildcase.xls"),
    ("桃園市", "H", "https://plvr.land.moi.gov.tw//Download?type=zip&fileName=h_lvr_buildcase.xls"),
]

all_records = []

for city_name, city_code, url in CITIES:
    local_file = f"{city_code.lower()}_lvr_buildcase.xls"
    print(f"處理中: {city_name} ...")
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as resp, open(local_file, 'wb') as f:
            f.write(resp.read())
    except Exception as e:
        print(f"下載 {city_name} 失敗，使用本機既有檔案或略過: {e}")

    if not os.path.exists(local_file):
        continue

    try:
        # 跳過第一列（英文欄位代號列）
        df = pd.read_excel(local_file, sheet_name='預售屋備查', skiprows=[1])
        for _, row in df.iterrows():
            households = row.get('層棟戶數')
            try:
                households = int(households)
            except:
                households = 0

            permit_date = str(row.get('建照核發日期', '')).replace('.0', '')
            if len(permit_date) == 7: # 民國年轉換
                permit_date = f"{int(permit_date[:3])+1911}-{permit_date[3:5]}-{permit_date[5:]}"

            record = {
                "city": city_name,
                "town": str(row.get('鄉鎮市區', '')).strip(),
                "name": str(row.get('建案名稱', '')).strip(),
                "builder": str(row.get('起造人', '')).strip(),
                "households": households,
                "zoning": str(row.get('使用分區', '')).strip(),
                "location": str(row.get('坐落街道', '')).strip(),
                "permitDate": permit_date
            }
            if record["name"] and record["name"] != "nan":
                all_records.append(record)
    except Exception as e:
        print(f"解析 {city_name} 失敗: {e}")

# 輸出成 data.json
with open('data.json', 'w', encoding='utf-8') as f:
    json.dump(all_records, f, ensure_ascii=False, indent=2)

print(f"完成！共匯入 {len(all_records)} 筆建案資料至 data.json")
