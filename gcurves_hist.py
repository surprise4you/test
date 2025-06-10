import requests
import pandas as pd
from datetime import datetime, timedelta
import time
import os
import logging

PERIODS = [0.25, 0.5, 0.75, 1, 2, 3, 5, 7, 10, 15, 20, 30]
SAVE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(SAVE_DIR, exist_ok=True)
EXCEL_FILE = os.path.join(SAVE_DIR, "gcurves_hist.xlsx")
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "gcurves_hist.log")

logging.basicConfig(
    filename=LOG_FILE,
    filemode='a',
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def daterange(start_date, end_date):
    for n in range((end_date - start_date).days + 1):
        yield start_date + timedelta(n)

def update_gcurve_file(to_date=None):
    logging.info("Старт update_gcurve_file()")
    start_date = datetime.strptime("2024-12-30", "%Y-%m-%d")
    if to_date is None:
        end_date = datetime.today()
    else:
        end_date = to_date if isinstance(to_date, datetime) else datetime.combine(to_date, datetime.min.time())
    all_rows = []

    if os.path.exists(EXCEL_FILE):
        df_old = pd.read_excel(EXCEL_FILE)
        all_rows = df_old.to_dict(orient="records")
        if not df_old.empty:
            last_date = pd.to_datetime(df_old["tradedate"], format="%d.%m.%Y").max()
            msg = f"Файл найден, последние данные: {last_date.strftime('%d.%m.%Y')}. Запрашиваем только новые даты."
            logging.info(msg)
            start_date = last_date + timedelta(days=1)
        else:
            logging.info("Файл есть, но пустой — скачиваем с начала периода.")
    else:
        logging.info("Файл не найден, скачиваем с начала периода.")

    new_rows = []
    for date in daterange(start_date, end_date):
        dstr = date.strftime("%Y-%m-%d")
        dstr_fmt = date.strftime("%d.%m.%Y")
        url = f"https://iss.moex.com/iss/engines/stock/zcyc.json?date={dstr}&iss.meta=off&iss.json=extended&iss.only=yearyields"
        try:
            r = requests.get(url, timeout=5)
            data = r.json()
            found = False
            for block in data:
                if isinstance(block, dict) and "yearyields" in block:
                    df = pd.DataFrame(block["yearyields"])
                    if not df.empty:
                        row = {"tradedate": dstr_fmt}
                        for p in PERIODS:
                            val = df.loc[df["period"] == p, "value"]
                            row[str(p)] = val.values[0] if not val.empty else None
                        new_rows.append(row)
                        found = True
                    break
            msg = f"{dstr}: {'+' if found else 'нет данных'}"
            logging.info(msg)
            time.sleep(0.25)
        except Exception as e:
            msg = f"{dstr}: ошибка ({e})"
            logging.error(msg)

    all_rows = all_rows + new_rows
    df_all = pd.DataFrame(all_rows)
    df_all["tradedate"] = pd.to_datetime(df_all["tradedate"], format="%d.%m.%Y")
    df_all = df_all.drop_duplicates(subset=["tradedate"], keep="first")
    df_all = df_all.sort_values("tradedate").reset_index(drop=True)
    # ОГРАНИЧЕНИЕ итогового датафрейма по to_date:
    if to_date is not None:
        df_all = df_all[df_all["tradedate"] <= end_date]
    df_all["tradedate"] = df_all["tradedate"].dt.strftime("%d.%m.%Y")
    df_all.to_excel(EXCEL_FILE, index=False)
    msg = f"\nСохранено: {len(df_all)} дней в {EXCEL_FILE}"
    logging.info(msg)
    return EXCEL_FILE


if __name__ == "__main__":
    update_gcurve_file()