import pandas as pd
import os
import logging

SAVE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
SRC_FILE = os.path.join(SAVE_DIR, "gcurves_hist.xlsx")
DEST_FILE = os.path.join(SAVE_DIR, "ofz_10.xlsx")
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ofz_10.log")

# Настройка логирования
logging.basicConfig(
    filename=LOG_FILE,
    filemode='a',
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def update_ofz10_file(to_date=None):
    logging.info("Старт update_ofz10_file()")
    if not os.path.exists(SRC_FILE):
        msg = f"Нет исходного файла {SRC_FILE}! Сначала соберите gcurves_hist."
        logging.warning(msg)
        return None
    df = pd.read_excel(SRC_FILE)
    if "10" not in df.columns or "tradedate" not in df.columns:
        msg = "В исходной таблице нет столбца '10' или 'tradedate'!"
        logging.warning(msg)
        return None
    df_10y = df[["tradedate", "10"]].copy()
    df_10y.rename(columns={"tradedate": "date", "10": "price"}, inplace=True)
    df_10y = df_10y.dropna(subset=["price"])
    if to_date is not None:
        if not isinstance(to_date, pd.Timestamp):
            to_date = pd.to_datetime(to_date)
        df_10y["date_dt"] = pd.to_datetime(df_10y["date"], format="%d.%m.%Y", errors="coerce")
        df_10y = df_10y[df_10y["date_dt"] <= to_date]
        df_10y = df_10y.drop(columns=["date_dt"])
    if df_10y.empty:
        msg = "Нет данных по 10-летним значениям."
        logging.warning(msg)
        return None
    df_10y.to_excel(DEST_FILE, index=False)
    msg = f"Таблица по 10-летней доходности сохранена в {DEST_FILE} ({len(df_10y)} строк)"
    logging.info(msg)
    return DEST_FILE

if __name__ == "__main__":
    update_ofz10_file()
