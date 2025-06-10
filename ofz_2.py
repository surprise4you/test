import pandas as pd
import os
import logging

SAVE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
SRC_FILE = os.path.join(SAVE_DIR, "gcurves_hist.xlsx")
DEST_FILE = os.path.join(SAVE_DIR, "ofz_2.xlsx")
LOG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ofz_2.log")

# Настройка логирования
logging.basicConfig(
    filename=LOG_FILE,
    filemode='a',
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def update_ofz2_file(to_date=None):
    logging.info("Старт update_ofz2_file()")
    if not os.path.exists(SRC_FILE):
        msg = f"Нет исходного файла {SRC_FILE}! Сначала соберите gcurves_hist."
        logging.warning(msg)
        return None
    df = pd.read_excel(SRC_FILE)
    if "2" not in df.columns or "tradedate" not in df.columns:
        msg = "В исходной таблице нет столбца '2' или 'tradedate'!"
        logging.warning(msg)
        return None
    df_2y = df[["tradedate", "2"]].copy()
    df_2y.rename(columns={"tradedate": "date", "2": "price"}, inplace=True)
    df_2y = df_2y.dropna(subset=["price"])
    if to_date is not None:
        if not isinstance(to_date, pd.Timestamp):
            to_date = pd.to_datetime(to_date)
        df_2y["date_dt"] = pd.to_datetime(df_2y["date"], format="%d.%m.%Y", errors="coerce")
        df_2y = df_2y[df_2y["date_dt"] <= to_date]
        df_2y = df_2y.drop(columns=["date_dt"])
    if df_2y.empty:
        msg = "Нет данных по 2-летним значениям."
        logging.warning(msg)
        return None
    df_2y.to_excel(DEST_FILE, index=False)
    msg = f"Таблица по 2-летней доходности сохранена в {DEST_FILE} ({len(df_2y)} строк)"
    logging.info(msg)
    return DEST_FILE


if __name__ == "__main__":
    update_ofz2_file()
