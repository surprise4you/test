import os

SAVE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
os.makedirs(SAVE_DIR, exist_ok=True)

import datetime
import requests
import pandas as pd
import yfinance as yf
import zipfile
import re
from io import StringIO
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.constants import ChatAction
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, ContextTypes
import logging

from gcurves_hist import update_gcurve_file
from ofz_2 import update_ofz2_file
from ofz_10 import update_ofz10_file


TOKEN = ""
ADMIN_CHAT_ID = 123


LOG_FILE = os.path.join("imoex.log")

# Настройка логирования
logging.basicConfig(
    filename=LOG_FILE,
    filemode='a',
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)


TICKERS = {
    "IMOEX":   {"source": "moex", "code": "IMOEX", "display": "IMOEX"},
    "RTSI":    {"source": "moex", "code": "RTSI", "display": "RTSI"},
    "RGBI":    {"source": "moex", "code": "RGBI", "display": "RGBI"},
    "^GSPC":   {"source": "yahoo", "code": "^GSPC", "display": "S&P 500 (^GSPC)"},
    "^HSI":    {"source": "yahoo", "code": "^HSI", "display": "Hang Seng (^HSI)"},
    "^STOXX50E": {"source": "yahoo", "code": "^STOXX50E", "display": "Euro Stoxx 50 (^STOXX50E)"},
    "^TNX":    {"source": "yahoo", "code": "^TNX", "display": "10Y Treasury (^TNX)"},
    "BZ=F":    {"source": "yahoo", "code": "BZ=F", "display": "Brent Crude Oil (BZ=F)"},
    "GC=F":    {"source": "yahoo", "code": "GC=F", "display": "Gold Futures (GC=F)"},
    "EURUSD=X": {"source": "yahoo", "code": "EURUSD=X", "display": "EUR/USD"},
    "CNY=X": {"source": "yahoo", "code": "CNY=X", "display": "CNY/USD"},
    "^N225": {"source": "yahoo", "code": "^N225", "display": "Nikkei 225 (^N225)"},
    "DX-Y.NYB": {"source": "yahoo", "code": "DX-Y.NYB", "display": "US Dollar Index (DXY)"},
    "BTC-USD": {"source": "yahoo", "code": "BTC-USD", "display": "Bitcoin (BTC-USD)"},
    "USDRUB": {"source": "cbr", "code": "R01235", "display": "USD/RUB"},
    "EURRUB": {"source": "cbr", "code": "R01239", "display": "EUR/RUB"},
    "CNYRUB": {"source": "moex_currency", "code": "CNYRUB_TOM", "display": "CNY/RUB"},
    "G_CURVE": {"source": "custom", "code": "gcurves_hist", "display": "КБД"},
    "OFZ2":   {"source": "custom", "code": "ofz_2", "display": "ОФЗ 2"},
    "OFZ10":   {"source": "custom", "code": "ofz_10", "display": "ОФЗ 10"},
}


def get_filename(k):
    # Для custom — явно указываем имена файлов
    if k == "G_CURVE":
        return os.path.join(SAVE_DIR, "gcurves_hist.xlsx")
    elif k == "OFZ2":
        return os.path.join(SAVE_DIR, "ofz_2.xlsx")
    elif k == "OFZ10":
        return os.path.join(SAVE_DIR, "ofz_10.xlsx")
    else:
        return os.path.join(SAVE_DIR, f"{k}.xlsx")


# Автосоздание файлов, если их нет
gcurve_path = os.path.join(SAVE_DIR, "gcurves_hist.xlsx")
ofz2_path = os.path.join(SAVE_DIR, "ofz_2.xlsx")
ofz10_path = os.path.join(SAVE_DIR, "ofz_10.xlsx")

if not os.path.exists(gcurve_path):
    print(f"Автоматически создаю {gcurve_path}...")
    update_gcurve_file()
if not os.path.exists(ofz10_path):
    print(f"Автоматически создаю {ofz10_path}...")
    update_ofz10_file()

# Источники
def fetch_index_closing_table(index: str, start_date: str, end_date: str) -> pd.DataFrame:
    rows = []
    start = 0
    limit = 100
    while True:
        url = (
            f"https://iss.moex.com/iss/history/engines/stock/markets/index/securities/{index}.json"
            f"?from={start_date}&till={end_date}&start={start}&limit={limit}&sort_order=TRADEDATE"
            f"&sort_order_desc=asc&iss.meta=off&iss.json=extended&lang=ru"
        )
        resp = requests.get(url)
        js = resp.json()
        history_block = None
        if isinstance(js, list):
            for block in js:
                if isinstance(block, dict) and "history" in block:
                    history_block = block["history"]
                    break
            if history_block is None:
                raise Exception(f"Нет блока 'history' в ответе MOEX для {index}!")
        else:
            raise Exception("Ответ MOEX не похож на list, обнови функцию!")
        if not history_block:
            break
        rows.extend(history_block)
        if len(history_block) < limit:
            break
        start += limit
    df = pd.DataFrame(rows)
    print(f"[DEBUG][{index}] всего строк: {len(df)}")
    df = df[['TRADEDATE', 'CLOSE']].copy()
    df = df.sort_values('TRADEDATE')
    df.rename(columns={'TRADEDATE': 'date', 'CLOSE': 'price'}, inplace=True)
    df['date'] = pd.to_datetime(df['date']).dt.strftime('%d.%m.%Y')
    return df[['date', 'price']]


def fetch_yahoo_closing_history(ticker, start_date, end_date) -> pd.DataFrame:
    data = yf.download(ticker, start=start_date, end=end_date)
    if data.empty:
        raise Exception(f"Нет данных по {ticker} за выбранный период!")
    data = data.reset_index()
    df = data[["Date", "Close"]].copy()
    df.rename(columns={"Date": "date", "Close": "price"}, inplace=True)
    df['date'] = pd.to_datetime(df['date']).dt.strftime('%d.%m.%Y')
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    return df[['date', 'price']]

def fetch_cbr_currency_history(currency_code: str, start_date: str, end_date: str) -> pd.DataFrame:
    url = (
        "https://www.cbr.ru/currency_base/dynamics/"
        f"?UniDbQuery.Posted=True&UniDbQuery.mode=1"
        f"&UniDbQuery.date_req1={start_date}"
        f"&UniDbQuery.date_req2={end_date}"
        f"&UniDbQuery.VAL_NM_RQ={currency_code}"
    )
    resp = requests.get(url)
    dfs = pd.read_html(StringIO(resp.text), decimal=',', thousands='\xa0')
    df = dfs[-1]
    header_row_idx = None
    for i, row in df.iterrows():
        if any('Дата' in str(x) for x in row):
            header_row_idx = i
            break
    if header_row_idx is None:
        raise Exception("Не удалось найти строку с названиями колонок!")
    df.columns = df.iloc[header_row_idx]
    df = df.iloc[header_row_idx+1:].reset_index(drop=True)
    date_pattern = re.compile(r'^\d{2}\.\d{2}\.\d{4}$')
    df = df[df['Дата ▼'].apply(lambda x: bool(date_pattern.match(str(x))))].copy()
    df = df[['Дата ▼', 'Курс']].copy()
    df.rename(columns={'Дата ▼': 'date', 'Курс': 'price'}, inplace=True)
    df['price'] = df['price'].astype(str).str.replace('\xa0', '').str.replace(',', '.').astype(float)
    df['date'] = pd.to_datetime(df['date'], dayfirst=True)
    df = df.sort_values('date').reset_index(drop=True)
    df['date'] = df['date'].dt.strftime('%d.%m.%Y')
    df = df.drop_duplicates(subset=["date"], keep='last')
    return df[['date', 'price']]

def fetch_moex_currency_pair(pair_code, start_date, end_date):
    rows = []
    start = 0
    limit = 100
    while True:
        url = (
            f"https://iss.moex.com/iss/history/engines/currency/markets/selt/securities/{pair_code}.json"
            f"?from={start_date}&till={end_date}&start={start}&limit={limit}&iss.meta=off&iss.json=extended"
        )
        resp = requests.get(url)
        js = resp.json()
        history_block = None
        if isinstance(js, list):
            for block in js:
                if isinstance(block, dict) and "history" in block:
                    history_block = block["history"]
                    break
        elif isinstance(js, dict):
            if "history" in js:
                history_block = js["history"]
        if not history_block:
            break
        rows.extend(history_block)
        if len(history_block) < limit:
            break
        start += limit
    if not rows:
        raise Exception(f"Нет данных по {pair_code} за период!")
    df = pd.DataFrame(rows)
    if 'BOARDID' in df.columns:
        df = df[df['BOARDID'] == 'CETS']
    price_col = None
    for col in ["CLOSE", "WAPRICE"]:
        if col in df.columns:
            price_col = col
            break
    if price_col is None:
        raise Exception(f"В данных нет ни CLOSE, ни WAPRICE! Колонки: {df.columns}")
    df = df[['TRADEDATE', price_col]].copy()
    df.rename(columns={'TRADEDATE': 'date', price_col: 'price'}, inplace=True)
    df = df.sort_values(["date"]).reset_index(drop=True)
    df = df.dropna(subset=['price'])
    df = df.groupby("date", as_index=False).first()
    df['date'] = pd.to_datetime(df['date']).dt.strftime('%d.%m.%Y')
    return df[['date', 'price']]

# --- Функция обновления всех тикеров ---
def update_all_tickers(start_date, end_date, start_cbr, end_cbr, start_moex_currency):
    for k, v in TICKERS.items():
        try:
            if v["source"] == "yahoo":
                df = fetch_yahoo_closing_history(v["code"], start_date, end_date)
                filename = get_filename(k)
                df.to_excel(filename, index=False)
            elif v["source"] == "moex":
                df = fetch_index_closing_table(v["code"], start_date, end_date)
                filename = get_filename(k)
                df.to_excel(filename, index=False)
            elif v["source"] == "cbr":
                df = fetch_cbr_currency_history(v["code"], start_cbr, end_cbr)
                filename = get_filename(k)
                df.to_excel(filename, index=False)
            elif v["source"] == "moex_currency":
                df = fetch_moex_currency_pair(v["code"], start_moex_currency, end_date)
                filename = get_filename(k)
                df.to_excel(filename, index=False)
            elif v["source"] == "custom":
                ensure_custom_file(v["code"])
        except Exception as e:
            print(f"{k}: {e}")

# Сводная таблица
# --- Округление значений по тикерам ---
ROUND_STYLE = {
    # Индексы
    "IMOEX": 0, "RTSI": 0, "RGBI": 2, "^GSPC": 0, "^HSI": 0, "^STOXX50E": 0, "^N225": 0,
    # Валюты
    "USDRUB": 2, "EURRUB": 2, "CNYRUB": 2, "EURUSD=X": 2, "CNY=X": 2,
    # Доходности/проценты
    "G_CURVE": 2, "OFZ2": 2, "OFZ10": 2, "^TNX": 2,
    # Товары
    "BZ=F": 2, "GC=F": 0,
    # Индекс доллара
    "DX-Y.NYB": 2,
    # Крипта
    "BTC-USD": 0,
}

# --- Красивое отображение тикеров ---
TICKER_DISPLAY = {
    "IMOEX": "IMOEX",
    "RTSI": "RTSI",
    "RGBI": "RGBI",
    "^GSPC": "SPX",
    "^HSI": "HSI",
    "^STOXX50E": "STOXX50E",
    "^TNX": "US10Y",
    "BZ=F": "Brent",
    "GC=F": "Gold",
    "EURUSD=X": "EURUSD",
    "CNY=X": "CNYUSD",
    "^N225": "N225",
    "DX-Y.NYB": "DXY",
    "BTC-USD": "BTC",
    "USDRUB": "USDRUB",
    "EURRUB": "EURRUB",
    "CNYRUB": "CNYRUB",
    "G_CURVE": "КБД",
    "OFZ2": "ОФЗ 2г",
    "OFZ10": "ОФЗ 10л",
}

import calendar

def get_last_day_of_prev_quarter(today):
    quarter = (today.month - 1) // 3 + 1
    if quarter == 1:
        year = today.year - 1
        month = 12
    else:
        year = today.year
        month = 3 * (quarter - 1)
    day = calendar.monthrange(year, month)[1]
    return datetime.date(year, month, day)

def build_summary_table():
    records = []
    today = datetime.date.today()
    week_ago = today - datetime.timedelta(days=7)
    last_q_date = get_last_day_of_prev_quarter(today)
    year_start = datetime.date(today.year, 1, 1)

    def find_prev_value(df, col_date, col_price, ref_date):
        ref_date = pd.Timestamp(ref_date)
        candidates = df[df[col_date] <= ref_date].sort_values(col_date)
        if len(candidates) == 0:
            return None, None
        val = candidates.iloc[-1][col_price]
        d = candidates.iloc[-1][col_date]
        return val, d

    def find_first_after(df, col_date, col_price, ref_date):
        ref_date = pd.Timestamp(ref_date)
        candidates = df[df[col_date] >= ref_date].sort_values(col_date)
        if len(candidates) > 0:
            val = candidates.iloc[0][col_price]
            d = candidates.iloc[0][col_date]
            return val, d
        val = df.iloc[0][col_price]
        d = df.iloc[0][col_date]
        return val, d

    def fmt(val, k):
        r = ROUND_STYLE.get(k, 2)
        try:
            if val is None or pd.isnull(val):
                return "-"
            val = float(val)
            if r == 0:
                return int(round(val))
            else:
                return f"{val:,.{r}f}".replace(",", " ").replace(".", ",")
        except Exception:
            return str(val) if val is not None else "-"

    for k, v in TICKERS.items():
        if k == "G_CURVE":
            continue  # КБД исключаем из свода
        fname = get_filename(k)
        if not os.path.exists(fname):
            continue
        try:
            df = pd.read_excel(fname)
            col_date = "tradedate" if k == "G_CURVE" else "date"
            col_price = "10" if k == "G_CURVE" else "price"
            if col_date not in df.columns or col_price not in df.columns:
                continue
            df = df.dropna(subset=[col_price])
            df[col_date] = pd.to_datetime(df[col_date], dayfirst=True)
            df = df.sort_values(col_date)
            if df.empty:
                continue

            last = df.iloc[-1]
            last_value = last[col_price]
            last_date = last[col_date]

            prev_week, prev_week_date = find_prev_value(df, col_date, col_price, week_ago)
            qtd_value, qtd_date = find_first_after(df, col_date, col_price, last_q_date)
            ytd_value, ytd_date = df.iloc[0][col_price], df.iloc[0][col_date]

            def delta(val1, val2):
                if val1 is None or val2 is None or val2 == 0 or pd.isnull(val1) or pd.isnull(val2):
                    return ""
                d = round((val1 - val2) / val2 * 100, 1)
                sign = "+" if d > 0 else ""
                return f"{sign}{d:.1f}%"

            rec = {
                "Тикер": TICKER_DISPLAY.get(k, k),
                "Код": k,
                "Значение": fmt(last_value, k),
                "Дата_последнего": last_date,
                "Неделя": fmt(prev_week, k),
                "Дата_недели": prev_week_date,
                "Квартал": fmt(qtd_value, k),
                "Дата_квартала": qtd_date,
                "Год": fmt(ytd_value, k),
                "Дата_года": ytd_date,
                "Изм. за неделю": delta(last_value, prev_week),
                "Изм. с начала года": delta(last_value, ytd_value),
                "Изм. с начала квартала": delta(last_value, qtd_value),
            }
            records.append(rec)
        except Exception as ex:
            logging.error(f"Ошибка в build_summary_table для {k}: {ex}")
            continue

    columns = [
        "Тикер", "Код",
        "Значение", "Дата_последнего",
        "Неделя", "Дата_недели",
        "Квартал", "Дата_квартала",
        "Год", "Дата_года",
        "Изм. за неделю", "Изм. с начала года", "Изм. с начала квартала"
    ]
    return pd.DataFrame(records, columns=columns)



# Бот
async def load_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    keyboard = [
        [InlineKeyboardButton("🗂️ Загрузить все", callback_data='load_all_zip')],
        [InlineKeyboardButton("📊 Сводная таблица", callback_data='summary_table')],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("Выберите действие:", reply_markup=reply_markup)
    await context.bot.send_message(
        chat_id=ADMIN_CHAT_ID,
        text=f"Пользователь {update.effective_user.id} ({update.effective_user.username}) вызвал /load"
    )


# async def load_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
#     keyboard = [
#         [InlineKeyboardButton(v["display"], callback_data=f'load_{k}')]
#         for k, v in TICKERS.items()
#     ]
#     keyboard.append([InlineKeyboardButton("🗂️Загрузить все", callback_data='load_all_zip')])
#     reply_markup = InlineKeyboardMarkup(keyboard)
#     await update.message.reply_text("Выберите тикер:", reply_markup=reply_markup)
#     await context.bot.send_message(
#         chat_id=ADMIN_CHAT_ID,
#         text=f"Пользователь {update.effective_user.id} ({update.effective_user.username}) вызвал /load"
#     )


def ensure_custom_file(code):
    if code == "gcurves_hist":
        path = update_gcurve_file()
        if not (path and os.path.exists(path)):
            raise Exception(f"Файл {path} не создан/не найден.")
        return path
    elif code == "ofz_2":
        update_gcurve_file()
        path = update_ofz2_file()
        if not (path and os.path.exists(path)):
            raise Exception(f"Файл {path} не создан/не найден.")
        return path
    elif code == "ofz_10":
        update_gcurve_file()
        path = update_ofz10_file()
        if not (path and os.path.exists(path)):
            raise Exception(f"Файл {path} не создан/не найден.")
        return path
    else:
        raise Exception(f"Неизвестный custom-тикер: {code}")


async def handle_ticker_press(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    start_date = "2024-12-28"
    end_date = datetime.date.today().strftime("%Y-%m-%d")
    start_cbr = "29.12.2024"
    end_cbr = datetime.date.today().strftime("%d.%m.%Y")
    start_moex_currency = "2024-12-30"

    def ensure_custom_file(code):
        # Создаёт файл (если надо), возвращает путь к нему
        if code == "gcurves_hist":
            path = update_gcurve_file()
            if not (path and os.path.exists(path)):
                raise Exception(f"Файл {path} не создан/не найден.")
            return path
        elif code == "ofz_2":
            update_gcurve_file()
            path = update_ofz2_file()
            if not (path and os.path.exists(path)):
                raise Exception(f"Файл {path} не создан/не найден.")
            return path
        elif code == "ofz_10":
            update_gcurve_file()  # обязательно обновить основную!
            path = update_ofz10_file()
            if not (path and os.path.exists(path)):
                raise Exception(f"Файл {path} не создан/не найден.")
            return path
        else:
            raise Exception(f"Неизвестный custom-тикер: {code}")

    if query.data == "load_all_zip":
        msg = await query.edit_message_text("⏳ Формирую архив со всеми таблицами, подождите...")
        await context.bot.send_chat_action(chat_id=query.message.chat.id, action=ChatAction.TYPING)
        files = []
        errors = []
        for k, v in TICKERS.items():
            try:
                filename = None
                if v["source"] == "yahoo":
                    df = fetch_yahoo_closing_history(v["code"], start_date, end_date)
                    filename = os.path.join(SAVE_DIR, f"{k}.xlsx")
                    df.to_excel(filename, index=False)
                elif v["source"] == "moex":
                    df = fetch_index_closing_table(v["code"], start_date, end_date)
                    filename = os.path.join(SAVE_DIR, f"{k}.xlsx")
                    df.to_excel(filename, index=False)
                elif v["source"] == "cbr":
                    df = fetch_cbr_currency_history(v["code"], start_cbr, end_cbr)
                    filename = os.path.join(SAVE_DIR, f"{k}.xlsx")
                    df.to_excel(filename, index=False)
                elif v["source"] == "moex_currency":
                    df = fetch_moex_currency_pair(v["code"], start_moex_currency, end_date)
                    filename = os.path.join(SAVE_DIR, f"{k}.xlsx")
                    df.to_excel(filename, index=False)
                elif v["source"] == "custom":
                    filename = ensure_custom_file(v["code"])
                else:
                    raise Exception("Неизвестный источник данных.")
                if filename and os.path.exists(filename):
                    files.append(filename)
                else:
                    errors.append(f"{k}: Файл {filename} не создан/не найден.")
            except Exception as e:
                errors.append(f"{k}: {e}")

        summary_file = make_summary_with_links(SAVE_DIR, TICKERS)
        if summary_file and os.path.exists(summary_file):
            files.append(summary_file)

        # Create zip
        zip_path = os.path.join(SAVE_DIR, "archive.zip")
        with zipfile.ZipFile(zip_path, "w") as zf:
            for f in files:
                if f and os.path.exists(f):
                    zf.write(f, arcname=os.path.basename(f))
        await context.bot.send_document(
            chat_id=query.message.chat.id,
            document=open(zip_path, "rb"),
            caption="Архив со всеми тикерами"
        )
        await context.bot.delete_message(
            chat_id=query.message.chat.id,
            message_id=msg.message_id
        )
        await context.bot.send_message(
            chat_id=ADMIN_CHAT_ID,
            text=f"Пользователь {query.from_user.id} запросил архив всех тикеров.\nОшибки: {', '.join(errors) if errors else 'нет'}"
        )
        return

    if query.data == "summary_table":
        msg = await query.edit_message_text("⏳ Формирую сводную таблицу, подождите...")
        await context.bot.send_chat_action(chat_id=query.message.chat.id, action=ChatAction.TYPING)
        try:
            update_all_tickers(
                start_date,
                end_date,
                start_cbr,
                end_cbr,
                start_moex_currency
            )
            df_summary = build_summary_table()
            fname = os.path.join(SAVE_DIR, "summary.xlsx")
            df_summary.to_excel(fname, index=False)
            await context.bot.send_document(
                chat_id=query.message.chat.id,
                document=open(fname, "rb"),
                caption="Сводная таблица (последние значения и динамика)"
            )
            minireport = make_minireport_w_links(df_summary)
            await context.bot.send_message(
                chat_id=query.message.chat.id,
                text=minireport,
                parse_mode="HTML",
                disable_web_page_preview=True
            )
            await context.bot.delete_message(
                chat_id=query.message.chat.id,
                message_id=msg.message_id
            )
        except Exception as e:
            await context.bot.send_message(chat_id=query.message.chat.id, text=f"Ошибка при формировании свода: {e}")
        return

    # --- Один тикер
    for k, v in TICKERS.items():
        if query.data == f'load_{k}':
            msg = await query.edit_message_text(f"⏳ Получаю данные {v['display']}, подождите...")
            await context.bot.send_chat_action(chat_id=query.message.chat.id, action=ChatAction.TYPING)
            try:
                filename = None
                last_date = ""
                if v["source"] == "yahoo":
                    df = fetch_yahoo_closing_history(v["code"], start_date, end_date)
                    filename = os.path.join(SAVE_DIR, f"{k}.xlsx")
                    df.to_excel(filename, index=False)
                    last_date = df["date"].iloc[-1]
                elif v["source"] == "moex":
                    df = fetch_index_closing_table(v["code"], start_date, end_date)
                    filename = os.path.join(SAVE_DIR, f"{k}.xlsx")
                    df.to_excel(filename, index=False)
                    last_date = df["date"].iloc[-1]
                elif v["source"] == "cbr":
                    df = fetch_cbr_currency_history(v["code"], start_cbr, end_cbr)
                    filename = os.path.join(SAVE_DIR, f"{k}.xlsx")
                    df.to_excel(filename, index=False)
                    last_date = df["date"].iloc[-1]
                elif v["source"] == "moex_currency":
                    df = fetch_moex_currency_pair(v["code"], start_moex_currency, end_date)
                    filename = os.path.join(SAVE_DIR, f"{k}.xlsx")
                    df.to_excel(filename, index=False)
                    last_date = df["date"].iloc[-1]
                elif v["source"] == "custom":
                    filename = ensure_custom_file(v["code"])
                    if filename and os.path.exists(filename):
                        df = pd.read_excel(filename)
                        # Для G_CURVE — tradedate, для OFZ10 — date
                        if v["code"] == "gcurves_hist":
                            last_date = df["tradedate"].iloc[-1] if not df.empty and "tradedate" in df.columns else ""
                        else:
                            last_date = df["date"].iloc[-1] if not df.empty and "date" in df.columns else ""
                    else:
                        raise Exception(f"Файл {filename} не создан/не найден.")
                else:
                    raise Exception("Неизвестный источник данных.")

                await context.bot.send_document(
                    chat_id=query.message.chat.id,
                    document=open(filename, "rb"),
                    caption=f"Последние доступные данные: {last_date}"
                )
                await context.bot.delete_message(
                    chat_id=query.message.chat.id,
                    message_id=msg.message_id
                )
                await context.bot.send_message(
                    chat_id=ADMIN_CHAT_ID,
                    text=f"Пользователь {query.from_user.id} запросил таблицу {v['display']} ({start_date} — {end_date}). Последняя дата в данных: {last_date}."
                )
            except Exception as e:
                await context.bot.send_message(chat_id=query.message.chat.id, text=f"Ошибка: {e}")
            break


import pandas as pd
import xlsxwriter
import os

def make_summary_with_links(save_dir, tickers):
    summary_path = os.path.join(save_dir, "summary_linked.xlsx")
    workbook = xlsxwriter.Workbook(summary_path)
    worksheet = workbook.add_worksheet("Сводная")
    headers = ["Тикер", "Последнее значение", "Дата"]
    for col, h in enumerate(headers):
        worksheet.write(0, col, h)

    i = 1
    for k, v in tickers.items():
        file_path = os.path.join(save_dir, f"{k}.xlsx")
        if not os.path.exists(file_path):
            continue
        try:
            df = pd.read_excel(file_path)
            if df.empty:
                continue
            # Найти последнюю строку (последнее значение)
            last_idx = len(df)
            value_cell = f"B{last_idx + 1}"  # B - price, +1 из-за заголовка
            date_cell = f"A{last_idx + 1}"   # A - date

            worksheet.write(i, 0, v.get('display', k))
            worksheet.write_formula(i, 1, f"='{k}.xlsx'!{value_cell}")
            worksheet.write_formula(i, 2, f"='{k}.xlsx'!{date_cell}")
            i += 1
        except Exception as e:
            print(f"Ошибка для {k}: {e}")
            continue

    workbook.close()
    return summary_path


def get_real_proof_url(ticker, date):
    # date — строка "YYYY-MM-DD" или дата
    d = date.strftime('%d.%m.%Y') if isinstance(date, (datetime.date, pd.Timestamp)) else str(date)
    d_us = date.strftime('%Y-%m-%d') if isinstance(date, (datetime.date, pd.Timestamp)) else str(date)
    if ticker in ("IMOEX", "RTSI", "RGBI"):
        # MOEX: страница архива
        return f"https://www.moex.com/ru/index/{ticker}/archive/"
    elif ticker in ("USDRUB", "EURRUB"):
        # CBR: сразу выводит таблицу за эту дату
        code = "R01235" if ticker == "USDRUB" else "R01239"
        return f"https://www.cbr.ru/currency_base/dynamics/?UniDbQuery.date_req1={d}&UniDbQuery.date_req2={d}&UniDbQuery.VAL_NM_RQ={code}"
    elif ticker == "CNYRUB":
        # CNY/RUB через MOEX
        return "https://www.moex.com/ru/issue.aspx?code=CNYRUB_TOM"
    elif ticker.startswith("^") or ticker.endswith("=X") or ticker in ("BZ=F", "GC=F", "BTC-USD", "DX-Y.NYB"):
        code = ticker.replace("^", "%5E")
        # Yahoo: история
        return f"https://finance.yahoo.com/quote/{code}/history?p={code}"
    elif ticker in ("OFZ2", "OFZ10"):
        # МОЕХ: история облигаций
        return "https://www.moex.com/ru/market/bonds/"
    else:
        return None

def make_minireport_w_links(df):
    report = []
    for _, row in df.iterrows():
        display = row['Тикер']
        ticker = row.get("Код", display)

        def date_str(val):
            if pd.isnull(val) or val is None:
                return "-"
            if isinstance(val, str):
                try:
                    return pd.to_datetime(val, dayfirst=True).strftime('%Y-%m-%d')
                except Exception:
                    return val
            elif isinstance(val, (pd.Timestamp, datetime.date, datetime.datetime)):
                return val.strftime('%Y-%m-%d')
            return str(val)

        def pruf_link(val, d):
            dstr = date_str(d)
            url = get_real_proof_url(ticker, d)
            if url and val != "-":
                return f'<a href="{url}">{val}</a> ({dstr})'
            else:
                return f"{val} ({dstr})"

        txt = (
            f"<b>{display}</b>\n"
            f"  Последнее значение: {pruf_link(row['Значение'], row.get('Дата_последнего'))}\n"
            f"  Неделя назад: {pruf_link(row.get('Неделя'), row.get('Дата_недели'))}\n"
            f"  С нач. квартала: {pruf_link(row.get('Квартал'), row.get('Дата_квартала'))}\n"
            f"  С нач. года: {pruf_link(row.get('Год'), row.get('Дата_года'))}\n"
        )
        report.append(txt)
    return "\n\n".join(report)

def main():
    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler("load", load_command))
    application.add_handler(CallbackQueryHandler(handle_ticker_press))
    application.run_polling()

if __name__ == "__main__":
    main()
