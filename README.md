# Financial Data Bot

A Telegram bot that fetches data for various financial tickers and generates Excel reports. It can gather data from MOEX, Yahoo Finance, and the Russian Central Bank. The bot also builds summary tables and packages all generated files in a single archive.

## Configuration

Before running the bot, provide your Telegram bot token and the administrator chat ID in `config.py`:

```python
TOKEN = "YOUR_TELEGRAM_BOT_TOKEN"
ADMIN_CHAT_ID = 123456789
```

The bot uses a `data` directory in the project root to store downloaded Excel files and logs. It will be created automatically if not present.

## Usage

Install the required dependencies and run the bot:

```bash
pip install -r requirements.txt
python main.py
```

Use `/load` in your Telegram bot chat to start interacting with the bot.

## Files

- `main.py` – main bot logic
- `gcurves_hist.py` – retrieves yield curve history from MOEX
- `ofz_2.py` – extracts 2‑year bond yields from `gcurves_hist.xlsx`
- `ofz_10.py` – extracts 10‑year bond yields from `gcurves_hist.xlsx`
- `config.py` – contains `TOKEN` and `ADMIN_CHAT_ID`

## Logs

The application writes logs to `imoex.log` and module specific log files such as `gcurves_hist.log`. These are useful for debugging issues with data retrieval.
