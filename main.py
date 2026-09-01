
import pandas as pd
import yfinance as yf
import requests
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

# הגדרת User-Agent למניעת חסימות מצד Yahoo Finance
session = requests.Session()
session.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'


# --- הגדרות אימייל (מלא את הפרטים שלך) ---
EMAIL_ADDRESS = os.environ.get('EMAIL_USER')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS')
RECIPIENT_EMAIL = "bf669907@gmail.com"  # למי לשלוח את הדוח


def send_email(report_df):
    """שליחת טבלת הניתוח כקובץ HTML מעוצב למייל"""
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = RECIPIENT_EMAIL
    msg['Subject'] = f"דוח ניתוח מניות יומי - {pd.Timestamp.now().strftime('%d/%m/%Y')}"

    # עיצוב הטבלה כ-HTML
    html_table = report_df.to_html(index=False, justify='center', border=1)

    body = f"""
    <html>
      <body dir="rtl">
        <h2>דוח סטטוס מניות שבועי</h2>
        <p>להלן הניתוח עבור המניות ברשימה שלך:</p>
        {html_table}
        <br>
        <p>הדוח הופק באופן אוטומטי על ידי ה-AI Agent שלך.</p>
      </body>
    </html>
    """

    msg.attach(MIMEText(body, 'html'))

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("המייל נשלח בהצלחה!")
    except Exception as e:
        print(f"שגיאה בשליחת המייל: {e}")




def get_stocks_from_sheets(sheets_url):
    """קריאת רשימת מניות מגיליון גוגל שפורסם כ-CSV"""
    try:
        response = requests.get(sheets_url)
        response.raise_for_status()

        # קריאת ה-CSV לתוך DataFrame
        # header=None אם אין שורת כותרת, או header=0 אם יש
        df_sheets = pd.read_csv(io.StringIO(response.text))

        # שליפת עמודה 2 (אינדקס 1 בפייתון)
        # אנחנו מנקים רווחים ומוודאים שזה טקסט
        symbols = df_sheets.iloc[:, 1].dropna().astype(str).str.strip().tolist()

        return symbols
    except Exception as e:
        print(f"שגיאה בגישה לגיליון: {e}")
        return []




def analyze_portfolio(symbols):
    results = []

    # הגדרת סשן עם User-Agent כדי למנוע חסימת בוטים מ-GitHub Actions
    session = requests.Session()
    session.headers[
        'User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'

    for sym in symbols:
        ticker = f"{sym}.TA" if not sym.endswith(".TA") else sym

        try:
            # הורדת נתונים מורחבת (15 ימים)
            df = yf.download(ticker, period="15d", session=session, progress=False)

            # אם התקבל DataFrame ריק
            if df.empty:
                print(f"לא נשלפו נתונים עבור {ticker}")
                continue

            # טיפול במבנה עמודות חדש של yfinance (לפעמים העמודות הן MultiIndex)
            if isinstance(df.columns, pd.MultiIndex):
                close_series = df['Close'][ticker] if ticker in df['Close'] else df['Close'].iloc[:, 0]
            else:
                close_series = df['Close']

            # ניקוי ערכי NaN
            close_series = close_series.dropna()

            if len(close_series) < 7:
                print(f"אין מספיק ימי מסחר תקינים עבור {ticker} (נמצאו {len(close_series)})")
                continue

            # שליפת מחירים
            price_now = float(close_series.iloc[-1])
            price_7_days_ago = float(close_series.iloc[-7])

            # חישוב אחוז שינוי
            pct_change = ((price_now - price_7_days_ago) / price_7_days_ago) * 100

            # בדיקת רצף עליות של 3 ימים בתוך ה-7 האחרונים
            daily_diff = close_series.tail(8).diff() > 0
            is_up_streak = daily_diff.tail(7).rolling(window=3).apply(lambda x: x.all()).max() == 1

            # סטטוס ואימוג'ים
            status_parts = []
            if pct_change > 5:
                status_parts.append("📈 עליה מעל 5%")
            elif pct_change < -5:
                status_parts.append("📉 ירידה מעל 5%")
            else:
                status_parts.append("ניטרלי")

            if is_up_streak:
                status_parts.append("🔥 רצף 3 ימי עליות!")

            final_status = " | ".join(status_parts)

            results.append({
                'מניה': sym,
                'שינוי שבועי': f"{pct_change:.2f}%",
                'סטטוס': final_status,
                'רצף 3 ימים': "כן ✅" if is_up_streak else "לא"
            })

        except Exception as e:
            print(f"שגיאה בניתוח המניה {sym}: {e}")

    return pd.DataFrame(results)

# --- הרצה ---
# החלף את ה-URL בקישור ה-CSV שקיבלת מ-"Publish to web"
SHEETS_CSV_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSwRcF6OGLtNkZK9QldRT9xeeC-eQ-2uF2Jef6naqGbO1H9s9rXPF4pX2r9D683mh9JP729qX7X_2vw/pub?gid=0&single=true&output=csv"

stock_symbols = get_stocks_from_sheets(SHEETS_CSV_URL)
if stock_symbols:
    final_report = analyze_portfolio(stock_symbols)
    print(final_report.to_string(index=False))

    # כאן תוכל להוסיף את הפונקציה send_email(final_report) שכתבנו קודם
    send_email(final_report)