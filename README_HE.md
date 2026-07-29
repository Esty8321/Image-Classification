# מערכת Web לבדיקת תמונות

## מה המערכת עושה

1. המשתמש המורשה מתחבר לחשבון Google.
2. מעלה קובץ XLSX.
3. השרת קורא את עמודות B, E ו-F בהתאם לקוד המקורי.
4. השרת מוריד כל תמונה ושולח אותה ל-endpoint.
5. התוצאות מתווספות לקובץ היסטוריה מקומי.
6. נוצר Google Sheet חדש בחשבון Google המחובר.
7. ניתן למחוק את ההיסטוריה המקומית דרך המסך.

## מדוע זו אינה מערכת Client Side בלבד

הדפדפן אינו מקום בטוח לשמירת:
- כתובת ושיטת התקשורת עם endpoint פנימי.
- סודות OAuth של Google.
- token מתמשך של Google.
- קובץ CSV משותף הנשמר בין הרצות.

בנוסף, דפדפן עלול לחסום הורדת תמונות מאתרים אחרים בגלל CORS. לכן נדרש שרת Python קטן.

## התקנה

```bash
cd image_classifier_web
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

ערכי חובה בקובץ `.env`:

```env
FLASK_SECRET_KEY=מחרוזת-אקראית-ארוכה
ALLOWED_GOOGLE_EMAIL=כתובת-החשבון-המורשה
GOOGLE_REDIRECT_URI=http://127.0.0.1:5000/oauth2callback
```

אפשר ליצור secret כך:

```bash
python -c "import secrets; print(secrets.token_hex(32))"
```

## הגדרת Google Cloud

1. ליצור Project ב-Google Cloud.
2. להפעיל Google Sheets API וגם Google Drive API.
3. להגדיר OAuth consent screen.
4. ליצור OAuth Client מסוג Web application.
5. להוסיף Authorized redirect URI:
   `http://127.0.0.1:5000/oauth2callback`
6. להוריד את קובץ ה-OAuth, לשנות את שמו ל-`client_secret.json`,
   ולשמור אותו בתיקיית הפרויקט.

## הפעלה

```bash
python app.py
```

פתיחה בדפדפן:

```text
http://127.0.0.1:5000
```

## שימוש ממחשב אחר ברשת

בקובץ `.env`:

```env
FLASK_HOST=0.0.0.0
GOOGLE_REDIRECT_URI=http://IP-OF-SERVER:5000/oauth2callback
```

יש להוסיף את אותה כתובת Redirect גם ב-Google Cloud.

להפעלה יציבה יותר:

```bash
gunicorn --workers 1 --threads 4 --bind 0.0.0.0:5000 app:app
```

נשמר worker אחד בלבד, משום שקובץ ההיסטוריה הוא קובץ CSV מקומי.

## קבצים חשובים

- `app.py` — ממשק Web, OAuth, העלאה, מחיקת היסטוריה.
- `classifier_core.py` — הלוגיקה המקורית, עם אפשרות לקבל Google client מחובר.
- `data/image_results_history.csv` — ההיסטוריה.
- `data/image_results_history_backup.json` — גיבוי JSON.
- `data/google_token.json` — token של החשבון המחובר.
- `uploads/` — קבצים זמניים; נמחקים אחרי הרצה.

## אבטחה

- אין להעלות ל-Git את `.env`, `client_secret.json` או `data/google_token.json`.
- מומלץ להפעיל את המערכת רק ברשת פנימית או מאחורי HTTPS.
- `ALLOWED_GOOGLE_EMAIL` מגביל את השימוש לחשבון אחד.
- כפתור מחיקת ההיסטוריה דורש הקלדת `DELETE`.
