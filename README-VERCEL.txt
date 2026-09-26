הגדרה חד-פעמית ב-Vercel
======================

1. העלה ל-GitHub את:
   - index.html
   - Code.gs
   - api/order.js

2. ב-Vercel פתח:
   Project > Settings > Environment Variables

3. צור משתנה:
   Name: GOOGLE_SCRIPT_URL
   Value: כתובת ה-Web App של Google Apps Script שמסתיימת ב-/exec

4. סמן Production (ומומלץ גם Preview), שמור ובצע Redeploy.

חשוב:
- אין צורך לשים יותר את כתובת Google Apps Script בתוך index.html.
- Code.gs נשאר בצד של Google Apps Script ולא רץ ב-Vercel.
- אם שינית Code.gs, יש לבצע Deploy > Manage deployments > Edit > New version > Deploy ב-Google Apps Script.
