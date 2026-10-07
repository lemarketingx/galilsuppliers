# פריסה לשרת פנימי

המערכת היא אפליקציית React סטטית. לאחר שלב הבנייה אין צורך בתהליך Node.js קבוע.

## דרישות

- Node.js 20.19 ומעלה לצורך Build בלבד
- npm
- שרת קבצים סטטיים: IIS, Nginx או Apache
- HTTPS או גישה דרך רשת פנימית מאובטחת / VPN

## Build רגיל

```bash
npm ci
npm run build
```

יש לפרסם את תוכן התיקייה `dist` בשורש האתר. אין לפרסם את קוד המקור, `node_modules` או את תיקיית `.git`.

## Docker

```bash
docker build -t galil-engineering-suite:5.1.0 .
docker run --name galil-engineering-suite -p 8080:80 galil-engineering-suite:5.1.0
```

בדיקת תקינות:

```text
http://SERVER:8080/healthz
```

## IIS

1. להריץ `npm ci` ולאחר מכן `npm run build` במחשב Build.
2. להעתיק לשרת רק את תוכן `dist`.
3. להגדיר את `index.html` כמסמך ברירת המחדל.
4. להגדיר HTTPS, או להגביל את האתר לרשת החברה בלבד.
5. למנוע Directory Browsing.

## מגבלה חשובה

הגרסה הנוכחית שומרת פרויקטים וספקים ב-`localStorage` של הדפדפן. לכן הנתונים אינם משותפים בין עובדים ואינם מגובים בשרת. לפני שימוש רב-משתמשים יש להוסיף API, מסד נתונים, אחסון מסמכים והזדהות ארגונית.

