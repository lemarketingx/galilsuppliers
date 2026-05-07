# Galil Engineering Suite V7

גרסה מתוקנת לפריסה ב-Vercel.

מה תוקן:
- הוסר package-lock.json כדי למנוע כשל npm install ב-Vercel בגלל registry פנימי.
- שמירה אוטומטית ל-localStorage אחרי טעינה תקינה.
- מניעת כפילויות בהעלאת BOQ וספקים.
- העלאת Excel ממזגת נתונים קיימים במקום להחליף.
- גיבוי ושחזור JSON.
- מדריך שימוש מובנה.
- אישור לפני איפוס נתונים.

## הפעלה מקומית
npm install
npm run dev

## פריסה ב-Vercel
Framework: Vite
Build command: npm run build
Output directory: dist
