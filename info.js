Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        renderButtons();
    }
});

// ריכוז כל הנתונים מהמסמך שלך
const requestTypes = [
    {
        id: "internet",
        label: "🌐 יציאה מיוחדת לאינטרנט",
        subject: "בקשה ליציאה מיוחדת של רכיב לאינטרנט",
        questions: [
            "שם הרכיב שנדרש לצאת לאינטרנט",
            "כתובות ה-IP של הרכיב (במידה וקיימת כתובת קבועה)",
            "רשימת אתרים / כתובות IP יעד (נימוק אם נדרש הכל)",
            "תיאור הצורך ביציאה לאינטרנט",
            "הפורט הנדרש ליציאה לעולם",
            "הסבר על סיבת הפורט הספציפי"
        ]
    },
    {
        id: "privileges",
        label: "🔑 הרשאות פריבילגיות",
        subject: "בקשה למתן הרשאות פריבילגיות ברשת",
        questions: [
            "מטרת ההרשאה והגדרות התפקיד",
            "לאיזה קבוצות חזקות המשתמש יצטרף",
            "מייל המשתמש",
            "תפקיד",
            "באיזה סביבה המשתמש ימצא (פנימית/עננית/אפליקטיבית)"
        ]
    },
    {
        id: "generic",
        label: "👤 משתמש גנרי / סרביס",
        subject: "בקשה לפתיחת משתמש גנרי (לרבות סרביס)",
        questions: [
            "שם המשתמש הגנרי",
            "תפקיד המשתמש / סיבה לפתיחתו",
            "הרשאות נדרשות",
            "עמדות או מערכות אליהם יהיה רשאי לגשת"
        ]
    },
    {
        id: "software",
        label: "📦 תוכנה / מערכת חדשה",
        subject: "בקשה לתוכנה / תוסף / מערכת חדשה",
        questions: [
            "שם התוכנה",
            "שם החברה",
            "קישור לאתר הרלוונטי",
            "סוג התוכנה (מקומית / עננית / לא ידוע)"
        ]
    },
    {
        id: "gritta",
        label: "♻️ תיעוד גריטה",
        subject: "תיעוד גריטה - אבטחת מידע",
        questions: [
            "תאריך ביצוע הגריטה",
            "שם מבצע הגריטה",
            "מקום ביצוע הגריטה",
            "הרכיב שבוצעה עליו גריטה",
            "המידע שהיה על הרכיב"
        ]
    },
    {
        id: "survey",
        label: "📝 פרסום טופס או סקר",
        subject: "בקשה לפרסום טופס או סקר",
        questions: [
            "שם המחלקה",
            "תאריך פתיחה",
            "תאריך סגירה",
            "שם הטופס / סקר",
            "הנתונים שיאספו",
            "כתובת הטופס",
            "על איזו מערכת הוא מתבסס (Forms, Google וכו')"
        ]
    },
    {
        id: "general",
        label: "🛡️ אישור כללי אחר",
        subject: "בקשה לאישור כללי - אבטחת מידע",
        questions: [
            "פירוט פרטי האישור המבוקש",
            "הצורך העסקי / טכני"
        ]
    }
];

// פונקציה ליצירת הכפתורים בחלונית
function renderButtons() {
    const list = document.getElementById("button-list");
    if (!list) return;

    requestTypes.forEach(type => {
        const btn = document.createElement("button");
        btn.className = "request-btn"; // משתמש ב-CSS החדש שכתבנו
        btn.innerHTML = `<span>${type.label}</span>`;
        
        btn.onclick = () => openEmailForm(type);
        list.appendChild(btn);
    });
}

// פונקציה לפתיחת המייל עם הטבלה
function openEmailForm(type) {
    // בניית שורות הטבלה על בסיס השאלות של אותו נושא
    let tableRows = type.questions.map(q => `
        <tr>
            <td style="border: 1px solid #cccccc; padding: 10px; background-color: #f2f2f2; width: 40%; font-weight: bold;">${q}:</td>
            <td style="border: 1px solid #cccccc; padding: 10px;"></td>
        </tr>
    `).join("");

    // בניית גוף המייל המלא ב-HTML
    const htmlBody = `
        <div dir="rtl" style="font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6;">
            <p>שלום רב,</p>
            <p>להלן פרטי בקשה בנושא: <strong>${type.label.replace(/[^א-ת\s]/g, '').trim()}</strong></p>
            <table style="border-collapse: collapse; width: 100%; max-width: 650px; border: 1px solid #cccccc; text-align: right;">
                <thead>
                    <tr style="background-color: #0078d4; color: white;">
                        <th style="padding: 10px; border: 1px solid #cccccc;">שדה</th>
                        <th style="padding: 10px; border: 1px solid #cccccc;">מידע למילוי</th>
                    </tr>
                </thead>
                <tbody>
                    ${tableRows}
                </tbody>
            </table>
            <br>
            <p style="color: #666;">המייל נשלח באמצעות תוסף "פורמטים לאבטחת מידע".</p>
        </div>
    `;

    // פקודת ה-Office לפתיחת חלון הודעה חדשה
    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["InfoSec@your-org.com"], // עדכני כאן את כתובת המייל שלכם
        subject: type.subject,
        htmlBody: htmlBody
    });
}