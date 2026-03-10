// וידוא שהסביבה של אופיס מוכנה
Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // אם הדף כבר נטען, נריץ מיד. אם לא, נחכה לטעינה.
        if (document.readyState === "complete" || document.readyState === "interactive") {
            renderButtons();
        } else {
            document.addEventListener("DOMContentLoaded", renderButtons);
        }
    }
});

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

function renderButtons() {
    const list = document.getElementById("button-list");
    if (!list) return;

    // ניקוי הרשימה לפני הוספה (למניעת כפילויות)
    list.innerHTML = "";

    requestTypes.forEach(type => {
        const btn = document.createElement("button");
        btn.className = "request-btn";
        btn.innerHTML = `<span>${type.label}</span>`;
        btn.onclick = () => openEmailForm(type);
        list.appendChild(btn);
    });
}

function openEmailForm(type) {
    let tableRows = type.questions.map(q => `
        <tr>
            <td style="border: 1px solid #cccccc; padding: 10px; background-color: #f2f2f2; width: 40%; font-weight: bold; text-align: right;">${q}:</td>
            <td style="border: 1px solid #cccccc; padding: 10px; text-align: right;"></td>
        </tr>
    `).join("");

    const htmlBody = `
        <div dir="rtl" style="font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; text-align: right;">
            <p>שלום רב,</p>
            <p>להלן פרטי בקשה בנושא: <strong>${type.label.replace(/[^א-ת\s]/g, '').trim()}</strong></p>
            <table dir="rtl" style="border-collapse: collapse; width: 100%; max-width: 650px; border: 1px solid #cccccc;">
                <thead>
                    <tr style="background-color: #0078d4; color: white;">
                        <th style="padding: 10px; border: 1px solid #cccccc; text-align: right;">שדה</th>
                        <th style="padding: 10px; border: 1px solid #cccccc; text-align: right;">מידע למילוי</th>
                    </tr>
                </thead>
                <tbody>
                    ${tableRows}
                </tbody>
            </table>
            <br>
            <p style="color: #666;">המייל נשלח באמצעות תוסף "InfoSec".</p>
        </div>
    `;

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["InfoSec@your-org.com"], 
        subject: type.subject,
        htmlBody: htmlBody
    });
}
