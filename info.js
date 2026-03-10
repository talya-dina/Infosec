function initializeApp() {
    renderButtons();
}

Office.onReady((info) => {
    if (info.host) {
        initializeApp();
    }
});

if (!window.officeInitialized && (window.location.host.includes('github.io') || window.location.host.includes('localhost'))) {
    if (document.readyState === "complete" || document.readyState === "interactive") {
        initializeApp();
    } else {
        document.addEventListener("DOMContentLoaded", initializeApp);
    }
}

const requestTypes = [
    { id: "internet", label: "🌐 יציאה מיוחדת לאינטרנט", subject: "בקשה ליציאה מיוחדת של רכיב לאינטרנט", questions: ["שם הרכיב", "כתובות IP", "רשימת אתרים/יעד", "תיאור הצורך", "פורט נדרש", "הסבר לפורט"] },
    { id: "privileges", label: "🔑 הרשאות פריבילגיות", subject: "בקשה למתן הרשאות פריבילגיות ברשת", questions: ["מטרת ההרשאה", "קבוצות נדרשות", "מייל המשתמש", "תפקיד", "סביבה (ענן/פנים)"] },
    { id: "generic", label: "👤 משתמש גנרי / סרביס", subject: "בקשה לפתיחת משתמש גנרי", questions: ["שם המשתמש", "סיבת הפתיחה", "הרשאות נדרשות", "מערכות יעד"] },
    { id: "software", label: "📦 תוכנה / מערכת חדשה", subject: "בקשה לתוכנה / מערכת חדשה", questions: ["שם התוכנה", "שם החברה", "קישור לאתר", "סוג התקנה"] },
    { id: "gritta", label: "♻️ תיעוד גריטה", subject: "תיעוד גריטה - אבטחת מידע", questions: ["תאריך", "מבצע הגריטה", "מקום", "רכיב", "המידע שנמחק"] },
    { id: "survey", label: "📝 פרסום טופס או סקר", subject: "בקשה לפרסום טופס או סקר", questions: ["מחלקה", "תאריך פתיחה/סגירה", "שם הסקר", "נתונים שיאספו", "מערכת (Forms/Google)"] },
    { id: "general", label: "🛡️ אישור כללי אחר", subject: "בקשה לאישור כללי - אבטחת מידע", questions: ["פירוט האישור", "צורך טכני/עסקי"] }
];

function renderButtons() {
    const list = document.getElementById("button-list");
    if (!list) return;
    list.innerHTML = "";
    requestTypes.forEach(type => {
        const btn = document.createElement("button");
        btn.className = "request-btn";
        // הוספת החץ כאן
        btn.innerHTML = `<span class="btn-text">${type.label}</span><span class="btn-arrow">←</span>`;
        btn.onclick = () => handleAction(type);
        list.appendChild(btn);
    });
}

async function handleAction(type) {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) {
        alert("הפעולה זמינה רק מתוך Outlook.");
        return;
    }

    const tableHtml = generateCyberTable(type);

    // בדיקה אם אנחנו במצב כתיבת מייל (Compose) או קריאה (Read)
    if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message && Office.context.mailbox.item.displayNewMessageForm === undefined) {
        // אנחנו בתוך מייל פתוח - נזריק את הטבלה
        Office.context.mailbox.item.body.setSelectedDataAsync(tableHtml, { coercionType: Office.CoercionType.Html }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                // אם ההזרקה נכשלה (למשל המייל לא במצב עריכה), נפתח מייל חדש
                createNewEmail(type, tableHtml);
            }
        });
    } else {
        // אנחנו לא בתוך מייל - נפתח חלון חדש
        createNewEmail(type, tableHtml);
    }
}

function createNewEmail(type, htmlBody) {
    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["info@ofirsec.co.il"],
        subject: type.subject,
        htmlBody: htmlBody
    });
}

function generateCyberTable(type) {
    const rows = type.questions.map(q => `
        <tr>
            <td style="border: 1px solid #1a2a3a; padding: 12px; background-color: #f8f9fa; color: #1a2a3a; font-weight: bold; width: 35%;">${q}</td>
            <td style="border: 1px solid #1a2a3a; padding: 12px; background-color: #ffffff;"></td>
        </tr>
    `).join("");

    return `
        <div dir="rtl" style="font-family: 'Segoe UI', Tahoma, sans-serif; max-width: 600px; color: #1a2a3a; line-height: 1.6;">
            <div style="background-color: #001529; color: #ffffff; padding: 15px; border-radius: 8px 8px 0 0; border-bottom: 4px solid #0078d4;">
                <h2 style="margin: 0; font-size: 18px;">בקשה בנושא: ${type.label.replace(/[^א-ת\s]/g, '').trim()}</h2>
            </div>
            <table style="width: 100%; border-collapse: collapse; border: 1px solid #1a2a3a;">
                ${rows}
            </table>
            <div style="margin-top: 20px; padding-top: 15px; border-top: 1px solid #eeeeee;">
                <p style="margin: 0; font-weight: bold; color: #001529;">תודה רבה על שיתוף הפעולה!</p>
                <p style="margin: 5px 0; color: #0078d4;">צוות אבטחת מידע OFIRSEC</p>
                <img src="https://ofirsec.co.il/wp-content/uploads/2024/06/logo-big-cyber-1-1-768x336.png" alt="OFIRSEC Logo" style="width: 180px; margin-top: 10px;">
            </div>
        </div>
    `;
}
