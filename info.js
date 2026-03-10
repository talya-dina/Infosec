function initializeApp() {
    renderButtons();
}

Office.onReady((info) => {
    if (info.host) {
        initializeApp();
    }
});

// מאפשר תצוגה גם ב-GitHub
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
        btn.innerHTML = `<span>${type.label}</span>`;
        btn.onclick = () => openNewEmail(type);
        list.appendChild(btn);
    });
}

function openNewEmail(type) {
    if (typeof Office === 'undefined' || !Office.context || !Office.context.mailbox) {
        alert("הפעולה זמינה רק מתוך Outlook.");
        return;
    }

    // יצירת Timestamp/ID ייחודי
    const uniqueId = Date.now();
    const fullSubject = `OFIRSEC Security (ID: ${uniqueId}) - ${type.subject}`;
    
    const tableHtml = generateCyberTable(type);

    Office.context.mailbox.displayNewMessageForm({
        toRecipients: ["info@ofirsec.co.il"],
        subject: fullSubject,
        htmlBody: tableHtml
    });
}

function generateCyberTable(type) {
    const rows = type.questions.map(q => `
        <tr>
            <td style="border: 1px solid #e0e0e0; padding: 15px; background-color: #fcfcfc; color: #333333; font-weight: bold; width: 40%; text-align: right; font-size: 14px;">${q}:</td>
            <td style="border: 1px solid #e0e0e0; padding: 15px; background-color: #ffffff; text-align: right;"></td>
        </tr>
    `).join("");

    return `
        <div dir="rtl" style="font-family: 'Segoe UI', Tahoma, Arial, sans-serif; max-width: 650px; color: #333; line-height: 1.6; margin: 0 auto; text-align: right;">
            <div style="background-color: #f8f9fa; padding: 20px; border: 1px solid #e0e0e0; border-bottom: 5px solid #0078d4; border-radius: 8px 8px 0 0;">
                <h2 style="margin: 0; font-size: 20px; color: #0078d4;">טופס בקשה: ${type.label.replace(/[^\u0590-\u05FF\s]/g, '').trim()}</h2>
                <p style="margin: 5px 0 0 0; font-size: 13px; color: #666;">אנא מלא את הפרטים בטבלה מטה והשב למייל זה.</p>
            </div>
            
            <table style="width: 100%; border-collapse: collapse; margin-top: 0; border: 1px solid #e0e0e0;">
                ${rows}
            </table>

            <div style="margin-top: 30px; text-align: center; border-top: 1px solid #eeeeee; padding-top: 20px;">
                <p style="margin: 0; font-size: 16px; font-weight: bold; color: #001529;">תודה רבה על שיתוף הפעולה!</p>
                <p style="margin: 5px 0 15px 0; color: #0078d4; font-size: 14px;">צוות אבטחת מידע OFIRSEC</p>
                <img src="https://ofirsec.co.il/wp-content/uploads/2024/06/logo-big-cyber-1-1-768x336.png" alt="OFIRSEC Logo" style="width: 250px; height: auto;">
            </div>
        </div>
    `;
}
