// 1. בוחרים את הטופס לפי ה-ID שלו
const form = document.getElementById("emailForm");

// 2. מאזינים לאירוע "submit" של הטופס
form.addEventListener("submit", function(event) {
    // 3. מונעים את ההתנהגות הרגילה של הטופס (reload של הדף)
    event.preventDefault();

    // 4. אוספים את הערכים שהמשתמש הזין בכל שדה
    const subject = document.getElementById("subject").value;
    const recipients = document.getElementById("recipients").value;
    const body = document.getElementById("body").value;
    const fileInput = document.getElementById("attachment");

    // 5. יוצרים אובייקט FormData כדי לשלוח נתונים + קבצים
    const formData = new FormData();
    formData.append("subject", subject);
    formData.append("recipients", recipients); // רשימה מופרדת בפסיקים
    formData.append("body", body);

    // 6. אם המשתמש העלה קובץ, מוסיפים אותו ל-FormData
    if (fileInput.files.length > 0) {
        formData.append("attachment", fileInput.files[0]);
    }

    // 7. שולחים את הנתונים ל-Python דרך HTTP POST
fetch("/create_drafts", {
    method: "POST",
    body: formData
})
.then(response => response.json())
.then(data => {
    console.log("Response from Python:", data);
    alert("טיוטות נוצרו בהצלחה ב-Outlook!");
})
.catch(error => {
    console.error("Error:", error);
    alert("אירעה שגיאה ביצירת הטיוטות.");
});
});
