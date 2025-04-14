```markdown
# Bulk WhatsApp Message

Automate sending WhatsApp messages in bulk using data from a Google Sheet.

## 🚀 Features

- Google Sheets integration
- Node.js backend to serve WhatsApp message triggers
- Public endpoint using VS Code port forwarding
- Custom Google Apps Script to trigger messages

---

## 📋 Setup Instructions

### 1. Create a Google Sheet Service
- Go to [Google Cloud Console](https://console.cloud.google.com/).
- Create a new project.
- Enable the Google Sheets API.
- Generate a **Service Account** key and download the JSON file.
- Place the JSON credential file inside the project folder.

### 2. Configure `index.js`
- Open `index.js` and update the following:
  - `sheetId` → Your Google Sheet ID
  - `sheetName` → Your Google Sheet tab name
  - `credentialsFile` → Your downloaded JSON file name

### 3. Add Google Apps Script to Sheet
- In your Google Sheet, go to `Extensions` > `Apps Script`.
- Add the provided script (see below) to handle sending requests.
- Save and deploy as needed.

### 4. Run the Node Server
```bash
node index.js
```

- Make sure the server is running locally.

### 5. Create Public Endpoint
- Use **VS Code Remote Explorer** or **Port Forwarding** feature.
- Forward the port used by the Node server.
- Set port visibility to **public**.

### 6. Add Public URL to Sheet
- Copy the generated **public URL**.
- Paste it into the **last column** of your Google Sheet for each row/message.

---

## 🛠 Example Google Apps Script

```javascript
function sendWhatsAppMessages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const name = data[i][0];
    const number = data[i][1];
    const message = data[i][2];
    const endpoint = data[i][3]; // public URL

    const payload = {
      name: name,
      number: number,
      message: message,
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
    };

    UrlFetchApp.fetch(endpoint, options);
  }
}
```

---

## 🧾 License

MIT License

---

## 🤝 Contributions

Feel free to open issues or submit PRs to enhance the project.
```
