const fs = require('fs');
const makeWASocket = require('@whiskeysockets/baileys').default;
const { DisconnectReason, useMultiFileAuthState, fetchLatestBaileysVersion } = require('@whiskeysockets/baileys');
const express = require('express');
const cors = require('cors');
const { google } = require("googleapis");
const { JWT } = require("google-auth-library");

const SPREADSHEET_ID = "1pmt2xvKsynM3C3WLvGydBXO6BfkqQKAOnZpTc8Z8NoI";
const SHEET_NAME = "CurrentSheet";
const AUTH_JSON = require("./whatsappservice-452312-5a0dead8da2b.json");

const auth = new JWT({
    email: AUTH_JSON.client_email,
    key: AUTH_JSON.private_key,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"]
});
const sheets = google.sheets({ version: "v4", auth });

const AUTH_FOLDER = 'auth';
const store = {};
let sock;

const app = express();
app.use(express.json());
app.use(cors());

async function connectToWhatsApp() {
    try {
        if (!fs.existsSync(AUTH_FOLDER)) {
            fs.mkdirSync(AUTH_FOLDER);
        }

        const { state, saveCreds } = await useMultiFileAuthState(AUTH_FOLDER);
        const { version } = await fetchLatestBaileysVersion();

        sock = makeWASocket({
            printQRInTerminal: true,
            auth: state,
            version
        });

        sock.ev.process(async (events) => {
            if (events['connection.update']) {
                const { connection, lastDisconnect } = events['connection.update'];

                if (connection === 'close') {
                    const shouldReconnect = lastDisconnect?.error?.output?.statusCode !== DisconnectReason.loggedOut;
                    if (!shouldReconnect) {
                        console.log('User logged out, deleting auth folder...');
                        fs.rmSync(AUTH_FOLDER, { recursive: true, force: true });
                    }
                    console.log('Reconnecting in 5 seconds...');
                    setTimeout(connectToWhatsApp, 5000);
                    return;
                }

                if (connection === 'open') {
                    console.log("✅ WhatsApp connected!");
                }
            }

            if (events['creds.update']) {
                await saveCreds();
            }
        });

    } catch (error) {
        console.error("Error in connectToWhatsApp:", error.message);
        console.log("Retrying in 5 seconds...");
        setTimeout(connectToWhatsApp, 5000);
    }
}

// Batch update Google Sheets status by row number
async function batchUpdateSheetStatus(updates) {
    if (updates.length === 0) return;

    try {
        let updateRequests = updates.map(({ row, status }) => ({
            range: `${SHEET_NAME}!D${row}`,
            values: [[status]]
        }));

        await sheets.spreadsheets.values.batchUpdate({
            spreadsheetId: SPREADSHEET_ID,
            requestBody: {
                valueInputOption: "RAW",
                data: updateRequests
            }
        });

        console.log(`Batch update completed for ${updates.length} rows.`);
    } catch (error) {
        console.error("Error in batch update:", error.message);
    }
}

// Process messages in BATCHES of 5
async function processMessages(messages) {
    const batchSize = 5;
    
    for (let i = 0; i < messages.length; i += batchSize) {
        const batch = messages.slice(i, i + batchSize);
        let updates = [];

        // Mark batch as "Processing"
        batch.forEach(({ row }) => updates.push({ row, status: "Processing" }));
        await batchUpdateSheetStatus(updates);

        // Process each message in batch with delays
        await Promise.all(batch.map(async (entry) => {
            const { row, phone, message } = entry;
            if (!row || !phone || !message) return;

            const formattedPhone = phone.startsWith("91") ? `${phone}@s.whatsapp.net` : `91${phone}@s.whatsapp.net`;

            // Random delay (3-5 sec)
            const delay = Math.floor(Math.random() * (5000 - 3000 + 1)) + 3000;
            await new Promise(resolve => setTimeout(resolve, delay));

            try {
                const [exists] = await sock.onWhatsApp(formattedPhone);
                if (!exists) {
                    updates.push({ row, status: "Number Not Registered" });
                } else {
                    await sock.sendMessage(formattedPhone, { text: message });
                    updates.push({ row, status: "Sent" });
                }
            } catch (e) {
                updates.push({ row, status: "Failed" });
            }
        }));

        // Update status after processing the batch
        await batchUpdateSheetStatus(updates);
    }
}


// API Endpoint to receive messages & update sheet status
app.post("/send-messages", async (req, res) => {
    console.log("Received request:", req.body);

    if (!req.body || !req.body.data || !Array.isArray(req.body.data)) {
        console.error("Invalid request format:", req.body);
        return res.status(400).json({ error: "Invalid request format. Expected { data: [...] }" });
    }

    let processingUpdates = [];

    req.body.data.forEach((entry) => {
        if (entry.row && entry.phone && entry.message) {
            processingUpdates.push({ row: entry.row, status: "Processing" });
        }
    });

    await batchUpdateSheetStatus(processingUpdates);
    processMessages(req.body.data);

    res.json({ success: true, message: "Messages are being processed." });
});


// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server running on port ${PORT}`);
});

// Start the WhatsApp connection
connectToWhatsApp();
