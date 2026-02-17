const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
require("dotenv").config();

const SCOPES = ["https://www.googleapis.com/auth/drive.file"];
const TOKEN_PATH = path.join(__dirname, "../token.json");
const CREDENTIALS_PATH = path.join(__dirname, "../credentials.json");
const BACKUP_FOLDER_ID = process.env.GOOGLE_DRIVE_BACKUP_FOLDER_ID;

function loadCredentials() {
    return JSON.parse(fs.readFileSync(CREDENTIALS_PATH));
}

async function authorize() {
    const { installed } = loadCredentials();

    const oAuth2Client = new google.auth.OAuth2(
        installed.client_id,
        installed.client_secret,
        installed.redirect_uris[0],
    );

    oAuth2Client.on("tokens", (tokens) => {
        const old = fs.existsSync(TOKEN_PATH) ?
            JSON.parse(fs.readFileSync(TOKEN_PATH)) :
            {};

        fs.writeFileSync(
            TOKEN_PATH,
            JSON.stringify({...old, ...tokens }, null, 2),
        );
    });

    if (fs.existsSync(TOKEN_PATH)) {
        oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
        return oAuth2Client;
    }

    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: "offline",
        prompt: "consent",
        scope: SCOPES,
    });

    console.log("\n🔗 AUTHORIZE THIS APP:\n", authUrl);

    const readline = require("readline").createInterface({
        input: process.stdin,
        output: process.stdout,
    });

    const code = await new Promise((resolve) =>
        readline.question("Enter code: ", (c) => {
            readline.close();
            resolve(c);
        }),
    );

    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));

    console.log("✅ Google Drive authorized successfully.");
    return oAuth2Client;
}

async function uploadToDrive(filePath) {
    const auth = await authorize();
    const drive = google.drive({ version: "v3", auth });

    const res = await drive.files.create({
        resource: {
            name: path.basename(filePath),
            parents: [BACKUP_FOLDER_ID],
        },
        media: {
            mimeType: "application/gzip",
            body: fs.createReadStream(filePath),
        },
        fields: "id, webViewLink",
    });

    return res.data.webViewLink;
}
async function cleanupOldDriveBackups() {
    const auth = await authorize();
    const drive = google.drive({ version: "v3", auth });

    const MAX_AGE_HOURS = Number(process.env.BACKUP_RETENTION_HOURS || 48);
    const now = new Date();

    const res = await drive.files.list({
        q: `'${BACKUP_FOLDER_ID}' in parents and name contains 'stockAssistant-backup-'`,
        fields: "files(id, name, createdTime)",
    });

    const files = res.data.files.sort(
        (a, b) => new Date(b.createdTime) - new Date(a.createdTime),
    );

    if (files.length <= 1) return;

    const [, ...olderFiles] = files;

    for (const file of olderFiles) {
        const ageHours = (now - new Date(file.createdTime)) / (1000 * 60 * 60);

        if (ageHours > MAX_AGE_HOURS) {
            await drive.files.delete({ fileId: file.id });
            console.log(`🗑️ Deleted old Drive backup: ${file.name}`);
        }
    }
}

module.exports = { uploadToDrive, cleanupOldDriveBackups };