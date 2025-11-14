// // utils/googleDrive.js
// const fs = require("fs");
// const path = require("path");
// const { google } = require("googleapis");

// const SCOPES = ["https://www.googleapis.com/auth/drive.file"];
// const TOKEN_PATH = path.join(__dirname, "../token.json");

// async function authorize() {
//     const credentials = JSON.parse(
//         fs.readFileSync(path.join(__dirname, "../credentials.json"))
//     );
//     const { client_secret, client_id, redirect_uris } = credentials.installed;
//     const oAuth2Client = new google.auth.OAuth2(
//         client_id,
//         client_secret,
//         redirect_uris[0]
//     );

//     if (fs.existsSync(TOKEN_PATH)) {
//         oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
//         return oAuth2Client;
//     }

//     const authUrl = oAuth2Client.generateAuthUrl({
//         access_type: "offline",
//         scope: SCOPES,
//     });
//     console.log("\n🔗 Authorize this app:\n", authUrl);
//     const readline = require("readline").createInterface({
//         input: process.stdin,
//         output: process.stdout,
//     });
//     const code = await new Promise((r) =>
//         readline.question("Enter code: ", (c) => {
//             readline.close();
//             r(c);
//         })
//     );
//     const { tokens } = await oAuth2Client.getToken(code);
//     oAuth2Client.setCredentials(tokens);
//     fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
//     console.log("✅ Google Drive token saved.");
//     return oAuth2Client;
// }

// async function uploadToDrive(filePath) {
//     const auth = await authorize();
//     const drive = google.drive({ version: "v3", auth });
//     const fileName = path.basename(filePath);
//     const media = {
//         mimeType: "application/zip",
//         body: fs.createReadStream(filePath),
//     };
//     const res = await drive.files.create({
//         resource: { name: fileName },
//         media,
//         fields: "id, webViewLink",
//     });
//     console.log(`☁️  Uploaded to Drive: ${res.data.webViewLink}`);
//     return res.data.webViewLink;
// }

// module.exports = uploadToDrive;
// utils/googleDrive.js
const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
require("dotenv").config();

const SCOPES = ["https://www.googleapis.com/auth/drive.file"];
const TOKEN_PATH = path.join(__dirname, "../token.json");

async function authorize() {
    const credentials = JSON.parse(
        fs.readFileSync(path.join(__dirname, "../credentials.json"))
    );
    const { client_secret, client_id, redirect_uris } = credentials.installed;
    const oAuth2Client = new google.auth.OAuth2(
        client_id,
        client_secret,
        redirect_uris[0]
    );

    if (fs.existsSync(TOKEN_PATH)) {
        oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
        return oAuth2Client;
    }

    const authUrl = oAuth2Client.generateAuthUrl({
        access_type: "offline",
        scope: SCOPES,
    });
    console.log("\n🔗 Authorize this app:\n", authUrl);
    const readline = require("readline").createInterface({
        input: process.stdin,
        output: process.stdout,
    });
    const code = await new Promise((r) =>
        readline.question("Enter code: ", (c) => {
            readline.close();
            r(c);
        })
    );
    const { tokens } = await oAuth2Client.getToken(code);
    oAuth2Client.setCredentials(tokens);
    fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
    console.log("✅ Google Drive token saved.");
    return oAuth2Client;
}

async function uploadToDrive(filePath) {
    const auth = await authorize();
    const drive = google.drive({ version: "v3", auth });
    const fileName = path.basename(filePath);
    const media = {
        mimeType: "application/zip",
        body: fs.createReadStream(filePath),
    };
    const res = await drive.files.create({
        resource: { name: fileName },
        media,
        fields: "id, webViewLink, createdTime",
    });
    console.log(`☁️  Uploaded to Drive: ${res.data.webViewLink}`);
    return res.data.webViewLink;
}

// 🧹 Delete Drive backups older than retention period (keep latest)
async function cleanupOldDriveBackups() {
    const auth = await authorize();
    const drive = google.drive({ version: "v3", auth });
    const MAX_AGE_DAYS = Number(process.env.BACKUP_RETENTION_DAYS || 3);
    const now = new Date();

    try {
        const res = await drive.files.list({
            q: "name contains 'stockAssistant-backup-' and mimeType='application/zip'",
            fields: "files(id, name, createdTime)",
            spaces: "drive",
        });

        const files = res.data.files.sort(
            (a, b) => new Date(b.createdTime) - new Date(a.createdTime)
        );
        if (files.length <= 1)
            return console.log("✅ Only one Drive backup — skipping deletion.");

        const [latest, ...olderFiles] = files;
        const deletable = olderFiles.filter((file) => {
            const ageDays =
                (now - new Date(file.createdTime)) / (1000 * 60 * 60 * 24);
            return ageDays > MAX_AGE_DAYS;
        });

        if (deletable.length === 0) {
            console.log("✅ No old Drive backups to delete.");
            return;
        }

        for (const file of deletable) {
            await drive.files.delete({ fileId: file.id });
            console.log(`🗑️ Deleted old Drive backup: ${file.name}`);
        }
    } catch (err) {
        console.error("⚠️ Error deleting old Drive backups:", err.message);
    }
}

module.exports = { uploadToDrive, cleanupOldDriveBackups };