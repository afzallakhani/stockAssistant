const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
require("dotenv").config();

const SERVICE_ACCOUNT_PATH = path.join(__dirname, "../service-account.json");
const BACKUP_FOLDER_ID = process.env.GOOGLE_DRIVE_BACKUP_FOLDER_ID;

function getDriveClient() {
    const auth = new google.auth.GoogleAuth({
        keyFile: SERVICE_ACCOUNT_PATH,
        scopes: ["https://www.googleapis.com/auth/drive"],
    });

    return google.drive({
        version: "v3",
        auth,
    });
}

async function uploadToDrive(filePath) {
    const drive = getDriveClient();

    const res = await drive.files.create({
        resource: {
            name: path.basename(filePath),
            parents: [BACKUP_FOLDER_ID],
        },
        media: {
            mimeType: "application/gzip",
            body: fs.createReadStream(filePath),
        },
        fields: "id, webViewLink, createdTime",
    });

    console.log(`☁️ Uploaded to Drive: ${res.data.webViewLink}`);
    return res.data.webViewLink;
}

async function cleanupOldDriveBackups() {
    const drive = getDriveClient();
    const MAX_AGE_HOURS = Number(process.env.BACKUP_RETENTION_HOURS || 36);
    const now = new Date();

    const res = await drive.files.list({
        q: `'${BACKUP_FOLDER_ID}' in parents and name contains 'stockAssistant-backup-'`,
        fields: "files(id, name, createdTime)",
        spaces: "drive",
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

module.exports = {
    uploadToDrive,
    cleanupOldDriveBackups,
};