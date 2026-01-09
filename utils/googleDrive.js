// // // // utils/googleDrive.js
// // // const fs = require("fs");
// // // const path = require("path");
// // // const { google } = require("googleapis");

// // // const SCOPES = ["https://www.googleapis.com/auth/drive.file"];
// // // const TOKEN_PATH = path.join(__dirname, "../token.json");

// // // async function authorize() {
// // //     const credentials = JSON.parse(
// // //         fs.readFileSync(path.join(__dirname, "../credentials.json"))
// // //     );
// // //     const { client_secret, client_id, redirect_uris } = credentials.installed;
// // //     const oAuth2Client = new google.auth.OAuth2(
// // //         client_id,
// // //         client_secret,
// // //         redirect_uris[0]
// // //     );

// // //     if (fs.existsSync(TOKEN_PATH)) {
// // //         oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
// // //         return oAuth2Client;
// // //     }

// // //     const authUrl = oAuth2Client.generateAuthUrl({
// // //         access_type: "offline",
// // //         scope: SCOPES,
// // //     });
// // //     console.log("\n🔗 Authorize this app:\n", authUrl);
// // //     const readline = require("readline").createInterface({
// // //         input: process.stdin,
// // //         output: process.stdout,
// // //     });
// // //     const code = await new Promise((r) =>
// // //         readline.question("Enter code: ", (c) => {
// // //             readline.close();
// // //             r(c);
// // //         })
// // //     );
// // //     const { tokens } = await oAuth2Client.getToken(code);
// // //     oAuth2Client.setCredentials(tokens);
// // //     fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
// // //     console.log("✅ Google Drive token saved.");
// // //     return oAuth2Client;
// // // }

// // // async function uploadToDrive(filePath) {
// // //     const auth = await authorize();
// // //     const drive = google.drive({ version: "v3", auth });
// // //     const fileName = path.basename(filePath);
// // //     const media = {
// // //         mimeType: "application/zip",
// // //         body: fs.createReadStream(filePath),
// // //     };
// // //     const res = await drive.files.create({
// // //         resource: { name: fileName },
// // //         media,
// // //         fields: "id, webViewLink",
// // //     });
// // //     console.log(`☁️  Uploaded to Drive: ${res.data.webViewLink}`);
// // //     return res.data.webViewLink;
// // // }

// // // module.exports = uploadToDrive;
// // // utils/googleDrive.js
// // const fs = require("fs");
// // const path = require("path");
// // const { google } = require("googleapis");
// // require("dotenv").config();

// // const SCOPES = ["https://www.googleapis.com/auth/drive.file"];
// // const TOKEN_PATH = path.join(__dirname, "../token.json");

// // async function authorize() {
// //     const credentials = JSON.parse(
// //         fs.readFileSync(path.join(__dirname, "../credentials.json"))
// //     );
// //     const { client_secret, client_id, redirect_uris } = credentials.installed;
// //     const oAuth2Client = new google.auth.OAuth2(
// //         client_id,
// //         client_secret,
// //         redirect_uris[0]
// //     );

// //     if (fs.existsSync(TOKEN_PATH)) {
// //         oAuth2Client.setCredentials(JSON.parse(fs.readFileSync(TOKEN_PATH)));
// //         return oAuth2Client;
// //     }

// //     const authUrl = oAuth2Client.generateAuthUrl({
// //         access_type: "offline",
// //         scope: SCOPES,
// //     });
// //     console.log("\n🔗 Authorize this app:\n", authUrl);
// //     const readline = require("readline").createInterface({
// //         input: process.stdin,
// //         output: process.stdout,
// //     });
// //     const code = await new Promise((r) =>
// //         readline.question("Enter code: ", (c) => {
// //             readline.close();
// //             r(c);
// //         })
// //     );
// //     const { tokens } = await oAuth2Client.getToken(code);
// //     oAuth2Client.setCredentials(tokens);
// //     fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens));
// //     console.log("✅ Google Drive token saved.");
// //     return oAuth2Client;
// // }

// // async function uploadToDrive(filePath) {
// //     const auth = await authorize();
// //     const drive = google.drive({ version: "v3", auth });
// //     const fileName = path.basename(filePath);
// //     const media = {
// //         mimeType: "application/zip",
// //         body: fs.createReadStream(filePath),
// //     };
// //     const res = await drive.files.create({
// //         resource: { name: fileName },
// //         media,
// //         fields: "id, webViewLink, createdTime",
// //     });
// //     console.log(`☁️  Uploaded to Drive: ${res.data.webViewLink}`);
// //     return res.data.webViewLink;
// // }

// // // 🧹 Delete Drive backups older than retention period (keep latest)
// // async function cleanupOldDriveBackups() {
// //     const auth = await authorize();
// //     const drive = google.drive({ version: "v3", auth });
// //     const MAX_AGE_DAYS = Number(process.env.BACKUP_RETENTION_DAYS || 3);
// //     const now = new Date();

// //     try {
// //         const res = await drive.files.list({
// //             q: "name contains 'stockAssistant-backup-' and mimeType='application/zip'",
// //             fields: "files(id, name, createdTime)",
// //             spaces: "drive",
// //         });

// //         const files = res.data.files.sort(
// //             (a, b) => new Date(b.createdTime) - new Date(a.createdTime)
// //         );
// //         if (files.length <= 1)
// //             return console.log("✅ Only one Drive backup — skipping deletion.");

// //         const [latest, ...olderFiles] = files;
// //         const deletable = olderFiles.filter((file) => {
// //             const ageDays =
// //                 (now - new Date(file.createdTime)) / (1000 * 60 * 60 * 24);
// //             return ageDays > MAX_AGE_DAYS;
// //         });

// //         if (deletable.length === 0) {
// //             console.log("✅ No old Drive backups to delete.");
// //             return;
// //         }

// //         for (const file of deletable) {
// //             await drive.files.delete({ fileId: file.id });
// //             console.log(`🗑️ Deleted old Drive backup: ${file.name}`);
// //         }
// //     } catch (err) {
// //         console.error("⚠️ Error deleting old Drive backups:", err.message);
// //     }
// // }

// // module.exports = { uploadToDrive, cleanupOldDriveBackups };

// // utils/googleDrive.js
// const fs = require("fs");
// const path = require("path");
// const { google } = require("googleapis");
// require("dotenv").config();

// const SCOPES = ["https://www.googleapis.com/auth/drive.file"];
// const TOKEN_PATH = path.join(__dirname, "../token.json");
// const CREDENTIALS_PATH = path.join(__dirname, "../credentials.json");

// function loadCredentials() {
//     return JSON.parse(fs.readFileSync(CREDENTIALS_PATH, "utf8"));
// }

// async function authorize() {
//     const { installed } = loadCredentials();

//     const oAuth2Client = new google.auth.OAuth2(
//         installed.client_id,
//         installed.client_secret,
//         installed.redirect_uris[0]
//     );

//     // 🔄 Auto-persist refreshed tokens
//     oAuth2Client.on("tokens", (tokens) => {
//         if (!tokens.access_token && !tokens.refresh_token) return;

//         const old = fs.existsSync(TOKEN_PATH) ?
//             JSON.parse(fs.readFileSync(TOKEN_PATH, "utf8")) : {};

//         fs.writeFileSync(
//             TOKEN_PATH,
//             JSON.stringify({...old, ...tokens }, null, 2)
//         );
//     });

//     // Load existing token
//     if (fs.existsSync(TOKEN_PATH)) {
//         oAuth2Client.setCredentials(
//             JSON.parse(fs.readFileSync(TOKEN_PATH, "utf8"))
//         );
//         return oAuth2Client;
//     }

//     // 🔐 First-time authorization (one time only)
//     const authUrl = oAuth2Client.generateAuthUrl({
//         access_type: "offline",
//         prompt: "consent", // 🔑 forces refresh_token
//         scope: SCOPES,
//     });

//     console.log("\n🔗 AUTHORIZE GOOGLE DRIVE ACCESS:\n", authUrl);

//     const readline = require("readline").createInterface({
//         input: process.stdin,
//         output: process.stdout,
//     });

//     const code = await new Promise((resolve) =>
//         readline.question("\nEnter authorization code: ", (c) => {
//             readline.close();
//             resolve(c);
//         })
//     );

//     const { tokens } = await oAuth2Client.getToken(code);
//     oAuth2Client.setCredentials(tokens);
//     fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));

//     console.log("✅ Google Drive authorized successfully.");
//     return oAuth2Client;
// }

// async function uploadToDrive(filePath) {
//     const auth = await authorize();
//     const drive = google.drive({ version: "v3", auth });

//     await fs.promises.access(filePath); // ensure file exists

//     const stream = fs.createReadStream(filePath);

//     stream.on("error", (err) => {
//         console.error("📛 ReadStream error:", err.message);
//     });

//     //   const res = await drive.files.create({
//     //     resource: { name: path.basename(filePath) },
//     //     media: {
//     //       mimeType: "application/gzip", // ✅ correct for .archive.gz
//     //       body: stream,
//     //     },
//     //     fields: "id, webViewLink",
//     //   });

//     //   // ✅ DO NOT wait for stream.close (Drive already consumed it)
//     //   return res.data.webViewLink;
//     // }
//     const res = await drive.files.create({
//         resource: {
//             name: path.basename(filePath),
//             parents: [process.env.GOOGLE_DRIVE_BACKUP_FOLDER_ID], // 📁 TARGET FOLDER
//         },
//         media: {
//             mimeType: "application/gzip",
//             body: stream,
//         },
//         fields: "id, webViewLink",
//     });

//     // 🧹 Cleanup old Drive backups (SAFE with auto-refresh)
//     async function cleanupOldDriveBackups() {
//         const auth = await authorize();
//         const drive = google.drive({ version: "v3", auth });

//         const MAX_AGE_DAYS = Number(process.env.BACKUP_RETENTION_DAYS || 3);
//         const now = new Date();

//         // const res = await drive.files.list({
//         //     q: "name contains 'stockAssistant-backup-' and mimeType='application/zip'",
//         //     fields: "files(id, name, createdTime)",
//         //     spaces: "drive",
//         // });
//         const res = await drive.files.list({
//             q: `'${process.env.GOOGLE_DRIVE_BACKUP_FOLDER_ID}' in parents and name contains 'stockAssistant-backup-'`,
//             fields: "files(id, name, createdTime)",
//             spaces: "drive",
//         });

//         const files = res.data.files.sort(
//             (a, b) => new Date(b.createdTime) - new Date(a.createdTime)
//         );

//         if (files.length <= 1) return;

//         const [, ...olderFiles] = files;

//         for (const file of olderFiles) {
//             const ageDays = (now - new Date(file.createdTime)) / (1000 * 60 * 60 * 24);

//             if (ageDays > MAX_AGE_DAYS) {
//                 await drive.files.delete({ fileId: file.id });
//                 console.log(`🗑️ Deleted old Drive backup: ${file.name}`);
//             }
//         }
//     }

//     module.exports = {
//         uploadToDrive,
//         cleanupOldDriveBackups,
//     };

const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");
require("dotenv").config();

const SCOPES = ["https://www.googleapis.com/auth/drive.file"];
const TOKEN_PATH = path.join(__dirname, "../token.json");
const CREDENTIALS_PATH = path.join(__dirname, "../credentials.json");
const BACKUP_FOLDER_ID = process.env.GOOGLE_DRIVE_BACKUP_FOLDER_ID;

// ---------------- AUTH ----------------

function loadCredentials() {
  return JSON.parse(fs.readFileSync(CREDENTIALS_PATH, "utf8"));
}

async function authorize() {
  const { installed } = loadCredentials();

  const oAuth2Client = new google.auth.OAuth2(
    installed.client_id,
    installed.client_secret,
    installed.redirect_uris[0]
  );

  // 🔄 Auto-save refreshed tokens
  oAuth2Client.on("tokens", (tokens) => {
    if (!tokens.access_token && !tokens.refresh_token) return;

    const old = fs.existsSync(TOKEN_PATH)
      ? JSON.parse(fs.readFileSync(TOKEN_PATH, "utf8"))
      : {};

    fs.writeFileSync(
      TOKEN_PATH,
      JSON.stringify({ ...old, ...tokens }, null, 2)
    );
  });

  if (fs.existsSync(TOKEN_PATH)) {
    oAuth2Client.setCredentials(
      JSON.parse(fs.readFileSync(TOKEN_PATH, "utf8"))
    );
    return oAuth2Client;
  }

  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: "offline",
    prompt: "consent",
    scope: SCOPES,
  });

  console.log("\n🔗 AUTHORIZE GOOGLE DRIVE ACCESS:\n", authUrl);

  const readline = require("readline").createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  const code = await new Promise((resolve) =>
    readline.question("\nEnter authorization code: ", (c) => {
      readline.close();
      resolve(c);
    })
  );

  const { tokens } = await oAuth2Client.getToken(code);
  oAuth2Client.setCredentials(tokens);
  fs.writeFileSync(TOKEN_PATH, JSON.stringify(tokens, null, 2));

  console.log("✅ Google Drive authorized.");
  return oAuth2Client;
}

// ---------------- UPLOAD ----------------

async function uploadToDrive(filePath) {
  const auth = await authorize();
  const drive = google.drive({ version: "v3", auth });

  await fs.promises.access(filePath);

  const stream = fs.createReadStream(filePath);
  stream.on("error", (err) =>
    console.error("📛 ReadStream error:", err.message)
  );

  const res = await drive.files.create({
    resource: {
      name: path.basename(filePath),
      parents: [BACKUP_FOLDER_ID],
    },
    media: {
      mimeType: "application/gzip",
      body: stream,
    },
    fields: "id, webViewLink, createdTime",
  });

  console.log(`☁️ Uploaded to Drive: ${res.data.webViewLink}`);
  return res.data.webViewLink;
}

// ---------------- CLEANUP ----------------

async function cleanupOldDriveBackups() {
  const auth = await authorize();
  const drive = google.drive({ version: "v3", auth });

  // const MAX_AGE_DAYS = Number(process.env.BACKUP_RETENTION_DAYS || 3);
  const MAX_AGE_HOURS = Number(process.env.BACKUP_RETENTION_HOURS || 48);

  const now = new Date();

  const res = await drive.files.list({
    q: `'${BACKUP_FOLDER_ID}' in parents and name contains 'stockAssistant-backup-'`,
    fields: "files(id, name, createdTime)",
    spaces: "drive",
  });

  const files = res.data.files.sort(
    (a, b) => new Date(b.createdTime) - new Date(a.createdTime)
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

// ---------------- EXPORTS ----------------

module.exports = {
  uploadToDrive,
  cleanupOldDriveBackups,
};
