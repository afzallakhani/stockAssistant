// // // utils/backupHelper.js
// const { exec } = require("child_process");
// const path = require("path");
// const fs = require("fs");

// // const DB_NAME = "your_database_name"; // change this
// // const BACKUP_PATH = path.join(__dirname, "../backups");

// // function runBackup(trigger = "manual") {
// //   const DATE = new Date().toISOString().replace(/[:.]/g, "-");
// //   const OUT_PATH = path.join(BACKUP_PATH, `${DB_NAME}-backup-${DATE}`);

// //   if (!fs.existsSync(BACKUP_PATH))
// //     fs.mkdirSync(BACKUP_PATH, { recursive: true });

// //   const cmd = `mongodump --db=${DB_NAME} --out="${OUT_PATH}" --gzip`;

// //   console.log(`🟡 Backup started (${trigger})...`);
// //   exec(cmd, (error, stdout, stderr) => {
// //     if (error) return console.error(`❌ Backup failed: ${error.message}`);
// //     console.log(
// //       `✅ Backup done (${trigger}) at ${new Date().toLocaleString()}`
// //     );
// //   });
// // }

// // module.exports = runBackup;
// let lastBackup = 0;
// let timeoutId;

// function runBackup(trigger = "manual") {
//   const now = Date.now();
//   const MIN_GAP = 60 * 1000; // 1 minute

//   clearTimeout(timeoutId);

//   if (now - lastBackup < MIN_GAP) {
//     // Wait 1 minute after the last change before backing up
//     timeoutId = setTimeout(() => doBackup(trigger), MIN_GAP);
//   } else {
//     doBackup(trigger);
//   }
// }

// function doBackup(trigger) {
//   lastBackup = Date.now();
//   const { exec } = require("child_process");
//   const path = require("path");
//   const fs = require("fs");

//   const DB_NAME = "stockAssistant";
//   const BACKUP_PATH = path.join(__dirname, "../backups");
//   const DATE = new Date().toISOString().replace(/[:.]/g, "-");
//   const OUT_PATH = path.join(BACKUP_PATH, `${DB_NAME}-backup-${DATE}`);

//   if (!fs.existsSync(BACKUP_PATH))
//     fs.mkdirSync(BACKUP_PATH, { recursive: true });

//   const MONGO_TOOLS = `"C:\\Program Files\\MongoDB\\Tools\\100\\bin\\mongodump.exe"`;
//   const cmd = `${MONGO_TOOLS} --db=${DB_NAME} --out="${OUT_PATH}" --gzip`;

//   console.log(`🟡 Backup started (${trigger})...`);
//   exec(cmd, (error) => {
//     if (error) console.error(`❌ Backup failed: ${error.message}`);
//     else console.log(`✅ Backup completed (${trigger})`);
//   });
// }

// module.exports = runBackup;
const { exec } = require("child_process");
const path = require("path");
const fs = require("fs");
const archiver = require("archiver");
const uploadToDrive = require("./googleDrive");

let lastBackup = 0;
let timeoutId;

function runBackup(trigger = "manual") {
    const now = Date.now();
    const MIN_GAP = 60 * 1000; // 1 min throttle for auto
    clearTimeout(timeoutId);
    if (now - lastBackup < MIN_GAP)
        timeoutId = setTimeout(() => doBackup(trigger), MIN_GAP);
    else doBackup(trigger);
}

async function doBackup(trigger) {
    lastBackup = Date.now();
    const DB_NAME = "stockAssistant";
    const BACKUP_PATH = path.join(__dirname, "../backups");
    const DATE = new Date().toISOString().replace(/[:.]/g, "-");
    const OUT_DIR = path.join(BACKUP_PATH, `${DB_NAME}-backup-${DATE}`);
    const ZIP_PATH = `${OUT_DIR}.zip`;

    if (!fs.existsSync(BACKUP_PATH))
        fs.mkdirSync(BACKUP_PATH, { recursive: true });

    const MONGO_TOOLS = `"C:\\Program Files\\MongoDB\\Tools\\100\\bin\\mongodump.exe"`;
    const cmd = `${MONGO_TOOLS} --db=${DB_NAME} --out="${OUT_DIR}" --gzip`;

    console.log(`🟡 Backup started (${trigger})...`);
    exec(cmd, async(error) => {
        if (error) return console.error(`❌ Backup failed: ${error.message}`);

        console.log("📦 Compressing...");
        await zipFolder(OUT_DIR, ZIP_PATH);

        console.log("☁️ Uploading to Google Drive...");
        try {
            const link = await uploadToDrive(ZIP_PATH);
            console.log(`✅ Backup uploaded (${trigger}) → ${link}`);
        } catch (err) {
            console.error("❌ Drive upload failed:", err.message);
        }
    });
}

function zipFolder(src, dest) {
    return new Promise((res, rej) => {
        const out = fs.createWriteStream(dest);
        const archive = archiver("zip", { zlib: { level: 9 } });
        out.on("close", res);
        archive.on("error", rej);
        archive.pipe(out);
        archive.directory(src, false);
        archive.finalize();
    });
}

module.exports = runBackup;

// mongorestore --db=stockAssistant --gzip "C:\StockAssistant\backups\stockAssistant-backup-2025-11-02T15-22-55-765Z\stockAssistant" --drop