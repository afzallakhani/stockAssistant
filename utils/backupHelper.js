// // const { exec } = require("child_process");
// // const path = require("path");
// // const fs = require("fs");
// // const archiver = require("archiver");
// // const { uploadToDrive, cleanupOldDriveBackups } = require("./googleDrive");
// // require("dotenv").config();

// // let lastBackup = 0;
// // let timeoutId;

// // function runBackup(trigger = "manual") {
// //     const now = Date.now();
// //     const MIN_GAP = 60 * 1000; // 1-minute throttle
// //     clearTimeout(timeoutId);
// //     if (now - lastBackup < MIN_GAP)
// //         timeoutId = setTimeout(() => doBackup(trigger), MIN_GAP);
// //     else doBackup(trigger);
// // }

// // async function doBackup(trigger) {
// //     lastBackup = Date.now();
// //     const DB_NAME = "stockAssistant";
// //     const BACKUP_PATH = path.join(__dirname, "../backups");
// //     const DATE = new Date().toISOString().replace(/[:.]/g, "-");
// //     const OUT_DIR = path.join(BACKUP_PATH, `${DB_NAME}-backup-${DATE}`);
// //     const ZIP_PATH = `${OUT_DIR}.zip`;

// //     if (!fs.existsSync(BACKUP_PATH))
// //         fs.mkdirSync(BACKUP_PATH, { recursive: true });

// //     const MONGO_TOOLS = `"C:\\Program Files\\MongoDB\\Tools\\100\\bin\\mongodump.exe"`;
// //     const cmd = `${MONGO_TOOLS} --db=${DB_NAME} --out="${OUT_DIR}" --gzip`;

// //     console.log(`🟡 Backup started (${trigger})...`);
// //     exec(cmd, async(error) => {
// //         if (error) return console.error(`❌ Backup failed: ${error.message}`);

// //         console.log("📦 Compressing...");
// //         await zipFolder(OUT_DIR, ZIP_PATH);

// //         console.log("☁️ Uploading to Google Drive...");
// //         try {
// //             const link = await uploadToDrive(ZIP_PATH);
// //             console.log(`✅ Backup uploaded (${trigger}) → ${link}`);

// //             deleteOldLocalBackups(); // 🧹 Local cleanup
// //             cleanupOldDriveBackups(); // ☁️ Drive cleanup
// //         } catch (err) {
// //             console.error("❌ Drive upload failed:", err.message);
// //         }
// //     });
// // }

// // function zipFolder(src, dest) {
// //     return new Promise((res, rej) => {
// //         const out = fs.createWriteStream(dest);
// //         const archive = archiver("zip", { zlib: { level: 9 } });
// //         out.on("close", res);
// //         archive.on("error", rej);
// //         archive.pipe(out);
// //         archive.directory(src, false);
// //         archive.finalize();
// //     });
// // }

// // // 🧹 Delete old local backups but always keep the newest
// // function deleteOldLocalBackups() {
// //     const BACKUP_PATH = path.join(__dirname, "../backups");
// //     const MAX_AGE_DAYS = Number(process.env.BACKUP_RETENTION_DAYS || 3);
// //     const now = Date.now();

// //     if (!fs.existsSync(BACKUP_PATH)) return;

// //     const files = fs
// //         .readdirSync(BACKUP_PATH)
// //         .map((f) => ({
// //             name: f,
// //             time: fs.statSync(path.join(BACKUP_PATH, f)).mtimeMs,
// //         }))
// //         .sort((a, b) => b.time - a.time);

// //     if (files.length <= 1)
// //         return console.log("✅ Only one local backup — skipping deletion.");

// //     const [latest, ...olderFiles] = files;

// //     for (const file of olderFiles) {
// //         const filePath = path.join(BACKUP_PATH, file.name);
// //         const ageDays = (now - file.time) / (1000 * 60 * 60 * 24);
// //         if (ageDays > MAX_AGE_DAYS) {
// //             try {
// //                 if (fs.lstatSync(filePath).isDirectory()) {
// //                     fs.rmSync(filePath, { recursive: true, force: true });
// //                 } else {
// //                     fs.unlinkSync(filePath);
// //                 }
// //                 console.log(`🗑️ Deleted old local backup: ${file.name}`);
// //             } catch (err) {
// //                 console.error("⚠️ Error deleting old local backup:", err.message);
// //             }
// //         }
// //     }
// // }

// // module.exports = runBackup;
// // const { exec } = require("child_process");
// // const path = require("path");
// // const fs = require("fs");
// // const archiver = require("archiver");
// // const { uploadToDrive, cleanupOldDriveBackups } = require("./googleDrive");
// // require("dotenv").config();

// // let lastBackup = 0;
// // let timeoutId;

// // function runBackup(trigger = "manual") {
// //     const now = Date.now();
// //     const MIN_GAP = 60 * 1000; // 1 minute throttle for auto
// //     clearTimeout(timeoutId);
// //     if (now - lastBackup < MIN_GAP)
// //         timeoutId = setTimeout(() => doBackup(trigger), MIN_GAP);
// //     else doBackup(trigger);
// // }

// // async function doBackup(trigger) {
// //     lastBackup = Date.now();
// //     const DB_NAME = "stockAssistant";
// //     const BACKUP_PATH = path.join(__dirname, "../backups");
// //     const PENDING_PATH = path.join(__dirname, "../pending-uploads");
// //     const DATE = new Date().toISOString().replace(/[:.]/g, "-");
// //     const OUT_DIR = path.join(BACKUP_PATH, `${DB_NAME}-backup-${DATE}`);
// //     const ZIP_PATH = `${OUT_DIR}.zip`;

// //     // Ensure directories exist
// //     [BACKUP_PATH, PENDING_PATH].forEach((dir) => {
// //         if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
// //     });

// //     const MONGO_TOOLS = `"C:\\Program Files\\MongoDB\\Tools\\100\\bin\\mongodump.exe"`;
// //     const cmd = `${MONGO_TOOLS} --db=${DB_NAME} --out="${OUT_DIR}" --gzip`;

// //     console.log(`🟡 Backup started (${trigger})...`);
// //     exec(cmd, async(error) => {
// //         if (error) return console.error(`❌ Backup failed: ${error.message}`);

// //         console.log("📦 Compressing...");
// //         await zipFolder(OUT_DIR, ZIP_PATH);

// //         // Retry any pending uploads first
// //         await retryPendingUploads(PENDING_PATH);

// //         console.log("☁️ Uploading to Google Drive...");
// //         try {
// //             const link = await uploadToDrive(ZIP_PATH);
// //             console.log(`✅ Backup uploaded (${trigger}) → ${link}`);

// //             deleteOldLocalBackups();
// //             cleanupOldDriveBackups();
// //         } catch (err) {
// //             console.error("❌ Drive upload failed:", err.message);
// //             // Move file to pending folder for later retry
// //             const pendingFile = path.join(PENDING_PATH, path.basename(ZIP_PATH));
// //             fs.copyFileSync(ZIP_PATH, pendingFile);
// //             console.log(`📦 Backup saved for retry: ${pendingFile}`);
// //         }
// //     });
// // }

// // function zipFolder(src, dest) {
// //     return new Promise((res, rej) => {
// //         const out = fs.createWriteStream(dest);
// //         const archive = archiver("zip", { zlib: { level: 9 } });
// //         out.on("close", res);
// //         archive.on("error", rej);
// //         archive.pipe(out);
// //         archive.directory(src, false);
// //         archive.finalize();
// //     });
// // }

// // // 🔁 Retry any pending uploads before new backups (with safe deletion)
// // async function retryPendingUploads(PENDING_PATH) {
// //     const pendingFiles = fs.existsSync(PENDING_PATH) ?
// //         fs.readdirSync(PENDING_PATH).filter((f) => f.endsWith(".zip")) :
// //         [];

// //     if (pendingFiles.length === 0) return;

// //     console.log(`🔄 Retrying ${pendingFiles.length} pending upload(s)...`);
// //     for (const file of pendingFiles) {
// //         const filePath = path.join(PENDING_PATH, file);
// //         try {
// //             const link = await uploadToDrive(filePath);
// //             console.log(`✅ Retried and uploaded: ${file} → ${link}`);

// //             // Wait a bit before deleting to let file handles close
// //             await new Promise((res) => setTimeout(res, 3000));

// //             try {
// //                 fs.unlinkSync(filePath);
// //                 console.log(`🗑️ Deleted pending file: ${file}`);
// //             } catch (err) {
// //                 console.warn(`⚠️ Could not delete ${file}: ${err.message}`);
// //             }
// //         } catch (err) {
// //             console.log(`⚠️ Retry failed for ${file}: ${err.message}`);
// //         }
// //     }
// // }

// // // 🧹 Delete old local backups but always keep latest
// // function deleteOldLocalBackups() {
// //     const BACKUP_PATH = path.join(__dirname, "../backups");
// //     const MAX_AGE_DAYS = Number(process.env.BACKUP_RETENTION_DAYS || 3);
// //     const now = Date.now();

// //     if (!fs.existsSync(BACKUP_PATH)) return;

// //     const files = fs
// //         .readdirSync(BACKUP_PATH)
// //         .map((f) => ({
// //             name: f,
// //             time: fs.statSync(path.join(BACKUP_PATH, f)).mtimeMs,
// //         }))
// //         .sort((a, b) => b.time - a.time);

// //     if (files.length <= 1) {
// //         console.log("✅ Only one local backup — skipping deletion.");
// //         return;
// //     }

// //     const [latest, ...olderFiles] = files;

// //     for (const file of olderFiles) {
// //         const filePath = path.join(BACKUP_PATH, file.name);
// //         const ageDays = (now - file.time) / (1000 * 60 * 60 * 24);
// //         if (ageDays > MAX_AGE_DAYS) {
// //             try {
// //                 if (fs.lstatSync(filePath).isDirectory()) {
// //                     fs.rmSync(filePath, { recursive: true, force: true });
// //                 } else {
// //                     fs.unlinkSync(filePath);
// //                 }
// //                 console.log(`🗑️ Deleted old local backup: ${file.name}`);
// //             } catch (err) {
// //                 console.error("⚠️ Error deleting old local backup:", err.message);
// //             }
// //         }
// //     }
// // }

// // module.exports = runBackup;
// const { exec } = require("child_process");
// const path = require("path");
// const fs = require("fs");
// const https = require("https");
// const archiver = require("archiver");
// const { uploadToDrive, cleanupOldDriveBackups } = require("./googleDrive");
// require("dotenv").config();

// let lastBackup = 0;
// let timeoutId;

// // 🧾 Logging Setup
// const LOG_DIR = path.join(__dirname, "../logs");
// if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR, { recursive: true });
// const LOG_PATH = path.join(
//     LOG_DIR,
//     `backup-${new Date().toISOString().split("T")[0]}.log`
// );
// const LOG_RETENTION_DAYS = Number(process.env.LOG_RETENTION_DAYS || 7);

// function logMessage(message) {
//     const timestamp = new Date().toLocaleString();
//     const line = `[${timestamp}] ${message}\n`;
//     console.log(message);
//     fs.appendFileSync(LOG_PATH, line, "utf8");
// }
// // 🧹 Delete logs older than LOG_RETENTION_DAYS (keep daily log files)
// function cleanupOldLogs() {
//     try {
//         const now = Date.now();
//         const files = fs
//             .readdirSync(LOG_DIR)
//             .filter((f) => f.startsWith("backup-") && f.endsWith(".log"));

//         files.forEach((file) => {
//             const full = path.join(LOG_DIR, file);
//             const stat = fs.statSync(full);
//             const ageDays = (now - stat.mtimeMs) / (1000 * 60 * 60 * 24);
//             if (ageDays > LOG_RETENTION_DAYS) {
//                 try {
//                     fs.unlinkSync(full);
//                     console.log(`🧹 Deleted old log: ${file}`);
//                 } catch (e) {
//                     console.warn(`⚠️ Could not delete old log ${file}: ${e.message}`);
//                 }
//             }
//         });
//     } catch (err) {
//         console.warn(`⚠️ Log cleanup error: ${err.message}`);
//     }
// }

// function runBackup(trigger = "manual") {
//     const now = Date.now();
//     const MIN_GAP = 60 * 1000; // 1 min throttle
//     clearTimeout(timeoutId);
//     if (now - lastBackup < MIN_GAP)
//         timeoutId = setTimeout(() => doBackup(trigger), MIN_GAP);
//     else doBackup(trigger);
// }

// async function doBackup(trigger) {
//     lastBackup = Date.now();
//     const DB_NAME = "stockAssistant";
//     const BACKUP_PATH = path.join(__dirname, "../backups");
//     const PENDING_PATH = path.join(__dirname, "../pending-uploads");
//     const DATE = new Date().toISOString().replace(/[:.]/g, "-");
//     const OUT_DIR = path.join(BACKUP_PATH, `${DB_NAME}-backup-${DATE}`);
//     const ZIP_PATH = `${OUT_DIR}.zip`;

//     [BACKUP_PATH, PENDING_PATH].forEach((dir) => {
//         if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
//     });

//     const MONGO_TOOLS = `"C:\\Program Files\\MongoDB\\Tools\\100\\bin\\mongodump.exe"`;
//     const cmd = `${MONGO_TOOLS} --db=${DB_NAME} --out="${OUT_DIR}" --gzip`;

//     logMessage(`🟡 Backup started (${trigger})...`);
//     exec(cmd, async(error) => {
//         if (error) return logMessage(`❌ Backup failed: ${error.message}`);

//         logMessage("📦 Compressing...");
//         await zipFolder(OUT_DIR, ZIP_PATH);

//         // Retry pending uploads first
//         await retryPendingUploads(PENDING_PATH);

//         logMessage("☁️ Uploading to Google Drive...");
//         try {
//             const link = await uploadToDrive(ZIP_PATH);
//             logMessage(`✅ Backup uploaded (${trigger}) → ${link}`);

//             deleteOldLocalBackups();
//             cleanupOldDriveBackups();
//         } catch (err) {
//             logMessage(`❌ Drive upload failed: ${err.message}`);
//             const pendingFile = path.join(PENDING_PATH, path.basename(ZIP_PATH));
//             fs.copyFileSync(ZIP_PATH, pendingFile);
//             logMessage(`📦 Backup saved for retry: ${pendingFile}`);
//         }
//     });
// }

// // 📦 Zip folder
// function zipFolder(src, dest) {
//     return new Promise((res, rej) => {
//         const out = fs.createWriteStream(dest);
//         const archive = archiver("zip", { zlib: { level: 9 } });
//         out.on("close", res);
//         archive.on("error", rej);
//         archive.pipe(out);
//         archive.directory(src, false);
//         archive.finalize();
//     });
// }

// // 🔁 Retry pending uploads (safe delete + logging)
// async function retryPendingUploads(PENDING_PATH) {
//     const pendingFiles = fs.existsSync(PENDING_PATH) ?
//         fs.readdirSync(PENDING_PATH).filter((f) => f.endsWith(".zip")) :
//         [];

//     if (pendingFiles.length === 0) return;

//     logMessage(`🔄 Retrying ${pendingFiles.length} pending upload(s)...`);
//     for (const file of pendingFiles) {
//         const filePath = path.join(PENDING_PATH, file);
//         try {
//             const link = await uploadToDrive(filePath);
//             logMessage(`✅ Retried and uploaded: ${file} → ${link}`);
//             await new Promise((res) => setTimeout(res, 1500));
//             try {
//                 fs.unlinkSync(filePath);
//                 logMessage(`🗑️ Deleted pending file: ${file}`);
//             } catch (err) {
//                 logMessage(`⚠️ Could not delete ${file}: ${err.message}`);
//             }
//         } catch (err) {
//             logMessage(`⚠️ Retry failed for ${file}: ${err.message}`);
//         }
//     }
// }

// // 🧹 Delete old local backups
// function deleteOldLocalBackups() {
//     const BACKUP_PATH = path.join(__dirname, "../backups");
//     const MAX_AGE_DAYS = Number(process.env.BACKUP_RETENTION_DAYS || 3);
//     const now = Date.now();

//     if (!fs.existsSync(BACKUP_PATH)) return;

//     const files = fs
//         .readdirSync(BACKUP_PATH)
//         .map((f) => ({
//             name: f,
//             time: fs.statSync(path.join(BACKUP_PATH, f)).mtimeMs,
//         }))
//         .sort((a, b) => b.time - a.time);

//     if (files.length <= 1) {
//         logMessage("✅ Only one local backup — skipping deletion.");
//         return;
//     }

//     const [latest, ...olderFiles] = files;

//     for (const file of olderFiles) {
//         const filePath = path.join(BACKUP_PATH, file.name);
//         const ageDays = (now - file.time) / (1000 * 60 * 60 * 24);
//         if (ageDays > MAX_AGE_DAYS) {
//             try {
//                 if (fs.lstatSync(filePath).isDirectory()) {
//                     fs.rmSync(filePath, { recursive: true, force: true });
//                 } else {
//                     fs.unlinkSync(filePath);
//                 }
//                 logMessage(`🗑️ Deleted old local backup: ${file.name}`);
//             } catch (err) {
//                 logMessage(`⚠️ Error deleting old local backup: ${err.message}`);
//             }
//         }
//     }
// }

// // 🕒 Periodic retry every 15 minutes
// function scheduleRetryUploads() {
//     const PENDING_PATH = path.join(__dirname, "../pending-uploads");
//     const RETRY_INTERVAL_MINUTES = 15;
//     setInterval(async() => {
//         const pendingFiles = fs.existsSync(PENDING_PATH) ?
//             fs.readdirSync(PENDING_PATH).filter((f) => f.endsWith(".zip")) :
//             [];
//         if (pendingFiles.length === 0) return;
//         logMessage(
//             `⏰ Scheduled retry: ${pendingFiles.length} pending backup(s)...`
//         );
//         await retryPendingUploads(PENDING_PATH);
//     }, RETRY_INTERVAL_MINUTES * 60 * 1000);
// }
// scheduleRetryUploads();
// // Run once on startup
// cleanupOldLogs();

// // Then run daily at ~03:05 local time
// function scheduleLogCleanupDaily() {
//     const MS_DAY = 24 * 60 * 60 * 1000;
//     const now = new Date();
//     const next = new Date(
//         now.getFullYear(),
//         now.getMonth(),
//         now.getDate(),
//         3,
//         5,
//         0,
//         0 // 03:05:00
//     );
//     if (next <= now) next.setDate(next.getDate() + 1);

//     setTimeout(() => {
//         cleanupOldLogs();
//         setInterval(cleanupOldLogs, MS_DAY);
//     }, next.getTime() - now.getTime());
// }
// scheduleLogCleanupDaily();

// // 🌐 Check internet connectivity
// function checkInternetConnection() {
//     return new Promise((resolve) => {
//         https
//             .get("https://www.google.com", () => resolve(true))
//             .on("error", () => resolve(false))
//             .setTimeout(3000, () => resolve(false));
//     });
// }

// // 🌐 Auto-detect when internet returns → instant retry
// function watchInternetAndRetry() {
//     const PENDING_PATH = path.join(__dirname, "../pending-uploads");
//     let wasOffline = false;

//     setInterval(async() => {
//         const isOnline = await checkInternetConnection();

//         if (isOnline && wasOffline) {
//             const pendingFiles = fs.existsSync(PENDING_PATH) ?
//                 fs.readdirSync(PENDING_PATH).filter((f) => f.endsWith(".zip")) :
//                 [];
//             if (pendingFiles.length > 0) {
//                 logMessage("🌐 Internet reconnected! Retrying pending uploads...");
//                 await retryPendingUploads(PENDING_PATH);
//             } else {
//                 logMessage("🌐 Internet reconnected — no pending uploads.");
//             }
//         }
//         wasOffline = !isOnline;
//     }, 30000); // every 30s
// }
// watchInternetAndRetry();

// module.exports = runBackup;


const { exec } = require("child_process");
const path = require("path");
const fs = require("fs");
const https = require("https");
const archiver = require("archiver");
require("dotenv").config();

let lastBackup = 0;
let timeoutId;

// 🧾 Logging setup
const LOG_DIR = path.join(__dirname, "../logs");
if (!fs.existsSync(LOG_DIR)) fs.mkdirSync(LOG_DIR, { recursive: true });
const LOG_PATH = path.join(
    LOG_DIR,
    `backup-${new Date().toISOString().split("T")[0]}.log`
);
const LOG_RETENTION_DAYS = Number(process.env.LOG_RETENTION_DAYS || 7);

function logMessage(message) {
    const timestamp = new Date().toLocaleString();
    const line = `[${timestamp}] ${message}\n`;
    console.log(message);
    fs.appendFileSync(LOG_PATH, line, "utf8");
}

// 🧹 Delete logs older than X days
function cleanupOldLogs() {
    try {
        const now = Date.now();
        const files = fs
            .readdirSync(LOG_DIR)
            .filter((f) => f.startsWith("backup-") && f.endsWith(".log"));

        files.forEach((file) => {
            const full = path.join(LOG_DIR, file);
            const stat = fs.statSync(full);
            const ageDays = (now - stat.mtimeMs) / (1000 * 60 * 60 * 24);
            if (ageDays > LOG_RETENTION_DAYS) {
                try {
                    fs.unlinkSync(full);
                    console.log(`🧹 Deleted old log: ${file}`);
                } catch (e) {
                    console.warn(`⚠️ Could not delete old log ${file}: ${e.message}`);
                }
            }
        });
    } catch (err) {
        console.warn(`⚠️ Log cleanup error: ${err.message}`);
    }
}

// Throttle auto backups
function runBackup(trigger = "manual") {
    const now = Date.now();
    const MIN_GAP = 60 * 1000; // 1-minute minimum gap
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

    logMessage(`🟡 Backup started (${trigger})...`);

    exec(cmd, async (error) => {
        if (error) return logMessage(`❌ Backup failed: ${error.message}`);

        logMessage("📦 Compressing...");
        await zipFolder(OUT_DIR, ZIP_PATH);

        logMessage(`✅ Local backup completed: ${ZIP_PATH}`);

        deleteOldLocalBackups();
    });
}

// Zip folder
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

// 🧹 Delete old local backups (keep newest)
function deleteOldLocalBackups() {
    const BACKUP_PATH = path.join(__dirname, "../backups");
    const MAX_AGE_DAYS = Number(process.env.BACKUP_RETENTION_DAYS || 3);
    const now = Date.now();

    if (!fs.existsSync(BACKUP_PATH)) return;

    const files = fs
        .readdirSync(BACKUP_PATH)
        .map((f) => ({
            name: f,
            time: fs.statSync(path.join(BACKUP_PATH, f)).mtimeMs,
        }))
        .sort((a, b) => b.time - a.time);

    if (files.length <= 1) {
        logMessage("⏳ Only one backup present — skipping old backup deletion.");
        return;
    }

    const [latest, ...olderFiles] = files;

    for (const file of olderFiles) {
        const filePath = path.join(BACKUP_PATH, file.name);
        const ageDays = (now - file.time) / (1000 * 60 * 60 * 24);
        if (ageDays > MAX_AGE_DAYS) {
            try {
                if (fs.lstatSync(filePath).isDirectory()) {
                    fs.rmSync(filePath, { recursive: true, force: true });
                } else {
                    fs.unlinkSync(filePath);
                }
                logMessage(`🗑️ Deleted old local backup: ${file.name}`);
            } catch (err) {
                logMessage(`⚠️ Error deleting old backup: ${err.message}`);
            }
        }
    }
}

// Daily log cleanup at 3:05 AM
function scheduleLogCleanupDaily() {
    const MS_DAY = 24 * 60 * 60 * 1000;
    const now = new Date();
    const next = new Date(
        now.getFullYear(),
        now.getMonth(),
        now.getDate(),
        3,
        5,
        0,
        0
    );
    if (next <= now) next.setDate(next.getDate() + 1);

    setTimeout(() => {
        cleanupOldLogs();
        setInterval(cleanupOldLogs, MS_DAY);
    }, next.getTime() - now.getTime());
}
scheduleLogCleanupDaily();

module.exports = runBackup;

// To restore DB:
// mongorestore --db=stockAssistant --gzip "<folder path>" --drop

// // // mongorestore --db=stockAssistant --gzip "C:\StockAssistant\backups\stockAssistant-backup-2025-11-02T15-22-55-765Z\stockAssistant" --drop