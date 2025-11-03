// // utils/backupHelper.js
const { exec } = require("child_process");
const path = require("path");
const fs = require("fs");

// const DB_NAME = "your_database_name"; // change this
// const BACKUP_PATH = path.join(__dirname, "../backups");

// function runBackup(trigger = "manual") {
//   const DATE = new Date().toISOString().replace(/[:.]/g, "-");
//   const OUT_PATH = path.join(BACKUP_PATH, `${DB_NAME}-backup-${DATE}`);

//   if (!fs.existsSync(BACKUP_PATH))
//     fs.mkdirSync(BACKUP_PATH, { recursive: true });

//   const cmd = `mongodump --db=${DB_NAME} --out="${OUT_PATH}" --gzip`;

//   console.log(`🟡 Backup started (${trigger})...`);
//   exec(cmd, (error, stdout, stderr) => {
//     if (error) return console.error(`❌ Backup failed: ${error.message}`);
//     console.log(
//       `✅ Backup done (${trigger}) at ${new Date().toLocaleString()}`
//     );
//   });
// }

// module.exports = runBackup;
let lastBackup = 0;
let timeoutId;

function runBackup(trigger = "manual") {
  const now = Date.now();
  const MIN_GAP = 60 * 1000; // 1 minute

  clearTimeout(timeoutId);

  if (now - lastBackup < MIN_GAP) {
    // Wait 1 minute after the last change before backing up
    timeoutId = setTimeout(() => doBackup(trigger), MIN_GAP);
  } else {
    doBackup(trigger);
  }
}

function doBackup(trigger) {
  lastBackup = Date.now();
  const { exec } = require("child_process");
  const path = require("path");
  const fs = require("fs");

  const DB_NAME = "stockAssistant";
  const BACKUP_PATH = path.join(__dirname, "../backups");
  const DATE = new Date().toISOString().replace(/[:.]/g, "-");
  const OUT_PATH = path.join(BACKUP_PATH, `${DB_NAME}-backup-${DATE}`);

  if (!fs.existsSync(BACKUP_PATH))
    fs.mkdirSync(BACKUP_PATH, { recursive: true });

  const cmd = `mongodump --db=${DB_NAME} --out="${OUT_PATH}" --gzip`;

  console.log(`🟡 Backup started (${trigger})...`);
  exec(cmd, (error) => {
    if (error) console.error(`❌ Backup failed: ${error.message}`);
    else console.log(`✅ Backup completed (${trigger})`);
  });
}

module.exports = runBackup;
