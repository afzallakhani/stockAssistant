// backup.js
const { exec } = require("child_process");
const path = require("path");
const fs = require("fs");

// ---- CONFIG ----
const DB_NAME = "stockAssistant"; // replace with your actual DB name
const BACKUP_PATH = path.join(__dirname, "backups");
const DATE = new Date().toISOString().replace(/[:.]/g, "-");
const OUT_PATH = path.join(BACKUP_PATH, `${DB_NAME}-backup-${DATE}`);

// Create backup folder if not exists
if (!fs.existsSync(BACKUP_PATH)) fs.mkdirSync(BACKUP_PATH, { recursive: true });

// ---- COMMAND ----
const cmd = `mongodump --db=${DB_NAME} --out="${OUT_PATH}" --gzip`;

// ---- EXECUTE ----
console.log(`\n🟡 Starting MongoDB backup for: ${DB_NAME}`);
exec(cmd, (error, stdout, stderr) => {
  if (error) {
    console.error(`❌ Backup failed: ${error.message}`);
    return;
  }
  if (stderr) console.warn(`⚠️ Warning: ${stderr}`);
  console.log(`✅ Backup completed successfully! Saved to: ${OUT_PATH}`);
});

// mongorestore --db=your_database_name --gzip "C:\StockAssistant\backups\stockAssistant-backup-2025-11-02T10-45-32\stockAssistant"
