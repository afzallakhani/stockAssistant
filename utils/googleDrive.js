// utils/googleDrive.js
const fs = require("fs");
const path = require("path");
const { google } = require("googleapis");

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
        fields: "id, webViewLink",
    });
    console.log(`☁️  Uploaded to Drive: ${res.data.webViewLink}`);
    return res.data.webViewLink;
}

module.exports = uploadToDrive;