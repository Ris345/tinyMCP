import { google } from "googleapis";
import path from "path";
import { fileURLToPath } from "url";


const CREDENTIALS_PATH = "./service.json";

const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];

const auth = new google.auth.GoogleAuth({
  keyFile: CREDENTIALS_PATH,
  scopes: SCOPES,
});

const sheets = google.sheets({ version: "v4", auth });

export async function getSheetData(spreadsheetId, range) {
  const client = await auth.getClient(); // Ensure client is authorized
  const res = await sheets.spreadsheets.values.get({
    auth: client,
    spreadsheetId,
    range,
  });
  return res.data.values;
}

export async function appendSheetData(spreadsheetId, range, values) {
  const client = await auth.getClient(); // Ensure client is authorized
  await sheets.spreadsheets.values.append({
    auth: client,
    spreadsheetId,
    range,
    valueInputOption: "RAW",
    requestBody: { values },
  });
}
