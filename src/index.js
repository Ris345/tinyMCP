// Fully working MCP server named "excel_automator" with CRUD operations
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { google } from "googleapis";
import fs from 'fs';
import { fileURLToPath } from "url";
import path from "path";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const CREDENTIALS_PATH = path.join(__dirname, "service.json");

const SPREADSHEET_ID = "1jnyASVXIfK-lTBJKRxarsJstQltzsBKF-OiAXJrahO8";
const RANGE = "Sheet1!A1:D";
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];

let credentials;
try {
  if (!fs.existsSync(CREDENTIALS_PATH)) throw new Error("service.json not found");
  credentials = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf-8'));
  console.log("✅ Loaded service account credentials successfully.");
} catch (err) {
  console.error("❌ Failed to load service account credentials:", err);
  process.exit(1);
}

let sheets;
try {
  const auth = new google.auth.GoogleAuth({ credentials, scopes: SCOPES });
  sheets = google.sheets({ version: "v4", auth });
  console.log("✅ Google Sheets client created.");
} catch (err) {
  console.error("❌ Error creating Google Sheets client:", err);
  process.exit(1);
}

async function getSheetData(spreadsheetId, range) {
  const res = await sheets.spreadsheets.values.get({ spreadsheetId, range });
  return res.data.values || [];
}

async function appendSheetData(spreadsheetId, range, values) {
  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range,
    valueInputOption: "RAW",
    requestBody: { values },
  });
}

async function updateRowData(rowNumber, values) {
  const range = `Sheet1!A${rowNumber}:D${rowNumber}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range,
    valueInputOption: "RAW",
    requestBody: { values: [values] },
  });
}

async function deleteRow(rowNumber) {
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        {
          deleteDimension: {
            range: {
              sheetId: 0,
              dimension: "ROWS",
              startIndex: rowNumber - 1,
              endIndex: rowNumber,
            },
          },
        },
      ],
    },
  });
}

const server = new McpServer({
  name: "excel_automator",
  version: "1.0.0",
  capabilities: { resources: {}, tools: {} },
});

server.tool("read-sheet", "Read rows from a Google Sheet", {}, async () => {
  try {
    const rows = await getSheetData(SPREADSHEET_ID, RANGE);
    const formatted = rows.map((r, i) => `${i + 1}. ${r.join(" | ")}`).join("\n");
    return {
      content: [{ type: "text", text: `Sheet Data:\n\n${formatted || "No data found."}` }],
    };
  } catch (err) {
    return { content: [{ type: "text", text: "Error reading sheet: " + err.message }] };
  }
});

server.tool("write-sheet", "Append a new row to Google Sheet", {
  values: z.array(z.string()).length(4).describe("Values for apples, bananas, oranges, grapes"),
}, async ({ values }) => {
  try {
    const headers = ["apples", "bananas", "oranges", "grapes"];
    const sheetData = await getSheetData(SPREADSHEET_ID, "Sheet1!A1:D1");
    if (!sheetData.length || sheetData[0].join() !== headers.join()) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: "Sheet1!A1:D1",
        valueInputOption: "RAW",
        requestBody: { values: [headers] },
      });
    }
    await appendSheetData(SPREADSHEET_ID, "Sheet1!A2", [values]);
    return {
      content: [{ type: "text", text: `✅ Added inventory counts: ${values.join(", ")}` }],
    };
  } catch (err) {
    return { content: [{ type: "text", text: "Error writing to sheet: " + err.message }] };
  }
});

server.tool("update-sheet-row", "Update a specific row in the Google Sheet", {
  rowNumber: z.number().min(2).describe("Row number to update (starting from 2)"),
  values: z.array(z.string()).length(4).describe("New values for the row (apples, bananas, oranges, grapes)"),
}, async ({ rowNumber, values }) => {
  try {
    await updateRowData(rowNumber, values);
    return { content: [{ type: "text", text: `✅ Updated row ${rowNumber} successfully.` }] };
  } catch (err) {
    return { content: [{ type: "text", text: "Error updating row: " + err.message }] };
  }
});

server.tool("delete-sheet-row", "Delete a specific row from the Google Sheet", {
  rowNumber: z.number().min(2).describe("Row number to delete (starting from 2)"),
}, async ({ rowNumber }) => {
  try {
    await deleteRow(rowNumber);
    return { content: [{ type: "text", text: `🗑️ Deleted row ${rowNumber} successfully.` }] };
  } catch (err) {
    return { content: [{ type: "text", text: "Error deleting row: " + err.message }] };
  }
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.log("🚀 Excel Automator MCP server is running!");
}

main().catch((err) => {
  console.error("❌ Fatal error in main():", err);
  process.exit(1);
});



// Use the write-sheet tool to log 15 apples, 30 bananas, 20 oranges, 10 grapes.
// Use the update-sheet-row tool to replace row 3 with 12 apples, 25 bananas, 18 oranges, 14 grapes.
// Use the delete-sheet-row tool to remove row 4 from the sheet.
