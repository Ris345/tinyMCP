# Excel Automator MCP

This is a Model Context Protocol (MCP) server that enables LLMs to interact with Google Sheets. It provides tools to read, write, update, and delete rows in a Google Sheet.

## Problem Solved
This project allows AI agents to directly manipulate data in Google Sheets, enabling automation of tasks like inventory management, data entry, and report generation directly from a chat interface.

## Prerequisites
- Node.js installed
- A Google Cloud Service Account with access to the Google Sheets API
- A `service.json` credential file from your Google Cloud project

## Setup

1.  **Install Dependencies**
    ```bash
    npm install
    ```

2.  **Configure Credentials**
    Place your Google Cloud service account credentials file named `service.json` in the `src/` directory.
    
    *Note: The application expects `service.json` to be available to the running script. Ensure it is copied to the build directory or accessible.*

3.  **Build the Project**
    ```bash
    npm run build
    ```
    *You may need to manually copy `src/service.json` to `build/service.json` if the build script doesn't handle it.*

## Usage

To run the server:

```bash
node build/index.js
```

### Available Tools

-   `read-sheet`: Read all rows from the configured Google Sheet.
-   `write-sheet`: Append a new row (specifically formatted for inventory: apples, bananas, oranges, grapes).
-   `update-sheet-row`: Update a specific row.
-   `delete-sheet-row`: Delete a specific row.

## Configuration
The spreadsheet ID is currently hardcoded in `src/index.js`. You may need to update `SPREADSHEET_ID` with your own sheet ID.
