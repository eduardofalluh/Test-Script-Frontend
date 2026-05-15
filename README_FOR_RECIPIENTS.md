# Excel Mapper Local Trial

Excel Mapper runs locally on your computer. By default it uses deterministic coded mapping logic and does not send uploaded files to Syntax GenAI Studio.

## Before You Start

Install Node.js LTS from https://nodejs.org.

You only need a Syntax GenAI Studio API key if you choose **AI-Assisted Deterministic**, **SAP CALM Test Script**, or **External AI Agent** mode. Paste it into the app when it opens. The key is saved locally in your browser on your computer, so it stays after refreshes. Use **Forget** in the app if you want to remove it.

## SAP CALM Test Script Mode

Click **SAP CALM Test Script** when you want to use the included SAP Cloud ALM Test Cases template. The app loads the template automatically. Upload the source workbook, then use the mapping prompt to clarify which source sheets and columns should become test case names, owners, activities, actions, instructions, expected results, evidence flags, tags, and any other SAP CALM fields.

## Start On Mac

Double-click:

```text
START_EXCEL_MAPPER_MAC.command
```

If macOS blocks the file, right-click it, choose **Open**, then confirm.

## Start On Windows

Double-click:

```text
START_EXCEL_MAPPER_WINDOWS.bat
```

## Browser URL

The app opens at:

```text
http://localhost:3010
```

Keep the terminal window open while using the app. Close the terminal window to stop Excel Mapper.

## First Run

The first run installs dependencies with `npm install`, so it can take a few minutes. Later starts are faster.

## Privacy Note

Files stay on your computer when using the default **Deterministic Mapper** mode. If you choose **AI-Assisted Deterministic** or **SAP CALM Test Script** mode, workbook summaries and your prompt are sent to Syntax GenAI Studio so it can return a JSON mapping plan, then Excel Mapper writes the workbook locally. If you choose **External AI Agent** mode, the app sends the uploaded workbooks and prompt to Syntax GenAI Studio for processing.

Do not upload client-confidential files unless you are authorized to send them to Syntax GenAI Studio.
