# Excel Mapper Local Trial

Excel Mapper runs locally on your computer. By default it uses deterministic coded mapping logic and does not send uploaded files to Syntax GenAI Studio.

## Before You Start

Install Node.js LTS from https://nodejs.org.

You only need a Syntax GenAI Studio API key if you choose **External AI Agent** mode. Paste it into the app when it opens. The key is saved locally in your browser on your computer, so it stays after refreshes. Use **Forget** in the app if you want to remove it.

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

Files stay on your computer when using the default **Deterministic Mapper** mode. If you choose **External AI Agent** mode, the app sends the uploaded workbooks and prompt to Syntax GenAI Studio for processing.

Do not upload client-confidential files unless you are authorized to send them to Syntax GenAI Studio.
