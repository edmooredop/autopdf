## **A Guide to Automating Production Documents**

**The Problem:** You're on set, and a last-minute change comes through. Do you have the latest call sheet? Are you frantically searching through your emails for the most recent unit list or movement order?

**The Solution:** This guide will help you set up a free, "set it and forget it" system that automatically saves the latest versions of key production documents to a dedicated folder in your Google Drive. Using an iOS Shortcut, you can then access the latest file with a single tap, confident you're always looking at the most current version.

### **How It Works: The Big Picture**

This system has three parts:

1.  **Google Drive:** A central folder acts as the home for your most current documents.
2.  **Google Apps Script:** A smart "robot" that lives in your Google account. It constantly watches your new emails for specific attachments.
3.  **iOS Shortcut:** A button on your iPhone home screen that instantly opens the latest file.

-----

### **Part 1: The Setup (10-Minute Guide)**

Follow these steps to build your automated system.

#### **Step 1: Create Your Google Drive Folder**

This is where your up-to-date documents will live.

1.  Go to [Google Drive](https://drive.google.com).

2.  Create a new folder. Let's call it `Film Production` or something similar.

3.  Open that new folder. Look at the URL in your browser's address bar. You need to copy the **Folder ID**, which is the long string of letters and numbers at the end of the URL.

    \*Example: `https://drive.google.com/drive/folders/`**`1a2b3c4d5e6f7g8h9i0j`** \<- Copy this part.

#### **Step 2: Create the Google Apps Script**

1.  Go to [script.google.com](https://script.google.com).
2.  Click **New project** in the top-left.
3.  Delete the sample `function myFunction() { ... }` code in the editor window.

#### **Step 3: Paste and Configure the Code**

1.  Copy the entire code block below and paste it into the empty script editor.
2.  Find the line that says `const ROOT_FOLDER_ID = 'YOUR_FOLDER_ID_HERE';`.
3.  Replace `YOUR_FOLDER_ID_HERE` with the actual **Folder ID** you copied in Step 1. Make sure to keep the single quotes around it.

<!-- end list -->

```javascript
/**
 * A production-grade script to automatically find and save specific email attachments.
 * This script is designed to be robust, preventing common issues and handling errors gracefully.
 *
 * v5: Implements LockService, try/catch error handling, and folder caching.
 *
 * @OnlyCurrentDoc
 */

// --- CONFIGURATION: PASTE YOUR FOLDER ID BELOW ---
// Find this by opening your target Google Drive folder and copying the last part of the URL.
const ROOT_FOLDER_ID = 'YOUR_FOLDER_ID_HERE';

// This section defines which documents to look for. You can customize this.
const FILE_RULES = {
  'callsheet.pdf': { keywords: ['call sheet', 'callsheet', 'CS'], folderName: 'Call Sheets' },
  'prepdiary.pdf': { keywords: ['prep diary', 'prepdiary'], folderName: 'Prep Diaries' },
  'unitlist.pdf': { keywords: ['unit list', 'unitlist'], folderName: 'Unit Lists' },
  'movementorder.pdf': { keywords: ['Movement Order', 'MO'], folderName: 'Movement Orders' }
};

// --- SCRIPT LOGIC (No need to edit below this line) ---

function saveMostRecentAttachments() {
  // Use a lock to prevent simultaneous runs. Wait up to 30 seconds.
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.warn("Could not obtain lock. Another instance of the script is likely running. Skipping this run.");
    return;
  }

  console.log("Script run started at: " + new Date().toLocaleTimeString());
  
  // Wrap all logic in a try...catch...finally block for resilience.
  try {
    // Cache for Drive folder objects to reduce API calls.
    const folderCache = {};
    const rootFolder = DriveApp.getFolderById(ROOT_FOLDER_ID);
    folderCache[ROOT_FOLDER_ID] = rootFolder; 

    const processedTypes = new Set();
    const searchQuery = 'is:unread has:attachment filename:pdf';
    const threads = GmailApp.search(searchQuery);

    console.log(`Found ${threads.length} unread email thread(s) with PDF attachments to process.`);
    if (threads.length === 0) {
      console.log("No new emails to process. Ending script run.");
      return;
    }

    threads.forEach(thread => {
      const messages = thread.getMessages();
      let threadMarkedAsRead = false;

      messages.reverse().forEach(message => {
        if (!message.isUnread()) return;

        const attachments = message.getAttachments();
        attachments.forEach(attachment => {
          if (attachment.getContentType() !== 'application/pdf') return;
          const attachmentName = attachment.getName();
          console.log(`-- Checking attachment: '${attachmentName}'`);

          for (const targetFilename in FILE_RULES) {
            if (processedTypes.has(targetFilename)) continue;

            const rule = FILE_RULES[targetFilename];
            const foundKeyword = rule.keywords.find(keyword => {
              const escapedKeyword = keyword.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
              const regex = new RegExp(`(^|[^A-Za-z0-9])${escapedKeyword}([^A-Za-z0-9]|$)`, 'i');
              return regex.test(attachmentName);
            });

            if (foundKeyword) {
              console.log(`   └─> Match found on keyword '${foundKeyword}'. Processing now.`);
              const typeFolder = getOrCreateSubfolder(rootFolder, rule.folderName, folderCache);
              const oldFolder = getOrCreateSubfolder(typeFolder, 'old', folderCache);

              processFile(typeFolder, oldFolder, targetFilename, attachment);
              processedTypes.add(targetFilename);
              if (!threadMarkedAsRead) {
                thread.markRead();
                threadMarkedAsRead = true;
              }
            }
          }
        });
      });
    });
    console.log("Script run finished successfully. Processed types: ", Array.from(processedTypes));

  } catch (e) {
    console.error("An error occurred during script execution: " + e.toString());
    console.error("Stack Trace: " + e.stack);

  } finally {
    // ALWAYS release the lock at the end.
    lock.releaseLock();
    console.log("Script lock released. Execution complete.");
  }
}

function getOrCreateSubfolder(parentFolder, folderName, cache) {
  const cacheKey = parentFolder.getId() + '_' + folderName;
  if (cache[cacheKey]) { return cache[cacheKey]; }
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    const folder = folders.next();
    cache[cacheKey] = folder;
    return folder;
  }
  console.log(`   └─> Creating subfolder: '${folderName}' inside '${parentFolder.getName()}'`);
  const newFolder = parentFolder.createFolder(folderName);
  cache[cacheKey] = newFolder;
  return newFolder;
}

function processFile(typeFolder, oldFolder, targetFilename, attachment) {
  const existingFiles = typeFolder.getFilesByName(targetFilename);
  if (existingFiles.hasNext()) {
    const oldFile = existingFiles.next();
    const timestamp = new Date().toISOString().replace(/:/g, '-').slice(0, 19);
    const oldFilename = `${targetFilename.replace('.pdf', '')}_${timestamp}.pdf`;
    console.log(`       └─> Moving existing file to 'old' folder.`);
    oldFile.moveTo(oldFolder).setName(oldFilename);
  }
  const blob = attachment.copyBlob();
  const newFile = typeFolder.createFile(blob).setName(targetFilename);
  console.log(`       └─> Saved new file to '${typeFolder.getName()}' folder.`);
}
```

#### **Step 4: Set the Automatic Trigger**

This makes the script run automatically for you.

1.  In the script editor, click the **Triggers** icon (looks like a clock) on the left-hand menu.
2.  Click the **+ Add Trigger** button in the bottom-right.
3.  Set up the trigger with these exact options:
      * **Choose which function to run**: `saveMostRecentAttachments`
      * **Choose which deployment should run**: `Head`
      * **Select event source**: `Time-driven`
      * **Select type of time-based trigger**: `Minutes timer`
      * **Select minute interval**: `Every 15 minutes`
4.  Click **Save**.

#### **Step 5: Grant Permissions**

The first time you save the trigger, Google will ask for permission for the script to access your Gmail and Drive.

1.  Choose your Google account.
2.  You will see a "Google hasn't verified this app" warning. This is **normal and expected** because *you* wrote the code.
3.  Click **Advanced**, then click **Go to [Your Project Name] (unsafe)**.
4.  Review the permissions it needs (to read emails and manage files) and click **Allow**.

**Your automation is now live\!** It will check for new documents every 15 minutes.

-----

### **Part 2: One-Tap Access with an iOS Shortcut**

Now let's create the button on your iPhone to instantly open the latest file.

1.  Make sure you have the official **Google Drive app** installed on your iPhone.
2.  Open the **Shortcuts** app and tap the **+** icon to create a new shortcut.
3.  Tap **"Add Action"**.
4.  Search for `Drive` and select the action **"Get File from Drive"**.
5.  In the action that appears, tap the blue "File" text. A Google Drive file picker will appear. Navigate to the file you want this shortcut to open (e.g., `Film Production` \> `Call Sheets` \> `callsheet.pdf`).
6.  Tap the search bar at the bottom again and search for the action **"Quick Look"**. Tap to add it. This will display the file you just selected.
7.  Name your shortcut at the top (e.g., "Latest Call Sheet") and tap **Done**.

You can add this shortcut to your iPhone's Home Screen for true one-tap access. Repeat these steps to create separate shortcuts for your Unit List, Movement Order, etc.

That's it\! Your digital assistant is now working for you, ensuring you always have the most current documents right at your fingertips.
