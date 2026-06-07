# OneDrive_cleanVersionHistory
Clean your full Onedrive because of Version History
How to find OneDrive ID
Method 1: The Address Bar (Easiest)

log into onedrive.live.com.

Click on My files on the left sidebar.

Look at the URL in the browser's address bar. It will usually look something like this:
https://onedrive.live.com/?id=%2Fpersonal%2F1234567890ABCDEF%2FDocuments

Those 16 characters right after %2Fpersonal%2F are their ID.

Method 2: The Network Tab (Foolproof)
If the URL doesn't show it, they can do exactly what you did:

Open the Developer Tools (F12) and go to the Network tab.

Right-click any file in their OneDrive and click Version history.

Delete an old version.

Look at the Network tab for the red/black request and check the URL for the /personal/xxxxxxxxxxxxxxxx/ string.

## For Business & University Accounts (`cleaner-business.js`)

If you are using a work or school account, your OneDrive runs on enterprise SharePoint servers.

### 1. Find your SharePoint URL
1. Log into your work/school OneDrive in your web browser.
2. Look at the address bar. You need to copy the base part of the URL up to the end of your email address. 
   * **Example:** `https://university-my.sharepoint.com/personal/jdoe_university_edu`
   * Do not include `/_layouts/...` or `/?id=...` at the end.

### 2. Configure the Script
1. Open the `cleaner-business.js` file and copy the code.
2. Paste it into a text editor.
3. At the top of the code, replace `"PASTE_YOUR_URL_HERE"` with the URL you just copied. 
4. *(Optional)* Change `EMPTY_RECYCLE_BIN` to `true` if you want the script to permanently delete the history to instantly get your storage back.

### 3. Run the Cleanup
1. Go back to your open OneDrive tab.
2. Press **F12** to open Developer Tools, and navigate to the **Console** tab.
3. Paste the code and press **Enter**.
