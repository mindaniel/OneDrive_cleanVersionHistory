# OneDrive_cleanVersionHistory
Clean your full Onedrive because of Version History
How to find OneDrive ID
Method 1: The Address Bar (Easiest)

Tell them to log into onedrive.live.com.

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
