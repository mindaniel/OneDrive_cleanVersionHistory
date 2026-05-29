// ==============================================================================
// ONEDRIVE VERSION HISTORY CLEANUP SCRIPT
// ==============================================================================

// --- CONFIGURATION ---
// 1. Replace the 16 characters below with your own personal OneDrive ID
const ONEDRIVE_ID = "YOUR_16_CHAR_ID_HERE"; 

// 2. How many recent versions do you want to keep as a safety net? (Default is 2)
const VERSIONS_TO_KEEP = 2;
// ---------------------

const SITE_URL = `https://onedrive.live.com/personal/${ONEDRIVE_ID}`;
const STARTING_FOLDER = `/personal/${ONEDRIVE_ID}/Documents`; 

let requestDigest = "";
let tokenFetchTime = 0;

async function getValidHeaders() {
    if (Date.now() - tokenFetchTime > 20 * 60 * 1000) {
        console.log("🔄 Fetching fresh security token...");
        const digestResponse = await fetch(`${SITE_URL}/_api/contextinfo`, {
            method: 'POST',
            headers: { 'Accept': 'application/json;odata=nometadata' }
        });
        const digestData = await digestResponse.json();
        requestDigest = digestData.FormDigestValue;
        tokenFetchTime = Date.now();
    }
    
    return {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": requestDigest
    };
}

async function processFolder(folderPath) {
    console.log(`\n📂 Scanning folder: ${folderPath}`);
    const encodedFolder = encodeURIComponent(folderPath).replace(/'/g, "%27");
    const headers = await getValidHeaders();

    // 1. Process Files
    const getFilesUrl = `${SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${encodedFolder}')/Files?$select=ServerRelativeUrl,Name`;
    const filesResponse = await fetch(getFilesUrl, { headers });
    
    if (filesResponse.ok) {
        const filesData = await filesResponse.json();
        const files = filesData.d ? filesData.d.results : [];
        
        for (const file of files) {
            await cleanFileVersions(file.ServerRelativeUrl, file.Name);
        }
    }

    // 2. Process Subfolders recursively
    const getFoldersUrl = `${SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${encodedFolder}')/Folders?$select=ServerRelativeUrl,Name`;
    const headersForFolders = await getValidHeaders();
    const foldersResponse = await fetch(getFoldersUrl, { headers: headersForFolders });
    
    if (foldersResponse.ok) {
        const foldersData = await foldersResponse.json();
        const subfolders = foldersData.d ? foldersData.d.results : [];
        
        for (const subfolder of subfolders) {
            if (subfolder.Name === "Forms" || subfolder.Name === "Attachments" || subfolder.Name.startsWith("_")) continue;
            await processFolder(subfolder.ServerRelativeUrl);
        }
    }
}

async function cleanFileVersions(filePath, fileName) {
    const encodedPath = encodeURIComponent(filePath).replace(/'/g, "%27");
    const headers = await getValidHeaders();
    const getVersionsUrl = `${SITE_URL}/_api/web/GetListItemUsingPath(decodedUrl='${encodedPath}')/versions?$select=VersionId,VersionLabel,IsCurrentVersion&$top=5000`;

    const versionsResponse = await fetch(getVersionsUrl, { headers });
    if (!versionsResponse.ok) return;

    const versionsData = await versionsResponse.json();
    const versions = versionsData.d ? versionsData.d.results : [];

    if (versions.length <= VERSIONS_TO_KEEP) return; 

    console.log(`  📄 ${fileName}: Found ${versions.length} versions. Trimming...`);
    const versionsToDelete = versions.slice(VERSIONS_TO_KEEP);

    for (const v of versionsToDelete) {
        const recycleUrl = `${SITE_URL}/_api/web/GetFileByServerRelativePath(decodedUrl='${encodedPath}')/versions/RecycleByLabel(versionLabel='${v.VersionLabel}')`;
        const actionHeaders = await getValidHeaders();
        const delResponse = await fetch(recycleUrl, { method: 'POST', headers: actionHeaders });
        
        if (delResponse.ok) {
            console.log(`    ✅ Recycled version ${v.VersionLabel}`);
        } else {
            console.error(`    ❌ Failed to recycle version ${v.VersionLabel}`);
        }
    }
}

async function startCleanup() {
    if (ONEDRIVE_ID === "YOUR_16_CHAR_ID_HERE") {
        console.error("🛑 STOP: You need to put your 16-character OneDrive ID in the configuration block at the top!");
        return;
    }
    console.log("🚀 Initializing personal OneDrive cleanup...");
    await processFolder(STARTING_FOLDER);
    console.log("\n🎉 Total recursive cleanup complete! Empty your Recycle Bin to reclaim the space.");
}

startCleanup();
