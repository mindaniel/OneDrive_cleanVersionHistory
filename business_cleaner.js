// ==============================================================================
// ONEDRIVE FOR BUSINESS / UNIVERSITY - VERSION CLEANUP SCRIPT
// ==============================================================================

// --- CONFIGURATION ---
// 1. Paste your personal SharePoint URL here. 
// It usually looks like: "https://yourcompany-my.sharepoint.com/personal/your_email_com"
const ONEDRIVE_URL = "PASTE_YOUR_URL_HERE";

// 2. How many recent versions do you want to keep as a safety net?
const VERSIONS_TO_KEEP = 2;

// 3. Automatically permanently delete the old versions to free up space instantly?
const EMPTY_RECYCLE_BIN = false;
// ---------------------

async function startCleanup() {
    if (ONEDRIVE_URL === "PASTE_YOUR_URL_HERE") {
        console.error("🛑 STOP: You need to paste your OneDrive URL in the configuration block at the top!");
        return;
    }

    // Auto-parse the URL to build the exact SharePoint API paths
    let cleanUrl = ONEDRIVE_URL.split('?')[0].replace(/\/$/, ""); // Remove queries and trailing slashes
    if (cleanUrl.endsWith('/Documents')) cleanUrl = cleanUrl.replace('/Documents', '');
    if (cleanUrl.includes('/_layouts')) cleanUrl = cleanUrl.split('/_layouts')[0];

    const urlObj = new URL(cleanUrl);
    const SITE_URL = cleanUrl;
    const STARTING_FOLDER = `${urlObj.pathname}/Documents`;

    console.log(`🚀 Initializing cleanup for: ${SITE_URL}`);
    console.log(`📁 Target Root: ${STARTING_FOLDER}`);

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
        console.log(`\n📂 Scanning: ${folderPath}`);
        const encodedFolder = encodeURIComponent(folderPath).replace(/'/g, "%27");
        const headers = await getValidHeaders();

        // 1. Process Files
        const getFilesUrl = `${SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${encodedFolder}')/Files?$select=ServerRelativeUrl,Name`;
        const filesResponse = await fetch(getFilesUrl, { headers });
        
        if (filesResponse.ok) {
            const filesData = await filesResponse.json();
            const files = filesData.d ? filesData.d.results : [];
            for (const file of files) {
                await cleanFileVersions(file.ServerRelativeUrl, file.Name, headers);
            }
        }

        // 2. Process Subfolders
        const getFoldersUrl = `${SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${encodedFolder}')/Folders?$select=ServerRelativeUrl,Name`;
        const foldersResponse = await fetch(getFoldersUrl, { headers: await getValidHeaders() });
        
        if (foldersResponse.ok) {
            const foldersData = await foldersResponse.json();
            const subfolders = foldersData.d ? foldersData.d.results : [];
            for (const subfolder of subfolders) {
                if (subfolder.Name === "Forms" || subfolder.Name === "Attachments" || subfolder.Name.startsWith("_")) continue;
                await processFolder(subfolder.ServerRelativeUrl);
            }
        }
    }

    async function cleanFileVersions(filePath, fileName, headers) {
        const encodedPath = encodeURIComponent(filePath).replace(/'/g, "%27");
        const getVersionsUrl = `${SITE_URL}/_api/web/GetListItemUsingPath(decodedUrl='${encodedPath}')/versions?$select=VersionId,VersionLabel,IsCurrentVersion&$top=5000`;

        const versionsResponse = await fetch(getVersionsUrl, { headers: await getValidHeaders() });
        if (!versionsResponse.ok) return;

        const versionsData = await versionsResponse.json();
        const versions = versionsData.d ? versionsData.d.results : [];

        if (versions.length <= VERSIONS_TO_KEEP) return; 

        console.log(`  📄 ${fileName}: Found ${versions.length} versions. Trimming...`);
        const versionsToDelete = versions.slice(VERSIONS_TO_KEEP);

        for (const v of versionsToDelete) {
            const recycleUrl = `${SITE_URL}/_api/web/GetFileByServerRelativePath(decodedUrl='${encodedPath}')/versions/RecycleByLabel(versionLabel='${v.VersionLabel}')`;
            const delResponse = await fetch(recycleUrl, { method: 'POST', headers: await getValidHeaders() });
            
            if (delResponse.ok) console.log(`    ✅ Recycled version ${v.VersionLabel}`);
            else console.error(`    ❌ Failed to recycle version ${v.VersionLabel}`);
        }
    }

    async function emptyRecycleBin() {
        console.log("\n🗑️ Commencing two-stage Recycle Bin purge...");
        const headers = await getValidHeaders();

        console.log("   Moving items to Second-Stage...");
        await fetch(`${SITE_URL}/_api/site/getrecyclebinitems(rowLimit='5000',isAscending='false',itemState=1,orderBy=3)/MoveAllToSecondStage`, { method: 'POST', headers });
        
        console.log("   Obliterating Second-Stage to reclaim storage...");
        const res = await fetch(`${SITE_URL}/_api/site/getrecyclebinitems(rowLimit='5000',isAscending='false',itemState=2,orderBy=3)/DeleteAllSecondStageItems`, { method: 'POST', headers });
        
        if (res.ok) console.log("   ✅ Storage space successfully reclaimed!");
    }

    await processFolder(STARTING_FOLDER);
    
    if (EMPTY_RECYCLE_BIN) await emptyRecycleBin();
    else console.log("\n🎉 Cleanup complete! Check your Recycle Bin to verify, then empty it manually.");
}

startCleanup();
