// ==============================================================================
// UNIVERSAL ONEDRIVE CLEANER (Works with both Business & Consumer)
// ==============================================================================

// --- Auto-detect account type and credentials ---
function detectAccountType() {
    const url = window.location.href;
    
    // Check if it's business (SharePoint)
    if (url.includes('sharepoint.com') || url.includes('my.sharepoint.com')) {
        return 'business';
    }
    
    // Check if it's consumer
    if (url.includes('onedrive.live.com') || url.includes('onedrive.com')) {
        return 'consumer';
    }
    
    // Try to detect from page elements
    const pageContent = document.documentElement.innerHTML;
    if (pageContent.includes('_spPageContextInfo') || pageContent.includes('SharePoint')) {
        return 'business';
    }
    
    // Default to trying business first
    return 'business';
}

// --- Auto-detect site URL and starting folder ---
function detectSiteAndFolder() {
    const origin = window.location.origin;
    const fullUrl = window.location.href;
    let siteUrl = null;
    let startingFolder = null;
    let accountType = detectAccountType();

    console.log(`🔍 Detecting ${accountType} OneDrive account...`);

    // Try to extract from URL
    const pathMatch = fullUrl.match(/\/personal\/([^\/?&]+)/);
    if (pathMatch) {
        const userPart = pathMatch[1];
        siteUrl = `${origin}/personal/${userPart}`;
        startingFolder = `/personal/${userPart}/Documents`;
        console.log(`✅ Detected: ${siteUrl}`);
        return { siteUrl, startingFolder, accountType };
    }

    // For business, check SharePoint context
    if (accountType === 'business' && window._spPageContextInfo) {
        const webUrl = window._spPageContextInfo.webServerRelativeUrl;
        if (webUrl && webUrl.startsWith('/personal/')) {
            siteUrl = origin + webUrl;
            startingFolder = webUrl + '/Documents';
            console.log(`✅ Detected business site: ${siteUrl}`);
            return { siteUrl, startingFolder, accountType };
        }
    }

    // Prompt user
    const userInput = prompt(
        `Could not auto-detect. Please enter:\n` +
        `1. Your OneDrive site URL (e.g., https://your-company.sharepoint.com/personal/username)\n` +
        `2. Or your user ID (the part after /personal/)\n` +
        `Leave blank to cancel:`
    );

    if (userInput && userInput.trim()) {
        try {
            const input = userInput.trim();
            if (input.includes('http')) {
                const url = new URL(input);
                const match = url.pathname.match(/\/personal\/([^\/]+)/);
                if (match) {
                    siteUrl = `${url.origin}/personal/${match[1]}`;
                    startingFolder = `/personal/${match[1]}/Documents`;
                }
            } else {
                // Assume it's just the user ID
                siteUrl = `${origin}/personal/${input}`;
                startingFolder = `/personal/${input}/Documents`;
            }
            console.log(`✅ Using user-provided: ${siteUrl}`);
            return { siteUrl, startingFolder, accountType };
        } catch (e) {
            console.error('❌ Invalid input');
        }
    }

    throw new Error('Could not detect OneDrive site.');
}

// --- Get detection ---
const detection = detectSiteAndFolder();
const SITE_URL = detection.siteUrl;
const STARTING_FOLDER = detection.startingFolder;
const ACCOUNT_TYPE = detection.accountType;

console.log(`📌 Account Type: ${ACCOUNT_TYPE}`);
console.log(`📌 Site: ${SITE_URL}`);
console.log(`📌 Starting: ${STARTING_FOLDER}`);

// --- CONFIGURATION ---
const VERSIONS_TO_KEEP = 2;
const EXTENSIONS_TO_SKIP = [];
const CONCURRENT_REQUESTS = 5;

let requestDigest = "";
let tokenFetchTime = 0;

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// --- Get authentication headers based on account type ---
async function getValidHeaders() {
    const headers = {
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-Requested-With": "XMLHttpRequest"
    };

    if (ACCOUNT_TYPE === 'business') {
        // Business account - use request digest
        if (Date.now() - tokenFetchTime > 20 * 60 * 1000) {
            console.log("🔄 Fetching fresh security token...");
            const digestResponse = await fetch(`${SITE_URL}/_api/contextinfo`, {
                method: 'POST',
                headers: { 'Accept': 'application/json;odata=nometadata' },
                credentials: 'include'
            });
            const digestData = await digestResponse.json();
            requestDigest = digestData.FormDigestValue;
            tokenFetchTime = Date.now();
        }
        headers['X-RequestDigest'] = requestDigest;
    } else {
        // Consumer account - try to get token from cookies
        const cookies = document.cookie.split(';');
        for (const cookie of cookies) {
            const [name, value] = cookie.trim().split('=');
            if (name === 'ODB_AccessToken' || name === 'access_token') {
                headers['Authorization'] = `Bearer ${decodeURIComponent(value)}`;
                break;
            }
        }
    }

    return headers;
}

// --- Process in batches (concurrent) ---
async function processInBatches(items, batchSize, processFn) {
    for (let i = 0; i < items.length; i += batchSize) {
        const batch = items.slice(i, i + batchSize);
        await Promise.all(batch.map(item => processFn(item)));
        await sleep(150);
    }
}

// --- Get files in folder ---
async function getFilesInFolder(folderPath, headers) {
    const encodedFolder = encodeURIComponent(folderPath).replace(/'/g, "%27");
    
    // Try different API endpoints based on account type
    let url;
    if (ACCOUNT_TYPE === 'business') {
        url = `${SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${encodedFolder}')/Files?$select=ServerRelativeUrl,Name,UIVersionLabel`;
    } else {
        // Consumer uses a different endpoint
        url = `${SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${encodedFolder}')/Files?$select=ServerRelativeUrl,Name,UIVersionLabel`;
    }
    
    const response = await fetch(url, { headers, credentials: 'include' });
    if (response.ok) {
        const data = await response.json();
        return data.d ? data.d.results : [];
    }
    return [];
}

// --- Get subfolders ---
async function getSubfolders(folderPath, headers) {
    const encodedFolder = encodeURIComponent(folderPath).replace(/'/g, "%27");
    const url = `${SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${encodedFolder}')/Folders?$select=ServerRelativeUrl,Name`;
    
    const response = await fetch(url, { headers, credentials: 'include' });
    if (response.ok) {
        const data = await response.json();
        return data.d ? data.d.results : [];
    }
    return [];
}

// --- Get versions for a file ---
async function getFileVersions(filePath, headers) {
    const encodedPath = encodeURIComponent(filePath).replace(/'/g, "%27");
    
    // Business account uses GetListItemUsingPath
    let url;
    if (ACCOUNT_TYPE === 'business') {
        url = `${SITE_URL}/_api/web/GetListItemUsingPath(decodedUrl='${encodedPath}')/versions?$select=VersionId,VersionLabel,IsCurrentVersion&$top=5000`;
    } else {
        // Consumer uses direct versions endpoint
        url = `${SITE_URL}/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/Versions?$select=ID,VersionLabel,IsCurrentVersion,Created`;
    }
    
    const response = await fetch(url, { headers, credentials: 'include' });
    if (response.ok) {
        const data = await response.json();
        return data.d ? data.d.results : [];
    }
    return [];
}

// --- Delete a version ---
async function deleteVersion(filePath, version, headers) {
    const encodedPath = encodeURIComponent(filePath).replace(/'/g, "%27");
    
    if (ACCOUNT_TYPE === 'business') {
        // Business: Use RecycleByLabel
        const recycleUrl = `${SITE_URL}/_api/web/GetFileByServerRelativePath(decodedUrl='${encodedPath}')/versions/RecycleByLabel(versionLabel='${version.VersionLabel}')`;
        const response = await fetch(recycleUrl, { 
            method: 'POST', 
            headers, 
            credentials: 'include' 
        });
        return response.ok;
    } else {
        // Consumer: Try different methods
        const versionId = version.ID;
        if (versionId) {
            // Try DELETE method first
            let url = `${SITE_URL}/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/Versions(${versionId})`;
            let response = await fetch(url, { 
                method: 'DELETE', 
                headers, 
                credentials: 'include' 
            });
            
            if (response.ok) return true;
            
            // Try RecycleByLabel as fallback
            if (version.VersionLabel) {
                const recycleUrl = `${SITE_URL}/_api/web/GetFileByServerRelativeUrl('${encodedPath}')/Versions/RecycleByLabel('${version.VersionLabel}')`;
                response = await fetch(recycleUrl, { 
                    method: 'POST', 
                    headers, 
                    credentials: 'include' 
                });
                return response.ok;
            }
        }
        return false;
    }
}

// --- Clean file versions ---
async function cleanFileVersions(filePath, fileName, headers) {
    const versions = await getFileVersions(filePath, headers);
    
    if (versions.length <= VERSIONS_TO_KEEP) return;
    
    console.log(`  📄 ${fileName}: Trimming ${versions.length - VERSIONS_TO_KEEP} old versions...`);
    
    // Sort by creation date (oldest first)
    const sortedVersions = versions.sort((a, b) => {
        const dateA = a.Created ? new Date(a.Created) : new Date(0);
        const dateB = b.Created ? new Date(b.Created) : new Date(0);
        return dateA - dateB;
    });
    
    // Filter out current version
    const nonCurrentVersions = sortedVersions.filter(v => !v.IsCurrentVersion);
    
    // Keep the newest versions
    const versionsToDelete = nonCurrentVersions.slice(0, nonCurrentVersions.length - (VERSIONS_TO_KEEP - 1));
    
    for (const v of versionsToDelete) {
        const success = await deleteVersion(filePath, v, headers);
        if (success) {
            console.log(`    ✅ Recycled version ${v.VersionLabel || v.ID}`);
        } else {
            console.log(`    ❌ Failed to recycle version ${v.VersionLabel || v.ID}`);
        }
        await sleep(100);
    }
}

// --- Process folder recursively ---
async function processFolder(folderPath) {
    console.log(`\n📂 Scanning: ${folderPath}`);
    const headers = await getValidHeaders();

    await sleep(200);

    // 1. Get files
    const files = await getFilesInFolder(folderPath, headers);
    const filesToScan = [];

    for (const file of files) {
        const fileName = file.Name.toLowerCase();
        if (EXTENSIONS_TO_SKIP.some(ext => fileName.endsWith(ext))) continue;
        const versionNum = parseFloat(file.UIVersionLabel);
        if (versionNum <= VERSIONS_TO_KEEP) continue;
        filesToScan.push(file);
    }

    if (filesToScan.length > 0) {
        console.log(`   🔥 Found ${filesToScan.length} files requiring history cleanup...`);
        await processInBatches(filesToScan, CONCURRENT_REQUESTS, async (file) => {
            await cleanFileVersions(file.ServerRelativeUrl, file.Name, await getValidHeaders());
        });
    }

    // 2. Get subfolders
    const subfolders = await getSubfolders(folderPath, headers);
    const filteredFolders = subfolders.filter(f => 
        f.Name !== "Forms" && 
        f.Name !== "Attachments" && 
        !f.Name.startsWith("_")
    );

    for (const subfolder of filteredFolders) {
        await processFolder(subfolder.ServerRelativeUrl);
        await sleep(200);
    }
}

// --- Main function ---
async function startCleanup() {
    console.log("\n🚀 Starting OneDrive version cleanup...");
    console.log(`📌 Account Type: ${ACCOUNT_TYPE}`);
    console.log(`📌 Keeping ${VERSIONS_TO_KEEP} most recent versions per file`);
    console.log(`📌 Starting from: ${STARTING_FOLDER}`);
    console.log("⏳ This may take a while...\n");

    try {
        await processFolder(STARTING_FOLDER);
        console.log("\n🎉 Cleanup complete!");
        console.log("💡 Check your Recycle Bin to verify deleted versions.");
        console.log("💡 Empty the Recycle Bin manually to reclaim storage space.");
    } catch (error) {
        console.error("❌ Error during cleanup:", error);
        console.log("\n💡 If you're on consumer OneDrive, try:");
        console.log("   1. Navigate to your Documents folder");
        console.log("   2. Refresh the page");
        console.log("   3. Run the script again");
    }
}

// --- Run it ---
startCleanup();
