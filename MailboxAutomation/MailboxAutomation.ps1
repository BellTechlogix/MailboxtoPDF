<#
    MailboxAutomation.ps1
    Created By - Kristopher Roy
    Created On - 2026-03-20
    Revised On - 2026-04-14
    Revised On - 2026-05-11 (Added better file filtering for signature blocks and name filtering for Display Names)
    Revised On - 2026-05-14 (Added legacy xls handling bypass, commented out xls conversion logic for now))
    Modules Required: Microsoft.Graph.Authentication, Microsoft.Graph.Users, Microsoft.Graph.Mail, ImportExcel

    .Important
    - EXECUTION ENVIRONMENT: This script is hardcoded for Linux paths (/opt/ap-automation/). 
      It will fail on Windows without WSL or significant path refactoring.
    - EXTERNAL DEPENDENCIES: Requires 'libreoffice' and 'img2pdf' to be available in the 
      system PATH for document and image conversion. As well as config.json.
    - AUTHENTICATION: Uses X.509 Certificate thumbprint/path. Ensure the service principal 
      has 'Mail.ReadWrite' permissions in Azure AD.
      Assumes permissions and mapping of the SMB folder paths.

    .DESCRIPTION
    Automates the ingestion of AP invoices from a Microsoft 365 mailbox. The script:
    Upon starting the Script will immediately verify the config file exists and that the values are acceptable for the scripts use.
    1. Authenticates via MS Graph using a certificate.
    2. Filters inbox messages for attachments and specific 'Test Mode' senders.
    3. Employs a 'Waterfall' matching logic to map senders to vendors via CSV.
    4. Sanitizes filenames, handles naming collisions, and converts non-PDF attachments 
       (Office docs, CSVs, Images) into standardized PDF format using LibreOffice and img2pdf.
    5. Routes files to a tiered SMB directory structure based on vendor naming.
    6. Added logging functionality: Circular (Bulk end-of-run trim), Verbose, Error, and Runtime history tracking.
       Includes Dual-pipe abstraction, correlation IDs, and global error trapping.
    7. Added Mailbox Routing: Marks emails as read and moves them from the Inbox to fuzzy-matched top-level mailbox folders (e.g., root "A - Invoices").

    .VERSION
    1.23

    .NOTES
    - Explicit regex sanitization is used for vendor names to prevent directory traversal.
    - Collision detection (suffixing files with (01), (02)) prevents overwriting 
      existing invoices on the SMB share.
    - Added additional robust try catch blocks to make sure that each loop is succesful, or gets skipped, and also to ensure that our configs load in succesfully and Graph connects succesfully
#>

# --- ENABLE LINUX GRAPHICS FOR EXCEL AUTOFIT ---
$env:DOTNET_System_Drawing_EnableUnixSupport = "true"

# --- GLOBAL ERROR TRAP & RUN IDENTITY ---
$ErrorActionPreference = "Stop" # Prevents silent failures
# Generates a unique ID to more easily group and filter log entries for each individual execution.
$global:RunID = "RUN-$(Get-Date -Format 'yyyyMMddHHmm')-$((Get-Random -Maximum 9999).ToString('0000'))"

# --- 0. CONCURRENCY LOCK ---
$lockFile = "/opt/ap-automation/staging/ap_automation.lock"
if (Test-Path -LiteralPath $lockFile) {
    $lockAge = (Get-Date) - (Get-Item $lockFile).LastWriteTime
    if ($lockAge.TotalMinutes -lt 15) {
        Write-Host "WARNING: Another instance is currently running. Exiting to prevent file collisions." -ForegroundColor Yellow
        exit
    } else {
        # Clear stale lock from a previous crash
        Remove-Item -LiteralPath $lockFile -Force
    }
}
New-Item -Path $lockFile -ItemType File -Force | Out-Null

# --- HELPER FUNCTIONS ---
function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("ERROR", "WARN", "INFO", "SUCCESS", "DEBUG")]
        [string]$Level = "INFO",
        [string]$LogType = "Error", # "Error" or "Runtime"
        [string]$MsgID = "SYS",
        [System.ConsoleColor]$Color = "Cyan"
    )

    # --- 1. CONSOLE OUTPUT (The Dual Pipe) ---
    if ($LogType -ne "Runtime") {
        Write-Host $Message -ForegroundColor $Color
    }

    # --- 2. CONFIGURATION GATEKEEPERS ---
    if ($LogType -eq "Runtime" -and $config.Logging.Runtime -ne $true) {
        return 
    }
    if ($LogType -eq "Error" -and $Level -notin @("ERROR", "WARN") -and $config.Logging.Verbose -ne $true) {
        return 
    }

    # --- 3. APPEND TO FILE (Fast Path) ---
    $logPath = if ($LogType -eq "Runtime") { $config.Paths.Runtime_Log } else { $config.Paths.Error_Log }
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedMessage = "[$timestamp] [$global:RunID] [$MsgID] [$Level] - $Message"

    $formattedMessage | Out-File -FilePath $logPath -Append -Encoding UTF8
}
# Strips all special characters except hyphens and spaces to prevent file system write errors.
# Also Consolidates multiple spaces/hyphens into a single hyphen for clean, predictable naming.
function Format-InvoiceName {
    param (
        [string]$SupplierName,
        [string]$OriginalFileName
    )
    # Added # and other reserved URI chars to the stripping list
    $cleanSupplier = $SupplierName -replace '[^a-zA-Z0-9\s\-]', ''
    $cleanFileName = $OriginalFileName -replace '[#%&{}\\<>*?/$!''"@+|=]', '' # Clean the attachment name too!
    
    $cleanSupplier = ($cleanSupplier -replace '\s+', '-') -replace '\-+', '-'
    return "$($cleanSupplier.TrimEnd('-'))-$cleanFileName"
}

# --- EXCEL FORMATTING FUNCTION ---
# --- EXCEL FORMATTING FUNCTION ---
function Format-ExcelForPdf {
    param (
        [string]$FilePath,
        [string]$MsgID = "SYS"
    )
    
    try {
        Import-Module ImportExcel -ErrorAction Stop
        Write-Log "  -> [FORMATTING] Adjusting Excel print settings (Landscape, Fit-to-Width)..." -Level "INFO" -Color "Cyan" -MsgID $MsgID
        
        # 1. ATTEMPT TO OPEN
        $pkg = Open-ExcelPackage -Path $FilePath -ErrorAction Stop
        
        foreach ($ws in $pkg.Workbook.Worksheets) {
            # NO AUTOFIT: Pure XML layout injection only!
            $ws.PrinterSettings.Orientation = [OfficeOpenXml.eOrientation]::Landscape
            $ws.PrinterSettings.FitToPage = $true
            $ws.PrinterSettings.FitToWidth = 1
            $ws.PrinterSettings.FitToHeight = 0
            $ws.PrinterSettings.LeftMargin = 0.25
            $ws.PrinterSettings.RightMargin = 0.25
        }
        
        # 2. ATTEMPT TO SAVE
        Close-ExcelPackage -ExcelPackage $pkg -ErrorAction Stop
        
        Write-Log "  -> [GOOD] Excel file pre-formatted successfully." -Level "SUCCESS" -Color "Green" -MsgID $MsgID
        return $true
        
    } catch {
        $errMsg = $_.Exception.Message
        
        if ($errMsg -match "Save" -or $errMsg -match "Error saving file") {
            Write-Log "  -> [WARN] Excel formatting skipped (Linux dependency missing). Raw file is valid. Proceeding to LibreOffice." -Level "WARN" -Color "Yellow" -MsgID $MsgID
            return $true 
        } else {
            Write-Log "  -> [FAIL] Excel file is corrupted or unreadable: $errMsg" -Level "ERROR" -Color "Red" -MsgID $MsgID
            return $false 
        }
    }
}

# --- INITIALIZATION of CONFIG FILE---
$config = Get-Content "/opt/ap-automation/configs/config.json" | ConvertFrom-Json

# --- PRE-FLIGHT VALIDATION & SCHEMA INTEGRITY BLOCK ---
Write-Log "Initializing Pre-Flight Environment Validation..." -Level "INFO" -Color "Cyan" -MsgID "SYS"

# 1. Critical Key & Path Existence Verification
$criticalPaths = @{
    "Staging"        = $config.Paths.Staging
    "CSVPath"        = $config.Paths.CSVPath
    "SMBDestination" = $config.Paths.SMBDestination
}

foreach ($item in $criticalPaths.GetEnumerator()) {
    if ([string]::IsNullOrWhiteSpace($item.Value)) {
        Write-Log "[FAIL] Pre-Flight Error: $($item.Name) path is null or empty in config.json." -Level "ERROR" -Color "Red" -MsgID "SYS"
        throw "ValidationFailed: Missing $($item.Name) path in configuration."
    }
    # Skip existence check on SMB if simulating SMB, otherwise require it.
    if ($item.Name -eq "SMBDestination" -and $simulateSMB -eq $true) {
        continue
    }
    if (-not (Test-Path -LiteralPath $item.Value)) {
        Write-Log "[FAIL] Pre-Flight Error: $($item.Name) path does not exist on disk: $($item.Value)" -Level "ERROR" -Color "Red" -MsgID "SYS"
        throw "ValidationFailed: Path not found - $($item.Value)"
    }
}

# 2. Path Safety Boundary Checks (Anti-Directory Traversal)
$safeAppRoot = "/opt/ap-automation/"
$unsafeSystemRoots = @("/", "/etc", "/bin", "/var", "/root", "/usr")

$resolvedStaging = (Resolve-Path $config.Paths.Staging -ErrorAction Stop).Path
if (-not $resolvedStaging.StartsWith($safeAppRoot)) {
    Write-Log "[FAIL] Security Exception: Staging path ($resolvedStaging) is outside the safe application boundary ($safeAppRoot)." -Level "ERROR" -Color "Red" -MsgID "SYS"
    throw "SecurityException: Staging boundary violation."
}

if ($config.Paths.SMBDestination -in $unsafeSystemRoots) {
    Write-Log "[FAIL] Security Exception: SMBDestination points to a protected system root." -Level "ERROR" -Color "Red" -MsgID "SYS"
    throw "SecurityException: SMB Destination root violation."
}

# 3. CSV Schema Verification
try {
    $mapping = Import-Csv $config.Paths.CSVPath -ErrorAction Stop
} catch {
    Write-Log "[FAIL] Pre-Flight Error: Could not read the CSV file at $($config.Paths.CSVPath). Is it locked?" -Level "ERROR" -Color "Red" -MsgID "SYS"
    throw "ValidationFailed: CSV Read Error."
}

if ($null -eq $mapping -or $mapping.Count -eq 0) {
    Write-Log "[FAIL] Pre-Flight Error: The CSV mapping file is empty or failed to load." -Level "ERROR" -Color "Red" -MsgID "SYS"
    throw "ValidationFailed: Empty Mapping Document."
}

$expectedHeaders = @(
    "Email - AP Vendor List", 
    "Email - Vendor Match", 
    "Domain", 
    "Supplier Name"
)

# Extract actual headers from the first row of the PSObject
$actualHeaders = $mapping[0].psobject.properties.name

foreach ($header in $expectedHeaders) {
    if ($header -notin $actualHeaders) {
        Write-Log "[FAIL] Schema Error: Required CSV column '$header' is missing from the mapping file." -Level "ERROR" -Color "Red" -MsgID "SYS"
        throw "ValidationFailed: Invalid CSV Schema."
    }
}

Write-Log "[GOOD] Pre-Flight Validation Passed. Configuration and Schema are sound." -Level "SUCCESS" -Color "Green" -MsgID "SYS"
# ------------------------------------------------------

# --- CONFIG VARIABLES ---
$certPath = $config.AzureAd.CertPath
$keyPath  = $config.AzureAd.KeyPath
$clientId = $config.AzureAd.ClientId
$tenantId = $config.AzureAd.TenantId
$targetMailbox = $config.Email.TargetMailbox
$genericDomains = $config.Email.GenericDomains
$internalDomains = $config.Email.InternalDomains
$allowedDocs   = $config.Email.AllowedDocs
$allowedImages = $config.Email.AllowedImages
$batchSize = if ($config.Email.BatchSize) { [int]$config.Email.BatchSize } else { 20 }
#Merges the allowed docs list and the allowed images list
$allowedExtensions= $allowedDocs + $allowedImages
$minImageSize = if ($config.Email.MinImageSizeBytes) { $config.Email.MinImageSizeBytes } else { 30000 }
#Build the strings powershell needs for parsing extension types
$imageRegex = if ($allowedImages) { "(?i)\.(" + ($allowedImages.Replace('.','') -join '|') + ")$" } else { "(?i)\.(NONE)$" }
$docRegex   = if ($allowedDocs) { "(?i)\.(" + ($allowedDocs.Replace('.','') -join '|') + ")$" } else { "(?i)\.(NONE)$" }

$testFromEnabled = $config.Email.TestFromEnabled
$testFromAddress = $config.Email.TestFromAddress
$keyWordExceptions = $config.Email.KeyWordExceptions
$simulateSMB = $config.Paths.SimulateSMB
$simulateMove = $config.Email.SimulateMove

# --- LOG FOLDER VERIFICATION ---
if (-not (Test-Path -LiteralPath $config.Paths.LogFolder)) {
    New-Item -ItemType Directory -Force -LiteralPath $config.Paths.LogFolder | Out-Null
}

# ==========================================
# --- 1. DEDICATED CONNECTION BLOCK ---
# ==========================================
try {
    Write-Log "Connecting to Graph API..." -Level "INFO" -Color "Cyan"
    
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPemFile($certPath, $keyPath)
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert -NoWelcome

    Write-Log "[GOOD] Connected to Graph successfully. Fetching Inbox Messages..." -Level "SUCCESS" -Color "Cyan"
}
catch [System.Security.Cryptography.CryptographicException] {
    Write-Log "[FAIL] CERTIFICATE ERROR: The PEM file or Key ($certPath) is invalid or inaccessible." -Level "ERROR" -Color "Red"
    throw "FatalError: Certificate Cryptographic Exception." # Terminates script immediately
}
catch [Microsoft.Graph.PowerShell.Authentication.CmdletException] {
    Write-Log "[FAIL] AUTHENTICATION ERROR: Azure AD rejected the connection. Verify ClientID, TenantID, and App Permissions." -Level "ERROR" -Color "Red"
    throw "FatalError: Graph API Authentication Denied." # Terminates script immediately
}
catch {
    Write-Log "[FAIL] UNEXPECTED CONNECTION ERROR: $($_.Exception.Message)" -Level "ERROR" -Color "Red"
    throw "FatalError: Unhandled Connection Exception." # Terminates script immediately
}

try {
    if ($testFromEnabled -eq $true) {
        Write-Log " [INFORMATIONAL] Test mode is ENABLED. Filtering only for emails from: $testFromAddress" -Level "INFO" -Color "Magenta"
        $filterQuery = "from/emailAddress/address eq '$testFromAddress' and hasAttachments eq true"
    } else {
        $filterQuery = "isRead eq false and hasAttachments eq true"
    }
    
    $messages = Get-MgUserMailFolderMessage -UserId $targetMailbox -MailFolderId "Inbox" -All -Filter $filterQuery -Select "id,subject,from,receivedDateTime,hasAttachments"

    if ($messages.Count -eq 0) {
        Write-Log "[INFORMATIONAL] No new emails with attachments found." -Level "INFO" -Color "Yellow"
    } else {
        Write-Log "Found $($messages.Count) email(s) to process." -Level "INFO" -Color "Cyan"
        
        # Pre-cache Top-Level Mailbox Folders for routing later
        try {
            Write-Log "Caching Top-Level Mailbox folders for Routing..." -Level "INFO" -Color "DarkGray"
            $mailboxFolders = Get-MgUserMailFolder -UserId $targetMailbox -All
        } catch {
            Write-Log "Failed to cache mailbox folders. Mailbox Routing may fail: $($_.Exception.Message)" -Level "WARN" -Color "Yellow"
            $mailboxFolders = @()
        }
    }
    # --- SMART BATCH TRACKING ---
    $processedCount = 0
    # Loop through each message
    foreach ($msg in $messages) {
        # 1. Check if we have hit our processing limit for this run
        if ($processedCount -ge $batchSize) {
            Write-Log " [BATCH LIMIT] Reached maximum of $batchSize processed invoices. Pausing remaining queue until next run." -Level "INFO" -Color "Cyan"
            break # Exits the foreach loop completely
        }
        # Generate short Message Correlation ID
        $MsgID = if ($msg.Id.Length -ge 8) { $msg.Id.Substring($msg.Id.Length - 8) } else { "UNKNOWN" }
        # Assume success until a failure occurs
        $allAttachmentsSuccessful = $true
        
        Write-Log "`n----------------------------------------" -Level "INFO" -Color "White" -MsgID $MsgID
        
        $senderEmail = $msg.From.EmailAddress.Address.ToLower()
        $senderDomain = ($senderEmail -split '@')[-1]
        $senderDisplayName = $msg.From.EmailAddress.Name
        
        Write-Log "Processing Email: $($msg.Subject) | Sender: $senderEmail" -Level "INFO" -Color "White" -MsgID $MsgID
        
        # --- Keyword Skipping ---
        # BUSINESS LOGIC: Prevents processing non-invoice financial documents (e.g., 'Account Statement').
        # TECHNICAL LOGIC: Dynamically builds a regex string from the config. 
        # INTENTIONAL FAIL-SAFE: If $keyWordExceptions is empty, the regex (?i)() is generated.
        # This matches ALL subject lines and file names, causing the script to skip the entire inbox. This prevents accidental ingestion of non-validated data if the configuration is cleared.
        
        $escapedKeywords = $keyWordExceptions | ForEach-Object { [regex]::Escape($_) }
        $killRegex = "(?i)(" + ($escapedKeywords -join "|") + ")"
        
        $hasKillKeyword = $false
        $matchedWord = ""
        
        if ($msg.Subject -match $killRegex) { 
            $hasKillKeyword = $true 
            $matchedWord = $matches[0]
        }
        
        $attachments = Get-MgUserMessageAttachment -UserId $targetMailbox -MessageId $msg.Id -Select "id,name,contentType,size,isInline"
        
        if (-not $hasKillKeyword -and $attachments) {
            foreach ($att in $attachments) {
                if ($att.Name -match $killRegex) { 
                    $hasKillKeyword = $true
                    $matchedWord = $matches[0]
                    break 
                }
            }
        }

        if ($hasKillKeyword) {
            # Injecting the senderEmail and the exact matched word into the Warning string
            Write-Log " [SKIP] Email from '$senderEmail' contains the keyword '$matchedWord'. Leaving untouched." -Level "WARN" -Color "Yellow" -MsgID $MsgID
            Write-Log "From:$senderEmail - Subject:$($msg.Subject) - AttachmentCount:$($attachments.Count) - Untouched (Keyword: $matchedWord)" -LogType "Runtime" -MsgID $MsgID
            continue
        }

        # --- 1. HIERARCHICAL VENDOR MAPPING (WATERFALL LOGIC) ---
        # BUSINESS LOGIC: Evaluates the sender's identity against the CSV schema in descending order of strictness.
        # Tier 1: Exact Match on 'Email - AP Vendor List' (Highest Confidence)
        # Tier 2: Exact Match on 'Email - Vendor Match'
        # Tier 3: Fuzzy Regex Match on Sender Display Name
        # Tier 4: Corporate Domain Fallback (Lowest Confidence, intentionally ignores generic domains like @gmail.com)
        $supplierName = "Unknown"
        $fallbackSupplier = "Unknown"
        Write-Host ">>> DEBUG: The raw Display Name PowerShell sees is: '$senderDisplayName'" -ForegroundColor Magenta

        foreach ($row in $mapping) {
            # Wrapping the row calls in "$()" prevents the "null-valued expression" crash if a column is missing or a cell is empty
            $csvApVendorList = "$($row.'Email - AP Vendor List')".Trim().ToLower()
            $csvVendorMatch  = "$($row.'Email - Vendor Match')".Trim().ToLower()
            $csvDomain       = "$($row.'Domain')".Trim().ToLower()
            $csvSupplierName = "$($row.'Supplier Name')".Trim()

            if ($csvApVendorList -ne "" -and $senderEmail -eq $csvApVendorList) {
                $supplierName = $csvSupplierName
                Write-Log " [GOOD] Matched via AP Vendor List -> $supplierName" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                break
            }
            if ($csvVendorMatch -ne "" -and $senderEmail -eq $csvVendorMatch) {
                $supplierName = $csvSupplierName
                Write-Log " [GOOD] Matched via Vendor Match -> $supplierName" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                break
            }
            # TIER 2: Display Name "Deep Scan" (Alphabetical Boundary Match)
            # Wraps the CSV names in negative lookarounds. It ensures that the matched name 
            # does not have alphabetical characters touching it on either side, preventing "SMI" from matching "Smith".
            
            $regexSupplier = "(?<![a-zA-Z])" + [regex]::Escape($csvSupplierName) + "(?![a-zA-Z])"
            $regexAPList   = "(?<![a-zA-Z])" + [regex]::Escape($csvApVendorList) + "(?![a-zA-Z])"
            $regexVendor   = "(?<![a-zA-Z])" + [regex]::Escape($csvVendorMatch) + "(?![a-zA-Z])"

            if (($csvSupplierName -ne "" -and $senderDisplayName -match $regexSupplier) -or 
                ($csvApVendorList -ne "" -and $senderDisplayName -match $regexAPList) -or
                ($csvVendorMatch  -ne "" -and $senderDisplayName -match $regexVendor)) {
                
                $supplierName = $csvSupplierName
                Write-Log " [GOOD] Matched via Display Name Tag -> $supplierName" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                break
            }
            if ($fallbackSupplier -eq "Unknown" -and $csvDomain -ne "" -and $senderDomain -eq $csvDomain -and $senderDomain -notin $genericDomains -and $senderDomain -notin $internalDomains) {
                # Store the domain match as a fallback, but continue the loop to check for exact email matches
                $fallbackSupplier = $csvSupplierName
            }
        }

        # Final check: If no Tier 1 or Tier 2 match was found, use the Tier 3 Domain Fallback
        if ($supplierName -eq "Unknown" -and $fallbackSupplier -ne "Unknown") {
            $supplierName = $fallbackSupplier
            Write-Log " [GOOD] Matched via Corporate Domain (Fallback) -> $supplierName" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
        }

        if ($supplierName -eq "Unknown") {
            # Check if the reason it's unknown is because it's an internal domain
            if ($internalDomains -and $senderDomain -in $internalDomains) {
                Write-Log " [SKIP] Internal Sender Detected: '$senderEmail'. Leaving email untouched in Inbox." -Level "WARN" -Color "Magenta" -MsgID $MsgID
                Write-Log "From:$senderEmail - Subject:$($msg.Subject) - AttachmentCount:$($attachments.Count) - Untouched (Internal Sender)" -LogType "Runtime" -MsgID $MsgID
                continue 
            } else {
                # It's an actual unknown external sender
                Write-Log " [WARNING] No Match Found for '$senderEmail'. Leaving email untouched in Inbox." -Level "WARN" -Color "Yellow" -MsgID $MsgID
                Write-Log "From:$senderEmail - Subject:$($msg.Subject) - AttachmentCount:$($attachments.Count) - Untouched (Unknown Vendor)" -LogType "Runtime" -MsgID $MsgID
                continue 
            }
        } else {
            # --- NEW SYNCHRONIZED FOLDER LOGIC (v1.11) ---
            $firstLetter = $supplierName.Substring(0,1).ToUpper()
            
            if ($firstLetter -match "[A-Z]") {
                $targetSubFolder = $firstLetter
                $folderNameForLog = "$targetSubFolder - Invoices"
                $expectedPattern = "(?i)^$targetSubFolder\s*-\s*Invoices$"
            } else {
                $targetSubFolder = "123 - Folder"
                $folderNameForLog = "123 - Folder"
                $expectedPattern = "(?i)^123\s*-\s*Folder$"
            }

            # 1. Resolve SMB Path (Discover actual folder name on share like 'A - Invoices' or '123 - Folder')
            try {
                $matchedSmbDir = Get-ChildItem -LiteralPath $config.Paths.SMBDestination -Directory -ErrorAction SilentlyContinue | 
                                 Where-Object { $_.Name -match $expectedPattern } | Select-Object -First 1
                
                $finalSmbPath = if ($matchedSmbDir) { $matchedSmbDir.FullName } else { Join-Path $config.Paths.SMBDestination $targetSubFolder }
            } catch {
                $finalSmbPath = Join-Path $config.Paths.SMBDestination $targetSubFolder
            }

            # 2. Resolve Mailbox Folder (Identify matching root-level Outlook folder)
            $targetMailFolder = $mailboxFolders | Where-Object { $_.DisplayName -match $expectedPattern } | Select-Object -First 1
            # ---------------------------------------------
        }

        # --- 2. ATTACHMENT INSPECTION & RENAMING BLOCK ---
        $validAttachments = @()
        $filesToMove = @()
        $processedFileNames = @()
        $dateStamp = Get-Date -Format "yyyyMMdd"

        if ($attachments) {
            foreach ($att in $attachments) {
                
                # 1. Figure out the extension FIRST so the logic below can use it
                $ext = [System.IO.Path]::GetExtension($att.Name).ToLower()

                # 2. THE INLINE GATE
                if ($att.IsInline -eq $true) {
                    # If it is an image AND (it's tiny OR named like a signature), skip it
                    if ($ext -match $imageRegex -and ($att.Size -lt $minImageSize -or $att.Name -match "(?i)(outlook|hoguebanner|facebook|logo|youtube|^image\d*\.png$)")) {
                        Write-Log "  -> [SKIP] Ignored Inline Signature Image: $($att.Name)" -Level "INFO" -Color "DarkGray" -MsgID $MsgID
                        continue
                    }
                }

                # 3. Standard allowed extensions check
                if ($ext -in $allowedExtensions) {
                    
                    # Catch tiny attached images that somehow weren't flagged as inline
                    if ($ext -match $imageRegex -and $att.Size -lt $minImageSize) {
                        Write-Log "  -> [SKIP] Ignored Tiny Attached Image: $($att.Name)" -Level "INFO" -Color "DarkGray" -MsgID $MsgID
                        continue
                    }

                    Write-Log "  -> [KEEP] Found Invoice: $($att.Name)" -Level "INFO" -Color "Green" -MsgID $MsgID
                    
                    # Ensure we use the function to sanitize both Supplier and Original File Name
                    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($att.Name)
                    $baseName = Format-InvoiceName -SupplierName $supplierName -OriginalFileName $nameWithoutExt
                    $baseName = "${baseName}"
                    #_${dateStamp}"
                    
                    # --- FIXED COLLISION LOGIC ---
                    # 1. Determine staging extension BEFORE the collision check
                    $stagingExt = if ($ext -match $docRegex -or $ext -match $imageRegex) { $ext } else { ".pdf" }
                    
                    # Determine the FINAL destination extension (.xls bypass keeps .xls, everything else gets .pdf)
                    $destExt = if ($ext -match "\.xls$") { $ext } else { ".pdf" }
                    
                    $counter = 0
                    
                    # Set initial target names
                    $finalPdfName = "$baseName$destExt"
                    $currentStagingName = "$baseName$stagingExt"
                    
                    # 2. Check BOTH the SMB target and the local staging folder
                    while ((Test-Path -LiteralPath (Join-Path $finalSmbPath $finalPdfName)) -or (Test-Path -LiteralPath (Join-Path $config.Paths.Staging $currentStagingName))) {
                        $counter++
                        $paddedCounter = "{0:D2}" -f $counter
                        
                        # Apply the dynamic extensions so bypassed files keep their format!
                        $finalPdfName = "$baseName($paddedCounter)$destExt"
                        $currentStagingName = "$baseName($paddedCounter)$stagingExt"
                    }
                    # -----------------------------
                    
                    $newFileName = $currentStagingName
                    Write-Log "  -> [RENAMING] New Invoice Name: $newFileName" -Level "INFO" -Color "Magenta" -MsgID $MsgID
                    
                    $stagingPath = Join-Path $config.Paths.Staging $newFileName
                    try {
                        Write-Log "  -> [DOWNLOADING] Fetching file data to Staging..." -Level "INFO" -Color "Cyan" -MsgID $MsgID
                        $uri = "https://graph.microsoft.com/v1.0/users/$targetMailbox/messages/$($msg.Id)/attachments/$($att.Id)"
                        # Added -ErrorAction Stop to guarantee Graph errors trigger the catch block
                        $rawAttachment = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                        [System.IO.File]::WriteAllBytes($stagingPath, [System.Convert]::FromBase64String($rawAttachment.contentBytes))
                        
                        # Short pause to ensure the OS has released the file handle before Excel or LibreOffice touches it
                        Start-Sleep -Milliseconds 500
                        
                        if ($ext -match $docRegex -and $ext -ne ".pdf") {
                            
                            
                            # --- 1. THE .XLS RAW BYPASS ---
                            if ($ext -match "\.xls$") {
                                Write-Log "  -> [BYPASS] Legacy .xls detected. Skipping PDF conversion and keeping raw file." -Level "INFO" -Color "Cyan" -MsgID $MsgID
                                
                                if (Test-Path -LiteralPath $stagingPath) {
                                    # Add the raw .xls file directly to the move queue
                                    $filesToMove += $stagingPath
                                    # Log the actual .xls name, not the .pdf name
                                    $processedFileNames += $currentStagingName
                                } else {
                                    Write-Log "  -> [FAIL] File not found in staging: $stagingPath" -Level "ERROR" -Color "Red" -MsgID $MsgID
                                    $allAttachmentsSuccessful = $false
                                }

<# ======================= OLD UPGRADE LOGIC COMMENTED OUT =======================
                            Write-Log "  -> [CONVERTING] Starting LibreOffice upgrade: $stagingPath" -Level "INFO" -Color "Cyan" -MsgID $MsgID
                            Write-Log "  -> [CONVERTING] Upgrading legacy .xls to .xlsx to apply Landscape formatting..." -Level "INFO" -Color "Cyan" -MsgID $MsgID

                            # --- DEBUG BLOCK START ---
                            Write-Log "  -> [DEBUG] XLS source exists BEFORE conversion: $(Test-Path $stagingPath)" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                            Write-Log "  -> [DEBUG] Target staging directory: $($config.Paths.Staging)" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                            # --- DEBUG BLOCK END ---
                            
                            $xlsProcess = Start-Process -FilePath "libreoffice" `
                                -ArgumentList "--headless", "--convert-to", "xlsx", "`"$stagingPath`"", "--outdir", "`"$($config.Paths.Staging)`"" `
                                -PassThru

                            # --- DEBUG BLOCK START ---
                            Write-Log "  -> [DEBUG] LibreOffice process started. PID: $($xlsProcess.Id)" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                            # --- DEBUG BLOCK END ---

                            if ($xlsProcess.WaitForExit(60000) -and $xlsProcess.ExitCode -eq 0) {
                                
                                # --- DEBUG BLOCK START ---
                                Write-Log "  -> [DEBUG] LibreOffice exited successfully. ExitCode: $($xlsProcess.ExitCode)" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                                # --- DEBUG BLOCK END ---
                                
                                $upgradedPath = [System.IO.Path]::ChangeExtension($stagingPath, ".xlsx")

                                # --- DEBUG BLOCK START ---
                                Write-Log "  -> [DEBUG] Expected upgraded file path: $upgradedPath" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                                Write-Log "  -> [DEBUG] XLSX exists AFTER conversion: $(Test-Path $upgradedPath)" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                                # --- DEBUG BLOCK END ---

                                if (Test-Path -LiteralPath $upgradedPath) {
                                    Write-Log "  -> [GOOD] XLS upgrade successful: $upgradedPath" -Level "SUCCESS" -Color "Green" -MsgID $MsgID

                                    # --- DEBUG BLOCK START ---
                                    Write-Log "  -> [DEBUG] Removing original XLS: $stagingPath" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                                    # --- DEBUG BLOCK END ---

                                    # Delete the old binary .xls
                                    Remove-Item -LiteralPath $stagingPath -Force

                                    # Trick the script into treating it like an .xlsx from here on out
                                    $stagingPath = $upgradedPath
                                    $ext = ".xlsx"

                                    # --- DEBUG BLOCK START ---
                                    Write-Log "  -> [DEBUG] Updated stagingPath to: $stagingPath" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                                    Write-Log "  -> [DEBUG] Updated extension to: $ext" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                                    # --- DEBUG BLOCK END ---
                                }
                                else {
                                    Write-Log "  -> [FAIL] XLS upgrade reported success but file NOT FOUND at expected path." -Level "ERROR" -Color "Red" -MsgID $MsgID
                                }
                            } else {
                                # --- DEBUG BLOCK START ---
                                Write-Log "  -> [DEBUG] LibreOffice exit code: $($xlsProcess.ExitCode)" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                                Write-Log "  -> [DEBUG] WaitForExit result: $($xlsProcess.HasExited)" -Level "DEBUG" -Color "Gray" -MsgID $MsgID
                                # --- DEBUG BLOCK END ---

                                Write-Log "  -> [WARN] Failed to upgrade .xls. LibreOffice will attempt to process raw file." -Level "WARN" -Color "Yellow" -MsgID $MsgID
                            }
============================================================================== #>

                                # This command skips the rest of the conversion logic below and jumps to the next attachment!
                                continue 
                            }
                            # ----------------------------------

                            # --- 2. THE FORMATTING GATE ---
                            # Legacy .xls files bypass this completely. This strictly formats modern .xlsx files.
                            if ($ext -match "\.xlsx$") {
                                # If it returns $false, the file is genuinely corrupted. Kill it.
                                if (-not (Format-ExcelForPdf -FilePath $stagingPath -MsgID $MsgID)) {
                                    Write-Log "   -> [FAIL] Skipping conversion due to fatal Excel error. Deleting corrupted file." -Level "ERROR" -Color "Red" -MsgID $MsgID
                                    Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue
                                    $allAttachmentsSuccessful = $false
                                    continue 
                                }
                            }

                            # --- 3. FINAL PDF CONVERSION ---
                            Write-Log "  -> [CONVERTING] Running LibreOffice Headless on $ext (Max 60s timeout)..." -Level "INFO" -Color "Cyan" -MsgID $MsgID
                            
                            # --- DYNAMIC EXPORT FILTER ---
                            # Protects Word Docs by ONLY applying the strict Calc filter to Excel files
                            $exportFormat = if ($ext -match "\.xlsx?$") { "pdf:calc_pdf_Export" } else { "pdf" }
                            
                            # Notice we swapped "pdf" for "$exportFormat" in the ArgumentList below
                            $process = Start-Process -FilePath "libreoffice" -ArgumentList "--headless", "--convert-to", "$exportFormat", "`"$stagingPath`"", "--outdir", "`"$($config.Paths.Staging)`"" -PassThru
                            
                            if ($process.WaitForExit(60000)) {
                                $pdfPath = [System.IO.Path]::ChangeExtension($stagingPath, ".pdf")
                                if ($process.ExitCode -eq 0 -and (Test-Path -LiteralPath $pdfPath)) {
                                    Write-Log "  -> [GOOD] Document successfully converted to PDF!" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                                    Remove-Item -LiteralPath $stagingPath -Force
                                    $filesToMove += $pdfPath
                                    $processedFileNames += $finalPdfName
                                } else {
                                    Write-Log "  -> [FAIL] LibreOffice failed. ExitCode: $($process.ExitCode). Output file not found." -Level "ERROR" -Color "Red" -MsgID $MsgID
                                    $allAttachmentsSuccessful = $false
                                }
                            } else {
                                Stop-Process -Id $process.Id -Force
                                Write-Log "  -> [FAIL] TIMEOUT: LibreOffice hung for over 60 seconds and was terminated." -Level "ERROR" -Color "Red" -MsgID $MsgID
                                $allAttachmentsSuccessful = $false
                            }
                        }
                        elseif ($ext -match $imageRegex -and $ext -ne ".pdf") {
                            Write-Log "  -> [CONVERTING] Running img2pdf on $ext (Max 60s timeout)..." -Level "INFO" -Color "Cyan" -MsgID $MsgID
                            
                            # 1. Explicitly define the target PDF path
                            $pdfPath = [System.IO.Path]::ChangeExtension($stagingPath, ".pdf")
                            
                            # 2. Execute conversion
                            $process = Start-Process -FilePath "img2pdf" -ArgumentList "`"$stagingPath`"", "-o", "`"$pdfPath`"" -PassThru
                            
                            if ($process.WaitForExit(60000)) {
                                # 3. PHYSICAL REALITY CHECK: Check ExitCode AND verify the file exists on the disk
                                if ($process.ExitCode -eq 0 -and (Test-Path -LiteralPath $pdfPath)) {
                                    Write-Log "  -> [GOOD] Image successfully converted to PDF!" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                                    
                                    # Cleanup: Remove the original image (e.g., the .jpg) now that the .pdf exists
                                    Remove-Item -LiteralPath $stagingPath -Force
                                    
                                    # Add ONLY the verified PDF path to the move list
                                    $filesToMove += $pdfPath
                                    $processedFileNames += $finalPdfName
                                } else {
                                    # This catches cases where img2pdf "finishes" but fails to write the file
                                    Write-Log "  -> [FAIL] img2pdf failed. ExitCode: $($process.ExitCode). Output file not found or inaccessible." -Level "ERROR" -Color "Red" -MsgID $MsgID
                                    $allAttachmentsSuccessful = $false
                                }
                            } else {
                                # Timeout logic
                                Stop-Process -Id $process.Id -Force
                                Write-Log "  -> [FAIL] TIMEOUT: img2pdf hung for over 60 seconds and was terminated." -Level "ERROR" -Color "Red" -MsgID $MsgID
                                $allAttachmentsSuccessful = $false
                            }
                        }
                        elseif ($ext -eq ".pdf") {
                            if (Test-Path -LiteralPath $stagingPath) {
                                $filesToMove += $stagingPath
                                $processedFileNames += $finalPdfName
                            } else {
                                Write-Log "  -> [FAIL] PDF download failed. File not found in staging: $stagingPath" -Level "ERROR" -Color "Red" -MsgID $MsgID
                                $allAttachmentsSuccessful = $false
                            }
                        }
                        $validAttachments += $att
                    } 
                    catch [System.Net.WebException], [Microsoft.Graph.PowerShell.Models.MicrosoftGraphODataErrorsOdataError] {
                        # Handles Network Timeouts, API Throttling, or Microsoft Graph Outages
                        Write-Log "  -> [FAIL] GRAPH API ERROR: Failed to fetch attachment '$($att.Name)'. Possible timeout or throttling." -Level "ERROR" -Color "Red" -MsgID $MsgID
                        $allAttachmentsSuccessful = $false
                    }
                    catch [System.IO.IOException] {
                        # Handles Local Disk Full, File Locked by another process, or Invalid File Name formats
                        Write-Log "  -> [FAIL] DISK I/O ERROR: Cannot write to $stagingPath. Check disk space or folder locks." -Level "ERROR" -Color "Red" -MsgID $MsgID
                        $allAttachmentsSuccessful = $false
                    }
                    catch [System.UnauthorizedAccessException] {
                        # Handles catastrophic permission loss on the Linux host
                        Write-Log "  -> [FAIL] PERMISSION ERROR: The script lacks write access to the staging directory." -Level "ERROR" -Color "Red" -MsgID $MsgID
                        throw "FatalError: Staging directory permissions invalid." # Terminate, no point in continuing
                    }
                    catch {
                        # The Fallback for anything truly unexpected
                        $errInfo = if ($config.Logging.Verbose) { "$($_.Exception.Message) (Line: $($_.InvocationInfo.ScriptLineNumber))" } else { $_.Exception.Message }
                        Write-Log "  -> [FAIL] Attachment Process Failed: $errInfo" -Level "ERROR" -Color "Red" -MsgID $MsgID
                        $allAttachmentsSuccessful = $false
                    } 
            }
            else {
                        Write-Log "  -> [SKIP] Unsupported file type '$ext' for attachment '$($att.Name)'. Skipping." -Level "WARN" -Color "Yellow" -MsgID $MsgID}
                }
        }

        if ($validAttachments.Count -eq 0) {
            continue 
        }

        # --- 3. SMB FOLDER ROUTING LOGIC ---
        Write-Log " Routing $($validAttachments.Count) file(s) to SMB Folder: $finalSmbPath" -Level "INFO" -Color "Cyan" -MsgID $MsgID
        
        # Short pause to ensure file handles are released by external binaries before moving
        Start-Sleep -Seconds 1 
        
        if ($simulateSMB -eq $true) {
            if (Test-Path -LiteralPath $finalSmbPath) {
                Write-Log "   [GOOD] Folder Exists on SMB Share." -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                Write-Log "From:$senderEmail - Subject:$($msg.Subject) - AttachmentCount:$($attachments.Count) - Processed (Simulation) - Attachments renamed to `"$($processedFileNames -join '", "')`" placed in folder `"$finalSmbPath`"" -LogType "Runtime" -MsgID $MsgID
            } else {
                Write-Log "   [FAIL] SMB Target Folder DOES NOT EXIST ($finalSmbPath)" -Level "ERROR" -Color "Red" -MsgID $MsgID
            }
        } else {
            if (Test-Path -LiteralPath $finalSmbPath) {
                foreach ($file in $filesToMove) {
                    try { 
                        if (Test-Path -LiteralPath $file) {
                            Move-Item -LiteralPath $file -Destination $finalSmbPath -Force 
                            Write-Log "   -> [MOVED] Successfully moved to SMB: $(Split-Path $file -Leaf)" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                        } else {
                            Write-Log "   -> [FAIL] Move Skipped: File disappeared from staging before move: $(Split-Path $file -Leaf)" -Level "ERROR" -Color "Red" -MsgID $MsgID
                        }
                    } catch { 
                        Write-Log "   -> [FAIL] Move Failed: $($_.Exception.Message)" -Level "ERROR" -Color "Red" -MsgID $MsgID
                    }
                }
                Write-Log "From:$senderEmail - Subject:$($msg.Subject) - TotalAttachmentCount:$($attachments.Count) - Processed:$($validAttachments.Count) - Attachments renamed to `"$($processedFileNames -join '", "')`" placed in folder `"$finalSmbPath`"" -LogType "Runtime" -MsgID $MsgID
            } else {
                Write-Log "   [FAIL] Target SMB Folder DOES NOT EXIST ($finalSmbPath). Files left in Staging." -Level "ERROR" -Color "Red" -MsgID $MsgID
            }
        }

        # --- 4. MAILBOX ROUTING LOGIC (Read & Move) ---
        if ($allAttachmentsSuccessful -eq $false) {
            Write-Log "   [HOLD] One or more attachments failed processing. Leaving email unread in Inbox for manual review." -Level "WARN" -Color "Yellow" -MsgID $MsgID
        }
        elseif ($simulateMove -eq $true) {
            if ($null -ne $targetMailFolder) {
                Write-Log "   [SIMULATION] Target mailbox folder found! Would move to: '$($targetMailFolder.DisplayName)'." -Level "INFO" -Color "Magenta" -MsgID $MsgID
            } else {
                Write-Log "   [SIMULATION WARNING] Target folder '$folderNameForLog' NOT FOUND at Root." -Level "WARN" -Color "Yellow" -MsgID $MsgID
            }
            Write-Log "   [SIMULATION] Email left unread and untouched in Inbox." -Level "INFO" -Color "DarkGray" -MsgID $MsgID
        } else {
            if ($null -ne $targetMailFolder) {
                try {
                    Update-MgUserMessage -UserId $targetMailbox -MessageId $msg.Id -IsRead -ErrorAction Stop | Out-Null
                    Move-MgUserMessage -UserId $targetMailbox -MessageId $msg.Id -DestinationId $targetMailFolder.Id -ErrorAction Stop | Out-Null
                    Write-Log "   [MOVED] Email marked as read and moved to mailbox folder: $($targetMailFolder.DisplayName)" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                    $processedCount++
                } catch {
                    Write-Log "   [FAIL] Failed to update/move email in mailbox: $($_.Exception.Message)" -Level "ERROR" -Color "Red" -MsgID $MsgID
                }
            } else {
                Write-Log "   [WARN] Target mailbox folder matching '$folderNameForLog' not found at Root. Email left untouched." -Level "WARN" -Color "Yellow" -MsgID $MsgID
            }
        }
    }
}
catch {
    $errInfo = if ($config.Logging.Verbose) { "$($_.Exception.Message) (Line: $($_.InvocationInfo.ScriptLineNumber))" } else { $_.Exception.Message }
    Write-Log "[FAIL] FATAL RUNTIME ERROR: $errInfo" -Level "ERROR" -Color "Red" -MsgID "SYS"
    
    # Force the script to exit with an error state so the scheduler (e.g., cron) registers a failure.
    throw "FatalError: Script execution aborted due to unhandled exception."
}
finally {
    if (Get-MgContext) { 
        Disconnect-MgGraph | Out-Null 
        Write-Log "`n[INFORMATIONAL] Disconnected from Graph API." -Level "INFO" -Color "DarkGray"
    }

    # --- FINAL CIRCULAR LOG CLEANUP (Bulk End-of-Run Trim) ---
    if ($config.Logging.Circular -eq $true) {
        $logPaths = @($config.Paths.Runtime_Log, $config.Paths.Error_Log)
        
        $sizeString = $config.Logging.Log_Size.ToUpper()
        $maxSizeBytes = 50MB
        if ($sizeString -match "(\d+)\s*MB") { $maxSizeBytes = [int]$matches[1] * 1MB }

        foreach ($logPath in $logPaths) {
            if (Test-Path $logPath) {
                $currentSize = (Get-Item $logPath).Length
                if ($currentSize -gt $maxSizeBytes) {
                    Write-Log "[CLEANUP] Log file $logPath exceeds limit ($([math]::Round($currentSize / 1MB, 2)) MB). Trimming..." -Level "INFO" -Color "Gray"
                    
                    # Read the log and determine how many entries (lines) to delete proportionally
                    $allLines = Get-Content $logPath
                    $totalLines = $allLines.Count
                    
                    # Proportional Calculation: If we are 20% over size, remove 20% of lines
                    $percentageOver = ($currentSize - $maxSizeBytes) / $currentSize
                    $linesToDelete = [math]::Ceiling($totalLines * $percentageOver)
                    
                    if ($linesToDelete -lt $totalLines) {
                        $allLines[$linesToDelete..($totalLines - 1)] | Set-Content $logPath
                        Write-Log "Log maintenance performed: Removed $linesToDelete oldest entries to maintain $sizeString limit." -Level "INFO" -Color "Gray"
                    }
                }
            }
        }
    }
    if (Test-Path $lockFile) {
        Remove-Item -LiteralPath $lockFile -Force -ErrorAction SilentlyContinue
    }
}
