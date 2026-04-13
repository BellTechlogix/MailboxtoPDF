<#
    MailboxAutomation.ps1
    Created By - Kristopher Roy
    Created On - 2026-03-20
    Revised On - 2026-04-13
    Revised On - 2026-04-13 (Descriptive Comments Added)
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
    1.8

    .NOTES
    - Explicit regex sanitization is used for vendor names to prevent directory traversal.
    - Collision detection (suffixing files with (01), (02)) prevents overwriting 
      existing invoices on the SMB share.
#>

# --- GLOBAL ERROR TRAP & RUN IDENTITY ---
$ErrorActionPreference = "Stop" # Prevents silent failures
$global:RunID = "RUN-$(Get-Date -Format 'yyyyMMddHHmm')-$((Get-Random -Maximum 9999).ToString('0000'))"

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

function Format-InvoiceName {
    param (
        [string]$SupplierName,
        [string]$OriginalFileName
    )
    $cleanSupplier = $SupplierName -replace '[^a-zA-Z0-9\s\-]', ''
    $cleanSupplier = ($cleanSupplier -replace '\s+', '-') -replace '\-+', '-'
    $cleanSupplier = $cleanSupplier.TrimEnd('-')
    return "$cleanSupplier-$OriginalFileName"
}

# --- NEW EXCEL FORMATTING FUNCTION ---
function Format-ExcelForPdf {
    param (
        [string]$FilePath,
        [string]$MsgID = "SYS"
    )
    
    try {
        # EXPLICITLY LOAD THE MODULE FIRST
        Import-Module ImportExcel -ErrorAction Stop
        
        Write-Log "  -> [FORMATTING] Adjusting Excel print settings (Landscape, Fit-to-Width)..." -Level "INFO" -Color "Cyan" -MsgID $MsgID
        
        # Open the Excel file in memory
        $pkg = Open-ExcelPackage -Path $FilePath
        
        # Loop through every sheet (tab) in the workbook
        foreach ($ws in $pkg.Workbook.Worksheets) {
            
            # Force the page orientation to Landscape
            $ws.PrinterSettings.Orientation = [OfficeOpenXml.eOrientation]::Landscape
            
            # Turn on the "Fit to Page" toggle
            $ws.PrinterSettings.FitToPage = $true
            
            # Constrain width to 1 page, but let height be unlimited (0)
            $ws.PrinterSettings.FitToWidth = 1
            $ws.PrinterSettings.FitToHeight = 0
            
            # Shrink the margins to give data more room
            $ws.PrinterSettings.LeftMargin = 0.25
            $ws.PrinterSettings.RightMargin = 0.25
        }
        
        # Save the changes back to the physical file and close it
        Close-ExcelPackage -ExcelPackage $pkg
        
        Write-Log "  -> [GOOD] Excel file pre-formatted successfully." -Level "SUCCESS" -Color "Green" -MsgID $MsgID
    } catch {
        # If it fails, we throw a warning but let LibreOffice try its best anyway
        Write-Log "  -> [WARN] Failed to pre-format Excel file: $($_.Exception.Message)" -Level "WARN" -Color "Yellow" -MsgID $MsgID
    }
}

# --- INITIALIZATION of CONFIG FILE---
$config = Get-Content "/opt/ap-automation/configs/config.json" | ConvertFrom-Json

# --- CONFIG VARIABLES ---
$certPath = $config.AzureAd.CertPath
$keyPath  = $config.AzureAd.KeyPath
$clientId = $config.AzureAd.ClientId
$tenantId = $config.AzureAd.TenantId
$targetMailbox = $config.Email.TargetMailbox
$mapping = Import-Csv $config.Paths.CSVPath
$genericDomains = $config.Email.GenericDomains
$internalDomains = $config.Email.InternalDomains
$allowedExtensions = $config.Email.AllowedExtensions
$minImageSize = if ($config.Email.MinImageSizeBytes) { $config.Email.MinImageSizeBytes } else { 30000 }

$testFromEnabled = $config.Email.TestFromEnabled
$testFromAddress = $config.Email.TestFromAddress
$keyWordExceptions = $config.Email.KeyWordExceptions
$simulateSMB = $config.Paths.SimulateSMB
$simulateMove = $config.Email.SimulateMove

# --- LOG FOLDER VERIFICATION ---
if (-not (Test-Path -Path $config.Paths.LogFolder)) {
    New-Item -ItemType Directory -Force -Path $config.Paths.LogFolder | Out-Null
}

try {
    Write-Log "Connecting to Graph API..." -Level "INFO" -Color "Cyan"
    
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPemFile($certPath, $keyPath)
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert -NoWelcome

    Write-Log "[GOOD] Connected to Graph successfully. Fetching Inbox Messages..." -Level "SUCCESS" -Color "Cyan"

    if ($testFromEnabled -eq $true) {
        Write-Log " [INFORMATIONAL] Test mode is ENABLED. Filtering only for emails from: $testFromAddress" -Level "INFO" -Color "Magenta"
        $filterQuery = "from/emailAddress/address eq '$testFromAddress' and hasAttachments eq true"
    } else {
        $filterQuery = "hasAttachments eq true"
    }
    
    $messages = Get-MgUserMailFolderMessage -UserId $targetMailbox -MailFolderId "Inbox" -all -Filter $filterQuery -Select "id,subject,from,receivedDateTime,hasAttachments"

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

    # Loop through each message
    foreach ($msg in $messages) {
        # Generate short Message Correlation ID
        $MsgID = if ($msg.Id.Length -ge 8) { $msg.Id.Substring($msg.Id.Length - 8) } else { "UNKNOWN" }
        
        Write-Log "`n----------------------------------------" -Level "INFO" -Color "White" -MsgID $MsgID
        
        $senderEmail = $msg.From.EmailAddress.Address.ToLower()
        $senderDomain = ($senderEmail -split '@')[-1]
        
        Write-Log "Processing Email: $($msg.Subject) | Sender: $senderEmail" -Level "INFO" -Color "White" -MsgID $MsgID
        
        # --- 0. STATEMENT & PAST DUE KILL SWITCH ---
        $escapedKeywords = $keyWordExceptions | ForEach-Object { [regex]::Escape($_) }
        $killRegex = "(?i)(" + ($escapedKeywords -join "|") + ")"
        
        $hasKillKeyword = $false
        if ($msg.Subject -match $killRegex) { $hasKillKeyword = $true }
        
        $attachments = Get-MgUserMessageAttachment -UserId $targetMailbox -MessageId $msg.Id -Select "id,name,contentType,size"
        
        if (-not $hasKillKeyword -and $attachments) {
            foreach ($att in $attachments) {
                if ($att.Name -match $killRegex) { $hasKillKeyword = $true; break }
            }
        }

        if ($hasKillKeyword) {
            Write-Log " [SKIP] Email contains Statement or Past Due keyword. Leaving untouched." -Level "WARN" -Color "Yellow" -MsgID $MsgID
            Write-Log "From:$senderEmail - Subject:$($msg.Subject) - AttachmentCount:$($attachments.Count) - Untouched (Past Due or Statement)" -LogType "Runtime" -MsgID $MsgID
            continue
        }

        # --- 1. THE WATERFALL MATCHING LOGIC ---
        $supplierName = "Unknown"

        foreach ($row in $mapping) {
            $csvApVendorList = $row.'Email - AP Vendor List'.Trim().ToLower()
            $csvVendorMatch  = $row.'Email - Vendor Match'.Trim().ToLower()
            $csvDomain       = $row.'Domain'.Trim().ToLower()

            if ($csvApVendorList -ne "" -and $senderEmail -eq $csvApVendorList) {
                $supplierName = $row.'Supplier Name'
                Write-Log " [GOOD] Matched via AP Vendor List -> $supplierName" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                break
            }
            if ($csvVendorMatch -ne "" -and $senderEmail -eq $csvVendorMatch) {
                $supplierName = $row.'Supplier Name'
                Write-Log " [GOOD] Matched via Vendor Match -> $supplierName" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                break
            }
            if ($csvDomain -ne "" -and $senderDomain -eq $csvDomain -and $senderDomain -notin $genericDomains -and $senderDomain -notin $internalDomains) {
                $supplierName = $row.'Supplier Name'
                Write-Log " [GOOD] Matched via Corporate Domain -> $supplierName" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                break
            }
        }

        if ($supplierName -eq "Unknown") {
            Write-Log " [WARNING] No Match Found. Leaving email untouched in Inbox." -Level "WARN" -Color "Yellow" -MsgID $MsgID
            Write-Log "From:$senderEmail - Subject:$($msg.Subject) - AttachmentCount:$($attachments.Count) - Untouched (Unknown Vendor)" -LogType "Runtime" -MsgID $MsgID
            continue 
        } else {
            # --- OLD SIMPLE FOLDER LOGIC (COMMENTED OUT) ---
            # $firstLetter = $supplierName.Substring(0,1).ToUpper()
            # $targetSubFolder = if ($firstLetter -match "[A-Z]") { $firstLetter } else { "#" }
            # $finalSmbPath = Join-Path $config.Paths.SMBDestination $targetSubFolder
            # -----------------------------------------------

            # --- NEW SYNCHRONIZED FOLDER LOGIC (v1.8) ---
            $firstLetter = $supplierName.Substring(0,1).ToUpper()
            $targetSubFolder = if ($firstLetter -match "[A-Z]") { $firstLetter } else { "#" }
            $expectedPattern = "(?i)^$targetSubFolder\s*-\s*Invoices$"

            # 1. Resolve SMB Path (Discover actual folder name on share like 'A - Invoices')
            try {
                $matchedSmbDir = Get-ChildItem -Path $config.Paths.SMBDestination -Directory -ErrorAction SilentlyContinue | 
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
                $ext = [System.IO.Path]::GetExtension($att.Name).ToLower()
                if ($ext -in $allowedExtensions) {
                    if ($ext -match "\.(jpg|jpeg|png|gif|bmp|tif|tiff)$" -and $att.Size -lt $minImageSize) {
                        Write-Log "  -> [SKIP] Ignored Tiny Image: $($att.Name)" -Level "INFO" -Color "DarkGray" -MsgID $MsgID
                        continue
                    }

                    Write-Log "  -> [KEEP] Found Invoice: $($att.Name)" -Level "INFO" -Color "Green" -MsgID $MsgID
                    
                    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($att.Name)
                    $cleanSupplier = ($supplierName -replace '[^a-zA-Z0-9\s\-]', '' -replace '\s+', '-').TrimEnd('-')
                    $baseName = "$cleanSupplier-$dateStamp-$nameWithoutExt"
                    
                    $counter = 0
                    $finalPdfName = "$baseName.pdf"
                    while (Test-Path (Join-Path $finalSmbPath $finalPdfName)) {
                        $counter++
                        $paddedCounter = "{0:D2}" -f $counter
                        $finalPdfName = "$baseName($paddedCounter).pdf"
                    }
                    
                    $newFileName = if ($counter -eq 0) { "$baseName$ext" } else { "$baseName($paddedCounter)$ext" }
                    Write-Log "  -> [RENAMING] New Invoice Name: $newFileName" -Level "INFO" -Color "Magenta" -MsgID $MsgID
                    
                    $stagingPath = Join-Path $config.Paths.Staging $newFileName
                    try {
                        Write-Log "  -> [DOWNLOADING] Fetching file data to Staging..." -Level "INFO" -Color "Cyan" -MsgID $MsgID
                        $uri = "https://graph.microsoft.com/v1.0/users/$targetMailbox/messages/$($msg.Id)/attachments/$($att.Id)"
                        $rawAttachment = Invoke-MgGraphRequest -Method GET -Uri $uri
                        [System.IO.File]::WriteAllBytes($stagingPath, [System.Convert]::FromBase64String($rawAttachment.contentBytes))
                        
                        if ($ext -match "\.(docx|doc|xlsx|xls|csv)$") {
                            
                            # --- INTERCEPT: FORMAT EXCEL FILES ---
                            if ($ext -match "\.(xlsx|xls)$") {
                                Format-ExcelForPdf -FilePath $stagingPath -MsgID $MsgID
                            }
                            # -------------------------------------

                            Write-Log "  -> [CONVERTING] Running LibreOffice Headless on $ext..." -Level "INFO" -Color "Cyan" -MsgID $MsgID
                            $process = Start-Process -FilePath "libreoffice" -ArgumentList "--headless", "--convert-to", "pdf", "`"$stagingPath`"", "--outdir", "`"$($config.Paths.Staging)`"" -Wait -PassThru
                            if ($process.ExitCode -eq 0) {
                                Write-Log "  -> [GOOD] Document successfully converted to PDF!" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                                Remove-Item -Path $stagingPath -Force
                                $filesToMove += [System.IO.Path]::ChangeExtension($stagingPath, ".pdf")
                                $processedFileNames += $finalPdfName
                            }
                        }
                        elseif ($ext -match "\.(jpg|jpeg)$") {
                            Write-Log "  -> [CONVERTING] Running img2pdf on $ext..." -Level "INFO" -Color "Cyan" -MsgID $MsgID
                            $convertedPdfPath = [System.IO.Path]::ChangeExtension($stagingPath, ".pdf")
                            $process = Start-Process -FilePath "img2pdf" -ArgumentList "`"$stagingPath`"", "-o", "`"$convertedPdfPath`"" -Wait -PassThru
                            if ($process.ExitCode -eq 0) {
                                Write-Log "  -> [GOOD] Image successfully converted to PDF!" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                                Remove-Item -Path $stagingPath -Force
                                $filesToMove += $convertedPdfPath
                                $processedFileNames += $finalPdfName
                            }
                        }
                        elseif ($ext -eq ".pdf") {
                            $filesToMove += $stagingPath
                            $processedFileNames += $finalPdfName
                        }
                        $validAttachments += $att
                    } catch {
                        $errInfo = if ($config.Logging.Verbose) { "$($_.Exception.Message) (Line: $($_.InvocationInfo.ScriptLineNumber))" } else { $_.Exception.Message }
                        Write-Log "  -> [FAIL] Attachment Process Failed: $errInfo" -Level "ERROR" -Color "Red" -MsgID $MsgID
                    }
                }
            }
        }

        if ($validAttachments.Count -eq 0) {
            continue 
        }

        # --- 3. SMB FOLDER ROUTING LOGIC ---
        Write-Log " Routing $($validAttachments.Count) file(s) to SMB Folder: $finalSmbPath" -Level "INFO" -Color "Cyan" -MsgID $MsgID
        
        if ($simulateSMB -eq $true) {
            if (Test-Path -Path $finalSmbPath) {
                Write-Log "   [GOOD] Folder Exists on SMB Share." -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                Write-Log "From:$senderEmail - Subject:$($msg.Subject) - AttachmentCount:$($attachments.Count) - Processed (Simulation) - Attachments renamed to `"$($processedFileNames -join '", "')`" placed in folder `"$finalSmbPath`"" -LogType "Runtime" -MsgID $MsgID
            } else {
                Write-Log "   [FAIL] SMB Target Folder DOES NOT EXIST ($finalSmbPath)" -Level "ERROR" -Color "Red" -MsgID $MsgID
            }
        } else {
            if (Test-Path -Path $finalSmbPath) {
                foreach ($file in $filesToMove) {
                    try { 
                        Move-Item -Path $file -Destination $finalSmbPath -Force 
                        Write-Log "   -> [MOVED] Successfully moved to SMB: $(Split-Path $file -Leaf)" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                    } catch { 
                        Write-Log "   -> [FAIL] Move Failed: $($_.Exception.Message)" -Level "ERROR" -Color "Red" -MsgID $MsgID
                    }
                }
                Write-Log "From:$senderEmail - Subject:$($msg.Subject) - AttachmentCount:$($attachments.Count) - Processed - Attachments renamed to `"$($processedFileNames -join '", "')`" placed in folder `"$finalSmbPath`"" -LogType "Runtime" -MsgID $MsgID
            } else {
                Write-Log "   [FAIL] Target SMB Folder DOES NOT EXIST ($finalSmbPath). Files left in Staging." -Level "ERROR" -Color "Red" -MsgID $MsgID
            }
        }

        # --- 4. MAILBOX ROUTING LOGIC (Read & Move) ---
        if ($simulateMove -eq $true) {
            # --- UPGRADED SIMULATION FEEDBACK ---
            if ($null -ne $targetMailFolder) {
                Write-Log "   [SIMULATION] Target mailbox folder found! Would move to: '$($targetMailFolder.DisplayName)'." -Level "INFO" -Color "Magenta" -MsgID $MsgID
            } else {
                Write-Log "   [SIMULATION WARNING] Target folder '$targetSubFolder - Invoices' NOT FOUND at Root." -Level "WARN" -Color "Yellow" -MsgID $MsgID
            }
            Write-Log "   [SIMULATION] Email left unread and untouched in Inbox." -Level "INFO" -Color "DarkGray" -MsgID $MsgID
            # ------------------------------------
        } else {
            # --- OLD FOLDER DISCOVERY (MOVED TO TOP IN v1.6) ---
            # Fuzzy match folder names like "A - Invoices", "A- Invoices", "A -Invoices", etc.
            # $expectedPattern = "(?i)^$targetSubFolder\s*-\s*Invoices$"
            # $targetMailFolder = $inboxSubfolders | Where-Object { $_.DisplayName -match $expectedPattern } | Select-Object -First 1
            # ---------------------------------------------------

            if ($null -ne $targetMailFolder) {
                try {
                    Update-MgUserMessage -UserId $targetMailbox -MessageId $msg.Id -IsRead -ErrorAction Stop | Out-Null
                    Move-MgUserMessage -UserId $targetMailbox -MessageId $msg.Id -DestinationId $targetMailFolder.Id -ErrorAction Stop | Out-Null
                    Write-Log "   [MOVED] Email marked as read and moved to mailbox folder: $($targetMailFolder.DisplayName)" -Level "SUCCESS" -Color "Green" -MsgID $MsgID
                } catch {
                    Write-Log "   [FAIL] Failed to update/move email in mailbox: $($_.Exception.Message)" -Level "ERROR" -Color "Red" -MsgID $MsgID
                }
            } else {
                Write-Log "   [WARN] Target mailbox folder matching '$targetSubFolder - Invoices' not found at Root. Email left untouched." -Level "WARN" -Color "Yellow" -MsgID $MsgID
            }
        }
    }
}
catch {
    $errInfo = if ($config.Logging.Verbose) { "$($_.Exception.Message) (Line: $($_.InvocationInfo.ScriptLineNumber))" } else { $_.Exception.Message }
    Write-Log "[FAIL] CRITICAL ERROR: $errInfo" -Level "ERROR" -Color "Red"
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
}
