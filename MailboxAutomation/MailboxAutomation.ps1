<#
    GraphMailCheck.ps1
    Created By - Kristopher Roy
    Created On - 2026-03-20
    Revised On - 2026-04-10
    Modules Required: Microsoft.Graph.Authentication, Microsoft.Graph.Users

    .Important
    - This script uses a hardcoded Linux-style path (/opt/ap-automation/). Ensure the execution environment is Linux or a compatible WSL/Container instance.
    - Requires a valid X.509 certificate and private key in PEM format for service principal authentication.
#>

# --- HELPER FUNCTIONS ---
function Format-InvoiceName {
    param (
        [string]$SupplierName,
        [string]$OriginalFileName
    )
    
    # 1. Strip out everything except letters, numbers, spaces, and hyphens
    $cleanSupplier = $SupplierName -replace '[^a-zA-Z0-9\s\-]', ''
    
    # 2. Replace all spaces with dashes, and collapse multiple dashes into a single dash
    $cleanSupplier = ($cleanSupplier -replace '\s+', '-') -replace '\-+', '-'
    
    # 3. Trim any trailing dashes just in case, then combine with the original file name
    $cleanSupplier = $cleanSupplier.TrimEnd('-')
    
    return "$cleanSupplier-$OriginalFileName"
}

# --- INITIALIZATION ---
$config = Get-Content "/opt/ap-automation/configs/config.json" | ConvertFrom-Json

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

# --- NEW CONFIG VARIABLES ---
$testFromEnabled = $config.Email.TestFromEnabled
$testFromAddress = $config.Email.TestFromAddress
$keyWordExceptions = $config.Email.KeyWordExceptions

try {
    Write-Host "Connecting to Graph..." -ForegroundColor Cyan
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::CreateFromPemFile($certPath, $keyPath)
    Connect-MgGraph -ClientId $clientId -TenantId $tenantId -Certificate $cert -NoWelcome

    Write-Host "[GOOD] Connected. Fetching Inbox Messages..." -ForegroundColor Cyan

    # --- INBOX FILTERING LOGIC ---
    # Deprecated: #$messages = Get-MgUserMailFolderMessage -UserId $targetMailbox -MailFolderId "Inbox" -Top 25 -Filter "hasAttachments eq true" -Select "id,subject,from,receivedDateTime,hasAttachments"
    # Deprecated: $messages = Get-MgUserMailFolderMessage -UserId $targetMailbox -MailFolderId "Inbox" -all -Filter "from/emailAddress/address eq 'kroy@belltechlogix.com' and hasAttachments eq true" -Select "id,subject,from,receivedDateTime,hasAttachments"
    
    if ($testFromEnabled -eq $true) {
        Write-Host " [INFORMATIONAL] Test mode is ENABLED. Filtering only for emails from: $testFromAddress" -ForegroundColor Magenta
        $filterQuery = "from/emailAddress/address eq '$testFromAddress' and hasAttachments eq true"
    } else {
        $filterQuery = "hasAttachments eq true"
    }
    
    $messages = Get-MgUserMailFolderMessage -UserId $targetMailbox -MailFolderId "Inbox" -all -Filter $filterQuery -Select "id,subject,from,receivedDateTime,hasAttachments"

    if ($messages.Count -eq 0) {
        Write-Host "[INFORMATIONAL] No new emails with attachments found." -ForegroundColor Yellow
    }

    # Loop through each message
    foreach ($msg in $messages) {
        Write-Host "`n----------------------------------------"
        
        $senderEmail = $msg.From.EmailAddress.Address.ToLower()
        $senderDomain = ($senderEmail -split '@')[-1]
        
        Write-Host "Processing Email: $($msg.Subject)" -ForegroundColor White
        Write-Host "Sender Email:   $senderEmail" -ForegroundColor Gray
        
        # --- 0. STATEMENT & PAST DUE KILL SWITCH ---
        # Deprecated: $killRegex = "(?i)(statement|statements|past due|past-due|pastdue)"
        
        # Dynamically build the Regex from the JSON array
        $escapedKeywords = $keyWordExceptions | ForEach-Object { [regex]::Escape($_) }
        $killRegex = "(?i)(" + ($escapedKeywords -join "|") + ")"
        
        $hasKillKeyword = $false
        
        # Check Subject Line
        if ($msg.Subject -match $killRegex) {
            $hasKillKeyword = $true
        }
        
        # Fetch attachments early so we can check their names for the kill words
        $attachments = Get-MgUserMessageAttachment -UserId $targetMailbox -MessageId $msg.Id -Select "id,name,contentType,size"
        
        if (-not $hasKillKeyword -and $attachments) {
            foreach ($att in $attachments) {
                if ($att.Name -match $killRegex) {
                    $hasKillKeyword = $true
                    break
                }
            }
        }

        if ($hasKillKeyword) {
            Write-Host " [SKIP] Email contains Statement or Past Due keyword. Leaving untouched." -ForegroundColor Yellow
            continue
        }

        # --- 1. THE WATERFALL MATCHING LOGIC ---
        $supplierName = "Unknown"

        foreach ($row in $mapping) {
            $csvApVendorList = $row.'Email - AP Vendor List'.Trim().ToLower()
            $csvVendorMatch  = $row.'Email - Vendor Match'.Trim().ToLower()
            $csvDomain       = $row.'Domain'.Trim().ToLower()

            # WATERFALL 1: Exact Email Match (Primary)
            if ($csvApVendorList -ne "" -and $senderEmail -eq $csvApVendorList) {
                $supplierName = $row.'Supplier Name'
                Write-Host " [GOOD] Matched via AP Vendor List -> $supplierName" -ForegroundColor Green
                break
            }
            
            # WATERFALL 2: Exact Email Match (Secondary)
            # Was: if ($senderEmail -eq $csvMatchEmail) { ...
            if ($csvVendorMatch -ne "" -and $senderEmail -eq $csvVendorMatch) {
                $supplierName = $row.'Supplier Name'
                Write-Host " [GOOD] Matched via Vendor Match -> $supplierName" -ForegroundColor Green
                break
            }
            
            # WATERFALL 3: Safe Domain Match
            if ($csvDomain -ne "" -and $senderDomain -eq $csvDomain -and $senderDomain -notin $genericDomains -and $senderDomain -notin $internalDomains) {
                $supplierName = $row.'Supplier Name'
                Write-Host " [GOOD] Matched via Corporate Domain -> $supplierName" -ForegroundColor Green
                break
            }
        }

        # --- NEW EARLY FOLDER ROUTING (Needed for Collision Detection) ---
        if ($supplierName -eq "Unknown") {
            Write-Host " [WARNING] No Match Found. Leaving email untouched in Inbox." -ForegroundColor Yellow
            continue 
        } else {
            $firstLetter = $supplierName.Substring(0,1).ToUpper()
            if ($firstLetter -match "[A-Z]") {
                $targetSubFolder = $firstLetter
            } else {
                $targetSubFolder = "#"
            }
        }
        $finalSmbPath = Join-Path $config.Paths.SMBDestination $targetSubFolder

        # --- 2. ATTACHMENT INSPECTION & RENAMING BLOCK ---
        # Deprecated: $attachments = Get-MgUserMessageAttachment -UserId $targetMailbox -MessageId $msg.Id -Select "id,name,contentType,size"
        # (Attachments are now fetched in Step 0)
        
        $validAttachments = @()
        $dateStamp = Get-Date -Format "yyyyMMdd"

        if ($attachments) {
            foreach ($att in $attachments) {
                $ext = [System.IO.Path]::GetExtension($att.Name).ToLower()
                
                if ($ext -in $allowedExtensions) {
                    
                    # --- IMAGE SIZE GATE ---
                    $isImage = $ext -match "\.(jpg|jpeg|png|gif|bmp|tif|tiff)$"
                    if ($isImage -and $att.Size -lt $minImageSize) {
                        Write-Host "  -> [SKIP] Ignored Tiny Image (Likely Signature, $($att.Size) bytes): $($att.Name)" -ForegroundColor DarkGray
                        continue
                    }
                    # -----------------------

                    Write-Host "  -> [KEEP] Found Invoice: $($att.Name)" -ForegroundColor Green
                    
                    # --- HYBRID FILE NAMING & COLLISION DETECTION ---
                    # Deprecated: $newFileName = Format-InvoiceName -SupplierName $supplierName -OriginalFileName $att.Name
                    
                    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($att.Name)
                    $cleanSupplier = $supplierName -replace '[^a-zA-Z0-9\s\-]', ''
                    $cleanSupplier = ($cleanSupplier -replace '\s+', '-') -replace '\-+', '-'
                    $cleanSupplier = $cleanSupplier.TrimEnd('-')
                    
                    # Base Name: Vendor Name - Date Stamp - OriginalFileName
                    $baseName = "$cleanSupplier-$dateStamp-$nameWithoutExt"
                    $counter = 0
                    
                    # Check SMB destination to see if a PDF with this name already exists
                    $finalPdfName = "$baseName.pdf"
                    while (Test-Path (Join-Path $finalSmbPath $finalPdfName)) {
                        $counter++
                        $paddedCounter = "{0:D2}" -f $counter
                        $finalPdfName = "$baseName($paddedCounter).pdf"
                    }
                    
                    # Apply the generated suffix to the staging file so it retains its native extension for conversion
                    if ($counter -eq 0) {
                        $newFileName = "$baseName$ext"
                    } else {
                        $newFileName = "$baseName($paddedCounter)$ext"
                    }
                    
                    Write-Host "  -> [RENAMING] New Invoice Name: $newFileName" -ForegroundColor Magenta
                    # ------------------------------------------------
                    
                    # --- DOWNLOAD & CONVERSION ENGINE ---
                    $stagingPath = Join-Path $config.Paths.Staging $newFileName
                    Write-Host "  -> [DOWNLOADING] Fetching file data to Staging..." -ForegroundColor Cyan
                    
                    try {
                        ## 1. Download the raw file via direct request (bypasses the SDK type-casting bug)
                        $uri = "https://graph.microsoft.com/v1.0/users/$targetMailbox/messages/$($msg.Id)/attachments/$($att.Id)"
                        $rawAttachment = Invoke-MgGraphRequest -Method GET -Uri $uri

                        # Convert the raw Base64 data back into actual file bytes
                        $fileBytes = [System.Convert]::FromBase64String($rawAttachment.contentBytes)

                        # Save to Staging
                        [System.IO.File]::WriteAllBytes($stagingPath, $fileBytes)
                        Write-Host "  -> [GOOD] File saved to Staging: $stagingPath" -ForegroundColor Green
                        
                        # 2. LibreOffice Conversion for Docs and Sheets
                        if ($ext -match "\.(docx|doc|xlsx|xls|csv)$") {
                            Write-Host "  -> [CONVERTING] Running LibreOffice Headless on $ext..." -ForegroundColor Cyan
                            $process = Start-Process -FilePath "libreoffice" -ArgumentList "--headless", "--convert-to", "pdf", "`"$stagingPath`"", "--outdir", "`"$($config.Paths.Staging)`"" -Wait -PassThru
                            
                            if ($process.ExitCode -eq 0) {
                                Write-Host "  -> [GOOD] Document successfully converted to PDF!" -ForegroundColor Green
                                # CLEANUP: Delete original raw file
                                Remove-Item -Path $stagingPath -Force
                                Write-Host "  -> [CLEANUP] Deleted original raw file: $newFileName" -ForegroundColor DarkGray
                            } else {
                                Write-Host "  -> [FAIL] LibreOffice conversion failed (Exit Code: $($process.ExitCode))" -ForegroundColor Red
                            }
                        }
                        # 3. img2pdf Conversion for Images
                        elseif ($ext -match "\.(jpg|jpeg)$") {
                            Write-Host "  -> [CONVERTING] Running img2pdf on $ext..." -ForegroundColor Cyan
                            $convertedPdfPath = [System.IO.Path]::ChangeExtension($stagingPath, ".pdf")
                            
                            $process = Start-Process -FilePath "img2pdf" -ArgumentList "`"$stagingPath`"", "-o", "`"$convertedPdfPath`"" -Wait -PassThru
                            
                            if ($process.ExitCode -eq 0) {
                                Write-Host "  -> [GOOD] Image successfully converted to PDF!" -ForegroundColor Green
                                # CLEANUP: Delete original raw file
                                Remove-Item -Path $stagingPath -Force
                                Write-Host "  -> [CLEANUP] Deleted original raw file: $newFileName" -ForegroundColor DarkGray
                            } else {
                                Write-Host "  -> [FAIL] img2pdf conversion failed (Exit Code: $($process.ExitCode))" -ForegroundColor Red
                            }
                        }

                        $validAttachments += $att
                    } catch {
                        Write-Host "  -> [FAIL] Could not save or process file. Error: $($_.Exception.Message)" -ForegroundColor Red
                    }
                    # ------------------------------------

                } else {
                    Write-Host "  -> [SKIP] Ignored Junk File: $($att.Name)" -ForegroundColor DarkGray
                }
            }
        }

        if ($validAttachments.Count -eq 0) {
            Write-Host " [WARNING] No valid invoice files found. Skipping email." -ForegroundColor Yellow
            continue 
        }

        # --- 3. FOLDER ROUTING LOGIC ---
        # <--- OLD ROUTING LOGIC DEPRECATED (Moved to Step 1 for Collision Support) --->
        # if ($supplierName -eq "Unknown") {
        #     Write-Host " [WARNING] No Match Found. Routing to Exceptions." -ForegroundColor Yellow
        #     $targetSubFolder = "Exceptions"
        # } else {
        #     $firstLetter = $supplierName.Substring(0,1).ToUpper()
        #     if ($firstLetter -match "[A-Z]") {
        #         $targetSubFolder = $firstLetter
        #     } else {
        #         $targetSubFolder = "#"
        #     }
        # }
        # $finalSmbPath = Join-Path $config.Paths.SMBDestination $targetSubFolder
        
        Write-Host " Routing $($validAttachments.Count) file(s) to SMB Folder: $finalSmbPath" -ForegroundColor Cyan
        
        # --- SIMULATE ACCESS (READ-ONLY TEST) ---
        if (Test-Path -Path $finalSmbPath) {
            Write-Host "   [GOOD] Folder Exists on SMB Share." -ForegroundColor Green
            try {
                $itemCount = (Get-ChildItem -Path $finalSmbPath -ErrorAction Stop).Count
                Write-Host "   [INFORMATIONAL] Successfully opened folder. It currently contains $itemCount items." -ForegroundColor DarkGray
            } 
            catch {
                Write-Host "   [FAIL] Folder exists, but READ ACCESS DENIED." -ForegroundColor Red
                Write-Host "       Error: $($_.Exception.Message)" -ForegroundColor DarkRed
            }
        } 
        else {
            Write-Host "   [FAIL] Folder DOES NOT EXIST ($finalSmbPath)" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "[FAIL] CRITICAL ERROR" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
}
finally {
    if (Get-MgContext) { 
        Disconnect-MgGraph | Out-Null 
        Write-Host "`n[INFORMATIONAL] Disconnected from Graph." -ForegroundColor DarkGray
    }
}