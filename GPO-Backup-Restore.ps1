Add-Type -AssemblyName System.Windows.Forms

function Select-FolderDialog([string]$description, [string]$initialPath) {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $description
    $folderBrowser.SelectedPath = $initialPath
    $folderBrowser.ShowNewFolderButton = $true
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    } else {
        return $null
    }
}

function Select-GPOs-GUI {
    param (
        [string]$Mode,
        [string]$BackupPath
    )

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Select GPO(s) for $Mode"
    $Form.Size = '500,500'
    $Form.StartPosition = 'CenterScreen'
    $Form.Topmost = $true

    $ListBox = New-Object System.Windows.Forms.CheckedListBox
    $ListBox.Size = '460,360'
    $ListBox.Location = '10,10'

    if ($Mode -eq "Backup") {
        $GPOs = Get-GPO -All | Sort-Object DisplayName
    } else {
        $GPOs = @()
        
        # Get all backup folders (GUIDs)
        $backupFolders = Get-ChildItem -Path $BackupPath -Directory | Where-Object { $_.Name -match "^{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}}$" }
        
        foreach ($folder in $backupFolders) {
            $manifestFile = Get-ChildItem -Path $folder.FullName -Filter "manifest.xml" -ErrorAction SilentlyContinue
            if ($manifestFile) {
                try {
                    [xml]$manifest = Get-Content $manifestFile.FullName
                    $gpoDisplayName = $manifest.Backups.BackupInst.GPODisplayName
                    
                    if ($gpoDisplayName) {
                        $GPOs += [PSCustomObject]@{
                            DisplayName = $gpoDisplayName
                            Path = $folder.FullName
                            BackupId = $folder.Name.Trim('{}')
                        }
                    }
                } catch {
                    # Skip invalid manifest files
                }
            }
        }
        
        $GPOs = $GPOs | Sort-Object DisplayName
    }

    foreach ($gpo in $GPOs) {
        $ListBox.Items.Add($gpo.DisplayName) | Out-Null
    }

    # If no GPOs found, show a message
    if ($GPOs.Count -eq 0 -and $Mode -eq "Restore") {
        $noGpoLabel = New-Object System.Windows.Forms.Label
        $noGpoLabel.Text = "No backup files found in the selected folder.`nPlease select a folder containing GPO backups."
        $noGpoLabel.AutoSize = $false
        $noGpoLabel.Size = '440,60'
        $noGpoLabel.Location = '20,180'
        $noGpoLabel.TextAlign = 'MiddleCenter'
        $noGpoLabel.ForeColor = [System.Drawing.Color]::Red
        $Form.Controls.Add($noGpoLabel)
    }

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = "Select All"
    $btnSelectAll.Location = '10,380'
    $btnSelectAll.Add_Click({ for ($i = 0; $i -lt $ListBox.Items.Count; $i++) { $ListBox.SetItemChecked($i, $true) } })

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = "Clear All"
    $btnClear.Location = '120,380'
    $btnClear.Add_Click({ for ($i = 0; $i -lt $ListBox.Items.Count; $i++) { $ListBox.SetItemChecked($i, $false) } })

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Text = "OK"
    $OKButton.Location = '350,380'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $Form.AcceptButton = $OKButton

    $Form.Controls.AddRange(@($ListBox, $btnSelectAll, $btnClear, $OKButton))
    
    if ($Form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selected = @()
        for ($i = 0; $i -lt $ListBox.Items.Count; $i++) {
            if ($ListBox.GetItemChecked($i)) {
                $name = $ListBox.Items[$i]
                $selected += $GPOs | Where-Object { $_.DisplayName -eq $name }
            }
        }
        return $selected
    } else {
        return @()
    }
}

function Backup-GPOs {
    $path = Select-FolderDialog -description "Select or create backup folder." -initialPath ""
    if (-not $path) { return }
    if (!(Test-Path $path)) { New-Item -ItemType Directory -Path $path | Out-Null }

    $gpos = Select-GPOs-GUI -Mode "Backup" -BackupPath $path
    if ($gpos.Count -eq 0) { return }

    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Backing up GPO(s)..."
    $progressForm.Size = '400,150'
    $progressForm.StartPosition = 'CenterScreen'
    $progressForm.Topmost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.AutoSize = $true
    $label.Location = '20,20'
    $label.Text = "Starting backup..."
    $progressForm.Controls.Add($label)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = '20,50'
    $progressBar.Size = '340,25'
    $progressBar.Minimum = 0
    $progressBar.Maximum = $gpos.Count
    $progressBar.Value = 0
    $progressForm.Controls.Add($progressBar)

    $progressForm.Show()

    $index = 0
    foreach ($gpo in $gpos) {
        $index++
        $label.Text = "Backing up: $($gpo.DisplayName) ($index / $($gpos.Count))"
        $progressBar.Value = $index
        $progressForm.Refresh()

        # Backup the GPO
        $backupResult = Backup-GPO -Name $gpo.DisplayName -Path $path -Comment "Backup: $(Get-Date)"
        
        # Create manifest.xml for easier restore
        $backupId = $backupResult.Id.ToString()
        $manifestContent = @"
<?xml version="1.0" encoding="utf-8"?>
<Backups xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://www.microsoft.com/GroupPolicy/GPOOperations/Manifest">
  <BackupInst>
    <GPOGuid>{$($gpo.Id.ToString())}</GPOGuid>
    <GPODisplayName>$($gpo.DisplayName)</GPODisplayName>
    <GPODomainName>$($gpo.DomainName)</GPODomainName>
    <BackupTime>$(Get-Date -Format "yyyy-MM-ddTHH:mm:ss")</BackupTime>
    <ID>{$backupId}</ID>
    <Comment>Backup: $(Get-Date)</Comment>
  </BackupInst>
</Backups>
"@
        
        # Save manifest.xml in the backup folder
        $backupFolder = Get-ChildItem -Path $path -Directory | Where-Object { $_.Name -eq "{$backupId}" } | Select-Object -First 1
        if ($backupFolder) {
            $manifestPath = Join-Path -Path $backupFolder.FullName -ChildPath "manifest.xml"
            $manifestContent | Out-File -FilePath $manifestPath -Encoding UTF8
            Write-Host "Created manifest.xml for $($gpo.DisplayName) in folder {$backupId}"
        } else {
            Write-Host "Warning: Could not find backup folder {$backupId} for $($gpo.DisplayName)"
        }
    }

    $label.Text = "✅ Backup completed."
    Start-Sleep -Seconds 2
    $progressForm.Close()
    
    $completionForm = New-Object System.Windows.Forms.Form
    $completionForm.Size = '1,1'
    $completionForm.WindowState = 'Minimized'
    $completionForm.ShowInTaskbar = $false
    $completionForm.TopMost = $true
    $completionForm.Show()
    [System.Windows.Forms.MessageBox]::Show($completionForm, "Selected GPO(s) have been backed up successfully.", "Backup Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    $completionForm.Close()
}

function Test-ADWSConnection {
    param (
        [string]$Server = $null
    )
    
    try {
        if ($Server) {
            Get-ADDomain -Server $Server -ErrorAction Stop | Out-Null
        } else {
            Get-ADDomain -ErrorAction Stop | Out-Null
        }
        return $true
    } catch {
        return $false
    }
}

function Select-DomainController {
    param (
        [string]$CurrentDomain = ""
    )
    
    $dcForm = New-Object System.Windows.Forms.Form
    $dcForm.Text = "Domain Controller Selection"
    $dcForm.Size = '520,420'
    $dcForm.StartPosition = 'CenterScreen'
    $dcForm.Topmost = $true
    $dcForm.FormBorderStyle = 'FixedDialog'
    $dcForm.MaximizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "ADWS connection issue detected. Please select a domain controller option:"
    $label.AutoSize = $false
    $label.Size = '480,30'
    $label.Location = '20,20'
    $label.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $dcForm.Controls.Add($label)

    $radioAuto = New-Object System.Windows.Forms.RadioButton
    $radioAuto.Text = "🔄 Auto-detect (recommended)"
    $radioAuto.Size = '250,20'
    $radioAuto.Location = '20,60'
    $radioAuto.Checked = $true
    $dcForm.Controls.Add($radioAuto)

    $radioPDC = New-Object System.Windows.Forms.RadioButton
    $radioPDC.Text = "🎯 Use PDC Emulator"
    $radioPDC.Size = '250,20'
    $radioPDC.Location = '20,90'
    $dcForm.Controls.Add($radioPDC)

    $radioCustom = New-Object System.Windows.Forms.RadioButton
    $radioCustom.Text = "⚙️ Custom domain controller:"
    $radioCustom.Size = '250,20'
    $radioCustom.Location = '20,120'
    $dcForm.Controls.Add($radioCustom)

    $textCustomDC = New-Object System.Windows.Forms.TextBox
    $textCustomDC.Location = '40,145'
    $textCustomDC.Size = '350,20'
    $textCustomDC.Text = if ($CurrentDomain) { "dc.$CurrentDomain" } else { "dc.yourdomain.com" }
    $textCustomDC.Enabled = $false
    $dcForm.Controls.Add($textCustomDC)

    $radioCustom.Add_CheckedChanged({ $textCustomDC.Enabled = $radioCustom.Checked })

    $radioLocal = New-Object System.Windows.Forms.RadioButton
    $radioLocal.Text = "🖥️ Use local machine as DC (testing only)"
    $radioLocal.Size = '350,20'
    $radioLocal.Location = '20,180'
    $dcForm.Controls.Add($radioLocal)

    $btnTest = New-Object System.Windows.Forms.Button
    $btnTest.Text = "Test Connection"
    $btnTest.Size = '120,25'
    $btnTest.Location = '40,210'
    $btnTest.Add_Click({
        $testDC = $null
        if ($radioAuto.Checked) {
            $testDC = $null
        } elseif ($radioPDC.Checked) {
            try { $testDC = (Get-ADDomain).PDCEmulator } catch { $testDC = $null }
        } elseif ($radioCustom.Checked) {
            $testDC = $textCustomDC.Text.Trim()
        } elseif ($radioLocal.Checked) {
            # For local machine, try multiple formats
            $testDC = "localhost"
        }
        
        $testResult = Test-ADWSConnection -Server $testDC
        $testMsg = if ($testResult) { "✅ Connection successful!" } else { "❌ Connection failed!" }
        
        # If local machine test fails, try alternative formats
        if (-not $testResult -and $radioLocal.Checked) {
            $testDC = $env:COMPUTERNAME
            $testResult = Test-ADWSConnection -Server $testDC
            if ($testResult) {
                $testMsg = "✅ Connection successful using computer name!"
            } else {
                # Try FQDN
                try {
                    $domain = (Get-ADDomain).DNSRoot
                    $testDC = "$($env:COMPUTERNAME).$domain"
                    $testResult = Test-ADWSConnection -Server $testDC
                    if ($testResult) {
                        $testMsg = "✅ Connection successful using FQDN!"
                    } else {
                        $testMsg = "❌ Connection failed with all formats (localhost, $($env:COMPUTERNAME), $testDC)"
                    }
                } catch {
                    $testMsg = "❌ Connection failed - Unable to determine domain"
                }
            }
        }
        
        [System.Windows.Forms.MessageBox]::Show($dcForm, $testMsg, "Connection Test", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
    $dcForm.Controls.Add($btnTest)

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "🔧 Troubleshooting Tips:`n• Ensure ADWS service is running on the domain controller`n• Check network connectivity and firewall (port 9389)`n• Try running the script directly on a domain controller`n• Verify you have appropriate permissions`n• Consider using 'runas /user:domain\admin' to run with elevated privileges`n• For local DC: Check 'Get-Service ADWS' and 'netstat -an | findstr :9389'"
    $infoLabel.AutoSize = $false
    $infoLabel.Size = '480,120'
    $infoLabel.Location = '20,250'
    $infoLabel.ForeColor = [System.Drawing.Color]::Blue
    $dcForm.Controls.Add($infoLabel)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = '320,380'
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dcForm.AcceptButton = $btnOK
    $dcForm.Controls.Add($btnOK)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = '410,380'
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dcForm.Controls.Add($btnCancel)

    $result = $dcForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($radioAuto.Checked) {
            return $null
        } elseif ($radioPDC.Checked) {
            try {
                $pdc = (Get-ADDomain).PDCEmulator
                return $pdc
            } catch {
                return $null
            }
        } elseif ($radioCustom.Checked -and $textCustomDC.Text.Trim()) {
            return $textCustomDC.Text.Trim()
        } elseif ($radioLocal.Checked) {
            # For local machine, try multiple formats to find the best one
            $localDC = $null
            
            # Try localhost first
            if (Test-ADWSConnection -Server "localhost") {
                $localDC = "localhost"
            }
            # Try computer name
            elseif (Test-ADWSConnection -Server $env:COMPUTERNAME) {
                $localDC = $env:COMPUTERNAME
            }
            # Try FQDN
            else {
                try {
                    $domain = (Get-ADDomain).DNSRoot
                    $fqdn = "$($env:COMPUTERNAME).$domain"
                    if (Test-ADWSConnection -Server $fqdn) {
                        $localDC = $fqdn
                    }
                } catch {
                    # If all else fails, use computer name
                    $localDC = $env:COMPUTERNAME
                }
            }
            
            return $localDC
        }
    }
    
    return "CANCEL"
}

function Restore-GPOs {
    $path = Select-FolderDialog -description "Select the backup folder to restore from." -initialPath ""
    if (-not $path) { return }

    $gpos = Select-GPOs-GUI -Mode "Restore" -BackupPath $path
    if ($gpos.Count -eq 0) { return }

    # Get current domain info first
    $currentDomain = $null
    $currentDC = $null
    $adwsConnectionGood = $false
    $selectedDC = $null
    
    try {
        $currentDomain = (Get-ADDomain).DNSRoot
        $currentDC = (Get-ADDomain).PDCEmulator
        $adwsConnectionGood = Test-ADWSConnection -Server $currentDC
        Write-Host "Current domain: $currentDomain"
        Write-Host "Current DC: $currentDC"
        Write-Host "ADWS connection status: $adwsConnectionGood"
    } catch {
        Write-Host "Warning: Could not determine current domain. Will use local machine as DC."
        $adwsConnectionGood = $false
    }
    
    # If ADWS connection is problematic, automatically use local machine as DC
    if (-not $adwsConnectionGood) {
        Write-Host "ADWS connection issue detected. Using local machine as DC..."
        
        # Try localhost first, then computer name, then FQDN
        $localDC = $null
        if (Test-ADWSConnection -Server "localhost") {
            $localDC = "localhost"
            Write-Host "✓ Using localhost as DC"
        } elseif (Test-ADWSConnection -Server $env:COMPUTERNAME) {
            $localDC = $env:COMPUTERNAME
            Write-Host "✓ Using computer name ($($env:COMPUTERNAME)) as DC"
        } else {
            try {
                $domain = (Get-ADDomain).DNSRoot
                $fqdn = "$($env:COMPUTERNAME).$domain"
                if (Test-ADWSConnection -Server $fqdn) {
                    $localDC = $fqdn
                    Write-Host "✓ Using FQDN ($fqdn) as DC"
                } else {
                    # Use localhost as fallback
                    $localDC = "localhost"
                    Write-Host "⚠ No successful connection test, using localhost as fallback"
                }
            } catch {
                $localDC = "localhost"
                Write-Host "⚠ Could not determine FQDN, using localhost as fallback"
            }
        }
        
        $currentDC = $localDC
        Write-Host "Selected local DC: $currentDC"
        
        # Check if local machine is actually a DC
        try {
            $isDC = (Get-WmiObject -Class Win32_OperatingSystem).ProductType -eq 2
            if ($isDC) {
                Write-Host "✓ Local machine is confirmed as a domain controller."
            } else {
                Write-Host "⚠ Warning: Local machine may not be a domain controller, but continuing anyway."
            }
        } catch {
            Write-Host "Could not verify if local machine is a DC. Continuing anyway."
        }
        
        # Try to get domain info with local DC
        try {
            if (-not $currentDomain) {
                $currentDomain = (Get-ADDomain -Server $currentDC).DNSRoot
                Write-Host "Got domain info using local DC: $currentDomain"
            }
        } catch {
            Write-Host "Could not get domain info from local DC. Will proceed without domain parameter."
        }
    }
    
    Write-Host "Final configuration - Domain: $currentDomain, DC: $currentDC"

    $progressForm = New-Object System.Windows.Forms.Form
    $progressForm.Text = "Restoring GPO(s)..."
    $progressForm.Size = '450,180'
    $progressForm.StartPosition = 'CenterScreen'
    $progressForm.Topmost = $true

    $label = New-Object System.Windows.Forms.Label
    $label.AutoSize = $true
    $label.Location = '20,20'
    $label.Text = "Starting restore..."
    $progressForm.Controls.Add($label)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = '20,50'
    $progressBar.Size = '400,25'
    $progressBar.Minimum = 0
    $progressBar.Maximum = $gpos.Count
    $progressBar.Value = 0
    $progressForm.Controls.Add($progressBar)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.AutoSize = $true
    $statusLabel.Location = '20,90'
    $statusLabel.Text = "Preparing..."
    $progressForm.Controls.Add($statusLabel)

    $progressForm.Show()

    $index = 0
    $successCount = 0
    $errorCount = 0
    $errorDetails = @()

    foreach ($gpo in $gpos) {
        $index++
        $label.Text = "Restoring: $($gpo.DisplayName) ($index / $($gpos.Count))"
        if ($index -ge $progressBar.Minimum -and $index -le $progressBar.Maximum) {
            $progressBar.Value = $index
        } elseif ($index -gt $progressBar.Maximum) {
            $progressBar.Value = $progressBar.Maximum
        } else {
            $progressBar.Value = $progressBar.Minimum
        }
        $statusLabel.Text = "Processing GPO: $($gpo.DisplayName)"
        $progressForm.Refresh()

        try {
            Write-Host "Restoring GPO: $($gpo.DisplayName) with BackupId: $($gpo.BackupId)"
            
            # Test if GPO already exists - use selected DC if available
            $existingGPO = $null
            try {
                if ($currentDomain -and $currentDC) {
                    Write-Host "Checking if GPO exists using domain: $currentDomain and DC: $currentDC"
                    $existingGPO = Get-GPO -Name $gpo.DisplayName -Domain $currentDomain -Server $currentDC -ErrorAction SilentlyContinue
                } elseif ($currentDomain) {
                    Write-Host "Checking if GPO exists using domain: $currentDomain"
                    $existingGPO = Get-GPO -Name $gpo.DisplayName -Domain $currentDomain -ErrorAction SilentlyContinue
                } else {
                    Write-Host "Checking if GPO exists using default connection"
                    $existingGPO = Get-GPO -Name $gpo.DisplayName -ErrorAction SilentlyContinue
                }
                
                if ($existingGPO) {
                    Write-Host "✓ GPO '$($gpo.DisplayName)' exists"
                } else {
                    Write-Host "○ GPO '$($gpo.DisplayName)' does not exist - will be created"
                }
            } catch {
                # GPO doesn't exist or can't be accessed - assume it doesn't exist
                Write-Host "Could not check GPO existence: $($_.Exception.Message) - assuming it doesn't exist"
                $existingGPO = $null
            }
            
            # Import the GPO with appropriate parameters
            if ($currentDomain -and $currentDC) {
                Write-Host "Using current domain: $currentDomain and DC: $currentDC"
                if ($existingGPO) {
                    Write-Host "GPO exists. Importing into existing GPO..."
                    Import-GPO -BackupId $gpo.BackupId -TargetName $gpo.DisplayName -Path $path -Domain $currentDomain -Server $currentDC -ErrorAction Stop
                } else {
                    Write-Host "GPO does not exist. Creating new GPO..."
                    Import-GPO -BackupId $gpo.BackupId -TargetName $gpo.DisplayName -Path $path -Domain $currentDomain -Server $currentDC -CreateIfNeeded -ErrorAction Stop
                }
            } elseif ($currentDomain) {
                Write-Host "Using current domain: $currentDomain"
                if ($existingGPO) {
                    Write-Host "GPO exists. Importing into existing GPO..."
                    Import-GPO -BackupId $gpo.BackupId -TargetName $gpo.DisplayName -Path $path -Domain $currentDomain -ErrorAction Stop
                } else {
                    Write-Host "GPO does not exist. Creating new GPO..."
                    Import-GPO -BackupId $gpo.BackupId -TargetName $gpo.DisplayName -Path $path -Domain $currentDomain -CreateIfNeeded -ErrorAction Stop
                }
            } else {
                Write-Host "Using default connection"
                if ($existingGPO) {
                    Write-Host "GPO exists. Importing into existing GPO..."
                    Import-GPO -BackupId $gpo.BackupId -TargetName $gpo.DisplayName -Path $path -ErrorAction Stop
                } else {
                    Write-Host "GPO does not exist. Creating new GPO..."
                    Import-GPO -BackupId $gpo.BackupId -TargetName $gpo.DisplayName -Path $path -CreateIfNeeded -ErrorAction Stop
                }
            }
            
            $successCount++
            $statusLabel.Text = "✓ Restored: $($gpo.DisplayName)"
            Write-Host "✓ Successfully restored: $($gpo.DisplayName)"
            
        }
        catch {
            $errorCount++
            $errorMessage = $_.Exception.Message
            $errorDetails += "GPO: $($gpo.DisplayName) - Error: $errorMessage"
            
            Write-Host "✗ Error restoring GPO '$($gpo.DisplayName)': $errorMessage"
            $statusLabel.Text = "✗ Failed: $($gpo.DisplayName)"
            
            # Analyze error type and provide specific guidance
            $guidanceMsg = ""
            if ($errorMessage -match "ADWS|Active Directory Web Services") {
                $guidanceMsg = "`n🔧 ADWS Issue - Try:`n• Restart ADWS service on DC`n• Run: Restart-Service ADWS`n• Check port 9389 connectivity"
            } elseif ($errorMessage -match "network|RPC|endpoint") {
                $guidanceMsg = "`n🔧 Network Issue - Try:`n• Check network connectivity to DC`n• Verify firewall settings`n• Run from a domain-joined machine"
            } elseif ($errorMessage -match "permission|access|denied") {
                $guidanceMsg = "`n🔧 Permission Issue - Try:`n• Run as Domain Admin`n• Check Group Policy permissions`n• Verify delegation rights"
            } elseif ($errorMessage -match "not found|does not exist") {
                $guidanceMsg = "`n🔧 GPO Not Found - Try:`n• Check if backup ID is correct`n• Verify backup path is accessible`n• Ensure backup files are intact"
            } else {
                $guidanceMsg = "`n🔧 General troubleshooting:`n• Check Event Viewer for more details`n• Try running script on a DC`n• Verify backup integrity"
            }
            
            # Show detailed error message with guidance
            $errorMsg = "Error restoring GPO '$($gpo.DisplayName)':`n$errorMessage"
            $errorMsg += $guidanceMsg
            $errorMsg += "`n`nWould you like to continue with the remaining GPOs?"
            
            $result = [System.Windows.Forms.MessageBox]::Show($progressForm, $errorMsg, "Restore Error", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Error)
            if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                Write-Host "User chose to stop restore operation after error."
                break
            }
        }
        
        Start-Sleep -Milliseconds 500
    }

    # Döngü sonunda progress bar'ı tamamla
    $progressBar.Value = $progressBar.Maximum
    $statusLabel.Text = "Finalizing restore process..."
    $progressForm.Refresh()
    Start-Sleep -Milliseconds 500
    $progressForm.Close()

    # Domain replikasyonunu arka planda başlat
    Write-Host "Forcing domain replication..."
    try {
        if ($currentDC) {
            Start-Process -FilePath "repadmin" -ArgumentList "/syncall $currentDC /A /e /P" -NoNewWindow -WindowStyle Hidden -ErrorAction SilentlyContinue
        } else {
            Start-Process -FilePath "repadmin" -ArgumentList "/syncall /A /e /P" -NoNewWindow -WindowStyle Hidden -ErrorAction SilentlyContinue
        }
        Write-Host "✓ Domain replication initiated (background)"
    } catch {
        Write-Host "Could not force domain replication: $($_.Exception.Message)"
    }

    # Show completion summary
    $summaryMsg = "🔄 GPO Restore Operation Summary`n`n"
    $summaryMsg += "Current Domain: $currentDomain`n"
    $summaryMsg += "✅ Successfully restored: $successCount GPO(s)`n"
    if ($errorCount -gt 0) {
        $summaryMsg += "❌ Failed to restore: $errorCount GPO(s)`n`n"
        $summaryMsg += "Failed GPOs:`n"
        foreach ($errorDetail in $errorDetails) {
            $summaryMsg += "• $errorDetail`n"
        }
        $summaryMsg += "`n"
    }
    $summaryMsg += "📋 Next Steps:`n"
    $summaryMsg += "• Refresh Group Policy Management Console (F5)`n"
    $summaryMsg += "• Wait a few minutes for AD replication`n"
    $summaryMsg += "• Check GPO links and permissions`n"
    $summaryMsg += "• Run 'gpupdate /force' on client machines if needed"
    
    $completionForm = New-Object System.Windows.Forms.Form
    $completionForm.Size = '1,1'
    $completionForm.WindowState = 'Minimized'
    $completionForm.ShowInTaskbar = $false
    $completionForm.TopMost = $true
    $completionForm.Show()
    
    $icon = if ($errorCount -gt 0) { [System.Windows.Forms.MessageBoxIcon]::Warning } else { [System.Windows.Forms.MessageBoxIcon]::Information }
    [System.Windows.Forms.MessageBox]::Show($completionForm, $summaryMsg, "Restore Complete", [System.Windows.Forms.MessageBoxButtons]::OK, $icon)
    $completionForm.Close()
}

# === Main Menu ===
$form = New-Object System.Windows.Forms.Form
$form.Text = "GPO Backup & Restore Tool"
$form.Size = '400,230'
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.Topmost = $true

$btnBackup = New-Object System.Windows.Forms.Button
$btnBackup.Text = "Backup GPO(s)"
$btnBackup.Size = '100,40'
$btnBackup.Location = '50,50'
$btnBackup.Add_Click({ Backup-GPOs })

$btnRestore = New-Object System.Windows.Forms.Button
$btnRestore.Text = "Restore GPO(s)"
$btnRestore.Size = '100,40'
$btnRestore.Location = '160,50'
$btnRestore.Add_Click({ Restore-GPOs })

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "Exit"
$btnExit.Size = '100,40'
$btnExit.Location = '270,50'
$btnExit.Add_Click({ $form.Close() })

# Footer Label
$lblFooter = New-Object System.Windows.Forms.Label
$lblFooter.Text = "Developer: https://github.com/birolbenli/"
$lblFooter.AutoSize = $true
$lblFooter.Location = '110,150'
$lblFooter.Font = New-Object System.Drawing.Font("Segoe UI",8,[System.Drawing.FontStyle]::Italic)
$lblFooter.ForeColor = [System.Drawing.Color]::Gray
$form.Controls.Add($lblFooter)

$form.Controls.AddRange(@($btnBackup, $btnRestore, $btnExit))
$form.ShowDialog()
