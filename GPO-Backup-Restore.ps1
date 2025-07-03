Add-Type -AssemblyName System.Windows.Forms

# Default backup directory
$BackupPath = "C:\GPO_Backup"

function Select-FolderDialog([string]$description, [string]$initialPath) {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = $description
    $folderBrowser.SelectedPath = $initialPath
    if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowser.SelectedPath
    } else {
        return $null
    }
}

function Select-GPOs-GUI {
    param (
        [string]$Mode
    )

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Select GPO(s) for $Mode"
    $Form.Size = '500,500'
    $Form.StartPosition = 'CenterScreen'

    $ListBox = New-Object System.Windows.Forms.CheckedListBox
    $ListBox.Size = '460,360'
    $ListBox.Location = '10,10'

    if ($Mode -eq "Backup") {
        $GPOs = Get-GPO -All | Sort-Object DisplayName
    } else {
        $GPOs = Get-ChildItem -Path $BackupPath -Directory | ForEach-Object {
            [PSCustomObject]@{
                DisplayName = $_.Name
                Path = $_.FullName
            }
        }
    }

    foreach ($gpo in $GPOs) {
        $ListBox.Items.Add($gpo.DisplayName) | Out-Null
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
    $path = Select-FolderDialog -description "Select or create backup folder." -initialPath $BackupPath
    if (-not $path) { return }
    $BackupPath = $path
    if (!(Test-Path $BackupPath)) { New-Item -ItemType Directory -Path $BackupPath | Out-Null }

    $gpos = Select-GPOs-GUI -Mode "Backup"
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

        $folder = Join-Path -Path $BackupPath -ChildPath $gpo.DisplayName
        if (!(Test-Path $folder)) {
            New-Item -ItemType Directory -Path $folder | Out-Null
        }

        Backup-GPO -Name $gpo.DisplayName -Path $folder -Comment "Backup: $(Get-Date)" | Out-Null
    }

    $label.Text = "✅ Backup completed."
    Start-Sleep -Seconds 2
    $progressForm.Close()
}

function Restore-GPOs {
    $path = Select-FolderDialog -description "Select the backup folder to restore from." -initialPath $BackupPath
    if (-not $path) { return }
    $BackupPath = $path

    $gpos = Select-GPOs-GUI -Mode "Restore"
    if ($gpos.Count -eq 0) { return }

    foreach ($gpo in $gpos) {
        Restore-GPO -Name $gpo.DisplayName -Path $gpo.Path -Confirm:$false
    }

    [System.Windows.Forms.MessageBox]::Show("Selected GPO(s) have been restored successfully.", "Operation Complete")
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
