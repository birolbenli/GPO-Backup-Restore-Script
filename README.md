# GPO Backup & Restore Tool (GUI)

A user-friendly PowerShell GUI tool for backing up and restoring Group Policy Objects (GPOs) in Active Directory environments.

## Features

- Easy-to-use graphical interface
- Select multiple GPOs to back up or restore
- Choose backup and restore folder locations interactively
- Progress bar for backup operations
- Clear separation between backup and restore processes
- Lightweight, no external dependencies

## Requirements

- Windows PowerShell (tested on PowerShell 5.1)
- Active Directory Module for Windows PowerShell (part of RSAT tools)
- Appropriate permissions to backup and restore GPOs

## Usage

1. Download or clone this repository.
2. Open PowerShell **as Administrator**.
3. (Optional) Temporarily allow script execution:

    ```powershell
    Set-ExecutionPolicy Bypass -Scope Process -Force
    ```
4. Run the script:

    ```powershell
    .\GPO-Backup-Restore.ps1
    ```

5. Use the GUI to select:

   - Backup folder and GPOs to back up  
   - Restore folder and GPOs to restore

## Notes

- Backups are saved in folders named after each GPO under the selected backup directory.
- Restores use the selected backup directory to locate saved GPOs.
- Progress bar shows backup progress; restore runs silently with a completion message.

## License

This project is provided "as is" without warranty of any kind.

---

*Developed by Birol Benli*

