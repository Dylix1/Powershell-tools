# PowerShell Administrative Tools

A collection of PowerShell scripts for managing Exchange Online and Active Directory tasks through an interactive menu interface.

## Features

### 1. Exchange Online Archive Manager
- Enable and manage auto-expanding archives
- Check archive status and quotas
- Automated connection handling with retry logic
- Supports force enable option

### 2. Active Directory Group Member Export
- Export AD group members to CSV
- Supports local execution or script generation
- Includes display name, email, and SAM account details
- Timestamp-based file naming
- Recursive group membership resolution

### 3. Distribution Group Member Export
- Export distribution group members to CSV
- Display members in console with formatting
- Flexible output path selection
- Includes member names and primary SMTP addresses

### 4. Calendar Permissions Manager
- View and modify calendar permissions
- Support for all permission levels
- Handles both new and existing permissions
- Interactive permission level selection
- Built-in help system for access rights

## Prerequisites

- PowerShell 5.1 or later
- Exchange Online PowerShell V2 module (`Install-Module -Name ExchangeOnlineManagement`)
- Active Directory module for AD-related functions (`Install-Module -Name ActiveDirectory`)
- Appropriate administrative permissions for Exchange Online and/or Active Directory

## Installation

1. Clone or download all script files to a local directory
2. Ensure all scripts are in the same directory
3. If required, unblock the files:
   ```powershell
   Get-ChildItem -Path .\*.ps1 | Unblock-File
   ```

## Usage

1. Open PowerShell as administrator
2. Navigate to the script directory
3. Run the main menu script:
   ```powershell
   .\MainMenu.ps1
   ```
4. Select the desired tool from the menu

## Tool-Specific Instructions

### Exchange Online Archive Manager
- Requires Exchange Online admin credentials
- Automatically retries connection if initial attempt fails
- Shows current archive status before enabling auto-expand

### Active Directory Group Export
- Can be run locally or generate a portable script
- Supports special characters in group names
- Creates timestamped output files
- Includes detailed user properties

### Distribution Group Export
- Supports custom output paths
- Optional CSV export
- Displays member count and details
- Automatic session cleanup

### Calendar Permissions Manager
- Supports all standard Exchange permission levels
- Interactive permission management
- Built-in access rights documentation
- Handles both new and existing permissions

## Permission Levels

Calendar permission levels available:
- Owner
- PublishingEditor
- Editor
- PublishingAuthor
- Author
- NonEditingAuthor
- Reviewer
- Contributor
- None

## Error Handling

- All tools include comprehensive error handling
- Failed operations are logged with detailed error messages
- Automatic cleanup of Exchange Online sessions
- Retry logic for connection attempts

## Security Notes

- Scripts use modern authentication by default
- Basic authentication available when required
- Credentials are never stored
- Sessions are properly terminated after use

## Troubleshooting

If you encounter issues:

1. Ensure you have the required permissions
2. Verify all prerequisite modules are installed
3. Check for any network connectivity issues
4. Review error messages in the console output
5. Ensure scripts are run with appropriate execution policy

## Contributing

Contributions are welcome! Please feel free to submit pull requests with improvements or bug fixes.

## License

This project is licensed under the MIT License - see the LICENSE file for details.