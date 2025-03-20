# Alternative example using a simpler format with the folder numbers directly in the name
# Note: The security groups will still be created with just the base name (without the number)
$exampleUsage2 = @'
   .\SharePoint-Folder-Creator.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/YourSite" -LibraryName "Documents" -FolderStructure @(
       "00 HR",
       "01 Finance",
       "02 IT",
       "03 Marketing"
   ) -SecurityGroupPrefix "SG-" -CreateGroupsIfNotExist
'@# SharePoint Folder Structure Creator with Unique Permissions
# This script creates folders in a SharePoint site and assigns unique permissions
# to Microsoft security groups that match the folder names

# Parameters
param(
    [Parameter(Mandatory=$true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory=$true)]
    [string]$LibraryName,
    
    [Parameter(Mandatory=$true)]
    [array]$FolderStructure,
    
    [Parameter(Mandatory=$false)]
    [string]$SecurityGroupPrefix = "SG-",
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateGroupsIfNotExist
)

# Load required SharePoint CSOM assemblies
try {
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
} catch {
    Write-Error "Failed to load SharePoint CSOM assemblies. Make sure they are installed."
    exit
}

# Connect to SharePoint site
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)
$credentials = Get-Credential -Message "Enter SharePoint Admin credentials"
$ctx.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($credentials.UserName, $credentials.Password)

# Connect to Microsoft Graph API for group operations
Connect-MgGraph -Scopes "Group.ReadWrite.All", "Directory.ReadWrite.All"

# Get SharePoint lists/libraries
$web = $ctx.Web
$lists = $web.Lists
$ctx.Load($web)
$ctx.Load($lists)
$ctx.ExecuteQuery()

# Get the document library
$library = $lists.GetByTitle($LibraryName)
$ctx.Load($library)
$ctx.ExecuteQuery()

# Function to check if Microsoft security group exists
function Check-SecurityGroupExists {
    param (
        [string]$groupName
    )
    
    try {
        $group = Get-MgGroup -Filter "displayName eq '$groupName' and securityEnabled eq true" -ErrorAction SilentlyContinue
        return ($null -ne $group)
    } catch {
        return $false
    }
}

# Function to create Microsoft security group if it doesn't exist
function Create-SecurityGroup {
    param (
        [string]$groupName,
        [string]$description
    )
    
    try {
        $params = @{
            displayName = $groupName
            description = $description
            mailEnabled = $false
            securityEnabled = $true
            mailNickname = ($groupName -replace '\s','')
        }
        
        $newGroup = New-MgGroup -BodyParameter $params
        Write-Host "Created new security group: $groupName" -ForegroundColor Green
        return $newGroup.Id
    } catch {
        Write-Error "Failed to create security group: $groupName. Error: $_"
        return $null
    }
}

# Function to break permission inheritance and set unique permissions
function Set-UniquePermissions {
    param (
        [Microsoft.SharePoint.Client.Folder]$folder,
        [string]$groupName
    )
    
    try {
        # Break inheritance
        $folder.ListItemAllFields.BreakRoleInheritance($false, $true)
        $ctx.ExecuteQuery()
        
        # Get the SharePoint group
        $spGroup = $web.SiteGroups.GetByName($groupName)
        $ctx.Load($spGroup)
        $ctx.ExecuteQuery()
        
        # Get the Full Control permission level
        $roleDefinition = $web.RoleDefinitions.GetByType([Microsoft.SharePoint.Client.RoleType]::Contributor)
        $ctx.Load($roleDefinition)
        $ctx.ExecuteQuery()
        
        # Create a role assignment with the Full Control permission level
        $roleAssignment = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
        $roleAssignment.Add($roleDefinition)
        
        # Assign the permission to the folder
        $folder.ListItemAllFields.RoleAssignments.Add($spGroup, $roleAssignment)
        $ctx.ExecuteQuery()
        
        Write-Host "Assigned unique permissions to folder '$($folder.Name)' for group '$groupName'" -ForegroundColor Green
    } catch {
        Write-Error "Failed to set unique permissions for folder '$($folder.Name)' and group '$groupName'. Error: $_"
    }
}

# Process each folder
foreach ($folderInfo in $FolderStructure) {
    # If array of strings is passed, convert to structured format
    if ($folderInfo -is [string]) {
        $folderNumber = $null
        $folderName = $folderInfo
    } else {
        $folderNumber = $folderInfo.Number
        $folderName = $folderInfo.Name
    }
    
    # Format folder name with number if provided
    $formattedFolderName = $folderName
    if ($folderNumber) {
        $formattedFolderName = "{0} {1}" -f $folderNumber, $folderName
    }
    
    Write-Host "Processing folder: $formattedFolderName" -ForegroundColor Cyan
    
    # Create the SharePoint folder
    $folderUrl = "$LibraryName/$formattedFolderName"
    $folder = $web.GetFolderByServerRelativeUrl($folderUrl)
    
    # Check if folder exists
    $ctx.Load($folder)
    try {
        $ctx.ExecuteQuery()
        Write-Host "Folder '$formattedFolderName' already exists." -ForegroundColor Yellow
    } catch {
        # Folder doesn't exist, create it
        $folderCreationInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation
        $folderCreationInfo.UnderlyingObjectType = 1 # Folder
        $folderCreationInfo.LeafName = $formattedFolderName
        
        $folderItem = $library.AddItem($folderCreationInfo)
        $folderItem["Title"] = $formattedFolderName
        $folderItem.Update()
        $ctx.ExecuteQuery()
        
        # Reload the folder
        $folder = $web.GetFolderByServerRelativeUrl($folderUrl)
        $ctx.Load($folder)
        $ctx.ExecuteQuery()
        
        Write-Host "Created folder: $formattedFolderName" -ForegroundColor Green
    }
    
    # Construct security group name - use the base folder name without the number prefix
    $securityGroupName = "$SecurityGroupPrefix$folderName"
    
    # Check if security group exists
    $groupExists = Check-SecurityGroupExists -groupName $securityGroupName
    
    if (-not $groupExists) {
        if ($CreateGroupsIfNotExist) {
            $groupDescription = "Security group for SharePoint folder: $formattedFolderName"
            $groupId = Create-SecurityGroup -groupName $securityGroupName -description $groupDescription
            
            if ($null -eq $groupId) {
                Write-Warning "Skipping permission assignment for folder '$formattedFolderName' due to group creation failure."
                continue
            }
            
            # Need to add this security group to SharePoint now
            # This part requires additional steps to add security group to SharePoint
            # Complex to implement in a script, consider using PnP PowerShell for this
        } else {
            Write-Warning "Security group '$securityGroupName' does not exist. Use -CreateGroupsIfNotExist to create it automatically."
            continue
        }
    }
    
    # Set unique permissions
    Set-UniquePermissions -folder $folder -groupName $securityGroupName
}

Write-Host "Script execution completed." -ForegroundColor Green

# Disconnect from Microsoft Graph
Disconnect-MgGraph