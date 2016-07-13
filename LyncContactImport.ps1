<#
 #
 #  Author: Richard Rios
 #  Created: Jan. 15, 2016
 #
 #  Purpose:
 #     The purpose of this script is to allow a user to automatically import
 #     all their Skype contacts in to Outlook without creating duplicate contacts.
 #
 #     This was tested using Skype for Business 2013 and Outlook 2013
 #>


FUNCTION WriteInformationMessage($msg)
{
    Write-Host -ForegroundColor Yellow "[INFORMATION]" -NoNewline
    Write-Host " $msg"
}

if(-not (Get-Module -Name Microsoft.Lync.Model))
{
    try
    {
        $location = Get-Location
        $path = "$PSScriptRoot\Assemblies\Desktop\Microsoft.Lync.Model.dll"
        Add-Type -Path $path
    }
    catch
    {
        Write-Host -ForegroundColor Red "[FATAL ERROR]" -NoNewline
        Write-Host " .NET Framework 3.5 or 4, or the Assembly folder is missing."
        exit
    }
}

Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"

$outlook = New-Object -ComObject Outlook.Application
$ol_namespace = $outlook.GetNameSpace("MAPI")
$ol_contact_folder = [Microsoft.Office.Interop.Outlook.olDefaultFolders]::olFolderContacts
$ol_contact_obj = $ol_namespace.GetDefaultFolder($ol_contact_folder).Items
$lync_client = [Microsoft.Lync.Model.LyncClient]::GetClient()
# Get the current contact groups
$contact_groups = $lync_client.ContactManager.Groups
$skype_contacts = @()
$ol_contacts = @()
$contact_enum = @{ 
    "DisplayName" = 10; 
    "Email" = 12; 
    "Title" = 14; 
    "Company" = 15; 
    "Phone" = 27; 
    "FirstName" = 37; 
    "LastName" = 39; 
    "MiddleName" = 40
}

WriteInformationMessage "Gathering Lync contact list"
# Iterate through each Skype group
$contact_groups | ForEach-Object {
    $grp = $_
    # Iterate through each contact in the Skype group
    $grp | ForEach-Object {
        $obj = @{
            "Name" = $_.GetContactInformation($contact_enum."DisplayName");
            "Email" = $_.GetContactInformation($contact_enum."Email");
        }
        $skype_contacts += New-Object -TypeName PSObject -Property $obj
    }
}

WriteInformationMessage "Gathering Outlook contact list"
# Create a list of outlook contacts
$ol_contact_obj | ForEach-Object {
    $obj = @{
        "Name" = $_.FullName;
        "Email" = $_.IMAddress;
    }
    $ol_contacts += New-Object -TypeName PSObject -Property $obj
}

WriteInformationMessage "Evaluating contacts for importation"
$added_count = 0
# Determine if the Skype user is already an outlook contact. If not, add them to Outlook Contacts
$skype_contacts | ForEach-Object {
    $new_contact = $ol_contact_obj.Add()
    if($ol_contacts.Name -notcontains $_.Name)
    {
        ++$added_count
        $new_contact.FullName = $_.Name
        $new_contact.IMAddress = $_.Email
        $new_contact.Save()
        Write-Host -ForegroundColor Yellow "[ADDED] " -NoNewline
        Wirte-Host $_.Name
    }
}

if($added_count -eq 0)
{
    WriteInformationMessage "All Skype contacts already exist in Outlook"
}
Write-Host -ForegroundColor Green "`n`r`n`r------------------ Import Script Finished ------------------"