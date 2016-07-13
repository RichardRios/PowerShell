<#
 #
 #  Author: Richard Rios
 #  Created: Jan. 12, 2016
 #
 #  Purpose:
 #     The purpose of this script is to allow a simple and efficient way for
 #     a user to manage their outlook rules and folders. The script is designed
 #     to create a subfolder in the inbox with subfolders of contacts names and
 #     apply rules to those subfolders.
 #     
 #     This script was tested using Skype for Business 2013 with Outlook 2013.
 #
 #     Example Structure:
 #        Inbox
 #        |-- Subfolder
 #            |-- Selected Contacts
 #>


<#
 # This function is intended to determine if a folder already exists in
 # in the Outlook application.
 #>
FUNCTION CheckForFolderExistance($root, $subFolderName, $_exists = $false)
{
    $exists = $_exists

    if($subFolderName -eq $root.Name)
    {
        $exists = $true
    }
    else
    {
        foreach($f in $root.Folders)
        {
            if($f.Name -eq $subFolderName)
            {
                $exists = $true
                break
            }
        }
    }
    return $exists
}

<# 
 # Function to create subfolders of a root folder object.
 #>
FUNCTION CreateSubFolders($root, $contact_list)
{
    $folder_count = 0
    $skipped_folders = 0
    foreach($item in $contact_list.GetEnumerator())
    {
        if(!(CheckForFolderExistance $root $item.Key))
        {
            $root.Folders.Add($item.Key) | Out-Null
            ++$folder_count | Out-Null
        }
        else
        {
            ++$skipped_folders
            Write-Host -ForegroundColor Yellow "[SKIPPING $($item.Key)] " -NoNewline
            Write-Host "Folder Already exists in $($root.Name)"
        }
    }
    Write-Host "`n`r-------------- Folder Synopsis --------------"
    if($folder_count -gt 0)
    {
        Write-Host -ForegroundColor Green "[STATUS] " -NoNewline 
        Write-Host "Successfully created $folder_count subfolders in $($root.Name)"
    }
    if($skipped_folders -gt 0)
    {
        Write-Host -ForegroundColor Yellow "[STATUS] " -NoNewline
        Write-Host "Skipped creation of $skipped_folders subfolders in $($root.Name)"
    }
    Write-Host "`n`r`n`r"
}

<#
 # Apply the rules to all the subfolders in the root folder
 #>
FUNCTION ApplyRules($root, $rules, $contacts, $rules_to_apply)
{
    # Get the RuleReceive rule type to set rules for messages received
    $olRuleReceive = [Microsoft.Office.Interop.Outlook.OlRuleType]::olRuleReceive
    $rules = $outlook_namespace.DefaultStore.GetRules()

    # Iterate through all the selected contacts to apply rules to
    foreach($rcp in $contacts)
    {
        $rule_name = "PS Rule: $($rcp.Name)"
        $rule = $rules.Create($rule_name, $olRuleReceive)
        $folder = $root.Folders.Item($rcp.Name)
        $rule.Conditions.From.Recipients.Add($rcp.Value) | Out-Null
        $rule.Conditions.From.Recipients.ResolveAll() | Out-Null
        $rule.Conditions.From.Enabled = $true
        # Add each rule to the specified contact
        foreach($r in $rules_to_apply.GetEnumerator())
        {   
            if($r.Key -eq "Delete On Receipt")
            {
                $action = $rule.Actions.Delete
                $action.Enabled = $true
                $rules.Save()
            }
            if($r.Key -eq "Delete Permanently")
            {
                $action = $rule.Actions.DeletePermanently
                $action.Enabled = $true
                $rules.Save()
            }
            if($r.Key -eq "Display Desktop Alert")
            {
                $action = $rule.Actions.DesktopAlert
                $action.Enabled = $true
                $rules.Save()
            }
            if($r.Key -eq "Copy To Folder")
            {
                $action = $rule.Actions.CopyToFolder
                $action.Enabled = $true
                [Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction].InvokeMember(
                "Folder",
                [System.Reflection.BindingFlags]::SetProperty,
                $null,
                $action,
                $folder)
                $rules.Save()
            }
            if($r.Key -eq "Mark As Read")
            {
                Write-Host -ForegroundColor Yellow "Marking messages originally from $($rcp.Name) as ""read"". This may take a while"
                # run through the inbox folder first and convert "unread" to "read" on any unmoved messages
                $msgCount = 0
                for($i = 1; $i -lt $root.Parent.Items.Count+1; $i++)
                {
                    $msg = $root.Parent.Items.Item($i)
                    if($msg.SenderEmailAddress -eq $rcp.Value -and $msg.UnRead -eq $true)
                    {
                        $msg.UnRead = $false
                    }
                }
                
                # Run through the specified folder and mark all as read
                for($i = 1; $i -ne $folder.Items.Count+1; $i++)
                {
                    $msg = $folder.Items.Item($i)
                    if($msg.UnRead -eq $true)
                    {
                        $msg.UnRead = $false
                    }
                }
            }
            if($r.Key -eq "Play Sound")
            {
                $action = $rule.Actions.PlaySound
                $action.FilePath = "C:\Windows\Media\chimes.wav"
                $action.Enabled = $true
                $rules.Save()
            }
            if($r.Key -eq "Move Messages")
            {
                $action = $rule.Actions.MoveToFolder
                $action.Enabled = $true
                [Microsoft.Office.Interop.Outlook._MoveOrCopyRuleAction].InvokeMember(
                "Folder",
                [System.Reflection.BindingFlags]::SetProperty,
                $null,
                $action,
                $folder)
                $rules.Save()
            }
        }
    }
}

<#
 # Function to create a dialog window and receive user input for the folder to create
 #>
FUNCTION GetFolderNameWindow()
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    $folder_name = ""

    # Create the form object
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Folder to Create"
    $form.AutoSize = $true
    $form.AutoSizeMode = "GrowAndShrink"
    $form.StartPosition = "CenterScreen"

    # Create a label for the form
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Size(10,20)
    $label.Size = New-Object System.Drawing.Size(280, 30)
    $label.AutoSize = $true
    $label.Text = "Enter the folder you would like your messages copied to below"
    $form.Controls.Add($label)

    # Create a textbox for the form
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Size(10, 60)
    $textBox.Size = New-Object System.Drawing.Size(280, 20)
    $form.Controls.Add($textBox)

    # Add an OK and Cancel button to the form
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Size(75,120)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Size(150,120)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $CancelButton

    # Set up key-bindings for ESC and Enter
    $form.KeyPreview = $true
    $form.Add_KeyDown({ 
        if($_.KeyCode -eq "Enter")
        {
            $folder_name = $textBox.Text;
            $form.Close();
        }
    })

    $form.Add_KeyDown({
        if($_.KeyCode -eq "Escape")
        {
            $form.Close()
        }
    })

    $form.TopMost = $true
    $dialogResult = $form.ShowDialog()

    if(($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) -and (-not [string]::IsNullOrWhiteSpace($textBox.Text)))
    {
        $folder_name = $textbox.Text
    }
    else
    {
        Write-Host -ForegroundColor Red "[FATAL ERROR] " -NoNewline 
        Write-Host "No folder name entered!"
        exit
    }

    $form.Dispose()

    return $folder_name
}

<#
 # Function to create a dialog window that displays all available rules to apply to the subfolders
 # that are created.
 #>
FUNCTION GetRulesWindow()
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $rules_dict = @{
        "Delete On Receipt" = [Microsoft.Office.Interop.Outlook.olRuleActionType]::olRuleActionDelete;
        "Delete Permanently" = [Microsoft.Office.Interop.Outlook.olRuleActionType]::olRuleActionDeletePermanently;
        "Display Desktop Alert" = [Microsoft.Office.Interop.Outlook.olRuleActionType]::olRuleActionDesktopAlert;
        "Copy To Folder" = [Microsoft.Office.Interop.Outlook.olRuleActionType]::olRuleActionCopyToFolder;
        "Mark As Read" = [Microsoft.Office.Interop.Outlook.olRuleActionType]::olRuleActionMarkRead;
        "Play Sound" = [Microsoft.Office.Interop.Outlook.olRuleActionType]::olRuleActionPlaySound;
        "Move Messages" = "Move Messages"
    }
    $rules_array = @()
    $selected_rules = @{}
    foreach($item in $rules_dict.Keys)
    {
        $rules_array += $item
    }

    # FORM
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Rules to Apply"
    $form.AutoSize = $true
    $form.AutoSizeMode = "GrowAndShrink"
    $form.StartPosition = "CenterScreen"
    $form.MinimizeBox = $false
    $form.MaximizeBox = $false

    # OK BUTTON
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(40,200)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton

    # CANCEL BUTTON
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(115,200)
    $cancelButton.Size = New-Object System.Drawing.Size(75,23)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton

    # LABEL
    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Available Rules (Ctrl or Shift to select multiple)"
    $label.Location = New-Object System.Drawing.Point(10, 40)
    $label.AutoSize = $true
    
    #LIST BOX
    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.Location = New-Object System.Drawing.Point(50,60)
    $listBox.AutoSize = $true
    $listBox.SelectionMode = "MultiExtended"
    foreach($r in $rules_array)
    {
        [void] $listBox.Items.Add($r)
    }

    # ADD CONTROLS
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
    $form.Controls.Add($label)
    $form.Controls.Add($listBox)

    $form.Topmost = $true
    $form.Add_Shown({ $form.Activate() })
    $dialogResult = $form.ShowDialog()

    # SET RULES SELECTED FOR RETURN
    if(($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) -and ($listBox.SelectedItems.Count -gt 0))
    {
        foreach($item in $listBox.SelectedItems)
        {
            $selected_rules.Add($item, $rules_dict.$item)
        }
    }
    else
    {
        Write-Host -ForegroundColor Red "[FATAL ERROR] " -NoNewline
        Write-Host "No Rules Selected!"
        exit
    }

    $form.Dispose()

    return $selected_rules
}

########################################################
#                                                      #
#        BEGIN OBJECTS AND VARIABLES SECTION           #
#                                                      #
########################################################
# Add the Microsoft COM object assembly for Outlook
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
$outlook = New-Object -ComObject Outlook.Application
$outlook_namespace = $outlook.GetNameSpace("MAPI")
$ruleName = "PS Rule: Copy To Folder - "
$olFolderInbox = [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox
$olFolderContacts = [Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderContacts
$contacts = $outlook_namespace.GetDefaultFolder($olFolderContacts).Items
$folderName = GetFolderNameWindow
# Hash table for the selected contacts
$contact_hash = @{}
########################################################
#                                                      #
#        END OBJECTS AND VARIABLES SECTION             #
#                                                      #
########################################################

# Add each contact and contacts address to a hash table
foreach($c in $contacts)
{
    $contact_hash.Add($c.FullName, $c.IMAddress)
}
# Display in a window for user selection, store in "selected_contacts"
$contact_hash.GetEnumerator() | 
    Sort-Object -Property Name | 
    Out-GridView -Title "Select Outlook contacts" -PassThru -OutVariable selected_contacts | 
    Out-Null
if($selected_contacts.Count -eq 0)
{
    Write-Host -ForegroundColor Red "[FATAL ERROR] " -NoNewline
    Write-Host "Must Select Atleast 1 Contact!"
    exit
}

# Get the rules to apply to the new folder
$apply_rules = GetRulesWindow

# Get the location of the inbox
$inbox = $outlook_namespace.GetDefaultFolder($olFolderInbox)

# Check if the destination folder exists
$folder_exists = $false
foreach($fold in $inbox.Folders)
{
    $folder_exists = CheckForFolderExistance $fold $folderName
    if($folder_exists)
    {
        Write-Host -ForegroundColor Yellow "[$folderName EXISTS] " -NoNewline
        Write-Host "Creating subfolders."
        $subFolder = $inbox.Folders.Item($folderName)
        CreateSubFolders $subFolder $selected_contacts
        break
    }
}
# Create the inbox subfolder if it does not exist
if(!$folder_exists)
{
    Write-Host -ForegroundColor Yellow "[Creating folder] " -NoNewline
    Write-Host "`"$folderName`""
    # Suppress the output of adding the folder
    $inbox.Folders.Add($folderName) | Out-Null
    Write-Host -ForegroundColor Green "`"$folderName`" created successfully"
    $subFolder = $inbox.Folders.Item("$folderName")
    CreateSubFolders $subFolder $selected_contacts
}

# Set the root folder for the ApplyRules function to be the inbox subfolder created earlier
$root = $inbox.Folders.Item($folderName)
Write-Host -ForegroundColor Yellow "[STATUS] " -NoNewline 
Write-Host "Applying Outlook Rules... May take some time"
ApplyRules $root $rules $selected_contacts $apply_rules
Write-Host -ForegroundColor Green "-------- Sucessfully Completed Creating Subfolders and Applying Outlook Rules --------"