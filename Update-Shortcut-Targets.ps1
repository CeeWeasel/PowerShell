Function Update-Shortcut-Targets
{
    <#
        .SYNOPSIS
            Changes the target of the specified shortcut (*.lnk) file(s).

        .DESCRIPTION
            Changes the target of specified shortcuts.  The shortcuts are changed
            per specified user on each specified computer.  Only shortcut files (*.lnk)
            are changed.

        .PARAMETER User
            The user profile from where to check for *.lnk files
            If unspecified, the current user's profile will be checked for .lnk files.

        .PARAMETER AllUsers
            If specified, all local profiles will be checked for .lnk files.

        .PARAMETER ProfileDirectory
            The folder within the specified user's profile that will be searched for shortcuts.

        .PARAMETER OldTarget
            Shortcuts with targets meeting this requirement will be changed.  Can be a directory, file, or string.
            To reference a directory within a user's profile, use '~Profile~' (e.g. '~Profile~'\Desktop)
            To reference the system drive, use '~System~' (e.g. '~System~'\Windows' )
            To reference a specific drive, use '~Drive Letter~' (e.g. '~C~\Program Files' )
            
        .PARAMETER NewTarget
            How to change shortcut targets.  Can be a directory, file, or string.

        .PARAMETER ComputerName
            The name of the computer to search for shortcuts.  If not specified, local accounts on the local computer will be searched.

        .EXAMPLE
            Update-Shortcut-Targets

            Description
            -----------
            Returns information about the .lnk files found in the logged in users'
            desktop directory.

            Update-Shortcut-Targets -User JohnSmith

            Description
            -----------
            Returns information about the .lnk files found in the local profile
            for the user JohnSmith

            Update-Shortcut-Targets -AllUsers

            Description
            -----------
            Returns information about the .lnk files found in all local profiles.

        .NOTES
            FunctionName : Update-Shortcut-Targets
            Created by   : weacli
            Date Coded   : 01/22/2014 16:46:54

        .LINK
            https://github.com/CeeWeasel/PowerShell/wiki/Update-Shortcut-Targets

        .TODO
            Add functionality for detecting files and strings in both $OldTarget and $NewTarget
            Add functionality for checking remote computers
            Clean-up the code
    #>

    [CmdletBinding()]
    Param
    (
        [Parameter(
            ParameterSetName="User",
            HelpMessage="The username(s) of the accounts local to the specified computer whose shortcuts should be changed.  If no user is specified and -AllUsers is not used, the account running the script will be used.",
            Mandatory=$false
        )]
        [array]$UserName = $env:USERNAME,

        [Parameter(
            ParameterSetName="AllUsers",
            HelpMessage="Changes the shortcut for users local to the specified computer.  If no user is specified and -AllUsers is not used, the account running the script will be used.",
            Mandatory=$false
        )]
        [switch]$AllUsers,

        [Parameter(
            HelpMessage="The folder within the specified user's profile that should be used to search for shortcuts to change.  If not used, all shortcuts within all folders will searched.",
            Mandatory=$false
        )]
        [string]$ProfileDirectory = "",

        [Parameter(
            HelpMessage="Shortcuts with targets meeting this requirement will be changed.  Can be a directory, file, or string.",
            Mandatory=$true
        )]
        [string]$OldTarget=$Null,

        [Parameter(
            HelpMessage="How to change shortcut targets.  Can be a directory, file, or string.",
            Mandatory=$true
        )]
        [string]$NewTarget,

        [Parameter(
            ValueFromPipeline=$True,
            ValueFromPipelineByPropertyName=$True,
            HelpMessage="The name of the computer to search for shortcuts.  If not specified, local accounts on the local computer will be searched.",
            Mandatory=$false
        )]
        [array]$ComputerName = $env:COMPUTERNAME,

        [Parameter(
            HelpMessage="If used, a backup of the shortcuts will be made.",
            Mandatory=$false
        )]
        [switch]$Backup

    )

    Begin
    {
        $Links = @()
        if ( ( Test-Path $OldTarget ) ) { $OldTargetLinks = get-childitem -filter *.lnk $OldTarget } else { $OldTarget = $Null }
        if ( ( Test-Path $NewTarget ) ) { $NewTargetLinks = get-childitem -filter *.lnk $NewTarget } else { $NewTarget = $Null }
        $shell = new-object -com wscript.shell
    }

    Process
    {
    $ErrorActionPreference = "Stop"
        foreach ( $Computer in $ComputerName )
        {
            $ProfileParent = $Null
            if ( $Computer -like "*\*" ) { $Computer = ($Computer.split("\"))[1] } # Try to catch computer names that include the domain. ( e.g. domain\computername )
            if ( $Computer -eq $env:COMPUTERNAME )
            {
                $XPTest = $env:SystemDrive + "\Documents and Settings"
                $7Test = $env:SystemDrive + "\Users"
                if ( Test-Path $XPTest ) { $ProfileParent = $XPTest }
                if ( Test-Path $7Test ) { $ProfileParent = $7Test }
            }
            try
            {
                Invoke-Command -ComputerName $Computer -ScriptBlock {
                    $XPTest = "\\" + $Computer + "\"  + $env:SystemDrive[0] + "\Documents and Settings"
                    $7Test = "\\" + $Computer + "\"  + $env:SystemDrive[0] + "\Users"
                    if ( Test-Path $XPTest ) { $ProfileParent = $XPTest ; $SystemDrive = $env:SystemDrive }
                    if ( Test-Path $7Test ) { $ProfileParent = $7Test ; $SystemDrive = $env:SystemDrive }
                }
            }
            catch
            { $error[0].Message }
                
            if ( ! $ProfileParent )
            {
                try
                {
                    $DrivesToTry = "c","d","e","f" | ForEach-Object {
                        $XPTest = "\\" + $Computer + "\" + $_ + "$\Documents and Settings"
                        $7Test = "\\" + $Computer + "\" + $_ + "$\Users"
                        if ( Test-Path $XPTest ) { $ProfileParent = $XPTest ; $SystemDrive = $_ }
                        if ( Test-Path $7Test ) { $ProfileParent = $7Test ; $SystemDrive = $_ }
                    }
                }
                catch { $error[0].Message }
            }
            Write-Host "ProfileParent on $Computer is $ProfileParent"

            <#
            At this point, we have tried to get the multiple ways to to get the $ProfileParent
            I'm at the point where if we can't get it, then the computer is either offline,
            we don't have invoke access on the computer, or UNC access.

            From here, if we have a $ProfileParent, then we are good to go.
            #>

            if ( $ProfileParent )
            {
                if ( $AllUsers )
                {
                    $UserName = $Null
                    $UserName = Get-ChildItem -Directory $ProfileParent
                }
                foreach ( $Account in $UserName )
                {
                    $SearchResults = @()
                    if ( $Account -like "*\*" ) { $Account = ($Account.split("\"))[1] } # Strip the domain from usernames
                    $ShortcutDirectory = $ProfileParent + "\" + $Account # e.g. c:\users\jimsmith
                    if ( $ProfileDirectory ) { $ShortcutDirectory += "\" + $ProfileDirectory } # E.G. c:\users\jimsmith\desktop
                    Write-Host "Checking $ShortcutDirectory"
                    if ( Test-Path $ShortcutDirectory ) { $SearchResults += Get-ChildItem -filter *.lnk $ShortcutDirectory -Recurse }
                    Write-Host "$SearchResults"
                    foreach ( $Result in $SearchResults )
                    {
                        ###
                    }
                }
            }
        }
        Return "Stopping for now"
<# This was my original way of doing this.
        $Win7Style = ( Test-Path "$env:SystemDrive\Users" )
        #region Original Code
        If ( $Win7Style ) { $UsersDir = "$env:SystemDrive\Users" }
        else { $UsersDir = "$env:SystemDrive\Documents and Settings" }
        
        If ( $UserName )
        {
            ForEach ( $UserName  in $UserName )
            {
                $DesktopDirs += Get-ChildItem "$UsersDir\$UserName\Desktop"
            }
        }

        If ( $AllUsers ) { $DesktopDirs = Get-ChildItem $UsersDir }

        ForEach ( $Directory in $DesktopDirs )
        {
            $Links = @()
            $Links = Get-ChildItem -filter *.lnk $Directory.FullName  -recurse
            If ( $Links )
            {
                ForEach ( $Link in $Links )
                {
                    $templink = $shell.createShortcut($Link.fullname)
                    $templinkname = ($templink.FullName).split("`\")
                    $templinkname = $templinkname[$templinkname.length - 1]
                    $OldTargetLink = $Null
                    If ( $OldTargetLinks )
                    {
                        ForEach ( $OldTargetLink in $OldTargetLinks )
                        {
                            $oldlink = $shell.createShortcut($OldTargetLink.FullName)
                            $oldlinkname = ($oldlink.FullName).split("`\")
                            $oldlinkname = $oldlinkname[$oldlinkname.length - 1]
                            if ( $templink.TargetPath -eq $oldlink.TargetPath )
                            {
                                ForEach ( $NewTargetLink in $NewTargetLinks )
                                {
                                    $newlink = $shell.CreateShortcut($NewTargetLink.FullName)
                                    $newlinkname = ($newlink.FullName).split("`\")
                                    $newlinkname = $newlinkname[$newlinkname.length - 1]
                                    if ( $oldlinkname -eq $newlinkname )
                                    {
                                        if ( $Backup )
                                        {
                                            $BackupPath = $Link.DirectoryName + "\Backup"
                                            If ( ! (Test-Path $BackupPath) )
                                            {
                                                New-Item $BackupPath -ItemType directory | Out-Null
                                            }
                                            $BackupName = $BackupPath + "\" + $Link.Name
                                            Copy-Item $Link.FullName $BackupName -Force
                                    
                                        }
                                        Remove-Item $Link.FullName
                                        Copy-Item $NewTargetLink.FullName $Link.FullName
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        #endregion
#>
    }
    End {}
}
Update-Shortcut-Targets -OldTarget "C:\Test" -NewTarget "C:\Test" -ComputerName "AGHXEN1","AGHXEN30","RAPIDCOMM","PC0110","AGHRDP02" -AllUsers -ProfileDirectory Desktop