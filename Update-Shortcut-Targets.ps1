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
            https://github.com/CeeWease/PowerShell/wiki/Update-Shortcut-Targets

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
        [string]$ComputerName = $env:COMPUTERNAME,

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
        <#
        foreach ( $Computer in $ComputerName )
        {
            
        }
        #>

        $Win7Style = ( Test-Path "$env:SystemDrive\Users" )
        #region Original Code
        If ( $Win7Style ) { $UsersDir = "$env:SystemDrive\Users" ; Write-Log -Message "Using Win7 style profile paths" -Path C:\MyLog.txt -EventLogName 'Application'}
        else { $UsersDir = "$env:SystemDrive\Documents and Settings" ; Write-Log -Message "Using WinXP style profile paths" -Path C:\MyLog.txt -EventLogName 'Application'}
        Write-Log -Message "Users directory set to $UsersDir" -Path C:\MyLog.txt -EventLogName 'Application' -Level Info
        
        If ( $UserName )
        {
            ForEach ( $UserName  in $UserName )
            {
                $DesktopDirs += Get-ChildItem "$UsersDir\$UserName\Desktop"
            }
        }
        else
        {
            $DesktopDirs = Get-ChildItem "$UsersDir\$env:UserName\Desktop"
        }
        If ( $AllUsers ) { $DesktopDirs = Get-ChildItem $UsersDir }

        ForEach ( $Directory in $DesktopDirs )
        {
            Write-Log -Message "Gathering links from $Directory" -Path C:\MyLog.txt -EventLogName 'Application' -Level Info
            $Links = @()
            $Links = Get-ChildItem -filter *.lnk $Directory.FullName  -recurse
            Write-Log -Message "Updating desktop shortcuts for $Directory, using $OldTarget and $NewTarget" -EventLogName 'Application' -Level Info
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
    }
    End {}
}