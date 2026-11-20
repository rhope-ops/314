<#
FakePnP.ps1 - stand-ins for the PnP.PowerShell cmdlets the scan calls, so the tests need no network.

  Purpose:     Defines global functions with the names of the real PnP cmdlets. Module code that calls
               those names reaches these fakes. $global:FakePnP holds the fake site and switches.
  Inputs:      $global:FakePnP (set by Reset-FakePnP, changed by the tests).
  Outputs:     fake site data; $global:FakePnP.Calls lists each cmdlet called.
  How to run:  dot-sourced by Run-Tests.ps1; not on its own.
  Run target:  PowerShell 7 (CI and dev machine); written to also parse on Windows PowerShell 5.1.
  Network:     none.
  Writes:      none.
  Credentials: none (the fake sign-ins accept nothing and store nothing).

  Why global:  module code finds a cmdlet by name in the global scope, so the fakes are global functions,
               and the state they share ($global:FakePnP) is global too. Tests change it only through
               Set-FakePnPOption and read it through Test-FakePnPCall.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'Fakes of global cmdlets share global state.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Read by the fakes.')]
param()

function Reset-FakePnP {
    <#
    .SYNOPSIS
      Sets up the fake site: library '/sites/TeamA/Shared Documents', channels General and Ops, the
      given files, file download through -AsMemoryStream.
    .PARAMETER Files
      Server-relative file path -> file bytes.
    .OUTPUTS
      None (sets $global:FakePnP).
    #>
    param([hashtable]$Files)
    $global:FakePnP = @{
        Files       = $Files
        Folders     = @('General', 'Ops', 'Forms')
        ListFails   = $false
        ConnectFails = $false                                     # the web login to a site fails (an error before the scan)
        SiteFails   = $false                                      # Get-PnPSite fails (the site's group cannot be read)
        TypedSite   = 'https://tenant.example.com/teams/TeamA'   # what the fake Read-Host returns
        SiteGroupId = [guid]'11111111-2222-3333-4444-555555555555'
        Calls       = New-Object System.Collections.Generic.List[string]
        Progress    = New-Object System.Collections.Generic.List[string]
    }
}

function Set-FakePnPOption {
    # Changes one switch of the fake site: ListFails, ConnectFails or SiteFails.
    param([ValidateSet('ListFails', 'ConnectFails', 'SiteFails')][string]$Name, [bool]$Value)
    $global:FakePnP[$Name] = $Value
}

function Set-FakeFolders {
    # Sets the top-level folders of the fake library (Get-PnPFolderItem).
    param([string[]]$Names)
    $global:FakePnP.Folders = $Names
}

function Set-FakeSiteAnswer {
    <#
    .SYNOPSIS
      Sets what Read-Host returns (the typed site address) and the group ID of the connected site.
    .PARAMETER TypedSite
      What the fake Read-Host returns; 'cancel' = the prompt fails (stands in for a cancel).
    .PARAMETER SiteGroupId
      The group ID of the connected site (Get-PnPSite).
    .OUTPUTS
      None (changes $global:FakePnP).
    #>
    param([string]$TypedSite, [guid]$SiteGroupId = '11111111-2222-3333-4444-555555555555')
    $global:FakePnP.TypedSite = $TypedSite
    $global:FakePnP.SiteGroupId = $SiteGroupId
}

function Get-FakePnPConnection {
    # Every site the fake Connect-PnPOnline was called with since the last Reset-FakePnP, in order.
    return @($global:FakePnP.Calls | Where-Object { $_ -like 'Connect *' } | ForEach-Object { $_.Substring(8) })
}

function Get-FakeProgressText {
    # Everything the progress bar showed since the last Reset-FakePnP, as one string.
    return ($global:FakePnP.Progress -join ' | ')
}

function Get-FakeCallCount {
    # How many times the named fake call was made since the last Reset-FakePnP.
    param([string]$Name)
    return @($global:FakePnP.Calls | Where-Object { $_ -eq $Name }).Count
}

function Test-FakePnPCall {
    # True when the named fake cmdlet was called since the last Reset-FakePnP.
    param([string]$Name)
    return $global:FakePnP.Calls.Contains($Name)
}

# Read the site: library, its root folder, its channel folders and its items.
function global:Get-PnPList {
    # Fake: the library exists only under its URL name.
    [CmdletBinding()] param($Identity)
    $global:FakePnP.Calls.Add('Get-PnPList')
    if ($Identity -eq 'Shared Documents') { [pscustomobject]@{ Title = 'Documents' } }
}

function global:Get-PnPProperty {
    # Fake: the library's root folder.
    [CmdletBinding()] param($ClientObject, $Property)
    [pscustomobject]@{ ServerRelativeUrl = '/sites/TeamA/Shared Documents' }
}

function global:Get-PnPFolderItem {
    # Fake: the top-level folders of the library.
    [CmdletBinding()] param($FolderSiteRelativeUrl, $ItemType)
    $global:FakePnP.Folders | ForEach-Object { [pscustomobject]@{ Name = $_ } }
}

function global:Get-PnPListItem {
    # Fake: one folder item plus one item per fake file; throws when ListFails is set.
    [CmdletBinding()] param($List, $PageSize, $Fields)
    if ($global:FakePnP.ListFails) { throw 'FAKE: list read refused' }
    @{ FileLeafRef = 'General'; FileRef = '/sites/TeamA/Shared Documents/General'; FSObjType = 1 }
    foreach ($path in $global:FakePnP.Files.Keys) {
        @{ FileLeafRef = ($path -split '/')[-1]; FileRef = $path; FSObjType = 0
           File_x0020_Size = [string]$global:FakePnP.Files[$path].Length; Modified = [datetime]'2026-11-01' }
    }
}

function global:Get-PnPFile {
    # Fake: the file bytes as a memory stream (a later PnP version's -AsMemoryStream).
    [CmdletBinding()] param($Url, [switch]$AsMemoryStream)
    New-Object System.IO.MemoryStream(, $global:FakePnP.Files[$Url])
}

function Remove-FakeMemoryStreamSwitch {
    # Redefines the fake Get-PnPFile without -AsMemoryStream, as in PnP versions that lack it.
    function global:Get-PnPFile { [CmdletBinding()] param($Url) throw 'FAKE: should not be called' }
}

function global:Get-PnPContext {
    <#
    .SYNOPSIS
      Fake client context for the OpenBinaryStream path: Web.GetFileByServerRelativeUrl(url)
      .OpenBinaryStream() gives an object whose Value is the file stream; ExecuteQuery does nothing.
    .OUTPUTS
      [pscustomobject]
    #>
    [CmdletBinding()] param()
    $global:FakePnP.Calls.Add('Get-PnPContext')
    $web = [pscustomobject]@{}
    $web | Add-Member -MemberType ScriptMethod -Name GetFileByServerRelativeUrl -Value {
        param($url)
        $file = [pscustomobject]@{ Url = $url }
        $file | Add-Member -MemberType ScriptMethod -Name OpenBinaryStream -Value {
            [pscustomobject]@{ Value = (New-Object System.IO.MemoryStream(, $global:FakePnP.Files[$this.Url])) }
        }
        return $file
    }
    $context = [pscustomobject]@{ Web = $web }
    $context | Add-Member -MemberType ScriptMethod -Name ExecuteQuery -Value { }
    return $context
}

# The progress bar is on screen too: the fake records its text so tests can check it is masked.
function global:Write-Progress {
    # Fake: records the activity and status text instead of drawing a bar.
    [CmdletBinding()] param($Activity, $Status, $PercentComplete, [switch]$Completed, $Id, $ParentId)
    $global:FakePnP.Progress.Add(('{0} {1}' -f $Activity, $Status))
    if ($Completed) { $global:FakePnP.Calls.Add('Progress done ' + $Id) }
    elseif ($null -ne $ParentId) { $global:FakePnP.Calls.Add('Progress ' + $Id + ' under ' + $ParentId) }
}

# Sign-in, sign-out and the team behind the site.
function global:Connect-PnPOnline {
    # Fake: the web login works unless ConnectFails is set.
    [CmdletBinding()] param($Url, [switch]$UseWebLogin)
    $global:FakePnP.Calls.Add('Connect ' + $Url)
    if ($PSBoundParameters['WarningAction'] -eq 'SilentlyContinue') { $global:FakePnP.Calls.Add('Quiet sign-in ' + $Url) }
    if ($UseWebLogin -and $global:FakePnP.ConnectFails) { throw 'FAKE: sign-in refused' }
}

function global:Disconnect-PnPOnline {
    # Fake: nothing to close; records the call.
    [CmdletBinding()] param()
    $global:FakePnP.Calls.Add('Disconnect-PnPOnline')
}

function global:Get-PnPSite {
    # Fake: the site belongs to a team (a Microsoft 365 group).
    [CmdletBinding()] param($Includes)
    if ($global:FakePnP.SiteFails) { throw 'FAKE: site refused' }
    [pscustomobject]@{ GroupId = $global:FakePnP.SiteGroupId }
}

function global:Read-Host {
    # Fake: the address the person would type.
    [CmdletBinding()] param($Prompt)
    $global:FakePnP.Calls.Add('Read-Host')
    # A real cancel stops the whole pipeline (the test run too); an error leaves through the same 'finally'.
    if ($global:FakePnP.TypedSite -eq 'cancel') { throw 'FAKE: prompt cancelled' }
    return $global:FakePnP.TypedSite
}

