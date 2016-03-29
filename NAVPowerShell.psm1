$script:CurrentDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SettingsFile = Join-Path $script:CurrentDirectory 'settings.config'
$global:JSATestModuleConfig = Get-Item $script:SettingsFile

[xml]$script:xml = Get-Content $script:SettingsFile
$script:watchFolder = $script:xml.settings.watchfolder
$script:backupfolder = $script:xml.settings.backupfolder
$script:LicenseFileLocation = $script:xml.settings.LicenseFileLocation
if(!(Test-Path $script:LicenseFileLocation -PathType Leaf)){
  Write-Warning "License file not fount at: ${script:LicenseFileLocation}"
}

function Compare-Files
{
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory)]
    [String]$Left,
    [Parameter(Mandatory)]
    [String]$Right
    )
  Process
  {
    if((Test-path $Left) -and (Test-path $Right))
    {
      $Leftpath = Get-Item $Left
      $Leftpath = $Leftpath.FullName
      $RightPath = Get-Item $Right
      $RightPath = $RightPath.FullName
      $VisualStudio = Join-Path $env:VS120COMNTOOLS '..\IDE\devenv.exe'
      Start-Process -FilePath $VisualStudio -ArgumentList "/Diff `"$Leftpath`" `"$RightPath`""
    }
  }
}

function Split-LatestNAVObjectFile
{
  [CmdletBinding()]
  Param
  (
    [Parameter(Mandatory)]
    [String]$DestinationPath
    )
  Process
  {
    $file = (Get-ChildItem -Path $script:backupfolder -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
    Write-Host "File found:`n$file`nPress enter to split to $DestinationPath"
    $keypress = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if ($keypress.VirtualKeyCode -eq 13)
    {
      Split-NAVObjectFile -SourcePath $file -DestinationPath $DestinationPath
    }
  }
}

function Start-BackupScript
{
  [CmdletBinding()]
  Param
  (
    [Parameter()]
    [String]$WatchFolder = $script:watchFolder,
    [Parameter()]
    [String]$DestinationFolder = $script:backupfolder
    )
  Process 
  {
    $filter = '*.*'                           
    $fsw = New-Object IO.FileSystemWatcher $WatchFolder, $filter -Property @{IncludeSubdirectories = $false;NotifyFilter = [IO.NotifyFilters]'FileName, LastWrite'} 
    $action = {
      $fileMissing = $false 
      $FileInUseMessage = $false 
      $copied = $false 
      $file = Get-Item $Args.FullPath 
      $dateString = Get-Date -format "_yyyy-MM-dd_HH-mm-ss" 
      $DestinationFolder = $event.MessageData 
      $DestinationFileName = $file.basename + $dateString + $file.extension 
      $resultfilename = Join-Path $DestinationFolder $DestinationFileName 
      Write-Output ""
      while(!$copied) { 
        try { 
          Move-Item -Path $file.FullName -Destination $resultfilename -ErrorAction Stop -ErrorVariable notused
          $copied = $true 
        }  
        catch [System.IO.IOException] { 
          if(!$FileInUseMessage) { 
            Write-Output "$(Get-Date -Format "yyyy-MM-dd @ HH:mm:ss") - $file in use. Waiting to move file"
            $FileInUseMessage = $true 
          } 
          Start-Sleep -s 1 
        }  
        catch [System.Management.Automation.ItemNotFoundException] { 
          $fileMissing = $true 
          $copied = $true 
        } 
      } 
      if($fileMissing) { 
        Write-Output "$(Get-Date -Format "yyyy-MM-dd @ HH:mm:ss") - $file not found!"
        } else { 
          Write-Output "$(Get-Date -Format "yyyy-MM-dd @ HH:mm:ss") - Moved $file to backup! `n`tFilename: `"$resultfilename`""
        }
      }
      $global:backupscript = Register-ObjectEvent -InputObject $fsw -EventName "Created" -Action $action -MessageData $DestinationFolder
      Write-Host "Started. WatchFolder: `"$($WatchFolder)`" DestinationFolder: `"$($DestinationFolder)`". Job is in: `$backupscript"
    }
  }

  function Get-BackupScript
  {
    [CmdletBinding()]
    Param
    (
      [Parameter()]
      [String]$BackupScriptJob = $global:backupscript,
      [Switch]$WithWarnings
      )
    Process 
    {
      if ($WithWarnings) {
        Receive-Job -Job $BackupScriptJob -Keep
        } else {
          Receive-Job -Job $BackupScriptJob -Keep -ErrorAction SilentlyContinue    
        }
        Write-Host "`n"
      }
    }

    function Split-NAVObjectFile
    {
      [CmdletBinding()]
      Param
      (
        [Parameter(Mandatory)]
        [String]$SourcePath,
        [Parameter(Mandatory)]
        [String]$DestinationPath
        )   
      Process
      {
        $index = 0
        $SourceFolder = (Get-Item $SourcePath).Directory.FullName
        $TempFolderName = [System.Guid]::NewGuid().ToString()
        $SplitPath = Join-Path $env:TEMP $TempFolderName
        if(!(Test-Path $SplitPath))
        {
          New-Item $SplitPath -type directory | Out-null
        }
        Split-NAVApplicationObjectFile -Source $SourcePath -Destination $SplitPath -Force -PreserveFormatting
        $files = Get-ChildItem -Path $SplitPath -Filter "*.txt"
        Write-Progress -Activity "Moving files to $DestinationPath" -Id 1 -ParentId -1 -PercentComplete 0
        foreach($file in $files)
        {
          $index = $index + 1
          Move-NavApplicationObjectFile -SourcePath $file.FullName -DestinationFolder $DestinationPath
          $percentcomplete = [math]::round($index / $files.Length * 100)
          Write-Progress -Activity "Moving files to $DestinationPath" -CurrentOperation "File $index of $($files.Length): $percentcomplete %" -Id 1 -ParentId -1 -PercentComplete $percentcomplete
        }
        Remove-Item $SplitPath -Force
        Write-Progress -Activity "Moving files to $DestinationPath" -Id 1 -Completed
      }
    }

    function Move-NavApplicationObjectFile
    {
      [CmdletBinding()]
      Param
      (
        [Parameter(Mandatory)]
        [String]$SourcePath,
        [Parameter(Mandatory)]
        [String]$DestinationFolder,
        [Parameter()]
        [String]$Version,
        [Parameter()]
        [switch]$Modified
        )
      Process
      {
        if($Version -or $Modified) {
          $navobjectinfo = Get-NAVApplicationObjectProperty -Source $SourcePath -WarningAction SilentlyContinue
        }
        if($navobjectinfo)
        {
          $type = $navobjectinfo.ObjectType.ToString()
          $id = $navobjectinfo.id.ToString()
          if ($Modified -and !($local:navobjectinfo.Modified))
          {
            Write-Warning "$type $id not modified"
          }
          if ( $Version -and !($local:navobjectinfo.VersionList -like $Version))
          {
            Write-Warning "$type $id doesnt contain $Version in VersionList!"
          }
        }
        else 
        {
          $SourceFile = Get-Item $SourcePath
          switch($SourceFile.BaseName.Substring(0, 3))
          {
            "COD" {$type = "Codeunit"}
            "MEN" {$type = "Menusuite"}
            "PAG" {$type = "Page"}
            "QUE" {$type = "Query"}
            "REP" {$type = "Report"}
            "TAB" {$type = "Table"}
            "XML" {$type = "XMLport"}
          }
          $id = $SourceFile.BaseName.Substring(3)
        }
        Write-Verbose "Current object: $type - $id"
        $stringID = $id.PadLeft(10, '0')
        $filename = Join-Path $DestinationFolder (Join-Path $type "$type`_$stringID.txt")
        if(!(Test-path -Path (Split-Path $filename)))
        {
          New-Item (Split-Path $filename) -type directory | Out-Null
        }
        Move-Item -Path $SourcePath -Destination $filename -Force
      }
    }

    function Set-NAVManagement
    {
      [CmdletBinding()]
      Param
      (
        [Parameter(Mandatory)]
        [String]$ServiceInstance
        )
      Process
      {
        $service = gwmi win32_service | ? {$_.Name -eq "MicrosoftDynamicsNavServer`$$ServiceInstance"}
        if(!$service){
          Throw "ServiceInstance: '$ServiceInstance' not found"
        }
        $SearchString = "Microsoft.Dynamics.Nav.Server.exe"
        $Executeable = $service.PathName.SubString(1, $service.PathName.IndexOf($SearchString) + $SearchString.Length - 1)
        Select-NAVManagement -Version (Get-Item $Executeable).VersionInfo.ProductVersion
      }
    }

    function Select-NAVManagement
    {
      [CmdletBinding()]
      Param
      (
        [Parameter()]
        [String]$Version,
        [Parameter()]
        [Switch]$Default
        )
      Process
      {
        if($Version){
          $versionNode = ($script:xml.settings.navversions.navversion | ? {$_.version -eq $Version})
          if(!$versionNode){
            Write-Warning "Version $Version not found in settingsfile. Setting to Default"
            $Default = $true
          }
        }
        if($Default){
          $versionNode = ($script:xml.settings.navversions.navversion.servicepath | ? {$_.default})
        } 
        $servicePath = $versionNode.servicepath.InnerText
        $management = Join-Path $servicePath 'Microsoft.Dynamics.Nav.Management.dll'
        $importModule = "Import-Module -Name '$management'"
        Start-Process powershell -ArgumentList "-NoExit","-Command $importModule" -Verb RunAs
      }
    }

    function Select-NAVModel
    {
      [CmdletBinding()]
      Param
      (
        [Parameter()]
        [String]$Version,
        [Parameter()]
        [Switch]$Default
        )
      Process
      {
        if($Version){
          $clientPath = ($script:xml.settings.navversions.navversion | ? {$_.version -eq $Version}).clientpath
          if(!$clientPath){
            Write-Warning "Version $Version not found in settingsfile. Setting to Default"
            $Default = $true
          }
        }
        if($Default){
          $clientPath = ($script:xml.settings.navversions.navversion.clientpath | ? {$_.default}).InnerText
        } 
        $module = Join-Path $clientPath 'Microsoft.Dynamics.Nav.Model.Tools.dll'
        Import-Module -Name "$module" -Global
        $module = Get-Module 'Microsoft.Dynamics.Nav.Model.Tools'
        Write-Host "$($module.Name) version $($module.Version) loaded"
      }
    }

    function Get-NAVVersionNode
    {
      [CmdletBinding()]
      Param
      (
        [Parameter(Mandatory)]
        [String]$Version,
        [Parameter(Mandatory)]
        [String]$ServicePath,
        [Parameter(Mandatory)]
        [String]$ClientPath
        )
      Process
      {
        $navversion = $script:xml.CreateElement("navversion")
        $VersionElement = $script:xml.CreateElement("version")
        $VersionElement.InnerText = "$Version"
        $navversion.AppendChild($VersionElement)
        $ServicePathElement = $script:xml.CreateElement("servicepath")
        $ServicePathElement.InnerText = "$ServicePath"
        $navversion.AppendChild($ServicePathElement)
        $ClientPath = $script:xml.CreateElement("clientpath")
        $ClientPath.InnerText = "$ClientPath"
        $navversion.AppendChild($ClientPath)
        return $navversion
      }
    }

    function Search-NAVVersionNode
    {
      [CmdletBinding()]
      Param
      (
        [Parameter()]
        [String]$ServicePath = $env:ProgramFiles,
        [Parameter()]
        [String]$ClientPath = ${env:ProgramFiles(x86)},
        [Parameter()]
        [Switch]$AutoAdd
        )
      Process
      {
        throw "still todo."
        $Services = Get-ChildItem -Path $ServicePath -Filter "Microsoft.Dynamics.Nav.Server.exe" -Recurse
        $ServiceArray = $Services.VersionInfo.ProductVersion
        $Clients = Get-ChildItem -Path $ClientPath -Filter "Microsoft.Dynamics.Nav.Client.exe" -Recurse
        $ClientArray = $Clients.VersionInfo.ProductVersion
        $Uniques = $ServiceArray + $ClientArray | Select -uniq
        if($ServiceArray.Contains($Uniques)){
          Write-Warning "Service contains a versions that clients doesn't"
        }
        if($ClientArray.Contains($Uniques)){
          Write-Warning "Client contains a versions that service doesn't"
        }
        if(!$script:xml.settings.navversions.navversion.version -eq $ServiceArray){}
        return $navversion
      }
    }

    Write-Host "$($MyInvocation.MyCommand) Loaded!"
    #Select-NAVModel -Default
    Write-Warning "Using hardcoded module in this module!!"
    Import-Module "C:\Program Files (x86)\Microsoft Dynamics NAV\90\RoleTailored Client\Microsoft.Dynamics.Nav.Model.Tools.dll"

    Export-ModuleMember -Function * -Cmdlet *