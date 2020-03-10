<#
.SYNOPSIS
    Gathers CentreStack Server and Client information useful for support troubleshooting 
.DESCRIPTION
	Gathers CentreStack Server Client information useful for support troubleshooting
.EXAMPLE
    .\Get-CSInfo.ps1 -Ticket 8675309
.PARAMETER Ticket
    The CentreStack support ticket number
.NOTES 
    Author: Jeff Reed
    Name: Get-CSInfo.ps1
    Created: 2019-08-08
    
    Version History
    2019-08-09  1.0.0   Initial version
    2019-08-12  1.0.1   Move Copy to S3 function to external script
    2019-08-13  1.0.2   Renamed Copy-CSSupport.ps1 script
    2019-08-13  1.1.0   Renamed Get-CSInfo.ps1 script to support both Server and Clients. Added DebugView download and execute.
    2019-08-13  1.1.1   Fixed detection of Windows Client vs Server Agent. Refactored Get-WebFile function.
    2019-08-13  1.1.2   Tweaked Get-WebFile to handle file sizes in excess of 4 GB.
    2019-08-13  1.1.3   Change name of uploaded files based on Server, Server Agent, or Windows Client
    2019-08-13  1.1.4   Changed parameter name for Copy-CSSupport.ps1 script
    2019-08-30  1.1.5   Added InfoOnly switch argument. Changed Ticket parameter to uint32
    2019-09-05  1.1.6   Report .NET Framework version  
    2019-09-06  1.1.7   Fixed example in New-Archive function
    2019-09-06  1.1.8   Fixed Expand-Zip function example
    2019-09-12  1.1.9   Changed order of system info output   
#>
#Requires -Version 2

#region script parameters
[CmdletBinding(DefaultParameterSetName='Ticket')]
Param
(
    [Parameter(
        ParameterSetName='Ticket',
        Position=0,    
        Mandatory=$true,
        ValueFromPipelineByPropertyName=$false,
        HelpMessage="The CentreStack support ticket number."
    )]
    [uint32]
    $Ticket,

    [Parameter(
        ParameterSetName='Info',
        Mandatory=$false,
        ValueFromPipeline=$false,
        HelpMessage="Only displays information to screen."
    )]
    [switch] $InfoOnly = $false

)
#endregion script parameters
#region functions
function New-Archive
{
    <#
    .SYNOPSIS
    Archives the contents of a directory into a zip file
    .DESCRIPTION
	Archives the contents of a directory into a zip file
    .EXAMPLE
    .\New-Archive -Source "C:\Temp" -ZipFile "C:\Temp.zip"
    .PARAMETER Source
    The fully qualified path to the source directory whose contents will be archived
    .PARAMETER ZipFile
    The fully qualified path to the zip file that will archive the source directory
    #>
    [CmdletBinding()]
    Param (
        [Parameter(
            Position=0,
            Mandatory=$true, 
            ValueFromPipeline = $false
        )]
        [string] $Source,

        [Parameter(
            Position=1,
            Mandatory=$true, 
            ValueFromPipeline = $false
        )]
        
        [string] $ZipFile

    )        

    # Check if this is PowerShell 2
    if ((Get-Host).Version.Major -eq 2) {
        # With PowerShell 2 call the Shell.Application COM object
        Set-Content -path $ZipFile -value ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18)) -ErrorAction Stop
        $shell = New-Object -ComObject Shell.Application
        $archive = $shell.NameSpace($ZipFile)
        $archive.CopyHere($Source)
        Start-Sleep -Milliseconds 200
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    }
    else {
        # .NET class is available in PowerShell 3.0 or later
        Add-Type -Assembly System.IO.Compression.FileSystem
        $compressionLevel = [System.IO.Compression.CompressionLevel]::Optimal
        [System.IO.Compression.ZipFile]::CreateFromDirectory($Source, $ZipFile, $compressionLevel, $false)
    }
} # end function New-Archive

function Expand-Zip
{
    <#
    .SYNOPSIS
    Expands the contents of a zip file into a new directory
    .DESCRIPTION
	Expands the contents of a zip file into a new directory
    .EXAMPLE
    .\Expand-Zip -ZipFile -ZipFile "C:\Temp.zip" -Destination "C:\Temp"
    .PARAMETER ZipFile
    The fully qualified path to the zip file that will archive the source directory
    .PARAMETER Destination
    The fully qualified path to the source directory whose contents will be archived
    #>
    [CmdletBinding()]
    Param (
        [Parameter(
            Position=0,
            Mandatory=$true, 
            ValueFromPipeline = $false
        )]
        [string] $ZipFile,

        [Parameter(
            Position=1,
            Mandatory=$true, 
            ValueFromPipeline = $false
        )]
        
        [string] $Destination

    )        

    # Check if this is PowerShell 2
    if ((Get-Host).Version.Major -eq 2) {
        # With PowerShell 2 call the Shell.Application COM object
        $shell = New-Object -ComObject Shell.Application
        $items = $shell.NameSpace($ZipFile).Items()
        $shell.NameSpace($Destination).CopyHere($items)
        Start-Sleep -Milliseconds 1000
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
    }
    else {
        # .NET class is available in PowerShell 3.0 or later
        Add-Type -Assembly System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipFile, $Destination)
    }
} # end function Expand-Zip


function Get-PSBitness() {
    # Returns $True if 64 bit PowerShell, else $False
    return [IntPtr]::size -eq 8
}
function Get-WebFile {
    <#
    .SYNOPSIS
    Downloads a file from a URI similar to wget
    .DESCRIPTION
    Downloads a file from a URI to wget. Useful for older PowerShell that lacks Invoke-WebRequest, 
    or to work around performance problem when Invoke-WebRequest displays the PS progress bar.
    .EXAMPLE
    .\Get-WebFile -URI 'https://download.sysinternals.com/files/DebugView.zip' -OutFile "C:\DebugView.zip" 
    .PARAMETER URI
    The web uniform resource identifier of the source file to be downloaded.
    .PARAMETER OutFile
    The destination output file.
    .PARAMETER Quiet
    When specified, the PowerShell progress bar will be supressed
    #>
    [CmdletBinding()]
    Param (
        [Parameter(
            Position=0,
            Mandatory=$true, 
            ValueFromPipeline = $false
        )]
        [string] $URI,

        [Parameter(
            Position=1,
            Mandatory=$true, 
            ValueFromPipeline = $false
        )]
        [string] $OutFile,
        [Parameter(
            Position=2,
            Mandatory=$false, 
            ValueFromPipeline = $false
        )]
        [switch]$Quiet

    )       
    $request = [System.Net.HttpWebRequest]::Create($URI)
    $response = $request.GetResponse()
  
    if ($response.StatusCode -eq 200) {
        Write-Verbose ("GET {0}" -f $URI)
        [int64] $fileLength = $response.ContentLength
        $streamReader= $response.GetResponseStream()
        try {
            $streamWriter = New-Object System.IO.FileStream $OutFile, "Create"
        }
        catch {
            throw $PSItem
        }
        [byte[]] $buffer = New-Object byte[] 4096
        [int64] $bytesWritten = 0
        [int] $count = 0
        $swProgress = [System.Diagnostics.Stopwatch]::StartNew()
        do {
            $count = $streamReader.Read($buffer, 0, $buffer.Length)
            $streamWriter.Write($buffer, 0, $count)
            $bytesWritten += $count
            # Only call the Write-Progress cmdlet if the Quiet switch was not specified and only every half second. This improves performance greatly (8x faster)
            if ((-not ($Quiet)) -and ($swProgress.Elapsed.TotalMilliseconds -ge 500)) {
                if ($fileLength -gt 0) {
                    Write-Progress -Activity ("Downloading to {0}" -f $OutFile) -Status ("Writing {0:N0} of {1:N0} bytes" -f $bytesWritten, $fileLength ) -id 0 -PercentComplete (($bytesWritten/$fileLength)*100)
                    # Reset and restart the stopwatch
                    $swProgress.Reset()
                    $swProgress.Start()
                } 
            }
       } while ($count -gt 0)
      
        $streamReader.Close()
        $streamWriter.Flush()
        $streamWriter.Close()
        Write-Verbose ("Received {0} bytes of content type {1}" -f $bytesWritten, $response.ContentType)
    }
    $response.Close()
} #end function Get-WebFile

function Get-DotNetVersion {

    $release = (Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full").Release

    if ($release -ge 528040) {return "Microsoft .NET 4.8 or later ($release)" }
    if ($release -ge 461808) {return "Microsoft .NET 4.7.2 ($release)"}
    if ($release -ge 461308) {return "Microsoft .NET 4.7.1 ($release)"}
    if ($release -ge 460798) {return "Microsoft .NET 4.7 ($release)"}
    if ($release -ge 394802) {return "Microsoft .NET 4.6.2 ($release)"}
    if ($release -ge 394254) {return "Microsoft .NET 4.6.1 ($release)"}      
    if ($release -ge 393295) {return "Microsoft .NET 4.6 ($release)"}    
    if ($release -ge 379893) {return "Microsoft .NET 4.5.2 ($release)"}      
    if ($release -ge 378675) {return "Microsoft .NET 4.5.1 ($release)"}      
    if ($release -ge 378389) {return "Microsoft .NET 4.5 ($release)"}  
    # This next line should not execute
    return "No 4.5 or later version detected";
    
} # end function Get-DotNetVersion

#endregion functions

#region Script Body

# Get this script
$ThisScript = $Script:MyInvocation.MyCommand
# Get the directory of this script
$scriptDir = Split-Path $ThisScript.Path -Parent
$CopyScript = Join-Path $scriptDir "Copy-CSSupport.ps1"
# Check that the Copy-CSSupport.ps1 exists. This script is called to upload files to S3
if (-not (Test-Path $CopyScript)) {
    $m = "Cannot continue as {0} is required." -f $CopyScript
    Throw $m
}

# This will return '32-bit' or '64-bit' for the operating system architecture
$os = Get-WmiObject Win32_OperatingSystem
if (($os.OSArchitecture -eq '64-bit') -and (-not (Get-PSBitness))) {
    Throw "32-bit PowerShell on a 64-bit operating system is not supported. Run the script in 64-bit PowerShell."
}

# CentreStack server registry key
if ($os.OSArchitecture -eq '64-bit') {
    # 64 bit Windows
    $regKey = 'HKLM:\SOFTWARE\WOW6432Node\Gladinet\Enterprise'
}
else {
    # 32 bit Windows
    $regKey = 'HKLM:\SOFTWARE\WOW6432Node\Gladinet\Enterprise'
}
$isCentreStack = $false
if (Test-Path $regKey) {
    # Attempt to determine if the machine is running the CentreStack web server
    try {
        $installDir = (Get-ItemProperty -Path $regKey -Name 'InstallDir').InstallDir
        $appImage = Join-Path $installDir "namespace\bin\userlib.dll"
        if (Test-Path ($appImage)) {
            # This is a CentreStack server
            $isCentreStack = $true
        }
        else {
            Throw "Unable to locate application image for the CentreStack Server." 
        } 
    }
    catch {
        Write-Verbose "Unable to open InstallDir registry key. Perhaps this is not a CentreStack server"
    }
}

# If it's not CentreStack than perhaps it is a CentreStack client
if (-not $isCentreStack) {
    # Server Agent and Windows Client share this same key
    if ($os.OSArchitecture -eq '64-bit') {
        # 64 bit Windows
        $regKey = 'HKLM:\SOFTWARE\Wow6432Node\Gladinet\AutoUpdate'
    }
    else {
        # 32 bit Windows
        $regKey = 'HKLM:\SOFTWARE\Gladinet\AutoUpdate'
    }

    try {
            $installPath = (Get-ItemProperty -Path $regKey -Name 'InstallPath').InstallPath
    }
    catch {
        throw $PSItem
    }

    # Server Agent and Windows Client share many of the same binaries. Examine the ProductName reg value to determine which is installed
    try {
        $productName = (Get-ItemProperty -Path $regKey -Name 'ProductName').ProductName
    }
    catch {
        throw $PSItem
    }

    if ($productName -like 'ServerAgent*') {
        $isServerAgent = $True
        $appImage = Join-Path $installPath "GladGroupSvc.exe"
        if (Test-Path ($appImage)) {
            $dataDir = Join-Path $env:ProgramData "gteamclient"
        }
        else {
            Throw "Unable to locate application image for the CentreStack Server Agent."
        }
    }
    else {
        $isServerAgent = $False
        $appImage = Join-Path $installPath "CoDesktopClient.exe"
        if (Test-Path ($appImage)) {
            $dataDir = Join-Path $env:LOCALAPPDATA "gteamclient"
        }
        else {
            Throw "Unable to locate application image for the CentreStack Windows Client."
        }
    }

}
# Get the product version information
$productVersion = (Get-Item -Path $appImage).VersionInfo.ProductVersion -replace '\s',''
$dotNetVersion = Get-DotNetVersion
# Use the old PowerShell 2 method of creating a customer psobject with ordered properties
$info = New-Object psobject
$info | Add-Member -MemberType NoteProperty -Name "ProductVersion" -Value $productVersion
$info | Add-Member -MemberType NoteProperty -Name "Image" -Value $appImage
$info | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value $env:COMPUTERNAME
$info | Add-Member -MemberType NoteProperty -Name "OS" -Value $os.Caption
$info | Add-Member -MemberType NoteProperty -Name "OSVersion" -Value $os.Version
$info | Add-Member -MemberType NoteProperty -Name "OSArchitecture" -Value $os.OSArchitecture
$info | Add-Member -MemberType NoteProperty -Name "TotalPhysicalRAMinMB" -Value ([math]::Round($os.TotalVisibleMemorySize / 1024))
$info | Add-Member -MemberType NoteProperty -Name "FreePhysicalRAMinMB" -Value ([math]::Round($os.FreePhysicalMemory / 1024))
$info | Add-Member -MemberType NoteProperty -Name "TotalVirtualRAMinMB" -Value ([math]::Round($os.TotalVirtualMemorySize / 1024))
$info | Add-Member -MemberType NoteProperty -Name "FreeVirtualRAMinMB" -Value ([math]::Round($os.FreeVirtualMemory / 1024))
$info | Add-Member -MemberType NoteProperty -Name "DotNETVersion" -Value $dotNetVersion

# Output system info to the screen
$info | Format-List

if ($PSBoundParameters.ContainsKey('InfoOnly')) {
    break
}

$now = Get-Date -f "yyyy-MM-dd_hh-mm-ss"
# Create a temp directory
$tempDir = [System.IO.Path]::GetTempPath()
# This variable will be unique each time the script runs due to the timestamp
if ($isCentreStack) {
    $sessionName = "CS_Server_Info_{0}_{1}" -f $Ticket.ToString(), $now
}
elseif ($isServerAgent) {
    $sessionName = "CS_ServerAgent_Info_{0}_{1}" -f $Ticket.ToString(), $now
}
else {
    $sessionName = "CS_WindowsClient_Info_{0}_{1}" -f $Ticket.ToString(), $now
}
# Get the user's MyDocuments directory
$myDocs = [Environment]::GetFolderPath("MyDocuments")

# Create a temporary destination folder to hold contents prior to zipping
$destRoot = New-Item -ItemType Directory -Path (Join-Path $tempDir $sessionName)
# Output info to root folder
$infoFile = Join-Path $destRoot ($sessionName + ".txt") 
$info | Out-File -FilePath $infoFile -Encoding utf8
. $CopyScript -FullName $infoFile -Ticket $Ticket
# Prompt operator to run DebugView
Add-Type -AssemblyName PresentationFramework
$msgBoxInput =  [System.Windows.MessageBox]::Show('Would you like to start SysInternals DebugView?','DebugView Prompt','YesNo','Question')
switch  ($msgBoxInput) {
    'No' {
        Write-Verbose "Skipping DebugView"
    }
    'Yes' {
        
        $DebugViewDir = Join-Path $tempDir "DebugView"
        $DebugView = Join-Path $DebugViewDir "dbgview.exe"
        # Download DebugView if necessary
        if (-not (Test-Path $DebugView)) {
            Write-Verbose "Downloading SysInternals DebugView from Microsoft"
            $DebugViewZip = Join-Path $tempDir "DebugView.zip"
            if (-not (Test-Path $DebugViewZip)) {
                $Uri = 'https://download.sysinternals.com/files/DebugView.zip'
                $m = Measure-Command {Get-WebFile -URI $Uri -OutFile $DebugViewZip}
                Write-Verbose ("Completed download in: {0:g}" -f $m)
            }
            if (Test-Path $DebugViewZip) {
                $DebugViewDir = Join-Path $tempDir "DebugView"
                if (-not (Test-Path $DebugViewDir)) {New-Item $DebugViewDir -ItemType Directory | Out-Null}
                Expand-Zip -ZipFile $DebugViewZip -Destination $DebugViewDir
            }
        }
        # Make sure dbgview.exe exists
        if (-not (Test-Path $DebugView)) {
            Throw ("Unable to locate '{0}'" -f $DebugView)
        }
        else {
            # Build dbgview command line arguments
            $dbgViewLog = Join-Path $myDocs ("DebugView_{0}_{1}.log" -f $info.ComputerName, $now)
            $argList = @('/f', '/om', '/l', $dbgViewLog)
            # DebugView must run with Global Win32 on CentreStack Server or Server Agent, but not Windows Client
            if (($isCentreStack) -or ($isServerAgent) ) { 
                $argList += '/g'
            }
            Write-Verbose ("Executing: {0} {1}" -f $DebugView, [String]::Join(" ", $argList))
            # Execute dbgview command line
            if (($isCentreStack) -or ($isServerAgent) ) { 
                # Run DebugView As Administrator on Server Agent or CentreStack server
                $proc = Start-Process -FilePath $DebugView -ArgumentList $argList -PassThru -Wait -Verb RunAs
            }
            else {
                $proc = Start-Process -FilePath $DebugView -ArgumentList $argList -PassThru -Wait
            }
            Write-Verbose ("DebugView exit code: {0}" -f $proc.ExitCode)
            Write-Output ("DebugView log: {0}" -f $dbgViewLog)
            Copy-Item $dbgViewLog -Destination $destRoot
        }
    }
}
if ($isCentreStack) {
    # Get some server files
    Copy-Item (Join-Path $installDir "root\web.config") -Destination $destRoot
}
else {
    # Create a gteamclient folder
    $destgteamclient = New-Item -ItemType Directory -Path (Join-Path $destRoot.FullName "gteamclient")
    # Copy the "C:\ProgramData\gteamclient\gsettings.db" to temp destination
    Copy-Item (Join-Path $dataDir "gsettings.db") -Destination $destgteamclient
    # Copy contents of logging directory to tmpDir
    $srcDir = Join-Path $dataDir "logging"
    $destDir = Join-Path $destgteamclient "logging\"
    New-Item $destDir -ItemType Directory | Out-Null
    # Only copy logging files from the last 30 days
    Get-ChildItem -Path $srcDir | 
        Where-Object {$_.LastWriteTime -ge (Get-Date).AddDays(-30)} |
            Copy-Item -Destination $destDir
}
$zipFile = Join-Path $myDocs ($sessionName + ".zip")
# Make a zip file
New-Archive -Source $destRoot -ZipFile $zipFile
# Remove the temp files 
Remove-Item $destRoot -Recurse
# Call external script to upload file
. $CopyScript -Path $zipFile -Ticket $Ticket
#endregion Script Body