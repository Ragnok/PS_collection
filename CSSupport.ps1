<#
.SYNOPSIS
    Uploads a file to CentreStack support's S3 bucket
.DESCRIPTION
	Uploads a file to CentreStack support's S3 bucket
.PARAMETER Ticket
    The CentreStack support ticket number
.PARAMETER FullName
    The full pathname to a file to be uploaded to CentreStack support's S3 bucket. 
.EXAMPLE
    .\Copy-CSSupport.ps1 -Ticket 8675309 -FullName C:\Temp\DebugView.zip

    Demonstrates uploading a single file.
.EXAMPLE
    Get-ChildItem C:\Dumps -Filter *.txt | .\Copy-CSSupport.ps1 -Ticket 8675309
    
    Demonstrates using the script with the PowerShell pipeline.
    Uses the Get-ChildItem (aka "dir") cmdlet to send *.txt files down the pipeline
    to be processed by the script
.EXAMPLE
    .\Copy-CSSupport.ps1 -Ticket 8675309 -FullName ("C:\Dumps\1.zip", "C:\Dumps\2.zip")
    
    Demonstrates sending an array of file names to the script to be uploaded sequentially. 
.NOTES 
    Author: Jeff Reed
    Name: Copy-CSSupport.ps1
    Created: 2019-08-12
    
    Version History
    2019-08-12  1.0.0   Initial version
    2019-08-13  1.0.1   Renamed script
    2019-08-13  1.1.0   Refactored script to support the object pipeline or an array of files with the FullName parameter
    2019-08-29  1.1.1   Tweaks to comment based help
    2019-08-30  1.1.2   Changed Ticket parameter to uint32
#>
#Requires -Version 2

#region script parameters
[CmdletBinding()]
Param
(
    [Parameter(
        Position=0,    
        Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true,
        HelpMessage="The file path of the file to be uploaded to S3."
    )]
    [Alias('Path')]
    [String[]]$FullName,

    [Parameter(
        Position=1,    
        Mandatory=$true,
        ValueFromPipeline=$false,
        HelpMessage="The CentreStack support ticket number."
    )]
    [uint32]
    $Ticket
)
#endregion script parameters

#region Script Body 
# This script supports a pipeline of objects
Begin {
    # Executes once before first item in pipeline is processed
    # S3 configuration
    $s3Bucket = "hadroncloud-support"
    $s3AccessKey="AKIAW2MRMDPHIN7Z7UXB"
    $s3SecretKey="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"
  }

Process {
    foreach ($path in $FullName) {
        try {
            # Get a FileInfo object for the object in the pipeline
            Write-Verbose ("Path: {0}" -f $path)
            $fi = Get-Item $path
        }
        catch {
            Write-Warning $PSItem.Exception.Message
        }
        try {
            $fileName = $fi.Name
            $relativePath = "/$s3Bucket/uploads/{0}/{1}" -f $Ticket.ToString(), $fileName
            $s3URL = "http://$s3Bucket.s3.amazonaws.com/uploads/{0}/{1}" -f $Ticket.ToString(), $fileName
        
            # build the S3 request
            $dateFormatted = get-date -UFormat "%a, %d %b %Y %T %Z00"
            $httpMethod="PUT"
            $contentType="application/octet-stream"
            $stringToSign="$httpMethod`n`n$contentType`n`nx-amz-date:$dateFormatted`n$relativePath"
        
            # generate the signature
            $hmacsha = New-Object System.Security.Cryptography.HMACSHA1
            $hmacsha.key = [Text.Encoding]::ASCII.GetBytes($s3SecretKey)
            $signature = $hmacsha.ComputeHash([Text.Encoding]::ASCII.GetBytes($stringToSign))
            $signature = [Convert]::ToBase64String($signature)
        
            # Set the URI of the web service
            $URI = [System.Uri]$s3URL;
        
            # Create a new web request
            $fileStream = [System.IO.File]::OpenRead($fi.FullName)
            $WebRequest = [System.Net.HttpWebRequest]::Create($URI);
            $WebRequest.Method = $httpMethod;
            $WebRequest.ContentType = $contentType
            $WebRequest.ContentLength = $fileStream.Length
            $WebRequest.PreAuthenticate = $true;
            $WebRequest.Accept = "*/*" # may not be required
            $WebRequest.UserAgent = "powershell" # not required
        
            # this is needed to generate a valid signature!
            $WebRequest.Headers.add('x-amz-date', $dateFormatted)
            $WebRequest.Headers["Authorization"] = "AWS $s3AccessKey`:$signature";
        
            # Create a webstream to write to
            $WebStream = $WebRequest.GetRequestStream()
            
            # Can't use this on PowerShell 2: $fileStream.CopyTo($WebStream)
            # Use this method to read bytes from the filestream and write then the webstream
            $swProgress = [System.Diagnostics.Stopwatch]::StartNew()
            $progressActivity = ("Uploading {0}" -f $fileName )      
            $buffer = New-Object Byte[] 10240
            [int64] $bytesWritten = 0
            while (($count = $fileStream.Read($buffer, 0, $buffer.Length)) -gt 0)
            {
                $WebStream.Write($buffer, 0, $count)
                $bytesWritten += $count
                $pctComplete = ($fileStream.Position / $fileStream.Length)
                # Only call the Write-Progress cmdlet every half second. This improves performance greatly (8x faster)
                if ($swProgress.Elapsed.TotalMilliseconds -ge 500) {
                    Write-Progress -Activity $progressActivity -Status ("{0:P0} complete:" -f $pctComplete) -PercentComplete ($pctComplete * 100)
                    $swProgress.Reset()
                    $swProgress.Start()                
                }
                
            }
            # Cleanup 
            Write-Progress -Activity $progressActivity -Status "Ready" -Completed
            $WebStream.Flush()
            $WebStream.Close()
            $response = $WebRequest.GetResponse()
            if ($response = 'OK') {
                Write-Output ("Uploaded {0}" -f $fi.FullName)
            }
            else {
                Write-Warning ("Failed to upload {0}" -f $fi.FullName)
            }
            Write-Verbose ("Sent {0:N0} bytes. Response: {1}" -f $bytesWritten, $response)
            $fileStream.Close()     
        }
        catch {
            throw $PSItem
        }    
    }

}

End {
    # Executes once after last pipeline object is processed
}



#endregion Script Body
