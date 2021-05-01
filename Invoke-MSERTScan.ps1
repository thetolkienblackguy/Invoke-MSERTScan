<###########################################################################################################################################
.DESCRIPTION 
This script automates the MSERT scan outlined in ED21-02 Supplemental Direction v2.  The script is meant to be executed as a weekly 
scheduled task but can also be manually executed.  It downloads the latest version of the MSERT tool from
https://go.microsoft.com/fwlink/?LinkId=212732 on each device passed to the computers parameter and runs the scan from that computer.  
Once complete the script will compress all logs from each endpoint and can send an email with the logs attached. 

.EXAMPLE
The computers parameter is mandatory all others while they have a default value are not. The parameter is an array thus accepts multiple 
endpoints if necessary. 
.\Invoke-MSERTScan.ps1 -Computers @("Exch01","Exch02")

Sending an email with logs
.\Invoke-MSERTScan.ps1 -Computers Exch01 -SendEmail -To "Test@test.com" -From "Automation@test.com" -SMTPServer "SomeSMTPRelay.com"

.NOTES
Author: Gabe Delaney
Version: 1.0
Date: 04/26/2021
Name: Invoke-MSERTScan


Version History:
1.0 - Original Release - Gabe Delaney 
###########################################################################################################################################>
#Requires -RunAsAdministrator
#Requires -Version 3.0
param (        
    [Parameter(Mandatory=$true)] 
    [array]$Computers,
    [Parameter(Mandatory=$false)] 
    [string]$MSERTLogFolder = "C:\Temp",
    [Parameter(Mandatory=$false)] 
    [switch]$SendEmail,
    [Parameter(Mandatory=$false)]
    [array]$To = "test@test.com",
    [Parameter(Mandatory=$false)]
    [array]$Cc = "test@test.com",
    [Parameter(Mandatory=$false)]
    [string]$From = "test@test.com",
    [Parameter(Mandatory=$false)]
    [string]$Subject = "MSERT Logs $(Get-Date -Format MMddyyyy)",
    [Parameter(Mandatory=$false)]
    [string]$Body = "See attached MSERT logs for review.",
    [Parameter(Mandatory=$false)]
    [string]$SMTPServer = "smtp.com"

)
<#
    This scriptblock variable will be passed to invoke-command to include all functions and paramters as well as leverage Invoke-Command's built in parrallel processing. 
    There may be a more elegant way to do this but this is ultimately the way I settled on after writing a couple hundred lines of code and realizing I'll need
    both parrallel processing and remoting a bit too late. 

#>
$scriptBlock = {
    $logFile = "C:\Temp\Invoke-MSERTScan_$(Get-Date -Format MMddyyyy_hhmmss).Log"
    $link = "https://go.microsoft.com/fwlink/?LinkId=212732"
    $outFile = "C:\Temp\MSERT.exe"
    Function Invoke-Logging {
        <#
            This function just helps streamline logging.  Specifically when logging to a file AND the console.
        
        #>
        [CmdletBinding()]
        param (        
            [Parameter(Mandatory=$true)]
            [string]$Message,
            [Parameter(Mandatory=$true)]
            [string]$LogFile,
            [Parameter(Mandatory=$false)]
            [switch]$WriteOutput,
            [Parameter(Mandatory=$false)]
            [ValidateSet(
                "Black",
                "Blue",
                "Cyan",
                "DarkBlue",
                "DarkCyan",
                "DarkGray",
                "DarkGreen",
                "DarkMagenta",
                "DarkRed",
                "DarkYellow",
                "Gray",
                "Green",
                "Magenta",
                "Red",
                "Yellow",
                "White"

            )]
            [string]$ForeGroundColor = "Yellow"

        )
        Begin {
            $message = "$("[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)) $message"
            $logPath = Split-Path $logFile -Parent
            If (!(Test-Path $logPath)) {
                New-Item -Path $logPath -ItemType Directory -Force | Out-Null
            
            }
        }
        Process {
            $message | Out-File $logFile -Append

        } End {
            If ($writeOutput) {
                Write-Host $message -ForegroundColor $foreGroundColor -BackgroundColor Black

            }
        }
    }
    Function Start-Download {
        <#
            Function just leverages Invoke-WebRequest to download a file.  The only additional functionality added is the built in
            usage of TLS1.2. It also allows me to expand upon the function in other scripts if necessary 

        #>
        [CmdletBinding()]
        Param (        
            [Parameter(Mandatory=$true,Position=0)]
            [string]$Link,      
            [Parameter(Mandatory=$true,Position=1)] 
            [string]$OutFile
        
        )
        Begin {   
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $webRequestParams = @{
                Uri = $link
                Outfile = $outFile

            }
        } Process {
            Invoke-WebRequest @webRequestParams 

        } End {
        
        }
    }
    Function Start-MSERTScanner {
        <#
            The function is a wrapper for the MSERT cmd line execution.   

        #>
        [CmdletBinding()]
        Param (        
            [Parameter(Mandatory=$false)]
            [Alias("MSERTPath","P")]
            [ValidateScript({
                If ($_) {  
                    Test-Path $_
                    
                    }
                }
            )]
            [string]$Path = "",        
            [Parameter(Mandatory=$false)] 
            [Alias("F")]
            [switch]$FullScan,
            [Parameter(Mandatory=$false)] 
            [Alias("N")]
            [switch]$DetectOnly,
            [Parameter(Mandatory=$false)] 
            [Alias("Q")]
            [switch]$Quiet,
            [Parameter(Mandatory=$false)]
            [Alias("H")] 
            [switch]$SevereOnly,
            [Parameter(Mandatory=$false)]
            [int]$TimeOut = 14400

        )
        Begin {
            $arguments = ""
            $arguments += If ($fullScan) {
                " /F"
            }
            $arguments += If ($detectOnly) {
                " /N"

            } Else {
                ":Y"
            }
            $arguments += If ($quiet) {
                " /Q"
            }
            $arguments += If ($severeOnly) {
                " /H"
            }
            $startProcessParams = @{
                FilePath = $path
                ArgumentList = $arguments
                Passthru = $true
            }
        } Process {
            If (![boolean](Get-Process MSERT -ErrorAction SilentlyContinue)) {
                $process = Start-Process @startProcessParams
                Try {
                    <#
                        The script will wait for the MSERT process to complete or timeout before continuting. 

                    #>
                    $process | Wait-Process -Timeout $timeOut -ErrorAction Stop
                
                }
                Catch [TimeoutException] {
                    Throw "WARNING: Microsoft security scanner took longer than $((New-TimeSpan -Seconds $timeOut).TotalHours) hours to complete, action terminated: $($_.Exception.Message)"
                    
                }
            } Else {
                Throw "MSERT is already running a scan on this system"
            
            }
        } End {
            
        }
    }
    #Start-Download parameters
    $downloadParams = @{
        Link = $Link
        OutFile = $outFile

    }
    #Start-MSERTScanner parameters
    $msertParams = @{
        Path = $outFile
        FullScan = $true
        DetectOnly = $true
        Quiet = $true

    }
    #Invoke-Logging parameters
    $invokeLoggingParams = @{
        LogFile = $logFile
        WriteOutput = $true 

    }
    Invoke-Logging -Message "Downloading newest version of MSERT on $env:ComputerName" @invokeLoggingParams
    Try {
        Start-Download @downloadParams

    } Catch {
        Invoke-Logging -Message "Failed to download MSERT.exe on $env:ComputerName.  $($error[0].Exception.Message)" @invokeLoggingParams

    }
    Try {
        Invoke-Logging -Message "Starting MSERT scanner on $env:ComputerName" @invokeLoggingParams
        Start-MSERTScanner @msertParams

    } Catch {
        Invoke-Logging -Message "Failed to start the MSERT scanner on $env:ComputerName. $($error[0].Exception.Message)" @invokeLoggingParams

    }
}
#Invoke-Command parameters
$invokeCmdParams = @{
    ComputerName = $computers
    ScriptBlock = $scriptBlock

}
Invoke-Command @invokeCmdParams
<#
    Copies the MSERT log from each server to a central location. Each log file is stamped with its computer name to differientiate from other endpoint logs

#>    
Foreach ($computer in $computers) {
    $path = "\\$computer\C$\Windows\Debug\MSERT.Log"
    #Copy-Item parameters
    $copyItemParams = @{
        Path = $path
        Destination = $msertLogFolder
        Force = $true

    }
    #Rename-Item parameters
    $renameItemParams = @{
        Path = "$msertLogFolder\MSERT.log"
        NewName = "$($computer)_MSERT.log"
        Force = $true

    }
    Copy-Item @copyItemParams
    Rename-Item @renameItemParams

}
#Compress-Archive parameters
$compress = @{
    Path = "$msertLogFolder\*MSERT.log"
    CompressionLevel = "Fastest"
    DestinationPath = "$msertLogFolder\MSERT.Zip"
    Force = $true

}
#Remove-Item parameters
$removeItemParams = @{
    Path = "$msertLogFolder\*MSERT*"
    Confirm = $false 

}
Compress-Archive @compress 
If ($sendEmail) {
    #Send-MailMessage parameters
    $sendMail = @{
        To = $to
        From = $from
        Subject = $subject
        SMTPServer = $smtpServer
        BodyAsHTML = $true
        Body = $body
        Attachment = "$msertLogFolder\MSERT.zip"
    
    }
    If ($cc) {
        $sendMail += @{
            Cc = $cc

        }
    }
    Send-MailMessage @sendMail

}    
#Folder clean-up
Remove-Item @removeItemParams
Exit
