# Invoke-MSERTScan
This script automates the MSERT scan outlined in [ED21-02 Supplemental Direction v2](https://cyber.dhs.gov/ed/21-02/#supplemental-direction-v2/).  The script is meant to be executed as a weekly scheduled task but can also be manually executed.  It downloads the latest version of the MSERT tool from [Microsoft](https://go.microsoft.com/fwlink/?LinkId=212732) on each device passed to the computers parameter and runs the scan from that computer.  Once complete the script will compress all logs from each endpoint and can send an email with the logs attached. 

## Dependencies


```powershell
#Requires -RunAsAdministrator
#Requires -Version 3.0
```

## Usage

```powershell
#Downloads and executes MSERT.exe on all computers in the array
.\Invoke-MSERTScan.ps1 -Computers @("Exch01","Exch02") 

#Sends email with MSERT logs.  Each log is stamped with the hostname of the device the logs were run on.
.\Invoke-MSERTScan.ps1 -Computers Exch01 -SendEmail -To "Test@test.com" -From "Automation@test.com" -SMTPServer "SomeSMTPRelay.com" 

```

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.


## License
[MIT](https://choosealicense.com/licenses/mit/)
