param ($run, $opt, $runr, $optr)
Import-Module PSWorkflow

try
    {
        $LOCUSR = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
    }
catch
    {
        $LOCUSR = "nobody.here"
    }
$SLOC= Split-Path -Parent $($global:MyInvocation.MyCommand.Definition)

$kfixes = @'
param ($run, $opt, $runr, $optr)
Import-Module PSWorkflow

try
    {
        $LOCUSR = ((Get-WMIObject -ClassName Win32_ComputerSystem).Username).Split('\')[1]
    }
catch
    {
        $LOCUSR = "nobody.here"
    }
$SLOC= Split-Path -Parent $($global:MyInvocation.MyCommand.Definition)


Function Net-Cleanup
    {    

        if ($LOCUSR -ne "nobody.here")
            {
                Get-Process -Name "*msedge" | Stop-Process -Force -ErrorAction SilentlyContinue
                Get-Process -Name "*Chrome*" | Stop-Process -Force -ErrorAction SilentlyContinue

                Write-Host "CLEARING INTERNET CACHE"
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Cache\Cache_Data\*" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Network\Cookies" -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Network\Cookies-journal" -Force -ErrorAction SilentlyContinue
                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Internet cache."

                Write-Host "CLEARING CHROME CACHE"
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cookies" -Recurse -Force -ErrorAction SilentlyContinue 
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Media Cache" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cookies-Journal" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Cookies" -Force -ErrorAction SilentlyContinue 
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Media Cache" -Force -ErrorAction SilentlyContinue
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Cookies-Journal" -Force -ErrorAction SilentlyContinue
                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Chrome cache."
                Write-Host "DONE"

                Remove-Item C:\users\$LOCUSR\AppData\Local\Temp\* -Recurse -Force -ErrorAction SilentlyContinue
            }

        Write-Host "CLEARING LOCAL CACHE"
        Remove-Item C:\Windows\Temp\* -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item C:\Windows\Prefetch\* -Recurse -Force -ErrorAction SilentlyContinue
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared local cache."
        Write-Host "DONE"

        Write-Host "FLUSHING DNS"
        cmd.exe /c ipconfig /flushdns
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Flushed DNS."
        Write-Host "DONE"
    }

Function Windows-Repair
    {
        Write-Host "REPAIRING WINDOWS IMAGE"
        $dismtimer = "Start-Sleep -Seconds 600
        cmd /c 'taskkill /IM dism.exe /F'"

        $dismtimer | Out-File "C:\temp\dt.ps1"
        Start-Sleep -Seconds 3
        Start-Job -FilePath C:\temp\dt.ps1 | Out-Null
        dism /online /cleanup-image /restorehealth
        Remove-Item "C:\temp\dt.ps1" 
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran DISM."
        Write-Host "DONE."

        Write-Host "RUNNING SYSTEM FILE CHECK"
        sfc /scannow
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran SFC."
        Write-Host "DONE"

        Write-Host "REPAIRING MICROSOFT COMPONENTS"
        Get-AppXPackage | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml"} -ErrorAction SilentlyContinue
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Re-registered Microsoft components."
        Write-Host "DONE"

        Write-Host "UPDATING GROUP POLICY"
        cmd.exe /c echo n | gpupdate /force /wait:0
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran GPUpdate."
        Write-Host "DONE"  
    }
Function Run-WinUpdate
    {    
        Write-Host "Running Windows Updates."    
        Install-Module -Name PSWindowsUpdate -Force
        Import-Module PSWindowsUpdate
        Get-WindowsUpdate | Out-Null
        Install-WindowsUpdate -AcceptAll
        Write-Host "Done."
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran Windows Updates."

    }
Function Get-RecentErrors
    {
        Write-Host "--- Recent Application Errors ---"
        Get-WinEvent -LogName Application -FilterXPath "*[System[Level=1 or Level=2]]" -MaxEvents 10 | 
            Select-Object TimeCreated, ProviderName, Message | Format-List

        Write-Host "--- Recent System Errors ---"
        Get-WinEvent -LogName System -FilterXPath "*[System[Level=1 or Level=2]]" -MaxEvents 10 | 
            Select-Object TimeCreated, ProviderName, Message | Format-List
    }
Function Reset-Network
    {
        Write-Host "Resetting Winsock Catalog..."
        cmd.exe /c "netsh winsock reset"

        Write-Host "Resetting TCP/IP Stack..."
        cmd.exe /c "netsh int ip reset"

        Write-Host "Releasing and Renewing IP Address..."
        cmd.exe /c "ipconfig /release"
        cmd.exe /c "ipconfig /renew"

        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Reset Winsock, TCP/IP Stack, and renewed IP address."
    }
Function Driver-Update
    {
        Function HP-Update
            {
                $download = 'C:\Temp\HPIA\Download'
                $report = 'C:\Temp\HPIA\Report'
                $log = 'C:\Temp\HPIA\Log'
            
                Function Run-HPIA
                    {
                        Write-Host "Analyzing Drivers"
                        try
                            {
                                Start-Process "C:\Program Files\HP\HPIA\HPImageAssistant.exe" -ArgumentList "/Operation:Analyze /Action:List /Silent /ReportFolder:$report" -wait
                            }
                        catch
                            {
                                Write-Host "Error analyzing drivers."
                                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Unable to run Driver Upgrades."
                                exit
                            }
                        $jsonFile = Get-ChildItem $report -Filter "*.json" | Where-Object { $_.PSIsContainer -eq $false }
                        $list = Get-Content $jsonFile.FullName | Where-Object {$_ -match '"Name":' -or $_ -match '#'}
                        $dreport = $list -replace '"Name":', "Upgrading Driver:" -replace ","," " -replace '"', " "
            
                        $dreport
                        try
                            {
                                Start-Process "C:\Program Files\HP\HPIA\HPImageAssistant.exe" -ArgumentList "/Operation:Analyze /Action:Install /Category:BIOS,Drivers,Firmware,Software /Selection:All /Noninteractive /SoftPaqDownloadFolder:$download /ReportFolder:$report /AutoCleanup /LogFolder:$log" -wait
                            }
                        catch
                            {
                                Write-Host "Error upgrading drivers."
                                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Unable to run Driver Upgrades."
                                exit
                            }

                        Write-Host "Drivers Upgraded."
                        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Ran Driver Upgrades."
                        $dreport[0] = " $($dreport[0])"
                        Add-Content -Path C:\Temp\k-log.txt -Value "$($dreport | ForEach-Object { "$_`n" -replace 'Upgrading Driver:', 'Upgraded Driver:' })"
                    }    
                Function Setup-HPIA
                    {
                        Function Install-HPIA
                            {
                                try
                                    {
                                        Start-Process "C:\Temp\hp-hpia-5.3.2.exe" -ArgumentList '/s /e /f "C:\Program Files\HP\HPIA"'
                                    }
                                catch
                                    {
                                        Write-Host "Error installing HPIA."
                                        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to install HPIA."
                                    }

                                Do {Start-Sleep -Seconds 5}
                                until (Test-Path -path "C:\Program Files\HP\HPIA\HPImageAssistant.exe" -PathType Leaf)
                                
                                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Set up HPIA."
                                Write-Host "Queuing Driver and Firmware upgrades."

                            }
                        Write-Host "Preparing Device Upgrades..."
                        If (Test-Path -path "C:\Temp\hp-hpia-5.3.2.exe")
                            {          
                                Install-HPIA
                            }
                        else
                        {
                            try
                                {
                                    Invoke-WebRequest -Uri https://hpia.hpcloud.hp.com/downloads/hpia/hp-hpia-5.3.2.exe -OutFile C:\Temp\hp-hpia-5.3.2.exe | Out-Null
                                }
                            catch
                                {
                                    Write-Host "Error downloading HPIA."
                                    Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to download HPIA."
                                    exit
                                }
                                Do {Start-Sleep -Seconds 5}
                                until (Test-Path -path "C:\Temp\hp-hpia-5.3.2.exe" -PathType Leaf)

                                $hash  = "E1B8214B0C0B3B6B4C37181CAC1914B4E16FA60E6B063A7F76A1D7A96908E1F0"
                                $fh = Get-FileHash -Path "C:\Temp\hp-hpia-5.3.2.exe" | Select-Object -ExpandProperty Hash

                                If ($hash -eq $fh)
                                    {
                                        Install-HPIA
                                    }
                                else
                                    {
                                        Write-Host "Hash Mismatch."
                                        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to download HPIA, hash mismatch."
                                        exit
                                    }

                        }
                    }
            
                If (Test-Path -path "C:\Program Files\HP\HPIA\HPImageAssistant.exe")
                    {
                    Run-HPIA
                    }
                Else
                    {
                    Setup-HPIA
                    Run-HPIA
                    }

            }


        if ((Get-WmiObject -Class Win32_ComputerSystem).Manufacturer -eq "HP")
            {
                HP-Update
            }
        else
            {
                Write-Host "Non-HP system, please run updates manually."    
            }
    }
Function Reinstall-Drivers
    {
        $drivers = Get-PnpDevice
        $duplicateDrivers = $drivers | Group-Object -Property HardwareID | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Group

        foreach ($duplicateDriver in $duplicateDrivers)
        {
            $duplicateDriver | Select-Object -Skip 1 | ForEach-Object { &"pnputil" /remove-device $_.InstanceId }
        }

        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Checked for and uninstalled duplicate or corrupt device drivers."
                
        foreach ($dev in (Get-PnpDevice | Where-Object{$_.Name -Match "##DRVR##"}))
        {
            &"pnputil" /remove-device $dev.InstanceId
        }

        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Reinstalled ##DRVR## drivers."

        Start-Process "cmd.exe" -ArgumentList " /c C:\Windows\System32\pnputil.exe /scan-devices" -Wait -ErrorAction SilentlyContinue
        Write-Host "DONE"
    }
Function Clear-PrintQueue
    {
        Write-Host "Stopping the Print Spooler service."
        Stop-Service -Name Spooler -Force
        
        $spoolPath = "C:\Windows\System32\spool\PRINTERS\*"
        Write-Host "Clearing files from $spoolPath"
        try
            {
                Remove-Item -Path $spoolPath -Force -ErrorAction Stop
                Write-Host "Print queue cleared successfully."
                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared print queue."
            }
        catch
            {
                Write-Host "Error clearing some files."
                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Failed to clear print queue."
            }
        
        Write-Host "Restarting the Print Spooler service."
        Start-Service -Name Spooler
    }
Function Repair-Office
    {
        Get-Process -Name "WINWORD", "EXCEL", "POWERPNT", "OUTLOOK" -ErrorAction SilentlyContinue | Stop-Process -Force
        $ExePath = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"

        if ($ExePath)
            {
                Write-Host "REPAIRING OFFICE"
                try
                    {
                        cmd.exe /C  "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" scenario=Repair platform=x86 culture=en-us RepairType=FullRepair forceappshutdown=True DisplayLevel=False
                        Write-Host "DONE"
                        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Ran Online Office Repair."
                    }
                catch
                    {
                        Write-Error "Unable to run Office Repair."
                    }
            }
        else
            {
                Write-Error "Repair Client not found."
            }
    }
Function Clear-Slack
    {
        if ($LOCUSR -eq "nobody.here")
            {
                $LOCUSR = Read-Host "Enter username of Slack account to clear"
            }
        Get-Process -Name *Slack* -ErrorAction SilentlyContinue | Stop-Process -Force 
        Remove-Item -Recurse -Path "C:\Users\$LOCUSR\AppData\Roaming\Slack\Cache\Cache_Data"
        Remove-Item -Recurse -Path "C:\Users\$LOCUSR\AppData\Roaming\Slack\Code Cache\js"
        Remove-Item -Recurse -Path "C:\Users\$LOCUSR\AppData\Roaming\Slack\Code Cache\wasm"
        Remove-Item -Recurse -Path "C:\Users\$LOCUSR\AppData\Roaming\Slack\GPUCache"
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Slack cache." 
    }
Function Post-Image
    {
        Clear-Host
        Function PC-Stats {
            $PCName = hostname
            $Ram = [Math]::Round((Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1GB)
            $CPU = Get-CimInstance -ClassName Win32_Processor | Select-Object Name
            $BIOS = (Get-WmiObject win32_bios).SMBIOSBIOSVersion.Split(' ')[-1].TrimStart('0')

                                        if ($CPU.Name -like "*i7*")
                                            {
                                                $inum = "i7"
                                            }
                                        elseif ($CPU.Name -like "*i3*")
                                            {
                                                $inum = "i3"
                                            }
                                        else
                                            {
                                                $inum = "Unknown"
                                            }

                                    Write-Host "PC Name:" -BackgroundColor White -ForegroundColor Black -NoNewLine
                                    Write-Host " $PCName"
                                    Write-Host "PC Stats:"-BackgroundColor White -ForegroundColor Black -NoNewLine
                                    Write-Host " $Ram GB  /  $inum"
                                    Write-Host "BIOS Updates:"-BackgroundColor White -ForegroundColor Black -NoNewLine
                                    Write-Host " Version $BIOS"

                        }


        Function Software-Check {

            $Crowdstrike = Test-Path -path "C:\Program Files\CrowdStrike"
            $Drive = Test-Path -path "C:\Program Files\Google\Drive File Stream"
            $Bit = Get-BitLockerVolume -MountPoint C: | Select-Object ProtectionStatus

                                    if ($Bit.ProtectionStatus -like "*On*")
                                    {
                                        Write-Host "Bitlocker ON:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " True" -ForegroundColor Green
                                    }
                                else
                                    {
                                        Write-Host "Bitlocker ON:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " False" -ForegroundColor Red
                                    }

            
                                if ($Crowdstrike)
                                    {
                                        Write-Host "CrowdStrike / Ninja:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " Pass" -ForegroundColor Green
                                    }
                                else
                                    {
                                        Write-Host "CrowdStrike / Ninja:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " Fail" -ForegroundColor Red
                                    }

                                if ($Drive)
                                    {
                                        Write-Host "Google Drive Installed:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " Pass" -ForegroundColor Green
                                    }
                                else
                                    {
                                        Write-Host "Google Drive Installed:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " Fail" -ForegroundColor Red
                                    }
            
                

                        }
            
        Function HD-Test {

            $HDSize = Get-PhysicalDisk | Select-Object @{Name='Size(GB)'; Expression={[math]::Round($_.Size / 1GB, 2)}} | Select-Object -ExpandProperty 'Size(GB)'
            $HDType = Get-PhysicalDisk | Select-Object FriendlyName, Manufacturer

                        if ($HDType.Manufacturer -ne  $null)
                            {
                                $HDBrand = $HDType.Manufacturer
                            }
                        elseif ($HDType.Manufacturer -eq  $null)
                            {
                                $HDBrand = Get-PhysicalDisk | Select-Object -ExpandProperty FriendlyName | ForEach-Object { $_.Substring(0, 3) }
                            }
                        else
                            {
                                $HDBrand = "Unknown"
                            }

            Write-Host "SSD/Size Test:"-BackgroundColor White -ForegroundColor Black -NoNewline
            Write-Host " Brand - $HDBrand    Size - $HDSize"

                        }

        Function Wifi-Test {
            $wifiAdapter = Get-NetAdapter -Name *Wi-Fi* -ErrorAction SilentlyContinue

                                if ($wifiAdapter) {
                                    Write-Host "Wi-Fi Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                    Write-Host " Pass" -ForegroundColor Green
                                }
                                else {
                                    Write-Host "Wi-Fi Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                    Write-Host " Fail" -ForegroundColor Red
                                }
                            }
            

        Function Battery-Test
            {

                powercfg /batteryreport -output C:\Temp\batt.html | Out-Null
                    
                    function Get-BatteryValue
                        {
                            param
                                (
                                    [Parameter(Mandatory=$true)]
                                    [string]$htmlFilePath,
                                    [Parameter(Mandatory=$true)]
                                    [string]$valueToFind
                                )

                            $htmlContent = Get-Content -Path $htmlFilePath -Raw
                            $pattern = "(?s)<span class=`"label`">$valueToFind</span></td><td>(?<value>[\d,]+ mWh)"
                            $match = [regex]::Match($htmlContent, $pattern)

                                return $match.Groups["value"].Value.Trim()
                    }
        `

                $batteryReportPath = "C:\Temp\batt.html"

                $designCapacity = Get-BatteryValue -htmlFilePath $batteryReportPath -valueToFind "DESIGN CAPACITY"
                $fullChargeCapacity = Get-BatteryValue -htmlFilePath $batteryReportPath -valueToFind "FULL CHARGE CAPACITY"

                $designNumber = [double]($designCapacity -replace ',', '' -replace ' mWh', '')
                $fullChargeNumber = [double]($fullChargeCapacity -replace ',', '' -replace ' mWh', '')

                $DesignRound = $designNumber / 1000
                $FullChargeRound = $fullChargeNumber / 1000

                $batt = "$([math]::Round($FullChargeRound)) / $([math]::Round($DesignRound))"
                $BatPer = $FullChargeRound / $DesignRound

                if ($BatPer -gt .5)
                    {
                        Write-Host "Battery Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                        Write-Host " Pass " -ForegroundColor Green -NoNewline
                        Write-Host $Batt
                    }
                else
                    {
                        Write-Host "Battery Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                        Write-Host " Fail " -ForegroundColor Red -NoNewline
                        Write-Host $Batt
                    }


                Remove-Item -Path C:\Temp\batt.html
            }

        Function Windows-Update
            {
                Install-Module -Name PSWindowsUpdate -Force
                Import-Module PSWindowsUpdate
                Get-WindowsUpdate | Out-Null
                Install-WindowsUpdate -AcceptAll
                Write-Host "Done."

            }

        PC-Stats
        Software-Check
        HD-Test
        Wifi-Test
        Battery-Test
        Write-Host "Running Windows Updates." -BackgroundColor White -ForegroundColor Black
        Windows-Update
        Write-Host "Press R to reboot and complete Windows Updates." -BackgroundColor White -ForegroundColor Black

        do
            {
        $keypress = Read-Host ">"
            } until  ($keypress -eq 'r' -or $keypress -eq "R")

        if ($keypress -eq 'r' -or $keypress -eq "R")
            {
                Restart-Computer
            }

    }
Function Reboot-PC
    {
        Write-Host "REBOOTING WORKSTATION"
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Rebooted."
        Start-Sleep -Seconds 10 ; Restart-Computer -Force
    }

Function Delete-Self
    {
        Remove-Item C:\Temp\kfixes.ps1
    } 

Switch ($run)
    {
        Default
            {
                Write-Host "Unrecognized Command. Run 'help' for a list of commands."
            }
        repair
            {
                Net-Cleanup
                Windows-Repair
                Driver-Update
                Run-WinUpdate
            }
        lightrepair
            {
                Windows-Repair
            }
        cache
            {
                Net-Cleanup
            }
        winupdate
            {
                Run-WinUpdate
            }
        hpdrivers
            {
                Driver-Update
            }
        driverfix
            {
                Reinstall-Drivers
            }
        officerep
            {
                Repair-Office
            }
        slackcache
            {
                Clear-Slack
            }
        postimage
            {
                Post-Image
            }
        network
            {
                Reset-Network
            }
        printq
            {
                Clear-PrintQueue
            }
        errorlog
            {
                Get-RecentErrors
            }
        pkill
            {
                P-Kill
            }

    }
Switch ($opt)
    {
    Default
            {
                Write-Host "DONE"
                Delete-Self
            }
    auto    
            {
                Delete-Self
                Reboot-PC
            }
                
    reboot
            {
                Write-Host "DONE"
                Reboot-PC
            }
    delete
            {
                Write-Host "DONE"
                Delete-Self
            } 
    }
'@


Function Net-Cleanup
    {    
        if ($LOCUSR -ne "nobody.here")
            {
                Get-Process -Name "*msedge" | Stop-Process -Force -ErrorAction SilentlyContinue
                Get-Process -Name "*Chrome*" | Stop-Process -Force -ErrorAction SilentlyContinue

                Write-Host "CLEARING INTERNET CACHE"
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Cache\Cache_Data\*" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Network\Cookies" -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Network\Cookies-journal" -Force -ErrorAction SilentlyContinue
                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Internet cache."

                Write-Host "CLEARING CHROME CACHE"
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cookies" -Recurse -Force -ErrorAction SilentlyContinue 
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Media Cache" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cookies-Journal" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Cookies" -Force -ErrorAction SilentlyContinue 
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Media Cache" -Force -ErrorAction SilentlyContinue
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Cookies-Journal" -Force -ErrorAction SilentlyContinue
                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Chrome cache."
                Write-Host "DONE"

                Remove-Item C:\users\$LOCUSR\AppData\Local\Temp\* -Recurse -Force -ErrorAction SilentlyContinue
            }

        Write-Host "CLEARING LOCAL CACHE"
        Remove-Item C:\Windows\Temp\* -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item C:\Windows\Prefetch\* -Recurse -Force -ErrorAction SilentlyContinue
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared local cache."
        Write-Host "DONE"

        Write-Host "FLUSHING DNS"
        cmd.exe /c ipconfig /flushdns
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Flushed DNS."
        Write-Host "DONE"
    }

Function Windows-Repair
    {
        Write-Host "REPAIRING WINDOWS IMAGE"
        $dismtimer = "Start-Sleep -Seconds 600
        cmd /c 'taskkill /IM dism.exe /F'"

        $dismtimer | Out-File "C:\temp\dt.ps1"
        Start-Sleep -Seconds 3
        Start-Job -FilePath C:\temp\dt.ps1 | Out-Null
        dism /online /cleanup-image /restorehealth
        Remove-Item "C:\temp\dt.ps1" 
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran DISM."
        Write-Host "DONE."

        Write-Host "RUNNING SYSTEM FILE CHECK"
        sfc /scannow
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran SFC."
        Write-Host "DONE"

        Write-Host "REPAIRING MICROSOFT COMPONENTS"
        Get-AppXPackage | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml"} -ErrorAction SilentlyContinue
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Re-registered Microsoft components."
        Write-Host "DONE"

        Write-Host "UPDATING GROUP POLICY"
        cmd.exe /c echo n | gpupdate /force /wait:0
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran GPUpdate."
        Write-Host "DONE"   
    }
Function Run-WinUpdate
    {    
        Write-Host "Running Windows Updates."    
        Install-Module -Name PSWindowsUpdate -Force
        Import-Module PSWindowsUpdate
        Get-WindowsUpdate | Out-Null
        Install-WindowsUpdate -AcceptAll
        Write-Host "Done."
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran Windows Updates."

    }
Function Reset-Network
    {
        Write-Host "Resetting Winsock Catalog..."
        cmd.exe /c "netsh winsock reset"

        Write-Host "Resetting TCP/IP Stack..."
        cmd.exe /c "netsh int ip reset"

        Write-Host "Releasing and Renewing IP Address..."
        cmd.exe /c "ipconfig /release"
        cmd.exe /c "ipconfig /renew"

        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Reset Winsock, TCP/IP Stack, and renewed IP address."
    }
Function Clear-PrintQueue
    {
        Write-Host "Stopping the Print Spooler service."
        Stop-Service -Name Spooler -Force
        
        $spoolPath = "C:\Windows\System32\spool\PRINTERS\*"
        Write-Host "Clearing files from $spoolPath"
        try
            {
                Remove-Item -Path $spoolPath -Force -ErrorAction Stop
                Write-Host "Print queue cleared successfully."
                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared print queue."
            }
        catch
            {
                Write-Host "Error clearing some files."
                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Failed to clear print queue."
            }
        
        Write-Host "Restarting the Print Spooler service."
        Start-Service -Name Spooler
    }
Function Driver-Update
    {
            Function HP-Update
        {
            $download = 'C:\Temp\HPIA\Download'
            $report = 'C:\Temp\HPIA\Report'
            $log = 'C:\Temp\HPIA\Log'
        
            Function Run-HPIA
                {
                    Write-Host "Analyzing Drivers"
                    try
                        {
                            Start-Process "C:\Program Files\HP\HPIA\HPImageAssistant.exe" -ArgumentList "/Operation:Analyze /Action:List /Silent /ReportFolder:$report" -wait
                        }
                    catch
                        {
                            Write-Host "Error analyzing drivers."
                            Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Unable to run Driver Upgrades."
                            exit
                        }
                    $jsonFile = Get-ChildItem $report -Filter "*.json" | Where-Object { $_.PSIsContainer -eq $false }
                    $list = Get-Content $jsonFile.FullName | Where-Object {$_ -match '"Name":' -or $_ -match '#'}
                    $dreport = $list -replace '"Name":', "Upgrading Driver:" -replace ","," " -replace '"', " "
        
                    $dreport
                    try
                        {
                            Start-Process "C:\Program Files\HP\HPIA\HPImageAssistant.exe" -ArgumentList "/Operation:Analyze /Action:Install /Category:BIOS,Drivers,Firmware,Software /Selection:All /Noninteractive /SoftPaqDownloadFolder:$download /ReportFolder:$report /AutoCleanup /LogFolder:$log" -wait
                        }
                    catch
                        {
                            Write-Host "Error upgrading drivers."
                            Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Unable to run Driver Upgrades."
                            exit
                        }

                    Write-Host "Drivers Upgraded."
                    Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Ran Driver Upgrades."
                    $dreport[0] = " $($dreport[0])"
                    Add-Content -Path C:\Temp\k-log.txt -Value "$($dreport | ForEach-Object { "$_`n" -replace 'Upgrading Driver:', 'Upgraded Driver:' })"
                }    
            Function Setup-HPIA
                {
                    Function Install-HPIA
                        {
                            try
                                {
                                    Start-Process "C:\Temp\hp-hpia-5.3.2.exe" -ArgumentList '/s /e /f "C:\Program Files\HP\HPIA"'
                                }
                            catch
                                {
                                    Write-Host "Error installing HPIA."
                                    Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to install HPIA."
                                }

                            Do {Start-Sleep -Seconds 5}
                            until (Test-Path -path "C:\Program Files\HP\HPIA\HPImageAssistant.exe" -PathType Leaf)
                            
                            Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Set up HPIA."
                            Write-Host "Queuing Driver and Firmware upgrades."

                        }
                    Write-Host "Preparing Device Upgrades..."
                    If (Test-Path -path "C:\Temp\hp-hpia-5.3.2.exe")
                        {          
                            Install-HPIA
                        }
                    else
                    {
                        try
                            {
                                Invoke-WebRequest -Uri https://hpia.hpcloud.hp.com/downloads/hpia/hp-hpia-5.3.2.exe -OutFile C:\Temp\hp-hpia-5.3.2.exe | Out-Null
                            }
                        catch
                            {
                                Write-Host "Error downloading HPIA."
                                Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to download HPIA."
                                exit
                            }
                            Do {Start-Sleep -Seconds 5}
                            until (Test-Path -path "C:\Temp\hp-hpia-5.3.2.exe" -PathType Leaf)

                            $hash  = "E1B8214B0C0B3B6B4C37181CAC1914B4E16FA60E6B063A7F76A1D7A96908E1F0"
                            $fh = Get-FileHash -Path "C:\Temp\hp-hpia-5.3.2.exe" | Select-Object -ExpandProperty Hash

                            If ($hash -eq $fh)
                                {
                                    Install-HPIA
                                }
                            else
                                {
                                    Write-Host "Hash Mismatch."
                                    Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to download HPIA, hash mismatch."
                                    exit
                                }

                    }
                }
        
            If (Test-Path -path "C:\Program Files\HP\HPIA\HPImageAssistant.exe")
                {
                Run-HPIA
                }
            Else
                {
                Setup-HPIA
                Run-HPIA
                }

        }


    if ((Get-WmiObject -Class Win32_ComputerSystem).Manufacturer -eq "HP")
        {
            HP-Update
        }
    else
        {
            Write-Host "Non-HP system, please run updates manually."    
        }
    }

Function Reinstall-Drivers
    {
        Write-Host "Reinstall which drivers?

1. Audio
2. Bluetooth
3. Cameras
4. Graphics Display
5. Wi-Fi
6. Ethernet
7. USB Controller
8. Cancel
"
$select = Read-Host "Select Number"

Switch ($select)
    {
        Default
            {
                Write-Host "Unrecognized Selection."
                exit
            }
        1
            {
              Set-Variable -name drvr -Value "audio"
            }
        2
            {
              Set-Variable -name drvr -Value "bluetooth"
            }
        3
            {
              Set-Variable -name drvr -Value "camera|webcam"
            }
        4
            {
              Set-Variable -name drvr -Value "graphics"
            }
        5
            {
              Set-Variable -name drvr -Value "Wi-Fi"
            }
        6
            {
              Set-Variable -name drvr -Value "Ethernet"
            }
        7
            {
              Set-Variable -name drvr -Value "eXtensible"
            }
        8
            {
              exit
            }
      }

$drivers = Get-PnpDevice
$duplicateDrivers = $drivers | Group-Object -Property HardwareID | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty Group

foreach ($duplicateDriver in $duplicateDrivers)
  {
    $duplicateDriver | Select-Object -Skip 1 | ForEach-Object { &"pnputil" /remove-device $_.InstanceId }
  }

  Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Checked for and uninstalled duplicate or corrupt device drivers."
        
foreach ($dev in (Get-PnpDevice | Where-Object{$_.Name -Match "$drvr"}))
  {
    &"pnputil" /remove-device $dev.InstanceId
  }

  Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Reinstalled $drvr drivers."

Start-Process "cmd.exe" -ArgumentList " /c C:\Windows\System32\pnputil.exe /scan-devices" -Wait -ErrorAction SilentlyContinue
Write-Host "DONE"
    }

Function Repair-Office
    {
        Get-Process -Name "WINWORD", "EXCEL", "POWERPNT", "OUTLOOK" -ErrorAction SilentlyContinue | Stop-Process -Force
        $ExePath = "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe"

        if ($ExePath)
            {
                Write-Host "REPAIRING OFFICE"
                try
                    {
                        cmd.exe /C  "C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeC2RClient.exe" scenario=Repair platform=x86 culture=en-us RepairType=FullRepair forceappshutdown=True DisplayLevel=False
                        Write-Host "DONE"
                        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Ran Online Office Repair."
                    }
                catch
                    {
                        Write-Error "Unable to run Office Repair."
                    }
            }
        else
            {
                Write-Error "Repair Client not found."
            }
    }

Function Clear-Slack
    {
        if ($LOCUSR -eq "nobody.here")
            {
                $LOCUSR = Read-Host "Enter username of Slack account to clear"
            }
        Get-Process -Name *Slack* | Stop-Process -Force -ErrorAction SilentlyContinue
        Remove-Item -Recurse -Path "C:\Users\$LOCUSR\AppData\Roaming\Slack\Cache\Cache_Data"
        Remove-Item -Recurse -Path "C:\Users\$LOCUSR\AppData\Roaming\Slack\Code Cache\js"
        Remove-Item -Recurse -Path "C:\Users\$LOCUSR\AppData\Roaming\Slack\Code Cache\wasm"
        Remove-Item -Recurse -Path "C:\Users\$LOCUSR\AppData\Roaming\Slack\GPUCache"
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Slack cache." 
    }

Function Get-RecentErrors
    {
        Write-Host "--- Recent Application Errors ---"
        Get-WinEvent -LogName Application -FilterXPath "*[System[Level=1 or Level=2]]" -MaxEvents 10 | 
            Select-Object TimeCreated, ProviderName, Message | Format-List

        Write-Host "--- Recent System Errors ---"
        Get-WinEvent -LogName System -FilterXPath "*[System[Level=1 or Level=2]]" -MaxEvents 10 | 
            Select-Object TimeCreated, ProviderName, Message | Format-List
    }
Function Wlan-Report
    { 
        cmd /c netsh wlan show wlanreport
        Start-Sleep -Seconds 12
        start-process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" C:\ProgramData\Microsoft\Windows\WlanReport\wlan-report-latest.html
    }

Function Battery-Report
    {
        $HN = hostname
        powercfg /batteryreport /output C:\temp\$HN-battery_report.html
        start-process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" C:\temp\$opt-battery_report.html      
    }
Function Get-Notes
    {
        Clear-Content C:\temp\tix.txt
        (Get-Content -Path "\\$opt\C$\Temp\k-log.txt" -Raw) -split '%' | 
            Select-Object -Last 1 | 
            ForEach-Object { 
                ($_ -split "`n") | 
                ForEach-Object { 
                    if ($_ -notlike "*Upgraded Driver: *") {
                        if ($_.Length -gt 15) {
                            $_.Substring(15) 
                        } else {
                            $_
                        }
                    } else {
                        $_
                    }
                }
            } | 
            Out-File c:\temp\tix.txt
            Start-Process 'C:\Windows\Notepad.exe' C:\temp\tix.txt
        Exit
    }

Function P-Kill
    {
        Get-Process -Name "*$opt*" | Stop-Process -Force
    }
Function Post-Image
    {
        Clear-Host
        Function PC-Stats {
            $PCName = hostname
            $Ram = [Math]::Round((Get-CimInstance -ClassName Win32_ComputerSystem).TotalPhysicalMemory / 1GB)
            $CPU = Get-CimInstance -ClassName Win32_Processor | Select-Object Name
            $BIOS = (Get-WmiObject win32_bios).SMBIOSBIOSVersion.Split(' ')[-1].TrimStart('0')

                                        if ($CPU.Name -like "*i7*")
                                            {
                                                $inum = "i7"
                                            }
                                        elseif ($CPU.Name -like "*i3*")
                                            {
                                                $inum = "i3"
                                            }
                                        else
                                            {
                                                $inum = "Unknown"
                                            }

                                    Write-Host "PC Name:" -BackgroundColor White -ForegroundColor Black -NoNewLine
                                    Write-Host " $PCName"
                                    Write-Host "PC Stats:"-BackgroundColor White -ForegroundColor Black -NoNewLine
                                    Write-Host " $Ram GB  /  $inum"
                                    Write-Host "BIOS Updates:"-BackgroundColor White -ForegroundColor Black -NoNewLine
                                    Write-Host " Version $BIOS"

                        }


        Function Software-Check {

            $Crowdstrike = Test-Path -path "C:\Program Files\CrowdStrike"
            $Drive = Test-Path -path "C:\Program Files\Google\Drive File Stream"
            $Bit = Get-BitLockerVolume -MountPoint C: | Select-Object ProtectionStatus

                                    if ($Bit.ProtectionStatus -like "*On*")
                                    {
                                        Write-Host "Bitlocker ON:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " True" -ForegroundColor Green
                                    }
                                else
                                    {
                                        Write-Host "Bitlocker ON:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " False" -ForegroundColor Red
                                    }

            
                                if ($Crowdstrike)
                                    {
                                        Write-Host "CrowdStrike / Ninja:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " Pass" -ForegroundColor Green
                                    }
                                else
                                    {
                                        Write-Host "CrowdStrike / Ninja:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " Fail" -ForegroundColor Red
                                    }

                                if ($Drive)
                                    {
                                        Write-Host "Google Drive Installed:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " Pass" -ForegroundColor Green
                                    }
                                else
                                    {
                                        Write-Host "Google Drive Installed:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                        Write-Host " Fail" -ForegroundColor Red
                                    }
            
                

                        }
            
        Function HD-Test {

            $HDSize = Get-PhysicalDisk | Select-Object @{Name='Size(GB)'; Expression={[math]::Round($_.Size / 1GB, 2)}} | Select-Object -ExpandProperty 'Size(GB)'
            $HDType = Get-PhysicalDisk | Select-Object FriendlyName, Manufacturer

                        if ($HDType.Manufacturer -ne  $null)
                            {
                                $HDBrand = $HDType.Manufacturer
                            }
                        elseif ($HDType.Manufacturer -eq  $null)
                            {
                                $HDBrand = Get-PhysicalDisk | Select-Object -ExpandProperty FriendlyName | ForEach-Object { $_.Substring(0, 3) }
                            }
                        else
                            {
                                $HDBrand = "Unknown"
                            }

            Write-Host "SSD/Size Test:"-BackgroundColor White -ForegroundColor Black -NoNewline
            Write-Host " Brand - $HDBrand    Size - $HDSize"

                        }

        Function Wifi-Test {
            $wifiAdapter = Get-NetAdapter -Name *Wi-Fi* -ErrorAction SilentlyContinue

                                if ($wifiAdapter) {
                                    Write-Host "Wi-Fi Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                    Write-Host " Pass" -ForegroundColor Green
                                }
                                else {
                                    Write-Host "Wi-Fi Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                    Write-Host " Fail" -ForegroundColor Red
                                }
                            }
            

        Function Typing-Test
            {
                $ttest = "the quick brown fox jumps over the lazy dog"
                $typed = Read-Host "Type '$ttest'"

                if ($typed -eq $ttest)
                    {
                        Write-Host "Keyboard Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                        Write-Host " Pass" -ForegroundColor Green
                    }
                else
                    {
                        do
                            {
                                $typed = Read-Host "Try again, or press s to skip"
                                if ($typed -eq 's') { break }    
                            } while ($typed -ne $ttest)

                        if ($typed -eq $ttest)
                            {
                                Write-Host "Keyboard Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                Write-Host " Pass" -ForegroundColor Green
                            }
                        else
                            {
                                Write-Host "Keyboard Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                                Write-Host " Fail" -ForegroundColor Red
                            }
                }
            }

        Function Battery-Test
            {

                powercfg /batteryreport -output C:\Temp\batt.html | Out-Null
                    
                    function Get-BatteryValue
                        {
                            param
                                (
                                    [Parameter(Mandatory=$true)]
                                    [string]$htmlFilePath,
                                    [Parameter(Mandatory=$true)]
                                    [string]$valueToFind
                                )

                            $htmlContent = Get-Content -Path $htmlFilePath -Raw
                            $pattern = "(?s)<span class=`"label`">$valueToFind</span></td><td>(?<value>[\d,]+ mWh)"
                            $match = [regex]::Match($htmlContent, $pattern)

                                return $match.Groups["value"].Value.Trim()
                    }
        `

                $batteryReportPath = "C:\Temp\batt.html"

                $designCapacity = Get-BatteryValue -htmlFilePath $batteryReportPath -valueToFind "DESIGN CAPACITY"
                $fullChargeCapacity = Get-BatteryValue -htmlFilePath $batteryReportPath -valueToFind "FULL CHARGE CAPACITY"

                $designNumber = [double]($designCapacity -replace ',', '' -replace ' mWh', '')
                $fullChargeNumber = [double]($fullChargeCapacity -replace ',', '' -replace ' mWh', '')

                $DesignRound = $designNumber / 1000
                $FullChargeRound = $fullChargeNumber / 1000

                $batt = "$([math]::Round($FullChargeRound)) / $([math]::Round($DesignRound))"
                $BatPer = $FullChargeRound / $DesignRound

                if ($BatPer -gt .5)
                    {
                        Write-Host "Battery Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                        Write-Host " Pass " -ForegroundColor Green -NoNewline
                        Write-Host $Batt
                    }
                else
                    {
                        Write-Host "Battery Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                        Write-Host " Fail " -ForegroundColor Red -NoNewline
                        Write-Host $Batt
                    }


                Remove-Item -Path C:\Temp\batt.html
            }


        Function Trackpad-Test
            {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Trackpad Test"
        $form.Size = New-Object System.Drawing.Size(350, 150)
        $form.StartPosition = "CenterScreen"

        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10, 20)
        $label.Size = New-Object System.Drawing.Size(300, 20)
        $label.Text = "Use the Trackpad to click the button."
        $form.Controls.Add($label)

        $button = New-Object System.Windows.Forms.Button
        $button.Location = New-Object System.Drawing.Point(100, 60)
        $button.Size = New-Object System.Drawing.Size(150, 30)
        $button.Text = "Trackpad Test: Pass"

        $button.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.AcceptButton = $button
        $form.Controls.Add($button)

        $result = $form.ShowDialog()

        if ($result -eq [System.Windows.Forms.DialogResult]::OK)
            {
                Write-Host "Trackpad Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                Write-Host " Pass" -ForegroundColor Green
            }
        else
            {
                Write-Host "Trackpad Test:" -BackgroundColor White -ForegroundColor Black -NoNewline
                Write-Host " Fail" -ForegroundColor Red
            }

        }

        Function Webcam-Test
            {
            Start-Process -FilePath "C:\Program Files\Google\Chrome\Application\chrome.exe" -ArgumentList "https://webcamtests.com/"
            }

        Function Windows-Update
            {
                Install-Module -Name PSWindowsUpdate -Force
                Import-Module PSWindowsUpdate
                Get-WindowsUpdate | Out-Null
                Install-WindowsUpdate -AcceptAll
                Write-Host "Done."

            }

        PC-Stats
        Software-Check
        HD-Test
        Wifi-Test
        Battery-Test
        Trackpad-Test
        Typing-Test
        Write-Host "Running Windows Updates. Please test the webcam and check the USB ports while they install." -BackgroundColor White -ForegroundColor Black
        Webcam-Test
        Windows-Update
        Write-Host "Press R to reboot and complete Windows Updates. Check for Global Protect VPN on the login screen afterwards." -BackgroundColor White -ForegroundColor Black

        do
            {
        $keypress = Read-Host ">"
            } until  ($keypress -eq 'r' -or $keypress -eq "R")

        if ($keypress -eq 'r' -or $keypress -eq "R")
            {
                Restart-Computer
            }

    }
Function kfixes-Remote
    {
        if (-not(Test-Path -path C:\PsExec.exe -PathType Leaf))
            {
                    Write-Host "Preparing Remote Execution."
                    try
                        {
                            Invoke-WebRequest -Uri https://download.sysinternals.com/files/PSTools.zip -OutFile C:\Temp\PSTools.zip
                        }
                    catch
                        {
                            Write-Host "Unable to setup PsExec for remote execution, download error."
                            exit
                        }

                    $ehash  = "A9CA77DFE03CE15004157727BB43BA66F00CEB215362C9B3D199F000EDAA8D61"
                    $efh = Get-FileHash -Path "C:\Temp\PSTools.zip" | Select-Object -ExpandProperty Hash

                    If ($ehash -ne $efh)
                        {
                            Write-Host "Unable to setup PsExec for remote execution, hash mismatch."
                            exit
                        }
                    Expand-Archive -Path "C:\Temp" -DestinationPath "C:\"
                    Remove-Item -Path C:\Temp\PSTools.zip
            }

        if($runr -eq "repair")
            {
                Set-Variable -name Command -Value "repair"
            }
        elseif($runr -eq "lightrepair")
            {
                Set-Variable -name Command -Value "lightrepair"
            }
        elseif($runr -eq "cache")
            {
                Set-Variable -name Command -Value "cache"
            }
        elseif($runr -eq "winupdate")
            {
                Set-Variable -name Command -Value "winupdate"
            }
        elseif($runr -eq "printq")
            {
                Set-Variable -name Command -Value "printq"
            }
        elseif($runr -eq "network")
            {
                Set-Variable -name Command -Value "network"
            }
        elseif($runr -eq "hpdrivers")
            {
                Set-Variable -name Command -Value "hpdrivers"
            }
        elseif($runr -eq "driverfix")
    {
        if (Test-Path -path \\$opt\c$\Temp)
            {
                $kfixes | Out-File \\$opt\C$\Temp\kfixes.ps1
            }
        else
            {
                Write-Host "HOST NOT FOUND."
                exit
            }
        Write-Host "Reinstall which drivers?

1. Audio
2. Bluetooth
3. Cameras
4. Graphics Display
5. Wi-Fi
6. Ethernet
7. USB Controller
8. Cancel
"
            $select = Read-Host "Select Number"

            Switch ($select)
                {
                    Default
                        {
                            Write-Host "Unrecognized Selection."
                            exit
                        }
                    1
                        {
                        Set-Variable -name drvr -Value "audio"
                        }
                    2
                        {
                        Set-Variable -name drvr -Value "bluetooth"
                        }
                    3
                        {
                        Set-Variable -name drvr -Value "camera|webcam"
                        }
                    4
                        {
                        Set-Variable -name drvr -Value "graphics"
                        }
                    5
                        {
                        Set-Variable -name drvr -Value "Wi-Fi"
                        }
                    6
                        {
                        Set-Variable -name drvr -Value "Ethernet"
                        }
                    7
                        {
                        Set-Variable -name drvr -Value "eXtensible"
                        }
                    8
                        {
                        exit
                        }
                }

                    (Get-Content \\$opt\C$\Temp\kfixes.ps1) -replace "##DRVR##", $drvr | Set-Content \\$opt\C$\Temp\kfixes.ps1
                    Set-Variable -name Command -Value "driverfix"
            }
        elseif($runr -eq "officerep")
            {
                Set-Variable -name Command -Value "officerep"
            }
        elseif($runr -eq "slackcache")
            {
                Set-Variable -name Command -Value "slackcache"
            }
        elseif($runr -eq "postimage")
            {
                Set-Variable -name Command -Value "postimage"
            }
        elseif($runr -eq "errorlog")
            {
                Set-Variable -name Command -Value "errorlog"
            }
        elseif($runr -eq "wlan")
            {
                $optPath = "\\$opt\c$\ProgramData\Microsoft\Windows\WlanReport\*"
                Start-Sleep -Seconds 15
                [void] (New-PSDrive -ErrorAction "Stop" -name "WlanTestDrive" -root "\\$opt\C$" -PSProvider FileSystem -Scope Local)
                C:\PsExec.exe \\$opt netsh wlan show wlanreport
                if (-not (Test-Path -Path 'C:\Temp\WLanReport'))
                    {
                        New-Item 'C:\Temp\WLanReport' -ItemType Directory
                    }
                Copy-Item -path $optPath -Destination "C:\temp\WlanReport\" | Out-Null
                start-process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" C:\Temp\WLanReport\wlan-report-latest.html 
                Remove-PSDrive -name "WlanTestDrive"
                exit
            }
        elseif($runr -eq "battery")
            {
                [void] (New-PSDrive -ErrorAction "Stop" -name "BatTestDrive" -root "\\$opt\C$" -PSProvider FileSystem -Scope Local)
                C:\PsExec.exe \\$opt powercfg /batteryreport /output C:\temp\$opt-battery_report.html
                Copy-Item \\$opt\c$\temp\$opt-battery_report.html -Destination c:\temp
                start-process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" C:\temp\$opt-battery_report.html
                Remove-PSDrive -name "BatTestDrive"
                exit
            }
        elseif($runr -eq "pkill")
            {
            C:\PsExec.exe \\$opt powershell.exe "Get-Process -Name "*$optr*" | Stop-Process -Force"
            }
        elseif($runr -eq "progress")
            {
                cmd /c "ipconfig /flushdns"
                    if (Test-Path -path \\$opt\c$\Temp)
                        {
                            (Get-Content -Path "\\$opt\C$\Temp\k-log.txt" -Raw) -split '%' | Select-Object -Last 1
                            exit
                        }
                    else
                        {
                        Write-Host "Waiting for Host..."
                            Do {
                                    Start-Sleep -seconds 30
                                    Write-Host "Still Waiting..."
                                }
                            Until (Test-Path -path \\$opt\c$\Temp)
                            Get-Content \\$opt\C$\Temp\k-log.txt
                            exit
                        }
            }
        else
            {
                Set-Variable -name Command -Value ""
                Write-Host "Unrecognized Command. Run 'help' for a list of commands."
                exit
            }

        if($optr -eq "auto")
            {
                Set-Variable -name Option -Value "auto"
            }
        elseif($optr -eq "reboot")
            {
                Set-Variable -name Option -Value "reboot"
            }
        elseif($optr -eq "delete")
            {
                Set-Variable -name Option -Value "delete"
            }
        else
            {
                Set-Variable -name Option -Value ""
            }        

        if (Test-Path -path \\$opt\c$\Temp)
            {   
                if ($null -eq $select)
                    {
                        $kfixes | Out-File \\$opt\C$\Temp\kfixes.ps1
                    }
            }
        else
            {
                Write-Host "HOST NOT FOUND."
                exit
            }

        Add-Content -Path "\\$opt\C$\Temp\k-log.txt" "`n%"
        C:\PsExec.exe \\$opt powershell.exe -ExecutionPolicy RemoteSigned -command "c:\temp\kfixes.ps1 $Command $Option"
        Clear-Variable -Name "Command"
        Clear-Variable -Name "Option"

        if (Test-Path -path \\$opt\c$\Temp)
            {
                Clear-Content C:\temp\tix.txt
                (Get-Content -Path "\\$opt\C$\Temp\k-log.txt" -Raw) -split '%' | 
                    Select-Object -Last 1 | 
                    ForEach-Object { 
                        ($_ -split "`n") | 
                        ForEach-Object { 
                            if ($_ -notlike "*Upgraded Driver: *") {
                                if ($_.Length -gt 15) {
                                    $_.Substring(15) 
                                } else {
                                    $_
                                }
                            } else {
                                $_
                            }
                        }
                    } | 
                    Out-File c:\temp\tix.txt
                Start-Process 'C:\Windows\Notepad.exe' C:\temp\tix.txt
            }
        else
            {
            Write-Host "CONNECTION LOST."
            }
        
    exit
    
    }
Function Reboot-PC
    {
        Write-Host "REBOOTING WORKSTATION"
        Add-Content -Path C:\Temp\k-log.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Rebooted."
        Set-ExecutionPolicy -ExecutionPolicy Default
        Start-Sleep -Seconds 10 ; Restart-Computer -Force
    }

Function Delete-Self
    {
        Set-ExecutionPolicy -ExecutionPolicy Default
        Remove-Item C:\Temp\kfixes.ps1
    } 
Function Zatanna-Translate
    {
        "$opt"[-1..-$opt.Length] -join ''
    }


Switch ($run)
    {
        Default
            {
                Write-Host "Unrecognized Command. Run 'help' for a list of commands."
            }
        repair
            {
                Net-Cleanup
                Windows-Repair
                Driver-Update
                Run-WinUpdate
            }
        lightrepair
            {
                Windows-Repair
            }
        cache
            {
                Net-Cleanup
            }
        winupdate
            {
                Run-WinUpdate
            }
        hpdrivers
            {
                Driver-Update
            }
        driverfix
            {
                Reinstall-Drivers
            }
        network
            {
                Reset-Network
            }
        officerep
            {
                Repair-Office
            }
        slackcache
            {
                Clear-Slack
            }
        printq
            {
                Clear-PrintQueue
            }
        wlan
            {
                Wlan-Report
            }
        battery
            {
                Battery-Report
            }
        pkill
            {
                P-Kill
            }
        postimage
            {
                Post-Image
            }
        errorlog
            {
                Get-RecentErrors
            }
        zatanna
            {
                Zatanna-Translate
            }
        remote
            {
                kfixes-Remote
            }
        notes
            {
                Get-Notes
            }
        help
            {
                Write-Host "------COMMANDS------
            repair		Clear cache (browser and local), run Windows repairs, HP Driver updates, and Windows updates
            lightrepair	Run Windows repairs only
            cache		Clear cache (browser and local)
            winupdate	Run Windows updates
            network     Resets Winsock, TCP/IP stack, and renews IP address
            hpdrivers	Run HP Imaging Assistant to automatically update drivers (HP only)
            driverfix   Reinstall drivers for specific hardware (Does not download new drivers, only reinstalls the current drivers.) 
            officerep	Run Online Office repair
            printq      Clear printer queue
            slackcache	Clear Slack cache
            pkill		Kills a process by name Syntax: c:\temp\kfixes.ps1 pkill chrome
            errorlog    Displays recent Application and System errors
            wlan		Displays wifi connection log
            battery		Displays Windows battery report
            postimage	Runs PostImage script; hardware tests are skipped if run remotely
            remote		Executes the script on a remote machine. Syntax: c:\temp\kfixes.ps1 remote HOSTNAME command flag
            progress	Checks to see if a remote machine is back online after a network disconnect, and shows the progress of the script once it is.
            notes	After running the script on a TM's machine, run this locally to generate ticket notes based on the script's logs. Syntax: c:\temp\kfixes.ps1 notes HOSTNAME
            ------FLAGS------
            delete		Deletes script after execution, but doesn't reboot.
            reboot		Prevents script from deleting itself after execution, then reboots.
            auto		Automatically reboots and self-deletes after execution
           
            ------SYNTAX------
            c:\temp\kfixes.ps1 wlan
            c:\temp\kfixes.ps1 cache delete
            c:\temp\kfixes.ps1 hpdrivers reboot
            c:\temp\kfixes.ps1 repair auto
            c:\temp\kfixes.ps1 remote HOSTNAME cache auto"
            }
  
    }
Switch ($opt)
    {
    Default
            {
                Set-ExecutionPolicy -ExecutionPolicy Default
                Write-Host "DONE"
            }
    auto    
            {
                Delete-Self
                Reboot-PC
            }
                
    reboot
            {
                Write-Host "DONE"
                Reboot-PC
            }
    delete
            {
                Write-Host "DONE"
                Delete-Self
            } 
    }
