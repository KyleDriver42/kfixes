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

$ktool = @'
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

                echo "CLEARING INTERNET CACHE"
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Cache\Cache_Data\*" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Network\Cookies" -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Network\Cookies-journal" -Force -ErrorAction SilentlyContinue
                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Internet cache."

                echo "CLEARING CHROME CACHE"
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cookies" -Recurse -Force -ErrorAction SilentlyContinue 
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Media Cache" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cookies-Journal" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Cookies" -Force -ErrorAction SilentlyContinue 
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Media Cache" -Force -ErrorAction SilentlyContinue
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Cookies-Journal" -Force -ErrorAction SilentlyContinue
                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Chrome cache."
                echo "DONE"

                Remove-Item C:\users\$LOCUSR\AppData\Local\Temp\* -Recurse -Force -ErrorAction SilentlyContinue
            }

        echo "CLEARING LOCAL CACHE"
        Remove-Item C:\Windows\Temp\* -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item C:\Windows\Prefetch\* -Recurse -Force -ErrorAction SilentlyContinue
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared local cache."
        echo "DONE"

        echo "FLUSHING DNS"
        cmd.exe /c ipconfig /flushdns
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Flushed DNS."
        echo "DONE"
    }

Function Windows-Repair
    {
        #This is the dumbest possible way to do this but it's the only way I could get to work in this version of PS.
        echo "REPAIRING WINDOWS IMAGE"
        $dismtimer = "Start-Sleep -Seconds 600
        cmd /c 'taskkill /IM dism.exe /F'"

        $dismtimer | Out-File "C:\temp\dt.ps1"
        Start-Sleep -Seconds 3
        Start-Job -FilePath C:\temp\dt.ps1 | Out-Null
        dism /online /cleanup-image /restorehealth
        Remove-Item "C:\temp\dt.ps1" 
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran DISM."
        echo "DONE."

        echo "RUNNING SYSTEM FILE CHECK"
        sfc /scannow
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran SFC."
        echo "DONE"

        echo "REPAIRING MICROSOFT COMPONENTS"
        Get-AppXPackage -allusers | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue}
        echo "DONE"

        echo "UPDATING GROUP POLICY"
        cmd.exe /c echo n | gpupdate /force /wait:0
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Re-Registered Microsoft components."
        echo "DONE"   
    }
Function Run-WinUpdate
    {    
        Write-Host "Running Windows Updates."    
        Install-Module -Name PSWindowsUpdate -Force
        Import-Module PSWindowsUpdate
        Get-WindowsUpdate | Out-Null
        Install-WindowsUpdate -AcceptAll
        Write-Host "Done."
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran Windows Updates."

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

        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Reset Winsock, TCP/IP Stack, and renewed IP address."
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
                                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Unable to run Driver Upgrades."
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
                                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Unable to run Driver Upgrades."
                                exit
                            }

                        Write-Host "Drivers Upgraded."
                        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Ran Driver Upgrades."
                        $dreport[0] = " $($dreport[0])"
                        Add-Content -Path C:\Temp\ktlog.txt -Value "$($dreport | ForEach-Object { "$_`n" -replace 'Upgrading Driver:', 'Upgraded Driver:' })"
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
                                        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to install HPIA."
                                    }

                                Do {Start-Sleep -Seconds 5}
                                until (Test-Path -path "C:\Program Files\HP\HPIA\HPImageAssistant.exe" -PathType Leaf)
                                
                                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Set up HPIA."
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
                                    Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to download HPIA."
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
                                        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to download HPIA, hash mismatch."
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
                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared print queue."
            }
        catch
            {
                Write-Host "Error clearing some files."
                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Failed to clear print queue."
            }
        
        Write-Host "Restarting the Print Spooler service."
        Start-Service -Name Spooler
    }
Function Repair-Office
    {
        Get-Process -Name *Excel* -ErrorAction SilentlyContinue | Stop-Process -Force 
        Get-Process -Name *Word* -ErrorAction SilentlyContinue | Stop-Process -Force
        $command64 = ' 
        cmd.exe /C "C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe" scenario=Repair platform=x64 culture=en-us RepairType=QuickRepair forceappshutdown=True DisplayLevel=False
        '
        $command86 = ' 
        cmd.exe /C "C:\Program Files\Microsoft Office 15\ClientX86\OfficeClickToRun.exe" scenario=Repair platform=x86 culture=en-us RepairType=QuickRepair forceappshutdown=True DisplayLevel=False
        '
        echo "REPAIRING OFFICE"
        if(Test-Path -Path "C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe")
            {
            Invoke-Expression -Command:$command64
            }
        elseif(Test-PAth -Path "C:\Program Files\Microsoft Office 15\ClientX32\OfficeClickToRun.exe")
            {
                Invoke-Expression -Command:$command86
            }
        echo "DONE"
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran Online Office repair."
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
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Slack cache." 
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
        echo "REBOOTING WORKSTATION"
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Rebooted."
        Start-Sleep -Seconds 10 ; Restart-Computer -Force
    }

Function Delete-Self
    {
        Remove-Item C:\Temp\ktool.ps1
    } 

Switch ($run)
    {
        Default
            {
                echo "Unrecognized Command. Run 'help' for a list of commands."
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

    }
Switch ($opt)
    {
    Default
            {
                echo "DONE"
                Delete-Self
            }
    auto    
            {
                Delete-Self
                Reboot-PC
            }
                
    reboot
            {
                echo "DONE"
                Reboot-PC
            }
    delete
            {
                echo "DONE"
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

                echo "CLEARING INTERNET CACHE"
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Cache\Cache_Data\*" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Network\Cookies" -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Local\Microsoft\Windows\Edge\User Data\Default\Network\Cookies-journal" -Force -ErrorAction SilentlyContinue
                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Internet cache."

                echo "CLEARING CHROME CACHE"
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cookies" -Recurse -Force -ErrorAction SilentlyContinue 
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Media Cache" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -path "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Cookies-Journal" -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Cookies" -Force -ErrorAction SilentlyContinue 
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Media Cache" -Force -ErrorAction SilentlyContinue
                Remove-Item "C:\Users\$LOCUSR\AppData\Google\Chrome\User Data\Default\Network\Cookies-Journal" -Force -ErrorAction SilentlyContinue
                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Chrome cache."
                echo "DONE"

                Remove-Item C:\users\$LOCUSR\AppData\Local\Temp\* -Recurse -Force -ErrorAction SilentlyContinue
            }

        echo "CLEARING LOCAL CACHE"
        Remove-Item C:\Windows\Temp\* -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item C:\Windows\Prefetch\* -Recurse -Force -ErrorAction SilentlyContinue
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared local cache."
        echo "DONE"

        echo "FLUSHING DNS"
        cmd.exe /c ipconfig /flushdns
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Flushed DNS."
        echo "DONE"
    }

Function Windows-Repair
    {
        #This is the dumbest possible way to do this but it's the only way I could get to work in this version of PS.
        echo "REPAIRING WINDOWS IMAGE"
        $dismtimer = "Start-Sleep -Seconds 600
        cmd /c 'taskkill /IM dism.exe /F'"

        $dismtimer | Out-File "C:\temp\dt.ps1"
        Start-Sleep -Seconds 3
        Start-Job -FilePath C:\temp\dt.ps1 | Out-Null
        dism /online /cleanup-image /restorehealth
        Remove-Item "C:\temp\dt.ps1" 
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran DISM."
        echo "DONE."

        echo "RUNNING SYSTEM FILE CHECK"
        sfc /scannow
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran SFC."
        echo "DONE"

        echo "REPAIRING MICROSOFT COMPONENTS"
        Get-AppXPackage -allusers | Foreach {Add-AppxPackage -DisableDevelopmentMode -Register "$($_.InstallLocation)\AppXManifest.xml" -ErrorAction SilentlyContinue}
        echo "DONE"

        echo "UPDATING GROUP POLICY"
        cmd.exe /c echo n | gpupdate /force /wait:0
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Re-Registered Microsoft components."
        echo "DONE"   
    }
Function Run-WinUpdate
    {    
        Write-Host "Running Windows Updates."    
        Install-Module -Name PSWindowsUpdate -Force
        Import-Module PSWindowsUpdate
        Get-WindowsUpdate | Out-Null
        Install-WindowsUpdate -AcceptAll
        Write-Host "Done."
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran Windows Updates."

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

        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Reset Winsock, TCP/IP Stack, and renewed IP address."
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
                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared print queue."
            }
        catch
            {
                Write-Host "Error clearing some files."
                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Failed to clear print queue."
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
                            Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Unable to run Driver Upgrades."
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
                            Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Unable to run Driver Upgrades."
                            exit
                        }

                    Write-Host "Drivers Upgraded."
                    Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format 'MM.dd.yy hh:mm') Ran Driver Upgrades."
                    $dreport[0] = " $($dreport[0])"
                    Add-Content -Path C:\Temp\ktlog.txt -Value "$($dreport | ForEach-Object { "$_`n" -replace 'Upgrading Driver:', 'Upgraded Driver:' })"
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
                                    Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to install HPIA."
                                }

                            Do {Start-Sleep -Seconds 5}
                            until (Test-Path -path "C:\Program Files\HP\HPIA\HPImageAssistant.exe" -PathType Leaf)
                            
                            Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Set up HPIA."
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
                                Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to download HPIA."
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
                                    Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Unable to download HPIA, hash mismatch."
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

Function Repair-Office
    {
        Get-Process -Name *Excel* -ErrorAction SilentlyContinue | Stop-Process -Force 
        Get-Process -Name *Word* -ErrorAction SilentlyContinue | Stop-Process -Force
        $command64 = ' 
        cmd.exe /C "C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe" scenario=Repair platform=x64 culture=en-us RepairType=QuickRepair forceappshutdown=True DisplayLevel=False
        '
        $command86 = ' 
        cmd.exe /C "C:\Program Files\Microsoft Office 15\ClientX86\OfficeClickToRun.exe" scenario=Repair platform=x86 culture=en-us RepairType=QuickRepair forceappshutdown=True DisplayLevel=False
        '
        echo "REPAIRING OFFICE"
        if(Test-Path -Path "C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe")
            {
            Invoke-Expression -Command:$command64
            }
        elseif(Test-PAth -Path "C:\Program Files\Microsoft Office 15\ClientX32\OfficeClickToRun.exe")
            {
                Invoke-Expression -Command:$command86
            }
        echo "DONE"
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Ran Online Office repair."
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
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Cleared Slack cache." 
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
Function Get-Info
    {
        $cinfo= Get-ComputerInfo -Property CsModel, CsDomain, WindowsRegisteredOrganization, OsLocalDateTime, OsLastBootUpTime
        $uinfo= get-ADUser $LOCUSR -Properties passwordlastset, passwordexpired, passwordneverexpires
        echo "COMPUTER INFORMATION"
        echo $cinfo

        echo "USER INFORMATION"
        echo $uinfo
        exit
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
Function Ktool-Remote
    {
        if (-not(Test-Path -path C:\PsExec.exe -PathType Leaf))
            {
                    Write-Host "Preparing Remote Execution."
                    Invoke-WebRequest -Uri https://download.sysinternals.com/files/PSTools.zip -OutFile C:\Temp\PSTools.zip
                    Expand-Archive -Path "C:\Temp" -DestinationPath "C:\"
                    Remove-Item -Path C:\Temp\PSTools.zip
            }

        if($runr -eq "repair")
            {
                Set-Variable -name Command -Value "repair"
            }
        elseif($runr -eq "cache")
            {
                Set-Variable -name Command -Value "cache"
            }
        elseif($runr -eq "winudate")
            {
                Set-Variable -name Command -Value "winupdate"
            }
        elseif($runr -eq "hpdrivers")
            {
                Set-Variable -name Command -Value "hpdrivers"
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
                Set-Variable -name Command -Value "slackcache"
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
                Remove-Item 'C:\Temp\WLanReport' -Recurse
                New-Item 'C:\Temp\WLanReport' -ItemType Directory
                Copy-Item -path $optPath -Destination "C:\temp\WlanReport\"
                start-process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" C:\Temp\WLanReport\wlan-report-latest.html 
                Remove-PSDrive -name "WlanTestDrive"
                Clear-Variable -Name "Command"
                exit
            }
        elseif($runr -eq "battery")
            {
                [void] (New-PSDrive -ErrorAction "Stop" -name "WlanTestDrive" -root "\\$opt\C$" -PSProvider FileSystem -Scope Local)
                C:\PsExec.exe \\$opt powercfg /batteryreport /output C:\temp\$opt-battery_report.html
                Copy-Item \\$opt\c$\temp\$opt-battery_report.html -Destination c:\temp
                start-process "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" C:\temp\$opt-battery_report.html
                Remove-PSDrive -name "WlanTestDrive"
                Clear-Variable -Name "Command"
                exit
            }
        elseif($runr -eq "pkill")
            {
            C:\PsExec.exe \\$opt powershell.exe "Get-Process -Name "*$optr*" | Stop-Process -Force"
            Clear-Variable -Name "Command"
            }
        elseif($runr -eq "notes")
            {
                Clear-Content C:\temp\tix.txt
                $today = Get-Date -Format "MM.dd.yy"
                Get-Content -Path "\\$opt\C$\Temp\ktlog.txt" | Where-Object { $_.StartsWith($today) } | Where-Object { $_ -match "^\d{2}" } | ForEach-Object { $_.Substring(15) } | Out-File c:\temp\tix.txt
                Start-Process 'C:\Windows\Notepad.exe' C:\temp\tix.txt
                Clear-Variable -Name "Command"
                Exit
            }
        elseif($runr -eq "progress")
            {
                cmd /c "ipconfig /flushdns"
                    if (Test-Path -path \\$opt\c$\Temp)
                        {
                            Get-Content \\$opt\C$\Temp\ktlog.txt
                            Clear-Variable -Name "Command"
                            exit
                        }
                    else
                        {
                        echo "Waiting for Host..."
                            Do {
                                    Start-Sleep -seconds 30
                                    echo "Still Waiting..."
                                }
                            Until (Test-Path -path \\$opt\c$\Temp)
                            Get-Content \\$opt\C$\Temp\ktlog.txt
                            Clear-Variable -Name "Command"
                            exit
                        }
            }
        else
            {
                echo "Unrecognized Command. Run 'help' for a list of commands."
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
                $ktool | Out-File \\$opt\C$\Temp\ktool.ps1
            }
        else
            {
                echo "HOST NOT FOUND."
                exit
            }

        C:\PsExec.exe \\$opt powershell.exe -ExecutionPolicy RemoteSigned -command "c:\temp\ktool.ps1 $Command $Option"
        Clear-Variable -Name "Command"
        Clear-Variable -Name "Option"
        Clear-Content C:\temp\tix.txt
        $today = Get-Date -Format "MM.dd.yy"
        Get-Content -Path "\\$opt\C$\Temp\ktlog.txt" | Where-Object { $_.StartsWith($today) } | Where-Object { $_ -match "^\d{2}" } | ForEach-Object { $_.Substring(15) } | Out-File c:\temp\tix.txt
        Start-Process 'C:\Windows\Notepad.exe' C:\temp\tix.txt
    
    }
Function Reboot-PC
    {
        echo "REBOOTING WORKSTATION"
        Add-Content -Path C:\Temp\ktlog.txt -Value "$(Get-Date -Format "MM.dd.yy hh:mm") Rebooted."
        Set-ExecutionPolicy -ExecutionPolicy Default
        Start-Sleep -Seconds 10 ; Restart-Computer -Force
    }

Function Delete-Self
    {
        Set-ExecutionPolicy -ExecutionPolicy Default
        Remove-Item C:\Temp\ktool.ps1
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
                Ktool-Remote
            }
        help
            {
                echo "------COMMANDS------
            repair		Clear cache (browser and local), run Windows repairs, HP Driver updates, and Windows updates
            lightrepair	Run Windows repairs only
            cache		Clear cache (browser and local)
            winupdate	Run Windows updates
            network     Resets Winsock, TCP/IP stack, and renews IP address
            hpdrivers	Run HP Imaging Assistant to automatically update drivers (HP only) 
            officerep	Run Online Office repair
            printq      Clear printer queue
            slackcache	Clear Slack cache
            pkill		Kills a process by name Syntax: c:\temp\ktool.ps1 pkill chrome
            errorlog    Displays recent Application and System errors
            wlan		Displays wifi connection log
            battery		Displays Windows battery report
            postimage	Runs PostImage script; hardware tests are skipped if run remotely
            remote		Executes the script on a remote machine. Syntax: c:\temp\ktool.ps1 remote HOSTNAME command flag
            progress	Checks to see if a remote machine is back online after a network disconnect, and shows the progress of the script once it is.
            notes	After running the script on a TM's machine, run this locally to generate ticket notes based on the script's logs. Syntax: c:\temp\ktool.ps1 notes HOSTNAME
            ------FLAGS------
            delete		Deletes script after execution, but doesn't reboot.
            reboot		Prevents script from deleting itself after execution, then reboots.
            auto		Automatically reboots and self-deletes after execution
           
            ------SYNTAX------
            c:\temp\ktool.ps1 wlan
            c:\temp\ktool.ps1 cache delete
            c:\temp\ktool.ps1 hpdrivers reboot
            c:\temp\ktool.ps1 repair auto
            c:\temp\ktool.ps1 remote HOSTNAME cache auto"
            }
  
    }
Switch ($opt)
    {
    Default
            {
                Set-ExecutionPolicy -ExecutionPolicy Default
                echo "DONE"
            }
    auto    
            {
                Delete-Self
                Reboot-PC
            }
                
    reboot
            {
                echo "DONE"
                Reboot-PC
            }
    delete
            {
                echo "DONE"
                Delete-Self
            } 
    }
