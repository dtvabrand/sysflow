import os, sys, inquirer

try:
    print("By running this script you accept the license terms (ToS/EULA) of all software installed via winget/msstore")
    answers = inquirer.prompt([
        inquirer.Checkbox(
            'tasks',
            message="Choose actions",
            choices=[
                'Windows Updates','Winget','Create daily scheduled task for Windows Update + Winget','Bloatware',
                'Set DNS','Disable DNS','Set Time Zone','Set Short Date Format','Disable Sticky Keys','Show desktop',
                'Enable Middle Button TouchPad','Android SDK Platform Tools','Microsoft Visual Studio Code',
                'Microsoft 365','PC On/Off Time','RedNotebook','Scanner Windows','Snipping Tool','Steam',
                'Telegram Desktop','VLC Media Player','WhatsApp','WinRAR'
            ]
        )
    ], raise_keyboard_interrupt=True)
    if not answers or not answers.get('tasks'): print("No actions selected."); sys.exit(0)

    if 'Windows Updates' in answers['tasks']:
        os.system(r'powershell -NoProfile -ExecutionPolicy Bypass -Command "try { if (-not (Get-Module -ListAvailable PSWindowsUpdate)) { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force; Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; Install-Module -Name PSWindowsUpdate -Scope AllUsers -Force } ; Import-Module PSWindowsUpdate } catch { exit 1 }"')
        os.system('powershell -NoProfile -ExecutionPolicy Bypass -Command "Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -Verbose"')
    
    if 'Winget' in answers['tasks']:
        os.system('winget upgrade --all --accept-source-agreements --accept-package-agreements')

    if 'Create daily scheduled task for Windows Update + Winget' in answers['tasks']:
        rc = os.system(r"""powershell -NoProfile -ExecutionPolicy Bypass -Command "$tn='Windows updates + winget'; if(Get-ScheduledTask -TaskName $tn -ErrorAction SilentlyContinue){ Write-Host 'Task already exists' } else { $p='Start-Sleep 10; try{ if(-not (Get-Module -ListAvailable PSWindowsUpdate)){ Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force; Set-PSRepository -Name PSGallery -InstallationPolicy Trusted; Install-Module PSWindowsUpdate -Scope AllUsers -Force }; Import-Module PSWindowsUpdate } catch{}; try{ Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -Verbose } catch{}; $w=(Get-Command winget.exe -EA 0).Source; if($w){ & $w upgrade --all --silent --accept-source-agreements --accept-package-agreements }'; $b=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($p)); Register-ScheduledTask -TaskName $tn -Principal (New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest) -Trigger (New-ScheduledTaskTrigger -Daily -At 07:00) -Settings (New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable) -Action (New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand '+$b)) | Out-Null }" """)
        if rc == 0: print("Daily task for Windows Update + Winget created (7 am)")
        
    if 'Bloatware' in answers['tasks']:
        bloatware = [
            "Microsoft.GamingApp","Microsoft.XboxApp","Microsoft.Xbox.TCUI","Microsoft.XboxGamingOverlay",
            "Microsoft.XboxGameOverlay","Microsoft.XboxSpeechToTextOverlay","Microsoft.WindowsTerminal",
            "Clipchamp.Clipchamp","Microsoft.PowerAutomateDesktop","DolbyLaboratories.DolbyAccess",
            "Microsoft.OutlookForWindows","Microsoft.RemoteDesktop","Microsoft.ZuneVideo","Microsoft.WindowsPhone",
            "Microsoft.CommsPhone","Microsoft.YourPhone","Microsoft.BingWeather","Microsoft.BingSports",
            "Microsoft.BingNews","Microsoft.BingFinance","Microsoft.MicrosoftSolitaireCollection",
            "Microsoft.MicrosoftStickyNotes","microsoft.windowscommunicationsapps","Microsoft.WindowsSoundRecorder",
            "Microsoft.Todos","Microsoft.WindowsFeedbackHub","Microsoft.WindowsMaps","Microsoft.People",
            "Microsoft.MicrosoftOfficeHub","Microsoft.Office.OneNote","Microsoft.Office.Sway"
        ]
        print("Removed and deprovisioned:")
        for app in bloatware:
            ps = (
                "$ErrorActionPreference='SilentlyContinue'; "
                f"$n='{app}'; "
                "Get-AppxPackage -AllUsers ^| "
                "Where-Object {{ $_.Name -eq $n -or $_.PackageFamilyName -like ($n+'*') }} ^| "
                "Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue; "
                "Get-AppxProvisionedPackage -Online ^| "
                "Where-Object {{ $_.DisplayName -eq $n -or $_.PackageName -like ($n+'*') }} ^| "
                "Remove-AppxProvisionedPackage -Online -AllUsers -ErrorAction SilentlyContinue"
            )
            os.system(f"powershell -NoProfile -ExecutionPolicy Bypass -Command \"{ps}\"")
            print(app)

    if 'Set DNS' in answers['tasks']:
        os.system('powershell -Command "$adapter = Get-NetAdapter | Where-Object { $_.Status -eq \'Up\' }; $ifIndex = $adapter.ifIndex; Set-DnsClientServerAddress -InterfaceIndex $ifIndex -ServerAddresses (\'94.140.14.15\', \'94.140.14.16\', \'2a10:50c0::bad1:ff\', \'2a10:50c0::bad2:ff\')"')
        print("AdGuard DNS server addresses set successfully!")
    
    if 'Disable DNS' in answers['tasks']:
        os.system('powershell -Command "$adapter = Get-NetAdapter | Where-Object { $_.Status -eq \'Up\' }; $ifIndex = $adapter.ifIndex; Set-DnsClientServerAddress -InterfaceIndex $ifIndex -ResetServerAddresses"')
        print("DNS server addresses reset to automatic!")
    
    if 'Set Time Zone' in answers['tasks']:
        os.system('tzutil /s "W. Europe Standard Time"')
        print("Western Europe Standard Time +1 set successfully!")

    if 'Set Short Date Format' in answers['tasks']:
        os.system('powershell -NoProfile -ExecutionPolicy Bypass -Command "Set-ItemProperty -Path ''HKCU:\\Control Panel\\International'' -Name ''sShortDate'' -Value ''d MMM yyyy''"')
        print("Short date format set to: d MMM yyyy")
    
    if 'Disable Sticky Keys' in answers['tasks']:
        os.system('reg add "HKEY_CURRENT_USER\\Control Panel\\Accessibility\\StickyKeys" /v "Flags" /t REG_SZ /d "506" /f')
    
    if 'Show desktop' in answers['tasks']:
        os.system('reg add "HKEY_CURRENT_USER\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "TaskbarSd" /t REG_DWORD /d "1" /f')
    
    if 'Enable Middle Button TouchPad' in answers['tasks']:
        os.system('reg add "HKEY_CURRENT_USER\\Software\\Elantech\\SmartPad" /v "n3A2PST_Middle_Mode" /t REG_DWORD /d "1" /f')
        print("RESTART REQUIRED!")
    
    if 'Android SDK Platform Tools' in answers['tasks']:
        os.system('winget install Google.PlatformTools --scope machine')
        print("ADB installed. Follow the README section 'Android Debloat Helper' to pair/connect your device.")
    
    if 'Microsoft Visual Studio Code' in answers['tasks']:
        os.system('winget install Microsoft.VisualStudioCode --scope machine')
    
    if 'Microsoft 365' in answers['tasks']:
        print("Microsoft 365 requires a valid license. This script only automates installation.")
        office_apps = ['Word', 'Excel', 'Access', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'PowerPoint', 'Publisher', 'Teams', 'Bing']
        selected_apps = inquirer.prompt([inquirer.Checkbox('office_apps', message="Seleziona i programmi Microsoft 365 da installare:", choices=office_apps)])['office_apps']
        exclude_lines = '\n'.join([f'      <ExcludeApp ID="{app}" />' for app in office_apps if app not in selected_apps])
        config_template = f"""<Configuration>
    <Add OfficeClientEdition="64" Channel="Current">
        <Product ID="O365ProPlusRetail">
        <Language ID="it-it" />
        <Language ID="en-us" />
    {exclude_lines}
        </Product>
    </Add>
    <Updates Enabled="TRUE" />
    <RemoveMSI />
    <Display Level="Full" AcceptEULA="TRUE" />
    </Configuration>""".strip()
        config_path = os.path.join(os.path.expanduser("~"), "Desktop", "Configuration.xml")
        with open(config_path, 'w') as file:
            file.write(config_template)
        os.system(f'winget install Microsoft.Office --override "/configure {config_path}"')
        os.remove(config_path)
    
    if 'PC On/Off Time' in answers['tasks']:
        os.system('powershell -Command "Invoke-WebRequest -Uri \'https://www.neuber.com/free/pctime/pctime.zip\' -OutFile \"$env:userprofile\\Desktop\\pctime.zip\""')
        os.system('powershell -Command "Expand-Archive -Path \"$env:userprofile\\Desktop\\pctime.zip\" -DestinationPath \"$env:userprofile\\Desktop\""')
        os.system('powershell -Command "Remove-Item \"$env:userprofile\\Desktop\\pctime.zip\", \"$env:userprofile\\Desktop\\PcOnOffTime.chm\""')
    
    if 'RedNotebook' in answers['tasks']:
        print("RedNotebook diary will be stored in OneDrive. Installing OneDrive...")
        os.system('winget install "Microsoft.OneDrive"')
        os.system('powershell -NoProfile -Command "if ($env:OneDrive) { Start-Process explorer.exe $env:OneDrive } else { Start-Process explorer.exe shell:OneDrive }"')
        os.system('powershell -NoProfile -Command "$t=Get-Date; Write-Host \'>>> Waiting for OneDrive to finish syncing and for the .rednotebook folder to exist...\'; while (!(Test-Path \\"$env:USERPROFILE\\OneDrive\\.rednotebook\\") -and ((Get-Date)-$t).TotalMinutes -lt 15) { Start-Sleep -Seconds 3 }; if (!(Test-Path \\"$env:USERPROFILE\\OneDrive\\.rednotebook\\")) { Write-Host \'OneDrive sync timed out. Run again after sync.\'; exit 1 }"')
        print("Installing RedNotebook...")
        os.system('winget install "RedNotebook"')
        print("Linking diary to OneDrive...")
        rc = os.system('powershell -NoProfile -Command "if (Test-Path \\"$env:USERPROFILE\\.rednotebook\\") { Write-Host \\"Symlink already exists. Skipping...\\" } else { New-Item -Path \\"$env:USERPROFILE\\.rednotebook\\" -ItemType SymbolicLink -Target \\"$env:USERPROFILE\\OneDrive\\.rednotebook\\" | Out-Null 2>$null }"')
        print("RedNotebook installed and diary linked to OneDrive successfully!" if rc == 0 else "Symlink failed — ensure OneDrive is signed in and the .rednotebook folder exists, then run again.")
    
    if 'Scanner Windows' in answers['tasks']:
        os.system('winget install --id 9WZDNCRFJ3PV --source msstore --accept-source-agreements --accept-package-agreements')

    if 'Snipping Tool' in answers['tasks']:
        os.system('winget install --id 9MZ95KL8MR0L --source msstore --silent --accept-source-agreements --accept-package-agreements')
            
    if 'Steam' in answers['tasks']:
        os.system('winget install Valve.Steam --scope machine')
    
    if 'Telegram Desktop' in answers['tasks']:
        os.system('winget install Telegram.TelegramDesktop --scope machine')
    
    if 'VLC Media Player' in answers['tasks']:
        os.system('winget install VideoLAN.VLC --scope machine')
    
    if 'WhatsApp' in answers['tasks']:
        os.system('winget install --id 9NKSQGP7F2NH --source msstore --silent --accept-source-agreements --accept-package-agreements')
    
    if 'WinRAR' in answers['tasks']:
        os.system('winget install RARLab.WinRAR --scope machine')
        print(r"If you own a WinRAR license, place your personal 'rarreg.key' in C:\Program Files\WinRAR (not included)")

except KeyboardInterrupt: print("\nCancelled by user."); sys.exit(130)
