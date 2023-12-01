<#
.SYNOPSIS
RemoveHP is a PowerShell script designed to remove pre-installed HP software, commonly referred to as bloatware, from a Windows system.

.DESCRIPTION
This script performs the following actions:
- Sets the execution policy to unrestricted for the current user to allow scripts to run.
- Defines a list of HP AppX packages and traditional programs that are considered bloatware.
- Stops and disables HP-related services to prepare for bloatware removal.
- Attempts to uninstall each piece of bloatware using various methods:
  - Uninstalls AppX packages from all users.
  - Uninstalls provisioned AppX packages (which would affect new user accounts).
  - Uninstalls traditional HP programs using the Uninstall-Package cmdlet and msiexec where necessary.
- Provides verbose output to the console with a pirate-themed narrative.
- Performs a post-removal check to confirm the uninstallation of the HP software.
- Prompts the user to restart the computer to complete the bloatware removal process.

.EXAMPLE
To execute the script, simply run the RemoveHP function after defining it in your PowerShell session.
RemoveHP

.NOTES
Author: John Bignold
Additional Credits: Adapted from migration scripts by https://gist.github.com/mark05e
#>




Function RemoveHP {
    Write-Host -BackgroundColor DarkYellow "Ahoy Matey! Welcome to Captain John's Mighty Bloatware Plunderin' Script" -ForegroundColor Black    

    # Let the cannons loose (Allow the scripts to run)
    Set-ExecutionPolicy Unrestricted -Scope CurrentUser

    # List of cursed treasures (apps) to remove
    $PlunderList = @(
        "AD2F1837.HPJumpStarts"
        "AD2F1837.HPPCHardwareDiagnosticsWindows"
        "AD2F1837.HPPowerManager"
        "AD2F1837.HPPrivacySettings"
        "AD2F1837.HPSupportAssistant"
        "AD2F1837.HPSureShieldAI"
        "AD2F1837.HPSystemInformation"
        "AD2F1837.HPQuickDrop"
        "AD2F1837.HPWorkWell"
        "AD2F1837.myHP"
        "AD2F1837.HPDesktopSupportUtilities"
        "AD2F1837.HPQuickTouch"
        "AD2F1837.HPEasyClean"
        "AD2F1837.HPSystemInformation"
    )

    # List of vile sea monsters (programs) to hunt down
    $RivalShips = @(
        "HP Device Access Manager"
        "HP Client Security Manager"
        "HP Connection Optimizer"
        "HP Documentation"
        "HP MAC Address Manager"
        "HP Notifications"
        "HP System Info HSA Service"
        "HP Security Update Service"
        "HP System Default Settings"
        "HP Sure Click"
        "HP Sure Click Security Browser"
        "HP Sure Run"
        "HP Sure Run Module"
        "HP Sure Recover"
        "HP Sure Sense"
        "HP Sure Sense Installer"
        "HP Wolf Security"
        "HP Wolf Security - Console"
        "HP Wolf Security Application Support for Sure Sense"
        "HP Wolf Security Application Support for Windows"
    )

    $HPMap = "AD2F1837"

    $SpottedTreasures = Get-AppxPackage -AllUsers `
    | Where-Object { ($PlunderList -contains $_.Name) -or ($_.Name -match "^$HPMap") }

    $HiddenTreasures = Get-AppxProvisionedPackage -Online `
    | Where-Object { ($PlunderList -contains $_.DisplayName) -or ($_.DisplayName -match "^$HPMap") }

    $RivalPirates = Get-Package | Where-Object { $RivalShips -contains $_.Name }

    # Anchor the HP ships (Stop HP Services)
    Function AnchorHPShips($ship) {
        if (Get-Service -Name $ship -ea SilentlyContinue) {
            Stop-Service -Name $ship -Force -Confirm:$False
            Set-Service -Name $ship -StartupType Disabled
        }
    }

    AnchorHPShips -name "HotKeyServiceUWP"
    AnchorHPShips -name "HPAppHelperCap"
    AnchorHPShips -name "HP Comm Recover"
    AnchorHPShips -name "HPDiagsCap"
    AnchorHPShips -name "HotKeyServiceUWP"
    AnchorHPShips -name "LanWlanWwanSwitchgingServiceUWP" # do we need to stop this?
    AnchorHPShips -name "HPNetworkCap"
    AnchorHPShips -name "HPSysInfoCap"
    AnchorHPShips -name "HP TechPulse Core"

    # Setting sail to challenge rival pirates and seize their loot!
    $RivalPirates | ForEach-Object {

        # Spotting a rival pirate ship on the horizon
        Write-Host -Object "Ahoy, mateys! A rival pirate ship by the name of [$($_.Name)] looms on the horizon! Hoist the Jolly Roger and prepare to engage!" -ForegroundColor Yellow

        Try {
            # Engage in naval warfare to claim their bounty
            $Null = $_ | Uninstall-Package -AllVersions -Force -ErrorAction Stop
            Write-Host -Object "Yarrr! We've bested the scurvy crew of [$($_.Name)] and claimed their booty!" -ForegroundColor Green
        }
        Catch {
            # They’re putting up a good fight; time to employ some cunning tactics!
            Write-Host -Object "By Davy Jones' locker! The crew of [$($_.Name)] be putting up a fierce resistance! Man the cannons!" -ForegroundColor Yellow
        
            Try {
                # Identify their ship's weak point
                $shipWeakPoint = Get-WmiObject win32_product | where { $_.name -like "$($_.Name)" }
                if ($_ -ne $null) {
                    # Fire a cannonball directly at their ship's weak point
                    msiexec /x $shipWeakPoint.IdentifyingNumber /quiet /noreboot
                    Write-Host -Object "Direct hit! We've sunk the ship of [$($_.Name)] and plundered their treasure!" -ForegroundColor Green
                }
                else { Write-Host -Message "Blast! A miss! The ship [$($_.Name)] managed to evade our cannons!" -ForegroundColor Red }
            }
            Catch { Write-Host -Message "Curses! We couldn’t sink the ship of [$($_.Name)]. They'll live to sail another day." -ForegroundColor Red }
        }
    }


    Write-Host -Object "On the horizon, a monstrous shadow emerges! Prepare to engage with HP Wolf Security" -ForegroundColor Yellow

    # Fallback plan 1 to banish the HP Wolf Security with our secret weapon (msiexec)
    Try {
        MsiExec /x "{0E2E04B0-9EDD-11EB-B38C-10604B96B11E}" /qn /norestart
        Write-Host -Object "Using the ol' MSI scroll to drive the HP Wolf Security down to the briny deep!" -ForegroundColor Yellow
        Write-Host -Object "Yarr! We've banished the fierce beast!" -ForegroundColor Green
    }
    Catch {
        Write-Warning -Object "Shiver me timbers! The cursed scroll led us astray. HP Wolf Security remains: $($_.Exception.Message)"
    }

    # Fallback plan 2, another attempt to drive the HP Wolf Security back to Davy Jones' locker
    Try {
        MsiExec /x "{4DA839F0-72CF-11EC-B247-3863BB3CB5A8}" /qn /norestart
        Write-Host -Object "By Davy Jones' locker! Raise the Jolly Roger and fire the cannons to repel HP Wolf Security!" -ForegroundColor Yellow
        Write-Host -Object "Yarr! We've banished the fierce beast!" -ForegroundColor Green
    }
    Catch {
        Write-Warning -Object "Blimey! Our efforts be thwarted again! The sea beast remains. Here be the wretched reason: $($_.Exception.Message)"
    }

    # Banish the appx treasures buried deep - AppxProvisionedPackage
    ForEach ($ProvPackage in $HiddenTreasures) {

        Write-Host -Object "Searching the shores to recover the buried booty: [$($ProvPackage.DisplayName)]..." -ForegroundColor Yellow

        Try {
            $Null = Remove-AppxProvisionedPackage -PackageName $ProvPackage.PackageName -Online -ErrorAction Stop
            Write-Host -Object "Huzzah! We've seized the treasure: [$($ProvPackage.DisplayName)] and it be ours!" -ForegroundColor Green
        }
        Catch { Write-Warning -Message "By Blackbeard's beard! We couldn't unearth the booty: [$($ProvPackage.DisplayName)]. It be cursed!" }
    }

    # Banishing appx ghost ships - AppxPackage
    ForEach ($AppxPackage in $SpottedTreasures) {
                                        
        Write-Host -Object "All hands on deck! Prepare to engage the ghostly vessel: [$($AppxPackage.Name)]..." -ForegroundColor Yellow

        Try {
            $Null = Remove-AppxPackage -Package $AppxPackage.PackageFullName -AllUsers -ErrorAction Stop
            Write-Host -Object "Yarrr! We've sent the phantom ship [$($AppxPackage.Name)] to the depths!" -ForegroundColor Green
        }
        Catch { Write-Warning -Message "Ahoy, the phantom vessel [$($AppxPackage.Name)] slipped through the fog! It's out of our grasp!" }
    }

    # Raise the Jolly Roger and begin the inspection
    Write-Host "Hoisting the colors and beginning the post-battle inspection!" -ForegroundColor Cyan

    # Search the waters for Appx Packages 
    Write-Host "Scouring the seas for Packages bearing the HP flag..." -ForegroundColor Cyan
    $FoundAppx = Get-AppxPackage -AllUsers | where { $_.Name -like "*HP*" }
    if ($FoundAppx) {
        $FoundAppx | Format-List
    }
    else {
        Write-Host "The seas be clear of any Packages waving the HP colors." -ForegroundColor Green
    }

    # Seek out any provisioned curses
    Write-Host "Searching for any mysterious curses hidden away..." -ForegroundColor Cyan
    $FoundProvisioned = Get-AppxProvisionedPackage -Online | where { $_.DisplayName -like "*HP*" }
    if ($FoundProvisioned) {
        $FoundProvisioned | Format-List
    }
    else {
        Write-Host "No hidden curses with the HP mark. The coast be clear!" -ForegroundColor Green
    }

    # Look out for any lingering HP brigands
    Write-Host "Keeping a sharp eye out for any HP brigands lurking about..." -ForegroundColor Cyan
    $FoundPackages = Get-Package | select Name, FastPackageReference, ProviderName, Summary | Where { $_.Name -like "*HP*" }
    if ($FoundPackages) {
        $FoundPackages | Format-List
    }
    else {
        Write-Host "The horizon be clear of any HP brigands! We've secured our digital domain, matey!" -ForegroundColor Green
    }


    # Decide to drop the anchor or set sail again (reboot or not)
    $command = Read-Host "Drop anchor and rest (restart) now, aye or nay? [y/n]"
    switch ($command) {
        y { Restart-computer -Force -Confirm:$false }
        n { exit }
        default { write-warning "Setting course without rest!" }
    }

}
