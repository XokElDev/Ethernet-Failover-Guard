<#
.SYNOPSIS
    Advanced Ethernet Status Monitor (V6.8) - Complete Edition
.DESCRIPTION
    Robust monitoring of wired connection with auto-switch to Wi-Fi.
    Features: Auto-Recovery, Enhanced Audio Alerts, Tray Icon Control.

.NOTES
    ===========================================================================
    [TASK SCHEDULER SETUP GUIDE] - HOW TO RUN AUTOMATICALLY ON STARTUP
    ===========================================================================
    To make this script run silently in the background every time you turn on 
    your PC (without asking for admin permission), follow these steps:

    STEP 1: Open Task Scheduler
      - Press the 'Windows' key, type "Task Scheduler" (or "Programador de tareas"), and open it.
      - In the right-hand panel, click "Create Task..." (NOT "Create Basic Task").

    STEP 2: "General" Tab
      - Name: Give it a name like "Network Monitor Script".
      - Security Options:
        * Check "Run with highest privileges" (CRITICAL: This allows admin access).
        * Select "Run only when user is logged on" (Required for the Tray Icon to appear).

    STEP 3: "Triggers" Tab (The Delay)
      - Click "New...".
      - Begin the task: Select "At log on".
      - Delay Settings:
        * Check "Delay task for:" and select "30 seconds" or "1 minute".
        * (This ensures Windows finishes loading drivers before the script starts).
      - Click OK.

    STEP 4: "Actions" Tab (Launch the Script)
      - Click "New...".
      - Action: "Start a program".
      - Program/script: Type "powershell.exe"
      - Add arguments (optional): Copy and paste the line below (UPDATE THE PATH):
        -ExecutionPolicy Bypass -File "C:\Path\To\Your\monitor_red.ps1"
        
        (Tip: Shift + Right-click your script file -> "Copy as path").

    STEP 5: "Conditions" Tab (Laptop Users)
      - Uncheck "Start the task only if the computer is on AC power".
      - (This ensures it runs even if you are on battery).

    STEP 6: Finish
      - Click OK. Restart your computer to test it.
    ===========================================================================
#>

# ==============================================================================
# [0] AUTO-ADMIN ELEVATION (Fallback)
# ==============================================================================
# Even with Task Scheduler, this safety block ensures the script always has 
# the necessary permissions to disable/enable adapters.
$CheckAdmin = $true   
if($CheckAdmin){
    if (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "Requesting Administrator privileges..." -ForegroundColor Yellow
        $CommandLine = "-NoProfile -ExecutionPolicy Bypass -File `"" + $MyInvocation.MyCommand.Path + "`""
        Start-Process -FilePath PowerShell.exe -Verb RunAs -ArgumentList $CommandLine
        Stop-Process -Id $PID
    }
}

# ==============================================================================
# [1] GENERAL CONFIGURATION
# ==============================================================================

# [TUTORIAL] Adapter Names
# Use 'Get-NetAdapter' in PowerShell to find your exact names.
$EthernetAdapterName = "Ethernet 2"
$WifiAdapterName = "Wi-Fi"

# [EXPLANATION] Speed Threshold (1000 = 1Gbps)
# If speed drops below this (e.g. to 100 Mbps), it triggers the alert.
$MinimumSpeed = 650 

# [EXPLANATION] Startup Wait
# Internal delay to allow network negotiation if run manually.
$StartupWaitSeconds = 30

# ==============================================================================
# [1.5] AUDIO CONFIGURATION
# ==============================================================================

# [EXPLANATION] Alert Volume Level (0 to 100)
# Forces system volume to this level specifically for the alert sound.
$VolumenAlerta = 100

# ==============================================================================
# [2] DETECTION AND ACTION CONFIGURATION
# ==============================================================================

# Check for internet access (Ping) in addition to physical connection?
$UsePingCheck = $true   

# Should we disable the cable to force Wi-Fi usage on failure?
$ForceWifiOnFail = $true

# ==============================================================================
# [3] RECOVERY CONFIGURATION
# ==============================================================================

# Try to re-enable the cable automatically to check if it's fixed?
$EnableRetry = $true 

# Minutes to wait before re-enabling the cable.
$RetryWaitMinutes = 0.1

# ==============================================================================
# [4] LOGGING CONFIGURATION
# ==============================================================================

$LogFile = Join-Path $PSScriptRoot "network_log.txt"
$StartHidden = $true

# ==============================================================================
# SYSTEM TRAY & WINDOWS API SETUP
# ==============================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import User32 methods for Window Control
$user32_def = @"
    [DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(int hWnd);
    [DllImport("user32.dll")] public static extern bool IsIconic(int hWnd); 
"@
Add-Type -MemberDefinition $user32_def -Name Api -Namespace User32

# --- AUDIO VOLUME CONTROL C# DEFINITION ---
$AudioCode = @"
    using System;
    using System.Runtime.InteropServices;

    namespace AudioTools {
        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IAudioEndpointVolume {
            int RegisterControlChangeNotify(IntPtr pNotify);
            int UnregisterControlChangeNotify(IntPtr pNotify);
            int GetChannelCount(out int pnChannelCount);
            int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
            int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
            int GetMasterVolumeLevel(out float pfLevelDB);
            int GetMasterVolumeLevelScalar(out float pfLevel);
            int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext);
            int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext);
            int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
            int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
            int SetMute([MarshalAs(UnmanagedType.Bool)] bool bMute, Guid pguidEventContext);
            int GetMute(out bool pbMute);
            int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
            int VolumeStepUp(Guid pguidEventContext);
            int VolumeStepDown(Guid pguidEventContext);
            int QueryHardwareSupport(out uint pdwHardwareSupportMask);
            int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IMMDevice {
            int Activate(ref Guid id, int clsCtx, int activationParams, out object interfacePtr);
        }

        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IMMDeviceEnumerator {
            int EnumAudioEndpoints(int dataFlow, int dwState, out object ppDevices);
            int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppEndpoint);
        }

        [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumeratorComObject { }

        public class Volume {
            private static IAudioEndpointVolume GetMasterVolumeObject() {
                var enumerator = new MMDeviceEnumeratorComObject() as IMMDeviceEnumerator;
                IMMDevice speakers;
                const int eRender = 0;
                const int eMultimedia = 1;
                enumerator.GetDefaultAudioEndpoint(eRender, eMultimedia, out speakers);
                
                Guid IID_IAudioEndpointVolume = typeof(IAudioEndpointVolume).GUID;
                object o;
                speakers.Activate(ref IID_IAudioEndpointVolume, 0, 0, out o);
                
                return o as IAudioEndpointVolume;
            }

            public static float GetVolume() {
                float currentVolume;
                GetMasterVolumeObject().GetMasterVolumeLevelScalar(out currentVolume);
                return currentVolume * 100;
            }

            public static void SetVolume(float level) {
                if (level < 0) level = 0;
                if (level > 100) level = 100;
                GetMasterVolumeObject().SetMasterVolumeLevelScalar(level / 100, Guid.Empty);
            }

            public static void Unmute() {
                GetMasterVolumeObject().SetMute(false, Guid.Empty);
            }
        }
    }
"@
Add-Type -TypeDefinition $AudioCode -Language CSharp

$consolePtr = [System.Diagnostics.Process]::GetCurrentProcess().MainWindowHandle

# Constants
$SW_HIDE = 0
$SW_SHOW = 5
$SW_RESTORE = 9

# --- SETUP TRAY ICON ---
$TrayIcon = New-Object System.Windows.Forms.NotifyIcon
$TrayIcon.Icon = [System.Drawing.SystemIcons]::Information 
$TrayIcon.Text = "Network Monitor Active"
$TrayIcon.Visible = $true

# --- SETUP TRAY CONTEXT MENU ---
$ContextMenu = New-Object System.Windows.Forms.ContextMenu

$MenuItemTestFail = New-Object System.Windows.Forms.MenuItem "[Sound] Test Failure"
$MenuItemTestFail.add_Click({
    Write-Host "[TEST] Testing Failure Sound..." -ForegroundColor Yellow
    Play-Sound -Type "Failure"
})

$MenuItemTestSuccess = New-Object System.Windows.Forms.MenuItem "[Sound] Test Success"
$MenuItemTestSuccess.add_Click({
    Write-Host "[TEST] Testing Success Sound..." -ForegroundColor Green
    Play-Sound -Type "Success"
})

$MenuItemShow = New-Object System.Windows.Forms.MenuItem "Show/Hide Console"
$MenuItemShow.add_Click({
    if ([User32.Api]::IsWindowVisible($consolePtr)) {
        [User32.Api]::ShowWindow($consolePtr, $SW_HIDE) | Out-Null
    } else {
        [User32.Api]::ShowWindow($consolePtr, $SW_RESTORE) | Out-Null
        [User32.Api]::ShowWindow($consolePtr, $SW_SHOW) | Out-Null
    }
})

$MenuItemExit = New-Object System.Windows.Forms.MenuItem "Exit Monitor"
$MenuItemExit.add_Click({
    $TrayIcon.Visible = $false
    $TrayIcon.Dispose()
    Stop-Process -Id $PID
})

# Add items to menu
$ContextMenu.MenuItems.Add($MenuItemTestFail)
$ContextMenu.MenuItems.Add($MenuItemTestSuccess)
$ContextMenu.MenuItems.Add("-") 
$ContextMenu.MenuItems.Add($MenuItemShow)
$ContextMenu.MenuItems.Add($MenuItemExit)
$TrayIcon.ContextMenu = $ContextMenu

if ($StartHidden) {
    [User32.Api]::ShowWindow($consolePtr, $SW_HIDE) | Out-Null
}

# ==============================================================================
# CORE FUNCTIONS
# ==============================================================================

function Write-Log {
    param ([string]$Message)
    $Line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Add-Content -Path $LogFile -Value $Line -Encoding UTF8
    Write-Host $Line
}

function Play-Sound {
    param ([string]$Type)
    
    # --- AUDIO VOLUME OVERRIDE LOGIC ---
    try {
        # 1. Get current volume to restore it later
        $OldVolume = [AudioTools.Volume]::GetVolume()
        
        # 2. Force System Unmute and Set Target Volume
        [AudioTools.Volume]::Unmute()
        [AudioTools.Volume]::SetVolume($VolumenAlerta)
        
        # 3. Play Sound Patterns
        if ($Type -eq "Failure") {
            # PATTERN: Neutral -> DROP to Low (Grave)
            # 1000Hz (Neutral) -> 400Hz (Deep/Grave)
            [System.Console]::Beep(1000, 200)
            Start-Sleep -Milliseconds 100
            [System.Console]::Beep(400, 500)
        }
        elseif ($Type -eq "Success") {
            # PATTERN: Neutral -> RISE to High (Acute)
            # 1000Hz (Neutral) -> 2500Hz (Sharp/Acute)
            [System.Console]::Beep(1000, 200)
            Start-Sleep -Milliseconds 100 
            [System.Console]::Beep(2500, 400)
        }
        elseif ($Type -eq "Startup") {
            # PATTERN: Simple Short Neutral Chirp
            [System.Console]::Beep(1000, 150)
            Start-Sleep -Milliseconds 50
            [System.Console]::Beep(1000, 150)
        }
        
        # 4. Wait a bit longer (1s) to ensure sound finishes before lowering volume
        Start-Sleep -Milliseconds 1000
        [AudioTools.Volume]::SetVolume($OldVolume)
        
    } catch {
        # Fallback if advanced audio fails
        if ($Type -eq "Failure") { 
            [System.Console]::Beep(1000, 200)
            [System.Console]::Beep(400, 500) 
        }
        if ($Type -eq "Success") { 
            [System.Console]::Beep(1000, 200)
            [System.Console]::Beep(2500, 400) 
        }
        if ($Type -eq "Startup") {
             [System.Console]::Beep(1000, 300) 
        }
    }
}

function Show-Notification {
    param ([string]$Title, [string]$Message, [string]$IconType = "Warning")
    
    if ($IconType -eq "Error") { $TrayIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Error }
    elseif ($IconType -eq "Warning") { $TrayIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Warning }
    else { $TrayIcon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info }

    $TrayIcon.BalloonTipTitle = $Title
    $TrayIcon.BalloonTipText = $Message
    $TrayIcon.ShowBalloonTip(5000) 
}

# ==============================================================================
# MAIN LOGIC LOOP
# ==============================================================================

Write-Log "--- NETWORK MONITOR V6.8 (COMPLETE) STARTED ---"

Start-Sleep -Seconds $StartupWaitSeconds

# [NEW] Play sound to indicate script is ready and volume control works
Play-Sound -Type "Startup"

# Guard to ensure adapters are enabled
if ($ForceWifiOnFail) {
    Get-NetAdapter -Name $WifiAdapterName -ErrorAction SilentlyContinue | Enable-NetAdapter -Confirm:$false -ErrorAction SilentlyContinue
}
Get-NetAdapter -Name $EthernetAdapterName -ErrorAction SilentlyContinue | Enable-NetAdapter -Confirm:$false -ErrorAction SilentlyContinue

$PreviousState = "Startup"
Write-Log "Monitoring Active. Right-click Tray Icon to Test Sounds."

while ($true) {
    # [IMPORTANT] DoEvents handles Tray Menu clicks
    [System.Windows.Forms.Application]::DoEvents()

    # Tray Minimization Logic
    if ([User32.Api]::IsIconic($consolePtr)) {
        [User32.Api]::ShowWindow($consolePtr, $SW_HIDE) | Out-Null
    }

    try {
        $Adapter = Get-NetAdapter -Name $EthernetAdapterName -ErrorAction SilentlyContinue

        if ($null -eq $Adapter -or $Adapter.Status -eq "Disabled") {
            if ($PreviousState -ne "DisabledByScript" -and $PreviousState -ne "DisabledExternally") {
                Write-Log "Ethernet adapter disabled or missing."
                $PreviousState = "DisabledExternally"
            }
            Start-Sleep -Milliseconds 500
            continue
        }

        $CurrentStatus = $Adapter.Status 
        $LinkSpeed = $Adapter.LinkSpeed   
        
        $NumericSpeed = 0
        if ($LinkSpeed -match "Gbps") {
            $NumericSpeed = [double]($LinkSpeed -replace "[^0-9.]","") * 1000
        } elseif ($LinkSpeed -match "Mbps") {
            $NumericSpeed = [double]($LinkSpeed -replace "[^0-9.]","")
        }

        $PingOK = $true
        if ($UsePingCheck -and $CurrentStatus -eq "Up") {
            $PingOK = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet
        }

        # --- FAILURE DETECTION LOGIC ---
        $FailureDetected = $false
        $FailureReason = ""

        if ($CurrentStatus -ne "Up") {
            if ($PreviousState -eq "Up" -or $PreviousState -eq "Startup") {
                $FailureDetected = $true
                $FailureReason = "Cable physically disconnected"
            }
        }
        elseif ($NumericSpeed -gt 0 -and $NumericSpeed -lt $MinimumSpeed) {
             $FailureDetected = $true
             $FailureReason = "Link speed degraded to $LinkSpeed"
        }
        elseif ($UsePingCheck -and -not $PingOK) {
            Start-Sleep -Milliseconds 500
            if (-not (Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet)) {
                $FailureDetected = $true
                $FailureReason = "No Internet access (Ping failed)"
            }
        }

        # --- ACTION ON FAILURE ---
        if ($FailureDetected) {
            Play-Sound -Type "Failure"
            Write-Log "ALERT: $FailureReason"
            
            if ($ForceWifiOnFail) {
                Show-Notification -Title "Ethernet Failure" -Message "$FailureReason. Switching to Wi-Fi." -IconType "Error"
                
                Enable-NetAdapter -Name $WifiAdapterName -Confirm:$false -ErrorAction SilentlyContinue
                Disable-NetAdapter -Name $EthernetAdapterName -Confirm:$false
                
                $PreviousState = "DisabledByScript"

                if ($EnableRetry) {
                    Write-Log "Wait mode: $RetryWaitMinutes min."
                    $EndTime = (Get-Date).AddMinutes($RetryWaitMinutes)
                    
                    while ((Get-Date) -lt $EndTime) {
                        [System.Windows.Forms.Application]::DoEvents()
                        if ([User32.Api]::IsIconic($consolePtr)) { [User32.Api]::ShowWindow($consolePtr, $SW_HIDE) | Out-Null }
                        Start-Sleep -Milliseconds 500
                    }
                    
                    Write-Log "Retrying Ethernet..."
                    Enable-NetAdapter -Name $EthernetAdapterName -Confirm:$false
                    Start-Sleep -Seconds 15 
                    continue 
                }
            } else {
                 Show-Notification -Title "Network Alert" -Message "$FailureReason" -IconType "Warning"
                 $PreviousState = "NotifiedFailure" 
            }
        }

        # --- RECOVERY SUCCESS ---
        if ($CurrentStatus -eq "Up" -and $NumericSpeed -ge $MinimumSpeed -and ($PreviousState -eq "DisabledByScript" -or $PreviousState -eq "NotifiedFailure")) {
             Play-Sound -Type "Success"
             Write-Log "RECOVERED: Stable at $LinkSpeed"
             Show-Notification -Title "Network Stable" -Message "Ethernet connection recovered." -IconType "Info"
             $PreviousState = "Up"
        }
        
        if (-not $FailureDetected -and $PreviousState -ne "DisabledByScript") {
            $PreviousState = $CurrentStatus
        }

    } catch {
        Write-Log "Error: $_"
    }

    # Wait loop kept simple for responsive Tray Icon
    for ($i=0; $i -lt 4; $i++) {
        [System.Windows.Forms.Application]::DoEvents()
        if ([User32.Api]::IsIconic($consolePtr)) { [User32.Api]::ShowWindow($consolePtr, $SW_HIDE) | Out-Null }
        Start-Sleep -Milliseconds 500
    }
}