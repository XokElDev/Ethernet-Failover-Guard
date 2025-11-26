# Ethernet-Failover-Guard
A robust PowerShell script that monitors Ethernet connection stability, detects physical disconnects or speed drops (e.g., 1Gbps to 100Mbps), and automatically switches to Wi-Fi and back on ETH if fixed, with audio alerts and system tray integration.

--------------------------------------------------------------------------

🛡️ Ethernet Failover Guard

Ethernet Failover Guard is a robust PowerShell script designed for Windows users who experience instability with their Ethernet connection (e.g., faulty cables, loose ports, or micro-cuts).

It actively monitors the connection status and link speed. If a physical disconnect or a speed drop (e.g., from 1Gbps down to 100Mbps) is detected, it automatically switches the system to Wi-Fi to maintain connectivity, plays an audio alert, and attempts to auto-recover the wired connection later.

✨ Features

⚡ Speed Degradation Detection: Instantly detects if your Gigabit connection drops to 100Mbps (often a sign of a bad cable/port).

🔄 Auto-Failover: Automatically enables Wi-Fi and disables the faulty Ethernet adapter to prevent slow speeds or packet loss.

🎧 Enhanced Audio Alerts: * Low Tone: Failure detected.

High Tone: Connection recovered.

Master Volume Override: Temporarily maximizes alert volume to ensure you hear the notification.

🛠️ System Tray Integration: Runs silently in the background with a tray icon. Right-click to test sounds or view the console.

🤖 Auto-Recovery: Periodically attempts to re-enable the Ethernet connection to check if the hardware issue has resolved itself.

🛡️ Auto-Elevation: Automatically restarts itself with Administrator privileges if needed.

🚀 Installation & Usage

1. Configuration

Open monitor_red.ps1 with a text editor (like Notepad or VS Code) and update the adapter names at the top of the file to match your system:

$EthernetAdapterName = "Ethernet"  # Change to your adapter name

$WifiAdapterName = "Wi-Fi"         # Change to your Wi-Fi adapter name

$MinimumSpeed = 1000               # Threshold in Mbps (1000 = 1Gbps)


To find your adapter names, open PowerShell and run:

Get-NetAdapter


2. Running the Script

Simply right-click monitor_red.ps1 and select "Run with PowerShell".

The script will request Administrator permissions.

It will start minimized in the System Tray (look for the blue 'i' icon near the clock).

3. (Optional) Run Automatically on Startup

The script includes a built-in guide to set up the Windows Task Scheduler.
Check the comment block at the top of the script named [TASK SCHEDULER SETUP GUIDE] for step-by-step instructions to have it run silently every time you log in.

🎮 Controls

The script runs in the background, but you can interact with it via the System Tray icon:

Right-Click Icon:

Test Failure Sound: Simulates a failure alert.

Test Success Sound: Simulates a recovery alert.

Show/Hide Console: Toggles the visibility of the log window.

Exit: Stops the script completely.

📝 License

This project is open-source. Feel free to modify and improve it!
