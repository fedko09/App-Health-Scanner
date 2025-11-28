App Health & Crash Diagnostic
=============================

Overview
--------
This tool is a Windows 10/11 diagnostic utility designed to analyze
application failures, crashes, hangs, and system-level issues. It provides
a graphical interface built with PowerShell + WPF and includes multiple
levels of detail for basic users, power users, and system administrators.

The tool gathers critical information from:
- Windows Event Logs (Application + System)
- Windows Reliability Monitor (Win32_ReliabilityRecords)
- System hardware and OS details (CPU, RAM, disks, uptime, network)
- Application-specific crash and hang entries
- Provider-based error sources and common WER crash data

Key Features
------------
• **Quick Scan**  
  Collects system-wide Critical and Error events within the selected  
  time window.

• **App-Focused Scan**  
  Filters events and crashes to a specific executable or app name.  
  Includes a built-in process picker to simplify selecting the correct app.

• **Full System Snapshot**  
  Complete diagnostic capture including system info, hardware summary,  
  network data, event logs, and reliability history.

• **Process Picker Window**  
  Lets the user browse and filter running processes (All / User / System)  
  and automatically populate the App Filter field.

• **Dynamic UI Loading Overlay**  
  A “Working…” overlay appears during heavy scans so the UI provides  
  feedback even when PowerShell is under load.

• **Tabbed Interface**  
  - Events  
  - Crashes  
  - System Info  
  - Summary  
  - Suggestions  

• **Export Diagnostic Bundle (ZIP)**  
  Creates a diagnostic package including:
  - Events.csv
  - Crashes.csv
  - SystemInfo.txt
  - Summary.txt
  - Suggestions.txt

• **Smart Suggestions Engine**  
  Provides guidance based on detected issues such as:
  - Application Error (Event ID 1000)
  - Application Hang (Event ID 1002)
  - Low disk conditions
  - WER crash entries
  - Profile-based advanced recommendations

• **Context Menus**  
  Right-click on Events or Crashes to copy selected entries to clipboard.

Requirements
------------
- Windows 10 or Windows 11
- PowerShell 5.x or PowerShell 7.x (Windows PowerShell recommended)
- Execution policy allowing script execution (e.g., RemoteSigned)
- Administrator privileges recommended for full log access

Notes
-----
- The tool runs all scan operations on the primary UI thread.  
  The loading overlay is forced to repaint before the heavy operations begin.
- Some extremely large event logs or corrupted entries may not parse,  
  but the script safely handles most exceptions automatically.

Author
------
Bogdan Fedko
