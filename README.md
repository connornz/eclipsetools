# Eclipse Swiss Army Knife - Documentation

## Overview

The Eclipse Swiss Army Knife is a PowerShell automation tool for Eclipse installation, configuration, and maintenance tasks. It streamlines setup for Eclipse DMS, IIS, SQL Server, and Eclipse services.

**Version:** 2.1.1  
**Created by:** Connor Brown for Eclipse Support Team  
**Requirements:** Windows 10/11 or Server 2016+, PowerShell 5.1+, Administrator privileges

---

## Quick Start

Download the latest version automatically by running the launcher script from GitHub:

1. Download `Launch-EclipseSwissArmyKnife.ps1` and `Run Me.bat` from https://github.com/connornz/eclipsetools

2. Run **Run Me.bat** → Automatically checks for updates and launches the latest version

3. Choose your mode:
   - **Interactive Mode:** Select individual options [0-20] for manual, step-by-step tasks
   - **Unattended Mode:** Press [U] to configure settings and select multiple tasks for automated execution

4. Follow on-screen prompts

---

## What's New in v2.1.1

### 🔄 Auto-Update System
- **GitHub Integration:** Script automatically checks for and downloads the latest version from GitHub
- **Version Tracking:** Smart version comparison ensures you always have the latest features
- **Offline Fallback:** Works offline using cached local version
- **Seamless Updates:** Updates happen automatically when you run the launcher

### 🔧 Enhanced Service Management (Options 9 & 10)
- **Automatic NSSM Installation:** Automatically downloads and installs NSSM if not present
- **Improved UI:** Better formatting with organized sections for configuration
- **Secure Password Input:** Passwords are now hidden when typing
- **Service Verification:** ESB Service checks for esbListener.exe to verify successful database connection
- **Automatic Startup:** Services start immediately after installation
- **Automatic (Delayed) Startup:** Services configured to start with system (delayed to ensure dependencies are ready)

### ✨ Previous Updates (v2.1.0)
- **Unattended Server Setup Mode [U]:** Configure once, run many tasks automatically
- **Silent Installations:** Most tasks now install silently without user interaction
- **Enhanced Task Functionality:** Improved installation processes across all options

---

For full documentation, see the detailed README in the repository.
