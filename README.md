# Eclipse Swiss Army Knife - Documentation

## Overview

The Eclipse Swiss Army Knife is a PowerShell automation tool for Eclipse installation, configuration, and maintenance tasks. It streamlines setup for Eclipse DMS, IIS, SQL Server, and Eclipse services.

**Version:** 2.1.0  
**Created by:** Connor Brown for Eclipse Support Team  
**Requirements:** Windows 10/11 or Server 2016+, PowerShell 5.1+, Administrator privileges

---

## Quick Start

1. Copy The Eclipse Swiss Army Knife folder to the client machine/server from `\\ubsnas1\Company\Support\Support Tools` or download `Eclipse Swiss Army Knife.zip`

2. Run **Run Me.bat** → Launches the PowerShell Script (as Administrator)

3. Choose your mode:
   - **Interactive Mode:** Select individual options [0-20] for manual, step-by-step tasks
   - **Unattended Mode:** Press [U] to configure settings and select multiple tasks for automated execution

4. Follow on-screen prompts

---

## What's New in v2.1.0

### ✨ Unattended Server Setup Mode [U]
- **Configure once, run many:** Set all configuration upfront (paths, SQL drive, client ID, domain, port)
- **Multi-task selection:** Interactive checkbox menu to select multiple tasks at once
- **Fully automated:** Tasks execute sequentially without prompts
- **Installation summary:** Shows completed/failed tasks with option to launch Eclipse Schema Update
- **Post-install checklist:** Displays outstanding configuration tasks for Eclipse Aura setups

### 🔧 Enhanced Task Functionality
Many tasks now perform **silent installations** instead of just downloading:
- Task 4: Silently installs SSMS 21 with `--quiet --norestart`
- Task 13: Installs .NET 8 runtimes via winget with fallback
- Task 15: Silently installs Eclipse Update Service and automatically configures client ID in database
- Task 17: Silently installs Eclipse Online Server to `C:\inetpub\Eclipse`
- Task 18: Creates IIS site, app pools, applications, self-signed certificate, and HTTPS binding
- Task 19: Silently installs Eclipse Smart Hub with default settings
- Task 20: Downloads, extracts, and launches Win-ACMEv2 (waits for user configuration)

---

## Menu Options Summary

### Core Installation (1-8)

**[1]** Create Eclipse Install directory structure (C:\Eclipse Install by default)

**[2]** Share Eclipse Install directory with proper permissions

**[3]** Download & Install SQL Server Express 2025 (~700MB, unattended installation with SA password: UBSubs123)

**[4]** Download & Install SQL Server Management Studio 21 (silent installation)

**[5]** Configure SQL Server network protocols (TCP/IP, Named Pipes) and firewall

**[6]** Install/Update Eclipse DMS 2032.25.290 (silent installation to default location)

**[7]** Export database registry keys to .reg file for deployment

**[8]** Add Ultimate domains to Internet Explorer Trusted Sites

### Eclipse Feed Services (9-10)

**[9]** ESB Service Management - Install/uninstall Eclipse ESB service (requires NSSM)

**[10]** Stock Export Service Management - Install/uninstall stock exporter service (requires NSSM)

⚠️ **WARNING:** Options 9 and 10 require manual database connection setup first. Run the ESB/Stock Export executables manually to configure database connections before installing them as services.

### Eclipse Online Services (11-20)

**[11]** Install IIS and all required Windows features

**[12]** Download & Install IIS URL Rewrite Module (requires IIS first)

**[13]** Install ASP.NET Core Hosting Bundle 8.0.22 & .NET Desktop Runtime 8.0.22 (via winget)

**[14]** Download & Install .NET Framework 4.8 (restart may be required)

**[15]** Download & Install Eclipse Update Service 1.52 (silent installation with client ID configuration)

**[16]** Download Eclipse Online Chrome 3.8.21

**[17]** Install Eclipse Online Server 12.0.89.0 (silent installation to C:\inetpub\Eclipse)

**[18]** Create IIS Application Pools & Site Binding
- Creates 'Eclipse' and 'EclipseAuraCore' application pools
- Creates IIS site 'Eclipse' with custom domain and port
- Generates self-signed SSL certificate 'EclipsePlaceholderSelfsigned'
- Creates HTTPS binding with certificate
- Creates 4 IIS applications: /AuraApi, /eclipseApi, /oAuth, /oAuth2
- Configures Windows Firewall rule

**[19]** Install Eclipse Smart Hub 12.0.101.0 (silent installation to default location)

**[20]** Install Win-ACMEv2 (downloads, extracts to Dependencies\Win-ACMEv2, and launches for SSL certificate configuration)

---

## Unattended Mode [U] - Detailed Guide

### What is Unattended Mode?

Unattended Mode allows you to configure all settings upfront, select multiple tasks via an interactive checkbox menu, and execute them sequentially without further prompts. Perfect for server deployments and standardized installations.

### How to Use Unattended Mode

1. **Launch the script** and press **[U]** from the main menu

2. **Configuration Wizard** - You'll be prompted for:
   - **Eclipse Install Folder Path** (default: `C:\Eclipse Install`)
   - **SQL Server Drive Letter** (e.g., C:, D:, E:)
   - **Eclipse Client ID** (4-digit number for Update Service registration)
   - **Eclipse Aura Hub Subdomain** (e.g., 'client' for client.eclipseaurahub.com.au)
   - **External Port Number** (e.g., 443 for HTTPS)

3. **Task Selection Menu**
   - Use **ARROW KEYS** to navigate
   - Press **SPACEBAR** to toggle tasks on/off
   - Press **A** to toggle all tasks
   - Press **ENTER** to start execution
   - Press **ESC** to cancel

4. **Automated Execution**
   - Tasks execute sequentially
   - Progress shown for each task
   - 3-second countdown between tasks
   - Log written to `C:\EclipseTechInstall.log`

5. **Installation Summary**
   - Shows which tasks succeeded/failed
   - Displays execution times
   - For Eclipse Aura setups (tasks 11-20), shows outstanding configuration tasks:
     - Map database in Eclipse Schema Update (press S to launch)
     - Activate Eclipse Aura using AWS Domain Manager
     - Whitelist domain on ws.dev.ultimate.net.au

### Available Tasks in Unattended Mode

✅ Task 1: Create Eclipse Install Directory Structure  
✅ Task 2: Share Eclipse Install Directory  
✅ Task 3: Download & Install SQL Server Express 2025  
✅ Task 4: Download & Install SQL Server Management Studio 21  
✅ Task 5: Configure SQL Server Network Protocols  
✅ Task 6: Install/Update Eclipse DMS 2032.25.290  
✅ Task 8: Add Ultimate Domains to Trusted Sites  
✅ Task 11: Install Microsoft IIS Server + Dependencies  
✅ Task 12: Download & Install IIS URL Rewrite Module  
✅ Task 13: Install ASP.NET Core Hosting Bundle 8.0.22 & .NET Desktop Runtime 8.0.22  
✅ Task 14: Download & Install .NET Framework 4.8  
✅ Task 15: Download & Install Eclipse Update Service 1.52  
✅ Task 17: Install Eclipse Online Server 12.0.89.0  
✅ Task 18: Create IIS Application Pools & Site Binding  
✅ Task 19: Install Eclipse Smart Hub 12.0.101.0  
✅ Task 20: Install Win-ACMEv2  

❌ Task 9 & 10: Not available in unattended mode (require manual database configuration)

### Example: Full Eclipse Aura Server Setup

For a complete Eclipse Aura Hub server setup, select these tasks in unattended mode:

1. **Infrastructure:** Tasks 1, 2, 3, 4, 5, 6, 8
2. **IIS & .NET:** Tasks 11, 12, 13, 14
3. **Eclipse Services:** Tasks 15, 17, 18, 19, 20

**Configuration Example:**
- Eclipse Install Path: `C:\Eclipse Install`
- SQL Server Drive: `D:`
- Eclipse Client ID: `5801`
- Subdomain: `client`
- Port: `443`

After execution, the script will:
- ✅ Install all components silently
- ✅ Configure IIS with HTTPS binding
- ✅ Set client ID in Eclipse Update Service database
- ✅ Launch Win-ACMEv2 for SSL certificate setup
- ✅ Show post-installation checklist

---

## Important Notes

### SQL Server Express 2025
- **Default Instance:** MSSQLSERVER
- **SA Password:** UBSubs123
- **Mixed Mode:** Enabled
- **Installation Time:** 10-15 minutes
- **Window Visible:** Progress window will show during installation

### SQL Server Management Studio 21
- **Silent Installation:** Uses `--quiet --norestart` arguments
- **Version:** 21.6.17 (latest SSMS 21.x)
- **Installation Time:** 5-10 minutes
- **Background Process:** Spawns Visual Studio Installer and setup processes
- **Verification:** Checks registry and file paths (SSMS may install to different location on Windows 11)

### Eclipse Update Service 1.52
- **Client ID Configuration:** Automatically sets ConsumerName in database
- **Silent Installation:** Uses `/qb` (shows progress bar)
- **Database Update:** Stops service, updates SQLite database, restarts service
- **Registration:** Service registers with API using client ID on first run

### Eclipse Online Server
- **Installation Path:** `C:\inetpub\Eclipse`
- **Silent Mode:** InstallShield `/S` wrapper with MSI properties
- **Installation Time:** 2-5 minutes

### IIS Application Pools & Site (Task 18)
- **Application Pools:** 
  - Eclipse (.NET CLR v4.0, Integrated, ApplicationPoolIdentity)
  - EclipseAuraCore (No Managed Code, Integrated, ApplicationPoolIdentity)
- **IIS Site:** 'Eclipse' with custom domain and HTTPS binding
- **Certificate:** Self-signed 'EclipsePlaceholderSelfsigned' (replace with production cert)
- **Applications:** /AuraApi, /eclipseApi, /oAuth, /oAuth2
- **Firewall:** Creates inbound rule for specified port

### Eclipse Smart Hub
- **Installation Path:** `C:\Program Files\Ultimate Business Systems\EclipseSmartHub` (default)
- **Silent Mode:** InstallShield `/S`
- **Services:** Creates SmartHubService and SmartHubUIService
- **Installation Time:** 1-3 minutes

### Win-ACMEv2
- **Installation Type:** Portable (extracted to Dependencies\Win-ACMEv2)
- **Interactive:** Launches after extraction for SSL certificate configuration
- **Usage:** Script waits for you to configure certificates and close Win-ACMEv2 before continuing

---

## Logging

### Interactive Mode
- No automatic logging
- Output shown in console only

### Unattended Mode
- **Log File:** `C:\EclipseTechInstall.log`
- **Contents:** 
  - Configuration settings
  - Task execution details
  - Success/failure status
  - Execution times
  - Error messages and stack traces

---

## Troubleshooting

### Script Won't Run
- Ensure you're running as Administrator
- Check PowerShell execution policy: `Set-ExecutionPolicy Bypass -Scope Process`

### SSMS Installation Shows "No installer processes detected"
- This is normal on some systems
- Verify installation via registry or file paths after 30-second wait
- SSMS may install to different location on Windows 11

### Eclipse Update Service Not Registering
- Check `C:\EclipseTechInstall.log` for database update status
- Verify ConsumerName in database: `C:\Program Files (x86)\Ultimate Business Systems\EclipseUpdateService\AppData\EclipseDB.db3`
- Restart the service: `Restart-Service EclipseUpdateService`

### IIS Site Not Accessible
- Check Windows Firewall rule was created for your port
- Verify IIS site 'Eclipse' is running in IIS Manager
- Check application pools 'Eclipse' and 'EclipseAuraCore' are started
- Verify self-signed certificate is bound to HTTPS binding

### SQL Server Installation Fails
- Check installer log: `{SqlDrive}:\SQL2025\SQLServer_Install.log`
- Check SQL Server setup logs: `C:\Program Files\Microsoft SQL Server\170\Setup Bootstrap\Log`
- Ensure sufficient disk space on target drive
- Verify no other SQL Server instances are installing

---

## Post-Installation Tasks (Eclipse Aura Setup)

After running tasks 11-20, complete these manual steps:

### 1. Map Database in Eclipse Schema Update
- Launch: `C:\Program Files (x86)\Eclipse\EclipseSchemaUpdate\EclipseSchemaUpdate.exe`
- Configure database connection to your Eclipse database
- Apply schema updates

### 2. Activate Eclipse Aura
- Use AWS Domain Manager to activate your domain
- Configure DNS settings as required

### 3. Whitelist Domain
- Connect to `ws.dev.ultimate.net.au` via SQL Server Management Studio
- Add your domain to the whitelist using the connection information provided

---

## Advanced Usage

### Custom Installation Paths
In unattended mode, you can specify:
- Custom Eclipse Install folder (default: `C:\Eclipse Install`)
- SQL Server installation drive (default: `C:`)

### Selective Task Execution
Use unattended mode to run only specific tasks:
- **SQL Server only:** Select tasks 3, 4, 5
- **IIS setup only:** Select tasks 11, 12, 13, 14
- **Eclipse Aura only:** Select tasks 15, 17, 18, 19, 20

### Silent Installation Arguments
All silent installations use appropriate flags:
- **MSI installers:** `/qb` or `/quiet /norestart`
- **InstallShield:** `/S /v"/qn INSTALLDIR=\"path\""`
- **Visual Studio Installer:** `--quiet --norestart`
- **winget packages:** `--silent --accept-package-agreements --accept-source-agreements`

---

## File Locations

### Downloads
- **Default:** `C:\Eclipse Install\Dependencies`
- **SQL Server:** `{SqlDrive}:\SQL2025`

### Installed Components
- **Eclipse DMS:** `C:\Program Files (x86)\Eclipse`
- **Eclipse Update Service:** `C:\Program Files (x86)\Ultimate Business Systems\EclipseUpdateService`
- **Eclipse Online Server:** `C:\inetpub\Eclipse`
- **Eclipse Smart Hub:** `C:\Program Files\Ultimate Business Systems\EclipseSmartHub`
- **Win-ACMEv2:** `C:\Eclipse Install\Dependencies\Win-ACMEv2`
- **SSMS 21:** `C:\Program Files (x86)\Microsoft SQL Server Management Studio 21` or `C:\Program Files\Microsoft SQL Server Management Studio 21`

### Log Files
- **Unattended Mode:** `C:\EclipseTechInstall.log`
- **SQL Server:** `{SqlDrive}:\SQL2025\SQLServer_Install.log`
- **MSI Verbose Logs:** `%TEMP%\MSI_*.log`

---

## Support

For issues or questions, contact the Eclipse Support Team.

**Support Resources:**
- Ultimate Business Systems: ultimate.net.au
- Development Server: ws.dev.ultimate.net.au
- Production Server: ws.ultimate.net.au

