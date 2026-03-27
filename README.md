# SharePoint Uploader Service

A Windows background service that monitors a local folder and automatically uploads new or modified files to a SharePoint document library using the Microsoft Graph API. Includes a GUI installer for easy configuration and deployment.

## How It Works

- **`uploader_service.py`** — A `pywin32`-based Windows service that watches a folder with `watchdog`. When a file is created or modified, it authenticates via MSAL (client credentials flow) and uploads the file to SharePoint. Files under 4 MB use a simple PUT; larger files use Graph API's chunked upload session. Temporary files, `Thumbs.db`, and `~$` Office lock files are ignored. A 15-second debounce prevents duplicate uploads.
- **`installer_gui.py`** — A `tkinter` GUI that collects your Azure AD credentials and folder paths, writes a `config.ini`, copies the service executable to `Program Files\SharePointUploader`, and registers/starts the Windows service — no command line required.

## Prerequisites

### Azure AD App Registration

Before deploying, you need an Azure AD (Entra ID) app registration with the following Microsoft Graph **application** permissions (not delegated):

- `Sites.ReadWrite.All` or `Files.ReadWrite.All`

Grant admin consent after adding the permissions. You will need:

| Field | Where to find it |
|---|---|
| **Tenant ID** | Azure Portal → Entra ID → Overview |
| **Client ID** | App Registration → Overview |
| **Client Secret** | App Registration → Certificates & Secrets |
| **SharePoint Site ID** | Graph Explorer: `GET /v1.0/sites/{hostname}:/sites/{site-name}` → copy `id` |
| **Document Library ID** | Graph Explorer: `GET /v1.0/sites/{site-id}/drives` → copy the drive `id` |

### Python Dependencies

```
pip install pywin32 msal requests watchdog pyinstaller
```

## Building Windows Executables

Both Python scripts must be compiled into standalone `.exe` files before the installer can bundle them. Run these commands from the project root on a **Windows** machine.

### 1. Build the Service Executable

```cmd
pyinstaller --onefile --noupx --hidden-import=win32timezone uploader_service.py
```

The output will be at `dist\uploader_service.exe`.

> **`--hidden-import=win32timezone`** is required — `pywin32` services need this import at runtime but PyInstaller won't detect it automatically.

### 2. Build the Installer GUI Executable

The installer bundles `uploader_service.exe` inside itself so it can copy it during installation:

```cmd
pyinstaller --onefile --windowed --add-data "dist\uploader_service.exe;." installer_gui.py
```

| Flag | Purpose |
|---|---|
| `--onefile` | Produces a single portable `.exe` |
| `--windowed` | Suppresses the console window for the GUI |
| `--add-data "dist\uploader_service.exe;."` | Bundles the service exe into the installer |

The final installer will be at `dist\installer_gui.exe`.

## Installation
### Windows
1. Run the build_installer.bat
The build_installer.bat will compile both the installer_gui.exe and the uploader_service.exe into a single installation SharePointUploaderSetup.exe

1. Run `dist\installer_gui.exe` as **Administrator** (it will prompt for elevation automatically).
2. Fill in your Azure AD credentials and paths:
   - **Tenant ID / Client ID / Client Secret** — from your app registration
   - **SharePoint Site ID / Document Library ID** — from Microsoft Graph
   - **Monitor Folder** — the local folder to watch for new/changed files
   - **SharePoint Target Folder** *(optional)* — subfolder path within the document library (e.g., `Uploads/Scans`). Leave blank to upload to the library root.
3. Click **Install / Update Service**.

The installer will:
- Stop and remove any existing version of the service
- Copy `uploader_service.exe` to `C:\Program Files\SharePointUploader\`
- Write `config.ini` with your settings
- Register and start the `SharePointUploaderService` Windows service (set to auto-start)

### Uninstalling

Run the installer again and click **Uninstall Service**. This stops the service, removes its registry entry, and deletes the install directory.

## Configuration

The service reads `config.ini` from its install directory. The installer writes this file automatically, but you can edit it manually if needed:

```ini
[Settings]
tenant_id         = your-tenant-id
client_id         = your-client-id
client_secret     = your-client-secret
sharepoint_site_id      = your-site-id
document_library_id     = your-drive-id
monitor_folder          = C:\Path\To\Watch
sharepoint_target_folder = Optional/Subfolder
log_file          = service.log
```

After editing `config.ini` manually, restart the service:

```cmd
sc stop SharePointUploaderService
sc start SharePointUploaderService
```

## Logs

The service writes logs to `service.log` in `C:\Program Files\SharePointUploader\`. Check this file to verify uploads or diagnose errors.

## Service Management

```cmd
sc start SharePointUploaderService
sc stop SharePointUploaderService
sc query SharePointUploaderService
```
