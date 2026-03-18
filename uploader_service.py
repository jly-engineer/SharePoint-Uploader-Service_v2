import win32serviceutil
import win32service
import win32event
import servicemanager
import configparser
import logging
import os
import sys
import time
import msal
import requests
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# Determine if running as a script or frozen exe
if getattr(sys, 'frozen', False):
    BUNDLE_DIR = os.path.dirname(sys.executable)
else:
    BUNDLE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_config():
    """Reads config.ini from the bundle directory."""
    config = configparser.ConfigParser()
    config_path = os.path.join(BUNDLE_DIR, 'config.ini')
    config.read(config_path)
    if 'Settings' not in config:
        return {}
    return config['Settings']

def setup_logging():
    """Sets up logging to file ONLY."""
    settings = get_config()
    log_file = settings.get('log_file', 'service.log')
    log_path = os.path.join(BUNDLE_DIR, log_file)
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
        handlers=[logging.FileHandler(log_path)]
    )

class SharePointUploader:
    """Helper class to handle the actual upload logic."""
    def get_token(self, config):
        try:
            authority = f"https://login.microsoftonline.com/{config.get('tenant_id')}"
            app = msal.ConfidentialClientApplication(
                config.get('client_id'),
                authority=authority,
                client_credential=config.get('client_secret')
            )
            scopes = ["https://graph.microsoft.com/.default"]
            result = app.acquire_token_for_client(scopes=scopes)
            return result.get('access_token')
        except Exception as e:
            logging.error(f"Token error: {e}")
            return None

    SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024       # 4 MB — Graph API hard limit for simple PUT
    CHUNK_SIZE = 10 * 1024 * 1024               # 10 MB chunks (must be a multiple of 320 KiB)

    def upload(self, file_path, config, token):
        if not os.path.exists(file_path):
            return

        monitor_root = config.get('monitor_folder')
        try:
            rel_path = os.path.relpath(file_path, monitor_root)
        except ValueError:
            rel_path = os.path.basename(file_path)

        remote_path = rel_path.replace(os.sep, '/')

        target_folder = config.get('sharepoint_target_folder', '').strip('/')
        if target_folder:
            final_path = f"{target_folder}/{remote_path}"
        else:
            final_path = remote_path

        final_path = final_path.replace('//', '/')
        site_id = config.get('sharepoint_site_id')
        drive_id = config.get('document_library_id')
        base = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}"

        try:
            file_size = os.path.getsize(file_path)
            logging.info(f"Uploading: {rel_path} ({file_size} bytes)")

            if file_size <= self.SIMPLE_UPLOAD_LIMIT:
                self._simple_upload(file_path, rel_path, base, final_path, token)
            else:
                self._chunked_upload(file_path, rel_path, file_size, base, final_path, token)

        except Exception as e:
            logging.error(f"Upload exception for {rel_path}: {e}")

    def _simple_upload(self, file_path, rel_path, base, final_path, token):
        url = f"{base}/root:/{final_path}:/content?@microsoft.graph.conflictBehavior=replace"
        headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/octet-stream'}
        with open(file_path, 'rb') as f:
            response = requests.put(url, headers=headers, data=f)
        if response.status_code in [200, 201]:
            logging.info(f"Done: {rel_path}")
        else:
            logging.error(f"Upload failed: {response.status_code} - {response.text}")

    def _chunked_upload(self, file_path, rel_path, file_size, base, final_path, token):
        # Step 1: Create an upload session
        session_url = f"{base}/root:/{final_path}:/createUploadSession"
        headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
        body = {"item": {"@microsoft.graph.conflictBehavior": "replace"}}
        resp = requests.post(session_url, headers=headers, json=body)
        if resp.status_code != 200:
            logging.error(f"Failed to create upload session for {rel_path}: {resp.status_code} - {resp.text}")
            return

        upload_url = resp.json().get('uploadUrl')
        if not upload_url:
            logging.error(f"No uploadUrl in session response for {rel_path}")
            return

        # Step 2: Upload in chunks — no Authorization header needed for the upload URL itself
        with open(file_path, 'rb') as f:
            offset = 0
            while offset < file_size:
                chunk = f.read(self.CHUNK_SIZE)
                if not chunk:
                    break
                end = offset + len(chunk) - 1
                chunk_headers = {
                    'Content-Length': str(len(chunk)),
                    'Content-Range': f'bytes {offset}-{end}/{file_size}',
                }
                chunk_resp = requests.put(upload_url, headers=chunk_headers, data=chunk)
                # 202 = more chunks expected, 200/201 = complete
                if chunk_resp.status_code not in [200, 201, 202]:
                    logging.error(f"Chunk upload failed at byte {offset} for {rel_path}: "
                                  f"{chunk_resp.status_code} - {chunk_resp.text}")
                    return
                offset += len(chunk)
                logging.info(f"Progress {rel_path}: {offset}/{file_size} bytes")

        logging.info(f"Done: {rel_path}")

class UploadHandler(FileSystemEventHandler):
    def __init__(self):
        self.uploader = SharePointUploader()
        # Dictionary to track last upload time per file to prevent double-uploads
        self.last_processed = {}

    def wait_for_file_ready(self, file_path, timeout=60):
        """
        Loops and waits until the file is no longer locked.
        """
        start_time = time.time()
        while (time.time() - start_time) < timeout:
            try:
                with open(file_path, 'rb'):
                    pass
                return True
            except IOError as e:
                if e.errno == 13: 
                    time.sleep(2) 
                else:
                    raise e
        return False

    def process(self, event):
        if event.is_directory: return
        
        filename = os.path.basename(event.src_path)
        
        # 1. FILTER: Ignore Thumbs.db and temporary files
        if filename.lower() == 'thumbs.db' or filename.startswith('~$') or filename.endswith('.tmp'):
            return

        # 2. DEBOUNCE: Check if we processed this file recently
        current_time = time.time()
        last_time = self.last_processed.get(event.src_path, 0)
        
        # If processed in the last 15 seconds, skip this event
        if (current_time - last_time) < 15:
            return

        # 3. READY CHECK: Wait for scanner to release lock
        try:
            if not self.wait_for_file_ready(event.src_path):
                logging.warning(f"Timeout waiting for file unlock: {filename}. Skipping.")
                return
        except Exception:
            return

        # Update the timestamp immediately before upload attempt
        self.last_processed[event.src_path] = time.time()

        try:
            config = get_config()
            token = self.uploader.get_token(config)
            if token:
                self.uploader.upload(event.src_path, config, token)
        except Exception as e:
            logging.error(f"Handler error: {e}")

    def on_created(self, event): self.process(event)
    def on_modified(self, event): self.process(event)

class AppServerSvc(win32serviceutil.ServiceFramework):
    _svc_name_ = "SharePointUploaderService"
    _svc_display_name_ = "SharePoint Uploader Service"
    _svc_description_ = "Monitors folders and uploads files to SharePoint."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.observer = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        logging.info("Service stopping...")
        win32event.SetEvent(self.stop_event)
        if self.observer:
            self.observer.stop()
            self.observer.join()
        logging.info("Service stopped.")

    def SvcDoRun(self):
        try:
            setup_logging()
            servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                                  servicemanager.PYS_SERVICE_STARTED,
                                  (self._svc_name_, ''))
            logging.info("Service starting...")
            
            config = get_config()
            path = config.get('monitor_folder')

            if path and os.path.isdir(path):
                self.observer = Observer()
                self.observer.schedule(UploadHandler(), path, recursive=True)
                self.observer.start()
                logging.info(f"Monitoring: {path}")
            else:
                logging.error(f"Invalid monitor path: {path}")

            win32event.WaitForSingleObject(self.stop_event, win32event.INFINITE)

        except Exception as e:
            logging.error(f"Service crash: {e}")
            servicemanager.LogErrorMsg(str(e))

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(AppServerSvc)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(AppServerSvc)