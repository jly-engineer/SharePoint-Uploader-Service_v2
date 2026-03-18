import tkinter as tk
from tkinter import filedialog, messagebox
import configparser
import os
import sys
import shutil
import subprocess
import ctypes
import time

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class InstallerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SharePoint Uploader Manager")
        self.root.geometry("500x700")

        self.create_label_entry("Tenant ID", "tenant_id")
        self.create_label_entry("Client ID", "client_id")
        self.create_label_entry("Client Secret", "client_secret", show="*")
        self.create_label_entry("SharePoint Site ID", "site_id")
        self.create_label_entry("Document Library ID", "drive_id")
        
        self.create_folder_picker("Monitor Folder", "monitor_folder")
        self.create_label_entry("SharePoint Target Folder (Optional)", "target_folder")
        
        self.install_btn = tk.Button(root, text="Install / Update Service", command=self.install, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
        self.install_btn.pack(pady=(20, 10), fill='x', padx=20)

        self.uninstall_btn = tk.Button(root, text="Uninstall Service", command=self.uninstall, bg="#D32F2F", fg="white", font=("Arial", 10, "bold"))
        self.uninstall_btn.pack(pady=5, fill='x', padx=20)

        self.status_label = tk.Label(root, text="Ready", fg="blue")
        self.status_label.pack(pady=5)

    def create_label_entry(self, label_text, key, show=None):
        frame = tk.Frame(self.root)
        frame.pack(fill='x', padx=10, pady=5)
        lbl = tk.Label(frame, text=label_text, anchor='w')
        lbl.pack(fill='x')
        entry = tk.Entry(frame, show=show)
        entry.pack(fill='x')
        setattr(self, f"entry_{key}", entry)

    def create_folder_picker(self, label_text, key):
        frame = tk.Frame(self.root)
        frame.pack(fill='x', padx=10, pady=5)
        lbl = tk.Label(frame, text=label_text, anchor='w')
        lbl.pack(fill='x')
        container = tk.Frame(frame)
        container.pack(fill='x')
        entry = tk.Entry(container)
        entry.pack(side='left', fill='x', expand=True)
        setattr(self, f"entry_{key}", entry)
        btn = tk.Button(container, text="Browse", command=lambda: self.browse_folder(entry))
        btn.pack(side='right', padx=5)

    def browse_folder(self, entry_widget):
        folder = filedialog.askdirectory()
        if folder:
            entry_widget.delete(0, tk.END)
            entry_widget.insert(0, folder)

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def kill_processes(self):
        """Kills known processes."""
        subprocess.run(["taskkill", "/F", "/IM", "uploader_service.exe"], capture_output=True)
        # Try to kill legacy NSSM if present, but ignore failure
        subprocess.run(["taskkill", "/F", "/IM", "nssm.exe"], capture_output=True)
        time.sleep(1)

    def cleanup_legacy_file(self, file_path):
        """Attempts to delete a file. If locked, ignores it (since we don't need it)."""
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except OSError:
                print(f"Warning: Could not remove locked legacy file {file_path}. Ignoring.")
                # We rename it so it's hidden/inactive, just in case
                try:
                    os.rename(file_path, file_path + ".trash")
                except:
                    pass

    def install(self):
        config_data = {
            'tenant_id': self.entry_tenant_id.get(),
            'client_id': self.entry_client_id.get(),
            'client_secret': self.entry_client_secret.get(),
            'sharepoint_site_id': self.entry_site_id.get(),
            'document_library_id': self.entry_drive_id.get(),
            'monitor_folder': self.entry_monitor_folder.get(),
            'sharepoint_target_folder': self.entry_target_folder.get(),
            'log_file': 'service.log'
        }

        if not all([config_data['tenant_id'], config_data['client_id'], config_data['monitor_folder']]):
            messagebox.showerror("Error", "Please fill in all required fields.")
            return

        install_dir = os.path.join(os.environ['ProgramFiles'], 'SharePointUploader')
        service_name = "SharePointUploaderService"

        try:
            self.update_status("Stopping old services...")
            subprocess.run(["sc", "stop", service_name], capture_output=True)
            time.sleep(2)
            # Remove old service registration
            subprocess.run(["sc", "delete", service_name], capture_output=True)
            
            self.kill_processes()

            if not os.path.exists(install_dir):
                os.makedirs(install_dir)

            # CLEANUP: Try to remove old nssm.exe if it exists, but don't crash if locked
            legacy_nssm = os.path.join(install_dir, "nssm.exe")
            self.cleanup_legacy_file(legacy_nssm)

            self.update_status("Copying files...")
            src_service = resource_path("uploader_service.exe")
            dst_service = os.path.join(install_dir, "uploader_service.exe")
            
            # This is the VITAL file. We must overwrite this.
            if os.path.exists(dst_service):
                try:
                    os.remove(dst_service)
                except OSError:
                    self.kill_processes()
                    os.remove(dst_service)
            
            shutil.copy(src_service, dst_service)

            # Write config
            config = configparser.ConfigParser()
            config['Settings'] = config_data
            with open(os.path.join(install_dir, 'config.ini'), 'w') as f:
                config.write(f)

            self.update_status("Registering service...")
            # Use PyWin32 self-registration
            subprocess.run([dst_service, '--startup', 'auto', 'install'], check=True, cwd=install_dir)

            self.update_status("Starting service...")
            subprocess.run(["sc", "start", service_name], check=True)

            self.update_status("Done!")
            messagebox.showinfo("Success", "Installation Complete! Service is running.")
            self.root.quit()

        except Exception as e:
            self.update_status("Error occurred.")
            messagebox.showerror("Installation Failed", f"An error occurred:\n{str(e)}")

    def uninstall(self):
        if not messagebox.askyesno("Confirm Uninstall", "Are you sure you want to remove the Service?"):
            return

        service_name = "SharePointUploaderService"
        install_dir = os.path.join(os.environ['ProgramFiles'], 'SharePointUploader')

        try:
            self.update_status("Stopping service...")
            subprocess.run(["sc", "stop", service_name], capture_output=True)
            time.sleep(2)
            
            self.update_status("Deleting service registry...")
            subprocess.run(["sc", "delete", service_name], capture_output=True)

            self.update_status("Killing processes...")
            self.kill_processes()

            self.update_status("Removing files...")
            if os.path.exists(install_dir):
                # Iterate and delete files individually to handle locks gracefully
                for filename in os.listdir(install_dir):
                    file_path = os.path.join(install_dir, filename)
                    try:
                        if os.path.isfile(file_path) or os.path.islink(file_path):
                            os.unlink(file_path)
                        elif os.path.isdir(file_path):
                            shutil.rmtree(file_path)
                    except Exception as e:
                        print(f"Skipping locked file: {file_path}")
                
                # Try to remove the folder itself
                try:
                    os.rmdir(install_dir)
                except OSError:
                    # If folder is not empty (because nssm is locked), that's fine. 
                    # We leave the folder with just the trash file in it.
                    pass
            
            self.update_status("Uninstalled.")
            messagebox.showinfo("Success", "Service uninstalled.\n(Some locked legacy files may remain until reboot, but the service is gone.)")
            self.root.quit()

        except Exception as e:
            self.update_status("Error.")
            messagebox.showerror("Uninstall Failed", f"An error occurred:\n{str(e)}")

if __name__ == "__main__":
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    else:
        root = tk.Tk()
        app = InstallerApp(root)
        root.mainloop()