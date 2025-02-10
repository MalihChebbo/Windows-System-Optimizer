import os
import subprocess
import sys
import ctypes
import winreg
import logging
from datetime import datetime

class WindowsOptimizer:
    def __init__(self):
        self.setup_logging()
        self.ensure_admin_privileges()

    def setup_logging(self):
        log_dir = os.path.join(os.environ['USERPROFILE'], 'WindowsOptimizer')
        os.makedirs(log_dir, exist_ok=True)
        
        log_file = os.path.join(log_dir, f'optimizer_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log')
        
        logging.basicConfig(
            filename=log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s: %(message)s'
        )
        self.logger = logging.getLogger()

    def ensure_admin_privileges(self):
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
            if not is_admin:
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", sys.executable, " ".join(sys.argv), None, 1
                )
                sys.exit()
        except Exception as e:
            self.logger.error(f"Admin privileges check failed: {e}")
            print(f"Error checking admin privileges: {e}")

    def clean_temp_files(self):
        print("\n[+] Cleaning Temporary Files...")
        self.logger.info("Starting temporary files cleanup")

        try:
            print("  - Running Disk Cleanup for system files...")
            subprocess.run(['cleanmgr', '/sagerun:1'], capture_output=True, text=True)
            self.logger.info("Disk Cleanup completed successfully")
        except Exception as e:
            print(f"  - Disk Cleanup failed: {e}")
            self.logger.error(f"Disk Cleanup failed: {e}")

    def optimize_drive(self):
        print("\n[+] Optimizing Drive...")
        self.logger.info("Starting drive optimization")

        try:
            subprocess.run(['defrag', 'C:', '/O'], capture_output=True, text=True)
            print("  - Drive optimization completed")
            self.logger.info("Drive optimization completed")
        except Exception as e:
            print(f"  - Drive optimization failed: {e}")
            self.logger.error(f"Drive optimization failed: {e}")

    def clean_browser_cache(self):
        print("\n[+] Cleaning Browser Caches...")
        self.logger.info("Starting browser cache cleanup")

        browsers = {
            'Chrome': os.path.join(os.environ['LOCALAPPDATA'], 'Google', 'Chrome', 'User Data', 'Default', 'Cache'),
            'Firefox': os.path.join(os.environ['APPDATA'], 'Mozilla', 'Firefox', 'Profiles'),
            'Edge': os.path.join(os.environ['LOCALAPPDATA'], 'Microsoft', 'Edge', 'User Data', 'Default', 'Cache')
        }

        for browser, cache_path in browsers.items():
            try:
                if os.path.exists(cache_path):
                    print(f"  - Cleaning {browser} cache...")
                    subprocess.run(['cmd', '/c', f'rmdir /s /q "{cache_path}"'], capture_output=True, text=True)
                    self.logger.info(f"{browser} cache cleaned")
            except Exception as e:
                print(f"  - Failed to clean {browser} cache: {e}")
                self.logger.error(f"{browser} cache cleanup failed: {e}")

    def disable_startup_programs(self):
        print("\n[+] Managing Startup Programs...")
        self.logger.info("Starting startup programs management")

        try:
            unnecessary_programs = [
                'Spotify', 'Steam', 'Discord', 'Dropbox', 
                'OneDrive', 'Zoom', 'Skype'
            ]

            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, 
                r"Software\Microsoft\Windows\CurrentVersion\Run", 
                0, 
                winreg.KEY_ALL_ACCESS
            )
            
            disabled_count = 0
            for program in unnecessary_programs:
                try:
                    winreg.DeleteValue(key, program)
                    print(f"  - Disabled startup program: {program}")
                    disabled_count += 1
                except FileNotFoundError:
                    pass
                
            winreg.CloseKey(key)
            print(f"  - Total startup programs disabled: {disabled_count}")
            self.logger.info(f"Disabled {disabled_count} startup programs")
        except Exception as e:
            print(f"  - Error managing startup programs: {e}")
            self.logger.error(f"Startup programs management failed: {e}")

    def reduce_visual_effects(self):
        print("\n[+] Optimizing Visual Effects...")
        self.logger.info("Starting visual effects optimization")

        try:
            subprocess.run(['powercfg', '-setactive', 'balanced'], capture_output=True, text=True)
            
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, 
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects", 
                0, 
                winreg.KEY_ALL_ACCESS
            )
            
            winreg.SetValueEx(key, "VisualFXSetting", 0, winreg.REG_DWORD, 2)
            winreg.CloseKey(key)
            
            print("  - Visual effects optimized for performance")
            self.logger.info("Visual effects optimized")
        except Exception as e:
            print(f"  - Error reducing visual effects: {e}")
            self.logger.error(f"Visual effects optimization failed: {e}")

    def display_menu(self):
        while True:
            print("\n===== Windows System Optimizer =====")
            print("1. Clean Temporary Files")
            print("2. Optimize Drive")
            print("3. Clean Browser Caches")
            print("4. Manage Startup Programs")
            print("5. Optimize Visual Effects")
            print("6. Run Complete System Optimization")
            print("7. Exit")
            
            try:
                choice = input("\nEnter your choice (1-7): ")
                
                if choice == '1':
                    self.clean_temp_files()
                elif choice == '2':
                    self.optimize_drive()
                elif choice == '3':
                    self.clean_browser_cache()
                elif choice == '4':
                    self.disable_startup_programs()
                elif choice == '5':
                    self.reduce_visual_effects()
                elif choice == '6':
                    print("\n[+] Running Complete System Optimization...")
                    self.clean_temp_files()
                    self.optimize_drive()
                    self.clean_browser_cache()
                    self.disable_startup_programs()
                    self.reduce_visual_effects()
                    print("\n[+] Complete System Optimization Finished!")
                elif choice == '7':
                    print("Exiting Windows System Optimizer. Goodbye!")
                    break
                else:
                    print("Invalid choice. Please enter a number between 1 and 7.")
                
                input("\nPress Enter to continue...")
            except KeyboardInterrupt:
                print("\nOperation cancelled by user.")
                break
            except Exception as e:
                print(f"An error occurred: {e}")

def main():
    print("Windows System Optimizer")
    print("WARNING: This tool requires administrator privileges.")
    print("Use with caution and ensure you have backups of important data.")
    
    optimizer = WindowsOptimizer()
    optimizer.display_menu()

if __name__ == "__main__":
    main()