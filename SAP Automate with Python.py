import sys
import subprocess
import time
import win32com.client
import tkinter as tk
from tkinter import messagebox, ttk

class SAPAutomation:
    def __init__(self):
        self.connection = None
        self.session = None
        self.new_session = None
        self.sap_gui_object = None
        self.application = None

    def connect_to_sap(self):
        """Connect to an existing SAP GUI session without disturbing it"""
        try:
            # Get the SAP GUI Scripting object
            self.sap_gui_object = win32com.client.GetObject("SAPGUI")
            
            if not self.sap_gui_object:
                messagebox.showerror("Error", "No SAP GUI instance found. Please login to SAP first.")
                return False
                
            # Get the first application instance
            self.application = self.sap_gui_object.GetScriptingEngine
            
            if not self.application:
                messagebox.showerror("Error", "No SAP application instance found.")
                return False
                
            # Get the first connection
            if self.application.Connections.Count > 0:
                self.connection = self.application.Connections(0)
                # Get the first session (the existing one)
                self.session = self.connection.Children(0)
                return True
            else:
                messagebox.showerror("Error", "No active SAP connections found.")
                return False
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to connect to SAP: {str(e)}")
            return False

    def open_new_session(self):
        """Open a new session window without touching the existing one"""
        try:
            if not self.session:
                return False
                
            # Method 1: Try using the New Session button function
            try:
                # This simulates clicking the "New Session" button in SAP GUI
                self.session.CreateSession()
                time.sleep(2)  # Wait for session creation
                
                # Find the newly created session (typically the last one)
                session_count = self.connection.Children.Count
                if session_count > 1:
                    self.new_session = self.connection.Children(session_count - 1)
                    if self.new_session:
                        '''try: #enable this to make the session invisible
                            print("Making the new session invisible...")
                            # Get the main window
                            window = self.new_session.FindById("wnd[0]")
                            if window:
                                # Set the window to invisible
                                window.Visible = False
                                print("Session window hidden successfully")
                        except Exception as e:
                            print(f"Warning: Could not hide window: {str(e)}")'''

                        return True
            except Exception as e:
                print(f"Method 1 failed: {str(e)}")   
                  
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open new session: {str(e)}")
            return False

    def perform_operations(self): #Enter SAP steps inside here
        """Perform the desired SAP operations in the new session"""
        try:
            if not self.new_session:
                return False
            
            # Store the session in a local variable to avoid using 'self'
            session = self.new_session
            
            # Your specific SAP steps
            session.FindById("wnd[0]").Maximize()  # Added parentheses to call the method
            
            
            # Just for demonstration, we'll add a small delay
            time.sleep(2)
            
            return True
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during SAP operations: {str(e)}")
            return False

    def close_session(self):
        """Close only the newly created session"""
        try:
            if self.new_session:
                try:
                    # First try to end the transaction
                    self.new_session.EndTransaction()
                    time.sleep(0.5)
                except:
                    pass
                    
                try:
                    # Try to close the session properly
                    window = self.new_session.FindById("wnd[0]")
                    if window:
                        window.Close()
                        time.sleep(0.5)
                        
                        # Check if we need to confirm closing
                        try:
                            popup = self.new_session.FindById("wnd[1]")
                            if popup:
                                popup.SendVKey(12)  # F12 is often "Yes" in close confirmation dialogs
                                time.sleep(0.5)
                        except:
                            pass
                except:
                    pass
                
                self.new_session = None
                return True
            return False
        except Exception as e:
            messagebox.showerror("Error", f"Error closing session: {str(e)}")
            return False

    def run_automation(self): #Main flow happens here
        """Run the complete automation sequence"""
        success = self.connect_to_sap()
        if not success:
            return "Failed to connect to SAP"
            
        success = self.open_new_session()
        if not success:
            return "Failed to open new SAP session"
            
        success = self.perform_operations()
        if not success:
            self.close_session()
            return "Failed during SAP operations"
            
        success = self.close_session()
        if not success:
            return "Failed to close SAP session properly"
            
        return "Automation completed successfully"


class SAPAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SAP Automation Tool")
        self.root.geometry("400x300")
        self.root.resizable(False, False)
        
        # Set style
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Arial", 12))
        self.style.configure("TLabel", font=("Arial", 12))
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create title label
        title_label = ttk.Label(main_frame, text="SAP Automation", font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Create description
        desc_text = "This tool will open a new SAP session and perform operations\nwithout disturbing your existing session."
        desc_label = ttk.Label(main_frame, text=desc_text)
        desc_label.pack(pady=(0, 20))
        
        # Create Run button
        self.run_button = ttk.Button(main_frame, text="Run Automation", command=self.run_automation)
        self.run_button.pack(pady=10)
        
        # Create status label
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_label = ttk.Label(main_frame, textvariable=self.status_var)
        status_label.pack(pady=10)
        
        # Create progress bar
        self.progress = ttk.Progressbar(main_frame, orient=tk.HORIZONTAL, length=300, mode='indeterminate')
        self.progress.pack(pady=10)
        
        # Create SAP automation instance
        self.sap_automation = SAPAutomation()
        
    def run_automation(self):
        """Run the SAP automation when the button is clicked"""
        # Disable button and show progress
        self.run_button.config(state=tk.DISABLED)
        self.status_var.set("Running automation...")
        self.progress.start()
        
        # Schedule the task to run after a small delay to allow GUI to update
        self.root.after(100, self.perform_automation)
    
    def perform_automation(self):
        """Perform the actual automation in a separate thread"""
        try:
            result = self.sap_automation.run_automation()
            self.status_var.set(result)
        except Exception as e:
            self.status_var.set(f"Error: {str(e)}")
        finally:
            # Re-enable button and stop progress
            self.run_button.config(state=tk.NORMAL)
            self.progress.stop()


def main():
    root = tk.Tk()
    app = SAPAutomationGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
