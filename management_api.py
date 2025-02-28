# By Andrei Epure, Microsoft Ltd. 2025. Use at your own risk. No warranties are given.
# DISCLAIMER:
# THIS CODE IS SAMPLE CODE. THESE SAMPLES ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND.
# MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING WITHOUT LIMITATION ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR
# A PARTICULAR PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLES REMAINS WITH YOU. IN NO EVENT SHALL
# MICROSOFT OR ITS SUPPLIERS BE LIABLE FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS PROFITS,
# BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS) ARISING OUT OF THE USE OF OR INABILITY TO USE THE
# SAMPLES, EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. BECAUSE SOME STATES DO NOT ALLOW THE EXCLUSION OR LIMITATION
# OF LIABILITY FOR CONSEQUENTIAL OR INCIDENTAL DAMAGES, THE ABOVE LIMITATION MAY NOT APPLY TO YOU.

"""
.SYNOPSIS Retrieves logs from the Office 365 Management API

.DESCRIPTION This script allows you to interact with the Office 365 Management API.  You must register your application in Azure AND grant admin consent prior to use.
For application registration instructions, please see https://learn.microsoft.com/en-us/previous-versions/office/developer/o365-enterprise-developers/jj984325(v=office.15)#register-your-application-in-azure-ad

Special thanks to David Barrett for the inspiration on this https://github.com/David-Barrett-MS/PowerShell/blob/main/Office%20365%20Management%20API/Test-ManagementActivityAPI.ps1
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkcalendar import DateEntry
import requests
from datetime import datetime, timedelta
import threading
import os
import json
import subprocess

# Function to execute PowerShell command (if needed for blob download)
def execute_powershell_script(script):
    try:
        result = subprocess.run(["powershell", "-Command", script], capture_output=True, text=True)
        if result.returncode == 0:
            return result.stdout
        else:
            return result.stderr
    except Exception as e:
        return f"Error executing PowerShell script: {e}"

# Function to download content blob (Python version)
def download_content_blob(content_url, auth_token, save_path):
    try:
        # Initialize variables
        audit_data = ""
        download_log = []
        file_extension = "json"

        # Send the GET request to download the content blob
        response = requests.get(content_url, headers={"Authorization": f"Bearer {auth_token}"})
        response.raise_for_status()
        audit_data = response.content

        if audit_data:
            if save_path:
                # Generate filename based on the URL's last part
                output_file_name = content_url.split("/")[-1]
                output_file = os.path.join(save_path, output_file_name)

                # Check if file already exists and if the content differs
                if os.path.exists(f"{output_file}.{file_extension}"):
                    with open(f"{output_file}.{file_extension}", 'r') as existing_file:
                        existing_blob = existing_file.read()
                    if existing_blob != audit_data.decode("utf-8"):
                        # If content differs, save as a new file
                        i = 1
                        while os.path.exists(f"{output_file}.{i}.{file_extension}"):
                            i += 1
                        output_file = f"{output_file}.{i}"
                    else:
                        # Data is already retrieved
                        download_log.append(f"Data already retrieved: {output_file}.{file_extension}")
                        output_file = ""  # Skip saving file
                if output_file:
                    # Save the data to the file
                    with open(f"{output_file}.{file_extension}", "wb") as file:
                        file.write(audit_data)
                    download_log.append(f"Saving data blob to: {output_file}.{file_extension}")
        else:
            download_log.append(f"No data returned from {content_url}")
    except requests.exceptions.RequestException as err:
        download_log.append(f"Request error: {err}")
    except Exception as e:
        download_log.append(f"An unexpected error occurred: {e}")
    
    return download_log

# Function to get an access token
def get_access_token(app_id, tenant_id, app_secret):
    auth_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    token_data = {
        "grant_type": "client_credentials",
        "client_id": app_id,
        "client_secret": app_secret,
        "scope": "https://manage.office.com/.default",
    }
    response = requests.post(auth_url, data=token_data)
    
    if response.status_code != 200:
        raise Exception(f"Failed to authenticate: {response.text}")

    return response.json().get("access_token")

# Function to fetch management activity logs
# Function to fetch management activity logs
def fetch_management_activity_logs():
    def background_task():
        try:
            # Gather user inputs
            app_id = app_id_entry.get().strip()
            tenant_id = tenant_id_entry.get().strip()
            app_secret = app_secret_entry.get().strip()
            save_path = save_path_var.get()
            selected_content_types = [var.get() for var in content_type_vars if var.get()]
            start_date = start_date_entry.get_date().strftime('%Y-%m-%d') + "T" + start_time_combobox.get().strip() + ":00Z"
            end_date = end_date_entry.get_date().strftime('%Y-%m-%d') + "T" + end_time_combobox.get().strip() + ":00Z"

            # Validate inputs
            if not app_id or not tenant_id or not app_secret:
                messagebox.showerror("Error", "App ID, Tenant ID, and App Secret are required!")
                return

            if not selected_content_types:
                messagebox.showerror("Error", "At least one Content Type must be selected!")
                return

            if not save_path:
                messagebox.showerror("Error", "Please select a folder to save the report!")
                return

            # Ensure the directory exists or create it
            if not os.path.exists(save_path):
                os.makedirs(save_path)

            processing_label.config(text="Fetching logs... Please wait.")
            root.update()

            # Get OAuth token
            token = get_access_token(app_id, tenant_id, app_secret)

            # API call to get content
            headers = {"Authorization": f"Bearer {token}"}
            all_content_data = []

            for content_type in selected_content_types:
                api_url = (
                    f"https://manage.office.com/api/v1.0/{tenant_id}/activity/feed/subscriptions/content"
                    f"?contentType={content_type}&startTime={start_date}&endTime={end_date}"
                )
                response = requests.get(api_url, headers=headers)

                if response.status_code != 200:
                    messagebox.showerror("Error", f"API call failed for {content_type}: {response.text}")
                    continue

                content_items = response.json()

                if not content_items:
                    messagebox.showinfo("Info", f"No logs found for {content_type} between {start_date} and {end_date}")
                    continue

                for item in content_items:
                    all_content_data.append(item)

                    # Extract and download content URL
                    content_url = item.get("contentUri")
                    if content_url:
                        download_log = download_content_blob(content_url, token, save_path)
                        all_content_data.append(download_log)

            # Remove the part that saves the "ManagementActivityLogs" file
            # The line below is no longer needed, so it has been removed:
            # output_path = os.path.join(save_path, f"ManagementActivityLogs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
            # with open(output_path, "w", encoding="utf-8") as file:
            #    json.dump(all_content_data, file, indent=4)

            processing_label.config(text="Processing complete!")
            messagebox.showinfo("Success", f"Logs saved to {save_path}")
        
        except requests.exceptions.RequestException as req_err:
            messagebox.showerror("Error", f"Request error: {req_err}")
        except Exception as e:
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        finally:
            processing_label.config(text="")

    threading.Thread(target=background_task, daemon=True).start()

# Function to browse folder for saving logs
def browse_folder():
    folder = filedialog.askdirectory()
    if folder:
        save_path_var.set(folder)

# Function to generate time options with 30-minute intervals for the combobox
def generate_time_options():
    times = []
    for hour in range(24):
        for minute in [0, 30]:
            times.append(f"{hour:02}:{minute:02}")
    return times

# GUI setup
root = tk.Tk()
root.title("Office 365 Management Activity Logs")

# Set application icon
icon_path = "C:/temp/Management API/Logo_management_api.ico"
try:
    root.iconbitmap(icon_path)
except Exception as e:
    print(f"Error loading icon: {e}")

main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=0, column=0, sticky="NSEW")

# App ID
ttk.Label(main_frame, text="App ID:").grid(row=0, column=0, sticky="W")
app_id_entry = ttk.Entry(main_frame, width=50)
app_id_entry.grid(row=0, column=1, padx=5, pady=5)

# Tenant ID
ttk.Label(main_frame, text="Tenant ID:").grid(row=1, column=0, sticky="W")
tenant_id_entry = ttk.Entry(main_frame, width=50)
tenant_id_entry.grid(row=1, column=1, padx=5, pady=5)

# App Secret
ttk.Label(main_frame, text="App Secret:").grid(row=2, column=0, sticky="W")
app_secret_entry = ttk.Entry(main_frame, show="*", width=50)
app_secret_entry.grid(row=2, column=1, padx=5, pady=5)

# Content Type Checkboxes
ttk.Label(main_frame, text="Content Types:").grid(row=3, column=0, sticky="W")
content_types = ["Audit.AzureActiveDirectory", "Audit.Exchange", "Audit.SharePoint", "Audit.General", "DLP.All" ]
content_type_vars = []
for i, content_type in enumerate(content_types):
    var = tk.StringVar()
    chk = ttk.Checkbutton(main_frame, text=content_type, variable=var, onvalue=content_type, offvalue="")
    chk.grid(row=3 + i, column=1, sticky="W")
    content_type_vars.append(var)

# Start Date Calendar
min_date = datetime.today() - timedelta(days=7)
max_date = datetime.today()
ttk.Label(main_frame, text="Start Date:").grid(row=8, column=0, sticky="W")
start_date_entry = DateEntry(main_frame, width=47, date_pattern="yyyy-mm-dd", mindate=min_date, maxdate=max_date)
start_date_entry.grid(row=8, column=1, padx=5, pady=5)

# Start Time (30-minute increments)
ttk.Label(main_frame, text="Start Time (HH:mm):").grid(row=8, column=2, sticky="W")
start_time_combobox = ttk.Combobox(main_frame, values=generate_time_options(), width=10)
start_time_combobox.grid(row=8, column=3, padx=5, pady=5)
start_time_combobox.set("00:00")  # Default value

# End Date Calendar
ttk.Label(main_frame, text="End Date:").grid(row=9, column=0, sticky="W")
end_date_entry = DateEntry(main_frame, width=47, date_pattern="yyyy-mm-dd", mindate=min_date, maxdate=max_date)
end_date_entry.grid(row=9, column=1, padx=5, pady=5)

# End Time (30-minute increments)
ttk.Label(main_frame, text="End Time (HH:mm):").grid(row=9, column=2, sticky="W")
end_time_combobox = ttk.Combobox(main_frame, values=generate_time_options(), width=10)
end_time_combobox.grid(row=9, column=3, padx=5, pady=5)
end_time_combobox.set("23:59")  # Default value

# Save Path
ttk.Label(main_frame, text="Save Path:").grid(row=10, column=0, sticky="W")
save_path_var = tk.StringVar(value="C:/temp/ManagementAPI-Logs")
save_path_entry = ttk.Entry(main_frame, textvariable=save_path_var, width=50)
save_path_entry.grid(row=10, column=1, padx=5, pady=5)

# Browse button
browse_button = ttk.Button(main_frame, text="Browse", command=browse_folder)
browse_button.grid(row=10, column=2, padx=5, pady=5)

# Process button
process_button = ttk.Button(main_frame, text="Fetch Logs", command=fetch_management_activity_logs)
process_button.grid(row=11, column=1, pady=10)

# Processing Label
processing_label = ttk.Label(main_frame, text="")
processing_label.grid(row=12, column=1, pady=10)

root.mainloop()
