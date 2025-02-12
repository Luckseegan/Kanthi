import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import datetime
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Alignment, Border, Side, PatternFill
from openpyxl.utils.cell import get_column_letter
from dateutil.relativedelta import relativedelta
import json
import pandas as pd
import warnings
import win32com.client
import time

# Suppress warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

# Suppress warnings
warnings.simplefilter(action='ignore', category=FutureWarning)


# SAP Login Function
def sap_login(username, password, connection_name, transaction_code):
    try:
        # Connect to SAP GUI
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        if not SapGuiAuto:
            raise Exception("SAP GUI is not running.")
        
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Connections(0)  # Use existing connection if available
        session = connection.Children(0)
        
        # Check if already logged in
        if session.Info.IsLowSpeedConnection == 0:  # Example condition
            print("Already logged in. Navigating to transaction.")
        else:
            raise Exception("No active session. Please login manually.")
        
    except Exception as e:
        # If not logged in, create a new connection
        print("Logging in...")
        connection = application.OpenConnection(connection_name, True)
        session = connection.Children(0)
        
        # Enter credentials
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
        session.findById("wnd[0]/tbar[0]/btn[0]").press()
    
        print("Login successful.")

    except Exception as e:
        print(f"Error during login: {str(e)}")
        exit(1)

    # Navigate to the transaction
    print(f"Navigating to transaction {transaction_code}...")
    session.findById("wnd[0]/tbar[0]/okcd").text = transaction_code
    session.findById("wnd[0]/tbar[0]/btn[0]").press()
    print("Transaction opened successfully.")
    return session

# Function to process valid POs with SAP
def process_valid_pos_with_sap(df, session, plant_mapping,default_mapping, rates_weights):
    # Function to check if PO is valid
    def is_valid_po(po):
        return str(po).isdigit() and str(po).startswith("4") and len(str(po)) == 10

    # Function to extract Merchant and Incoterm from SAP for valid POs
    def extract_from_sap(po_list, plant_mapping):
        merchant = set()
        incoterm = set()

        for purchase_order_number, plant_name in po_list:
            try:
                # Extract PO details from SAP
                session.findById("wnd[0]/tbar[1]/btn[17]").press()
                session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").text = purchase_order_number
                session.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-EBELN").caretPosition = len(purchase_order_number)
                session.findById("wnd[1]/tbar[0]/btn[0]").press()
                
                # Plant-specific handling
                if plant_name in plant_mapping:
                    session.findById(plant_mapping[plant_name]["tab_selector9"]).select()
                    created_by = session.findById(plant_mapping[plant_name]["created_by"]).text
                    session.findById(plant_mapping[plant_name]["tab_selector1"]).select()
                    inco_term = session.findById(plant_mapping[plant_name]["inco_term"]).text
                else:
                    session.findById(default_mapping["tab_selector9"]).select()
                    created_by = session.findById(default_mapping["created_by"]).text
                    session.findById(default_mapping["tab_selector1"]).select()
                    inco_term = session.findById(default_mapping["inco_term"]).text
                                
                    
                # Add extracted values to the sets
                merchant.add(created_by)
                incoterm.add(inco_term)
                
            except Exception as e:
                print(f"Error processing PO {purchase_order_number}: {str(e)}")
                continue

        # Return extracted merchant and incoterm as comma-separated values
        return ", ".join(merchant), ", ".join(incoterm)

    # Process each row and validate POs
    processed_rows = []
    for index, row in df.iterrows():
        plant_name = row[3]
        po_numbers = str(row[5]).split(",")

        # Create a PO list for the current row
        po_list = [(po_number.strip(), plant_name) for po_number in po_numbers]

        # Filter valid POs
        valid_pos = [po for po in po_list if is_valid_po(po[0])]
        
        if not valid_pos:
            # Skip if no valid POs found
            row[29]='Manual Check'
            row[30] = "Invalid PO"
            processed_rows.append(row)
            continue
        
        plant_name = row[3]  # 4th column
        if plant_name.lower() == "silueta":
        # Skip extraction if plant_name is silueta
            row[6] = "N/A"  # Set Merchant as "N/A"
            row[7] = "N/A"  # Set Terms as "N/A"
            row[29]='Manual Check'
            row[30] = "Extraction Skipped for Silueta"
        else:
            # Extract merchant and incoterm details from SAP
            merchant, incoterm = extract_from_sap(valid_pos, plant_mapping)
            row[6] = merchant  # Merchant column

        terms = row[7]  # Terms (Incoterm) column
        shipment_type = row[17] 
        booking_gross_kgs = float(row[13])
        volume_kgs = float(row[14])
        max_volume = max(booking_gross_kgs, volume_kgs) # Assuming shipment type is in the 11th column

        if plant_name.lower() == "bodyline":
            # Set financial approval status
            row[12] = "APP"  # Assuming the 15th column is for financial approval
            
            # Calculate cost based on shipment type and terms
            if shipment_type.lower() == "hong kong":
                if "EXW" in terms:
                    row[18] = round(max_volume * rates_weights["hong_kong_exw_rate"]/10)*10
                elif "FOB" in terms:
                    row[18] = round(max_volume * rates_weights["hong_kong_fob_rate"] / 10) * 10
            elif shipment_type.lower() == "china":
                if "EXW" in terms:
                    row[18] = round(max_volume * rates_weights["china_exw_rate"] / 10) * 10
                elif "FOB" in terms:
                    row[18] = round(max_volume * rates_weights["china_fob_rate"] / 10) * 10
            
            # Set status
            row[28] = "Financial Approval(Bodyline)"
        
        elif plant_name.lower() == "silueta":
            # Silueta-specific logic
            if shipment_type.lower() == "hong kong":
                if max_volume > rates_weights["max_volume_hong_kong"]:
                    row[12] = "APP"  # Assign financial approval as "APP"
                    if "EXW" in terms:
                        row[18] = round(max_volume * rates_weights["hong_kong_exw_rate"] / 10) * 10
                    elif "FOB" in terms:
                        row[18] = round(max_volume * rates_weights["hong_kong_fob_rate"] / 10) * 10
                    row[28] = "Financial Approval(Silueta)"  # Assign status
                else:
                    row[12] = "CE"
                    row[28] = "Delivery term confirmation"  # Assign status

            elif shipment_type.lower() == "china":
                if max_volume > rates_weights["max_volume_china"]:
                    row[12] = "APP"  # Assign financial approval as "APP"
                    if "EXW" in terms:
                        row[18] = round(max_volume * rates_weights["china_exw_rate"] / 10) * 10
                    elif "FOB" in terms:
                        row[18] = round(max_volume * rates_weights["china_fob_rate"]/10)*10
                    row[28] = "Financial Approval(Silueta)"  # Assign status
                else:
                    row[12] = "CE"
                    row[28] = "Delivery term confirmation"  # Assign status

        elif "a058" in plant_name.lower() and any(po.strip().startswith("43") for po in str(row[5]).split(",")):  
            # Assuming POs are in the 5th column and separated by commas
            # Set status for Unichela-A058 with PO starting with 43
            row[28] = "All Good"  # Assign status
            row[29] = "Received approval"  # Assign status for mailing (assuming column 22)

        else:
            # Check if any PO starts with "49"
            if any(po.strip().startswith("49") for po in str(row[5]).split(",")):  # Column 5 for POs
                row[12] = "Cstd"  # Assign "Cstd" in the financial approval column (Column 13)
                print(f"POs {row[4]} contain one starting with '49'. Financial approval set to 'Cstd'.")

                # Check if the SAP extracted incoterm (incoterm) is in the terms column (Column 8)
                extracted_incoterm = str(row[7])  # Column 8 for terms
                print(f"Extracted incoterm from row[7]: {extracted_incoterm}, comparing with SAP incoterm: {incoterm}.")
                
                if incoterm in extracted_incoterm:
                    # If incoterm matches, set status as "All Good" and prepare mail
                    row[28] = "All Good"  # Set status as "All Good" (Column 21)
                    row[29] = "Received approval"  # Set mail status (Column 22)
                    print(f"Incoterm {incoterm} found in extracted terms. Status set to 'All Good'. Mail to Received Approval.")
                else:
                    # If incoterm does not match, set status to "Delivery Term Confirmation"
                    row[28] = "Delivery term confirmation"  # Set status as "Delivery Term Confirmation" (Column 21)
                    print(f"Incoterm {incoterm} not found in extracted terms. Status set to 'Delivery Term Confirmation'.")

            else:
                if shipment_type.lower() == "hong kong":
                    if max_volume > rates_weights["max_volume_hong_kong"]:
                        row[12] = "APP"  # Assign financial approval as "APP"
                        if "EXW" in terms:
                            row[18] = round(max_volume * rates_weights["hong_kong_exw_rate"]/10)*10
                        elif "FOB" in terms:
                            row[18] = round(max_volume * rates_weights["hong_kong_fob_rate"]/10)*10
                        row[28] = "Approval"  # Assign status
                        print(f"Hong Kong shipment, volume > {rates_weights['max_volume_hong_kong']}, Financial approval set to 'APP'.")

                    else:
                        # Check if the extracted incoterm matches and adjust accordingly
                        if incoterm in terms:
                            row[28] = "All Good"  # Status for matching incoterm
                            row[29] = "Received approval"  # Mail status
                            print(f"Hong Kong shipment, volume <= {rates_weights['max_volume_hong_kong']}, Incoterm match found. Status: 'All Good'. Mail Received Approval.")
                        else:
                            row[12] = "CE"  # Assign "CE" in financial approval
                            row[28] = "Delivery term confirmation"  # Assign status for no match
                            print(f"Hong Kong shipment, volume <= {rates_weights['max_volume_hong_kong']}, Incoterm mismatch. Status: 'Delivery Term Confirmation'.")

                elif shipment_type.lower() == "china":
                    if max_volume > rates_weights["max_volume_china"]:
                        row[12] = "APP"  # Assign financial approval as "APP"
                        if "EXW" in terms:
                            row[18] = round(max_volume * rates_weights["china_exw_rate"]/10)*10
                        elif "FOB" in terms:
                            row[18] = round(max_volume * rates_weights["china_fob_rate"]/10)*10
                        row[28] = "Approval"  # Assign status
                        print(f"China shipment, volume > {rates_weights['max_volume_china']}, Financial approval set to 'APP'.")
                    
                    else:
                        # Check if the extracted incoterm matches and adjust accordingly
                        if incoterm in terms:
                            row[28] = "All Good"  # Status for matching incoterm
                            row[21] = "Received approval"  # Mail status
                            print(f"China shipment, volume <= {rates_weights['max_volume_china']}, Incoterm match found. Status: 'All Good'. Mail to Received Approval.")
                        else:
                            row[12] = "CE"  # Assign "CE" in financial approval
                            row[28] = "Delivery term confirmation"  # Assign status for no match
                            print(f"China shipment, volume <= {rates_weights['max_volume_china']}, Incoterm mismatch. Status: 'Delivery Term Confirmation'.")

        processed_rows.append(row)

    # Create a new DataFrame with only the relevant columns
    processed_df = pd.DataFrame(processed_rows, columns=df.columns)
    return processed_df

def get_updated_plant_mapping(sap_gui_value):
    return {
        "bodyline": {
            "tab_selector9": f"wnd[0]/usr/subSUB0:SAPLMEGUI:{sap_gui_value}/subSUB1:SAPLMEVIEWS:1100/"
                             "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/"
                             "tabpTABHDT9",
            "created_by": f"wnd[0]/usr/subSUB0:SAPLMEGUI:{sap_gui_value}/subSUB1:SAPLMEVIEWS:1100/"
                          "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/"
                          "tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/txtMEPO1222-EKNAM",
            "tab_selector1": f"wnd[0]/usr/subSUB0:SAPLMEGUI:{sap_gui_value}/subSUB1:SAPLMEVIEWS:1100/"
                             "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/"
                             "tabpTABHDT1",
            "inco_term": f"wnd[0]/usr/subSUB0:SAPLMEGUI:{sap_gui_value}/subSUB1:SAPLMEVIEWS:1100/"
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/"
                         "tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-INCO1"
        }
    }

# Function to update placeholders in default (else block)
def get_updated_default_mapping(default_sap_gui_value):
    return {
        "tab_selector9": f"wnd[0]/usr/subSUB0:SAPLMEGUI:{default_sap_gui_value}/subSUB1:SAPLMEVIEWS:1100/"
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/"
                         "tabpTABHDT9",
        "created_by": f"wnd[0]/usr/subSUB0:SAPLMEGUI:{default_sap_gui_value}/subSUB1:SAPLMEVIEWS:1100/"
                      "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/"
                      "tabpTABHDT9/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/txtMEPO1222-EKNAM",
        "tab_selector1": f"wnd[0]/usr/subSUB0:SAPLMEGUI:{default_sap_gui_value}/subSUB1:SAPLMEVIEWS:1100/"
                         "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/"
                         "tabpTABHDT1",
        "inco_term": f"wnd[0]/usr/subSUB0:SAPLMEGUI:{default_sap_gui_value}/subSUB1:SAPLMEVIEWS:1100/"
                     "subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/"
                     "tabpTABHDT1/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1226/ctxtMEPO1226-INCO1"
    }


# Function to check if a file is open
def is_file_open(file_path):
    try:
        with open(file_path, 'a'):
            pass
        return False
    except IOError:
        return True

# Function to show a pop-up message
def show_popup(message):
    messagebox.showinfo("File Open", message)

# Load configuration or initialize if not present
def load_config():
    config_file = 'config.json'
    if os.path.exists(config_file):
        with open(config_file, 'r') as f:
            return json.load(f)
    return {}

def save_config(config):
    with open('config.json', 'w') as f:
        json.dump(config, f)

# Function to allow the user to select files
def select_file(file_type):
    file_path = filedialog.askopenfilename(title=f"Select {file_type} File", filetypes=[("Excel Files", "*.xlsx")])
    return file_path

def parse_month_year(sheet_name):
    try:
        return datetime.datetime.strptime(sheet_name.upper(), "%b %Y")  # Parse uppercase month-year format
    except ValueError:
        return None

# Get the latest sheet based on the month-year format
def get_latest_sheet(sheets):
    parsed_sheets = []
    for sheet in sheets:
        parsed_date = parse_month_year(sheet)
        if parsed_date:
            parsed_sheets.append((sheet, parsed_date))
    
    # Sort sheets by date and return the most recent one
    if parsed_sheets:
        parsed_sheets.sort(key=lambda x: x[1], reverse=True)  # Sort descending by date
        return parsed_sheets[0][0]
    return None

# Define the main application
class FileUploaderApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Kanthi - China/HK Air")
        self.geometry("800x600")
        self.config = load_config()

        self.agent_file = self.config.get("agent_file", "")
        self.my_file = self.config.get("my_file", "")
        self.sap_gui_value = self.config.get("sap_gui_value", "0013")
        self.default_sap_gui_value = self.config.get("default_sap_gui_value", "0013")
        self.rates_weights = self.config.get("rates_weights", {
            "hong_kong_exw_rate": 10.0,
            "hong_kong_fob_rate": 12.0,
            "china_exw_rate": 8.0,
            "china_fob_rate": 9.0,
            "max_volume_hong_kong": 1000.0,
            "max_volume_china": 1500.0
        })

        self.create_widgets()

    def create_widgets(self):
        # Title Label
        self.title_label = tk.Label(self, text="Kanthi - China/HK Air", font=("Arial", 18, 'bold'), bg="#4CAF50", fg="white")
        self.title_label.pack(fill=tk.X, pady=10)

        # Notebook for tabs
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # File Upload Tab
        self.file_upload_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.file_upload_tab, text="File Upload")
        self.create_file_upload_tab()

        # SAP Processing Tab
        self.sap_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.sap_tab, text="SAP Processing")
        self.create_sap_tab()

    def create_file_upload_tab(self):
        # Frame for content
        content_frame = tk.Frame(self.file_upload_tab)
        content_frame.pack(pady=20, padx=20, fill=tk.X)

        # Display agent file info
        self.agent_label = tk.Label(content_frame, text=f"Agent File: {os.path.basename(self.agent_file) if self.agent_file else 'Not Selected'}", font=("Arial", 12))
        self.agent_label.grid(row=0, column=0, sticky="w", pady=5)

        # Select Agent File Button
        self.agent_button = tk.Button(content_frame, text="Choose Agent File", font=("Arial", 12), command=self.select_agent_file, bg="#4CAF50", fg="white", relief="raised", padx=10, pady=5)
        self.agent_button.grid(row=0, column=1, padx=10, pady=5)

        # Display My File info
        self.my_label = tk.Label(content_frame, text=f"My File: {os.path.basename(self.my_file) if self.my_file else 'Not Selected'}", font=("Arial", 12))
        self.my_label.grid(row=1, column=0, sticky="w", pady=5)

        # Select My File Button
        self.my_button = tk.Button(content_frame, text="Choose My File", font=("Arial", 12), command=self.select_my_file, bg="#4CAF50", fg="white", relief="raised", padx=10, pady=5)
        self.my_button.grid(row=1, column=1, padx=10, pady=5)

        # Frame for latest sheets
        latest_sheets_frame = tk.Frame(self.file_upload_tab)
        latest_sheets_frame.pack(pady=10, padx=20, fill=tk.X)

        # Label for latest sheets
        tk.Label(latest_sheets_frame, text="Latest Sheets", font=("Arial", 12, 'bold')).pack(anchor="w")

        # Frame to display latest sheets horizontally
        self.sheets_display_frame = tk.Frame(latest_sheets_frame)
        self.sheets_display_frame.pack(fill=tk.X, pady=5)

        # Start Process Button
        self.start_button = tk.Button(self.file_upload_tab, text="Start Processing", font=("Arial", 14, 'bold'), command=self.start_processing, bg="#FF9800", fg="white", relief="raised", padx=20, pady=10)
        self.start_button.pack(pady=10)

        # Update latest sheets display
        self.update_latest_sheets()

    def create_sap_tab(self):
        # Frame for SAP inputs
        sap_frame = tk.Frame(self.sap_tab)
        sap_frame.pack(pady=20, padx=20, fill=tk.X)

        # SAP GUI Value
        tk.Label(sap_frame, text="SAPLMEGUI Value for Bodyline:", font=("Arial", 12)).grid(row=0, column=0, sticky="w", pady=5)
        self.sap_gui_entry = tk.Entry(sap_frame, font=("Arial", 12))
        self.sap_gui_entry.insert(0, self.sap_gui_value)
        self.sap_gui_entry.grid(row=0, column=1, padx=10, pady=5)

        # Default SAP GUI Value
        tk.Label(sap_frame, text="Default SAPLMEGUI Value:", font=("Arial", 12)).grid(row=1, column=0, sticky="w", pady=5)
        self.default_sap_gui_entry = tk.Entry(sap_frame, font=("Arial", 12))
        self.default_sap_gui_entry.insert(0, self.default_sap_gui_value)
        self.default_sap_gui_entry.grid(row=1, column=1, padx=10, pady=5)

        # Rates and Weights
        tk.Label(sap_frame, text="Rates and Weights:", font=("Arial", 12, 'bold')).grid(row=2, column=0, sticky="w", pady=10)
        self.rates_entries = {}
        row_idx = 3
        for key, value in self.rates_weights.items():
            tk.Label(sap_frame, text=key.replace("_", " ").title() + ":", font=("Arial", 12)).grid(row=row_idx, column=0, sticky="w", pady=5)
            self.rates_entries[key] = tk.Entry(sap_frame, font=("Arial", 12))
            self.rates_entries[key].insert(0, str(value))
            self.rates_entries[key].grid(row=row_idx, column=1, padx=10, pady=5)
            row_idx += 1

        # Save SAP Config Button
        self.save_sap_button = tk.Button(sap_frame, text="Save SAP Config", font=("Arial", 12), command=self.save_sap_config, bg="#4CAF50", fg="white", relief="raised", padx=10, pady=5)
        self.save_sap_button.grid(row=row_idx, column=0, columnspan=2, pady=10)

        # Process SAP Button
        self.process_sap_button = tk.Button(self.sap_tab, text="Process SAP", font=("Arial", 14, 'bold'), command=self.process_sap, bg="#FF9800", fg="white", relief="raised", padx=20, pady=10)
        self.process_sap_button.pack(pady=10)

    def save_sap_config(self):
        self.sap_gui_value = self.sap_gui_entry.get()
        self.default_sap_gui_value = self.default_sap_gui_entry.get()
        self.rates_weights = {key: float(entry.get()) for key, entry in self.rates_entries.items()}

        self.config["sap_gui_value"] = self.sap_gui_value
        self.config["default_sap_gui_value"] = self.default_sap_gui_value
        self.config["rates_weights"] = self.rates_weights
        save_config(self.config)
        messagebox.showinfo("Success", "SAP configuration saved successfully.")

    def process_sap(self):
        if not self.my_file:
            show_popup("Please select 'My File' before processing SAP.")
            return

        try:
            df = pd.read_excel(self.my_file)
            session = sap_login(username='maheswaranl', password='Srilanka@2024', connection_name='PDM', transaction_code='me23n')
            plant_mapping = get_updated_plant_mapping(self.sap_gui_value)
            default_mapping = get_updated_default_mapping(self.default_sap_gui_value)
            df_with_merchant_and_terms = process_valid_pos_with_sap(df, session, plant_mapping, default_mapping, self.rates_weights)
            df_with_merchant_and_terms.to_excel(self.my_file, index=False)
            messagebox.showinfo("Success", "SAP processing completed and file updated.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during SAP processing: {e}")

    # Other methods (select_agent_file, select_my_file, update_latest_sheets, start_processing) remain unchanged

# Run the application
if __name__ == "__main__":
    app = FileUploaderApp()
    app.mainloop()