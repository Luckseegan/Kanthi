import os
from tkinter import ttk
import tkinter as tk
from tkinter import filedialog, messagebox
import datetime
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle, Alignment, Border, Side, PatternFill
from openpyxl.utils.cell import get_column_letter
from dateutil.relativedelta import relativedelta
import json
from tkinter import filedialog, messagebox, ttk
import warnings
import json
import pandas as pd
import warnings
import win32com.client
import time



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
        self.geometry("600x400")
        self.config = load_config()

        self.agent_file = self.config.get("agent_file", "")
        self.my_file = self.config.get("my_file", "")
        self.sap_username = self.config.get("sap_username", "")  # Load SAP username
        self.sap_password = self.config.get("sap_password", "") 
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

    def update_sap_treeview(self, df):
        """Update the Treeview to display extracted SAP data."""
        
        # Remove old tree if exists
        for widget in self.sap_output_frame.winfo_children():
            widget.destroy()

        # Create a new frame to hold the Treeview and scrollbars
        tree_frame = tk.Frame(self.sap_output_frame)
        tree_frame.pack(fill="both", expand=False)

        # Create scrollbars
        tree_scroll = ttk.Scrollbar(tree_frame, orient="vertical")
        tree_scroll.pack(side="right", fill="y")

        tree_hscroll = ttk.Scrollbar(tree_frame, orient="horizontal")
        tree_hscroll.pack(side="bottom", fill="x")

        # Create the Treeview widget
        self.tree = ttk.Treeview(tree_frame, columns=list(df.columns), show="headings", yscrollcommand=tree_scroll.set, xscrollcommand=tree_hscroll.set)

        # Configure columns
        self.tree.column("#0", width=0, stretch=tk.NO)  # Hide first empty column
        for col in df.columns:
            self.tree.column(col, anchor="center", width=120)
            self.tree.heading(col, text=col, anchor="center")

        # Insert Data
        for i, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

        # Attach Treeview to the scrollbars
        tree_scroll.config(command=self.tree.yview)
        tree_hscroll.config(command=self.tree.xview)

        # Pack the Treeview widget into the frame
        self.tree.pack(fill="both", expand=False)

    def create_sap_tab(self):
        """SAP Tab with left (inputs) and right (output table) layout."""

        # 🟢 Create PanedWindow to Split Left (Inputs) & Right (Table Output)
        self.sap_paned_window = tk.PanedWindow(self.sap_tab, orient=tk.HORIZONTAL)
        self.sap_paned_window.pack(fill="both", expand=True)

        # 🟢 Left Section (1/3) → SAP Inputs & Rates/Weights
        self.sap_input_frame = tk.Frame(self.sap_paned_window, padx=10, pady=10, width=350, height=500)
        self.sap_input_frame.pack_propagate(False)

        # 🔹 SAP Username
        tk.Label(self.sap_input_frame, text="SAP Username:", font=("Arial", 8)).pack(anchor="w", pady=5)
        self.sap_username_entry = tk.Entry(self.sap_input_frame, font=("Arial", 8))
        self.sap_username_entry.insert(0, self.sap_username)
        self.sap_username_entry.pack(fill="x", padx=10, pady=5)

        # 🔹 SAP Password
        tk.Label(self.sap_input_frame, text="SAP Password:", font=("Arial", 8)).pack(anchor="w", pady=5)
        self.sap_password_entry = tk.Entry(self.sap_input_frame, font=("Arial", 8), show="*")
        self.sap_password_entry.insert(0, self.sap_password)
        self.sap_password_entry.pack(fill="x", padx=10, pady=5)

        # 🔹 SAP GUI Values
        tk.Label(self.sap_input_frame, text="SAPLMEGUI Value for Bodyline:", font=("Arial", 8)).pack(anchor="w", pady=5)
        self.sap_gui_entry = tk.Entry(self.sap_input_frame, font=("Arial", 8))
        self.sap_gui_entry.insert(0, self.sap_gui_value)
        self.sap_gui_entry.pack(fill="x", padx=10, pady=5)

        tk.Label(self.sap_input_frame, text="Default SAPLMEGUI Value:", font=("Arial", 8)).pack(anchor="w", pady=5)
        self.default_sap_gui_entry = tk.Entry(self.sap_input_frame, font=("Arial", 8))
        self.default_sap_gui_entry.insert(0, self.default_sap_gui_value)
        self.default_sap_gui_entry.pack(fill="x", padx=10, pady=5)

        # 🟠 Rates & Weights Section
        tk.Label(self.sap_input_frame, text="Rates and Weights:", font=("Arial", 8, "bold")).pack(anchor="w", pady=10)
        self.rates_entries = {}

        for key, value in self.rates_weights.items():
            frame = tk.Frame(self.sap_input_frame)
            frame.pack(fill="x", padx=10, pady=2)
            tk.Label(frame, text=key.replace("_", " ").title() + ":", font=("Arial", 8)).pack(side="left")
            entry = tk.Entry(frame, font=("Arial", 8), width=10)
            entry.insert(0, str(value))
            entry.pack(side="right")
            self.rates_entries[key] = entry

        # 🟢 Buttons
        self.save_sap_button = tk.Button(self.sap_input_frame, text="Save SAP Config", font=("Arial", 8),
                                        command=self.save_sap_config, bg="#4CAF50", fg="white", relief="raised",
                                        padx=5, pady=5)
        self.save_sap_button.pack(pady=5, fill="x")

        self.process_sap_button = tk.Button(self.sap_input_frame, text="Process SAP", font=("Arial", 10, 'bold'),
                                            command=self.process_sap, bg="#FF9800", fg="white", relief="raised",
                                            padx=5, pady=5)
        self.process_sap_button.pack(pady=5, fill="x")

        # 🟢 Progress Bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.sap_input_frame, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress_bar.pack(pady=8, fill="x")

        # Add Left Frame to PanedWindow
        self.sap_paned_window.add(self.sap_input_frame, width=350)

        # 🟠 Right Section (2/3) → Data Output & Future Work
        self.sap_output_frame = tk.Frame(self.sap_paned_window, width=650, height=500)
        self.sap_output_frame.pack_propagate(False)

        # **Top Half - Treeview Table with Scrolling**
        self.tree_frame = tk.Frame(self.sap_output_frame, height=250)  # Half of the right section
        self.tree_frame.pack(fill="both", expand=True)

        self.tree_scroll = ttk.Scrollbar(self.tree_frame, orient="vertical")
        self.tree_scroll.pack(side="right", fill="y")

        self.tree_hscroll = ttk.Scrollbar(self.tree_frame, orient="horizontal")
        self.tree_hscroll.pack(side="bottom", fill="x")

        self.tree = ttk.Treeview(
            self.tree_frame,
            show="headings",
            yscrollcommand=self.tree_scroll.set,
            xscrollcommand=self.tree_hscroll.set
        )

        # Attach scrollbars
        self.tree_scroll.config(command=self.tree.yview)
        self.tree_hscroll.config(command=self.tree.xview)

        # Pack Table
        self.tree.pack(fill="both", expand=True)

        # **Bottom Half - Reserved for Future Work**
        self.future_work_frame = tk.Frame(self.sap_output_frame, height=250, bg="lightgray")  # Placeholder
        self.future_work_frame.pack(fill="both", expand=True)

        # Add Excel update button in the bottom frame
        self.update_button = tk.Button(self.future_work_frame, text="Update Excel", command=self.update_excel)
        self.update_button.pack(pady=20)

        # Add Output Frame to PanedWindow
        self.sap_paned_window.add(self.sap_output_frame, width=650)

        # Function to display DataFrame dynamically
    
    def display_dataframe(self, df):
            """Dynamically update Treeview table with DataFrame contents."""
            # Clear previous columns & data
            self.tree["columns"] = list(df.columns)
            
            for col in df.columns:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=100)  # Default width, can be adjusted

            # Insert data
            for _, row in df.iterrows():
                self.tree.insert("", "end", values=list(row))


    def save_sap_config(self):
        self.sap_gui_value = self.sap_gui_entry.get()
        self.default_sap_gui_value = self.default_sap_gui_entry.get()
        self.sap_username = self.sap_username_entry.get()
        self.sap_password = self.sap_password_entry.get()
        self.rates_weights = {key: float(entry.get()) for key, entry in self.rates_entries.items()}

        self.config["sap_gui_value"] = self.sap_gui_value
        self.config["default_sap_gui_value"] = self.default_sap_gui_value
        self.config["sap_username"] = self.sap_username
        self.config["sap_password"] = self.sap_password
        self.config["rates_weights"] = self.rates_weights
        save_config(self.config)
        messagebox.showinfo("Success", "SAP configuration saved successfully.")


    # Other methods (select_agent_file, select_my_file, update_latest_sheets, start_processing) remain unchanged

    def select_agent_file(self):
        self.agent_file = select_file("Agent")
        self.config["agent_file"] = self.agent_file
        save_config(self.config)
        self.agent_label.config(text=f"Agent File: {os.path.basename(self.agent_file)}")
        self.update_latest_sheets()

    def select_my_file(self):
        self.my_file = select_file("My")
        self.config["my_file"] = self.my_file
        save_config(self.config)
        self.my_label.config(text=f"My File: {os.path.basename(self.my_file)}")
        self.update_latest_sheets()

    def update_latest_sheets(self):
        # Clear the previous sheets display
        for widget in self.sheets_display_frame.winfo_children():
            widget.destroy()

        # Get latest sheets from agent and my files
        agent_latest_sheet = ""
        my_latest_sheet = ""

        if self.agent_file and os.path.exists(self.agent_file):
            try:
                agent_wb = load_workbook(self.agent_file, data_only=True)
                agent_sheets = agent_wb.sheetnames
                agent_latest_sheet = get_latest_sheet(agent_sheets) or "No valid sheets"
            except Exception as e:
                agent_latest_sheet = "Error loading file"

        if self.my_file and os.path.exists(self.my_file):
            try:
                my_wb = load_workbook(self.my_file, data_only=True)
                my_sheets = my_wb.sheetnames
                my_latest_sheet = get_latest_sheet(my_sheets) or "No valid sheets"
            except Exception as e:
                my_latest_sheet = "Error loading file"

        # Display latest sheets in small boxes
        tk.Label(self.sheets_display_frame, text=f"Agent Latest Sheet: {agent_latest_sheet}", font=("Arial", 10), bg="#E0E0E0", relief="solid", padx=10, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Label(self.sheets_display_frame, text=f"My Latest Sheet: {my_latest_sheet}", font=("Arial", 10), bg="#E0E0E0", relief="solid", padx=10, pady=5).pack(side=tk.LEFT, padx=5)

    def start_processing(self):
        if not self.agent_file or not self.my_file:
            show_popup("Please select both agent and my files before proceeding.")
            return
        
        if is_file_open(self.agent_file):
            show_popup(f"Please close the file '{os.path.basename(self.agent_file)}' and try again.")
            return

        if is_file_open(self.my_file):
            show_popup(f"Please close the file '{os.path.basename(self.my_file)}' and try again.")
            return

        try:
            # Load the workbooks
            agent_wb = load_workbook(self.agent_file, data_only=True)
            my_wb = load_workbook(self.my_file)

            # Get available sheet names from the agent's workbook
            agent_sheets = agent_wb.sheetnames
            my_sheets = my_wb.sheetnames

            latest_sheet = get_latest_sheet(agent_sheets)

            if latest_sheet:
                print(f"Latest sheet in agent workbook: {latest_sheet}")
            else:
                print("No valid sheets found in the agent workbook.")
                latest_sheet = agent_sheets[0]
            self.latest_sheet = latest_sheet
            agent_ws = agent_wb[latest_sheet]

            if latest_sheet in my_sheets:
                show_popup(f"Sheet '{latest_sheet}' already exists in your workbook.")
                my_ws = my_wb[latest_sheet]
            else:
                show_popup(f"Sheet '{latest_sheet}' does not exist in your workbook. Creating a new sheet...")
                my_ws = my_wb.create_sheet(title=latest_sheet)

                if len(my_sheets) > 0:
                    previous_month_ws = my_wb[my_sheets[-1]]
                    for col in range(1, previous_month_ws.max_column + 1):
                        header_cell = my_ws.cell(row=1, column=col)
                        header_cell.value = previous_month_ws.cell(row=1, column=col).value
                        header_cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                    show_popup(f"Headers copied from the previous month's sheet '{my_sheets[-1]}' to the new sheet.")

            last_row_my_sheet = my_ws.max_row
            while last_row_my_sheet > 1 and not any(my_ws.cell(row=last_row_my_sheet, column=col).value for col in range(1, my_ws.max_column + 1)):
                last_row_my_sheet -= 1

            last_so_number = my_ws.cell(row=last_row_my_sheet, column=1).value

            agent_last_row = None
            for row in range(1, agent_ws.max_row + 1):
                if agent_ws.cell(row=row, column=1).value == last_so_number:
                    agent_last_row = row
                    break

            start_row = (agent_last_row + 1) if agent_last_row else 2

            rows_to_append = []
            for row in range(start_row, agent_ws.max_row + 1):
                if agent_ws.cell(row=row, column=1).value:
                    rows_to_append.append([agent_ws.cell(row=row, column=col).value for col in range(1, agent_ws.max_column + 1)])

            date_style = NamedStyle(name="short_date", number_format="MM/DD/YYYY")
            if "short_date" not in my_wb.named_styles:
                my_wb.add_named_style(date_style)

            date_columns = set()
            for col in range(1, agent_ws.max_column + 1):
                for row in range(2, min(10, agent_ws.max_row)):
                    cell_value = agent_ws.cell(row=row, column=col).value
                    if isinstance(cell_value, datetime.datetime):
                        date_columns.add(col)
                        break

            border_style = Border(left=Side(border_style="thin"), right=Side(border_style="thin"), top=Side(border_style="thin"), bottom=Side(border_style="thin"))

            for i, row_data in enumerate(rows_to_append, start=last_row_my_sheet + 1):
                for col, value in enumerate(row_data, start=1):
                    cell = my_ws.cell(row=i, column=col, value=value)
                    if col in date_columns and isinstance(value, datetime.datetime):
                        cell.style = "short_date"
                    cell.border = border_style

            my_wb.save(self.my_file)
            messagebox.showinfo("Success", f"{len(rows_to_append)} new rows appended to {os.path.basename(self.my_file)}.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


    def process_sap(self):
        """Process SAP data and display with a progress bar."""
        
        if not self.my_file:
            show_popup("Please select 'My File' before processing SAP.")
            return

        if not hasattr(self, 'latest_sheet') or not self.latest_sheet:
            show_popup("Latest sheet not found. Please start processing first.")
            return

        try:
            # Load the Excel file
            df = pd.read_excel(self.my_file, sheet_name=self.latest_sheet)

            # Filter out rows where the "status" column is blank
            df_filtered = df[df['status'].isna()]
            
            if df_filtered.empty:
                show_popup("No rows to process. All rows have a status.")
                return
            
            # Get the total number of rows for progress tracking
            total_rows = len(df_filtered)
            self.progress_var.set(0)  # Reset progress bar
            
            # SAP Login
            session = sap_login(username='maheswaranl', password='Srilanka@2024', connection_name='PDM', transaction_code='me23n')
            plant_mapping = get_updated_plant_mapping(self.sap_gui_value)
            default_mapping = get_updated_default_mapping(self.default_sap_gui_value)

            processed_rows = []
            for i, row in df_filtered.iterrows():
                # Process each row
                processed_row = process_valid_pos_with_sap(pd.DataFrame([row]), session, plant_mapping, default_mapping, self.rates_weights)
                processed_rows.append(processed_row)
                
                # Update progress
                progress_percent = ((i + 1) / total_rows) * 100
                self.progress_var.set(progress_percent)
                self.sap_tab.update_idletasks()  # Refresh UI to show progress
            
            # Combine processed data
            self.df_with_merchant_and_terms = pd.concat(processed_rows, ignore_index=True)
            
            # Select Relevant Columns
            relevant_columns = [0, 3, 5, 6, 7, 12, 13, 14, 17, 18, 28, 29, 30]
            df_display = self.df_with_merchant_and_terms.iloc[:, relevant_columns]

            # Display the result in Treeview
            self.update_sap_treeview(df_display)

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during SAP processing: {e}")

    def update_excel(self):
        """Update the Excel sheet when the button is pressed."""
        if self.df_with_merchant_and_terms.empty:
            show_popup("No processed data available to update the Excel sheet.")
            return
        if is_file_open(self.agent_file):
            show_popup(f"Please close the file '{os.path.basename(self.agent_file)}' and try again.")
            return

        if is_file_open(self.my_file):
            show_popup(f"Please close the file '{os.path.basename(self.my_file)}' and try again.")
            return
        

        # Load the workbook and the sheet
        wb = load_workbook(self.my_file)
        ws = wb[self.latest_sheet]

        # Read the main sheet into a DataFrame
        df = pd.read_excel(self.my_file, sheet_name=self.latest_sheet)

        # Ensure df_with_merchant_and_terms contains only the rows to be updated
        for index, row in self.df_with_merchant_and_terms.iterrows():
            # Find the row in the main dataframe that needs to be updated (matching on the first column, 'SO')
            matching_row = df[df.iloc[:, 0] == row.iloc[0]]  # Matching based on first column (SO)

            if not matching_row.empty:
                for col in self.df_with_merchant_and_terms.columns:
                    if col in df.columns:
                        # Find the Excel row number
                        excel_row_index = df.index[df.iloc[:, 0] == row.iloc[0]].tolist()[0] + 2  # Adjust for header row
                        
                        # Find the Excel column letter
                        excel_col_index = df.columns.get_loc(col) + 1
                        excel_col_letter = get_column_letter(excel_col_index)

                        # Get the value to update
                        new_value = row[col]

                        # Preserve date format if the column contains dates
                        if pd.api.types.is_datetime64_any_dtype(df[col]):
                            ws[f"{excel_col_letter}{excel_row_index}"].value = new_value  # Update value
                            ws[f"{excel_col_letter}{excel_row_index}"].number_format = "DD-MMM-YY"  # Apply short date format
                        else:
                            ws[f"{excel_col_letter}{excel_row_index}"].value = new_value  # Update non-date values

        # Preserve filters (reapply if they existed)
        if ws.auto_filter.ref:
            ws.auto_filter.ref = ws.auto_filter.ref

        # Save the workbook (preserving formatting, date formats, and filters)
        wb.save(self.my_file)

        print("Excel file updated successfully while preserving short date format.")





# Run the application
if __name__ == "__main__":
    app = FileUploaderApp()
    app.mainloop()