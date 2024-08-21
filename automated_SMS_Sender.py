import pandas as pd
import re
import logging
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinter import ttk
from twilio.rest import Client
import time
import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

# Set up logging
logging.basicConfig(filename='opt_in.log', level=logging.INFO, 
                    format='%(asctime)s:%(levelname)s:%(message)s')

# Global variable to hold customer data
customer_data = None
opted_out_numbers = set()

# Function to validate phone numbers
def is_valid_phone_number(phone):
    return re.match(r'^\+\d+$', phone) is not None

# Function to load customer data
def load_customer_data():
    global customer_data
    try:
        # Inform user about data loading
        load_button.configure(text="Loading...", fg_color="orange")
        root.update_idletasks()
        
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return None
        
        customer_data = pd.read_excel(file_path)
        
        # Check for required columns
        if 'Name' not in customer_data.columns or 'Phone Number' not in customer_data.columns:
            raise ValueError("The spreadsheet must contain 'Name' and 'Phone Number' columns.")
        
        messagebox.showinfo("Success", "Customer data loaded successfully!")
        load_button.configure(text="Data Loaded", fg_color="green")  # Visual confirmation
        return customer_data
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load data: {str(e)}")
        logging.error(f"Failed to load data: {str(e)}")
        load_button.configure(text="Load Failed", fg_color="red")  # Visual indication of failure
        return None

# Function to send a single SMS message and return the message SID
def send_single_sms(client, twilio_phone_number, customer_name, customer_phone, message_body):
    try:
        message = client.messages.create(
            body=message_body,
            from_=twilio_phone_number,
            to=customer_phone
        )
        logging.info(f"Message sent to {customer_name} at {customer_phone} with SID: {message.sid}")
        return message.sid  # Return the message SID for status checking later
    except Exception as e:
        logging.error(f"Failed to send message to {customer_name} at {customer_phone}: {str(e)}")
        return None

# Function to check the status of a message after it has been sent
def check_message_status(client, message_sid, customer_name, customer_phone):
    try:
        # Loop to check message status until it is delivered or fails
        for _ in range(10):  # Limit the number of checks to avoid infinite loops
            msg_status = client.messages(message_sid).fetch().status
            if msg_status == "delivered":
                logging.info(f"Message delivered to {customer_name} at {customer_phone}")
                return True
            elif msg_status in ["failed", "undelivered"]:
                logging.error(f"Message failed to {customer_name} at {customer_phone}")
                return False
            time.sleep(2)  # Wait before checking again
        
        # If the status is still unknown after all checks
        logging.warning(f"Message delivery status for {customer_name} at {customer_phone} is unknown")
        return False
    
    except Exception as e:
        logging.error(f"Failed to check status for message to {customer_name} at {customer_phone}: {str(e)}")
        return False

# Function to send SMS messages
def send_sms():
    global customer_data, opted_out_numbers
    if customer_data is None:
        messagebox.showerror("Error", "Please load customer data first.")
        return
    
    account_sid = sid_entry.get()
    auth_token = token_entry.get()
    twilio_phone_number = phone_entry.get()
    
    if not account_sid or not auth_token or not twilio_phone_number:
        messagebox.showerror("Error", "Please enter all Twilio credentials.")
        return
    
    message_template = message_entry.get("1.0", "end-1c")
    if not message_template:
        messagebox.showerror("Error", "Please enter the message text.")
        return
    
    try:
        client = Client(account_sid, auth_token)
        client.api.accounts(sid=account_sid).fetch()
    except Exception as e:
        error_message = (
            "Failed to authenticate Twilio credentials.\n"
            "Please check your Account SID, Auth Token, and Twilio Phone Number."
        )
        messagebox.showerror("Authentication Error", error_message)
        logging.error(f"Failed to authenticate Twilio credentials: {str(e)}")
        send_button.configure(text="Sending Failed", fg_color="red")
        return
    
    start_time = datetime.datetime.now()
    logging.info(f"SMS sending started at {start_time}")
    
    progress["maximum"] = len(customer_data)
    progress["value"] = 0  # Reset progress bar
    
    message_sids = []  # Store message SIDs for later status checking
    
    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = []
        for index, row in customer_data.iterrows():
            customer_name = row['Name']
            customer_phone = str(row['Phone Number'])
            
            if not is_valid_phone_number(customer_phone):
                logging.warning(f"Invalid phone number skipped: {customer_phone}")
                continue
            
            if customer_phone in opted_out_numbers:
                logging.info(f"Skipping opted-out number: {customer_phone}")
                continue
            
            message_body = message_template.replace("{Name}", customer_name)
            
            futures.append(executor.submit(send_single_sms, client, twilio_phone_number, customer_name, customer_phone, message_body))
        
        for future in as_completed(futures):
            message_sid = future.result()
            if message_sid:
                message_sids.append((message_sid, row['Name'], row['Phone Number']))
            progress["value"] += 1
            root.update_idletasks()
    
    # Check the statuses after all messages have been sent
    for message_sid, customer_name, customer_phone in message_sids:
        status = check_message_status(client, message_sid, customer_name, customer_phone)
        if status:
            print("Message sent")
        else:
            print("Message haven't sent")
    
    end_time = datetime.datetime.now()
    logging.info(f"SMS sending finished at {end_time}. Duration: {end_time - start_time}")
    
    messagebox.showinfo("Success", "SMS messages processed!")
    send_button.configure(text="Messages Processed", fg_color="green")

# Set up the GUI
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.geometry("500x600")
root.title("SMS Sender")

# Ensure the window pops up on the current tab and in the foreground
root.attributes('-topmost', True)
root.after(100, lambda: root.attributes('-topmost', False))

label = ctk.CTkLabel(root, text="Automated SMS Sender", font=ctk.CTkFont(size=24, weight="bold"))
label.pack(pady=20)

sid_label = ctk.CTkLabel(root, text="Twilio Account SID:", font=ctk.CTkFont(size=16))
sid_label.pack(pady=10)
sid_entry = ctk.CTkEntry(root, width=400)
sid_entry.pack(pady=5)

token_label = ctk.CTkLabel(root, text="Twilio Auth Token:", font=ctk.CTkFont(size=16))
token_label.pack(pady=10)
token_entry = ctk.CTkEntry(root, width=400, show="*")
token_entry.pack(pady=5)

phone_label = ctk.CTkLabel(root, text="Twilio Phone Number:", font=ctk.CTkFont(size=16))
phone_label.pack(pady=10)
phone_entry = ctk.CTkEntry(root, width=400)
phone_entry.pack(pady=5)

message_label = ctk.CTkLabel(root, text="Enter Message Text (use {Name} for customer name):", font=ctk.CTkFont(size=16))
message_label.pack(pady=10)

# Create the paragraph box using CTkTextbox
message_entry = ctk.CTkTextbox(root, width=400, height=150, wrap="word")  # Enable word wrapping
message_entry.pack(pady=10)

# Separate buttons for loading data and sending SMS
load_button = ctk.CTkButton(root, text="Load Customer Data", command=load_customer_data, font=ctk.CTkFont(size=16, weight="bold"))
load_button.pack(pady=10)

send_button = ctk.CTkButton(root, text="Send SMS Messages", command=send_sms, font=ctk.CTkFont(size=16, weight="bold"))
send_button.pack(pady=20)

# Add progress bar to track message sending progress
progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress.pack(pady=10)

# Start the GUI event loop
root.mainloop()
