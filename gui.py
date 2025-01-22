import json
import tkinter as tk
from tkinter import messagebox, simpledialog

# File to save the contacts
CONTACTS_FILE = "contacts.json"

# Load contacts from file
def load_contacts():
    try:
        with open(CONTACTS_FILE, "r") as file:
            contacts = json.load(file)
    except (FileNotFoundError, json.JSONDecodeError):
        contacts = {}
    return contacts

# Save contacts to file
def save_contacts(contacts):
    with open(CONTACTS_FILE, "w") as file:
        json.dump(contacts, file)

# Add a new contact
def add_contact():
    name = simpledialog.askstring("Add Contact", "Enter contact name:")
    if not name:
        return
    if name in contacts:
        messagebox.showerror("Error", f"Contact '{name}' already exists.")
        return
    phone = simpledialog.askstring("Add Contact", "Enter phone number:")
    email = simpledialog.askstring("Add Contact", "Enter email:")
    contacts[name] = {"phone": phone, "email": email}
    messagebox.showinfo("Success", f"Contact '{name}' added successfully.")
    update_contacts_list()

# Delete a contact
def delete_contact():
    name = simpledialog.askstring("Delete Contact", "Enter the name of the contact to delete:")
    if name in contacts:
        del contacts[name]
        messagebox.showinfo("Success", f"Contact '{name}' deleted successfully.")
        update_contacts_list()
    else:
        messagebox.showerror("Error", f"Contact '{name}' not found.")

# Search for a contact
def search_contact():
    name = simpledialog.askstring("Search Contact", "Enter the name of the contact to search:")
    if name in contacts:
        contact_info = contacts[name]
        messagebox.showinfo("Contact Found", f"Name: {name}\nPhone: {contact_info['phone']}\nEmail: {contact_info['email']}")
    else:
        messagebox.showerror("Error", f"Contact '{name}' not found.")

# Update contact information
def update_contact():
    name = simpledialog.askstring("Update Contact", "Enter the name of the contact to update:")
    if name in contacts:
        phone = simpledialog.askstring("Update Contact", f"Enter new phone number (current: {contacts[name]['phone']}):")
        email = simpledialog.askstring("Update Contact", f"Enter new email (current: {contacts[name]['email']}):")
        contacts[name] = {"phone": phone, "email": email}
        messagebox.showinfo("Success", f"Contact '{name}' updated successfully.")
        update_contacts_list()
    else:
        messagebox.showerror("Error", f"Contact '{name}' not found.")

# Display all contacts
def update_contacts_list():
    contacts_list.delete(0, tk.END)
    for name, info in contacts.items():
        contacts_list.insert(tk.END, f"Name: {name}, Phone: {info['phone']}, Email: {info['email']}")

# Save contacts when closing the app
def on_closing():
    save_contacts(contacts)
    root.destroy()

# Main GUI setup
root = tk.Tk()
root.title("Contact Book")

# Contacts list box
contacts_list = tk.Listbox(root, width=50, height=15)
contacts_list.pack(pady=10)

# Buttons
btn_add = tk.Button(root, text="Add Contact", command=add_contact)
btn_add.pack(pady=5)

btn_delete = tk.Button(root, text="Delete Contact", command=delete_contact)
btn_delete.pack(pady=5)

btn_search = tk.Button(root, text="Search Contact", command=search_contact)
btn_search.pack(pady=5)

btn_update = tk.Button(root, text="Update Contact", command=update_contact)
btn_update.pack(pady=5)

# Load existing contacts and display them
contacts = load_contacts()
update_contacts_list()

# Handle window close
root.protocol("WM_DELETE_WINDOW", on_closing)

# Run the main loop
root.mainloop()