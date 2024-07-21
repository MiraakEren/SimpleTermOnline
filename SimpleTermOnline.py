import os
import sys
import csv
import subprocess
import datetime
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog, PhotoImage
import pyperclip
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from PIL import Image, ImageTk

CONFIG_FILE = 'config.json'



class SimpleTermOnline:
    def __init__(self, root):
        self.root = root
        self.setup_icon()
        self.sheet_id = None
        self.json_keyfile_path = None
        self.sheet = None
        self.df = None
        self.results = []
        self.current_index = 0
        self.current_search_term = ""
        self.username = None

        self.load_config()
        self.authenticate_google_sheets()
        self.load_sheet()
        self.setup_gui()

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r') as file:
                config = json.load(file)
                self.sheet_id = config.get('sheet_id')
                self.json_keyfile_path = config.get('json_keyfile_path')
                self.username = config.get('username')

    def save_config(self):
        config = {
            'sheet_id': self.sheet_id,
            'json_keyfile_path': self.json_keyfile_path,
            'username': self.username
        }
        with open(CONFIG_FILE, 'w') as file:
            json.dump(config, file)

    def authenticate_google_sheets(self):
        try:
            if not self.sheet_id:
                self.sheet_id = simpledialog.askstring("Google Sheets ID", "Enter the ID between 'd/' and the next '/' in the URL:")
                if not self.sheet_id:
                    raise Exception("No Google Sheets ID entered")

            if not self.json_keyfile_path:
                self.json_keyfile_path = filedialog.askopenfilename(
                    title="Select Google Sheets API Key File",
                    filetypes=[("JSON files", "*.json")]
                )
                if not self.json_keyfile_path:
                    raise Exception("No key file selected")

            if not self.username:
                self.username = simpledialog.askstring("Username", "Please enter a username")
                if not self.username:
                    raise Exception("No username entered")

            self.save_config()

            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name(self.json_keyfile_path, scope)
            client = gspread.authorize(creds)
            self.sheet = client.open_by_key(self.sheet_id).sheet1
        except Exception as e:
            messagebox.showerror("Error", f"Error authenticating with Google Sheets:\n{str(e)}")
            sys.exit()

    def load_sheet(self):
        try:
            records = self.sheet.get_all_records()
            self.df = pd.DataFrame(records)
            
            # Convert 'Reviewed' column to boolean
            if 'Reviewed' in self.df.columns:
                self.df['Reviewed'] = self.df['Reviewed'].apply(lambda x: str(x).strip().lower() == 'true')
            
        except Exception as e:
            self.df = None
            messagebox.showerror("Error", f"Error loading Google Sheet:\n{str(e)}")
            sys.exit()

    def refresh_sheet(self, event=None):
        try:
            self.df = None
            self.load_sheet()
            if self.df is not None:
                self.results = []
                self.update_display()
                self.root.title("Updated!")
                self.root.after(2000, lambda: self.root.title("SimpleTerm Online"))
            else:
                messagebox.showwarning("Update Error", "Failed to refresh data from Google Sheets.")
        except Exception as e:
            messagebox.showerror("Error", f"Error refreshing data from Google Sheets:\n{str(e)}")

    def find_equivalent(self, term):
        results = []
        if self.df is not None:
            for index, row in self.df.iterrows():
                if row['Source Term'].lower() == term.lower():
                    notes = row['Notes']
                    reviewed = row.get('Reviewed', False)
                    if pd.isna(notes):
                        notes = ""
                    results.append({
                        'target_term': row['Target Term'],
                        'notes': notes,
                        'reviewed': reviewed
                    })
        return results

    def search_term(self, event=None):
        term = self.entry.get().strip()
        if term and self.df is not None:
            try:
                self.results = self.find_equivalent(term)
                self.current_index = 0
                self.update_display()
                self.current_search_term = term
            except Exception as e:
                print(f"Error occurred: {str(e)}")
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
        else:
            messagebox.showwarning("Input Error", "Please enter a term to search or load the Google Sheet first.")

    def update_display(self):
        if self.results:
            result = self.results[self.current_index]
            self.result_label.config(text=result['target_term'])

            notes_text = result['notes'] if result['notes'] else ""
            self.notes_text.config(state=tk.NORMAL)
            self.notes_text.delete(1.0, tk.END)
            self.notes_text.insert(tk.END, notes_text)
            self.notes_text.config(state=tk.DISABLED)

            if result['reviewed']:
                try:
                    icon_path = self.resource_path('checkmark.png')
                    icon = Image.open(icon_path)
                    icon = icon.resize((16, 16))
                    self.reviewed_icon = ImageTk.PhotoImage(icon)
                    self.result_label.config(image=self.reviewed_icon, compound=tk.RIGHT)
                except Exception as e:
                    print(f"Error loading icon: {str(e)}")
            else:
                self.result_label.config(image='', compound=tk.NONE)

            self.result_label.config(fg='blue' if len(self.results) > 1 else 'black')
        else:
            self.result_label.config(text='Term not found.', image='', compound=tk.NONE)
            self.notes_text.config(state=tk.NORMAL)
            self.notes_text.delete(1.0, tk.END)
            self.notes_text.config(state=tk.DISABLED)

        self.root.update_idletasks()
        self.root.geometry(f"{self.root.winfo_width()}x{self.root.winfo_height()}")

    def download_sheet(self, event=None):
        if self.sheet is None:
            messagebox.showwarning("Download Error", "Google Sheet is not loaded.")
            return

        try:
            data = self.sheet.get_all_records()
            now = datetime.datetime.now()
            timestamp = now.strftime("%Y%m%d_%H%M%S")

            filename = f'downloaded_sheet_{timestamp}.csv'

            with open(filename, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(self.df.columns)
                for row in data:
                    writer.writerow(row.values())

            messagebox.showinfo("Download Success", f"Sheet downloaded successfully as {filename}.")
        except Exception as e:
            messagebox.showerror("Error", f"Error downloading Google Sheet:\n{str(e)}")

    def setup_gui(self):
        if self.sheet_id:
            self.root.title(f"SimpleTerm Online")

        self.root.attributes('-topmost', True)
        self.root.configure(bg='#f0f0f0')

        input_frame = tk.Frame(self.root, padx=10, pady=10, bg='#f0f0f0')
        input_frame.pack(padx=10, pady=10, fill=tk.X)
        input_frame.pack_propagate(False)

        self.entry = tk.Entry(input_frame, font=('Arial', 14), relief=tk.FLAT, width=20)
        self.entry.grid(row=0, column=0, padx=(0, 10), sticky='ew')

        self.result_label = tk.Label(input_frame, text='', font=('Arial', 14), anchor='w', bg='#f0f0f0')
        self.result_label.grid(row=0, column=1, padx=(10, 0), sticky='ew')

        input_frame.grid_columnconfigure(0, weight=0)

        result_frame = tk.Frame(self.root, padx=10, pady=0, bg='#f0f0f0')
        result_frame.pack(padx=10, pady=0, fill=tk.BOTH, expand=True)
        self.notes_text = tk.Text(result_frame, wrap=tk.WORD, font=('Arial', 12), state=tk.DISABLED, bg='#f0f0f0', relief=tk.FLAT)
        self.notes_text.pack(fill=tk.BOTH, expand=True)
        result_frame.pack_propagate(False)

        self.entry.bind('<Return>', self.search_term)
        self.root.bind('<Control-n>', lambda event: self.open_add_term_dialog())
        self.root.bind('<Control-o>', lambda event: self.open_google_sheet())
        self.root.bind('<Right>', self.navigate_results)
        self.root.bind('<Left>', self.navigate_results)
        self.root.bind('<Tab>', self.navigate_results)
        self.root.bind('<Control-c>', self.copy_result_term)
        self.root.bind('<Control-plus>', self.increase_font_size)
        self.root.bind('<Control-minus>', self.decrease_font_size)
        self.root.bind('<Control-Shift-plus>', self.increase_notes_font_size)
        self.root.bind('<Control-Shift-minus>', self.decrease_notes_font_size)
        self.root.bind('<F5>', self.refresh_google_sheet)
        self.root.bind('<F3>', self.display_help)
        self.root.bind('<F2>', self.change_sheet_config)
        self.root.bind('<Control-d>', self.download_sheet)
        self.root.bind('<Control-e>', self.open_edit_entry_dialog)
        self.root.bind('<Control-r>', self.open_reviewer_mode)

        self.entry.focus_set()
        self.root.after(100, self.entry.focus_force)

                
    def open_google_sheet(self, event=None):
        url = f"https://docs.google.com/spreadsheets/d/{self.sheet_id}/edit"
        if sys.platform.startswith('win'):
            os.startfile(url)
        elif sys.platform.startswith('darwin'):
            subprocess.call(['open', url])
        else:
            subprocess.call(['xdg-open', url])
    
    def change_sheet_config(self, event=None):
        def save_changes():
            new_sheet_id = sheet_id_entry.get().strip()
            new_keyfile_path = keyfile_path_entry.get().strip()

            if new_sheet_id and new_keyfile_path:
                self.sheet_id = new_sheet_id
                self.json_keyfile_path = new_keyfile_path
                self.save_config()  # Save the new configuration
                self.authenticate_google_sheets()  # Re-authenticate with the new credentials
                messagebox.showinfo("Success", "Configuration updated successfully.")
                config_dialog.destroy()
            else:
                messagebox.showwarning("Missing Fields", "Please enter both Google Sheet ID and JSON key file path.")

        config_dialog = tk.Toplevel(self.root)
        config_dialog.title("Change Google Sheet Config")

        dialog_width = 450
        dialog_height = 100

        screen_width = config_dialog.winfo_screenwidth()
        screen_height = config_dialog.winfo_screenheight()

        x = (screen_width // 2) - (dialog_width // 2)
        y = (screen_height // 2) - (dialog_height // 2)

        config_dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        tk.Label(config_dialog, text="Google Sheet ID:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        sheet_id_entry = tk.Entry(config_dialog, width=50, relief=tk.FLAT)
        sheet_id_entry.grid(row=0, column=1, padx=10, pady=5)
        sheet_id_entry.insert(0, self.sheet_id if self.sheet_id else "")

        tk.Label(config_dialog, text="JSON Key File Path:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        keyfile_path_entry = tk.Entry(config_dialog, width=50, relief=tk.FLAT)
        keyfile_path_entry.grid(row=1, column=1, padx=10, pady=5)
        keyfile_path_entry.insert(0, self.json_keyfile_path if self.json_keyfile_path else "")

        tk.Button(config_dialog, text="Save", command=save_changes).grid(row=2, column=1, padx=10, pady=10, sticky='e')

        config_dialog.bind('<Return>', lambda event: save_changes())
        config_dialog.bind('<Escape>', lambda event: config_dialog.destroy())

        config_dialog.transient(self.root)
        config_dialog.grab_set()
        config_dialog.focus_set()
        sheet_id_entry.focus()

    def open_edit_entry_dialog(self, event=None):
        if not self.results:
            messagebox.showwarning("No Results", "No results available to edit.")
            return

        result = self.results[self.current_index]
        current_source_term = self.current_search_term
        current_target_term = result['target_term']
        current_notes = result['notes']
        user_info = self.username

        
        def save_changes():
            new_source_term = source_entry.get().strip()
            new_target_term = target_entry.get().strip()
            new_notes = notes_entry.get().strip()
            user_info = self.username

            if new_source_term and new_target_term:
                try:
                    # Find the row index for the current entry
                    index = self.df[(self.df['Source Term'] == current_source_term) &
                                    (self.df['Target Term'] == current_target_term)].index[0]

                    # Update the DataFrame
                    self.df.at[index, 'Source Term'] = new_source_term
                    self.df.at[index, 'Target Term'] = new_target_term
                    self.df.at[index, 'Notes'] = new_notes
                    self.df.at[index, 'User Info'] = user_info

                    # Update the Google Sheet
                    self.sheet.update_cell(index + 2, 1, new_source_term)  # Column 1 for 'Source Term'
                    self.sheet.update_cell(index + 2, 2, new_target_term)  # Column 2 for 'Target Term'
                    self.sheet.update_cell(index + 2, 3, new_notes)  # Column 3 for 'Notes'
                    self.sheet.update_cell(index + 2, 4, user_info)

                    messagebox.showinfo("Success", "Entry updated successfully.")
                    edit_dialog.destroy()
                    self.refresh_sheet()  # Refresh the display to show updated data
                except Exception as e:
                    messagebox.showerror("Error", f"Error updating Google Sheet:\n{str(e)}")
            else:
                messagebox.showwarning("Missing Fields", "Please enter both Source Term and Target Term.")

        edit_dialog = tk.Toplevel(self.root)
        edit_dialog.title("Edit Entry")

        dialog_width = 400
        dialog_height = 100

        screen_width = edit_dialog.winfo_screenwidth()
        screen_height = edit_dialog.winfo_screenheight()

        x = (screen_width // 2) - (dialog_width // 2)
        y = (screen_height // 2) - (dialog_height // 2)

        edit_dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        tk.Label(edit_dialog, text="Source Term:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        source_entry = tk.Entry(edit_dialog, width=40, relief=tk.FLAT)
        source_entry.grid(row=0, column=1, padx=10, pady=5)
        source_entry.insert(0, current_source_term)

        tk.Label(edit_dialog, text="Target Term:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        target_entry = tk.Entry(edit_dialog, width=40, relief=tk.FLAT)
        target_entry.grid(row=1, column=1, padx=10, pady=5)
        target_entry.insert(0, current_target_term)

        tk.Label(edit_dialog, text="Notes:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
        notes_entry = tk.Entry(edit_dialog, width=40, relief=tk.FLAT)
        notes_entry.grid(row=2, column=1, padx=10, pady=5)
        notes_entry.insert(0, current_notes)

        edit_dialog.bind('<Return>', lambda event: save_changes())
        edit_dialog.bind('<Escape>', lambda event: edit_dialog.destroy())

        edit_dialog.transient(self.root)
        edit_dialog.grab_set()
        edit_dialog.focus_set()
        source_entry.focus()

    def navigate_results(self, event):
        if self.results:
            if event.keysym == 'Right' or event.keysym == 'Tab':
                self.current_index = (self.current_index + 1) % len(self.results)
            elif event.keysym == 'Left':
                self.current_index = (self.current_index - 1) % len(self.results)
            self.update_display()

    def copy_result_term(self, event):
        if self.results:
            pyperclip.copy(self.results[self.current_index]['target_term'])
            self.result_label.config(bg='green')
            self.root.after(500, lambda: self.result_label.config(bg=self.root.cget('bg')))

    def increase_font_size(self, event=None):
        current_font_size = int(self.entry.cget('font').split()[1])
        new_font_size = current_font_size + 1
        self.entry.config(font=('Arial', new_font_size))
        self.result_label.config(font=('Arial', new_font_size))

    def decrease_font_size(self, event=None):
        current_font_size = int(self.entry.cget('font').split()[1])
        new_font_size = max(current_font_size - 1, 8)
        self.entry.config(font=('Arial', new_font_size))
        self.result_label.config(font=('Arial', new_font_size))

    def increase_notes_font_size(self, event=None):
        current_font_size = int(self.notes_text.cget('font').split()[1])
        new_font_size = current_font_size + 1
        self.notes_text.config(font=('Arial', new_font_size))

    def decrease_notes_font_size(self, event=None):
        current_font_size = int(self.notes_text.cget('font').split()[1])
        new_font_size = max(current_font_size - 1, 8)
        self.notes_text.config(font=('Arial', new_font_size))

    def refresh_google_sheet(self, event=None):
        self.refresh_sheet()

    def display_help(self, event=None):
        help_text = (
            "Shortcuts\n"
            "Tab Key, Left and Right arrows: Navigate between results\n"
            "Ctrl+C: Copy the displayed result\n"
            "Ctrl+N: Add a new term\n"
            "Ctrl+O: Open the used Google Sheet\n"
            "Ctrl+E: Edit the searched entry"
            "Ctrl+D: Download the sheet as a .csv"
            "Ctrl+R Open the reviewer mode"
            "Ctrl++/-: Increase or decrease the font size\n"
            "Ctrl+Shift++/-: Increase or decrease the notes font size\n"
            "F5: Refresh the Google Sheet\n"
            "F3: Display this help\n"
            "F2: Change the sheet being used"
        )
        messagebox.showinfo("Help", help_text)

    def open_add_term_dialog(self, event=None):
        def save_term():
            source_term = source_entry.get().strip()
            target_term = target_entry.get().strip()
            notes = notes_entry.get().strip()

            if source_term and target_term:
                new_data = {
                    'Source Term': source_term,
                    'Target Term': target_term,
                    'Notes': notes if notes else '',
                    'Username': self.username,
                    'Reviewed': False
                }

                try:
                    self.sheet.append_row([source_term, target_term, new_data['Notes'], new_data['Username'], new_data['Reviewed']])
                    self.root.title("Updated!")
                    self.root.after(2000, lambda: self.root.title("SimpleTerm Online"))
                    new_term_dialog.destroy()
                except Exception as e:
                    messagebox.showerror("Error", f"Error saving to Google Sheet:\n{str(e)}")
            else:
                messagebox.showwarning("Missing Fields", "Please enter Source Term and Target Term.")

        new_term_dialog = tk.Toplevel(self.root)
        new_term_dialog.title("Add New Term")

        dialog_width = 400
        dialog_height = 100

        screen_width = new_term_dialog.winfo_screenwidth()
        screen_height = new_term_dialog.winfo_screenheight()

        x = (screen_width // 2) - (dialog_width // 2)
        y = (screen_height // 2) - (dialog_height // 2)

        new_term_dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        tk.Label(new_term_dialog, text="Source Term:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
        source_entry = tk.Entry(new_term_dialog, width=40, relief=tk.FLAT)
        source_entry.grid(row=0, column=1, padx=10, pady=5)

        tk.Label(new_term_dialog, text="Target Term:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
        target_entry = tk.Entry(new_term_dialog, width=40, relief=tk.FLAT)
        target_entry.grid(row=1, column=1, padx=10, pady=5)

        tk.Label(new_term_dialog, text="Notes:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
        notes_entry = tk.Entry(new_term_dialog, width=40, relief=tk.FLAT)
        notes_entry.grid(row=2, column=1, padx=10, pady=5)

        new_term_dialog.bind('<Return>', lambda event: [save_term(), self.refresh_sheet()])
        new_term_dialog.bind('<Escape>', lambda event: new_term_dialog.destroy())

        source_entry.insert(0, self.current_search_term)

        new_term_dialog.transient(self.root)
        new_term_dialog.grab_set()
        new_term_dialog.focus_set()
        source_entry.focus()


    def open_reviewer_mode(self, event=None):
        # Create a new window for Reviewer Mode
        reviewer_window = tk.Toplevel()
        reviewer_window.title("Reviewer Mode")

        # Create a frame to hold the canvas and scrollbar
        frame = tk.Frame(reviewer_window)
        frame.pack(fill=tk.BOTH, expand=True)

        # Create a canvas for scrollable content
        canvas = tk.Canvas(frame)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Add a scrollbar
        scrollbar = tk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Create a frame to contain the content within the canvas
        content_frame = tk.Frame(canvas)
        
        # Add content_frame to the canvas
        canvas.create_window((0, 0), window=content_frame, anchor=tk.NW)

        # Configure scrollbar and canvas
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.configure(yscrollcommand=scrollbar.set)

        # Bind mousewheel scrolling
        def on_mousewheel(event):
            # Scroll up or down
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)  # Windows and macOS

        # Filter out reviewed entries
        self.check_vars = []
        self.entries = []

        # Get unreviewed entries
        unreviewed_entries = self.df[self.df['Reviewed'] == False]

        for index, row in unreviewed_entries.iterrows():
            var = tk.BooleanVar()
            chk = tk.Checkbutton(content_frame, text=row['Source Term'] + ' - ' + row['Target Term'], variable=var)
            chk.grid(row=index, column=0, sticky='w')
            self.check_vars.append(var)
            self.entries.append(index)  # Use row index for DataFrame

        # Button to confirm changes
        confirm_button = tk.Button(reviewer_window, text="Update Selected", command=lambda: self.update_selected(reviewer_window))
        confirm_button.pack(pady=10)

    def update_selected(self, reviewer_window):
        # Update the DataFrame and Google Sheet
        for var, index in zip(self.check_vars, self.entries):
            if var.get():  # If checkbox is checked
                self.df.at[index, 'Reviewed'] = True
                self.df.at[index, 'Reviewer'] = self.username
                # Update Google Sheet
                self.sheet.update_cell(index + 2, 5, 'TRUE')  # 5th column for 'Reviewed'
                self.sheet.update_cell(index + 2, 6, self.username)  # 6th column for 'Reviewer'

        # Close the reviewer window
        reviewer_window.destroy()
        
    def resource_path(self, relative_path):
        if hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)
        return os.path.join(base_path, relative_path)


    def setup_icon(self):
        """ Set the application icon. """
        icon_path = self.resource_path('app_icon.png')
        try:
            icon = PhotoImage(file=icon_path)
            self.root.iconphoto(True, icon)
        except Exception as e:
            print(f"Error loading icon: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleTermOnline(root)
    root.mainloop()