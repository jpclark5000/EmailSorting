import os
import extract_msg
from tkinter import filedialog, messagebox
import tkinter as tk
from datetime import datetime

def rename_msg_files_with_date():
    """
    Retrieve all .msg files in a selected folder, extract their sent date/time,
    and rename them with the format: [initial_title]_[YYYYMMDD-HH'h'MM].msg
    """
    
    # Create a root window and hide it
    root = tk.Tk()
    root.withdraw()
    
    # Open directory selector
    folder_path = filedialog.askdirectory(
        title="Select folder containing .msg files"
    )
    
    if not folder_path:
        messagebox.showinfo("Cancelled", "No folder selected. Operation cancelled.")
        root.destroy()
        return
    
    try:
        # Get all .msg files in the selected folder (non-recursive)
        msg_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.msg')]
        
        if not msg_files:
            messagebox.showinfo("No Files", "No .msg files found in the selected folder.")
            root.destroy()
            return
        
        # Dictionary to map original filename with extracted date
        file_date_mapping = {}
        renamed_files = []
        errors = []
        
        for msg_file in msg_files:
            msg = None  # Initialize msg variable
            try:
                file_path = os.path.join(folder_path, msg_file)
                
                # Open and parse the .msg file
                msg = extract_msg.Message(file_path)
                
                # Extract sent date
                sent_date = msg.date
                
                if sent_date is None:
                    # Try alternative date fields if sent date is not available
                    sent_date = getattr(msg, 'receivedTime', None) or getattr(msg, 'creationTime', None)
                
                if sent_date is None:
                    errors.append(f"No date found for {msg_file}")
                    continue
                
                # Format date as YYYYMMDD-HH'h'MM
                if isinstance(sent_date, str):
                    # Parse string date if needed
                    try:
                        sent_date = datetime.fromisoformat(sent_date.replace('Z', '+00:00'))
                    except:
                        # Try other common formats
                        for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%dT%H:%M:%S']:
                            try:
                                sent_date = datetime.strptime(sent_date, fmt)
                                break
                            except:
                                continue
                
                formatted_date = sent_date.strftime("%Y%m%d-%Hh%M")
                
                # Store mapping
                file_date_mapping[msg_file] = formatted_date
                
                # Create new filename
                original_name = os.path.splitext(msg_file)[0]  # Remove .msg extension
                new_filename = f"{formatted_date}_ {original_name}.msg"
                
                # Handle potential duplicate names
                counter = 1
                temp_filename = new_filename
                while os.path.exists(os.path.join(folder_path, temp_filename)):
                    name_part = f"{formatted_date}_ {original_name}_{counter}"
                    temp_filename = f"{name_part}.msg"
                    counter += 1
                new_filename = temp_filename
                
                # Close the message object BEFORE renaming
                if msg:
                    msg.close()
                    msg = None
                
                # Rename the file
                old_path = os.path.join(folder_path, msg_file)
                new_path = os.path.join(folder_path, new_filename)
                
                os.rename(old_path, new_path)
                renamed_files.append(f"{msg_file} → {new_filename}")
                
            except Exception as e:
                error_msg = f"Error processing {msg_file}: {str(e)}"
                errors.append(error_msg)
                print(error_msg)
            finally:
                # Ensure the message object is always closed
                if msg:
                    try:
                        msg.close()
                    except:
                        pass
        
        # Show results
        result_message = f"Processing completed!\n\n"
        result_message += f"Successfully renamed {len(renamed_files)} files:\n"
        
        if renamed_files:
            # Show first 10 renamed files to avoid overwhelming the message box
            display_files = renamed_files[:10]
            result_message += "\n".join(display_files)
            if len(renamed_files) > 10:
                result_message += f"\n... and {len(renamed_files) - 10} more files"
        
        if errors:
            result_message += f"\n\nErrors encountered ({len(errors)}):\n"
            result_message += "\n".join(errors[:5])  # Show first 5 errors
            if len(errors) > 5:
                result_message += f"\n... and {len(errors) - 5} more errors"
        
        messagebox.showinfo("Results", result_message)
        
        # Print detailed mapping to console
        print("\n" + "="*50)
        print("FILE DATE MAPPING:")
        print("="*50)
        for original_file, extracted_date in file_date_mapping.items():
            print(f"{original_file} → Date: {extracted_date}")
        
    except Exception as e:
        error_msg = f"An error occurred: {str(e)}"
        messagebox.showerror("Error", error_msg)
        print(error_msg)
    
    finally:
        root.destroy()


def extract_attachments_from_msg():
    """
    Extract all attachments from .msg files in a selected folder.
    Renames emails with incremental numbers and saves attachments with structured naming.
    Format: [email_number].[attachment_number]_[original_attachment_name]
    """
    
    # Create a root window and hide it
    root = tk.Tk()
    root.withdraw()
    
    # Open directory selector
    folder_path = filedialog.askdirectory(
        title="Select folder containing .msg files"
    )
    
    if not folder_path:
        messagebox.showinfo("Cancelled", "No folder selected. Operation cancelled.")
        root.destroy()
        return
    
    try:
        # Get all .msg files in the selected folder
        msg_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.msg')]
        
        if not msg_files:
            messagebox.showinfo("No Files", "No .msg files found in the selected folder.")
            root.destroy()
            return
        
        # Sort files to ensure consistent ordering
        msg_files.sort()
        
        extracted_attachments = []
        renamed_emails = []
        emails_with_attachments = 0
        total_attachments = 0
        errors = []
        
        # Process each email with incremental numbering
        for email_index, msg_file in enumerate(msg_files, 1):
            msg = None
            try:
                file_path = os.path.join(folder_path, msg_file)
                
                # Open and parse the .msg file
                msg = extract_msg.Message(file_path)
                
                # Create new email filename with incremental number
                original_name = os.path.splitext(msg_file)[0]
                new_email_name = f"{email_index}.0_ {original_name}.msg"
                
                # Handle potential duplicate names for emails
                counter = 1
                temp_email_name = new_email_name
                while os.path.exists(os.path.join(folder_path, temp_email_name)) and temp_email_name != msg_file:
                    name_part = f"{email_index}.0_ {original_name}_{counter}"
                    temp_email_name = f"{name_part}.msg"
                    counter += 1
                new_email_name = temp_email_name
                
                # Process attachments if they exist
                if msg.attachments:
                    emails_with_attachments += 1
                    
                    # Extract each attachment with structured naming
                    for attachment_index, attachment in enumerate(msg.attachments, 1):
                        try:
                            # Get attachment filename
                            if hasattr(attachment, 'longFilename') and attachment.longFilename:
                                original_attachment_name = attachment.longFilename
                            elif hasattr(attachment, 'shortFilename') and attachment.shortFilename:
                                original_attachment_name = attachment.shortFilename
                            else:
                                original_attachment_name = f"attachment_ {attachment_index}"
                            
                            # Ensure filename is safe for filesystem
                            safe_attachment_name = "".join(c for c in original_attachment_name if c.isalnum() or c in (' ', '-', '_', '.')).rstrip()
                            if not safe_attachment_name:
                                safe_attachment_name = f"attachment_ {attachment_index}"
                            
                            # Create structured attachment filename: [email_number].[attachment_number]_[original_name]
                            structured_attachment_name = f"{email_index}.{attachment_index}_ {safe_attachment_name}"
                            
                            # Handle duplicate attachment filenames
                            original_structured_name = structured_attachment_name
                            dup_counter = 1
                            while os.path.exists(os.path.join(folder_path, structured_attachment_name)):
                                name, ext = os.path.splitext(original_structured_name)
                                structured_attachment_name = f"{name}_ dup{dup_counter}{ext}"
                                dup_counter += 1
                            
                            # Save the attachment
                            attachment_path = os.path.join(folder_path, structured_attachment_name)
                            
                            # Save attachment data
                            if hasattr(attachment, 'data') and attachment.data:
                                with open(attachment_path, 'wb') as f:
                                    f.write(attachment.data)
                                
                                extracted_attachments.append(f"Email {email_index} → {structured_attachment_name}")
                                total_attachments += 1
                                print(f"Extracted: {structured_attachment_name} from email {email_index}")
                            else:
                                errors.append(f"No data found for attachment {attachment_index} in email {email_index}")
                        
                        except Exception as e:
                            error_msg = f"Error extracting attachment {attachment_index} from email {email_index}: {str(e)}"
                            errors.append(error_msg)
                            print(error_msg)
                
                # Close the message object BEFORE renaming the email
                if msg:
                    msg.close()
                    msg = None
                
                # Rename the email file (only if the name is different)
                if new_email_name != msg_file:
                    old_email_path = os.path.join(folder_path, msg_file)
                    new_email_path = os.path.join(folder_path, new_email_name)
                    
                    os.rename(old_email_path, new_email_path)
                    renamed_emails.append(f"{msg_file} → {new_email_name}")
                    print(f"Renamed email: {msg_file} → {new_email_name}")
                else:
                    renamed_emails.append(f"{msg_file} (no rename needed)")
                
            except Exception as e:
                error_msg = f"Error processing email {email_index} ({msg_file}): {str(e)}"
                errors.append(error_msg)
                print(error_msg)
            finally:
                # Ensure the message object is always closed
                if msg:
                    try:
                        msg.close()
                    except:
                        pass
        
        # Show results
        result_message = f"Attachment extraction completed!\n\n"
        result_message += f"Processed {len(msg_files)} .msg files\n"
        result_message += f"Renamed {len(renamed_emails)} emails with incremental numbers\n"
        result_message += f"Found {emails_with_attachments} emails with attachments\n"
        result_message += f"Extracted {total_attachments} attachments total\n\n"
        
        if extracted_attachments:
            result_message += "Sample attachment extractions:\n"
            # Show first 10 extractions to avoid overwhelming the message box
            display_extractions = extracted_attachments[:10]
            result_message += "\n".join(display_extractions)
            if len(extracted_attachments) > 10:
                result_message += f"\n... and {len(extracted_attachments) - 10} more attachments"
        
        if errors:
            result_message += f"\n\nErrors encountered ({len(errors)}):\n"
            result_message += "\n".join(errors[:5])  # Show first 5 errors
            if len(errors) > 5:
                result_message += f"\n... and {len(errors) - 5} more errors"
        
        messagebox.showinfo("Extraction Results", result_message)
        
        # Print detailed results to console
        print("\n" + "="*60)
        print("ATTACHMENT EXTRACTION SUMMARY:")
        print("="*60)
        print(f"Total .msg files processed: {len(msg_files)}")
        print(f"Emails renamed: {len(renamed_emails)}")
        print(f"Emails with attachments: {emails_with_attachments}")
        print(f"Total attachments extracted: {total_attachments}")
        
        if renamed_emails:
            print("\nEmail renaming log:")
            for rename in renamed_emails:
                print(f"  {rename}")
        
        if extracted_attachments:
            print("\nAttachment extraction log:")
            for extraction in extracted_attachments:
                print(f"  {extraction}")
        
    except Exception as e:
        error_msg = f"An error occurred: {str(e)}"
        messagebox.showerror("Error", error_msg)
        print(error_msg)
    
    finally:
        root.destroy()


def main_menu():
    """
    Display a modern menu to choose between renaming files or extracting attachments.
    """
    root = tk.Tk()
    root.title("Email Management Tool")
    root.geometry("500x400")
    root.resizable(False, False)
    root.configure(bg='#2c3e50')
    
    # Center the window
    root.eval('tk::PlaceWindow . center')
    
    # Modern color scheme
    bg_color = '#2c3e50'
    accent_color = '#3498db'
    button_color = '#34495e'
    button_hover = '#4a6741'
    text_color = '#ecf0f1'
    button_text = '#ffffff'
    
    # Title frame
    title_frame = tk.Frame(root, bg=bg_color)
    title_frame.pack(pady=30)
    
    # Main title with modern styling
    main_label = tk.Label(
        title_frame, 
        text="📧 Email Sorting Tool", 
        font=("Segoe UI", 24, "bold"),
        fg=text_color,
        bg=bg_color
    )
    main_label.pack()

    # Separator line
    separator = tk.Frame(root, height=2, bg=accent_color)
    separator.pack(fill='x', padx=50, pady=1)
    
    # Buttons frame
    buttons_frame = tk.Frame(root, bg=bg_color)
    buttons_frame.pack(pady=1)
    
    # Modern button style function
    def create_modern_button(parent, text, icon, command, description):
        button_frame = tk.Frame(parent, bg=bg_color)
        button_frame.pack(pady=15, padx=40, fill='x')
        
        # Main button
        btn = tk.Button(
            button_frame,
            text=f"{icon}  {text}",
            command=command,
            font=("Segoe UI", 12, "bold"),
            fg=button_text,
            bg=button_color,
            activebackground=button_hover,
            activeforeground=button_text,
            relief='flat',
            bd=0,
            padx=20,
            pady=15,
            cursor='hand2',
            width=35
        )
        btn.pack(fill='x')
        
        # Description label
        desc_label = tk.Label(
            button_frame,
            text=description,
            font=("Segoe UI", 9),
            fg='#95a5a6',
            bg=bg_color
        )
        desc_label.pack(pady=(5, 0))
        
        # Hover effects
        def on_enter(e):
            btn.configure(bg=accent_color)
        
        def on_leave(e):
            btn.configure(bg=button_color)
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return btn
    
    # Create modern buttons
    rename_btn = create_modern_button(
        buttons_frame,
        "Rename MSG Files with Date",
        "📅",
        lambda: [root.destroy(), rename_msg_files_with_date()],
        "Automatically rename email files with their sent date and time"
    )
    
    extract_btn = create_modern_button(
        buttons_frame,
        "Extract Attachments from MSG Files",
        "📎",
        lambda: [root.destroy(), extract_attachments_from_msg()],
        "Extract and organize all attachments with structured naming"
    )
    
    # Bottom section
    bottom_frame = tk.Frame(root, bg=bg_color)
    bottom_frame.pack(side='bottom', fill='x', pady=20)
    
 
    # Footer info
    footer_label = tk.Label(
        bottom_frame,
        text="v1.0 | Professional Email Management Solution",
        font=("Segoe UI", 8),
        fg='#7f8c8d',
        bg=bg_color
    )
    footer_label.pack(side='bottom', pady=(10, 0))
    
    root.mainloop()


# Example usage
if __name__ == "__main__":
    main_menu()