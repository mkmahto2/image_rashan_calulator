import os
import win32com.client as win32

# Function to send email with PDF attachment using Outlook
def send_outlook_mail(subject, body, recipient, attachment_path):
    # Create an instance of Outlook
    olApp = win32.Dispatch('Outlook.Application')
    olNS = olApp.GetNameSpace('MAPI')

    # Construct the mail item object
    mailItem = olApp.CreateItem(0)  # 0: mail item
    
    mailItem.Subject = subject
    mailItem.BodyFormat = 1  # 1 = Plain text
    mailItem.To = recipient
    
    # Attach the PDF file if it exists
    if os.path.exists(attachment_path):
        mailItem.Attachments.Add(attachment_path)
        print(f"Attachment added: {attachment_path}")
    else:
        print(f"Attachment not found: {attachment_path}")
    
    mailItem.Body = body
    mailItem.Display()  # Display the email (optional)
    
    # Save and send the email
    mailItem.Save()  # Save a copy of the mail in drafts (optional)
    mailItem.Send()  # Send the email
    print(f"Email sent to {recipient} successfully")

# Example usage
if __name__ == "__main__":
    recipient = 'Mukeshmh001@gmail.com'  # Replace with actual recipient email
    subject = 'Your Subject Here'
    body = 'Please find the attached PDF.'
    attachment_path = 'product_bill.pdf'  # Path to your PDF file

    send_outlook_mail(subject, body, recipient, attachment_path)
