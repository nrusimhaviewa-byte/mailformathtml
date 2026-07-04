import sys
import os
import argparse

try:
    import win32com.client
except ImportError:
    print("Error: pywin32 is not installed.")
    print("Please install it using: pip install pywin32")
    sys.exit(1)

try:
    from premailer import transform
    HAS_PREMAILER = True
except ImportError:
    HAS_PREMAILER = False
    print("Warning: premailer is not installed. CSS will not be automatically inlined.")
    print("For best results in Outlook, install it using: pip install premailer\n")

def create_outlook_email(html_content, subject="HTML Email", to=""):
    try:
        # Initialize Outlook COM object
        outlook = win32com.client.Dispatch('outlook.application')
        mail = outlook.CreateItem(0) # 0 corresponds to a MailItem
        
        mail.Subject = subject
        if to:
            mail.To = to
            
        mail.HTMLBody = html_content
        
        # Display the email for user review before sending
        mail.Display()
        print("Outlook email draft created successfully.")
        
    except Exception as e:
        print(f"Failed to create Outlook email: {e}")

def main():
    parser = argparse.ArgumentParser(description="Convert an HTML file into an Outlook draft email.")
    parser.add_argument("html_file", help="Path to the source HTML file")
    parser.add_argument("-s", "--subject", default="Converted HTML Email", help="Subject of the email")
    parser.add_argument("-t", "--to", default="", help="Recipient email address")
    parser.add_argument("-o", "--output", default="", help="Save the fully rendered HTML to this output file")
    parser.add_argument("--no-inline", action="store_true", help="Skip CSS inlining even if premailer is installed")
    parser.add_argument("--no-outlook", action="store_true", help="Skip creating the Outlook draft (useful if you only want to output the file)")

    args = parser.parse_args()

    if not os.path.exists(args.html_file):
        print(f"Error: File '{args.html_file}' not found.")
        sys.exit(1)

    try:
        with open(args.html_file, 'r', encoding='utf-8') as f:
            html_content = f.read()
    except Exception as e:
        print(f"Error reading file: {e}")
        sys.exit(1)

    # Inline CSS styles for better Outlook compatibility
    if HAS_PREMAILER and not args.no_inline:
        print("Inlining CSS styles using premailer...")
        try:
            import logging
            import cssutils
            cssutils.log.setLevel(logging.CRITICAL)
            html_content = transform(html_content)
        except Exception as e:
            print(f"Error during CSS inlining: {e}")
            print("Proceeding with original HTML...")

    # Save output if requested
    if args.output:
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(html_content)
            print(f"Successfully saved fully rendered HTML to: {args.output}")
        except Exception as e:
            print(f"Failed to save output file: {e}")

    # Create the email unless explicitly skipped
    if not args.no_outlook:
        create_outlook_email(html_content, subject=args.subject, to=args.to)

if __name__ == "__main__":
    main()