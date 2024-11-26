import pandas as pd
from anthropic import Anthropic
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

def read_excel_data(file_path):
    """
    Read data from Excel file
    """
    try:
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def generate_email_content(recipient_data, anthropic_client):
    """
    Generate a highly personalized sales proposal email
    """
    prompt = f"""
    Craft a compelling, personalized sales proposal email with the following specifications:

    Recipient Details:
    - Name: {recipient_data['first_name']} {recipient_data['last_name']}
    - Company: {recipient_data['company']}
    - Position: {recipient_data['position']}

    Email Objectives:
    1. Create a highly personalized opening that demonstrates research into the recipient's company
    2. Highlight a specific pain point or challenge relevant to their industry
    3. Introduce our AI solution with a clear value proposition
    4. Include a compelling call-to-action for a meeting or demo
    5. Demonstrate immediate potential value and ROI

    Tone and Style:
    - Professional yet conversational
    - Show deep understanding of their business challenges
    - Use a consultative approach
    - Create a sense of urgency and opportunity
    - Make the email feel like a tailored solution, not a generic pitch

    Key Requirements:
    - Mention specific details about their company or industry
    - Quantify potential benefits (e.g., cost savings, efficiency improvements)
    - Sound confident but not arrogant
    - Create intrigue and desire for further discussion

    Email Signature:
    Name: Critical Future Ai Mailer
    Position: Expert salesAI

    Format:
    - Short, impactful paragraphs
    - Clear, concise language
    - Focus on recipient's potential gains
    """
    
    try:
        response = anthropic_client.messages.create(
            model="claude-3-opus-20240229",
            max_tokens=1000,
            temperature=0.7,
            system="You are a top-tier sales AI specialist crafting highly personalized, impactful sales proposal emails. Your goal is to convert recipients into interested prospects through precise, value-driven communication.",
            messages=[{
                "role": "user",
                "content": prompt
            }]
        )
        
        # Extract the text content from the response
        email_content = response.content[0].text
        
        # Append signature
        email_content += f"\n\n---\nBest regards,\n\nCritical Future Ai Mailer\nExpert salesAI"
        
        return email_content
    except Exception as e:
        print(f"Error generating email content: {e}")
        return None

def send_email(recipient_email, subject, body):
    """
    Send email using SMTP
    """
    # Email configuration
    sender_email = os.getenv('SENDER_EMAIL')
    sender_password = os.getenv('SENDER_PASSWORD')
    smtp_server = os.getenv('SMTP_SERVER')
    smtp_port = int(os.getenv('SMTP_PORT'))

    # Create message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    try:
        # Create SMTP session
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        
        # Send email
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False

def main():
    # Initialize Anthropic client
    anthropic_client = Anthropic(api_key=os.getenv('ANTHROPIC_API_KEY'))
    
    # Read Excel file
    excel_file = "recipients.xlsx"
    df = read_excel_data(excel_file)
    
    if df is None:
        return
    
    # Process each row
    for index, row in df.iterrows():
        # Generate email content
        email_content = generate_email_content(row, anthropic_client)
        
        if email_content:
            # Send email
            success = send_email(
                recipient_email=row['email'],
                subject=f"Exclusive AI Solution for {row['company']} - Transformative Opportunity",
                body=email_content
            )
            
            if success:
                print(f"Email sent successfully to {row['email']}")
            else:
                print(f"Failed to send email to {row['email']}")
                
            # Add delay to avoid hitting rate limits
            time.sleep(5)  # Increased delay for more cautious sending

if __name__ == "__main__":
    main()
