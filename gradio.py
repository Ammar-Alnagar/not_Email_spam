import pandas as pd
import gradio as gr
from anthropic import Anthropic
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import os
import io

def read_excel_data(file_obj):
    """
    Read data from uploaded Excel file
    """
    try:
        df = pd.read_excel(io.BytesIO(file_obj.read()))
        return df
    except Exception as e:
        return f"Error reading Excel file: {e}"

def generate_email_content(recipient_data, anthropic_api_key):
    """
    Generate a highly personalized sales proposal email
    """
    try:
        # Initialize Anthropic client
        client = Anthropic(api_key=anthropic_api_key)
        
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
        """
        
        response = client.messages.create(
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
        email_content += f"\n\n---\nBest regards,\n\nName: Critical Future Ai Mailer\nPosition: Expert salesAI"
        
        return email_content
    except Exception as e:
        return f"Error generating email content: {e}"

def send_test_email(row, smtp_email, smtp_password, smtp_server, smtp_port):
    """
    Send a test email and return status
    """
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = smtp_email
        msg['To'] = row['email']
        msg['Subject'] = f"Exclusive AI Solution for {row['company']} - Transformative Opportunity"

        # Generate email content
        email_body = generate_email_content(row, os.getenv('ANTHROPIC_API_KEY'))
        msg.attach(MIMEText(email_body, 'plain'))

        # Create SMTP session
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_email, smtp_password)
        
        # Send email
        server.send_message(msg)
        server.quit()
        return "Email sent successfully"
    except Exception as e:
        return f"Error sending email: {e}"

def process_emails(excel_file, smtp_email, smtp_password, smtp_server, smtp_port):
    """
    Process emails from the uploaded Excel file
    """
    # Read the Excel file
    df = read_excel_data(excel_file)
    
    if isinstance(df, str):  # Error occurred
        return df
    
    # Validate required columns
    required_columns = ['first_name', 'last_name', 'company', 'position', 'email']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        return f"Missing columns in Excel file: {', '.join(missing_columns)}"
    
    # Process emails with results
    results = []
    for index, row in df.iterrows():
        result = send_test_email(row, smtp_email, smtp_password, smtp_server, smtp_port)
        results.append(f"Row {index + 2}: {result}")
        # Add a delay between emails
        time.sleep(5)
    
    return "\n".join(results)

# Gradio Interface
def create_gradio_interface():
    with gr.Blocks() as demo:
        gr.Markdown("# Critical Future AI Email Automation")
        
        with gr.Row():
            with gr.Column():
                excel_input = gr.File(label="Upload Excel File", type="file", file_types=['.xlsx', '.xls'])
                
                smtp_email = gr.Textbox(label="SMTP Email", type="text")
                smtp_password = gr.Textbox(label="SMTP Password", type="password")
                smtp_server = gr.Textbox(label="SMTP Server", value="smtp.gmail.com")
                smtp_port = gr.Number(label="SMTP Port", value=587)
                
                submit_btn = gr.Button("Send Emails")
                
            with gr.Column():
                output = gr.Textbox(label="Results")
        
        submit_btn.click(
            fn=process_emails, 
            inputs=[excel_input, smtp_email, smtp_password, smtp_server, smtp_port],
            outputs=output
        )
    
    return demo

# Main execution
if __name__ == "__main__":
    # Create Gradio interface
    demo = create_gradio_interface()
    
    # Launch the Gradio app
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860
    )