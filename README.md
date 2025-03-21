# Personalized Sales Email Automation

This project automates the process of sending personalized sales proposal emails using data from an Excel file, content generated by the Anthropic Claude API, and SMTP for email delivery.

## Features

*   **Data-Driven Personalization:** Reads recipient data (name, company, position, email) from an Excel file to create highly personalized email content.
*   **AI-Powered Content Generation:** Utilizes the Anthropic Claude API to generate compelling and tailored sales proposal emails based on recipient information and predefined objectives.
*   **Automated Email Sending:**  Sends emails using SMTP (Simple Mail Transfer Protocol).
*   **Environment Variable Configuration:**  Uses environment variables to store sensitive information such as email credentials and API keys.
*   **Rate Limiting:** Implements a delay between email sends to avoid hitting rate limits.
*   **Error Handling:** Includes error handling for file reading, API calls, and email sending.

## Dependencies

*   pandas
*   anthropic
*   smtplib
*   email
*   python-dotenv
*   os
*   time

## Installation

1.  **Clone the repository:**

    ```bash
    git clone <repository_url> #Replace <repository_url> with url of the repo
    cd <repository_directory> # Navigate into project
    ```

2.  **Install dependencies:**

    ```bash
    pip install pandas anthropic python-dotenv
    ```

## Configuration

1.  **Set up environment variables:**

    Create a `.env` file in the project directory with the following variables:

    ```
    SENDER_EMAIL=your_email@example.com
    SENDER_PASSWORD=your_email_password
    SMTP_SERVER=smtp.example.com
    SMTP_PORT=587  # Or 465 for SSL
    ANTHROPIC_API_KEY=your_anthropic_api_key
    ```

    *   Replace `your_email@example.com` with the email address you will be sending from.
    *   Replace `your_email_password` with the password for that email address.  **Note:**  Consider using an app password if your email provider supports it, for enhanced security.
    *   Replace `smtp.example.com` with your email provider's SMTP server address.
    *   Replace `587` with the appropriate SMTP port for your provider (typically 587 for TLS, or 465 for SSL).
    *   Replace `your_anthropic_api_key` with your Anthropic API key.  Obtain this from the Anthropic website.

2.  **Create the Excel file:**

    Create an Excel file named `recipients.xlsx` in the project directory.  The file should have the following columns:

    *   `first_name`: Recipient's first name.
    *   `last_name`: Recipient's last name.
    *   `company`: Recipient's company name.
    *   `position`: Recipient's job title.
    *   `email`: Recipient's email address.

    Example `recipients.xlsx`:

    | first_name | last_name | company     | position          | email                 |
    | :---------- | :-------- | :---------- | :---------------- | :-------------------- |
    | John       | Doe       | Acme Corp   | CEO               | john.doe@acme.com     |
    | Jane       | Smith     | Beta Inc    | Marketing Manager | jane.smith@beta.com   |

## Usage

1.  **Run the script:**

    ```bash
    python your_script_name.py # Replace your_script_name with the name of your script
    ```

    The script will:

    *   Read the recipient data from `recipients.xlsx`.
    *   For each recipient, generate a personalized sales proposal email using the Anthropic Claude API.
    *   Send the email using SMTP.
    *   Print success or failure messages to the console.
    *   Pause for 5 seconds between each email to avoid rate limiting.

## Code Explanation

*   **`read_excel_data(file_path)`:** Reads data from an Excel file into a pandas DataFrame.
*   **`generate_email_content(recipient_data, anthropic_client)`:** Generates the email body using the Anthropic Claude API.  It creates a detailed prompt instructing the API to generate a personalized sales proposal based on the recipient's information.
*   **`send_email(recipient_email, subject, body)`:** Sends the generated email using SMTP.  It reads the email credentials and SMTP server settings from environment variables.
*   **`main()`:** The main function that orchestrates the process:
    *   Initializes the Anthropic client.
    *   Reads the Excel file.
    *   Iterates through each row in the DataFrame, generating and sending an email to each recipient.
    *   Includes a `time.sleep()` call to add a delay between emails, preventing rate limiting.

## Key Improvements and Considerations

*   **Environment Variables:** Using `.env` files and `os.getenv()` is crucial for security. Do not hardcode sensitive information directly in the script.
*   **Error Handling:** The `try...except` blocks handle potential errors during file reading, API calls, and email sending. This makes the script more robust.
*   **Rate Limiting:** The `time.sleep(5)` call is essential to avoid hitting rate limits imposed by email providers and the Anthropic API. Adjust the delay as needed.  You might also consider using a more sophisticated rate limiting strategy.
*   **HTML Email:**  Consider generating HTML emails instead of plain text emails for a more visually appealing presentation.  You would need to modify the `MIMEText` attachment in the `send_email` function.
*   **Email Tracking:** Integrate email tracking to monitor open rates and click-through rates, allowing you to refine your email content and targeting.
*   **Authentication:** Consider using OAuth 2.0 for email authentication instead of storing passwords directly.
*   **Data Validation:**  Add data validation to ensure the Excel file has the correct format and that the data is valid.
*   **Logging:** Implement logging to a file for better debugging and monitoring.
*   **Scalability:** For sending large volumes of emails, consider using a dedicated email marketing service (e.g., SendGrid, Mailgun) instead of relying on SMTP.
*   **Security:** Always be mindful of security best practices when handling email credentials and API keys.
