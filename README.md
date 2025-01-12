# Prevista Invoice Automation

## Overview

The Prevista Invoice Automation project is designed to streamline the process of generating and submitting invoices for employees and tutors. The application leverages Microsoft Graph API to interact with SharePoint and OneDrive, featuring fetch calendar events, validate sessions, and upload invoices and timesheets to the appropriate folders.

## Features

- **Invoice Generation**: Automatically generate invoices based on user input and predefined templates.
- **Timesheet Generation**: For tutors, generate timesheets based on session data.
- **Calendar Event Validation**: Validate session times against calendar events fetched from Microsoft Graph API.
- **SharePoint Integration**: Upload invoices and timesheets to SharePoint folders.
- **Email Notifications**: Send email notifications with logs and attachments.

## Requirements

- Python 3.7+
- Streamlit
- Microsoft Graph API
- Microsoft Authentication Library (MSAL)
- Pandas
- OpenPyXL
- Python-Docx
- Requests
- Dotenv
- Pytz
- Shutil
- Smtplib

## Installation

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/prevista-invoice-automation.git
    cd prevista-invoice-automation
    ```

2. Install the required packages:
    ```sh
    pip install -r requirements.txt
    ```

3. Create a `.env` file in the root directory and add your environment variables:
    ```env
    CLIENT_ID=your_client_id
    CLIENT_SECRET=your_client_secret
    TENANT_ID=your_tenant_id
    DRIVE_ID=your_drive_id
    EMAIL=your_email
    PASSWORD=your_email_password
    ```

## Usage

1. Run the Streamlit application:
    ```sh
    streamlit run app_v2.py
    ```

2. Open your web browser and navigate to the provided local URL (e.g., `http://localhost:8501`).

3. Follow the on-screen instructions to generate and submit your invoice.

## File Structure

- `app_v2.py`: Main application file.
- `resources/`: Directory containing template files and other resources.
- `README.md`: Project documentation.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License

This project is opensource & created for prevista.co.uk

Dev : https://linkedin.com/in/osamatech786