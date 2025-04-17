# Email Parser Application - Streamlit Version

A standalone Streamlit application for parsing emails to extract structured data like order information, tracking numbers, and shipping details.

## Features

- **Connect to Email Server**: Connect to your email server using IMAP
- **Fetch Emails**: Retrieve emails from specific folders with search criteria
- **Parse Email**: Extract structured data from emails using pattern recognition
- **Manual Input**: Paste email content for parsing without connecting to an email server
- **Export Data**: Export parsed data to CSV, JSON, or Excel formats

## Installation

1. Clone this repository:
```bash
git clone <repository-url>
cd email-parser-app
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install the required packages:
```bash
pip install -r requirements_streamlit.txt
```

4. Set up environment variables by creating a `.env` file with the following content:
```
EMAIL_SERVER=imap.example.com
EMAIL_PORT=993
EMAIL_USER=your@email.com
EMAIL_PASSWORD=yourpassword
```

## Running the Application

Run the Streamlit app:
```bash
streamlit run streamlit_app.py
```

The application will open in your default web browser at http://localhost:8501.

## Usage

1. **Connect to Email Server**:
   - Enter your email server details or use the ones from the `.env` file
   - Click "Connect" to establish a connection

2. **Fetch Emails**:
   - Select a folder, limit, and search criteria
   - Click "Fetch Emails" to retrieve emails

3. **Parse Email**:
   - Click "Parse" next to an email in the fetch results to extract data
   - Review the parsed data structure

4. **Manual Input**:
   - Paste email content directly into the text area
   - Click "Parse" to extract data

5. **Export Data**:
   - Choose a format (CSV, JSON, or Excel)
   - Click "Export Data" to save the results

## Security Note

The application stores your email password in the `.env` file. Make sure to:
- Never commit the `.env` file to version control
- Set appropriate file permissions to restrict access
- Consider using app-specific passwords if your email provider supports them 