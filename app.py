import streamlit as st
import os
import sys
import imaplib
import email
from email.header import decode_header
import re
from bs4 import BeautifulSoup
import pandas as pd
import json
from datetime import datetime
from dotenv import load_dotenv
import openai

# Load environment variables
load_dotenv()

# Set OpenAI API key
openai.api_key = os.getenv('OPENAI_API_KEY')

# Page configuration
st.set_page_config(
    page_title="Email Parser Application",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Define class equivalents of the original app inline

class EmailParser:
    def __init__(self):
        # Define fields to extract - can be expanded or modified
        self.fields_to_extract = [
            'vendor_name',
            'amount_due',
            'date_due',
            'order_number',
            'order_date',
            'total_amount',
            'shipping_address',
            'tracking_number',
            'email_from'
        ]
    
    def parse(self, email_content):
        """
        Parse email content and extract structured data using OpenAI's API
        """
        # Use OpenAI API to extract structured data
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "Extract structured data from emails."},
                {"role": "user", "content": f"Extract the following fields from the email content: {', '.join(self.fields_to_extract)}.\n\nEmail Content:\n{email_content}"}
            ],
            max_tokens=500,
            temperature=0.0
        )
        
        # Process the response
        extracted_data = self._process_openai_response(response)
        return extracted_data
    
    def _process_openai_response(self, response):
        """
        Process the response from OpenAI API to extract structured data
        """
        extracted_data = {}
        
        # Assuming the response is a JSON-like structure
        try:
            structured_data = json.loads(response.choices[0].message['content'].strip())
            for field in self.fields_to_extract:
                extracted_data[field] = structured_data.get(field, None)
        except json.JSONDecodeError:
            print("Failed to decode OpenAI response")
        
        return extracted_data
    
    def parse_html(self, html_content):
        """
        Parse HTML email content
        """
        # Convert HTML to plain text first
        soup = BeautifulSoup(html_content, 'html.parser')
        text_content = soup.get_text(' ', strip=True)
        
        # Extract structured data using LLM-like approach
        extracted_data = self._extract_structured_data(text_content, html_content)
        
        # Try to extract tables from HTML for items
        items = self._extract_items_from_html(soup)
        
        if items:
            extracted_data['items'] = items
        
        return extracted_data
    
    def parse_text(self, text_content):
        """
        Parse plain text email content
        """
        # Extract structured data using LLM-like approach
        extracted_data = self._extract_structured_data(text_content)
        
        # Try to extract items from text
        items = self._extract_items_from_text(text_content)
        
        if items:
            extracted_data['items'] = items
            
        return extracted_data
    
    def _extract_structured_data(self, text_content, html_content=None):
        """
        Extract structured data using a more sophisticated approach.
        This simulates what an LLM would do by using targeted extraction logic
        rather than purely regex.
        """
        extracted_data = {}
        
        # 1. Extract vendor_name
        vendor_name = self._extract_vendor_name(text_content)
        if vendor_name:
            extracted_data['vendor_name'] = vendor_name
        
        # 2. Extract amount_due
        amount_due = self._extract_amount_due(text_content, html_content)
        if amount_due:
            extracted_data['amount_due'] = amount_due
        
        # 3. Extract date_due
        date_due = self._extract_date_due(text_content)
        if date_due:
            extracted_data['date_due'] = date_due
        
        # 4. Extract order_number
        order_number = self._extract_order_number(text_content)
        if order_number:
            extracted_data['order_number'] = order_number
        
        # 5. Extract order_date
        order_date = self._extract_order_date(text_content)
        if order_date:
            extracted_data['order_date'] = order_date
        
        # 6. Extract total_amount
        total_amount = self._extract_total_amount(text_content, html_content)
        if total_amount:
            extracted_data['total_amount'] = total_amount
        
        # 7. Extract shipping_address
        shipping_address = self._extract_shipping_address(text_content)
        if shipping_address:
            extracted_data['shipping_address'] = shipping_address
        
        # 8. Extract tracking_number
        tracking_number = self._extract_tracking_number(text_content)
        if tracking_number:
            extracted_data['tracking_number'] = tracking_number
        
        # 9. Extract email_from
        email_from = self._extract_email_from(text_content)
        if email_from:
            extracted_data['email_from'] = email_from
        
        return extracted_data
    
    def _extract_vendor_name(self, text):
        """Extract vendor name using multiple approaches"""
        # First try common patterns
        patterns = [
            r'(?:From|Vendor|Seller|Company)[:\s]+([A-Za-z0-9\s,.]+)(?=\n|<|,|\()',
            r'Thank you for (?:your order|shopping) (?:from|with|at) ([A-Za-z0-9\s,.&]+)',
            r'([A-Za-z0-9\s,.&]+) Order Confirmation',
            r'Welcome to ([A-Za-z0-9\s,.&]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        # Try to find company name at start of email or in signature
        lines = text.split('\n')
        for i, line in enumerate(lines[:5]):  # Check first 5 lines
            if len(line.strip()) > 0 and len(line.strip()) < 50:  # Reasonable length for company name
                if not re.search(r'@|http|www|subject|dear|hi\s|hello', line, re.IGNORECASE):
                    return line.strip()
        
        # Check for possible company in email signature area
        for i in range(len(lines)-1, max(0, len(lines)-10), -1):  # Check last 10 lines
            if re.search(r'©|copyright|all rights reserved', lines[i], re.IGNORECASE):
                match = re.search(r'(?:©|copyright|all rights reserved)[,\s]+([A-Za-z0-9\s,.&]+)', 
                                 lines[i], re.IGNORECASE)
                if match:
                    return match.group(1).strip()
        
        return None
    
    def _extract_amount_due(self, text, html=None):
        """Extract amount due using multiple approaches"""
        patterns = [
            r'(?:Amount\s*Due|Balance\s*Due|Total\s*Due|Payment\s*Due)[:\s]*[$€£]?([0-9,.]+)',
            r'(?:Total\s*Amount\s*Due|Payment\s*Amount)[:\s]*[$€£]?([0-9,.]+)',
            r'(?:Please\s*Pay|Pay\s*Now)[:\s]*[$€£]?([0-9,.]+)',
            r'(?:Total\s*Balance|Outstanding\s*Balance)[:\s]*[$€£]?([0-9,.]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        # If HTML is available, look for specific HTML patterns that often indicate amount due
        if html:
            soup = BeautifulSoup(html, 'html.parser')
            
            # Look for amount due in structured format (like a table)
            amount_elements = soup.find_all(string=re.compile(r'amount\s*due', re.IGNORECASE))
            for element in amount_elements:
                parent = element.parent
                if parent.name in ['th', 'td']:
                    # Look for the value in the next cell
                    next_cell = parent.find_next_sibling('td')
                    if next_cell:
                        amount_match = re.search(r'[$€£]?([0-9,.]+)', next_cell.text)
                        if amount_match:
                            return amount_match.group(1).strip()
        
        # If amount due isn't found, try total amount as fallback
        return self._extract_total_amount(text, html)
    
    def _extract_date_due(self, text):
        """Extract due date using multiple approaches"""
        patterns = [
            r'(?:Due\s*Date|Payment\s*Due\s*(?:Date|By|On)|Date\s*Due)[:\s]*([A-Za-z0-9,\s]+)',
            r'(?:Pay\s*By|Payment\s*Deadline)[:\s]*([A-Za-z0-9,\s]+)',
            r'(?:due\s*on|due\s*by)[:\s]*([A-Za-z0-9,\s]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                date_str = match.group(1).strip()
                # Check if it's a valid date format
                if re.search(r'\d', date_str):  # Contains at least one digit
                    return date_str
        
        return None
    
    def _extract_order_number(self, text):
        """Extract order number using multiple approaches"""
        patterns = [
            r'Order\s*(?:Number|#|No\.)[:\s]*([A-Za-z0-9\-_]+)',
            r'(?:order|confirmation)[:\s]*#?\s*([A-Za-z0-9\-_]+)',
            r'Reference\s*(?:Number|#)[:\s]*([A-Za-z0-9\-_]+)',
            r'(?:Invoice|Receipt)\s*(?:Number|#)[:\s]*([A-Za-z0-9\-_]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        return None
    
    def _extract_order_date(self, text):
        """Extract order date using multiple approaches"""
        patterns = [
            r'Order\s*Date[:\s]*([A-Za-z0-9,\s]+)',
            r'Date\s*(?:of|on)[:\s]*Order[:\s]*([A-Za-z0-9,\s]+)',
            r'Ordered\s*on[:\s]*([A-Za-z0-9,\s]+)',
            r'Purchase\s*Date[:\s]*([A-Za-z0-9,\s]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                date_str = match.group(1).strip()
                # Check if it's a valid date format
                if re.search(r'\d', date_str):  # Contains at least one digit
                    return date_str
        
        return None
    
    def _extract_total_amount(self, text, html=None):
        """Extract total amount using multiple approaches"""
        patterns = [
            r'(?:Order\s*Total|Total)[:\s]*[$€£]?([0-9,.]+)',
            r'(?:Total\s*Amount|Grand\s*Total)[:\s]*[$€£]?([0-9,.]+)',
            r'(?:Amount|Payment)[:\s]*[$€£]?([0-9,.]+)',
            r'(?:Charged|Price)[:\s]*[$€£]?([0-9,.]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        # If HTML is available, look for structured format
        if html:
            soup = BeautifulSoup(html, 'html.parser')
            
            # Look for total in structured format (like a table)
            total_elements = soup.find_all(string=re.compile(r'(?:order\s*total|total\s*amount|grand\s*total)', re.IGNORECASE))
            for element in total_elements:
                parent = element.parent
                if parent.name in ['th', 'td']:
                    # Look for the value in the next cell
                    next_cell = parent.find_next_sibling('td')
                    if next_cell:
                        amount_match = re.search(r'[$€£]?([0-9,.]+)', next_cell.text)
                        if amount_match:
                            return amount_match.group(1).strip()
        
        return None
    
    def _extract_shipping_address(self, text):
        """Extract shipping address using multiple approaches"""
        patterns = [
            r'(?:Shipping|Delivery)\s*Address[:\s]*(.*?)(?=\n\n|\n[A-Z]|\Z)',
            r'(?:Ship\s*To|Deliver\s*To)[:\s]*(.*?)(?=\n\n|\n[A-Z]|\Z)',
            r'(?:Shipped\s*To|Delivered\s*To)[:\s]*(.*?)(?=\n\n|\n[A-Z]|\Z)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.DOTALL)
            if match:
                address = match.group(1).strip()
                # Clean up the address
                address = re.sub(r'\n+', ', ', address)
                address = re.sub(r'\s+', ' ', address)
                return address
        
        return None
    
    def _extract_tracking_number(self, text):
        """Extract tracking number using multiple approaches"""
        patterns = [
            r'(?:Tracking\s*(?:Number|#)|Track\s*Your\s*Package)[:\s]*([A-Za-z0-9]+)',
            r'(?:Tracking\s*ID|Shipment\s*ID)[:\s]*([A-Za-z0-9]+)',
            r'(?:Your\s*package\s*can\s*be\s*tracked\s*with)[:\s]*([A-Za-z0-9]+)',
            r'(?:Track)[:\s]*.*?(?:number)[:\s]*([A-Za-z0-9]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
        
        return None
    
    def _extract_email_from(self, text):
        """Extract sender email address"""
        patterns = [
            r'From:[:\s]*([A-Za-z0-9\s,.@<>]+)',
            r'Sender:[:\s]*([A-Za-z0-9\s,.@<>]+)',
            r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                email_addr = match.group(1).strip()
                # Extract just the email if it's in a format like "Name <email@example.com>"
                email_match = re.search(r'<([^>]+)>', email_addr)
                if email_match:
                    return email_match.group(1)
                return email_addr
        
        return None
    
    def _extract_items_from_html(self, soup):
        """
        Extract items from HTML content, specifically looking at tables
        """
        tables = soup.find_all('table')
        items = []
        
        for table in tables:
            if self._is_likely_items_table(table):
                items = self._extract_items_from_table(table)
                if items:
                    break
        
        return items
    
    def _is_likely_items_table(self, table):
        """
        Identify if a table is likely to contain order items
        """
        text = table.get_text().lower()
        
        # Check if table contains these keywords
        item_indicators = ['item', 'product', 'description', 'quantity', 'price', 'amount', 'subtotal']
        score = sum(1 for indicator in item_indicators if indicator in text)
        
        return score >= 3  # If at least 3 indicators are present
    
    def _extract_items_from_table(self, table):
        """
        Extract items from an HTML table
        """
        items = []
        rows = table.find_all('tr')
        
        # Skip header row
        if len(rows) < 2:
            return []
        
        # Try to determine the column indices
        header_cells = rows[0].find_all(['th', 'td'])
        header_text = [cell.get_text().strip().lower() for cell in header_cells]
        
        # Map column indices
        column_map = {}
        for i, text in enumerate(header_text):
            if 'item' in text or 'product' in text or 'description' in text:
                column_map['name'] = i
            elif 'qty' in text or 'quantity' in text:
                column_map['quantity'] = i
            elif 'price' in text or 'unit' in text or 'cost' in text:
                column_map['price'] = i
            elif 'total' in text or 'amount' in text or 'subtotal' in text:
                column_map['total'] = i
        
        # If we couldn't determine the columns, return empty
        if 'name' not in column_map:
            return []
        
        # Extract items from rows
        for row in rows[1:]:
            cells = row.find_all(['td', 'th'])
            if len(cells) < max(column_map.values()) + 1:
                continue
                
            item = {}
            
            if 'name' in column_map:
                item['name'] = cells[column_map['name']].get_text().strip()
                
            if 'quantity' in column_map:
                qty_text = cells[column_map['quantity']].get_text().strip()
                qty_match = re.search(r'(\d+)', qty_text)
                if qty_match:
                    item['quantity'] = int(qty_match.group(1))
            
            if 'price' in column_map:
                price_text = cells[column_map['price']].get_text().strip()
                price_match = re.search(r'[\$€£]?([0-9,.]+)', price_text)
                if price_match:
                    item['unit_price'] = price_match.group(1)
            
            if 'total' in column_map:
                total_text = cells[column_map['total']].get_text().strip()
                total_match = re.search(r'[\$€£]?([0-9,.]+)', total_text)
                if total_match:
                    item['total_price'] = total_match.group(1)
            
            # Only add non-empty items
            if item and 'name' in item and item['name']:
                items.append(item)
        
        return items
    
    def _extract_items_from_text(self, text):
        """
        Extract items from plain text
        """
        items = []
        
        # Look for patterns indicating items (multiple approaches)
        patterns = [
            # Pattern 1: quantity x product name, price
            r'(\d+)\s*x\s*([^,\n]+)[\s,]*(?:\$|EUR|£)?([0-9,.]+)',
            # Pattern 2: quantity product name @ price
            r'(\d+)\s+([^@\n]+)@\s*(?:\$|EUR|£)?([0-9,.]+)',
            # Pattern 3: product name (quantity) price
            r'([^()\n]+)\s*\((\d+)\)\s*(?:\$|EUR|£)?([0-9,.]+)'
        ]
        
        for pattern in patterns:
            matches = re.finditer(pattern, text)
            for match in matches:
                if pattern == patterns[0]:
                    quantity = int(match.group(1))
                    name = match.group(2).strip()
                    price = match.group(3).strip()
                elif pattern == patterns[1]:
                    quantity = int(match.group(1))
                    name = match.group(2).strip()
                    price = match.group(3).strip()
                elif pattern == patterns[2]:
                    name = match.group(1).strip()
                    quantity = int(match.group(2))
                    price = match.group(3).strip()
                
                item = {
                    'name': name,
                    'quantity': quantity,
                    'unit_price': price
                }
                
                items.append(item)
        
        # If no items found using patterns, try line-by-line analysis
        if not items:
            # Find sections that might contain items
            item_section_match = re.search(r'(?:Your Order|Order Details|Items|Products).*?(?=\n\n|\n[A-Z]|\Z)', 
                                          text, re.IGNORECASE | re.DOTALL)
            if item_section_match:
                item_text = item_section_match.group(0)
                lines = item_text.split('\n')
                
                for line in lines:
                    # Look for lines that have both product name and price indicators
                    if re.search(r'(?:\$|EUR|£)?([0-9,.]+)', line) and len(line.strip()) > 10:
                        # Extract price
                        price_match = re.search(r'(?:\$|EUR|£)?([0-9,.]+)', line)
                        if price_match:
                            price = price_match.group(1)
                            
                            # Extract name (everything before the price)
                            name_match = re.search(r'(.*?)(?=\$|EUR|£|[0-9]{1,3},[0-9]{3}|[0-9]+\.[0-9]+)', line)
                            if name_match:
                                name = name_match.group(1).strip()
                                
                                # Extract quantity if present
                                qty_match = re.search(r'(\d+)\s*x', name)
                                if qty_match:
                                    quantity = int(qty_match.group(1))
                                    name = re.sub(r'\d+\s*x\s*', '', name).strip()
                                else:
                                    quantity = 1
                                
                                items.append({
                                    'name': name,
                                    'quantity': quantity,
                                    'unit_price': price
                                })
        
        return items

class EmailConnector:
    def __init__(self, server, port, email, password):
        self.server = server
        self.port = int(port)
        self.email = email
        self.password = password
        self.connection = None
    
    def connect(self):
        """
        Connect to the email server using IMAP
        """
        try:
            self.connection = imaplib.IMAP4_SSL(self.server, self.port)
            self.connection.login(self.email, self.password)
            return f"Successfully connected to {self.email}"
        except Exception as e:
            raise Exception(f"Connection failed: {str(e)}")
    
    def disconnect(self):
        """
        Disconnect from the email server
        """
        if self.connection:
            try:
                self.connection.logout()
                return "Successfully disconnected"
            except Exception as e:
                raise Exception(f"Disconnect failed: {str(e)}")
        return "No active connection"
    
    def get_folders(self):
        """
        Get list of available folders
        """
        if not self.connection:
            raise Exception("Not connected to server")
        
        folders = []
        response, folder_list = self.connection.list()

        if response == 'OK':
            for folder in folder_list:
                folder_name = folder.decode().split('"/"')[-1].strip()
                folders.append(folder_name.replace('"', ''))
        
        return folders
    
    def fetch_emails(self, folder="INBOX", limit=10, criteria="ALL"):
        """
        Fetch emails from specified folder
        """
        if not self.connection:
            raise Exception("Not connected to server")
        
        try:
            print(f"Selecting folder: {folder}")
            response, data = self.connection.select(folder)
            if response != 'OK':
                raise Exception(f"Failed to select folder: {response}")
            
            print(f"Searching with criteria: {criteria}")
            response, messages = self.connection.search(None, criteria)
            if response != 'OK':
                raise Exception(f"Failed to search emails: {response}")
            
            email_ids = messages[0].split()
            print(f"Found {len(email_ids)} emails")
            
            # Ensure limit is an integer
            limit = int(limit)
            
            # Get the last 'limit' number of emails
            if len(email_ids) > limit:
                email_ids = email_ids[-limit:]
            
            emails = []
            
            for email_id in email_ids:
                print(f"Fetching email ID: {email_id}")
                response, msg_data = self.connection.fetch(email_id, '(RFC822)')
                if response != 'OK':
                    print(f"Failed to fetch email {email_id}: {response}")
                    continue
                
                raw_email = msg_data[0][1]
                
                # Parse the raw email
                msg = email.message_from_bytes(raw_email)
                
                subject = self._decode_email_header(msg['Subject'])
                from_address = self._decode_email_header(msg['From'])
                date = msg['Date']
                
                # Get email body
                body = ""
                
                if msg.is_multipart():
                    for part in msg.walk():
                        content_type = part.get_content_type()
                        content_disposition = str(part.get("Content-Disposition"))
                        
                        # Skip attachments
                        if "attachment" in content_disposition:
                            continue
                        
                        # Get text content
                        if content_type == "text/plain" or content_type == "text/html":
                            try:
                                body = part.get_payload(decode=True).decode('utf-8')
                            except:
                                try:
                                    body = part.get_payload(decode=True).decode('latin-1')
                                except:
                                    body = "Could not decode email body"
                            break
                else:
                    # If email is not multipart
                    try:
                        body = msg.get_payload(decode=True).decode('utf-8')
                    except:
                        try:
                            body = msg.get_payload(decode=True).decode('latin-1')
                        except:
                            body = "Could not decode email body"
                
                email_data = {
                    'id': email_id.decode(),
                    'subject': subject,
                    'from': from_address,
                    'date': date,
                    'body': body
                }
                
                emails.append(email_data)
            
            return emails
        except Exception as e:
            print(f"Error in fetch_emails: {str(e)}")
            raise
    
    def _decode_email_header(self, header):
        """
        Decode email headers
        """
        if header is None:
            return ""
            
        decoded_header = decode_header(header)
        header_parts = []
        
        for part, encoding in decoded_header:
            if isinstance(part, bytes):
                try:
                    if encoding:
                        header_parts.append(part.decode(encoding))
                    else:
                        header_parts.append(part.decode('utf-8'))
                except:
                    header_parts.append(part.decode('latin-1'))
            else:
                header_parts.append(part)
                
        return " ".join(header_parts)

class EmailController:
    def __init__(self):
        self.connector = None
        self.parser = EmailParser()
        self.export_directory = os.path.join(os.getcwd(), 'exports')
        
        # Create exports directory if it doesn't exist
        if not os.path.exists(self.export_directory):
            os.makedirs(self.export_directory)
    
    def connect(self, server, port, email, password):
        """
        Connect to email server
        """
        self.connector = EmailConnector(server, port, email, password)
        return self.connector.connect()
    
    def disconnect(self):
        """
        Disconnect from email server
        """
        if self.connector:
            return self.connector.disconnect()
        return "Not connected"
    
    def fetch_emails(self, folder="INBOX", limit=10, criteria="ALL"):
        """
        Fetch emails from specified folder
        """
        # Create a new connection using environment variables if none exists
        if not self.connector:
            server = os.getenv('EMAIL_SERVER')
            port = os.getenv('EMAIL_PORT')
            email = os.getenv('EMAIL_USER')
            password = os.getenv('EMAIL_PASSWORD')
            
            if not all([server, port, email, password]):
                raise Exception("Email connection not initialized and missing environment variables")
            
            self.connect(server, port, email, password)
        
        return self.connector.fetch_emails(folder, limit, criteria)
    
    def export_data(self, data, format_type='csv'):
        """
        Export parsed data to a file
        """
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"email_data_{timestamp}"
        
        if format_type == 'json':
            file_path = os.path.join(self.export_directory, f"{filename}.json")
            with open(file_path, 'w') as f:
                json.dump(data, f, indent=4)
            return file_path
        
        elif format_type == 'csv':
            file_path = os.path.join(self.export_directory, f"{filename}.csv")
            
            # Convert to DataFrame for CSV export
            # If data is a list of dictionaries, convert directly
            if isinstance(data, list) and all(isinstance(item, dict) for item in data):
                # Flatten nested dictionaries for CSV format
                flattened_data = []
                
                for item in data:
                    flat_item = {}
                    
                    for key, value in item.items():
                        if isinstance(value, dict):
                            for subkey, subvalue in value.items():
                                flat_item[f"{key}_{subkey}"] = subvalue
                        elif isinstance(value, list) and key == 'items':
                            # Special handling for item lists
                            for i, subitem in enumerate(value):
                                for subkey, subvalue in subitem.items():
                                    flat_item[f"item{i+1}_{subkey}"] = subvalue
                        else:
                            flat_item[key] = value
                    
                    flattened_data.append(flat_item)
                
                df = pd.DataFrame(flattened_data)
            else:
                # If it's a single dictionary, convert to a single-row DataFrame
                df = pd.DataFrame([data])
            
            df.to_csv(file_path, index=False)
            return file_path
        
        elif format_type == 'excel':
            file_path = os.path.join(self.export_directory, f"{filename}.xlsx")
            
            # Similar logic as CSV but for Excel
            if isinstance(data, list) and all(isinstance(item, dict) for item in data):
                flattened_data = []
                
                for item in data:
                    flat_item = {}
                    
                    for key, value in item.items():
                        if isinstance(value, dict):
                            for subkey, subvalue in value.items():
                                flat_item[f"{key}_{subkey}"] = subvalue
                        elif isinstance(value, list) and key == 'items':
                            for i, subitem in enumerate(value):
                                for subkey, subvalue in subitem.items():
                                    flat_item[f"item{i+1}_{subkey}"] = subvalue
                        else:
                            flat_item[key] = value
                    
                    flattened_data.append(flat_item)
                
                df = pd.DataFrame(flattened_data)
            else:
                df = pd.DataFrame([data])
            
            df.to_excel(file_path, index=False)
            return file_path
        
        else:
            raise ValueError(f"Unsupported format: {format_type}")

# Initialize session state variables if they don't exist
if 'emails' not in st.session_state:
    st.session_state.emails = []
if 'parsed_emails' not in st.session_state:
    st.session_state.parsed_emails = []
if 'current_email' not in st.session_state:
    st.session_state.current_email = None
if 'current_parsed_data' not in st.session_state:
    st.session_state.current_parsed_data = None
if 'email_controller' not in st.session_state:
    st.session_state.email_controller = EmailController()
if 'connection_status' not in st.session_state:
    st.session_state.connection_status = {"status": "", "message": ""}
if 'active_tab' not in st.session_state:
    st.session_state.active_tab = "connect"

# App title
st.title("Email Parser Application")

# Sidebar for navigation
with st.sidebar:
    st.header("Navigation")
    tabs = {
        "connect": "Connect to Email",
        "fetch": "Fetch Emails",
        "parse": "Parse Email",
        "manual": "Manual Input",
        "export": "Export Data"
    }
    
    for tab_id, tab_name in tabs.items():
        if st.button(tab_name, key=f"nav_{tab_id}"):
            st.session_state.active_tab = tab_id
            # Reset current email when switching tabs
            if tab_id != "parse":
                st.session_state.current_email = None

# Main content
def render_connect_tab():
    st.header("Connect to Email Server")
    
    with st.form("connect_form"):
        col1, col2 = st.columns(2)
        
        with col1:
            server = st.text_input("Email Server", value=os.getenv('EMAIL_SERVER', ''), placeholder="imap.example.com")
            port = st.text_input("Port", value=os.getenv('EMAIL_PORT', '993'), placeholder="993")
        
        with col2:
            email = st.text_input("Email Address", value=os.getenv('EMAIL_USER', ''), placeholder="your@email.com")
            password = st.text_input("Password", value=os.getenv('EMAIL_PASSWORD', ''), type="password")
        
        connect_button = st.form_submit_button("Connect")
        
        if connect_button:
            try:
                connection_status = st.session_state.email_controller.connect(
                    server=server or os.getenv('EMAIL_SERVER'),
                    port=port or os.getenv('EMAIL_PORT'),
                    email=email or os.getenv('EMAIL_USER'),
                    password=password or os.getenv('EMAIL_PASSWORD')
                )
                st.session_state.connection_status = {"status": "success", "message": connection_status}
            except Exception as e:
                st.session_state.connection_status = {"status": "error", "message": str(e)}
    
    # Display connection status
    if st.session_state.connection_status["status"] == "success":
        st.success(st.session_state.connection_status["message"])
    elif st.session_state.connection_status["status"] == "error":
        st.error(st.session_state.connection_status["message"])

def render_fetch_tab():
    st.header("Fetch Emails")
    
    with st.form("fetch_form"):
        col1, col2, col3 = st.columns(3)
        
        with col1:
            folder = st.text_input("Folder", value="INBOX", placeholder="INBOX")
        
        with col2:
            limit = st.number_input("Limit", min_value=1, value=10)
        
        with col3:
            criteria = st.text_input("Search Criteria", value="ALL", placeholder="ALL")
        
        fetch_button = st.form_submit_button("Fetch Emails")
        
        if fetch_button:
            with st.spinner("Fetching emails..."):
                try:
                    emails = st.session_state.email_controller.fetch_emails(
                        folder=folder,
                        limit=int(limit),
                        criteria=criteria
                    )
                    st.session_state.emails = emails
                    if emails:
                        st.success(f"Found {len(emails)} emails")
                    else:
                        st.info("No emails found")
                except Exception as e:
                    st.error(f"Failed to fetch emails: {str(e)}")
    
    # Display fetched emails
    if st.session_state.emails:
        st.subheader("Fetched Emails")
        
        # Create a DataFrame for display
        emails_df = pd.DataFrame([
            {
                "Subject": email.get("subject", "No Subject"),
                "From": email.get("from", "No Sender"),
                "Date": email.get("date", "No Date"),
                "Action": "Parse"
            } for email in st.session_state.emails
        ])
        
        # Display the table
        for i, row in emails_df.iterrows():
            col1, col2, col3, col4 = st.columns([2, 2, 2, 1])
            
            with col1:
                st.write(row["Subject"])
            with col2:
                st.write(row["From"])
            with col3:
                st.write(row["Date"])
            with col4:
                if st.button("Parse", key=f"parse_btn_{i}"):
                    st.session_state.current_email = st.session_state.emails[i]
                    st.session_state.active_tab = "parse"
                    st.rerun()
            
            st.divider()

def render_parse_tab():
    st.header("Parse Email")
    
    if st.session_state.current_email:
        # Display email info
        col1, col2 = st.columns(2)
        
        with col1:
            st.text_input("Subject", value=st.session_state.current_email.get("subject", ""), disabled=True)
        
        with col2:
            st.text_input("From", value=st.session_state.current_email.get("from", ""), disabled=True)
        
        st.text_input("Date", value=st.session_state.current_email.get("date", ""), disabled=True)
        
        # Parse the email content
        with st.spinner("Parsing email..."):
            try:
                parser = EmailParser()
                parsed_data = parser.parse(st.session_state.current_email.get("body", ""))
                st.session_state.current_parsed_data = parsed_data
            except Exception as e:
                st.error(f"Error parsing email: {str(e)}")
        
        # Display parsed data
        if st.session_state.current_parsed_data:
            st.subheader("Parsed Data")
            st.json(st.session_state.current_parsed_data)
            
            # Add to export button
            if st.button("Add to Export"):
                if st.session_state.current_parsed_data:
                    st.session_state.parsed_emails.append(st.session_state.current_parsed_data)
                    st.success("Added to export list")
    else:
        st.info("No email selected. Please go to the Fetch Emails tab and select an email to parse.")

def render_manual_tab():
    st.header("Manual Email Input")
    
    # Text area for pasting email content
    email_content = st.text_area("Paste Email Content", height=300, placeholder="Paste email content here...")
    
    if st.button("Parse"):
        if not email_content:
            st.warning("Please enter email content.")
        else:
            with st.spinner("Parsing content..."):
                try:
                    parser = EmailParser()
                    parsed_data = parser.parse(email_content)
                    st.session_state.current_parsed_data = parsed_data
                    
                    # Display parsed data
                    st.subheader("Parsed Data")
                    st.json(parsed_data)
                    
                    # Add to export button
                    if st.button("Add to Export", key="manual_export_btn"):
                        st.session_state.parsed_emails.append(parsed_data)
                        st.success("Added to export list")
                except Exception as e:
                    st.error(f"Error parsing content: {str(e)}")

def render_export_tab():
    st.header("Export Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Export format selection
        format_type = st.selectbox(
            "Export Format",
            options=["csv", "json", "excel"],
            index=0
        )
    
    # Export and clear buttons
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("Export Data"):
            if not st.session_state.parsed_emails:
                st.warning("No data to export.")
            else:
                with st.spinner("Exporting data..."):
                    try:
                        export_path = st.session_state.email_controller.export_data(
                            st.session_state.parsed_emails,
                            format_type
                        )
                        st.success(f"Data exported successfully to: {export_path}")
                    except Exception as e:
                        st.error(f"Export failed: {str(e)}")
    
    with col2:
        if st.button("Clear Data"):
            if st.session_state.parsed_emails:
                st.session_state.parsed_emails = []
                st.success("Data cleared successfully")
            else:
                st.info("No data to clear")
    
    # Display emails to export
    st.subheader(f"Emails to Export ({len(st.session_state.parsed_emails)})")
    
    for i, data in enumerate(st.session_state.parsed_emails):
        with st.expander(f"Item {i+1}: {data.get('vendor_name') or data.get('order_number') or f'Item {i+1}'}"):
            col1, col2 = st.columns([4, 1])
            
            with col1:
                st.write(f"**Order Total:** {data.get('total_amount', 'N/A')}")
                st.write(f"**Items:** {len(data.get('items', []))}")
                st.json(data)
            
            with col2:
                if st.button("Remove", key=f"remove_btn_{i}"):
                    st.session_state.parsed_emails.pop(i)
                    st.rerun()

# Render the active tab
if st.session_state.active_tab == "connect":
    render_connect_tab()
elif st.session_state.active_tab == "fetch":
    render_fetch_tab()
elif st.session_state.active_tab == "parse":
    render_parse_tab()
elif st.session_state.active_tab == "manual":
    render_manual_tab()
elif st.session_state.active_tab == "export":
    render_export_tab()

# Footer
st.markdown("---")
st.markdown("Email Parser Application - Streamlit Version") 