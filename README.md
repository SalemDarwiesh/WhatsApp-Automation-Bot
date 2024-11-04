## WhatsApp Automation Bot
This project automates sending WhatsApp messages via the desktop app using Python's pyautogui and other utilities. It reads contact numbers and messages from an Excel file, sends messages to each contact, and records the status back to the file.


## Overview
The WhatsApp Automation Bot is designed to streamline the process of sending bulk messages via WhatsApp. It reads an Excel file containing phone numbers, opens the WhatsApp desktop application, searches for each contact, sends a message, and updates the Excel file with the delivery status. This is especially useful for business messaging, notifications, and reminders.

## Features
Automated Message Sending: Opens the WhatsApp desktop app and sends predefined messages to contacts.
Excel Integration: Reads contacts from and writes message statuses to an Excel file.
Error Handling: Records "Failed to send" status if a message cannot be delivered.

## Requirements
- Python 3.x
- Required Python libraries:
- pandas
- pyautogui
- clipboard
- pytesseract
- Pillow
You can install the required libraries with:


``pip install pandas pyautogui clipboard pytesseract pillow ``
## Installation
1. Clone the repository:


`` git clone https://github.com/yourusername/whatsapp-automation-bot.git ``
cd whatsapp-automation-bot

2. Install the required libraries:


`` pip install -r requirements.txt ``
Set up the Tesseract OCR for text recognition (if needed for image-based number verification).

## Usage
1. Prepare the Excel File:

- Create an Excel file (example.xlsx) with a PhoneNumber column containing the contact numbers and a Result column (leave blank initially).

2. Run the Script:


``python whatsapp_bot.py``
3. The script will:

- Load the contacts from the Excel file.
- Send the specified message to each contact.
- Save the status of each message (sent/failed) back to the Excel file.
## Disclaimer
- This bot is intended for educational purposes only. Use it responsibly to avoid violating WhatsApp's terms of service.
- Excessive or unsolicited messaging may result in temporary or permanent account bans.
## Contributing
Contributions are welcome! Please submit a pull request or open an issue if you have suggestions.

```

This README provides a solid foundation for documenting your project on GitHub. Let me know if you'd like to add or adjust any sections! ​
