# Gmail Search and Update for Google Sheets

> Built with Claude + MCP (Multimodal Code Pad)

A Google Apps Script tool that automatically searches Gmail for sent emails and replies based on a list of contacts in a Google Sheet.

## Overview

This script helps you track email correspondence with a list of contacts. For each contact in your Google Sheet, the script will:

1. Search your Gmail "Sent" folder to find when you first emailed them
2. Search your Gmail "All Mail" to find if they've replied to you (excluding auto-replies)
3. Update your Google Sheet with these dates

Perfect for tracking outreach campaigns, cold emails, and follow-ups.

## Features

- **Automated Email Tracking**: Automatically finds and records when you first contacted someone and if/when they replied
- **Multiple Search Strategies**: Uses advanced search techniques to find emails even when standard searches fail
- **Intelligent Matching**: Uses various criteria to identify and match emails correctly
- **Auto-Reply Filtering**: Ignores out-of-office messages and bounce notifications
- **Custom Menu Integration**: Adds a custom menu item to your Google Sheet for easy access

## Requirements

- A Google Sheet with at least the following columns:
  - Name
  - Email
  - Cold Email Date
  - Status
- Optional but helpful columns:
  - Company
- Gmail access with the emails you want to track

## Installation

1. Open your Google Sheet
2. Click on "Extensions" > "Apps Script"
3. Delete any existing code in the editor
4. Paste the entire script from `gmail_search_script.gs`
5. Save the project (File > Save or Ctrl+S)
6. Refresh your Google Sheet

You should now see an "Email Lookup" menu item in your Google Sheet.

## Usage

1. Ensure your Google Sheet has the required columns
2. Fill in the "Name" and "Email" columns with your contacts' information
3. Click on "Email Lookup" > "Update Email Dates and Status"
4. Grant the necessary permissions when prompted
5. Wait for the script to complete
6. Your sheet will be updated with the dates of first contact and legitimate replies

## How It Works

The script uses multiple search strategies to find emails:

### For Finding First Contact:
- Direct email search in the "Sent" folder
- Name + domain search in the "Sent" folder

### For Finding Replies:
1. Direct email search
2. Username search (just the part before @)
3. Domain search (just the domain)
4. Name + domain search
5. Company + domain search (if available)
6. Conversation-based approach (looking at threads where you emailed the person)

### Auto-Reply Filtering
The script automatically detects and ignores:
- Out of office messages
- Automatic replies
- Delivery failure notifications
- Bounce messages
- Other automated system responses

Each email is checked against common patterns in both subject and body to determine if it's an automated response.

## Limitations

- The script can only search emails accessible to your Google account
- It may not find emails if they use unusual formatting in the "From" field
- Gmail API quotas may limit how many emails you can process at once
- The script needs appropriate permissions to access your Gmail

## Troubleshooting

If the script doesn't find certain emails:
1. Verify you can manually find the email in Gmail
2. Check that the email address in your sheet matches the one used in Gmail
3. Try modifying the search strategies in the script if needed

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
