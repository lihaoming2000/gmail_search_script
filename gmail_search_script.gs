/**
 * Gmail Search and Update Google Sheet Script
 * Script generation time: 2025-04-13 12:17
 * 
 * This script searches Gmail for sent emails and replies based on email addresses in a Google Sheet,
 * then updates the sheet with the found information.
 */

function updateEmailDatesAndStatus() {
  // Get the active spreadsheet and the active sheet
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  
  // Get all data from the sheet
  const data = sheet.getDataRange().getValues();
  
  // Find the column indexes
  const headers = data[0];
  const nameColIndex = headers.indexOf("Name");
  const emailColIndex = headers.indexOf("Email");
  const coldEmailDateColIndex = headers.indexOf("Cold Email Date");
  const statusColIndex = headers.indexOf("Status");
  const companyColIndex = headers.indexOf("Company"); // Add company for additional search help
  
  // Check if all required columns exist
  if (nameColIndex === -1 || emailColIndex === -1 || coldEmailDateColIndex === -1 || statusColIndex === -1) {
    throw new Error("One or more required columns are missing. Please ensure your sheet has Name, Email, Cold Email Date, and Status columns.");
  }
  
  // Loop through each row (skip header row)
  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameColIndex];
    const email = data[i][emailColIndex];
    const company = companyColIndex !== -1 ? data[i][companyColIndex] : "";
    
    // Skip empty emails
    if (!email) {
      continue;
    }
    
    try {
      // Search for the first sent email to this recipient
      const firstSentEmail = findFirstSentEmail(email, name);
      if (firstSentEmail) {
        sheet.getRange(i + 1, coldEmailDateColIndex + 1).setValue(firstSentEmail);
      }
      
      // Search for any legitimate replies from this sender (excluding auto-replies)
      const firstReply = findReplies(email, name, company);
      if (firstReply) {
        sheet.getRange(i + 1, statusColIndex + 1).setValue(firstReply);
      }
    } catch (e) {
      console.error("Error processing row " + (i + 1) + " (" + email + "): " + e.message);
    }
    
    // Add a small delay to avoid hitting rate limits
    Utilities.sleep(300);
  }
  
  // Show a completion message
  SpreadsheetApp.getUi().alert("Email search and update complete!");
}

/**
 * Find the date of the first email sent to the given recipient
 * @param {string} recipientEmail - The email address to search for
 * @param {string} recipientName - The name of the recipient
 * @return {Date|null} The date of the first sent email or null if none found
 */
function findFirstSentEmail(recipientEmail, recipientName) {
  // Try multiple queries to find sent emails
  const queries = [
    `to:(${recipientEmail}) in:sent`,
    `to:(${recipientName} ${getEmailDomain(recipientEmail)}) in:sent`
  ];
  
  for (let q = 0; q < queries.length; q++) {
    const query = queries[q];
    
    try {
      // Search for threads matching the query
      const threads = GmailApp.search(query, 0, 5);
      
      if (threads.length === 0) continue;
      
      // Check each thread for matching messages
      for (let i = 0; i < threads.length; i++) {
        try {
          const messages = threads[i].getMessages();
          
          for (let j = 0; j < messages.length; j++) {
            const message = messages[j];
            const to = message.getTo().toLowerCase();
            
            // Check if this message was sent to our target
            if (to.indexOf(recipientEmail.toLowerCase()) !== -1) {
              return message.getDate();
            }
          }
        } catch (e) {
          console.error(`Error in thread ${i+1}: ${e.message}`);
        }
      }
    } catch (e) {
      console.error(`Error executing query ${q+1}: ${e.message}`);
    }
  }
  
  return null;
}

/**
 * Find legitimate replies (excluding auto-replies) using multiple search methods
 * @param {string} senderEmail - The email address to search for
 * @param {string} senderName - The name of the sender
 * @param {string} senderCompany - The company of the sender
 * @return {Date|null} The date of the first legitimate reply or null if none found
 */
function findReplies(senderEmail, senderName, senderCompany) {
  // Build an array of strategies to try
  const strategies = [];
  
  // 1. Direct email search
  strategies.push(`from:(${senderEmail}) in:anywhere`);
  
  // 2. Extract and use just the email username part
  const username = senderEmail.split('@')[0];
  strategies.push(`from:(${username}) in:anywhere`);
  
  // 3. Domain-based search
  const domain = getEmailDomain(senderEmail);
  strategies.push(`from:(@${domain}) in:anywhere`);
  
  // 4. Name + domain search
  if (senderName && senderName.trim() !== "") {
    strategies.push(`from:(${senderName} @${domain}) in:anywhere`);
  }
  
  // 5. Company search if available
  if (senderCompany && senderCompany.trim() !== "") {
    strategies.push(`from:(@${domain}) ${senderCompany} in:anywhere`);
  }
  
  // Try each strategy
  for (let s = 0; s < strategies.length; s++) {
    const query = strategies[s];
    
    try {
      // Search for threads matching the query
      const threads = GmailApp.search(query, 0, 10);
      
      if (threads.length === 0) continue;
      
      // Check each thread
      for (let i = 0; i < threads.length; i++) {
        try {
          const messages = threads[i].getMessages();
          
          // Extract detailed information from each message
          for (let j = 0; j < messages.length; j++) {
            const message = messages[j];
            const from = message.getFrom().toLowerCase();
            const subject = message.getSubject().toLowerCase();
            const body = message.getPlainBody().toLowerCase();
            
            // Skip auto-replies and bounce notifications
            if (isAutoReply(subject, body)) {
              continue;
            }
            
            // Extract email from the From field using regex
            const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/;
            const fromEmailMatch = from.match(emailRegex);
            const fromEmail = fromEmailMatch ? fromEmailMatch[0].toLowerCase() : "";
            
            // Check for matches with various criteria
            const isDirectMatch = fromEmail === senderEmail.toLowerCase();
            const isDomainMatch = fromEmail.endsWith(domain.toLowerCase());
            const containsName = senderName && from.includes(senderName.toLowerCase());
            
            // If any match criteria is met, return this message date
            if (isDirectMatch || (isDomainMatch && containsName)) {
              return message.getDate();
            }
          }
        } catch (e) {
          console.error(`Error in thread ${i+1}: ${e.message}`);
        }
      }
    } catch (e) {
      console.error(`Error executing strategy ${s+1}: ${e.message}`);
    }
  }
  
  // If we've made it here, let's also try a conversation-based approach
  try {
    // Find conversations where we sent email to this person
    const sentQuery = `to:(${senderEmail}) in:sent`;
    
    const sentThreads = GmailApp.search(sentQuery, 0, 20);
    
    // Check each conversation for replies
    for (let i = 0; i < sentThreads.length; i++) {
      try {
        const thread = sentThreads[i];
        const messages = thread.getMessages();
        
        // Skip threads with only one message (no replies)
        if (messages.length <= 1) continue;
        
        // Look at each message after our sent message
        for (let j = 0; j < messages.length; j++) {
          const message = messages[j];
          const from = message.getFrom().toLowerCase();
          const subject = message.getSubject().toLowerCase();
          const body = message.getPlainBody().toLowerCase();
          
          // Skip auto-replies and bounce notifications
          if (isAutoReply(subject, body)) {
            continue;
          }
          
          // Check if this could be a reply from our target
          const possibleSender = (
            from.includes(domain.toLowerCase()) || 
            (senderName && from.includes(senderName.toLowerCase())) ||
            from.includes(username.toLowerCase())
          );
          
          if (possibleSender) {
            return message.getDate();
          }
        }
      } catch (e) {
        console.error(`Error in sent thread ${i+1}: ${e.message}`);
      }
    }
  } catch (e) {
    console.error(`Error in conversation approach: ${e.message}`);
  }
  
  return null;
}

/**
 * Determines if an email is an auto-reply or bounce notification based on subject and body
 * @param {string} subject - Email subject (lowercase)
 * @param {string} body - Email body (lowercase)
 * @return {boolean} True if the email appears to be an auto-reply or bounce
 */
function isAutoReply(subject, body) {
  // Common phrases in auto-reply subjects
  const autoReplySubjects = [
    "out of office",
    "automatic reply",
    "auto-reply",
    "autoreply",
    "auto reply",
    "vacation",
    "away from my email",
    "absence notification",
    "ooo",
    "not at my desk",
    "automated response",
    "delivery status notification",
    "undeliverable",
    "delivery failed",
    "mail delivery failed",
    "delivery notification",
    "mail delivery system",
    "mailer-daemon",
    "delivery failure",
    "failed delivery",
    "could not be delivered",
    "bounce",
    "returned mail",
    "delivery status"
  ];
  
  // Common phrases in auto-reply bodies
  const autoReplyBodies = [
    "out of office",
    "i am currently out of the office",
    "i'm away from the office",
    "i will be out of the office",
    "i am on vacation",
    "automatic reply",
    "automated response",
    "not in the office",
    "will return on",
    "i will respond to your email when i return",
    "i will have limited access to email",
    "away until",
    "thank you for your email. i am currently",
    "this is an automatic confirmation",
    "undeliverable message",
    "unable to deliver your message",
    "permanent error",
    "delivery has failed",
    "wasn't delivered",
    "delivery to the following recipients failed",
    "your message couldn't be delivered",
    "it will be delivered after approval",
    "this mailbox is not monitored",
    "email address you entered couldn't be found",
    "user unknown",
    "recipient address rejected",
    "mailbox unavailable",
    "recipient not found",
    "does not exist",
    "account has been disabled",
    "address is administratively disabled",
    "mailbox is full",
    "storage quota exceeded",
    "550 5.1.1",
    "550 5.4.1",
    "550 5.2.1"
  ];
  
  // Check subject for auto-reply indicators
  for (const phrase of autoReplySubjects) {
    if (subject.includes(phrase)) {
      return true;
    }
  }
  
  // Check body for auto-reply indicators (just check first 1000 chars to save processing)
  const bodyStart = body.substring(0, 1000);
  for (const phrase of autoReplyBodies) {
    if (bodyStart.includes(phrase)) {
      return true;
    }
  }
  
  return false;
}

/**
 * Extract the domain part from an email address
 */
function getEmailDomain(email) {
  const parts = email.split('@');
  return parts.length > 1 ? parts[1] : '';
}

/**
 * Format a date as MM/DD/YYYY
 */
function formatDate(date) {
  if (!date) return "";
  return (date.getMonth() + 1) + "/" + date.getDate() + "/" + date.getFullYear();
}

/**
 * Creates a menu item to run the script
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Lookup')
      .addItem('Update Email Dates and Status', 'updateEmailDatesAndStatus')
      .addToUi();
}
