// Google Apps Script - VELISSE Booking System
// Αντιγράψτε αυτόν τον κώδικα στο Google Apps Script σας

// ============================================
// CONFIGURATION - ΑΛΛΑΞΤΕ ΤΑ ΠΑΡΑΚΑΤΩ
// ============================================

const SPREADSHEET_ID = 'YOUR_SHEET_ID_HERE'; // Βρείτε στο URL του Sheet
const SHEET_NAME = 'Bookings'; // Όνομα του worksheet
const TWILIO_ACCOUNT_SID = 'YOUR_TWILIO_SID'; // Από Twilio
const TWILIO_AUTH_TOKEN = 'YOUR_TWILIO_TOKEN'; // Από Twilio
const TWILIO_PHONE = '+1234567890'; // Το Twilio phone number σας
const VIBER_API_TOKEN = 'YOUR_VIBER_TOKEN'; // Από Viber
const ADMIN_EMAIL = 'velisselashesland@gmail.com'; // Το email σας

// ============================================
// MAIN HANDLER
// ============================================

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    if (data.action === 'addBooking') {
      // Προσθήκη ραντεβού στο Sheet
      addBookingToSheet(data);
      
      // Αποθήκευση reminder για αργότερα
      scheduleReminders(data);
      
      // Ενημέρωση admin
      sendAdminNotification(data);
      
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        message: 'Booking saved successfully'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
  } catch (error) {
    Logger.log('Error: ' + error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// 1. SAVE TO GOOGLE SHEET
// ============================================

function addBookingToSheet(booking) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  
  // Headers (δημιουργία αν δεν υπάρχουν)
  const headers = [
    'id', 'name', 'phone', 'birthday', 'service', 'duration', 
    'data', 'start_time', 'end_time', 'price', 'final_price', 
    'client_type', 'discount', 'friend_name', 'friend_phone', 
    'is_first_time', 'is_student', 'is_bring_friend', 'gift', 
    'created_at', 'birthday_type', 'email'
  ];
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }
  
  // Add booking row
  const row = [
    booking.id,
    booking.name,
    booking.phone,
    booking.birthday || '',
    booking.service,
    booking.duration || 60,
    booking.data || booking.date,
    booking.start_time || booking.time,
    booking.end_time,
    booking.price || 0,
    booking.final_price || booking.price || 0,
    booking.client_type || 'regular',
    booking.discount || 0,
    booking.friend_name || '',
    booking.friend_phone || '',
    booking.is_first_time ? 'YES' : 'NO',
    booking.is_student ? 'YES' : 'NO',
    booking.is_bring_friend ? 'YES' : 'NO',
    booking.gift ? 'YES' : 'NO',
    booking.created_at,
    booking.birthday_type || 'none',
    booking.email || ''
  ];
  
  sheet.appendRow(row);
  Logger.log('✅ Booking added to sheet: ' + booking.name);
}

// ============================================
// 2. SCHEDULE REMINDERS
// ============================================

function scheduleReminders(booking) {
  const reminderSheet = SpreadsheetApp.openById(SPREADSHEET_ID)
    .getSheetByName('Reminders') || createRemindersSheet();
  
  const bookingTime = new Date(booking.data + 'T' + booking.start_time);
  const oneDayBefore = new Date(bookingTime.getTime() - 24*60*60*1000);
  const twoHoursBefore = new Date(bookingTime.getTime() - 2*60*60*1000);
  
  reminderSheet.appendRow([
    booking.id,
    booking.name,
    booking.phone,
    booking.email,
    '1 day before',
    oneDayBefore.toISOString(),
    'pending'
  ]);
  
  reminderSheet.appendRow([
    booking.id,
    booking.name,
    booking.phone,
    booking.email,
    '2 hours before',
    twoHoursBefore.toISOString(),
    'pending'
  ]);
  
  Logger.log('⏰ Reminders scheduled for: ' + booking.name);
}

function createRemindersSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.insertSheet('Reminders');
  sheet.appendRow([
    'booking_id', 'name', 'phone', 'email', 'reminder_type', 
    'send_time', 'status'
  ]);
  return sheet;
}

// ============================================
// 3. SEND ADMIN NOTIFICATION
// ============================================

function sendAdminNotification(booking) {
  const emailBody = `
    ✅ Νέο ραντεβού δημιουργήθηκε!
    
    👤 Πελάτης: ${booking.name}
    📱 Τηλέφωνο: ${booking.phone}
    📧 Email: ${booking.email}
    📅 Ημερομηνία: ${booking.data}
    ⏰ Ώρα: ${booking.start_time}
    💇 Υπηρεσία: ${booking.service}
    💰 Τιμή: ${booking.final_price}€
    👥 Τύπος: ${booking.client_type}
    
    ${booking.is_bring_friend ? '👭 Φέρνει φίλη: ' + booking.friend_name + ' (' + booking.friend_phone + ')' : ''}
  `;
  
  GmailApp.sendEmail(ADMIN_EMAIL, 'Νέο ραντεβού - ' + booking.name, emailBody);
  Logger.log('📧 Admin notification sent');
}

// ============================================
// 4. CHECK & SEND REMINDERS (Run every 5 mins)
// ============================================

function checkAndSendReminders() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Reminders');
  
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  for (let i = 1; i < data.length; i++) {
    const reminder = data[i];
    const sendTime = new Date(reminder[5]);
    const status = reminder[6];
    
    if (status === 'pending' && sendTime <= now) {
      const reminderType = reminder[4];
      const bookingId = reminder[0];
      const name = reminder[1];
      const phone = reminder[2];
      const email = reminder[3];
      
      // Send reminders
      sendWhatsAppReminder(name, phone, reminderType);
      sendViberReminder(name, phone, reminderType);
      sendEmailReminder(email, name, reminderType);
      
      // Mark as sent
      sheet.getRange(i + 1, 7).setValue('sent');
      Logger.log('✅ Reminder sent to ' + name);
    }
  }
}

// ============================================
// 5. WHATSAPP REMINDERS (Twilio)
// ============================================

function sendWhatsAppReminder(name, phone, reminderType) {
  if (!TWILIO_ACCOUNT_SID || TWILIO_ACCOUNT_SID.includes('YOUR_')) {
    Logger.log('⚠️ Twilio not configured');
    return;
  }
  
  const message = `Γεια σας ${name}! 👋\n\n⏰ Υπενθύμιση: Το ραντεβού σας με VELISSE είναι ${reminderType}.\n\n💇 Περιμένουμε να σας δούμε!\n🌟 VELISSE`;
  
  const payload = {
    From: TWILIO_PHONE,
    To: phone,
    Body: message,
    Channel: 'whatsapp'
  };
  
  try {
    const response = UrlFetchApp.fetch('https://api.twilio.com/2010-04-01/Accounts/' + TWILIO_ACCOUNT_SID + '/Messages.json', {
      method: 'post',
      headers: {
        Authorization: 'Basic ' + Utilities.base64Encode(TWILIO_ACCOUNT_SID + ':' + TWILIO_AUTH_TOKEN)
      },
      payload: payload,
      muteHttpExceptions: true
    });
    
    Logger.log('💬 WhatsApp sent to ' + phone);
  } catch (error) {
    Logger.log('WhatsApp error: ' + error);
  }
}

// ============================================
// 6. VIBER REMINDERS
// ============================================

function sendViberReminder(name, phone, reminderType) {
  if (!VIBER_API_TOKEN || VIBER_API_TOKEN.includes('YOUR_')) {
    Logger.log('⚠️ Viber not configured');
    return;
  }
  
  const message = `Γεια σας ${name}! 👋\n⏰ Υπενθύμιση: Το ραντεβού σας με VELISSE είναι ${reminderType}.\n💇 Περιμένουμε να σας δούμε!`;
  
  const payload = {
    receiver: phone,
    min_api_version: 1,
    sender: {
      name: 'VELISSE'
    },
    text: message
  };
  
  try {
    const response = UrlFetchApp.fetch('https://chatapi.viber.com/pa/send_message', {
      method: 'post',
      headers: {
        'X-Viber-Auth-Token': VIBER_API_TOKEN,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    Logger.log('📞 Viber sent to ' + phone);
  } catch (error) {
    Logger.log('Viber error: ' + error);
  }
}

// ============================================
// 7. EMAIL REMINDERS
// ============================================

function sendEmailReminder(email, name, reminderType) {
  const subject = 'Υπενθύμιση ραντεβού - VELISSE 💇';
  const emailBody = `
    Γεια σας ${name}! 👋
    
    ⏰ Υπενθύμιση: Το ραντεβού σας με VELISSE είναι ${reminderType}.
    
    💇 Περιμένουμε να σας δούμε!
    
    Ευχαριστούμε που επιλέγετε VELISSE ✨
  `;
  
  try {
    GmailApp.sendEmail(email, subject, emailBody);
    Logger.log('📧 Email sent to ' + email);
  } catch (error) {
    Logger.log('Email error: ' + error);
  }
}

// ============================================
// SETUP REMINDERS (Run manually once)
// ============================================

function setupRemindersTimer() {
  // Διαγράψτε παλιές triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  
  // Νέα trigger κάθε 5 λεπτά
  ScriptApp.newTrigger('checkAndSendReminders')
    .timeBased()
    .everyMinutes(5)
    .create();
  
  Logger.log('✅ Reminders timer set up!');
}
