// Google Apps Script - Create VELISSE Booking System from scratch
// Τρέξτε αυτή τη συνάρτηση μόνο μια φορά: setupVelisseBookingSystem()

function setupVelisseBookingSystem() {
  // Δημιουργία νέου Google Sheet
  const spreadsheet = SpreadsheetApp.create('VELISSE Booking System');
  const spreadsheetId = spreadsheet.getId();
  
  Logger.log('📊 Google Sheet δημιουργήθηκε: ' + spreadsheetId);
  
  // Διαγραφή του default sheet
  const defaultSheet = spreadsheet.getSheets()[0];
  spreadsheet.deleteSheet(defaultSheet);
  
  // ============================================
  // TAB 1: SERVICES (Υπηρεσίες)
  // ============================================
  const servicesSheet = spreadsheet.insertSheet('Services', 0);
  servicesSheet.appendRow(['Υπηρεσία', 'Διάρκεια (λεπτά)', 'Τιμή (€)']);
  servicesSheet.appendRow(['One by One', 120, 45]);
  servicesSheet.appendRow(['Half Set', 90, 35]);
  servicesSheet.appendRow(['Volume 2D-3D', 150, 50]);
  servicesSheet.appendRow(['Special Effects', 150, 60]);
  servicesSheet.appendRow(['Lash & Brow Lift', 50, 40]);
  servicesSheet.appendRow(['Καθαρισμό Φρυδιών', 30, 20]);
  
  // Format: Bold header
  const servicesRange = servicesSheet.getRange(1, 1, 1, 3);
  servicesRange.setFontWeight('bold');
  servicesRange.setBackground('#1a1a1a');
  servicesRange.setFontColor('#ffffff');
  servicesSheet.setColumnWidths([1, 20, 20, 15]);
  
  Logger.log('✅ Tab 1: Services δημιουργήθηκε');
  
  // ============================================
  // TAB 2: PENDING (Αναμονή Επιβεβαίωσης)
  // ============================================
  const pendingSheet = spreadsheet.insertSheet('Pending', 1);
  pendingSheet.appendRow([
    'ID', 'Όνομα', 'Τηλέφωνο', 'Email', 
    'Ημερομηνία', 'Ώρα Αρχής', 'Ώρα Λήξης', 
    'Υπηρεσία', 'Διάρκεια (λεπτά)', 'Τιμή (€)', 
    'Κατάσταση', 'Δημιουργήθηκε'
  ]);
  
  const pendingRange = pendingSheet.getRange(1, 1, 1, 12);
  pendingRange.setFontWeight('bold');
  pendingRange.setBackground('#FFC107');
  pendingRange.setFontColor('#000000');
  pendingSheet.setColumnWidths([1, 12, 12, 15, 15, 12, 12, 15, 15, 10, 12, 15]);
  
  Logger.log('✅ Tab 2: Pending δημιουργήθηκε');
  
  // ============================================
  // TAB 3: BOOKINGS (Επιβεβαιωμένα Ραντεβού)
  // ============================================
  const bookingsSheet = spreadsheet.insertSheet('Bookings', 2);
  bookingsSheet.appendRow([
    'ID', 'Όνομα', 'Τηλέφωνο', 'Email', 
    'Ημερομηνία', 'Ώρα Αρχής', 'Ώρα Λήξης', 
    'Υπηρεσία', 'Διάρκεια (λεπτά)', 'Τιμή (€)', 
    'Κατάσταση', 'Επιβεβαιώθηκε'
  ]);
  
  const bookingsRange = bookingsSheet.getRange(1, 1, 1, 12);
  bookingsRange.setFontWeight('bold');
  bookingsRange.setBackground('#4CAF50');
  bookingsRange.setFontColor('#ffffff');
  bookingsSheet.setColumnWidths([1, 12, 12, 15, 15, 12, 12, 15, 15, 10, 12, 15]);
  
  Logger.log('✅ Tab 3: Bookings δημιουργήθηκε');
  
  // ============================================
  // TAB 4: AVAILABILITY (Διαθεσιμότητα)
  // ============================================
  const availabilitySheet = spreadsheet.insertSheet('Availability', 3);
  availabilitySheet.appendRow(['Ημερομηνία', 'Ώρα', 'Διαθέσιμο', 'Υπηρεσία', 'Πελάτης']);
  
  const availabilityRange = availabilitySheet.getRange(1, 1, 1, 5);
  availabilityRange.setFontWeight('bold');
  availabilityRange.setBackground('#2196F3');
  availabilityRange.setFontColor('#ffffff');
  availabilitySheet.setColumnWidths([1, 15, 12, 12, 20, 20]);
  
  Logger.log('✅ Tab 4: Availability δημιουργήθηκε');
  
  // ============================================
  // TAB 5: CANCELLED (Ακυρωμένα)
  // ============================================
  const cancelledSheet = spreadsheet.insertSheet('Cancelled', 4);
  cancelledSheet.appendRow([
    'ID', 'Όνομα', 'Τηλέφωνο', 'Email', 
    'Ημερομηνία', 'Ώρα Αρχής', 'Ώρα Λήξης', 
    'Υπηρεσία', 'Διάρκεια (λεπτά)', 'Τιμή (€)', 
    'Κατάσταση', 'Ακυρώθηκε'
  ]);
  
  const cancelledRange = cancelledSheet.getRange(1, 1, 1, 12);
  cancelledRange.setFontWeight('bold');
  cancelledRange.setBackground('#F44336');
  cancelledRange.setFontColor('#ffffff');
  cancelledSheet.setColumnWidths([1, 12, 12, 15, 15, 12, 12, 15, 15, 10, 12, 15]);
  
  Logger.log('✅ Tab 5: Cancelled δημιουργήθηκε');
  
  // ============================================
  // Τελική Ενημέρωση
  // ============================================
  Logger.log('🎉 VELISSE Booking System είναι έτοιμο!');
  Logger.log('📊 Sheet ID: ' + spreadsheetId);
  Logger.log('🔗 Link: https://docs.google.com/spreadsheets/d/' + spreadsheetId);
  
  // Αντιγραφή του ID στα clipboard
  SpreadsheetApp.getUi().alert(
    '✅ Google Sheet δημιουργήθηκε!\n\n' +
    'Sheet ID: ' + spreadsheetId + '\n\n' +
    'Ανοίξτε: https://docs.google.com/spreadsheets/d/' + spreadsheetId
  );
}

// ============================================
// HELPER FUNCTIONS
// ============================================

function getServiceDuration(serviceName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const servicesSheet = ss.getSheetByName('Services');
  const data = servicesSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === serviceName) {
      return data[i][1]; // Διάρκεια σε λεπτά
    }
  }
  return 60; // Default
}

function getServicePrice(serviceName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const servicesSheet = ss.getSheetByName('Services');
  const data = servicesSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === serviceName) {
      return data[i][2]; // Τιμή
    }
  }
  return 0;
}

function calculateEndTime(startTime, durationMinutes) {
  const [hours, minutes] = startTime.split(':').map(Number);
  const totalMinutes = hours * 60 + minutes + durationMinutes;
  const endHours = Math.floor(totalMinutes / 60) % 24;
  const endMinutes = totalMinutes % 60;
  return `${String(endHours).padStart(2, '0')}:${String(endMinutes).padStart(2, '0')}`;
}

function blockAvailability(date, startTime, durationMinutes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const availabilitySheet = ss.getSheetByName('Availability');
  const endTime = calculateEndTime(startTime, durationMinutes);
  
  // Generate all 30-minute slots
  const slots = generateTimeSlots(startTime, endTime);
  
  for (const slot of slots) {
    availabilitySheet.appendRow([date, slot, 'FALSE', '', '']);
  }
}

function generateTimeSlots(startTime, endTime) {
  const slots = [];
  const [startHours, startMinutes] = startTime.split(':').map(Number);
  const [endHours, endMinutes] = endTime.split(':').map(Number);
  
  let currentMinutes = startHours * 60 + startMinutes;
  const endTotalMinutes = endHours * 60 + endMinutes;
  
  while (currentMinutes < endTotalMinutes) {
    const hours = Math.floor(currentMinutes / 60);
    const minutes = currentMinutes % 60;
    slots.push(`${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`);
    currentMinutes += 30; // 30-minute intervals
  }
  
  return slots;
}
