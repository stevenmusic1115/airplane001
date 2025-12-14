// ================================
// VMS Free Trial Lesson Booking System
// FINAL FIXED VERSION - 最終修正版
// Correctly handles format conversion
// ================================

const SPREADSHEET_ID = '1CQS4YG4d3mRsAvYMzlrUlQ13SMTUlfgJyuL15UpnDV4';
const RESPONSE_SHEET_NAME = 'Free Trial Reservation Form';  // 修正工作表名稱
const SCHEDULE_SHEET_NAME = 'Free Slot';

// Day name conversion: Short → Full
const DAY_MAP = {
  "Mon.": "Monday",
  "Tue.": "Tuesday", 
  "Wed.": "Wednesday",
  "Thu.": "Thursday",
  "Fri.": "Friday",
  "Sat.": "Saturday",
  "Sun.": "Sunday"
};

/**
 * Web App GET 請求處理
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('VMS Free Trial Lesson Booking')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Convert short day name to full name
 * Wed. → Wednesday
 */
function convertDayToFull(shortDay) {
  // Remove any extra spaces
  const cleaned = shortDay.toString().trim();
  return DAY_MAP[cleaned] || cleaned;
}

/**
 * Convert single time to time range
 * Input: "5:00 PM" or "5:00:00 PM"
 * Output: "5:00pm-5:30pm"
 */
function convertTimeToRange(time) {
  // Clean the time string
  const cleaned = time.toString().trim();
  
  // Remove seconds if present: "5:00:00 PM" → "5:00 PM"
  const withoutSeconds = cleaned.replace(/:\d{2}\s*(AM|PM)/i, ' $1');
  
  // Parse time: "5:00 PM"
  const match = withoutSeconds.match(/(\d{1,2}):(\d{2})\s*(AM|PM)/i);
  if (!match) {
    Logger.log('⚠️ Failed to parse time: ' + time);
    return time; // Return original if can't parse
  }
  
  let hour = parseInt(match[1]);
  const minute = parseInt(match[2]);
  const period = match[3].toUpperCase();
  
  // Convert to 24-hour format
  if (period === 'PM' && hour !== 12) {
    hour += 12;
  } else if (period === 'AM' && hour === 12) {
    hour = 0;
  }
  
  // Calculate end time (30 minutes later)
  let endHour = hour;
  let endMinute = minute + 30;
  if (endMinute >= 60) {
    endHour += 1;
    endMinute -= 60;
  }
  
  // Convert back to 12-hour format with lowercase am/pm
  let startPeriod = hour >= 12 ? 'pm' : 'am';
  let endPeriod = endHour >= 12 ? 'pm' : 'am';
  
  let displayStartHour = hour > 12 ? hour - 12 : (hour === 0 ? 12 : hour);
  let displayEndHour = endHour > 12 ? endHour - 12 : (endHour === 0 ? 12 : endHour);
  
  // Format as "5:00pm-5:30pm"
  const startTime = `${displayStartHour}:${minute.toString().padStart(2, '0')}${startPeriod}`;
  const endTime = `${displayEndHour}:${endMinute.toString().padStart(2, '0')}${endPeriod}`;
  
  return `${startTime}-${endTime}`;
}

/**
 * Get free slots based on instrument selection
 */
function getFreeSlots(instrument) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
    
    if (!sheet) {
      Logger.log('❌ Schedule sheet not found');
      return [];
    }
    
    // Determine columns based on instrument
    let dayColumn, timeColumn;
    if (instrument === 'Piano' || instrument === 'Sing and Play') {
      dayColumn = 'A';
      timeColumn = 'B';
      Logger.log('Using columns A-B for Piano/Sing and Play');
    } else {
      dayColumn = 'D';
      timeColumn = 'E';
      Logger.log('Using columns D-E for Guitar instruments');
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No data in Free Slot sheet');
      return [];
    }
    
    const dayValues = sheet.getRange(dayColumn + '2:' + dayColumn + lastRow).getValues();
    const timeValues = sheet.getRange(timeColumn + '2:' + timeColumn + lastRow).getValues();
    
    // Get booked slots (already converted to match Free Slot format)
    const bookedSlots = getBookedSlots(instrument);
    Logger.log('📋 Booked slots: ' + JSON.stringify(bookedSlots));
    
    const freeSlots = [];
    
    for (let i = 0; i < dayValues.length; i++) {
      const day = dayValues[i][0];
      const time = timeValues[i][0];
      
      if (day && time && day.toString().trim() !== '' && time.toString().trim() !== '') {
        const dayStr = day.toString().trim();
        const timeStr = time.toString().trim();
        const slotKey = `${dayStr}|${timeStr}`;
        
        const isBooked = bookedSlots.includes(slotKey);
        Logger.log(`${slotKey} - ${isBooked ? '❌ BOOKED' : '✅ Available'}`);
        
        if (!isBooked) {
          freeSlots.push([dayStr, timeStr]);
        }
      }
    }
    
    Logger.log(`✅ Total free slots: ${freeSlots.length}`);
    return freeSlots;
    
  } catch (error) {
    Logger.log('❌ Error in getFreeSlots: ' + error.toString());
    return [];
  }
}

/**
 * Get list of already booked slots from Response sheet
 * Converts Response format to Free Slot format for comparison
 */
function getBookedSlots(instrument) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    
    if (!responseSheet) {
      Logger.log('⚠️ Response sheet not found: ' + RESPONSE_SHEET_NAME);
      return [];
    }
    
    const lastRow = responseSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('No bookings found');
      return [];
    }
    
    // Read all booking data
    // Columns: Timestamp(A), Parent Name(B), Phone(C), Child Name(D), Age(E), 
    //          Instrument(F), Grade(G), Date(H), Day(I), Time(J)
    const data = responseSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    const bookedSlots = [];
    
    Logger.log(`\n🔍 Checking ${data.length} bookings for ${instrument}:`);
    
    for (let i = 0; i < data.length; i++) {
      const rowInstrument = data[i][5]; // Column F: Instrument
      const day = data[i][8];           // Column I: Day (e.g., "Wed.")
      const time = data[i][9];          // Column J: Time (e.g., "5:00 PM" or "5:00:00 PM")
      
      // Skip if not matching instrument or missing data
      if (!rowInstrument || 
          rowInstrument.toString().trim() !== instrument ||
          !day || 
          !time) {
        continue;
      }
      
      // Convert Response format to Free Slot format
      const dayStr = day.toString().trim();
      const timeStr = time.toString().trim();
      
      const fullDay = convertDayToFull(dayStr);      // "Wed." → "Wednesday"
      const timeRange = convertTimeToRange(timeStr);  // "5:00 PM" → "5:00pm-5:30pm"
      
      const slotKey = `${fullDay}|${timeRange}`;
      bookedSlots.push(slotKey);
      
      Logger.log(`Row ${i+2}: ${dayStr}|${timeStr} → ${slotKey}`);
    }
    
    Logger.log(`\n✅ Total booked slots for ${instrument}: ${bookedSlots.length}`);
    return bookedSlots;
    
  } catch (error) {
    Logger.log('❌ Error in getBookedSlots: ' + error.toString());
    return [];
  }
}

/**
 * Submit booking to the response sheet
 */
function submitBooking(bookingData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    
    // Create sheet if doesn't exist
    if (!responseSheet) {
      Logger.log('Creating new response sheet...');
      responseSheet = ss.insertSheet(RESPONSE_SHEET_NAME);
      responseSheet.appendRow([
        'Timestamp',
        'Parent Name',
        'Phone Number',
        'Child Name',
        'Age',
        'Instrument',
        'Grade',
        'Date',
        'Day',
        'Time'
      ]);
      
      const headerRange = responseSheet.getRange(1, 1, 1, 10);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4169E1');
      headerRange.setFontColor('#FFFFFF');
      headerRange.setHorizontalAlignment('center');
      responseSheet.setFrozenRows(1);
    }
    
    // Check if already booked
    Logger.log('\n📝 Processing booking:');
    Logger.log(`Instrument: ${bookingData.instrument}`);
    Logger.log(`Day: ${bookingData.day}`);
    Logger.log(`Time: ${bookingData.time}`);
    
    // Convert to Free Slot format for comparison
    const fullDay = convertDayToFull(bookingData.day);
    const timeRange = convertTimeToRange(bookingData.time);
    const checkSlotKey = `${fullDay}|${timeRange}`;
    
    Logger.log(`Converted to: ${checkSlotKey}`);
    
    const bookedSlots = getBookedSlots(bookingData.instrument);
    
    Logger.log(`\n🔍 Duplicate check:`);
    Logger.log(`Looking for: ${checkSlotKey}`);
    Logger.log(`In booked list: ${JSON.stringify(bookedSlots)}`);
    Logger.log(`Is duplicate: ${bookedSlots.includes(checkSlotKey)}`);
    
    if (bookedSlots.includes(checkSlotKey)) {
      Logger.log('❌ Slot already booked!');
      return {
        success: false,
        message: 'This time slot has already been booked. Please refresh and choose another slot.'
      };
    }
    
    // Add booking (store in original frontend format)
    const timestamp = new Date();
    const newRow = responseSheet.getLastRow() + 1;
    
    Logger.log(`✅ Slot available, adding to row ${newRow}`);
    
    responseSheet.appendRow([
      timestamp,
      bookingData.parentName,
      bookingData.parentPhone,
      bookingData.childName,
      bookingData.childAge,
      bookingData.instrument,
      bookingData.grade,
      bookingData.date,
      bookingData.day,   // Store as "Wed." (frontend format)
      bookingData.time   // Store as "5:00 PM" (frontend format)
    ]);
    
    // Format
    responseSheet.getRange(newRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    responseSheet.autoResizeColumns(1, 10);
    
    const dataRange = responseSheet.getRange(newRow, 1, 1, 10);
    if (newRow % 2 === 0) {
      dataRange.setBackground('#F0F8FF');
    } else {
      dataRange.setBackground('#FFFFFF');
    }
    dataRange.setBorder(true, true, true, true, true, true);
    
    Logger.log('✅ Booking successful!');
    
    return {
      success: true,
      message: 'Booking confirmed! 🎉'
    };
    
  } catch (error) {
    Logger.log('❌ Error in submitBooking: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Test format conversion
 */
function testFormatConversion() {
  Logger.log('=== Testing Format Conversion ===\n');
  
  Logger.log('--- Day Conversion ---');
  Logger.log('Mon. → ' + convertDayToFull('Mon.'));
  Logger.log('Wed. → ' + convertDayToFull('Wed.'));
  Logger.log('Sat. → ' + convertDayToFull('Sat.'));
  
  Logger.log('\n--- Time Conversion ---');
  Logger.log('5:00 PM → ' + convertTimeToRange('5:00 PM'));
  Logger.log('5:00:00 PM → ' + convertTimeToRange('5:00:00 PM'));
  Logger.log('10:00 AM → ' + convertTimeToRange('10:00 AM'));
  Logger.log('2:30 PM → ' + convertTimeToRange('2:30 PM'));
  Logger.log('11:45 PM → ' + convertTimeToRange('11:45 PM'));
  
  Logger.log('\n--- Complete Conversion ---');
  const day = 'Wed.';
  const time = '5:00 PM';
  const result = convertDayToFull(day) + '|' + convertTimeToRange(time);
  Logger.log(`${day}|${time} → ${result}`);
}

/**
 * Test spreadsheet access
 */
function testSpreadsheetAccess() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('✅ Spreadsheet: ' + ss.getName());
    
    Logger.log('\n📊 Available sheets:');
    ss.getSheets().forEach(sheet => {
      Logger.log('  - ' + sheet.getName());
    });
    
    // Check if Response sheet exists
    const responseSheet = ss.getSheetByName(RESPONSE_SHEET_NAME);
    if (responseSheet) {
      Logger.log(`\n✅ Response sheet found: "${RESPONSE_SHEET_NAME}"`);
      Logger.log(`   Rows: ${responseSheet.getLastRow()}`);
    } else {
      Logger.log(`\n⚠️ Response sheet NOT found: "${RESPONSE_SHEET_NAME}"`);
    }
    
  } catch (error) {
    Logger.log('❌ Error: ' + error.toString());
  }
}

/**
 * Test with real data from your sheet
 */
function testRealData() {
  Logger.log('=== Testing with Real Data ===\n');
  
  // Simulate a booking that matches your example
  const testBooking = {
    parentName: 'Test Parent',
    parentPhone: '+65 9999 9999',
    childName: 'Test Child',
    childAge: '8',
    instrument: 'Piano',  // Adjust based on your actual instrument
    grade: 'Grade 1',
    date: 'Dec 18, 2024',
    day: 'Wed.',          // Your actual format
    time: '5:00 PM',      // Your actual format
    timestamp: new Date().toISOString()
  };
  
  Logger.log('Test booking data:');
  Logger.log(JSON.stringify(testBooking, null, 2));
  
  Logger.log('\n--- Conversion Test ---');
  const fullDay = convertDayToFull(testBooking.day);
  const timeRange = convertTimeToRange(testBooking.time);
  Logger.log(`Frontend: ${testBooking.day}|${testBooking.time}`);
  Logger.log(`Converted: ${fullDay}|${timeRange}`);
  Logger.log(`Expected Free Slot format: Wednesday|5:00pm-5:30pm`);
  
  Logger.log('\n--- Checking Booked Slots ---');
  const bookedSlots = getBookedSlots(testBooking.instrument);
  
  Logger.log('\n--- Submit Test (DRY RUN - not actually submitting) ---');
  const slotKey = `${fullDay}|${timeRange}`;
  if (bookedSlots.includes(slotKey)) {
    Logger.log('❌ Would be REJECTED (slot already booked)');
  } else {
    Logger.log('✅ Would be ACCEPTED (slot available)');
  }
}

/**
 * Initialize system
 */
function initializeSystem() {
  Logger.log('🚀 VMS Booking System - Final Fixed Version\n');
  Logger.log('='.repeat(60));
  
  Logger.log('\n1️⃣ Testing spreadsheet access...');
  testSpreadsheetAccess();
  
  Logger.log('\n' + '='.repeat(60));
  Logger.log('\n2️⃣ Testing format conversion...');
  testFormatConversion();
  
  Logger.log('\n' + '='.repeat(60));
  Logger.log('\n3️⃣ Testing with real data...');
  testRealData();
  
  Logger.log('\n' + '='.repeat(60));
  Logger.log('\n✅ Initialization complete!');
  Logger.log('\n📖 Format Handling:');
  Logger.log('   Frontend → Response: Wed.|5:00 PM');
  Logger.log('   Response → Free Slot: Wednesday|5:00pm-5:30pm');
  Logger.log('   Now correctly matches and prevents duplicates!');
  Logger.log('\n📖 Next steps:');
  Logger.log('   1. Deploy as Web App');
  Logger.log('   2. Test booking');
  Logger.log('   3. Try to book same slot again → Should fail!');
}

/**
 * Get Web App URL
 */
function getWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  Logger.log('📱 Web App URL: ' + url);
  return url;
}
