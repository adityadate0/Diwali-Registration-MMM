// ============================================================================
// CONFIGURATION & SAFETY SETTINGS
// ============================================================================
const SHEET_NAME = 'Guest List & Check-in';
const STARTING_SEAT_NUMBER = 1;

// 🛡️ DEV / TESTING SAFETY SWITCH:
// Set DEV_MODE = true during testing. All emails route exclusively to DEV_TEST_EMAIL.
// Set DEV_MODE = false on the live event day.
const DEV_MODE = true;
const DEV_TEST_EMAIL = 'sebastianschnell54@gmail.com';

const MMM_LOGO_URL = 'https://mmmunich.com/wp-content/uploads/2019/03/mmmunich-lg.png';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('MMM Diwali Check-in Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Loads directory from the 16-column guest list with resilient error handling.
 */
function getGuestDirectory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`Sheet tab "${SHEET_NAME}" not found.`);

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const rows = data.slice(1);
  return rows.map((r, index) => {
    const rowNum = index + 2;

    // Resilient timestamp parsing (handles Date objects, ISO strings, and raw times)
    let timeStr = '';
    if (r[13]) {
      if (r[13] instanceof Date) {
        try {
          timeStr = Utilities.formatDate(r[13], 'Europe/Berlin', 'HH:mm:ss');
        } catch (e) {
          timeStr = String(r[13]);
        }
      } else {
        const s = String(r[13]).trim();
        const parts = s.split(' ');
        timeStr = parts.length > 1 ? parts[1] : s;
      }
    }

    return {
      rowIndex: rowNum,
      memberId: String(r[0] || '').trim(),
      firstName: String(r[1] || '').trim(),
      lastName: String(r[2] || '').trim(),
      fullName: String(r[3] || '').trim(),
      email: String(r[4] || '').trim(),
      adults: Number(r[5] || 0),
      kidsAbove12: Number(r[6] || 0),
      kids6to12: Number(r[7] || 0),
      kidsBelow6: Number(r[8] || 0),
      headcount: Number(r[9] || 0),
      foodCoupons: Number(r[10] || 0),
      seats: String(r[11] || '').trim(),
      status: String(r[12] || 'Pending').trim(),
      timestamp: timeStr,
      desk: String(r[14] || '').trim(),
      notes: String(r[15] || '').trim()
    };
  });
}

/**
 * Checks in guest, assigns seats, and dispatches confirmation email.
 */
function checkInAndAssignSeats(rowIndex, deskId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy handling another desk. Please tap Check In again.' };
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const targetRow = parseInt(rowIndex, 10);
    
    const guestData = sheet.getRange(targetRow, 1, 1, 16).getValues()[0];
    if (!guestData || !guestData[3]) {
      return { success: false, message: `Attendee row #${targetRow} was not found.` };
    }

    let assignedSeats = String(guestData[11] || '').trim();
    const isAlreadyCheckedIn = (String(guestData[12]).trim().toLowerCase() === 'checked in');

    // Dynamic contiguous seat allocation
    if (!assignedSeats) {
      let maxSeat = STARTING_SEAT_NUMBER - 1;
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const seatColValues = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
        for (let i = 0; i < seatColValues.length; i++) {
          const matches = String(seatColValues[i][0]).match(/\d+/g);
          if (matches) {
            matches.forEach(m => {
              const num = parseInt(m, 10);
              if (num > maxSeat) maxSeat = num;
            });
          }
        }
      }

      const headcount = parseInt(guestData[9], 10) || 1;
      const startSeat = maxSeat + 1;
      const endSeat = maxSeat + headcount;
      assignedSeats = (startSeat === endSeat) ? `Seat ${startSeat}` : `Seats ${startSeat} - ${endSeat}`;
    }

    const now = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd HH:mm:ss');

    sheet.getRange(targetRow, 12).setValue(assignedSeats);
    sheet.getRange(targetRow, 13).setValue('Checked In');
    sheet.getRange(targetRow, 14).setValue(now);
    sheet.getRange(targetRow, 15).setValue(deskId);

    let targetEmail = DEV_MODE ? DEV_TEST_EMAIL : String(guestData[4] || '').trim();
    let emailStatus = 'No email on record';

    if (targetEmail && targetEmail.includes('@')) {
      try {
        GmailApp.sendEmail(
          targetEmail,
          `MMM Diwali 2026: Check-In Confirmed (${assignedSeats})`,
          '',
          {
            name: 'Maharashtra Mandal Munich',
            htmlBody: buildCheckInEmailHtml(
              guestData[1], 
              assignedSeats, 
              guestData[0] === 'Non Mbr' ? 'Non-Member' : guestData[0], 
              guestData[9], 
              guestData[5], 
              guestData[6], 
              guestData[7], 
              guestData[8], 
              guestData[10], 
              deskId
            )
          }
        );
        emailStatus = `Check-in confirmation delivered to ${targetEmail}`;
      } catch (err) {
        emailStatus = `Email error: ${err.message}`;
      }
    }

    return {
      success: true,
      rowIndex: targetRow,
      bookingId: (guestData[0] === 'Non Mbr') ? 'Non-Member' : guestData[0],
      fullName: guestData[3],
      seats: assignedSeats,
      headcount: guestData[9],
      adults: guestData[5],
      kids12: guestData[6],
      kids6to12: guestData[7],
      kids6: guestData[8],
      foodCoupons: guestData[10],
      emailStatus: emailStatus,
      isRecheck: isAlreadyCheckedIn
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * Registers an on-the-spot walk-in attendee.
 */
function registerSpotWalkIn(guestData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (e) {
    return { success: false, message: 'Server busy. Please try again.' };
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    const lastRow = sheet.getLastRow();
    const walkInId = 'WALK-' + ('000' + lastRow).slice(-3);

    let maxSeat = STARTING_SEAT_NUMBER - 1;
    if (lastRow > 1) {
      const seatColValues = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
      for (let i = 0; i < seatColValues.length; i++) {
        const matches = String(seatColValues[i][0]).match(/\d+/g);
        if (matches) {
          matches.forEach(m => {
            const num = parseInt(m, 10);
            if (num > maxSeat) maxSeat = num;
          });
        }
      }
    }

    const adults = parseInt(guestData.adults, 10) || 0;
    const kids12 = parseInt(guestData.kidsAbove12, 10) || 0;
    const kids6to12 = parseInt(guestData.kids6to12, 10) || 0;
    const kids6 = parseInt(guestData.kidsBelow6, 10) || 0;
    const totalCount = adults + kids12 + kids6to12 + kids6;
    const foodCoupons = adults + kids12 + kids6to12;

    const startSeat = maxSeat + 1;
    const endSeat = maxSeat + (totalCount > 0 ? totalCount : 1);
    const assignedSeats = (startSeat === endSeat) ? `Seat ${startSeat}` : `Seats ${startSeat} - ${endSeat}`;
    const now = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd HH:mm:ss');

    const newRow = [
      walkInId,
      guestData.firstName,
      guestData.lastName,
      `${guestData.firstName} ${guestData.lastName}`.trim(),
      guestData.email,
      adults,
      kids12,
      kids6to12,
      kids6,
      totalCount,
      foodCoupons,
      assignedSeats,
      'Checked In',
      now,
      guestData.deskId || 'Desk 1',
      guestData.notes || 'Spot Walk-in'
    ];

    sheet.appendRow(newRow);

    let targetEmail = DEV_MODE ? DEV_TEST_EMAIL : guestData.email;
    let emailStatus = 'No email provided';

    if (targetEmail && targetEmail.includes('@')) {
      try {
        GmailApp.sendEmail(
          targetEmail,
          `MMM Diwali 2026: Check-In Confirmed (${assignedSeats})`,
          '',
          {
            name: 'Maharashtra Mandal Munich',
            htmlBody: buildCheckInEmailHtml(guestData.firstName, assignedSeats, walkInId, totalCount, adults, kids12, kids6to12, kids6, foodCoupons, guestData.deskId)
          }
        );
        emailStatus = `Check-in confirmation delivered to ${targetEmail}`;
      } catch (err) {
        emailStatus = `Email error: ${err.message}`;
      }
    }

    return {
      success: true,
      rowIndex: lastRow + 1,
      bookingId: walkInId,
      fullName: `${guestData.firstName} ${guestData.lastName}`.trim(),
      seats: assignedSeats,
      headcount: totalCount,
      adults: adults,
      kids12: kids12,
      kids6to12: kids6to12,
      kids6: kids6,
      foodCoupons: foodCoupons,
      emailStatus: emailStatus,
      isRecheck: false
    };

  } finally {
    lock.releaseLock();
  }
}

/**
 * Builds email HTML template.
 */
function buildCheckInEmailHtml(name, seats, bookingId, headcount, adults, kids12, kids6to12, kids6, foodCoupons, deskId) {
  const totalKids = (Number(kids12) || 0) + (Number(kids6to12) || 0) + (Number(kids6) || 0);
  return `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 520px; margin: auto; border: 1px solid #eee; border-radius: 14px; overflow: hidden; background-color: #ffffff;">
      <div style="background: linear-gradient(135deg, #b02a37, #e06d28); padding: 26px 20px 22px 20px; text-align: center; color: #ffffff;">
        <div style="margin-bottom: 12px;">
          <img src="${MMM_LOGO_URL}" alt="Maharashtra Mandal Munich" style="height: 60px; width: auto; background-color: rgba(255,255,255,0.95); padding: 6px 12px; border-radius: 8px;" />
        </div>
        <h2 style="margin: 0; font-size: 23px; font-weight: 700; color: #ffffff;">Welcome to Diwali 2026</h2>
        <p style="margin: 5px 0 0 0; font-size: 14px; color: #ffffff; opacity: 0.95;">Maharashtra Mandal Munich e.V.</p>
      </div>

      <div style="padding: 24px; color: #333333; line-height: 1.5;">
        <p style="font-size: 16px; margin-top: 0;">Namaskar <strong>${name || 'Guest'}</strong>,</p>
        <p>You have checked in at <strong>${deskId}</strong>. Here are your seating details:</p>
        
        <div style="background: #fff8eb; border: 1px solid #ffd8a8; border-left: 6px solid #e06d28; padding: 16px; margin: 20px 0; border-radius: 8px;">
          <span style="font-size: 12px; font-weight: 700; color: #8c4109; text-transform: uppercase;">Your Assigned Seating</span>
          <div style="font-size: 28px; font-weight: 800; color: #b02a37; margin-top: 2px;">${seats}</div>
        </div>

        <table style="width: 100%; border-collapse: collapse; font-size: 15px; margin-bottom: 20px;">
          <tr style="border-bottom: 1px solid #f0f0f0;">
            <td style="padding: 10px 0; color: #666;">Booking ID:</td>
            <td style="padding: 10px 0; font-weight: 600; text-align: right;">${bookingId}</td>
          </tr>
          <tr style="border-bottom: 1px solid #f0f0f0;">
            <td style="padding: 10px 0; color: #666;">Total Headcount:</td>
            <td style="padding: 10px 0; font-weight: 600; text-align: right;">${headcount} (${adults} Adults, ${totalKids} Kids)</td>
          </tr>
          <tr>
            <td style="padding: 10px 0; color: #666;">Dinner Food Coupons:</td>
            <td style="padding: 10px 0; font-weight: 600; text-align: right; color: #b02a37;">${foodCoupons} Coupons</td>
          </tr>
        </table>

        <p style="font-size: 14px; color: #666;">
          Please show this email if you need assistance reaching your table. Our floor volunteers will guide you.
        </p>

        <p style="margin-top: 24px; font-size: 15px; font-weight: bold; color: #b02a37; text-align: center;">
          Wishing you and your family a Joyous Diwali!
        </p>
      </div>
    </div>
  `;
}
