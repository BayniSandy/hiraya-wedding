/**
 * ============================================
 * HIRAYA WEDDING — Google Apps Script Backend
 * Rodnel & Divine · October 15, 2027
 * ============================================
 *
 * This script powers the wedding automation system:
 * - RSVP form processing
 * - Guest lookup & validation
 * - QR code generation
 * - Email notifications
 * - Admin API endpoints
 *
 * SETUP:
 * 1. Create a Google Sheet with sheets: GUEST_LIST, RSVP_LOG, TABLE_PLAN, EVENT_CONFIG, EMAIL_LOG
 * 2. Open Extensions > Apps Script
 * 3. Paste this code
 * 4. Deploy as Web App (Execute as: Me, Access: Anyone)
 * 5. Copy the Web App URL → use as API endpoint
 */

// ============ CONFIG ============
const SHEET_ID = '15M0Unr6I2p3NZgzFy2rGPg0O9DtGQGGzLaXH-3Pjp5c';
const ADMIN_KEY = 'hiraya-rd-2027-admin';

function getSpreadsheet() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

function getConfig(key) {
  const sheet = getSheet('EVENT_CONFIG');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1];
  }
  return null;
}

// ============ WEB APP ENTRY POINTS ============

function doGet(e) {
  const action = e.parameter.action || '';
  const cors = ContentService.createTextOutput('');

  try {
    let result;
    switch (action) {
      case 'lookup':
        result = lookupGuest(e.parameter);
        break;
      case 'tablemates':
        result = getTablemates(e.parameter);
        break;
      case 'config':
        result = getPublicConfig();
        break;
      case 'admin_guests':
        result = adminGetGuests(e.parameter);
        break;
      case 'admin_stats':
        result = adminGetStats(e.parameter);
        break;
      case 'admin_checkin_lookup':
        result = adminCheckinLookup(e.parameter);
        break;
      default:
        result = { error: 'Invalid action', status: 400 };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message, status: 500 });
  }
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || '';

    let result;
    switch (action) {
      case 'rsvp':
        result = processRsvp(payload);
        break;
      case 'admin_add_guest':
        result = adminAddGuest(payload);
        break;
      case 'admin_update_guest':
        result = adminUpdateGuest(payload);
        break;
      case 'admin_delete_guest':
        result = adminDeleteGuest(payload);
        break;
      case 'admin_assign_table':
        result = adminAssignTable(payload);
        break;
      case 'admin_reopen':
        result = adminReopenInvite(payload);
        break;
      case 'admin_send_invite':
        result = adminSendInvite(payload);
        break;
      case 'admin_checkin':
        result = adminCheckinGuest(payload);
        break;
      default:
        result = { error: 'Invalid action', status: 400 };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ error: err.message, status: 500 });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============ GUEST LOOKUP ============

/**
 * Lookup a guest by invite_code (from the invitation URL)
 * Returns: guest name, greeting, status (for the invitation page)
 */
function lookupGuest(params) {
  const code = (params.code || '').trim();
  if (!code) return { error: 'Walang invite code', status: 400 };

  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const codeCol = headers.indexOf('invite_code');

  for (let i = 1; i < data.length; i++) {
    if (data[i][codeCol] === code) {
      const row = rowToObj(headers, data[i]);

      // Check if expired
      if (row.invite_status === 'expired') {
        return {
          error: 'Expired na ang invitation code na ito. Makipag-ugnayan kay Divine o Rodnel.',
          status: 403,
          expired: true
        };
      }

      // Check expiry date
      if (row.invite_expiry && new Date() > new Date(row.invite_expiry)) {
        // Auto-expire
        updateGuestField(i + 1, headers, 'invite_status', 'expired');
        return {
          error: 'Expired na ang invitation code na ito. Makipag-ugnayan kay Divine o Rodnel.',
          status: 403,
          expired: true
        };
      }

      // Check RSVP deadline
      const deadline = getConfig('rsvp_deadline');
      if (deadline && new Date() > new Date(deadline)) {
        return {
          error: 'Lumipas na ang RSVP deadline (Setyembre 4, 2027).',
          status: 403,
          past_deadline: true
        };
      }

      return {
        status: 200,
        guest: {
          guest_id: row.guest_id,
          full_name: row.full_name,
          invite_status: row.invite_status,
          already_confirmed: row.invite_status === 'confirmed',
          already_declined: row.invite_status === 'declined',
        }
      };
    }
  }

  return { error: 'Hindi makita ang invite code na ito.', status: 404 };
}

// ============ RSVP PROCESSING ============

/**
 * Process an RSVP submission
 * payload: { code, email, full_name, attendance, proxy_name, special_needs }
 */
function processRsvp(payload) {
  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const codeCol = headers.indexOf('invite_code');

  const code = (payload.code || '').trim();
  if (!code) return { error: 'Walang invite code', status: 400 };

  // Find guest row
  let guestRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][codeCol] === code) {
      guestRow = i;
      break;
    }
  }

  if (guestRow === -1) return { error: 'Hindi makita ang invite code.', status: 404 };

  const guest = rowToObj(headers, data[guestRow]);

  // Validate: check if already responded
  if (['confirmed', 'declined', 'proxy'].includes(guest.invite_status)) {
    return {
      error: 'Nakapag-respond ka na dati. Kung gusto mong baguhin, makipag-ugnayan kay Divine o Rodnel.',
      status: 409,
      already_responded: true
    };
  }

  // Check deadline
  const deadline = getConfig('rsvp_deadline');
  if (deadline && new Date() > new Date(deadline)) {
    return { error: 'Lumipas na ang RSVP deadline.', status: 403 };
  }

  // Check expiry
  if (guest.invite_expiry && new Date() > new Date(guest.invite_expiry)) {
    updateGuestField(guestRow + 1, headers, 'invite_status', 'expired');
    return { error: 'Expired na ang invite code na ito.', status: 403 };
  }

  // Validate full name matches (case-insensitive)
  const submittedName = (payload.full_name || '').trim().toLowerCase();
  const guestName = (guest.full_name || '').trim().toLowerCase();
  if (submittedName !== guestName) {
    return {
      error: 'Hindi tugma ang pangalan. Ilagay ang inyong buong pangalan gaya ng nasa imbitasyon.',
      status: 400,
      name_mismatch: true
    };
  }

  // Validate email
  const email = (payload.email || '').trim();
  if (!email || !email.includes('@')) {
    return { error: 'Kailangan ng valid na email address.', status: 400 };
  }

  // Process based on attendance
  const attendance = payload.attendance; // 'attending', 'decline', 'proxy'
  const now = new Date();
  const sheetRow = guestRow + 1; // 1-indexed for sheet

  if (attendance === 'attending') {
    updateGuestFields(sheetRow, headers, {
      email: email,
      invite_status: 'confirmed',
      rsvp_date: now,
      special_needs: payload.special_needs ? JSON.stringify(payload.special_needs) : '',
      updated_at: now
    });

    // Log RSVP
    logRsvp(guest.guest_id, 'confirmed', payload);

    // Generate QR code
    const qrUrl = generateQR(guest.guest_id, guest.full_name);
    updateGuestField(sheetRow, headers, 'qr_code_url', qrUrl);

    // Send confirmation email with QR
    sendConfirmationEmail(email, guest.full_name, guest.guest_id, qrUrl);

    // Get tablemates for display
    const tablemates = getTablematesForGuest(guest.guest_id, guest.table_number);

    return {
      status: 200,
      message: 'Salamat sa iyong kumpirmasyon! Check your email para sa QR code.',
      confirmed: true,
      tablemates: tablemates
    };

  } else if (attendance === 'decline') {
    updateGuestFields(sheetRow, headers, {
      email: email,
      invite_status: 'declined',
      rsvp_date: now,
      updated_at: now
    });

    logRsvp(guest.guest_id, 'declined', payload);

    return {
      status: 200,
      message: 'Naiintindihan namin. Maraming salamat pa rin sa pagtugon!',
      declined: true
    };

  } else if (attendance === 'proxy') {
    const proxyName = (payload.proxy_name || '').trim();
    if (!proxyName) return { error: 'Ilagay ang pangalan ng kinatawan.', status: 400 };

    updateGuestFields(sheetRow, headers, {
      email: email,
      invite_status: 'proxy',
      rsvp_date: now,
      proxy_name: proxyName,
      special_needs: payload.special_needs ? JSON.stringify(payload.special_needs) : '',
      updated_at: now
    });

    logRsvp(guest.guest_id, 'proxy', payload);

    // QR goes to proxy name
    const qrUrl = generateQR(guest.guest_id, proxyName + ' (proxy for ' + guest.full_name + ')');
    updateGuestField(sheetRow, headers, 'qr_code_url', qrUrl);

    sendConfirmationEmail(email, guest.full_name, guest.guest_id, qrUrl, proxyName);

    return {
      status: 200,
      message: 'Salamat! Ang QR code para kay ' + proxyName + ' ay ipapadala sa iyong email.',
      proxy: true,
      proxy_name: proxyName
    };

  } else {
    return { error: 'Invalid attendance value.', status: 400 };
  }
}

// ============ TABLEMATES ============

function getTablemates(params) {
  const guestId = params.guest_id;
  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('guest_id');
  const tableCol = headers.indexOf('table_number');

  let tableNum = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === guestId) {
      tableNum = data[i][tableCol];
      break;
    }
  }

  if (!tableNum) return { status: 200, tablemates: [] };
  return { status: 200, tablemates: getTablematesForGuest(guestId, tableNum) };
}

function getTablematesForGuest(guestId, tableNum) {
  if (!tableNum) return [];

  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('guest_id');
  const nameCol = headers.indexOf('full_name');
  const tableCol = headers.indexOf('table_number');
  const statusCol = headers.indexOf('invite_status');

  const mates = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][tableCol] === tableNum && data[i][idCol] !== guestId) {
      mates.push({
        name: data[i][nameCol],
        status: ['confirmed', 'proxy'].includes(data[i][statusCol]) ? 'confirmed' : 'pending'
      });
    }
  }
  return mates;
}

// ============ QR CODE GENERATION ============

function generateQR(guestId, displayName) {
  // QR data — compact string for scanning
  const qrData = JSON.stringify({ id: guestId, name: displayName, event: '2027-10-15' });

  // Using goqr.me API (free, reliable, no auth needed)
  const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=400x400&format=png&data='
    + encodeURIComponent(qrData)
    + '&color=5c4a3d&bgcolor=faf8f5';

  return qrUrl;
}

// ============ EMAIL TEMPLATE HELPERS ============

const LOGO_URL = 'https://baynisandy.github.io/hiraya-wedding/logo-rd-web.png';

const DRESS_COLORS = [
  { name: 'Periwinkle',      hex: '#A7A4D8' },
  { name: 'Soft Lavender',   hex: '#C8B4E2' },
  { name: 'Dusty Mauve',     hex: '#A78BAA' },
  { name: 'Warm Taupe',      hex: '#B8A6A0' },
  { name: 'Champagne Beige', hex: '#C9B99F' },
  { name: 'Muted Olive',     hex: '#8E887A' }
];

function emailHeader() {
  return `
    <div style="background:linear-gradient(135deg,#5c4a3d 0%,#8e6f5e 100%);padding:32px 20px 24px;text-align:center;border-radius:12px 12px 0 0;">
      <img src="${LOGO_URL}" alt="R&D" width="120" style="display:block;margin:0 auto 12px;" />
      <div style="font-family:'Georgia',serif;font-size:14px;color:#e8d5c4;letter-spacing:6px;text-transform:uppercase;">Hiraya</div>
    </div>
  `;
}

function emailFooter() {
  return `
    <div style="background:#5c4a3d;padding:24px 20px;text-align:center;border-radius:0 0 12px 12px;">
      <p style="font-family:'Georgia',serif;font-size:13px;color:#e8d5c4;margin:0 0 4px;letter-spacing:2px;">Rodnel & Divine</p>
      <p style="font-family:Arial,sans-serif;font-size:11px;color:#b8a696;margin:0;">Oktubre 15, 2027 · Calumpit, Bulacan</p>
    </div>
  `;
}

function dressCodeBlock() {
  const swatches = DRESS_COLORS.map(c =>
    `<td style="padding:4px 6px;text-align:center;">
      <div style="width:44px;height:44px;border-radius:50%;background:${c.hex};margin:0 auto 6px;border:2px solid #f0ebe6;"></div>
      <span style="font-family:Arial,sans-serif;font-size:10px;color:#5c4a3d;line-height:1.2;display:block;">${c.name}</span>
    </td>`
  ).join('');

  return `
    <div style="background:#f8f4f0;border-radius:10px;padding:20px;margin:20px 0;text-align:center;">
      <p style="font-family:'Georgia',serif;font-size:14px;color:#5c4a3d;margin:0 0 4px;font-weight:bold;">Dress Code</p>
      <p style="font-family:Arial,sans-serif;font-size:12px;color:#8e6f5e;margin:0 0 16px;">Semi-formal · Mga sumusunod na kulay lamang</p>
      <table cellpadding="0" cellspacing="0" border="0" align="center" style="margin:0 auto;">
        <tr>${swatches}</tr>
      </table>
    </div>
  `;
}

function wrapEmail(bodyContent) {
  return `
<!DOCTYPE html>
<html>
<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"></head>
<body style="margin:0;padding:0;background:#f0ebe6;font-family:Arial,sans-serif;">
  <table cellpadding="0" cellspacing="0" border="0" width="100%" style="background:#f0ebe6;padding:24px 0;">
    <tr><td align="center">
      <table cellpadding="0" cellspacing="0" border="0" width="560" style="max-width:560px;width:100%;background:#ffffff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(92,74,61,0.08);">
        <tr><td>${emailHeader()}</td></tr>
        <tr><td style="padding:28px 32px;">${bodyContent}</td></tr>
        <tr><td>${emailFooter()}</td></tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;
}

// ============ CONFIRMATION EMAIL ============

function sendConfirmationEmail(email, guestName, guestId, qrUrl, proxyName) {
  try {
    const isProxy = !!proxyName;
    const subject = 'Hiraya — Kumpirmasyon ng Pagdalo · Rodnel & Divine · Oktubre 15, 2027';

    const bodyContent = `
      <p style="font-family:'Georgia',serif;font-size:18px;color:#5c4a3d;margin:0 0 16px;text-align:center;">
        Mahal naming ${guestName},
      </p>

      <p style="font-family:Arial,sans-serif;font-size:14px;color:#6b5b4f;line-height:1.7;margin:0 0 20px;text-align:center;">
        ${isProxy
          ? `Salamat sa pagtugon! Ang iyong kinatawan na si <strong>${proxyName}</strong> ay aming inaasahan sa aming kasal.`
          : 'Salamat sa kumpirmasyon ng iyong pagdalo! Lubos kaming nagagalak na makakasama ka sa aming espesyal na araw.'
        }
      </p>

      <div style="background:#faf8f5;border-radius:10px;padding:20px;margin:20px 0;border-left:4px solid #c4a265;">
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;width:100px;">Petsa</td>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">Oktubre 15, 2027 (Biyernes)</td>
          </tr>
          <tr>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;">Oras</td>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">2:00 ng hapon</td>
          </tr>
          <tr>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;">Lugar</td>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">Calumpit, Bulacan</td>
          </tr>
          <tr>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;">Wedding ID</td>
            <td style="padding:6px 0;font-family:'Courier New',monospace;font-size:14px;color:#5c4a3d;font-weight:bold;">${guestId}</td>
          </tr>
          ${isProxy ? `<tr>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;">Kinatawan</td>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">${proxyName}</td>
          </tr>` : ''}
        </table>
      </div>

      <div style="text-align:center;margin:24px 0;">
        <p style="font-family:'Georgia',serif;font-size:14px;color:#5c4a3d;margin:0 0 12px;">Iyong QR Code para sa Check-in</p>
        <img src="${qrUrl}" alt="QR Code" width="200" height="200" style="display:block;margin:0 auto;border-radius:8px;border:2px solid #f0ebe6;" />
        <p style="font-family:Arial,sans-serif;font-size:11px;color:#a0917f;margin:8px 0 0;">I-screenshot o i-save ang QR code na ito. Ipakita sa entrance sa araw ng kasal.</p>
      </div>

      ${dressCodeBlock()}

      <p style="font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;line-height:1.6;margin:20px 0 0;text-align:center;">
        Nagpapasalamat kami sa inyong pagtugon. Inaasahan namin ang inyong presensya sa aming pagdiriwang.
      </p>

      <p style="font-family:'Georgia',serif;font-size:15px;color:#5c4a3d;margin:20px 0 0;text-align:center;">
        Lubos na nagpapasalamat,<br/>
        <strong>Rodnel & Divine</strong>
      </p>
    `;

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: `Hiraya — Kumpirmasyon ng Pagdalo\n\nMahal naming ${guestName},\n\n${isProxy ? 'Ang kinatawan mo na si ' + proxyName + ' ay inaasahan sa kasal.' : 'Salamat sa kumpirmasyon!'}\n\nPetsa: Oktubre 15, 2027\nLugar: Calumpit, Bulacan\nOras: 2:00 PM\nWedding ID: ${guestId}\n\nQR Code: ${qrUrl}\n\nDress Code: Semi-formal (Periwinkle, Soft Lavender, Dusty Mauve, Warm Taupe, Champagne Beige, Muted Olive)\n\nNagpapasalamat,\nRodnel & Divine`,
      htmlBody: wrapEmail(bodyContent),
      name: 'Rodnel & Divine — Hiraya'
    });

    logEmail(guestId, 'confirmation', email, 'sent');

  } catch (err) {
    logEmail(guestId, 'confirmation', email, 'failed', err.message);
  }
}

// ============ ADMIN ENDPOINTS ============

function verifyAdmin(payload) {
  if (payload.admin_key !== ADMIN_KEY) {
    throw new Error('Unauthorized — invalid admin key');
  }
}

function adminGetGuests(params) {
  verifyAdmin(params);
  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const guests = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) guests.push(rowToObj(headers, data[i]));
  }

  return { status: 200, guests: guests, count: guests.length };
}

function adminGetStats(params) {
  verifyAdmin(params);
  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf('invite_status');

  const stats = { total: 0, confirmed: 0, declined: 0, proxy: 0, pending: 0, sent: 0, expired: 0 };
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    stats.total++;
    const s = data[i][statusCol];
    if (stats.hasOwnProperty(s)) stats[s]++;
    else stats.pending++;
  }

  return { status: 200, stats: stats };
}

function adminAddGuest(payload) {
  verifyAdmin(payload);
  const sheet = getSheet('GUEST_LIST');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const guestId = generateGuestId();
  const inviteCode = generateInviteCode();
  const now = new Date();

  const row = headers.map(h => {
    switch (h) {
      case 'guest_id': return guestId;
      case 'full_name': return payload.full_name || '';
      case 'email': return payload.email || '';
      case 'phone': return payload.phone || '';
      case 'guest_type': return payload.guest_type || 'tech-savvy';
      case 'group_label': return payload.group_label || '';
      case 'invite_code': return inviteCode;
      case 'invite_status': return 'pending';
      case 'table_number': return payload.table_number || '';
      case 'notes': return payload.notes || '';
      case 'side': return payload.side || '';
      case 'relation': return payload.relation || '';
      case 'is_memorial': return payload.is_memorial || 'FALSE';
      case 'wedding_id': return guestId;
      case 'created_at': return now;
      case 'updated_at': return now;
      default: return '';
    }
  });

  sheet.appendRow(row);

  return {
    status: 200,
    message: 'Guest added!',
    guest_id: guestId,
    invite_code: inviteCode
  };
}

function adminUpdateGuest(payload) {
  verifyAdmin(payload);
  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('guest_id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === payload.guest_id) {
      const fields = payload.fields || {};
      fields.updated_at = new Date();
      updateGuestFields(i + 1, headers, fields);
      return { status: 200, message: 'Guest updated!' };
    }
  }
  return { error: 'Guest not found', status: 404 };
}

function adminDeleteGuest(payload) {
  verifyAdmin(payload);
  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('guest_id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === payload.guest_id) {
      sheet.deleteRow(i + 1);
      return { status: 200, message: 'Guest deleted!' };
    }
  }
  return { error: 'Guest not found', status: 404 };
}

function adminAssignTable(payload) {
  verifyAdmin(payload);
  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('guest_id');

  // payload.assignments = [{ guest_id, table_number }]
  const assignments = payload.assignments || [];
  let updated = 0;

  for (const a of assignments) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol] === a.guest_id) {
        updateGuestFields(i + 1, headers, {
          table_number: a.table_number,
          updated_at: new Date()
        });
        updated++;
        break;
      }
    }
  }

  return { status: 200, message: `${updated} guest(s) assigned to tables.` };
}

function adminReopenInvite(payload) {
  verifyAdmin(payload);
  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('guest_id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === payload.guest_id) {
      const newExpiry = new Date();
      newExpiry.setDate(newExpiry.getDate() + 7); // +7 days from now

      updateGuestFields(i + 1, headers, {
        invite_status: 'sent',
        invite_expiry: newExpiry,
        updated_at: new Date()
      });

      logRsvp(payload.guest_id, 'reopened', { admin: true });

      return { status: 200, message: 'Invite reopened! New expiry: ' + newExpiry.toISOString() };
    }
  }
  return { error: 'Guest not found', status: 404 };
}

function adminSendInvite(payload) {
  verifyAdmin(payload);
  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('guest_id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === payload.guest_id) {
      const guest = rowToObj(headers, data[i]);

      if (!guest.email) return { error: 'Walang email ang guest na ito.', status: 400 };

      const inviteUrl = payload.base_url + '?code=' + guest.invite_code;
      const now = new Date();
      const expiry = new Date();
      expiry.setDate(expiry.getDate() + 7);

      // Send invitation email with HTML template
      try {
        const invSubject = 'Hiraya — Imbitasyon sa Kasal · Rodnel & Divine · Oktubre 15, 2027';

        const invBodyContent = `
          <p style="font-family:'Georgia',serif;font-size:18px;color:#5c4a3d;margin:0 0 16px;text-align:center;">
            Mahal naming ${guest.full_name},
          </p>

          <p style="font-family:Arial,sans-serif;font-size:14px;color:#6b5b4f;line-height:1.7;margin:0 0 20px;text-align:center;">
            Nais naming ibahagi sa inyo ang isa sa pinaka-espesyal na araw ng aming buhay. Kayo ay aming iniimbitahan sa aming kasal.
          </p>

          <div style="background:#faf8f5;border-radius:10px;padding:20px;margin:20px 0;border-left:4px solid #c4a265;">
            <table cellpadding="0" cellspacing="0" border="0" width="100%">
              <tr>
                <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;width:100px;">Petsa</td>
                <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">Oktubre 15, 2027 (Biyernes)</td>
              </tr>
              <tr>
                <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;">Oras</td>
                <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">2:00 ng hapon</td>
              </tr>
              <tr>
                <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;">Lugar</td>
                <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">Calumpit, Bulacan</td>
              </tr>
            </table>
          </div>

          <div style="text-align:center;margin:24px 0;">
            <a href="${inviteUrl}" style="display:inline-block;padding:14px 36px;background:linear-gradient(135deg,#c4a265,#d4b87a);color:#fff;text-decoration:none;border-radius:30px;font-family:Arial,sans-serif;font-size:14px;font-weight:bold;letter-spacing:1px;">KUMPIRMAHIN ANG PAGDALO</a>
          </div>

          <p style="font-family:Arial,sans-serif;font-size:11px;color:#a0917f;text-align:center;margin:0 0 20px;">
            Ang link na ito ay valid hanggang ${expiry.toLocaleDateString('en-PH', {year:'numeric',month:'long',day:'numeric'})}.<br/>
            RSVP Deadline: Setyembre 4, 2027
          </p>

          ${dressCodeBlock()}

          <p style="font-family:'Georgia',serif;font-size:15px;color:#5c4a3d;margin:20px 0 0;text-align:center;">
            Lubos na nagpapasalamat,<br/>
            <strong>Rodnel & Divine</strong>
          </p>
        `;

        MailApp.sendEmail({
          to: guest.email,
          subject: invSubject,
          body: `Hiraya — Imbitasyon sa Kasal\n\nMahal naming ${guest.full_name},\n\nKayo ay aming iniimbitahan sa aming kasal!\n\nPetsa: Oktubre 15, 2027\nLugar: Calumpit, Bulacan\nOras: 2:00 PM\n\nKumpirmahin: ${inviteUrl}\nValid hanggang: ${expiry.toLocaleDateString('en-PH')}\nRSVP Deadline: Setyembre 4, 2027\n\nDress Code: Semi-formal (Periwinkle, Soft Lavender, Dusty Mauve, Warm Taupe, Champagne Beige, Muted Olive)\n\nNagpapasalamat,\nRodnel & Divine`,
          htmlBody: wrapEmail(invBodyContent),
          name: 'Rodnel & Divine — Hiraya'
        });

        updateGuestFields(i + 1, headers, {
          invite_status: 'sent',
          invite_sent_date: now,
          invite_expiry: expiry,
          updated_at: now
        });

        logEmail(guest.guest_id, 'invitation', guest.email, 'sent');

        return { status: 200, message: 'Invitation sent to ' + guest.email };

      } catch (err) {
        logEmail(guest.guest_id, 'invitation', guest.email, 'failed', err.message);
        return { error: 'Failed to send: ' + err.message, status: 500 };
      }
    }
  }
  return { error: 'Guest not found', status: 404 };
}

// ============ SCHEDULED REMINDERS ============

/**
 * Set up a daily trigger to check and send reminders.
 * Run this once manually: setupReminderTrigger()
 */
function setupReminderTrigger() {
  ScriptApp.newTrigger('checkAndSendReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9) // 9 AM PHT
    .create();
}

function checkAndSendReminders() {
  const weddingDate = new Date(getConfig('wedding_date'));
  const now = new Date();
  const daysUntil = Math.ceil((weddingDate - now) / (1000 * 60 * 60 * 24));

  let reminderType = null;
  if (daysUntil === 90 && getConfig('reminder_3mo') === 'true') reminderType = 'reminder_3mo';
  else if (daysUntil === 30 && getConfig('reminder_1mo') === 'true') reminderType = 'reminder_1mo';
  else if (daysUntil === 7 && getConfig('reminder_1wk') === 'true') reminderType = 'reminder_1wk';
  else if (daysUntil === 1 && getConfig('reminder_1day') === 'true') reminderType = 'reminder_1day';

  if (!reminderType) return;

  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  for (let i = 1; i < data.length; i++) {
    const guest = rowToObj(headers, data[i]);
    if (guest.invite_status === 'confirmed' && guest.email) {
      sendReminderEmail(guest, reminderType, daysUntil);
    }
  }
}

function sendReminderEmail(guest, type, daysUntil) {
  const timeLabel = daysUntil === 1 ? 'bukas na'
    : daysUntil === 7 ? 'sa isang linggo na'
    : daysUntil === 30 ? 'sa isang buwan na'
    : daysUntil === 90 ? 'sa tatlong buwan na'
    : 'malapit na';

  const urgency = daysUntil <= 7 ? 'Malapit Na' : 'Paalala';

  try {
    const subject = `Hiraya — ${urgency}: ${timeLabel} ang kasal! · Oktubre 15, 2027`;

    const bodyContent = `
      <p style="font-family:'Georgia',serif;font-size:18px;color:#5c4a3d;margin:0 0 16px;text-align:center;">
        Mahal naming ${guest.full_name},
      </p>

      <p style="font-family:Arial,sans-serif;font-size:14px;color:#6b5b4f;line-height:1.7;margin:0 0 20px;text-align:center;">
        Paalala lang — <strong>${timeLabel}</strong> ang aming kasal! Inaasahan namin ang inyong pagdalo.
      </p>

      <div style="background:#faf8f5;border-radius:10px;padding:20px;margin:20px 0;border-left:4px solid #c4a265;">
        <table cellpadding="0" cellspacing="0" border="0" width="100%">
          <tr>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;width:100px;">Petsa</td>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">Oktubre 15, 2027 (Biyernes)</td>
          </tr>
          <tr>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;">Oras</td>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">2:00 ng hapon</td>
          </tr>
          <tr>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;">Lugar</td>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:14px;color:#5c4a3d;font-weight:bold;">Calumpit, Bulacan</td>
          </tr>
          <tr>
            <td style="padding:6px 0;font-family:Arial,sans-serif;font-size:13px;color:#8e6f5e;">Wedding ID</td>
            <td style="padding:6px 0;font-family:'Courier New',monospace;font-size:14px;color:#5c4a3d;font-weight:bold;">${guest.guest_id}</td>
          </tr>
        </table>
      </div>

      ${guest.qr_code_url ? `
      <div style="text-align:center;margin:24px 0;">
        <p style="font-family:'Georgia',serif;font-size:14px;color:#5c4a3d;margin:0 0 12px;">Iyong QR Code para sa Check-in</p>
        <img src="${guest.qr_code_url}" alt="QR Code" width="180" height="180" style="display:block;margin:0 auto;border-radius:8px;border:2px solid #f0ebe6;" />
        <p style="font-family:Arial,sans-serif;font-size:11px;color:#a0917f;margin:8px 0 0;">Ipakita ang QR code na ito sa entrance sa araw ng kasal.</p>
      </div>` : ''}

      ${dressCodeBlock()}

      <div style="background:#faf8f5;border-radius:10px;padding:16px;margin:16px 0;text-align:center;">
        <p style="font-family:Arial,sans-serif;font-size:12px;color:#8e6f5e;margin:0;line-height:1.6;">
          Huwag kalimutang dalhin ang iyong QR code at sundin ang dress code. Maraming salamat!
        </p>
      </div>

      <p style="font-family:'Georgia',serif;font-size:15px;color:#5c4a3d;margin:20px 0 0;text-align:center;">
        Nagpapasalamat,<br/>
        <strong>Rodnel & Divine</strong>
      </p>
    `;

    MailApp.sendEmail({
      to: guest.email,
      subject: subject,
      body: `Hiraya — Paalala\n\nMahal naming ${guest.full_name},\n\n${timeLabel} ang aming kasal!\n\nPetsa: Oktubre 15, 2027\nLugar: Calumpit, Bulacan\nOras: 2:00 PM\nWedding ID: ${guest.guest_id}\n${guest.qr_code_url ? 'QR Code: ' + guest.qr_code_url : ''}\n\nDress Code: Semi-formal (Periwinkle, Soft Lavender, Dusty Mauve, Warm Taupe, Champagne Beige, Muted Olive)\n\nNagpapasalamat,\nRodnel & Divine`,
      htmlBody: wrapEmail(bodyContent),
      name: 'Rodnel & Divine — Hiraya'
    });
    logEmail(guest.guest_id, type, guest.email, 'sent');
  } catch (err) {
    logEmail(guest.guest_id, type, guest.email, 'failed', err.message);
  }
}

// ============ HELPER FUNCTIONS ============

function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i]; });
  return obj;
}

function updateGuestField(rowNum, headers, field, value) {
  const col = headers.indexOf(field);
  if (col === -1) return;
  getSheet('GUEST_LIST').getRange(rowNum, col + 1).setValue(value);
}

function updateGuestFields(rowNum, headers, fields) {
  const sheet = getSheet('GUEST_LIST');
  for (const [field, value] of Object.entries(fields)) {
    const col = headers.indexOf(field);
    if (col !== -1) sheet.getRange(rowNum, col + 1).setValue(value);
  }
}

function generateGuestId() {
  const sheet = getSheet('GUEST_LIST');
  const lastRow = sheet.getLastRow();
  const num = lastRow; // header is row 1, so lastRow = count + 1, but row 1 is header so this gives next number
  return 'DR-2027-' + String(num).padStart(4, '0');
}

function generateInviteCode() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // no confusing chars (0/O, 1/I)
  let code = '';
  for (let i = 0; i < 8; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

function logRsvp(guestId, action, details) {
  const sheet = getSheet('RSVP_LOG');
  sheet.appendRow([
    sheet.getLastRow(), // log_id
    guestId,
    action,
    new Date(),
    '', // ip_address (not available in Apps Script)
    JSON.stringify(details || {})
  ]);
}

function logEmail(guestId, type, email, status, error) {
  const sheet = getSheet('EMAIL_LOG');
  sheet.appendRow([
    sheet.getLastRow(),
    guestId,
    type,
    email,
    new Date(),
    status,
    error || ''
  ]);
}

// ============ SHEET INITIALIZATION ============

/**
 * Run this once to create all sheets with headers.
 * Menu: Run > initializeSheets
 */
function initializeSheets() {
  const ss = getSpreadsheet();

  // GUEST_LIST
  createSheetWithHeaders(ss, 'GUEST_LIST', [
    'guest_id', 'full_name', 'email', 'phone', 'guest_type', 'group_label',
    'invite_code', 'invite_status', 'rsvp_date', 'proxy_name',
    'table_number', 'companions', 'special_needs', 'qr_code_url', 'wedding_id',
    'invite_sent_date', 'invite_expiry', 'notes', 'created_at', 'updated_at'
  ]);

  // RSVP_LOG
  createSheetWithHeaders(ss, 'RSVP_LOG', [
    'log_id', 'guest_id', 'action', 'timestamp', 'ip_address', 'details'
  ]);

  // TABLE_PLAN
  createSheetWithHeaders(ss, 'TABLE_PLAN', [
    'table_number', 'table_name', 'capacity', 'assigned_count', 'confirmed_count', 'notes'
  ]);

  // EVENT_CONFIG
  const configSheet = createSheetWithHeaders(ss, 'EVENT_CONFIG', ['key', 'value']);
  const defaults = [
    ['wedding_date', '2027-10-15'],
    ['ceremony_time', '14:00'],
    ['venue_name', '(TBD)'],
    ['venue_address', 'Calumpit, Bulacan'],
    ['rsvp_deadline', '2027-09-04'],
    ['invite_expiry_days', '7'],
    ['couple_name_1', 'Rodnel'],
    ['couple_name_2', 'Divine'],
    ['email_sender_name', 'Rodnel & Divine'],
    ['reminder_3mo', 'true'],
    ['reminder_1mo', 'true'],
    ['reminder_1wk', 'true'],
    ['reminder_1day', 'true']
  ];
  defaults.forEach(row => configSheet.appendRow(row));

  // EMAIL_LOG
  createSheetWithHeaders(ss, 'EMAIL_LOG', [
    'log_id', 'guest_id', 'email_type', 'sent_to', 'sent_at', 'status', 'error'
  ]);

  SpreadsheetApp.getUi().alert('✅ All sheets initialized! Handa na ang database.');
}

function createSheetWithHeaders(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  } else {
    sheet.clear();
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Bold header row
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f3f0ff');

  // Freeze header row
  sheet.setFrozenRows(1);

  return sheet;
}

// ============ PUBLIC CONFIG (for frontend) ============

function getPublicConfig() {
  return {
    status: 200,
    config: {
      wedding_date: getConfig('wedding_date'),
      ceremony_time: getConfig('ceremony_time'),
      venue_name: getConfig('venue_name'),
      venue_address: getConfig('venue_address'),
      rsvp_deadline: getConfig('rsvp_deadline'),
      couple: getConfig('couple_name_1') + ' & ' + getConfig('couple_name_2')
    }
  };
}

// ============ COORDINATOR CHECK-IN ============

/**
 * Lookup guest by QR code data or search term (GET)
 * Used by coordinator to find guest before check-in
 */
function adminCheckinLookup(params) {
  if (params.key !== ADMIN_KEY) return { error: 'Unauthorized', status: 401 };
  const query = (params.q || '').trim().toLowerCase();
  if (!query) return { error: 'No search query', status: 400 };

  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const results = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    const guest = rowToObj(headers, data[i]);
    const name = (guest.full_name || '').toLowerCase();
    const code = (guest.invite_code || '').toLowerCase();
    const gid = (guest.guest_id || '').toLowerCase();
    const table = (guest.table_number || '').toString();

    if (name.includes(query) || code === query || gid === query || query === 'all') {
      results.push({
        guest_id: guest.guest_id,
        full_name: guest.full_name,
        email: guest.email,
        invite_code: guest.invite_code,
        invite_status: guest.invite_status,
        table_number: guest.table_number,
        group_label: guest.group_label,
        checked_in: guest.checked_in || false,
        checkin_time: guest.checkin_time || '',
        notes: guest.notes || ''
      });
    }
  }

  return { status: 200, guests: results, total: results.length };
}

/**
 * Check in a guest (POST)
 * Sets checked_in = true and records timestamp
 */
function adminCheckinGuest(payload) {
  if (payload.key !== ADMIN_KEY) return { error: 'Unauthorized', status: 401 };
  const guestId = payload.guest_id;
  if (!guestId) return { error: 'Missing guest_id', status: 400 };

  const sheet = getSheet('GUEST_LIST');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('guest_id');

  // Ensure checked_in and checkin_time columns exist
  let checkinCol = headers.indexOf('checked_in');
  let timeCol = headers.indexOf('checkin_time');
  if (checkinCol === -1) {
    checkinCol = headers.length;
    sheet.getRange(1, checkinCol + 1).setValue('checked_in');
  }
  if (timeCol === -1) {
    timeCol = headers.indexOf('checkin_time') !== -1 ? headers.indexOf('checkin_time') : (checkinCol === headers.length ? headers.length + 1 : headers.length);
    if (headers.indexOf('checkin_time') === -1) sheet.getRange(1, timeCol + 1).setValue('checkin_time');
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === guestId) {
      const now = new Date().toISOString();
      const isUndo = payload.undo === true;

      sheet.getRange(i + 1, checkinCol + 1).setValue(isUndo ? '' : 'TRUE');
      sheet.getRange(i + 1, timeCol + 1).setValue(isUndo ? '' : now);

      const guest = rowToObj(headers, data[i]);
      return {
        status: 200,
        message: isUndo ? `${guest.full_name} check-in undone` : `${guest.full_name} checked in!`,
        guest_id: guestId,
        checked_in: !isUndo,
        checkin_time: isUndo ? '' : now,
        full_name: guest.full_name,
        table_number: guest.table_number || ''
      };
    }
  }

  return { error: 'Guest not found', status: 404 };
}
