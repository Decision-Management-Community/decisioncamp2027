const CONFIG = Object.freeze({
	SHEET_NAME: 'Registrations',
	TIME_ZONE: 'America/New_York',
	WEBSITE_URL: 'https://decision-management-community.github.io/decisioncamp2027/',
	CONTACT_EMAIL: 'decisionmanagementcommunity@gmail.com',
	EVENT_NAME: 'DecisionCAMP 2027',
	REMINDER_START: new Date('2027-09-10T00:00:00-04:00'),
	REMINDER_SEND_AT: new Date('2027-09-10T09:00:00-04:00'),
	REMINDER_END: new Date('2027-09-18T00:00:00-04:00'),
	HEADERS: [
		'Timestamp',
		'Name',
		'Email',
		'Organization',
		'Website',
		'Comments',
		'Consent',
		'Confirmation sent',
		'Reminder sent',
		'Source',
	],
});

const EVENTS = Object.freeze([
	{ day: 1, date: 'September 14, 2027', startUtc: '20270914T130000Z', endUtc: '20270914T190000Z' },
	{ day: 2, date: 'September 15, 2027', startUtc: '20270915T130000Z', endUtc: '20270915T190000Z' },
	{ day: 3, date: 'September 16, 2027', startUtc: '20270916T130000Z', endUtc: '20270916T190000Z' },
]);

function setupRegistrationSystem() {
	const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
	if (!spreadsheet) {
		throw new Error('Open this script from its Google Sheet before running setupRegistrationSystem.');
	}

	PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheet.getId());
	ensureRegistrationSheet_(spreadsheet);
	ensureReminderTrigger_();
}

function doPost(event) {
	try {
		const values = event && event.parameter ? event.parameter : {};
		if (values.company) {
			return registrationResponse_('success', 'Registration received.');
		}

		const registration = validateRegistration_(values);
		const lock = LockService.getScriptLock();
		lock.waitLock(20000);

		try {
			upsertRegistration_(registration);
		} finally {
			lock.releaseLock();
		}

		sendConfirmation_(registration);
		markConfirmationSent_(registration.email);

		return registrationResponse_(
			'success',
			'You are registered. Check your email for the confirmation and calendar files.',
		);
	} catch (error) {
		console.error(error);
		return registrationResponse_(
			'error',
			error && error.message ? error.message : 'Registration could not be completed. Please try again.',
		);
	}
}

function sendSeptemberReminder() {
	const now = new Date();
	if (now < CONFIG.REMINDER_START) return;

	if (now >= CONFIG.REMINDER_END) {
		deleteReminderTriggers_();
		return;
	}

	const sheet = getRegistrationSheet_();
	const lastRow = sheet.getLastRow();
	if (lastRow < 2) return;

	const values = sheet.getRange(2, 1, lastRow - 1, CONFIG.HEADERS.length).getValues();
	let remainingQuota = MailApp.getRemainingDailyQuota();
	let pending = 0;

	values.forEach((row, index) => {
		const email = String(row[2] || '').trim();
		const consent = String(row[6] || '').toLowerCase() === 'yes';
		const reminderSent = Boolean(row[8]);

		if (!email || !consent || reminderSent) return;
		pending += 1;
		if (remainingQuota < 1) return;

		MailApp.sendEmail({
			to: email,
			subject: 'DecisionCAMP 2027: check the website for the live-event link',
			body: reminderText_(row[1]),
			htmlBody: reminderHtml_(row[1]),
			name: 'DecisionCAMP 2027',
			replyTo: CONFIG.CONTACT_EMAIL,
		});

		sheet.getRange(index + 2, 9).setValue(new Date());
		remainingQuota -= 1;
		pending -= 1;
	});

	deleteReminderTriggers_();
	if (pending > 0) {
		ScriptApp.newTrigger('sendSeptemberReminder')
			.timeBased()
			.after(24 * 60 * 60 * 1000)
			.create();
	}
}

function validateRegistration_(values) {
	const name = cleanText_(values.name, 120);
	const email = String(values.email || '').trim().toLowerCase();
	const organization = cleanText_(values.organization, 160);
	const website = String(values.website || '').trim().slice(0, 300);
	const comments = cleanText_(values.comments, 2000);
	const consent = String(values.consent || '').toLowerCase() === 'yes';

	if (!name) throw new Error('Please enter your name.');
	if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Please enter a valid email address.');
	if (website && !/^https?:\/\//i.test(website)) throw new Error('Website must begin with http:// or https://.');
	if (!consent) throw new Error('Please agree to receive DecisionCAMP 2027 event emails.');

	return {
		name,
		email,
		organization,
		website,
		comments,
		consent: 'Yes',
		source: cleanText_(values.source || 'DecisionCAMP 2027 website', 120),
	};
}

function upsertRegistration_(registration) {
	const sheet = getRegistrationSheet_();
	const lastRow = sheet.getLastRow();
	let existingRow = 0;

	if (lastRow >= 2) {
		const emails = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();
		const match = emails.findIndex((row) => String(row[0]).trim().toLowerCase() === registration.email);
		if (match >= 0) existingRow = match + 2;
	}

	const row = [
		new Date(),
		safeSheetValue_(registration.name),
		safeSheetValue_(registration.email),
		safeSheetValue_(registration.organization),
		safeSheetValue_(registration.website),
		safeSheetValue_(registration.comments),
		registration.consent,
		'',
		'',
		safeSheetValue_(registration.source),
	];

	if (existingRow) {
		const priorReminder = sheet.getRange(existingRow, 9).getValue();
		row[8] = priorReminder;
		sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
	} else {
		sheet.appendRow(row);
	}
}

function sendConfirmation_(registration) {
	if (MailApp.getRemainingDailyQuota() < 1) {
		throw new Error('The confirmation email limit has been reached. Please contact the organizers.');
	}

	MailApp.sendEmail({
		to: registration.email,
		subject: 'You are registered for DecisionCAMP 2027',
		body: confirmationText_(registration.name),
		htmlBody: confirmationHtml_(registration.name),
		attachments: EVENTS.map(calendarAttachment_),
		name: 'DecisionCAMP 2027',
		replyTo: CONFIG.CONTACT_EMAIL,
	});
}

function confirmationText_(name) {
	return [
		`Hello ${name},`,
		'',
		'You are registered for DecisionCAMP 2027, held online September 14–16, 2027.',
		'Sessions run from 9:00 AM to 3:00 PM Eastern Time each day.',
		'',
		'Three calendar files are attached, one for each conference day.',
		'',
		'After September 10, please visit the “Register free” section of the DecisionCAMP 2027 website to receive the link for joining the live event:',
		CONFIG.WEBSITE_URL + '#register',
		'',
		'We look forward to seeing you!',
		'Decision Management Community',
	].join('\n');
}

function confirmationHtml_(name) {
	const safeName = escapeHtml_(name);
	return `
		<p>Hello ${safeName},</p>
		<p>You are registered for <strong>DecisionCAMP 2027</strong>, held online September 14–16, 2027.</p>
		<p>Sessions run from <strong>9:00 AM to 3:00 PM Eastern Time</strong> each day.</p>
		<p>Three calendar files are attached, one for each conference day. Open them to add the sessions to your calendar.</p>
		<p><strong>After September 10:</strong> visit the <a href="${CONFIG.WEBSITE_URL}#register">Register free section of the DecisionCAMP 2027 website</a> to receive the link for joining the live event.</p>
		<p>We look forward to seeing you!</p>
		<p>Decision Management Community</p>
	`;
}

function reminderText_(name) {
	return [
		`Hello ${name || 'DecisionCAMP participant'},`,
		'',
		'DecisionCAMP 2027 begins soon.',
		'Please visit the “Register free” section of the website for the link to join the live event:',
		CONFIG.WEBSITE_URL + '#register',
		'',
		'September 14–16, 2027 · 9:00 AM–3:00 PM ET daily',
	].join('\n');
}

function reminderHtml_(name) {
	return `
		<p>Hello ${escapeHtml_(name || 'DecisionCAMP participant')},</p>
		<p><strong>DecisionCAMP 2027 begins soon.</strong></p>
		<p>Please visit the <a href="${CONFIG.WEBSITE_URL}#register">Register free section of the website</a> for the link to join the live event.</p>
		<p>September 14–16, 2027 · 9:00 AM–3:00 PM ET daily</p>
	`;
}

function calendarAttachment_(event) {
	const description = `DecisionCAMP 2027 Day ${event.day}. After September 10, visit ${CONFIG.WEBSITE_URL}#register for the live-event link.`;
	const lines = [
		'BEGIN:VCALENDAR',
		'VERSION:2.0',
		'PRODID:-//Decision Management Community//DecisionCAMP 2027//EN',
		'CALSCALE:GREGORIAN',
		'METHOD:PUBLISH',
		'BEGIN:VEVENT',
		`UID:decisioncamp-2027-day-${event.day}@dmcommunity.org`,
		`DTSTAMP:${Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'")}`,
		`DTSTART:${event.startUtc}`,
		`DTEND:${event.endUtc}`,
		`SUMMARY:${CONFIG.EVENT_NAME} - Day ${event.day}`,
		`DESCRIPTION:${escapeIcs_(description)}`,
		'LOCATION:Online',
		`URL:${CONFIG.WEBSITE_URL}#register`,
		'END:VEVENT',
		'END:VCALENDAR',
	];

	return Utilities.newBlob(
		lines.join('\r\n'),
		'text/calendar',
		`decisioncamp-2027-day-${event.day}.ics`,
	);
}

function markConfirmationSent_(email) {
	const sheet = getRegistrationSheet_();
	const lastRow = sheet.getLastRow();
	if (lastRow < 2) return;

	const emails = sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues();
	const match = emails.findIndex((row) => String(row[0]).trim().toLowerCase() === email);
	if (match >= 0) sheet.getRange(match + 2, 8).setValue(new Date());
}

function getRegistrationSheet_() {
	const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
	if (!spreadsheetId) throw new Error('Registration system setup is incomplete.');
	return ensureRegistrationSheet_(SpreadsheetApp.openById(spreadsheetId));
}

function ensureRegistrationSheet_(spreadsheet) {
	let sheet = spreadsheet.getSheetByName(CONFIG.SHEET_NAME);
	if (!sheet) sheet = spreadsheet.insertSheet(CONFIG.SHEET_NAME);

	if (sheet.getLastRow() === 0) {
		sheet.appendRow(CONFIG.HEADERS);
		sheet.setFrozenRows(1);
		sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setFontWeight('bold');
		sheet.autoResizeColumns(1, CONFIG.HEADERS.length);
	}

	return sheet;
}

function ensureReminderTrigger_() {
	const exists = ScriptApp.getProjectTriggers().some(
		(trigger) => trigger.getHandlerFunction() === 'sendSeptemberReminder',
	);
	if (exists) return;

	ScriptApp.newTrigger('sendSeptemberReminder')
		.timeBased()
		.at(CONFIG.REMINDER_SEND_AT)
		.create();
}

function deleteReminderTriggers_() {
	ScriptApp.getProjectTriggers()
		.filter((trigger) => trigger.getHandlerFunction() === 'sendSeptemberReminder')
		.forEach((trigger) => ScriptApp.deleteTrigger(trigger));
}

function registrationResponse_(status, message) {
	const payload = JSON.stringify({ type: 'decisioncamp-registration', status, message });
	return HtmlService.createHtmlOutput(
		`<!doctype html><html><body><script>window.parent.postMessage(${payload}, '*');</script></body></html>`,
	).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function cleanText_(value, maxLength) {
	return String(value || '').trim().replace(/\s+/g, ' ').slice(0, maxLength);
}

function safeSheetValue_(value) {
	const text = String(value || '');
	return /^[=+\-@]/.test(text) ? `'${text}` : text;
}

function escapeHtml_(value) {
	return String(value || '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#039;');
}

function escapeIcs_(value) {
	return String(value || '')
		.replace(/\\/g, '\\\\')
		.replace(/\r?\n/g, '\\n')
		.replace(/,/g, '\\,')
		.replace(/;/g, '\\;');
}
