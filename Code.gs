function myFunction() {
	Logger.log(MailApp.getRemainingDailyQuota());
}

function onFormSubmit(e) {
	const lock = LockService.getDocumentLock() || LockService.getScriptLock();
	lock.waitLock(30000);

	try {
		processRow(e.range.getRow(), e.range.getSheet());
	} finally {
		lock.releaseLock();
	}
}

function retryEmailByRow(row) {
	row = Number(row);
	if (!row || row < 2) throw new Error('Invalid row number: ' + row);

	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
	if (!sheet) throw new Error('Sheet not found');

	const rowData = sheet.getRange(row, 1, 1, 14).getValues()[0];
	const storedUsername = String(rowData[12] || '').trim();
	const storedPassword = String(rowData[13] || '').trim();

	if (storedUsername && storedPassword && !storedUsername.startsWith('Error:')) {
		try {
			sendEmail(row, sheet, rowData, storedUsername, storedPassword);
		} catch (err) {
			sheet.getRange(row, 13).setValue('Error: ' + err.message);
			throw err;
		}
	} else {
		processRow(row, sheet, false);
	}
}

function processRow(row, sheet, increment = true) {
	const rowData = sheet.getRange(row, 1, 1, 13).getValues()[0];
	const track = getTrack(String(rowData[7] || '').trim());

	try {
		const { username, password } = fetchCredentials(track, increment);
		sendEmail(row, sheet, rowData, username, password);
	} catch (err) {
		sheet.getRange(row, 13).setValue('Error: ' + err.message);
		throw err;
	}
}

function getTrack(track) {
	if (track === 'วิทย์ - คณิต') return 'sm';
	if (track === 'ภาษา - คำนวน') return 'am';
	return 'el';
}

function fetchCredentials(track, increment = true) {
	const codes = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Codes');
	if (!codes) throw new Error('Sheet "Codes" not found');

	const indexCol = { sm: 9, am: 10, el: 11 }[track];
	const maxCol = { sm: 14, am: 15, el: 16 }[track];
	const credCol = { sm: 1, am: 3, el: 5 }[track];

	const meta = codes.getRange(2, indexCol, 1, maxCol - indexCol + 1).getValues()[0];
	const index = Number(meta[0]) || 1;
	const max = Number(meta[maxCol - indexCol]) || 0;

	if (index > max) throw new Error('Credential quota exceeded for track ' + track + ' (' + index + '/' + max + ')');

	const [username, password] = codes.getRange(index + 2, credCol, 1, 2).getValues()[0].map(v => String(v).trim());

	if (increment) codes.getRange(2, indexCol).setValue(index + 1);

	return { username, password };
}

function sendEmail(row, sheet, rowData, username, password) {
	if (MailApp.getRemainingDailyQuota() < 1) throw new Error('No email quota remaining today');

	const email = String(rowData[1] || '').trim();
	const name = String(rowData[2] || '').trim();
	const surname = String(rowData[3] || '').trim();

	if (!email) throw new Error('Missing email address');
	console.log('Sending to: %s | username: %s', email, username);

	const subject = 'ข้อมูลการเข้าสู่ระบบ Mock x Triam The Trilogy for TU90';
	const htmlBody = generateHtmlBody(username, password);
	const plainText = generatePlainText(name, surname, username, password);

	GmailApp.sendEmail(email, subject, plainText, {
		htmlBody: htmlBody,
		name: 'Mock x Triam The Trilogy',
		// noreply: true,
	});

	sheet.getRange(row, 13).setValue(username);
	sheet.getRange(row, 14).setValue(password);
	sheet.getRange(row, 15).setValue('Sent successfully');
	console.log(`Sent to ${email} (row ${row})`);
}

function generateHtmlBody(username, password) {
	const html = HtmlService.createHtmlOutputFromFile('inline').getContent();
	return html
		.replace(/\{username\}/g, escapeHtml(username))
		.replace(/\{password\}/g, escapeHtml(password));
}

function generatePlainText(name, surname, username, password) {
	const greeting = (name || surname) ? 'เรียนคุณ ' + (name + ' ' + surname).trim() + '\n\n' : '';
	return [
		greeting + 'ข้อมูลการเข้าสู่ระบบ Mock Test',
		'Mock x Triam The Trilogy for #TU90',
		'',
		'ข้อมูลบัญชีผู้ใช้',
		'Username: ' + username,
		'Password: ' + password,
		'',
		'รายละเอียดการสอบ',
		'เปิดระบบการสอบ: 15 พฤษภาคม 2569',
		'ปิดระบบการสอบ: 31 พฤษภาคม 2569',
		'',
		'เข้าสู่ระบบสอบ: https://auth.dugga.com/',
		'คู่มือการใช้ระบบการสอบ: https://drive.google.com/file/d/1wcg2npYd9e56uZzEPSE_bXZYknYxM85g/view?usp=sharing',
		'',
		'วิธีเข้าสู่ระบบ',
		'1. เข้าสู่เว็บไซต์ระบบการสอบออนไลน์',
		'2. กรอก Username และ Password ตามที่ระบุข้างต้น',
		'3. เริ่มทำข้อสอบตามวันและเวลาที่กำหนด',
		'',
		'ติดต่อสอบถามเพิ่มเติม',
		'อีเมล: admin@triamudomfamily.org',
		'Line Official Account: @triamudom',
		'',
		'ขอแสดงความนับถือ',
		'ทีมงานผู้ดูแลระบบ Mock Test',
		'สมาคมผู้ปกครองและครูโรงเรียนเตรียมอุดมศึกษา',
		'',
		'อีเมลฉบับนี้ถูกส่งโดยอัตโนมัติ กรุณาอย่าตอบกลับ',
	].join('\n');
}

function escapeHtml(text) {
	return String(text ?? '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}
