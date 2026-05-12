function onFormSubmit(e) {
	const sheet = e.range.getSheet();
	const row   = e.range.getRow();
	const rowData = sheet.getRange(row, 1, 1, 13).getValues()[0];
	sendEmail(row, sheet, rowData);
}

function retryEmailByRow(row) {
	row = Number(row);
	if (!row || row < 2) throw new Error('Invalid row number: ' + row);

	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses');
	if (!sheet) throw new Error('Sheet not found');

	const rowData = sheet.getRange(row, 1, 1, 13).getValues()[0];
	sendEmail(row, sheet, rowData);
	console.log('Retry succeeded for row %s', row);
}

function sendEmail(row, sheet, rowData) {
	if (MailApp.getRemainingDailyQuota() < 1) throw new Error('No email quota remaining today');

	const email = String(rowData[1] || '').trim();
	const name = String(rowData[2] || '').trim();
	const surname = String(rowData[3] || '').trim();

	if (!email) throw new Error('Missing email address');

	const subject = 'ข้อมูลการเข้าสู่ระบบ Mock Test — Mock x Triam The Trilogy for #TU90';
	const htmlBody = generateHtmlBody(name, surname);
	const plainText = generatePlainText(name, surname);

	GmailApp.sendEmail(email, subject, plainText, {
		htmlBody : htmlBody,
		name     : 'Mock x Triam The Trilogy',
		noReply  : true,
	});

	sheet.getRange(row, 13).setValue("Sent successfully");
	console.log(`Sent to ${email} (row ${row})`);
}

function generateHtmlBody(name, surname) {
	const html = HtmlService.createHtmlOutputFromFile('inline').getContent();
	return html
		.replace(/\{name\}/g,    escapeHtml(name))
		.replace(/\{surname\}/g, escapeHtml(surname));
}

function generatePlainText(name, surname) {
	const greeting = (name || surname) ? 'เรียนคุณ ' + (name + ' ' + surname).trim() + '\n\n' : '';
	return [
		greeting + 'ข้อมูลการเข้าสู่ระบบ Mock Test',
		'Mock x Triam The Trilogy for #TU90',
		'',
		'รายละเอียดการสอบ',
		'เปิดระบบการสอบ: 10 พฤษภาคม 2569',
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