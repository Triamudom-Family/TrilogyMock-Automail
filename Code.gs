function myFunction() {
	Logger.log(MailApp.getRemainingDailyQuota());
	// retryEmailByRow(546);
}

function findID(ID) {
	try {
		const idMap = getData();
		const key = String(ID).trim();

		if (Object.prototype.hasOwnProperty.call(idMap, key)) {
			const entry = idMap[key];
			return {
				found: true,
				row: entry.row,
				timestamp: entry.timestamp,
				name: entry.name,
				status: entry.status,
				notes: entry.notes,
			};
		}

		return { found: false };
	} catch (error) {
		return { found: false, error: error.message };
	}
}

function retryEmailByRow(row) {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses");
	if (!sheet) throw new Error('Response sheet not found');

	row = Number(row);
	if (!row || row < 2) throw new Error('Invalid row number');

	try {

	} catch (err) {
	}
}

function onFormSubmit(e) {
	const lock = LockService.getDocumentLock() || LockService.getScriptLock();
	lock.waitLock(30000);

	try {
		const formData = e.namedValues;
		const responseSheet = e.range.getSheet();
		const referenceSheet = responseSheet.getParent().getSheetByName('CurrentNumber');

		if (!referenceSheet) throw new Error('Sheet "CurrentNumber" not found');

		const track = (formData['แผนการเรียน']?.[0] || '').trim();

		if (!track) throw new Error("Track not specified in the form response");

		const counterCell = referenceSheet.getRange(3, getColumn(track));
		const membershipNumber = (Number(counterCell.getValue()) || 0) + 1;
		counterCell.setValue(membershipNumber);

		const membershipID = "89-" + String(membershipNumber);

		const row = e.range.getRow();
		responseSheet.getRange(row, 3).setValue(membershipID);

		const emailAddress = formData['Email Address']?.[0] || formData['ที่อยู่อีเมล']?.[0] || '';
		if (emailAddress) {
			const subject = 'ขอบคุณสำหรับการกรอกแบบฟอร์มสมัครสมาชิกและจองไซส์เสื้อ Triamudom Family';

			const emailData = {
				membershipID: membershipID,
				parentFirstName: formData["ชื่อผู้ปกครอง"]?.[0] || '',
				parentLastName: formData["นามสกุลผู้ปกครอง"]?.[0] || '',
				studentFirstName: formData["ชื่อนักเรียน"]?.[0] || '',
				studentLastName: formData["นามสกุลนักเรียน"]?.[0] || '',
				jerseySize: formData["ไซส์เสื้อ Jersey"]?.[0] || '-',
				jerseyText: formData["ตัวอักษรบนเสื้อ Jersey"]?.[0] || '-',
				jerseyNumber: formData["ตัวเลขบนเสื้อ Jersey"]?.[0] || '-',
				poloSize: formData["ไซส์เสื้อโปโล"]?.[0] || '-',
				poloGender: formData["เพศเสื้อโปโล"]?.[0] || '-'
			};

			const plainText = generatePlainText(emailData);
			const htmlBody = generateHtmlBody(emailData);

			GmailApp.sendEmail(emailAddress, subject, plainText, {
				htmlBody: htmlBody,
				name: 'ข้อมูลการสมัครสมาชิก',
				noReply: true
			});

			responseSheet.getRange(row, 32).setValue("TRUE");
		}

		responseSheet
			.getRange(row, 1, 1, responseSheet.getLastColumn())
			.setFontFamily('Anuphan');

		const monoRanges = ['A' + row, 'B' + row, 'C' + row, 'H' + row, 'J' + row, 'Q' + row, 'R' + row, 'S' + row, 'T' + row, 'AD' + row];
		responseSheet.getRangeList(monoRanges).setFontFamily('JetBrains Mono');

		responseSheet.getRange('A' + row).setHorizontalAlignment('right');
		responseSheet.getRangeList(['B' + row, 'AD' + row]).setHorizontalAlignment('left');
		responseSheet.getRangeList(['C' + row, 'H' + row, 'I' + row, 'J' + row, 'K' + row, 'L' + row, 'M' + row, 'Q' + row, 'R' + row, 'S' + row, 'T' + row, 'U' + row, 'AE' + row]).setHorizontalAlignment('center');
	} finally {
		lock.releaseLock();
	}
}

function generatePlainText(data) {
	return `
ขอบคุณสำหรับการสมัครสมาชิก

เรียนคุณ ${data.parentFirstName} ${data.parentLastName}

ขอบคุณสำหรับการกรอกแบบฟอร์มสมัครสมาชิก Triamudom Family และจองไซส์เสื้อ
ทางเราได้รับข้อมูลของท่านเรียบร้อยแล้ว

หมายเลขสมาชิก
${data.membershipID}

ข้อมูลสมาชิก
- ชื่อผู้ปกครอง: ${data.parentFirstName} ${data.parentLastName}
- ชื่อนักเรียน: ${data.studentFirstName} ${data.studentLastName}

ข้อมูลเสื้อ Jersey
- ไซส์เสื้อ Jersey: ${data.jerseySize}
- ชื่อบนเสื้อ Jersey: ${data.jerseyText}
- เบอร์บนเสื้อ Jersey: ${data.jerseyNumber}

ข้อมูลเสื้อโปโล
- ไซส์เสื้อโปโล: ${data.poloSize}
- เพศเสื้อโปโล: ${data.poloGender}

หากมีข้อมูลเพิ่มเติม ทางทีมงานจะแจ้งให้ท่านทราบอีกครั้ง

ขอขอบพระคุณอีกครั้ง

อีเมลฉบับนี้ถูกส่งโดยอัตโนมัติ กรุณาอย่าตอบกลับอีเมลนี้
	`.trim();
}

function generateHtmlBody(data) {
	return `
	`.trim();
}

function escapeHtml_(text) {
	return String(text ?? '')
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
		.replace(/'/g, '&#39;');
}

const MOCK_SHEET_NAME = 'Credentials';

const COL_EMAIL    = 1; // A — recipient email address
const COL_NAME     = 2; // B — student name (used in plain-text greeting)
const COL_USERNAME = 3; // C — Dugga username
const COL_PASSWORD = 4; // D — Dugga password
const COL_SENT     = 5; // E — "TRUE" once sent successfully
const COL_NOTES    = 6; // F — error messages / retry notes

function sendAllCredentials() {
	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MOCK_SHEET_NAME);
	if (!sheet) throw new Error('Sheet "' + MOCK_SHEET_NAME + '" not found');

	const lastRow = sheet.getLastRow();
	if (lastRow < 2) {
		console.log('No data rows found.');
		return;
	}

	const numCols = Math.max(COL_SENT, COL_NOTES);
	const data    = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
	let sent = 0, failed = 0;

	for (let i = 0; i < data.length; i++) {
		const row        = i + 2;
		const sentStatus = String(data[i][COL_SENT - 1] || '').trim();

		if (sentStatus === 'TRUE') continue;

		if (MailApp.getRemainingDailyQuota() < 1) {
			console.warn('Email quota exhausted. Stopped at row %s.', row);
			break;
		}

		try {
			mockSendRow_(row, sheet, data[i]);
			sent++;
		} catch (err) {
			sheet.getRange(row, COL_NOTES).setValue('Error: ' + err.message);
			console.error('Row %s failed: %s', row, err.message);
			failed++;
		}
	}

	console.log('Done. Sent: %s, Failed: %s', sent, failed);
}

function retrySendByRow(row) {
	row = Number(row);
	if (!row || row < 2) throw new Error('Invalid row number: ' + row);

	const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MOCK_SHEET_NAME);
	if (!sheet) throw new Error('Sheet "' + MOCK_SHEET_NAME + '" not found');

	const numCols = Math.max(COL_SENT, COL_NOTES);
	const rowData = sheet.getRange(row, 1, 1, numCols).getValues()[0];

	try {
		mockSendRow_(row, sheet, rowData);
		sheet.getRange(row, COL_NOTES).setValue('Retried OK');
		console.log('Retry succeeded for row %s', row);
	} catch (err) {
		sheet.getRange(row, COL_NOTES).setValue('Retry Failed: ' + err.message);
		throw err;
	}
}

function mockSendRow_(row, sheet, rowData) {
	if (MailApp.getRemainingDailyQuota() < 1) throw new Error('No email quota remaining today');

	const email    = String(rowData[COL_EMAIL    - 1] || '').trim();
	const name     = String(rowData[COL_NAME     - 1] || '').trim();
	const username = String(rowData[COL_USERNAME - 1] || '').trim();
	const password = String(rowData[COL_PASSWORD - 1] || '').trim();

	if (!email)    throw new Error('Missing email address');
	if (!username) throw new Error('Missing username');
	if (!password) throw new Error('Missing password');

	const subject   = 'ข้อมูลการเข้าสู่ระบบ Mock Test — Mock x Triam The Trilogy for #TU90';
	const htmlBody  = mockBuildHtmlBody_(username, password);
	const plainText = mockBuildPlainText_(username, password, name);

	GmailApp.sendEmail(email, subject, plainText, {
		htmlBody : htmlBody,
		name     : 'ระบบ Mock Test — Triamudom Family',
		noReply  : true,
	});

	sheet.getRange(row, COL_SENT ).setValue('TRUE');
	sheet.getRange(row, COL_NOTES).setValue('');
	console.log('Sent to %s (row %s)', email, row);
}

function mockBuildHtmlBody_(username, password) {
	const html = HtmlService.createHtmlOutputFromFile('inline').getContent();
	return html
		.replace(/\{username\}/g, escapeHtml_(username))
		.replace(/\{password\}/g, escapeHtml_(password));
}

function mockBuildPlainText_(username, password, name) {
	const greeting = name ? 'เรียนคุณ ' + name + '\n\n' : '';
	return [
		greeting + 'ข้อมูลการเข้าสู่ระบบ Mock Test',
		'Mock x Triam The Trilogy for #TU90',
		'',
		'Username : ' + username,
		'Password : ' + password,
		'',
		'เข้าสู่ระบบสอบ    : https://auth.dugga.com/',
		'เปิดระบบการสอบ : 10 พฤษภาคม 2569',
		'ปิดระบบการสอบ  : 31 พฤษภาคม 2569',
		'',
		'ติดต่อสอบถาม : admin@triamudomfamily.org | Line: @triamudom',
		'',
		'อีเมลฉบับนี้ถูกส่งโดยอัตโนมัติ กรุณาอย่าตอบกลับ',
	].join('\n');
}