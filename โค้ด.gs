// Google Apps Script - Update รองรับ 3 อันดับ + ผลการคัดเลือก
const SHEET_ID = '17WNCt-dBXrO0c-mT18gfDCky0tx1eKCsYkFYUSw3TUw'; 
const SHEET_NAME = 'Students';

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: 'active', message: 'API V2 Ready' })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    let result = {};

    if (action === 'getAllStudents') result = getAllStudents();
    else if (action === 'createStudent') result = createStudent(request.data);
    else if (action === 'createStudentsBulk') result = createStudentsBulk(request.data);
    else if (action === 'updateStudent') result = updateStudent(request.data);
    else if (action === 'updateScoresBulk') result = updateScoresBulk(request.data);
    else if (action === 'deleteStudent') result = deleteStudent(request.id);
    else if (action === 'deleteStudentsBulk') result = deleteStudentsBulk(request.ids);

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // เพิ่มคอลัมน์ Choice 1-3 และ Admission Result
    sheet.appendRow(['ID','Exam ID','Full Name','Previous School','Grade Level','Thai','Math','Science','English','Aptitude','Total','Rank','National ID', 'Choice 1', 'Choice 2', 'Choice 3', 'Admission Result']);
  }
  return sheet;
}

function getAllStudents() {
  const sheet = getSheet();
  // อ่านถึงคอลัมน์ 17 (Q)
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 17).getDisplayValues();
  if (data.length <= 1) return { success: true, students: [] };
  
  const students = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    students.push({
      id: row[0],
      exam_id: String(row[1]).replace(/^'/, ''),
      full_name: row[2],
      previous_school: row[3],
      grade_level: row[4],
      thai_score: parseFloat(row[5]) || 0,
      math_score: parseFloat(row[6]) || 0,
      science_score: parseFloat(row[7]) || 0,
      english_score: parseFloat(row[8]) || 0,
      aptitude_score: parseFloat(row[9]) || 0,
      total_score: parseFloat(row[10]) || 0,
      rank: parseInt(row[11]) || 0,
      national_id: String(row[12] || '').replace(/^'/, ''),
      choice_1: row[13] || '',
      choice_2: row[14] || '',
      choice_3: row[15] || '',
      admission_result: row[16] || ''
    });
  }
  return { success: true, students: students };
}

function createStudent(data) {
  const sheet = getSheet();
  const id = Utilities.getUuid();
  const newRow = [
    id, "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level,
    data.thai_score || 0, data.math_score || 0, data.science_score || 0, data.english_score || 0,
    data.aptitude_score || 0, data.total_score || 0, data.rank || 0, "'" + (data.national_id || ''),
    data.choice_1 || '', data.choice_2 || '', data.choice_3 || '', data.admission_result || ''
  ];
  sheet.appendRow(newRow);
  return { success: true, message: 'บันทึกสำเร็จ', id: id };
}

function createStudentsBulk(students) {
  const sheet = getSheet();
  if (students.length === 0) return { success: false, message: 'ไม่พบข้อมูล' };
  const newRows = students.map(data => [
    Utilities.getUuid(), "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level,
    data.thai_score || 0, data.math_score || 0, data.science_score || 0, data.english_score || 0,
    data.aptitude_score || 0, data.total_score || 0, data.rank || 0, "'" + (data.national_id || ''),
    data.choice_1 || '', data.choice_2 || '', data.choice_3 || '', data.admission_result || ''
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 17).setValues(newRows);
  return { success: true, message: `นำเข้าสำเร็จ ${newRows.length} รายการ` };
}

function updateStudent(data) {
  const sheet = getSheet();
  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      sheet.getRange(i + 1, 1, 1, 17).setValues([[
        data.id, "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level,
        data.thai_score || 0, data.math_score || 0, data.science_score || 0, data.english_score || 0,
        data.aptitude_score || 0, data.total_score || 0, data.rank || 0, "'" + (data.national_id || ''),
        data.choice_1 || '', data.choice_2 || '', data.choice_3 || '', data.admission_result || ''
      ]]);
      return { success: true, message: 'อัปเดตสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

// updateScoresBulk, deleteStudent, deleteStudentsBulk ใช้ของเดิมได้ (แต่ต้องแก้เลขคอลัมน์ถ้ามีการแตะต้องแถว แต่ฟังก์ชันพวกนี้อิง ID หรือ Loop Row อยู่แล้ว จึงไม่ต้องแก้มาก ยกเว้น updateScoresBulk)

function updateScoresBulk(scoresData) {
  const sheet = getSheet();
  const range = sheet.getDataRange();
  const allValues = range.getValues(); 
  let studentMap = {};
  for (let i = 1; i < allValues.length; i++) {
    let eid = String(allValues[i][1]).replace(/^'/, '').trim();
    studentMap[eid] = i;
  }
  let updatedCount = 0;
  scoresData.forEach(item => {
    let targetId = String(item.exam_id).trim();
    if (studentMap.hasOwnProperty(targetId)) {
      let r = studentMap[targetId];
      let sc = parseFloat(item.science_score) || 0;
      let ma = parseFloat(item.math_score) || 0;
      let en = parseFloat(item.english_score) || 0;
      allValues[r][6] = ma; allValues[r][7] = sc; allValues[r][8] = en; allValues[r][10] = sc + ma + en;
      updatedCount++;
    }
  });
  if (updatedCount > 0) {
    range.setValues(allValues);
    return { success: true, message: `อัปเดตคะแนนสำเร็จ ${updatedCount} รายการ` };
  }
  return { success: false, message: 'ไม่พบรหัสผู้สอบที่ตรงกัน' };
}

function deleteStudent(id) {
  const sheet = getSheet();
  const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ลบสำเร็จ' };
    }
  }
  return { success: false, message: 'ไม่พบข้อมูล' };
}

function deleteStudentsBulk(idsToDelete) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  let deletedCount = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (idsToDelete.includes(data[i][0])) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  return deletedCount > 0 ? { success: true, message: `ลบสำเร็จ ${deletedCount} รายการ` } : { success: false, message: 'ไม่พบข้อมูล' };
}
