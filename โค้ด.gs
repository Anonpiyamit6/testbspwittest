const SHEET_ID = '17WNCt-dBXrO0c-mT18gfDCky0tx1eKCsYkFYUSw3TUw'; // ตรวจสอบ ID Sheet
const SHEET_NAME = 'Students';

function doGet(e) { return ContentService.createTextOutput(JSON.stringify({ status: 'active' })).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.tryLock(10000);
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
  } catch (err) { return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() })).setMimeType(ContentService.MimeType.JSON); } finally { lock.releaseLock(); }
}

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // สร้าง Header 18 คอลัมน์
    sheet.appendRow(['ID','Exam ID','Full Name','Previous School','Grade Level','Thai','Math','Science','English','Aptitude','Total','Rank','National ID', 'Choice 1', 'Choice 2', 'Choice 3', 'Admission Result', 'Practical Score']);
  }
  return sheet;
}

function getAllStudents() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();
  
  // ถ้าไม่มีข้อมูลเลย ให้คืนค่าว่าง
  if (lastRow <= 1) return { success: true, students: [] };

  // บังคับอ่าน 18 คอลัมน์ (A ถึง R)
  const data = sheet.getRange(2, 1, lastRow - 1, 18).getDisplayValues();
  
  const students = data.map(row => ({
    id: row[0], 
    exam_id: String(row[1]).replace(/^'/, ''), 
    full_name: row[2], 
    previous_school: row[3], 
    grade_level: row[4],
    thai_score: parseFloat(row[5])||0, 
    math_score: parseFloat(row[6])||0, 
    science_score: parseFloat(row[7])||0, 
    english_score: parseFloat(row[8])||0, 
    aptitude_score: parseFloat(row[9])||0, 
    total_score: parseFloat(row[10])||0, 
    rank: parseInt(row[11])||0, 
    national_id: String(row[12]||'').replace(/^'/, ''),
    choice_1: row[13]||'', 
    choice_2: row[14]||'', 
    choice_3: row[15]||'', 
    admission_result: row[16]||'',
    practical_score: parseFloat(row[17])||0 // คอลัมน์ที่ 18 (index 17)
  }));
  
  return { success: true, students: students };
}

// ฟังก์ชันอื่นๆ คงเดิม (ตัดมาเฉพาะส่วนสำคัญเพื่อประหยัดพื้นที่ แต่ต้องมี create/update ครบ)
function createStudent(data) {
  const sheet = getSheet(); const id = Utilities.getUuid();
  const newRow = [id, "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level, data.thai_score||0, data.math_score||0, data.science_score||0, data.english_score||0, data.aptitude_score||0, data.total_score||0, data.rank||0, "'" + (data.national_id||''), data.choice_1||'', data.choice_2||'', data.choice_3||'', data.admission_result||'', data.practical_score||0];
  sheet.appendRow(newRow); return { success: true, message: 'บันทึกสำเร็จ', id: id };
}
function createStudentsBulk(students) {
  const sheet = getSheet(); if (students.length === 0) return { success: false, message: 'ไม่พบข้อมูล' };
  const newRows = students.map(data => [Utilities.getUuid(), "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level, data.thai_score||0, data.math_score||0, data.science_score||0, data.english_score||0, data.aptitude_score||0, data.total_score||0, data.rank||0, "'" + (data.national_id||''), data.choice_1||'', data.choice_2||'', data.choice_3||'', data.admission_result||'', data.practical_score||0]);
  sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 18).setValues(newRows); return { success: true, message: `นำเข้าสำเร็จ ${newRows.length} รายการ` };
}
function updateStudent(data) {
  const sheet = getSheet(); const allData = sheet.getDataRange().getValues();
  for (let i = 1; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      sheet.getRange(i + 1, 1, 1, 18).setValues([[data.id, "'" + data.exam_id, data.full_name, data.previous_school, data.grade_level, data.thai_score||0, data.math_score||0, data.science_score||0, data.english_score||0, data.aptitude_score||0, data.total_score||0, data.rank||0, "'" + (data.national_id||''), data.choice_1||'', data.choice_2||'', data.choice_3||'', data.admission_result||'', data.practical_score||0]]);
      return { success: true, message: 'อัปเดตสำเร็จ' };
    }
  } return { success: false, message: 'ไม่พบข้อมูล' };
}
function deleteStudent(id) { const sheet = getSheet(); const allData = sheet.getDataRange().getValues(); for (let i = 1; i < allData.length; i++) { if (allData[i][0] === id) { sheet.deleteRow(i + 1); return { success: true, message: 'ลบสำเร็จ' }; } } return { success: false, message: 'ไม่พบข้อมูล' }; }
function deleteStudentsBulk(ids) { const sheet = getSheet(); const data = sheet.getDataRange().getValues(); let count=0; for(let i=data.length-1; i>=1; i--){ if(ids.includes(data[i][0])){ sheet.deleteRow(i+1); count++; } } return count>0 ? {success:true, message:`ลบ ${count} รายการ`} : {success:false, message:'ไม่พบข้อมูล'}; }
function updateScoresBulk(d){ const s=getSheet(); const r=s.getDataRange(); const v=r.getValues(); let m={}; for(let i=1;i<v.length;i++) m[String(v[i][1]).replace(/^'/,'').trim()]=i; let c=0; d.forEach(x=>{ let k=String(x.exam_id).trim(); if(m[k]){ let i=m[k]; v[i][6]=x.math_score; v[i][7]=x.science_score; v[i][8]=x.english_score; v[i][17]=x.practical_score||0; v[i][10]=parseFloat(x.math_score)+parseFloat(x.science_score)+parseFloat(x.english_score)+(parseFloat(x.practical_score)||0); c++; } }); if(c>0){r.setValues(v); return{success:true, message:`อัปเดต ${c} รายการ`};} return{success:false, message:'ไม่พบข้อมูล'}; }
