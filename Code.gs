/**
 * ==============================================================================
 * 差勤自動化審核系統 (Attendance Review System) - Google Apps Script (GAS) 後端核心
 * ==============================================================================
 * 
 * 功能模組：
 * 1. Web App 介面提供 (doGet, doPost)
 * 2. 試算表自動結構化建置 (initSpreadsheet)
 * 3. Gemini AI 多模態 OCR 辨識 (手寫排班表、代班申請單、加退班單)
 * 4. 員工名單與綽號簡寫正規化對照 (Mapping Table + 動態 Prompt 注入)
 * 5. 特殊排班合併規則 (同日假日早班+假日晚班 -> 假日全班；早班+晚班 -> 全班；去除妝/髮/代班等雜訊)
 * 6. 打卡 CSV 解析與多維度差勤稽核結算 (代班覆蓋、加退班覆蓋、正職計分扣款、PT遲到扣0.5h、忘卡階梯扣假)
 * 7. 系統稽核日誌 (Audit Trail) 與自動定時備份 (Scheduled Backup)
 */

// ==========================================
// 1. 全域常數與設定
// ==========================================
const SCRIPT_PROP = PropertiesService.getScriptProperties();

/**
 * 取得 Gemini API Key (優先從 Script Properties 讀取，若無則從試算表「系統設定」頁讀取)
 */
function getGeminiApiKey() {
  let key = SCRIPT_PROP.getProperty('GEMINI_API_KEY');
  if (key && key.trim() !== '') return key.trim();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingSheet = ss.getSheetByName('系統設定');
    if (settingSheet) {
      const data = settingSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString().trim() === 'GEMINI_API_KEY') {
          key = data[i][1] ? data[i][1].toString().trim() : '';
          break;
        }
      }
    }
  } catch (e) {
    console.warn("從工作表讀取 API Key 失敗：" + e.toString());
  }

  return key || '';
}

/**
 * 取得使用的 Gemini 模型名稱 (預設 gemini-3.1-flash-lite)
 */
function getGeminiModel() {
  let model = SCRIPT_PROP.getProperty('GEMINI_MODEL');
  if (model && model.trim() !== '') return model.trim();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingSheet = ss.getSheetByName('系統設定');
    if (settingSheet) {
      const data = settingSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][0].toString().trim() === 'GEMINI_MODEL') {
          model = data[i][1] ? data[i][1].toString().trim() : '';
          break;
        }
      }
    }
  } catch (e) {}

  return model || 'gemini-3.1-flash-lite';
}

// ==========================================
// 2. Web App 入口與前後端路由
// ==========================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('差勤自動化審核系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;

    if (action === 'processCSV') {
      const res = processUploadedCSV(postData.csvText, postData.year, postData.month);
      return ContentService.createTextOutput(res).setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'getRules') {
      const res = getBusinessRules();
      return ContentService.createTextOutput(res).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: '未知操作' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. 系統初始化與工作表架構建置
// ==========================================

/**
 * 一鍵在當前試算表建立差勤系統的所有必要工作表與基礎資料
 */
function initSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 業務規則設定 (欄位分開：部門、職別；部門去識別化為 A部門、B部門)
    let rulesSheet = ss.getSheetByName('業務規則設定');
    if (!rulesSheet) {
      rulesSheet = ss.insertSheet('業務規則設定');
      rulesSheet.appendRow(['設定類別', '部門', '職別', '班別名稱', '預設上下班時間', '備註與計算規則']);
      rulesSheet.getRange('A1:F1').setBackground('#4A3934').setFontColor('#FFFFFF').setFontWeight('bold');
      
      const defaultRules = [
        ['班別設定', 'A部門', '正職', '晨班', '04:30 - 12:30', '正職常態班'],
        ['班別設定', 'A部門', '正職', '早班', '06:00 - 14:00', '正職常態班'],
        ['班別設定', 'A部門', '正職', '中班', '12:00 - 20:00', '正職常態班'],
        ['班別設定', 'A部門', '正職', '晚班', '14:30 - 22:30', '正職常態班'],
        ['班別設定', 'A部門', '正職', '假日早班', '05:00 - 14:00', '假日值班，不計入加班費'],
        ['班別設定', 'A部門', '正職', '假日晚班', '14:00 - 23:00', '假日值班，不計入加班費'],
        ['班別設定', 'A部門', '正職', '假日全班', '05:00 - 23:00', '假日值班，同天排假日早+假日晚自動合併'],
        ['班別設定', 'A部門', '正職', '假早班', '05:00 - 14:00', '假日早班別名'],
        ['班別設定', 'A部門', '正職', '假晚班', '14:00 - 23:00', '假日晚班別名'],
        ['班別設定', 'A部門', '正職', '假全班', '05:00 - 23:00', '假日全班別名'],
        ['班別設定', 'B部門', '正職', '早班', '08:30 - 17:30', '平日班'],
        ['班別設定', 'B部門', '正職', '晚班', '11:30 - 20:30', '平日班'],
        ['班別設定', 'B部門', '正職', '假日早班', '08:30 - 17:30', '假日班'],
        ['班別設定', 'B部門', '正職', '假日晚班', '14:00 - 23:00', '假日班'],
        ['班別設定', 'B部門', '正職', '假日全班', '08:30 - 23:00', '同天排假日早+假日晚自動合併'],
        ['班別設定', 'B部門', '計時 PT', '早班', '07:00 - 14:00', '遲到扣 0.5 小時時薪'],
        ['班別設定', 'B部門', '計時 PT', '晚班', '14:00 - 20:00', '遲到扣 0.5 小時時薪'],
        ['班別設定', 'B部門', '計時 PT', '全班', '07:00 - 20:00', '遲到扣 0.5 小時時薪'],
        ['班別設定', 'B部門', '計時 PT', '節目 1 班', '10:30 - 16:00', '允許調整起始時間（08:30 / 09:30）'],
        ['班別設定', 'B部門', '計時 PT', '節目 2 班', '13:00 - 16:00', '彈性支援班']
      ];
      rulesSheet.getRange(2, 1, defaultRules.length, 6).setValues(defaultRules);
      rulesSheet.autoResizeColumns(1, 6);
    }

    // 2. 系統設定
    let settingSheet = ss.getSheetByName('系統設定');
    if (!settingSheet) {
      settingSheet = ss.insertSheet('系統設定');
      settingSheet.appendRow(['設定項目', '設定值', '說明']);
      settingSheet.getRange('A1:C1').setBackground('#4A3934').setFontColor('#FFFFFF').setFontWeight('bold');
      const defaultSettings = [
        ['GEMINI_API_KEY', '', 'Google AI Studio API Key (若已在 Script Properties 設定可留空)'],
        ['GEMINI_MODEL', 'gemini-3.1-flash-lite', '預設使用 gemini-3.1-flash-lite'],
        ['BACKUP_FOLDER_ID', '1uorJ3Ii5u97IF7S2b4m6o5FGMjfFPqLy', '備份檔案放置的 Google 雲端硬碟資料夾 ID']
      ];
      settingSheet.getRange(2, 1, defaultSettings.length, 3).setValues(defaultSettings);
      settingSheet.autoResizeColumns(1, 3);
    }

    // 3. 員工名單 (含綽號簡寫) - 去識別化：不預載任何員工資料，由管理員於試算表自行建立
    let empSheet = ss.getSheetByName('員工名單');
    if (!empSheet) {
      empSheet = ss.insertSheet('員工名單');
      empSheet.appendRow(['員工姓名', '部門', '預設班別', '職別', '綽號簡寫 (逗號隔開)']);
      empSheet.getRange('A1:E1').setBackground('#4A3934').setFontColor('#FFFFFF').setFontWeight('bold');
      empSheet.autoResizeColumns(1, 5);
    }

    // 4. 排班資料庫
    let scheduleSheet = ss.getSheetByName('排班資料庫');
    if (!scheduleSheet) {
      scheduleSheet = ss.insertSheet('排班資料庫');
      scheduleSheet.appendRow(['排班日期', '員工姓名', '班別', '匯入時間']);
      scheduleSheet.getRange('A1:D1').setBackground('#4A3934').setFontColor('#FFFFFF').setFontWeight('bold');
      scheduleSheet.autoResizeColumns(1, 4);
    }

    // 5. 代班紀錄
    let swapSheet = ss.getSheetByName('代班紀錄');
    if (!swapSheet) {
      swapSheet = ss.insertSheet('代班紀錄');
      swapSheet.appendRow(['代班日期', '申請員工姓名', '代班人員', '調整後時間', '匯入時間']);
      swapSheet.getRange('A1:E1').setBackground('#4A3934').setFontColor('#FFFFFF').setFontWeight('bold');
      swapSheet.autoResizeColumns(1, 5);
    }

    // 6. 加退班紀錄
    let otSheet = ss.getSheetByName('加退班紀錄');
    if (!otSheet) {
      otSheet = ss.insertSheet('加退班紀錄');
      otSheet.appendRow(['加退班日期', '申請員工姓名', '遲到/應到時間', '早退/應退時間', '加班時數', '原因', '匯入時間']);
      otSheet.getRange('A1:G1').setBackground('#4A3934').setFontColor('#FFFFFF').setFontWeight('bold');
      otSheet.autoResizeColumns(1, 7);
    }

    // 7. 差勤結算總表
    let finalSheet = ss.getSheetByName('差勤結算總表');
    if (!finalSheet) {
      finalSheet = ss.insertSheet('差勤結算總表');
      finalSheet.appendRow([
        '員工姓名', '部門', '職別', '遲到總分鐘', '正職遲到扣款(元)', 
        'PT遲到次數(扣0.5h)', '忘卡次數', '應扣假天數', '加退班時數合計', '打卡出勤明細', '結算時間'
      ]);
      finalSheet.getRange('A1:K1').setBackground('#4A3934').setFontColor('#FFFFFF').setFontWeight('bold');
      finalSheet.autoResizeColumns(1, 11);
    }

    // 8. 系統日誌
    let logSheet = ss.getSheetByName('Log_系統日誌');
    if (!logSheet) {
      logSheet = ss.insertSheet('Log_系統日誌');
      logSheet.appendRow(['時間戳記', '操作者', '操作項目', '詳細內容', '狀態']);
      logSheet.getRange('A1:E1').setBackground('#6c757d').setFontColor('#FFFFFF').setFontWeight('bold');
      logSheet.autoResizeColumns(1, 5);
    }

    writeAuditLog('系統初始化', '成功建立所有差勤資料表結構', 'SUCCESS');
    return JSON.stringify({ status: 'success', message: '差勤系統資料表架構初始化完成！' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: '初始化失敗：' + e.toString() });
  }
}

/**
 * 清除動態資料 (只刪除：排班資料庫、代班紀錄、加退班紀錄、差勤結算總表之資料列)
 * 保留第一行標題列，保留業務規則設定與員工名單
 */
function clearTransactionData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheetNames = ['排班資料庫', '代班紀錄', '代班記錄', '加退班紀錄', '加退班記錄', '差勤結算總表'];
    let clearedCount = 0;
    const clearedSheets = [];

    targetSheetNames.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const lastRow = sheet.getLastRow();
        if (lastRow > 1) {
          sheet.deleteRows(2, lastRow - 1);
          clearedCount += (lastRow - 1);
        }
        if (!clearedSheets.includes(sheetName)) {
          clearedSheets.push(sheetName);
        }
      }
    });

    writeAuditLog('清除資料', `成功清除 ${clearedSheets.join('、')} 共 ${clearedCount} 筆資料`, 'SUCCESS');
    return JSON.stringify({
      status: 'success',
      message: `已清除排班資料庫、代班紀錄、加退班紀錄與差勤結算總表資料（共清除 ${clearedCount} 筆，標題列與員工名單規則完整保留）！`
    });
  } catch (e) {
    writeAuditLog('清除資料', `清除失敗：${e.toString()}`, 'ERROR');
    return JSON.stringify({ status: 'error', message: '清除資料失敗：' + e.toString() });
  }
}

// ==========================================
// 4. 稽核日誌與自動備份模組
// ==========================================

function writeAuditLog(action, detail, status = 'SUCCESS') {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Log_系統日誌');
    if (!logSheet) {
      logSheet = ss.insertSheet('Log_系統日誌');
      logSheet.appendRow(['時間戳記', '操作者', '操作項目', '詳細內容', '狀態']);
    }
    const user = Session.getActiveUser().getEmail() || 'Web使用者';
    logSheet.appendRow([new Date(), user, action, detail, status]);
  } catch (err) {
    console.error("寫入日誌失敗：" + err.toString());
  }
}

/**
 * 即時備份當前試算表至指定雲端硬碟資料庫位置
 * 雲端資料庫：https://drive.google.com/drive/folders/1uorJ3Ii5u97IF7S2b4m6o5FGMjfFPqLy
 */
function backupSpreadsheetNow() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const folderId = '1uorJ3Ii5u97IF7S2b4m6o5FGMjfFPqLy';
    const timeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Taipei', 'yyyyMMdd_HHmmss');
    const backupName = `Backup_${timeStr}_${ss.getName()}`;

    const currentFile = DriveApp.getFileById(ss.getId());
    const targetFolder = DriveApp.getFolderById(folderId);

    currentFile.makeCopy(backupName, targetFolder);
    writeAuditLog('即時備份', `成功備份至雲端資料庫：${backupName}`, 'SUCCESS');

    return JSON.stringify({
      status: 'success',
      message: `備份成功！已即時備份「${backupName}」至指定雲端資料庫。`
    });
  } catch (e) {
    writeAuditLog('即時備份', `備份失敗：${e.toString()}`, 'ERROR');
    return JSON.stringify({
      status: 'error',
      message: '備份失敗：' + e.toString()
    });
  }
}

/**
 * 每日定時自動異地/指定資料夾備份腳本 (可掛載 Time-driven 觸發器)
 */
function scheduledDailyBackup() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    const backupName = `Backup_${todayStr}_差勤資料庫`;

    // 檢查備份資料夾 ID
    let folderId = SCRIPT_PROP.getProperty('BACKUP_FOLDER_ID');
    if (!folderId) {
      const settingSheet = ss.getSheetByName('系統設定');
      if (settingSheet) {
        const data = settingSheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          if (data[i][0] === 'BACKUP_FOLDER_ID') {
            folderId = data[i][1] ? data[i][1].toString().trim() : '';
            break;
          }
        }
      }
    }
    if (!folderId) {
      folderId = '1uorJ3Ii5u97IF7S2b4m6o5FGMjfFPqLy';
    }

    const currentFile = DriveApp.getFileById(ss.getId());
    let targetFolder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    
    currentFile.makeCopy(backupName, targetFolder);
    writeAuditLog('自動定時備份', `成功建立備份：${backupName}`, 'SUCCESS');
    return "備份成功：" + backupName;
  } catch (e) {
    writeAuditLog('自動定時備份', `備份失敗：${e.toString()}`, 'ERROR');
    return "備份失敗：" + e.toString();
  }
}

// ==========================================
// 5. 字典檔與對照表存取
// ==========================================

/**
 * 取得員工資料對照表 (含簡寫與綽號反查正式全名)
 */
function getEmployeeMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('員工名單');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const empMap = {};

  for (let i = 1; i < data.length; i++) {
    const fullName = data[i][0] ? data[i][0].toString().trim() : "";
    if (!fullName) continue;

    const empObj = {
      fullName: fullName,
      dept: data[i][1] ? data[i][1].toString().trim() : "",
      defaultShift: data[i][2] ? data[i][2].toString().trim() : "",
      position: data[i][3] ? data[i][3].toString().trim() : "正職"
    };

    empMap[fullName] = empObj;

    // 處理綽號簡寫
    const nicknames = data[i][4] ? data[i][4].toString().trim() : "";
    if (nicknames) {
      nicknames.split(/[,，、\s]+/).forEach(alias => {
        const cleanAlias = alias.trim();
        if (cleanAlias) {
          empMap[cleanAlias] = empObj;
        }
      });
    }
  }

  return empMap;
}

/**
 * 取得員工簡寫字串清單 (注入至 Gemini Prompt 作為動態提示詞)
 */
function getEmployeeMappingPromptText() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('員工名單');
  if (!sheet) return "";

  const data = sheet.getDataRange().getValues();
  const mappingList = [];

  for (let i = 1; i < data.length; i++) {
    const fullName = data[i][0] ? data[i][0].toString().trim() : "";
    const nicknames = data[i][4] ? data[i][4].toString().trim() : "";
    if (fullName && nicknames) {
      nicknames.split(/[,，、\s]+/).forEach(alias => {
        const cleanAlias = alias.trim();
        if (cleanAlias && cleanAlias !== fullName) {
          mappingList.push(`${cleanAlias} -> ${fullName}`);
        }
      });
    }
  }

  return mappingList.length > 0 ? "【員工簡寫對照表】：\n" + mappingList.join(", ") : "";
}

/**
 * 取得班別上下班時間對照表 (格式：{ "部門_職別_班別": "HH:mm", "部門_班別": "HH:mm", "班別": "HH:mm" })
 */
function getShiftRulesMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('業務規則設定');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const rulesMap = {};

  for (let i = 1; i < data.length; i++) {
    // 欄位：[0:設定類別, 1:部門, 2:職別, 3:班別名稱, 4:預設上下班時間, 5:備註與計算規則]
    if (data[i][0] === '班別設定' && data[i][1] && data[i][3] && data[i][4]) {
      const dept = data[i][1].toString().trim();
      const position = data[i][2] ? data[i][2].toString().trim() : '';
      const shiftName = data[i][3].toString().trim();
      const timeStr = data[i][4].toString().trim();

      // 提取上班起始時間 (例如 "06:00 - 14:00" -> "06:00")
      const match = timeStr.match(/\d{2}:\d{2}/);
      if (match) {
        const startTime = match[0];
        if (dept && position) {
          rulesMap[`${dept}_${position}_${shiftName}`] = startTime;
        }
        if (dept) {
          rulesMap[`${dept}_${shiftName}`] = startTime;
        }
        rulesMap[shiftName] = startTime;

        // 假日班別雙向相容映射 (假早班 <-> 假日早班, 假晚班 <-> 假日晚班, 假全班 <-> 假日全班)
        const aliases = [];
        if (shiftName === '假全班') aliases.push('假日全班');
        if (shiftName === '假日全班') aliases.push('假全班');
        if (shiftName === '假早班') aliases.push('假日早班');
        if (shiftName === '假日早班') aliases.push('假早班');
        if (shiftName === '假晚班') aliases.push('假日晚班');
        if (shiftName === '假日晚班') aliases.push('假晚班');

        aliases.forEach(alias => {
          if (dept && position) rulesMap[`${dept}_${position}_${alias}`] = startTime;
          if (dept) rulesMap[`${dept}_${alias}`] = startTime;
          if (!rulesMap[alias]) rulesMap[alias] = startTime;
        });
      }
    }
  }

  return rulesMap;
}

/**
 * 前端查詢業務規則清單
 */
function getBusinessRules() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('業務規則設定');
    if (!sheet) return JSON.stringify({ status: 'error', message: '找不到「業務規則設定」工作表' });
    const data = sheet.getDataRange().getValues();
    return JSON.stringify({ status: 'success', data: data.slice(1) });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: e.toString() });
  }
}

/**
 * 取得已排定之排班對照表 (Key: `${day}_${fullName}`, Value: shiftName)
 */
function getScheduleMap(targetYear, targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('排班資料庫');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const schedMap = {};

  for (let i = 1; i < data.length; i++) {
    const dateCell = data[i][0];
    const name = data[i][1] ? data[i][1].toString().trim() : '';
    const shift = data[i][2] ? data[i][2].toString().trim() : '';
    if (!dateCell || !name) continue;

    const day = extractDay(dateCell, targetYear, targetMonth);
    if (day !== null) {
      schedMap[`${day}_${name}`] = mergeShiftList([shift]);
    }
  }

  return schedMap;
}

/**
 * 取得代班紀錄對照表 (Key: `${day}_${applicantFullName}`, Value: adjusted_time)
 */
function getSubstituteMap(targetYear, targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('代班紀錄');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const subMap = {};

  for (let i = 1; i < data.length; i++) {
    const dateCell = data[i][0];
    const applicant = data[i][1] ? data[i][1].toString().trim() : '';
    const adjustedTime = data[i][3] ? data[i][3].toString().trim() : '';
    if (!dateCell || !applicant) continue;

    const day = extractDay(dateCell, targetYear, targetMonth);
    if (day !== null && adjustedTime) {
      const match = adjustedTime.match(/\d{2}:\d{2}/);
      if (match) {
        subMap[`${day}_${applicant}`] = match[0];
      }
    }
  }

  return subMap;
}

/**
 * 取得加退班紀錄對照表 (Key: `${day}_${fullName}`, Value: { lateTime, earlyTime, otHours })
 */
function getOvertimeMap(targetYear, targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('加退班紀錄');
  if (!sheet) return {};

  const data = sheet.getDataRange().getValues();
  const otMap = {};

  for (let i = 1; i < data.length; i++) {
    const dateCell = data[i][0];
    const name = data[i][1] ? data[i][1].toString().trim() : '';
    if (!dateCell || !name) continue;

    const day = extractDay(dateCell, targetYear, targetMonth);
    if (day !== null) {
      otMap[`${day}_${name}`] = {
        lateTime: data[i][2] ? data[i][2].toString().trim() : "",
        earlyTime: data[i][3] ? data[i][3].toString().trim() : "",
        otHours: data[i][4] ? (parseFloat(data[i][4]) || 0) : 0
      };
    }
  }

  return otMap;
}

function extractDay(dateCell, filterYear, filterMonth) {
  if (dateCell instanceof Date) {
    if (filterYear && dateCell.getFullYear() !== parseInt(filterYear, 10)) return null;
    if (filterMonth && (dateCell.getMonth() + 1) !== parseInt(filterMonth, 10)) return null;
    return dateCell.getDate();
  }
  if (typeof dateCell === 'string') {
    const parts = dateCell.replace(/-/g, '/').split('/');
    if (parts.length === 3) {
      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      const d = parseInt(parts[2], 10);
      if (filterYear && y !== parseInt(filterYear, 10)) return null;
      if (filterMonth && m !== parseInt(filterMonth, 10)) return null;
      return d;
    }
    if (parts.length === 2) {
      return parseInt(parts[1], 10);
    }
    const num = parseInt(dateCell, 10);
    if (!isNaN(num) && num >= 1 && num <= 31) return num;
  }
  return null;
}

// ==========================================
// 6. Gemini AI OCR 辨識引擎
// ==========================================

/**
 * 調用 Gemini API 辨識照片
 * @param {string} base64Data - 圖片 Base64 字串
 * @param {string} mimeType - 圖片 MIME 格式 (如 image/jpeg)
 * @param {string} formType - 表單類型：'schedule' (排班), 'swap' (代班), 'overtime' (加退班)
 */
function processUploadedImage(base64Data, mimeType, formType = 'schedule') {
  const apiKey = getGeminiApiKey();
  if (!apiKey) {
    return JSON.stringify({ 
      status: 'error', 
      message: '尚未設定 GEMINI_API_KEY！請至試算表「系統設定」頁面填入，或在 Apps Script 專案屬性中設定。' 
    });
  }

  const model = getGeminiModel();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

  // 取得動態注入的員工對照表提示詞
  const dynamicMappingPrompt = getEmployeeMappingPromptText();

  let promptText = '';

  if (formType === 'swap') {
    promptText = `
你是一個專業的代班申請單 OCR 辨識助手。這是一張手寫或列印的「代班申請單」照片。
任務說明：
1. 照片最上方填寫的姓名代表「申請員工姓名」。
2. 請仔細辨識表格中「有實際填寫內容」的每一筆代班紀錄，【絕對忽略沒有填寫任何內容的空白列】。
3. 若代班人員欄位旁有英文名字或括號雜訊，請統一擷取正確名稱。
4. 特別注意：員工填寫時可能使用簡寫或綽號，請參照以下對照表自動正規化為正式全名：
${dynamicMappingPrompt}

請嚴格輸出為標準 JSON Array 格式，每個物件包含以下欄位：
- "date": 代班日期 (字串，例如 "2026/10/05" 或 "10/05")
- "applicant": 申請員工姓名 (字串)
- "substitute": 代班人員 (字串)
- "adjusted_time": 調整後時間 (字串，例如 "06:00 - 14:00" 或 "06:00")
`;
  } else if (formType === 'overtime') {
    promptText = `
你是一個專業的加退班申請單 OCR 辨識助手。這是一張「加退班申請單」照片。
任務說明：
1. 表單上可能分為多個區塊，每個區塊上方常有「姓名：某某某」。請將每一筆紀錄正確對應到該區塊的「申請員工姓名」。
2. 【絕對忽略完全未填寫資料的空白列】。
3. 將遲到/應到時間、早退/應退時間整理為乾淨的時間格式字串 (例如 "08:30")。
4. 若加班欄位出現負數 (例如 -4.5)，代表「早退扣除時數」，請務必記錄為負數數值；若為加班請記錄為正數；未填寫請設為 0。
5. 特別注意：員工常使用綽號或簡寫，請參照以下對照表自動正規化為全名：
${dynamicMappingPrompt}

請嚴格輸出為標準 JSON Array 格式，每個物件包含以下欄位：
- "date": 加退班日期 (字串，例如 "2026/10/08" 或 "10/08")
- "applicant": 申請員工姓名 (字串)
- "late_time": 遲到時間 (字串，無則留空 "")
- "expected_in": 應到時間 (字串，無則留空 "")
- "leave_early_time": 早退時間 (字串，無則留空 "")
- "expected_out": 應退時間 (字串，無則留空 "")
- "overtime_hours": 加退班時數 (數字，正數為加班，負數為早退扣除)
- "reason": 原因 (字串)
`;
  } else {
    // 預設為手寫排班表
    promptText = `
你是一個專業的手寫排班表 OCR 辨識助手。這是一張手寫或列印的班表照片（通常是矩陣式表格）。
任務說明：
1. 請仔細辨識照片中每一位員工的排班資料，找出所有「有實際填寫班別」的日期格子。
2. 忽略空白格子或無效劃記。
3. 特別注意：手寫排班表常使用員工綽號簡寫，請嚴格參照以下【員工簡寫對照表】自動替換為正式姓名：
${dynamicMappingPrompt}
若不在對照表內的新人名字，請保留照片上的原樣文字。
4. 【關鍵班別格式規則】：
   - 班別文字請去除「妝」、「髮」、「代班」等職別或註記雜訊。例如：「假日早班髮」請擷取為「假日早班」；「假日晚班代班妝」請擷取為「假日晚班」；「早班妝」請擷取為「早班」。
   - 請輸出標準乾淨的班別名稱。

請嚴格輸出為標準 JSON Array 格式，每個物件包含以下欄位：
- "date": 日期 (數字 1 到 31)
- "name": 員工姓名 (正規化後的全名或原樣)
- "shift": 班別名稱 (字串，例如 "晨班"、"早班"、"中班"、"晚班"、"假日早班"、"假日晚班"、"假日全班"、"全班"、"節目 1 班" 等)
`;
  }

  const payload = {
    contents: [
      {
        parts: [
          { text: promptText },
          {
            inline_data: {
              mime_type: mimeType,
              data: base64Data
            }
          }
        ]
      }
    ],
    generationConfig: {
      response_mime_type: "application/json"
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    if (responseCode !== 200) {
      return JSON.stringify({ status: 'error', message: `Gemini API 呼叫失敗 (HTTP ${responseCode})：${responseText}` });
    }

    const result = JSON.parse(responseText);
    if (!result.candidates || result.candidates.length === 0) {
      return JSON.stringify({ status: 'error', message: 'Gemini 未回傳有效辨識結果' });
    }

    const candidateText = result.candidates[0].content.parts[0].text;
    const parsedData = JSON.parse(candidateText);

    return JSON.stringify({
      status: 'success',
      data: parsedData,
      message: `成功辨識 ${parsedData.length} 筆資料`
    });
  } catch (err) {
    return JSON.stringify({ status: 'error', message: '圖片辨識過程發生例外錯誤：' + err.toString() });
  }
}

// ==========================================
// 7. 資料寫入與排班班別自動合併規則
// ==========================================

/**
 * 清理單一班別字串：去除「髮」、「妝」、「代班」等雜訊標記
 * 例如：
 * - 假日早班髮 -> 假日早班
 * - 假日早班妝 -> 假日早班
 * - 假日晚班代班妝 -> 假日晚班
 * - 假日晚班代班髮 -> 假日晚班
 */
function cleanSingleShift(shiftStr) {
  if (!shiftStr) return '';
  let s = shiftStr.toString().trim();
  if (!s) return '';

  // 1. 去除括號註記如 (代班)、（髮）、(妝) 等
  s = s.replace(/[\(（][^\)）]*[\)）]/g, '').trim();

  // 2. 去除代班、代班妝、代班髮、妝、髮等標記
  s = s.replace(/代班[妝髮]?/g, '');
  s = s.replace(/[妝髮]$/g, '');
  s = s.replace(/[妝髮]/g, '');
  s = s.replace(/代班/g, '');
  s = s.trim();

  // 3. 標準化名稱對映
  if (s === '假早' || s === '假早班') {
    s = '假日早班';
  } else if (s === '假晚' || s === '假晚班') {
    s = '假日晚班';
  } else if (s === '假全' || s === '假全班') {
    s = '假日全班';
  }

  return s;
}

/**
 * 核心班別清理與合併演算法：
 * 1. 把「圖一」格式（含 髮、妝、代班）正規化為「圖二」乾淨格式（多班別以「、」連接）
 * 2. 若同人同日同時有 (假日早班、假日晚班)，直接改為【假日全班】
 * 3. 若同日有 (早班、晚班)，合併為【全班】
 */
function mergeShiftList(shiftArray) {
  if (!shiftArray || !Array.isArray(shiftArray)) return '';

  const cleanTokens = [];
  shiftArray.forEach(item => {
    if (!item) return;
    const parts = item.toString().split(/[、,\/\s]+/);
    parts.forEach(p => {
      const cleaned = cleanSingleShift(p);
      if (cleaned && !cleanTokens.includes(cleaned)) {
        cleanTokens.push(cleaned);
      }
    });
  });

  if (cleanTokens.length === 0) return '';

  // 規則 2: 同日同時有 (假日早班、假日晚班)，直接改為【假日全班】
  const hasHolidayMorning = cleanTokens.includes('假日早班');
  const hasHolidayEvening = cleanTokens.includes('假日晚班');

  if (hasHolidayMorning && hasHolidayEvening) {
    const remaining = cleanTokens.filter(t => t !== '假日早班' && t !== '假日晚班');
    if (!remaining.includes('假日全班')) {
      remaining.unshift('假日全班');
    }
    return remaining.join('、');
  }

  // 規則：同日有 (早班、晚班)，合併為【全班】
  const hasMorning = cleanTokens.includes('早班');
  const hasEvening = cleanTokens.includes('晚班');
  if (hasMorning && hasEvening) {
    const remaining = cleanTokens.filter(t => t !== '早班' && t !== '晚班');
    if (!remaining.includes('全班')) {
      remaining.unshift('全班');
    }
    return remaining.join('、');
  }

  // 規則 1: 圖二乾淨顯示方式
  return cleanTokens.join('、');
}

/**
 * 儲存手寫排班表資料 (包含特殊合併邏輯與正規化)
 */
function saveScheduleData(scheduleArray, targetYear, targetMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('排班資料庫');
    if (!sheet) {
      initSpreadsheet();
      sheet = ss.getSheetByName('排班資料庫');
    }

    if (!scheduleArray || !Array.isArray(scheduleArray) || scheduleArray.length === 0) {
      return JSON.stringify({ status: 'error', message: '排班資料為空，無法寫入' });
    }

    const timestamp = new Date();
    const empMap = getEmployeeMap();

    // 1. 姓名二次正規化與雜訊過濾
    const validData = scheduleArray
      .filter(item => item.name && item.name.toString().trim() !== '' && item.name !== '圈')
      .map(item => {
        const rawName = item.name.toString().trim();
        const formalName = empMap[rawName] ? empMap[rawName].fullName : rawName;
        return {
          date: parseInt(item.date, 10),
          name: formalName,
          shift: item.shift ? item.shift.toString().trim() : ''
        };
      });

    // 2. 關鍵規則：同人同日班別清理與合併
    //    - 去除「髮/妝/代班」等標記（圖一 ➔ 圖二）
    //    - 若同時有 (假日早班、假日晚班)，改為「假日全班」
    //    - 早班 + 晚班 -> 合併為「全班」
    const shiftGroupMap = {};
    validData.forEach(item => {
      const key = `${item.date}_${item.name}`;
      if (!shiftGroupMap[key]) {
        shiftGroupMap[key] = {
          date: item.date,
          name: item.name,
          shifts: []
        };
      }
      if (item.shift) {
        shiftGroupMap[key].shifts.push(item.shift);
      }
    });

    const finalCleanData = Object.values(shiftGroupMap).map(group => {
      return {
        date: group.date,
        name: group.name,
        shift: mergeShiftList(group.shifts)
      };
    });

    // 3. 組裝寫入行
    const rowsToInsert = finalCleanData.map(item => {
      const formattedDate = `${targetYear}/${String(targetMonth).padStart(2, '0')}/${String(item.date).padStart(2, '0')}`;
      return [
        formattedDate,
        item.name,
        item.shift,
        timestamp
      ];
    });

    if (rowsToInsert.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    }

    writeAuditLog('匯入排班資料', `成功寫入 ${rowsToInsert.length} 筆排班資料 (${targetYear}/${targetMonth})`, 'SUCCESS');
    return JSON.stringify({
      status: 'success',
      message: `成功將 ${rowsToInsert.length} 筆排班寫入資料庫（已對照正式全名、去除妝/髮/代班雜訊，並合併假日全班）！`
    });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: '寫入排班資料庫失敗：' + e.toString() });
  }
}

/**
 * 一鍵整理「排班資料庫」中的歷史紀錄班別：
 * 1. 將圖一格式（含 髮、妝、代班）轉換為圖二乾淨格式
 * 2. 若同人同日包含 (假日早班、假日晚班)，直接轉換為【假日全班】
 */
function normalizeExistingScheduleSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('排班資料庫');
    if (!sheet) return JSON.stringify({ status: 'error', message: '找不到「排班資料庫」工作表' });

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return JSON.stringify({ status: 'success', message: '排班資料庫目前無資料需整理' });

    let updatedCount = 0;
    for (let i = 1; i < data.length; i++) {
      const rawShift = data[i][2] ? data[i][2].toString().trim() : '';
      if (rawShift) {
        const cleanedShift = mergeShiftList([rawShift]);
        if (cleanedShift !== rawShift) {
          data[i][2] = cleanedShift;
          updatedCount++;
        }
      }
    }

    if (updatedCount > 0) {
      sheet.getDataRange().setValues(data);
    }

    writeAuditLog('歷史班別整理', `成功整理 ${updatedCount} 筆排班紀錄（去除妝髮代班雜訊、假日早晚班改為假日全班）`, 'SUCCESS');
    return JSON.stringify({
      status: 'success',
      message: `成功整理 ${updatedCount} 筆歷史排班班別！已去除妝/髮/代班標記，並將 (假日早班、假日晚班) 自動更新為【假日全班】。`
    });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: '整理排班資料庫失敗：' + e.toString() });
  }
}

/**
 * 儲存代班單資料
 */
function saveSubstituteData(swapArray, targetYear, targetMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('代班紀錄');
    if (!sheet) {
      initSpreadsheet();
      sheet = ss.getSheetByName('代班紀錄');
    }

    const timestamp = new Date();
    const empMap = getEmployeeMap();

    const rowsToInsert = swapArray.map(item => {
      const rawApplicant = item.applicant ? item.applicant.toString().trim() : '';
      const applicantFull = empMap[rawApplicant] ? empMap[rawApplicant].fullName : rawApplicant;

      const rawSub = item.substitute ? item.substitute.toString().trim() : '';
      const subFull = empMap[rawSub] ? empMap[rawSub].fullName : rawSub;

      return [
        item.date,
        applicantFull,
        subFull,
        item.adjusted_time,
        timestamp
      ];
    });

    if (rowsToInsert.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    }

    writeAuditLog('匯入代班資料', `成功寫入 ${rowsToInsert.length} 筆代班紀錄`, 'SUCCESS');
    return JSON.stringify({ status: 'success', message: `成功寫入 ${rowsToInsert.length} 筆代班紀錄！` });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: '寫入代班紀錄失敗：' + e.toString() });
  }
}

/**
 * 儲存加退班單資料
 */
function saveOvertimeData(overtimeArray, targetYear, targetMonth) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('加退班紀錄');
    if (!sheet) {
      initSpreadsheet();
      sheet = ss.getSheetByName('加退班紀錄');
    }

    const timestamp = new Date();
    const empMap = getEmployeeMap();

    const rowsToInsert = overtimeArray.map(item => {
      const rawApplicant = item.applicant ? item.applicant.toString().trim() : '';
      const applicantFull = empMap[rawApplicant] ? empMap[rawApplicant].fullName : rawApplicant;

      return [
        item.date,
        applicantFull,
        item.late_time || item.expected_in || '',
        item.leave_early_time || item.expected_out || '',
        item.overtime_hours !== undefined ? item.overtime_hours : 0,
        item.reason || '',
        timestamp
      ];
    });

    if (rowsToInsert.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    }

    writeAuditLog('匯入加退班資料', `成功寫入 ${rowsToInsert.length} 筆加退班紀錄`, 'SUCCESS');
    return JSON.stringify({ status: 'success', message: `成功寫入 ${rowsToInsert.length} 筆加退班紀錄！` });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: '寫入加退班紀錄失敗：' + e.toString() });
  }
}

// 前端相容橋接名稱
function processSubstituteImage(base64Data, mimeType) {
  return processUploadedImage(base64Data, mimeType, 'swap');
}

function processOvertimeImage(base64Data, mimeType) {
  return processUploadedImage(base64Data, mimeType, 'overtime');
}

// ==========================================
// 8. CSV 打卡紀錄解析與結算引擎
// ==========================================

/**
 * 解析打卡 CSV 文字並執行稽核結算
 */
function processUploadedCSV(csvText, targetYear, targetMonth) {
  try {
    const rawData = Utilities.parseCsv(csvText);
    const structuredData = extractPunchData(rawData);
    const empMap = getEmployeeMap();

    // 1. 將打卡資料中的姓名比對回正式姓名
    const normalizedRecords = structuredData.map(record => {
      const formalName = empMap[record.name] ? empMap[record.name].fullName : record.name;
      return { ...record, name: formalName };
    });

    // 2. 彙整每日打卡記錄 (最早時間為上班卡，最晚時間為下班卡)
    const dailyRecords = analyzeDailyPunches(normalizedRecords);

    // 3. 讀取排班、代班、加退班及班別規則對照
    const scheduleMap = getScheduleMap(targetYear, targetMonth);
    const substituteMap = getSubstituteMap(targetYear, targetMonth);
    const overtimeMap = getOvertimeMap(targetYear, targetMonth);
    const rulesMap = getShiftRulesMap();

    // 4. 執行多維度差勤稽核計算
    const summaryData = calculateAttendance(dailyRecords, scheduleMap, substituteMap, overtimeMap, empMap, rulesMap);

    return JSON.stringify({
      status: 'success',
      message: `成功解析 ${dailyRecords.length} 筆每日出勤打卡，已綜合排班、代班覆蓋、加退班申請完成結算！`,
      summary: summaryData
    });
  } catch (error) {
    return JSON.stringify({ status: 'error', message: 'CSV 解析與結算失敗：' + error.toString() });
  }
}

/**
 * 門禁/打卡鐘 CSV 區塊解析器
 * 依據包含「工號：」與「姓名：」標籤定位員工區塊，欄位 1~31 依序對應 1~31 號
 */
function extractPunchData(data) {
  const records = [];
  let currentEmployee = "";
  let readingPunches = false;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const rowStr = row.join(" ");

    // 掃描是否包含「工號：」或「姓名：」
    if (rowStr.indexOf("工號：") !== -1 || rowStr.indexOf("姓名：") !== -1) {
      currentEmployee = "";
      for (let j = 0; j < row.length; j++) {
        const cell = row[j].toString();
        if (cell.includes("姓名：")) {
          currentEmployee = row[j + 1] ? row[j + 1].trim() : cell.replace("姓名：", "").trim();
          if (!currentEmployee) currentEmployee = cell.replace("姓名：", "").trim();
          break;
        }
      }
      readingPunches = true;
      i++; // 跳過日期標題列
      continue;
    }

    if (readingPunches && currentEmployee !== "") {
      // 若連續遇到整列空白，表示該員工區塊結束
      if (row.every(cell => cell.toString().trim() === "")) {
        readingPunches = false;
        continue;
      }

      for (let dayIndex = 0; dayIndex < 31; dayIndex++) {
        const cellData = row[dayIndex];
        if (cellData && cellData.toString().trim() !== "") {
          const times = cellData.toString().split('\n').map(t => t.trim()).filter(t => t !== "");
          if (times.length > 0) {
            let existingRecord = records.find(r => r.name === currentEmployee && r.date === (dayIndex + 1));
            if (existingRecord) {
              existingRecord.punches.push(...times);
            } else {
              records.push({ name: currentEmployee, date: dayIndex + 1, punches: [...times] });
            }
          }
        }
      }
    }
  }

  return records;
}

/**
 * 整理當日打卡：取最早為上班卡、最晚為下班卡，單筆判斷為缺卡
 */
function analyzeDailyPunches(structuredData) {
  return structuredData.map(record => {
    const sortedPunches = record.punches.sort();
    return {
      name: record.name,
      date: record.date,
      firstPunch: sortedPunches[0],
      lastPunch: sortedPunches.length > 1 ? sortedPunches[sortedPunches.length - 1] : "",
      isMissingPunch: (sortedPunches.length === 1)
    };
  });
}

/**
 * 差勤稽核計算核心邏輯
 */
function calculateAttendance(dailyRecords, scheduleMap, substituteMap, overtimeMap, empMap, rulesMap) {
  const summary = {};

  dailyRecords.forEach(record => {
    const emp = record.name;
    if (!summary[emp]) {
      const empInfo = empMap[emp] || { dept: "未分類", position: "正職" };
      summary[emp] = {
        name: emp,
        dept: empInfo.dept,
        position: empInfo.position,
        lateMins: 0,
        latePenalty: 0,
        ptLateCount: 0,
        missingCount: 0,
        leaveDeducted: 0,
        otHoursTotal: 0,
        punchLogs: []
      };
    }

    const empInfo = empMap[emp] || { dept: "", defaultShift: "", position: "正職" };
    const empDept = empInfo.dept;
    const isPT = (empInfo.position === "計時 PT" || empInfo.position === "計時PT" || empInfo.position === "PT");

    // 1. 決定基礎排班或預設班
    let empShift = scheduleMap[`${record.date}_${emp}`];
    if (!empShift && empInfo.defaultShift) {
      empShift = empInfo.defaultShift;
    }

    let shiftStart = null;
    if (empShift) {
      // 支援複合班別 (例如 晨班、中班)，優先以整串比對，若無則比對第一個班別
      const primaryShift = empShift.split('、')[0].trim();
      const shiftsToTry = empShift === primaryShift ? [empShift] : [empShift, primaryShift];

      for (const s of shiftsToTry) {
        if (empDept && empInfo.position) {
          shiftStart = rulesMap[`${empDept}_${empInfo.position}_${s}`];
        }
        if (!shiftStart && empDept) {
          shiftStart = rulesMap[`${empDept}_${s}`];
        }
        if (!shiftStart) {
          shiftStart = rulesMap[s];
        }
        if (shiftStart) break;
      }
    }

    // 2. 代班紀錄優先覆蓋上班時間
    const subStartTime = substituteMap[`${record.date}_${emp}`];
    if (subStartTime) {
      shiftStart = subStartTime;
      empShift = (empShift ? empShift + " " : "") + "(代班覆蓋)";
    }

    // 3. 加退班申請之應到時間覆蓋
    const otInfo = overtimeMap[`${record.date}_${emp}`];
    if (otInfo) {
      if (otInfo.lateTime) shiftStart = otInfo.lateTime;
      if (otInfo.otHours) summary[emp].otHoursTotal += otInfo.otHours;
    }

    // 組裝每日明細字串
    let dailyLogStr = `[${record.date}日] `;
    if (empShift) dailyLogStr += `(${empShift}) `;
    if (record.isMissingPunch) {
      dailyLogStr += `${record.firstPunch} (忘卡/缺卡)`;
    } else {
      dailyLogStr += `${record.firstPunch} ~ ${record.lastPunch}`;
    }
    summary[emp].punchLogs.push(dailyLogStr);

    // 忘卡累計
    if (record.isMissingPunch) {
      summary[emp].missingCount += 1;
    }

    // 計算遲到扣款與次數
    if (!record.isMissingPunch && shiftStart && record.firstPunch > shiftStart) {
      const lateMinutes = calculateMins(shiftStart, record.firstPunch);
      summary[emp].lateMins += lateMinutes;

      if (!isPT) {
        // 正職：每分鐘 10 元
        summary[emp].latePenalty += (lateMinutes * 10);
      } else {
        // 計時 PT：遲到直接扣 0.5 小時時薪 (累計 PT 遲到次數)
        summary[emp].ptLateCount += 1;
      }
    }
  });

  // 4. 忘卡扣假階梯式累計：滿 3 次扣 1 日休假；第 4 次起每次加扣 0.5 日
  for (const key in summary) {
    const mCount = summary[key].missingCount;
    if (mCount === 3) {
      summary[key].leaveDeducted = 1.0;
    } else if (mCount > 3) {
      summary[key].leaveDeducted = 1.0 + (mCount - 3) * 0.5;
    }
    summary[key].punchLogsStr = summary[key].punchLogs.join("； ");
  }

  return Object.values(summary);
}

function calculateMins(time1, time2) {
  const d1 = new Date("2026/01/01 " + time1);
  const d2 = new Date("2026/01/01 " + time2);
  const diff = Math.floor((d2 - d1) / 60000);
  return diff > 0 ? diff : 0;
}

/**
 * 將結算資料批次寫入「差勤結算總表」
 */
function saveAttendanceSummary(summaryData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('差勤結算總表');
    if (!sheet) {
      initSpreadsheet();
      sheet = ss.getSheetByName('差勤結算總表');
    }

    const timestamp = new Date();
    const rowsToInsert = summaryData.map(row => [
      row.name,
      row.dept || '',
      row.position || '',
      row.lateMins,
      row.latePenalty,
      row.ptLateCount || 0,
      row.missingCount,
      row.leaveDeducted,
      row.otHoursTotal || 0,
      row.punchLogsStr || (Array.isArray(row.punchLogs) ? row.punchLogs.join('； ') : ''),
      timestamp
    ]);

    if (rowsToInsert.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    }

    writeAuditLog('儲存差勤結算總表', `成功寫入 ${rowsToInsert.length} 位員工之差勤結算總表`, 'SUCCESS');
    return JSON.stringify({ status: 'success', message: '已成功匯出至 Google 試算表「差勤結算總表」！' });
  } catch (e) {
    return JSON.stringify({ status: 'error', message: '匯出至試算表失敗：' + e.toString() });
  }
}
