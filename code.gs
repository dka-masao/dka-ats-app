function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('ダイキアクシスATS');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- スプレッドシート全データ取得 ---
function getSpreadsheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const getSheetValues = (sheetName) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    const values = sheet.getDataRange().getDisplayValues();
    if (values.length <= 1) return [];
    
    const headers = values[0];
    return values.slice(1).map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] !== undefined ? String(row[index]) : "";
      });
      return obj;
    });
  };

  let selectionStatuses = getSheetValues("選考ステータスマスタ");
  if (selectionStatuses.length > 0) {
    selectionStatuses.sort((a, b) => {
      const idA = String(a.ステータスID || "");
      const idB = String(b.ステータスID || "");
      return idA.localeCompare(idB);
    });
  }

  // 🌟 システム設定表示用にスプレッドシートの情報を取得
  const settingsData = {
    spreadsheetUrl: ss.getUrl(),
    spreadsheetName: ss.getName()
  };

  return {
    candidates: getSheetValues("応募者"),
    jobs: getSheetValues("求人マスタ"),
    channels: getSheetValues("応募経路マスタ"), 
    selectionStatuses: selectionStatuses,
    systemMaster: getSheetValues("システムマスタ"), 
    history: getSheetValues("選考履歴"),
    settings: settingsData // 🌟 設定情報を追加
  };
}

// --- 応募者詳細情報の更新処理 ---
function updateCandidateInfo(canData) {
  return updateMasterRow("応募者", "応募者ID", canData.応募者ID, canData);
}

// --- アプリ画面からの新規候補者追加 ＆ リアルタイム自動採番 ---
function addNewCandidateFromApp(canData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("応募者");
    if (!sheet) throw new Error("応募者シートが見つかりません");
    const headers = sheet.getDataRange().getValues()[0];
    
    const dateStr = String(canData.応募日).trim();
    const yearMatch = dateStr.match(/^(\d{4})/);
    if (!yearMatch) throw new Error("応募日の日付形式が不正です");
    const yearStr = yearMatch[1].slice(-2);
    const typeCode = (canData.採用区分 === '新卒') ? 'N' : (canData.採用区分 === '中途' ? 'C' : '');
    const prefix = `${yearStr}-${typeCode}-`;
    
    const lastRow = sheet.getLastRow();
    let maxSeq = 0;
    if (lastRow > 1) {
      const idValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
      for (let i = 0; i < idValues.length; i++) {
        const currentId = String(idValues[i][0]);
        if (currentId.startsWith(prefix)) {
          const seqStr = currentId.replace(prefix, '');
          const seq = parseInt(seqStr, 10);
          if (!isNaN(seq) && seq > maxSeq) { maxSeq = seq; }
        }
      }
    }
    const generatedId = prefix + String(maxSeq + 1).padStart(4, '0');
    const defaultStatus = canData.採用区分 === '新卒' ? '①エントリー' : '選考中';
    const finalStatus = (canData.現在のステータス && canData.現在のステータス.trim() !== "") ? canData.現在のステータス : defaultStatus;
    
    const rowData = headers.map(header => {
      if (header === "応募者ID") return generatedId;
      if (header === "現在のステータス") return finalStatus;
      if (canData[header] !== undefined) return canData[header];
      return ""; 
    });
    sheet.appendRow(rowData);
    return { success: true, generatedId: generatedId };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// --- 💼 求人マスタの追加 ＆ 更新 ---
function addJobMaster(jobData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("求人マスタ");
    const lastRow = sheet.getLastRow();
    const nextId = "JOB-" + String(lastRow).padStart(3, '0');
    sheet.appendRow([nextId, jobData.求人タイトル, jobData.部署名, jobData.募集人数, jobData.ステータス]);
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}
function updateJobMaster(jobData) {
  return updateMasterRow("求人マスタ", "求人ID", jobData.求人ID, jobData);
}

// --- 📢 応募経路マスタの追加 ＆ 更新 ---
function addChannelMaster(chData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("応募経路マスタ");
    const lastRow = sheet.getLastRow();
    const nextId = "CH-" + String(lastRow).padStart(3, '0');
    sheet.appendRow([nextId, chData.経路名, chData.経路タイプ, chData.ステータス]);
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}
function updateChannelMaster(chData) {
  return updateMasterRow("応募経路マスタ", "応募経路ID", chData.応募経路ID, chData);
}

// --- 🗂️ 選考ステータスマスタの追加 ＆ 更新 ---
function addSelectionStatusMaster(statusData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("選考ステータスマスタ");
    const lastRow = sheet.getLastRow();
    const prefix = statusData.対象区分 === '新卒' ? 'ST-N-' : 'ST-C-';
    const nextId = prefix + String(lastRow).padStart(2, '0');
    sheet.appendRow([nextId, statusData.対象区分, statusData.ステータス名]);
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}
function updateSelectionStatusMaster(statusData) {
  return updateMasterRow("選考ステータスマスタ", "ステータスID", statusData.ステータスID, statusData);
}

// --- ⚙️ システムマスタの追加 ＆ 更新 ---
function addSystemMasterRow(sysData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("システムマスタ");
    sheet.appendRow([sysData.マスタ分類, sysData.コード, sysData.名称]);
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}
function updateSystemMasterRow(sysData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("システムマスタ");
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === sysData.old分類 && values[i][1] === sysData.oldコード) {
        sheet.getRange(i + 1, 1).setValue(sysData.マスタ分類);
        sheet.getRange(i + 1, 2).setValue(sysData.コード);
        sheet.getRange(i + 1, 3).setValue(sysData.名称);
        break;
      }
    }
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// --- 🧱 汎用マスタデータ上書きヘルパー関数 ---
function updateMasterRow(sheetName, idColumnName, idValue, dataObj) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(sheetName + "シートが見つかりません");
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const idIndex = headers.indexOf(idColumnName);
    
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idIndex]).trim() === String(idValue).trim()) {
        headers.forEach((header, colIdx) => {
          if (header !== idColumnName && dataObj[header] !== undefined) {
            sheet.getRange(i + 1, colIdx + 1).setValue(dataObj[header]);
          }
        });
        break;
      }
    }
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}

// --- 評価データの保存＆ステータス自動進行処理 ---
function saveInterviewResult(resultData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName("選考履歴");
    if (historySheet) {
      const headers = historySheet.getDataRange().getValues()[0];
      const rowData = headers.map(header => {
        if (header === "イベントID") return "EV-" + String(historySheet.getLastRow()).padStart(3, '0');
        if (header === "応募者ID") return resultData.candidateId;
        if (header === "選考フェーズ") return resultData.interviewType;
        if (header === "面談・面接日") return Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd");
        if (header === "対応者") return "採用担当者";
        if (header === "自分らしさの原点") return resultData.origin || "";
        if (header === "面談担当者まとめ") return resultData.comment || "";
        if (header === "学生からの質問・意見・感想") return resultData.studentQuestion || "";
        if (header === "知ったきっかけ") return resultData.trigger || "";
        if (header === "就職活動状況") return resultData.jobHuntingStatus || "";
        if (header === "その他採用課内情報") return resultData.internalInfo || "";
        if (header === "備考") return resultData.notes || "";
        if (header === "評価_身だしなみ") return resultData.ratings["身だしなみ"] || "";
        if (header === "評価_積極性") return resultData.ratings["積極性"] || "";
        if (header === "評価_態度") return resultData.ratings["態度"] || "";
        if (header === "評価_協調性") return resultData.ratings["協調性"] || "";
        if (header === "評価_対話") return resultData.ratings["対話"] || "";
        if (header === "評価_知識") return resultData.ratings["知識"] || "";
        if (header === "総合評価結果") {
          if (resultData.overallResult === 'ok') return '合格';
          if (resultData.overallResult === 'hold') return '再考察';
          if (resultData.overallResult === 'ng') return '不合格';
          return resultData.overallResult;
        }
        return "";
      });
      historySheet.appendRow(rowData);
    }
    
    const candidateSheet = ss.getSheetByName("応募者");
    if (candidateSheet) {
      const values = candidateSheet.getDataRange().getValues();
      const headers = values[0];
      const idIndex = headers.indexOf("応募者ID");
      const statusIndex = headers.indexOf("現在のステータス");
      const actionIndex = headers.indexOf("ネクストアクション");
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][idIndex] === resultData.candidateId) {
          let nextStatus = ""; let nextAction = "";
          if (resultData.overallResult === 'ok') {
            if (resultData.interviewType === '一次面談') { nextStatus = '⑤一次選考合格'; nextAction = '二次選考調整'; }
            else if (resultData.interviewType === '二次選考') { nextStatus = '⑦二次選考合格'; nextAction = '最終選考調整'; }
            else { nextStatus = '⑩内定通知'; nextAction = '内定承諾待ち'; }
          } else if (resultData.overallResult === 'hold') {
            nextStatus = '選考中（保留）'; nextAction = '社内協議';
          } else {
            nextStatus = '⑬選考不合格'; nextAction = '不合格通知送付';
          }
          if (statusIndex !== -1) candidateSheet.getRange(i + 1, statusIndex + 1).setValue(nextStatus);
          if (actionIndex !== -1) candidateSheet.getRange(i + 1, actionIndex + 1).setValue(nextAction);
          break;
        }
      }
    }
    return { success: true };
  } catch (err) { return { success: false, error: err.toString() }; }
}

function autoAssignApplicantId(e) {
  if (!e) return;
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== '応募者') return;
  const row = e.range.getRow();
  if (row <= 1) return;
  const idCell = sheet.getRange(row, 1);
  if (idCell.getValue() !== '') return;

  const dateStr = String(sheet.getRange(row, 2).getDisplayValue()).trim();
  const typeVal = String(sheet.getRange(row, 3).getDisplayValue()).trim();

  if (dateStr && typeVal) {
    const yearMatch = dateStr.match(/^(\d{4})/);
    if (!yearMatch) return;
    const yearStr = yearMatch[1].slice(-2);
    const typeCode = (typeVal === '新卒') ? 'N' : (typeVal === '中途' ? 'C' : '');
    if (!typeCode) return;
    const prefix = yearStr + '-' + typeCode + '-';
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) { idCell.setValue(prefix + "0001"); return; }
    
    const idValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    let maxSeq = 0;
    for (let i = 0; i < idValues.length; i++) {
      const currentId = String(idValues[i][0]);
      if (currentId.startsWith(prefix)) {
        const seqStr = currentId.replace(prefix, '');
        const seq = parseInt(seqStr, 10);
        if (!isNaN(seq) && seq > maxSeq) { maxSeq = seq; }
      }
    }
    const nextId = prefix + String(maxSeq + 1).padStart(4, '0');
    idCell.setValue(nextId);
  }
}
