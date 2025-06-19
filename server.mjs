// --- Polyfill for import GenerativeAI ---
import fetch, { Headers, Request, Response } from "node-fetch";
globalThis.fetch = fetch;
globalThis.Headers = Headers;
globalThis.Request = Request;
globalThis.Response = Response;
import express from "express";
import line from "@line/bot-sdk";
import { google } from "googleapis";
import { GoogleGenerativeAI } from "@google/generative-ai";
import axios from "axios";
import fs from "fs";

const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY);
const model = genAI.getGenerativeModel({ model: "gemini-pro" });

// ========== LINE BOT 設定 ==========
const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET
};
const client = new line.Client(config);

// ========== Google Sheets 設定 ==========
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
const STAFF_SHEET = process.env.STAFF_SHEET;
const ATTEND_SHEET = process.env.ATTEND_SHEET;
const HOLIDAY_SHEET = process.env.HOLIDAY_SHEET;
const BONUS_SHEET = process.env.BONUS_SHEET;
const DISASTER_SHEET = process.env.DISASTER_SHEET;
const credentials = JSON.parse(
  Buffer.from(process.env.GOOGLE_SERVICE_JSON_BASE64, 'base64').toString('utf-8')
);
const SCOPES = ["https://www.googleapis.com/auth/spreadsheets"];
const auth = new google.auth.GoogleAuth({
  credentials,
  scopes: SCOPES
});
const sheets = google.sheets({ version: "v4", auth });

// ========== Google Sheets 範圍 ==========
function pad(n) { return n < 10 ? "0" + n : n; }
function isWeekend(dateStr) { const d = new Date(dateStr); return d.getDay() === 0 || d.getDay() === 6; }
function today() { return new Date().toISOString().slice(0, 10); }
function nowStr() { return new Date().toLocaleString("zh-TW", { timeZone: "Asia/Taipei" }); }

// 補這一行
function getColIdx(header, colName) {
  return header.indexOf(colName);
}

// 國定假日自動同步，建議啟動時和每24hr自動跑一次
async function autoSyncTaiwanHolidays() {
  const year = (new Date()).getFullYear();
  const nextYear = year + 1;
  const fetchYear = async (y) => {
    try {
      const res = await axios.get(`https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/${y}.json`);
      return res.data;
    } catch { return []; }
  };
  const data1 = await fetchYear(year);
  const data2 = await fetchYear(nextYear);
  const existing = (await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID, range: HOLIDAY_SHEET + "!A2:A"
  })).data.values?.flat() || [];
  const needAdd = [...data1, ...data2]
    .filter(h => h.isHoliday)
    .filter(h => !existing.includes(h.date))
    .map(h => [h.date, h.name]);
  if (needAdd.length > 0) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID, range: HOLIDAY_SHEET + "!A1",
      valueInputOption: "USER_ENTERED",
      resource: { values: needAdd }
    });
  }
}

// AI 分析意圖（使用 Gemini）
async function getIntentByAI(msg) {
  try {
    const model = genAI.getGenerativeModel({ model: "gemini-1.5-pro" });
    const chat = model.startChat();

    const prompt = `
你是「品禾設計智慧出勤AI」，用 JSON 結構回應：
{
 "intent":"請假|打卡|外出|查詢薪資|新增天災假|新增獎金|其它",
 "假別":"事假",
 "日期":["2025-07-01","2025-07-02"],
 "說明":"我要陪家人"
}
使用者輸入：「${msg}」
如果資訊不足，請回「[補問]內容」
`;

    const result = await chat.sendMessage(prompt);
    const response = await result.response.text();

    if (/^\[補問\]/.test(response)) {
      return { intent: "補問", text: response.replace(/^\[補問\]/, "") };
    }

    return JSON.parse(response);
  } catch (err) {
    console.error("Gemini intent error:", err);
    return { intent: "其它" };
  }
}

function parseDateRange(datestr) {
  if (!datestr) return [];
  if (datestr.includes("~")) {
    const [start, end] = datestr.split("~");
    const sd = new Date(start), ed = new Date(end);
    let arr = [];
    for (let d = new Date(sd); d <= ed; d.setDate(d.getDate() + 1)) {
      arr.push(d.toISOString().slice(0, 10));
    }
    return arr;
  }
  // 只單日
  const m = datestr.match(/(\d{4}-\d{2}-\d{2})(\(.+\))?/);
  return m ? [m[1] + (m[2] || "")] : [datestr];
}

// ========== 歡迎說明 ==========
const welcomeMsg = `歡迎使用品禾LINE出勤系統！
以下為使用說明：
【註冊】輸入：「註冊+你的姓名」
【打卡】輸入：「上班」、「下班」、「打卡」
【請假】
　- 輸入「請假」開始（依指示回覆）
　- 或直接輸入「請假 事假 2025-07-01~2025-07-03」
 
【外出紀錄】
　- 輸入「外出」、「到工地」、「離開工地」自動記錄時間（需加地點說明）
【查薪資】
　- 輸入「薪資」查詢本月
　- 輸入「X月薪資」查詢指定月份
【查請假】輸入：「查請假」、「我的請假」

＊如需說明，輸入「說明」或「help」取得本訊息。`;

// ========== 主程式 ==========
const app = express();
const sessionMap = new Map();

app.post("/webhook", line.middleware(config), async (req, res) => {
  await autoSyncTaiwanHolidays(); // 開頭自動同步國定假日
  await Promise.all(req.body.events.map(event => smartHandleEvent(event)));
  res.send("ok");
});

async function smartHandleEvent(event) {
  if (event.type !== "message" || event.message.type !== "text") return;
  const userId = event.source.userId;
  const msg = event.message.text.trim();

  // ====== 取得使用者資料 ======
  const staffSheet = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: STAFF_SHEET + "!A1:Z" });
  const staffHeader = staffSheet.data.values[0];
  const staffRows = staffSheet.data.values.slice(1);
  const staffIdx = staffRows.findIndex(r => r[staffHeader.indexOf("LINE_ID")] === userId);
  const staffInfo = staffRows[staffIdx];
  const adminUserIds = staffRows.filter(r => r[staffHeader.indexOf("管理員")] === "是").map(r => r[staffHeader.indexOf("LINE_ID")]);
  function isAdmin(uid) { return adminUserIds.includes(uid); }

  // ====== AI意圖分流 ======
  let intentObj = await getIntentByAI(msg);

  // ====== 註冊 ======
  if (/^註冊/.test(msg)) {
    const name = msg.replace("註冊", "").trim();
    if (!staffInfo) {
      let newRow = [];
      newRow[getColIdx(staffHeader, "LINE_ID")] = userId;
      newRow[getColIdx(staffHeader, "姓名")] = name;
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID, range: STAFF_SHEET + "!A1",
        valueInputOption: "USER_ENTERED", resource: { values: [newRow] }
      });
      return client.replyMessage(event.replyToken, { type: "text", text: `註冊成功！請通知管理員審核。\n${welcomeMsg}` });
    }
    return client.replyMessage(event.replyToken, { type: "text", text: "你已註冊。" });
  }

  // ====== 說明 ======
  if (/^(hi|hello|您好|你好|help|說明|幫助)$/i.test(msg)) {
    return client.replyMessage(event.replyToken, { type: "text", text: welcomeMsg });
  }

  // ========== 打卡 ==========
  if (intentObj.intent === "打卡") {
    const todayStr = today(), time = nowStr();
    const holidayRows = (await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: HOLIDAY_SHEET + "!A2:A" })).data.values?.flat() || [];
    const disasterRows = (await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: DISASTER_SHEET + "!A2:A" })).data.values?.flat() || [];
    if (holidayRows.includes(todayStr) || disasterRows.includes(todayStr) || isWeekend(todayStr))
      return client.replyMessage(event.replyToken, { type: "text", text: "今天是假日/天災假/週末，不用打卡！" });

    const attendSheet = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1:N" });
    const attendHeader = attendSheet.data.values[0], attendRows = attendSheet.data.values.slice(1);
    const colIdx = n => attendHeader.indexOf(n);
    const todayIdx = attendRows.findIndex(r => r[colIdx("LINE_ID")] === userId && r[colIdx("日期")] === todayStr);
    if (todayIdx >= 0) {
      let row = attendRows[todayIdx];
      if (row[colIdx("請假")] === "V") return client.replyMessage(event.replyToken, { type: "text", text: "你今天請假不用打卡。" });
      if (!row[colIdx("下班時間")]) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `${ATTEND_SHEET}!${String.fromCharCode(colIdx("下班時間") + 65)}${todayIdx + 2}`,
          valueInputOption: "USER_ENTERED", resource: { values: [[time]] }
        });
        return client.replyMessage(event.replyToken, { type: "text", text: `下班打卡完成：${time}` });
      }
      return client.replyMessage(event.replyToken, { type: "text", text: "今日已完成打卡。" });
    } else {
      let newRow = [];
      newRow[colIdx("LINE_ID")] = userId;
      newRow[colIdx("姓名")] = staffInfo ? staffInfo[staffHeader.indexOf("姓名")] : "";
      newRow[colIdx("日期")] = todayStr;
      newRow[colIdx("上班時間")] = time;
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1",
        valueInputOption: "USER_ENTERED", resource: { values: [newRow] }
      });
      return client.replyMessage(event.replyToken, { type: "text", text: `上班打卡完成：${time}` });
    }
  }
  // ========== AI 補問 ==========
  if (intentObj.intent === "補問") {
    return client.replyMessage(event.replyToken, { type: "text", text: intentObj.text });
  }
  // ========== 外出 ==========
  if (intentObj.intent === "外出") {
    const nowDate = today();
    const note = intentObj.說明 || msg;
    const attendSheet = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1:N" });
    const attendHeader = attendSheet.data.values[0], attendRows = attendSheet.data.values.slice(1);
    const colIdx = n => attendHeader.indexOf(n);
    let idx = attendRows.findIndex(r => r[colIdx("LINE_ID")] === userId && r[colIdx("日期")] === nowDate);
    if (idx >= 0) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `${ATTEND_SHEET}!${String.fromCharCode(colIdx("外出說明") + 65)}${idx + 2}`,
        valueInputOption: "USER_ENTERED", resource: { values: [[note || "外出"]] }
      });
      return client.replyMessage(event.replyToken, { type: "text", text: "外出紀錄已登記。" });
    } else {
      let newRow = [];
      newRow[colIdx("LINE_ID")] = userId;
      newRow[colIdx("姓名")] = staffInfo ? staffInfo[staffHeader.indexOf("姓名")] : "";
      newRow[colIdx("日期")] = nowDate;
      newRow[colIdx("外出說明")] = note || "外出";
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1",
        valueInputOption: "USER_ENTERED", resource: { values: [newRow] }
      });
      return client.replyMessage(event.replyToken, { type: "text", text: "外出紀錄已登記。" });
    }
  }

  // ========== 請假 ==========
  if (intentObj.intent === "請假") {
    let dateList = intentObj.日期 || parseDateRange(msg.match(/\d{4}-\d{2}-\d{2}(~\d{4}-\d{2}-\d{2})?/g)?.[0]);
    // 自動排除國定假日/天災/週末
    const holidayRows = (await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: HOLIDAY_SHEET + "!A2:A" })).data.values?.flat() || [];
    const disasterRows = (await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: DISASTER_SHEET + "!A2:A" })).data.values?.flat() || [];
    let validDates = (dateList || []).filter(d => !holidayRows.includes(d) && !disasterRows.includes(d) && !isWeekend(d));
    if (!validDates.length) return client.replyMessage(event.replyToken, { type: "text", text: "全部日期都是國定假日、天災假或週末，不用請假！" });
    // 寫入Google Sheet出勤表
    const attendSheet = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1:N" });
    const attendHeader = attendSheet.data.values[0];
    for (let d of validDates) {
      let newRow = [];
      newRow[attendHeader.indexOf("LINE_ID")] = userId;
      newRow[attendHeader.indexOf("姓名")] = staffInfo ? staffInfo[staffHeader.indexOf("姓名")] : "";
      newRow[attendHeader.indexOf("日期")] = d;
      newRow[attendHeader.indexOf("請假")] = "V";
      newRow[attendHeader.indexOf("假別說明")] = intentObj.假別 || "";
      newRow[attendHeader.indexOf("請假狀態")] = "待審核";
      newRow[attendHeader.indexOf("說明")] = intentObj.說明 || "";
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1",
        valueInputOption: "USER_ENTERED", resource: { values: [newRow] }
      });
    }
    // 推播主管
    for (let adminId of adminUserIds) {
      await client.pushMessage(adminId, {
        type: "text",
        text: `[請假審核] ${staffInfo ? staffInfo[staffHeader.indexOf("姓名")] : ""} 申請 ${intentObj.假別}\n日期：${validDates.join("、")}\n說明：${intentObj.說明 || ""}\n請回覆「准假 張三 2025-07-01」或「需商議 張三 2025-07-01」`
      });
    }
    return client.replyMessage(event.replyToken, { type: "text", text: `請假已登記，日期：${validDates.join("、")}，待審核。` });
  }

  // ====== 請假審核（管理員） ======
  if (adminUserIds.includes(userId) && /^(准假|需商議)\s+/.test(msg)) {
    const arr = msg.split(/\s+/);
    if (arr.length < 3) return client.replyMessage(event.replyToken, { type: "text", text: "請用：准假 張三 2025-07-01" });
    const action = arr[0], name = arr[1], date = arr[2];
    const attendSheet = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1:N" });
    const attendHeader = attendSheet.data.values[0], attendRows = attendSheet.data.values.slice(1);
    const colIdx = n => getColIdx(attendHeader, n);
    let idx = attendRows.findIndex(r => r[colIdx("姓名")] === name && r[colIdx("日期")] === date && r[colIdx("請假狀態")] === "待審核");
    if (idx === -1) return client.replyMessage(event.replyToken, { type: "text", text: "查無該請假申請。" });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${ATTEND_SHEET}!${String.fromCharCode(colIdx("請假狀態") + 65)}${idx + 2}`,
      valueInputOption: "USER_ENTERED", resource: { values: [[action === "准假" ? "已通過" : "需商議"]] }
    });
    let userLineId = staffRows.find(r => r[getColIdx(staffHeader, "姓名")] === name)?.[getColIdx(staffHeader, "LINE_ID")];
    if (userLineId) {
      await client.pushMessage(userLineId, {
        type: "text",
        text: `您的${date}請假：${action === "准假" ? "已通過" : "需商議"}`
      });
    }
    return client.replyMessage(event.replyToken, { type: "text", text: "已審核。" });
  }

  // ====== 查詢我的請假紀錄 ======
  if (/^(查請假|我的請假)/.test(msg)) {
    const attendSheet = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1:N" });
    const attendHeader = attendSheet.data.values[0], attendRows = attendSheet.data.values.slice(1);
    const colIdx = n => attendHeader.indexOf(n);
    let mine = attendRows.filter(r => r[colIdx("LINE_ID")] === userId && r[colIdx("請假")] === "V");
    if (mine.length === 0) return client.replyMessage(event.replyToken, { type: "text", text: "查無請假紀錄。" });
    let txt = mine.map(r =>
      `${r[colIdx("日期")]}：${r[colIdx("假別說明")] || ""}，狀態：${r[colIdx("請假狀態")] || "—"}`
    ).join("\n");
    return client.replyMessage(event.replyToken, { type: "text", text: `你的請假紀錄：\n${txt}` });
  }

  // ====== 新增獎金 ======
  if (/^新增獎金\s+/.test(msg) && adminUserIds.includes(userId)) {
    const arr = msg.split(/\s+/);
    if (arr.length < 5) return client.replyMessage(event.replyToken, { type: "text", text: "請用：新增獎金 姓名 2025-07 12000 說明" });
    const name = arr[1], month = arr[2], amt = arr[3], note = arr.slice(4).join(' ');
    const staff = staffRows.find(r => r[getColIdx(staffHeader, "姓名")] === name);
    if (!staff) return client.replyMessage(event.replyToken, { type: "text", text: "查無該員工姓名" });
    const lineId = staff[getColIdx(staffHeader, "LINE_ID")];
    const bonusSheet = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: BONUS_SHEET + "!A1:G" });
    const bonusHeader = bonusSheet.data.values[0], bonusRows = bonusSheet.data.values.slice(1);
    let exists = bonusRows.find(r => r[bonusHeader.indexOf("LINE_ID")] === lineId && r[bonusHeader.indexOf("月份")] === month);
    if (exists) return client.replyMessage(event.replyToken, { type: "text", text: "本月已登錄，不可重複！" });
    let newRow = [];
    newRow[bonusHeader.indexOf("LINE_ID")] = lineId;
    newRow[bonusHeader.indexOf("職稱")] = staff[getColIdx(staffHeader, "職稱")];
    newRow[bonusHeader.indexOf("姓名")] = name;
    newRow[bonusHeader.indexOf("月份")] = month;
    newRow[bonusHeader.indexOf("獎金金額")] = amt;
    newRow[bonusHeader.indexOf("說明")] = note;
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID, range: BONUS_SHEET + "!A1",
      valueInputOption: "USER_ENTERED", resource: { values: [newRow] }
    });
    return client.replyMessage(event.replyToken, { type: "text", text: "已新增獎金紀錄。" });
  }

  // ====== 查詢薪資 ======
  if (/薪資|薪水|(\d{1,2}|[一二三四五六七八九十十一十二])月薪資/.test(msg)) {
    const now = new Date();
    let year = now.getFullYear(), month = now.getMonth();
    let match = msg.match(/(\d{4})[年/-]?(\d{1,2})月?/);
    if (match) {
      year = parseInt(match[1], 10); month = parseInt(match[2], 10) - 1;
    } else {
      match = msg.match(/([一二三四五六七八九十十一十二])月/);
      if (match) {
        const zhMap = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10, 十一: 11, 十二: 12 };
        month = zhMap[match[1]] - 1;
      }
    }
    if (!match) { if (month === 0) { year -= 1; month = 11; } else { month -= 1; } }
    const startDay = `${year}-${pad(month + 1)}-01`;
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const endDay = `${year}-${pad(month + 1)}-${daysInMonth}`;
    const readResult = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1:N" });
    const rows = readResult.data.values || [];
    const myRows = rows.filter(row => row[0] === userId && row[2] >= startDay && row[2] <= endDay);
    const userType = staffInfo[3];
    let resultText = `${staffInfo[1]} ${staffInfo[2]} ${year}年${month + 1}月薪資明細\n`;
    if (userType === "一般" || userType === "獎金") {
      const baseSalary = parseFloat(staffInfo[4]) || 0;
      const otRate = parseFloat(staffInfo[4]) || 1.33;
      let sumLeave = myRows.filter(r => r[5] === "V" && (!r[6] || !r[6].includes("特休"))).length;
      let sumOT = 0;
      myRows.forEach(row => {
        if (row[3] && row[4] && (!row[5] || row[5] !== "V")) {
          const t1 = new Date(`${row[2]}T${row[3].split(" ")[1] || row[3]}:00+08:00`);
          const t2 = new Date(`${row[2]}T${row[4].split(" ")[1] || row[4]}:00+08:00`);
          let wh = (t2 - t1) / 3600000;
          if (t1 < new Date(`${row[2]}T13:00:00+08:00`) && t2 > new Date(`${row[2]}T12:00:00+08:00`)) wh -= 1;
          const baseOff = new Date(t1.getTime() + 9 * 3600000);
          if (t2 > baseOff) {
            const over = (t2 - baseOff) / 60000;
            if (over >= 30) sumOT += Math.floor((over / 60) * 2) / 2;
          }
        }
      });
      const laborInsurance = staffInfo[staffHeader.indexOf("勞健保")];
      let startDate = staffInfo[8];
      let specialLeave = startDate ? calcTaiwanLeave(startDate) : 0;
      let thisYear = now.getFullYear().toString();
      let specialLeaveUsed = rows.filter(
        row => row[0] === userId && row[2].startsWith(thisYear) && row[5] === "V" && (row[6] || "").includes("特休")
      ).length;
      let specialLeaveRemain = specialLeave - specialLeaveUsed;
      const dailySalary = baseSalary / daysInMonth;
      const leaveDeduct = dailySalary * sumLeave;
      let otPay = 0;
      if (userType === "一般") otPay = (dailySalary / 9) * sumOT * otRate;
      let bonus = 0, bonusDesc = "";
      if (userType === "獎金") {
        const bonusSheet = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: BONUS_SHEET + "!A1:G" });
        const bonusRows = bonusSheet.data.values || [];
        const myBonus = bonusRows.find(r => r[0] === userId && r[3] === `${year}-${pad(month + 1)}`);
        bonus = myBonus ? parseInt(myBonus[5], 10) : 0;
        bonusDesc = myBonus ? myBonus[6] || "" : "";
      }
      resultText += `基本底薪：${baseSalary}\n`;
      if (userType === "一般") resultText += `加班時數：${sumOT.toFixed(1)} 加班費：${otPay.toFixed(0)} 元\n`;
      resultText += `勞健保：${laborInsurance ? laborInsurance : ""}（公司全額支付）\n`;
      resultText += `今年特休：${specialLeave}天，已用：${specialLeaveUsed}天，剩餘：${specialLeaveRemain}天\n`;
      resultText += `病/事假：${sumLeave}天，請假扣薪：${leaveDeduct.toFixed(0)}元\n`;
      if (userType === "獎金") {
        resultText += `本月獎金：${bonus} 元 ${bonusDesc ? "\n說明：" + bonusDesc : ""}\n`;
        resultText += `-------------------------\n`;
        resultText += `薪資總額：${baseSalary}+${bonus}-${leaveDeduct.toFixed(0)} = ${(baseSalary + bonus - leaveDeduct).toFixed(0)} 元\n`;
      } else {
        resultText += `-------------------------\n`;
        resultText += `薪資總額：${baseSalary.toFixed(0)}+${otPay.toFixed(0)}-${leaveDeduct.toFixed(0)} = ${(baseSalary + otPay - leaveDeduct).toFixed(0)} 元\n`;
      }
      return client.replyMessage(event.replyToken, { type: "text", text: resultText });
    }
    if (userType === "工讀生") {
      const wage = parseInt(staffInfo[8], 10) || 0;
      let totalMinutes = 0;
      myRows.forEach(row => {
        if (row[3] && row[4] && (!row[5] || row[5] !== "V")) {
          const t1 = new Date(`${row[2]}T${row[3].split(" ")[1] || row[3]}:00+08:00`);
          const t2 = new Date(`${row[2]}T${row[4].split(" ")[1] || row[4]}:00+08:00`);
          let min = (t2 - t1) / 60000;
          if (min >= 301) min -= 60;
          if (min > 0) totalMinutes += min;
        }
      });
      const totalHours = totalMinutes / 60;
      let sumLeave = myRows.filter(r => r[5] === "V").length;
      resultText += `時薪：${wage} 元\n總時數：${totalHours.toFixed(2)}\n請假：${sumLeave} 天\n`;
      resultText += `本月工讀薪資：${wage} x ${totalHours.toFixed(2)} = ${(wage * totalHours).toFixed(0)} 元\n`;
      return client.replyMessage(event.replyToken, { type: "text", text: resultText });
    }
    return client.replyMessage(event.replyToken, { type: "text", text: resultText + "(查無薪資型態設定)" });
  }

  // ====== 匯出報表 ======
  if (/^(導出出勤|導出請假|匯出)/.test(msg) && adminUserIds.includes(userId)) {
    const attendSheet = await sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: ATTEND_SHEET + "!A1:N" });
    const attendHeader = attendSheet.data.values[0], attendRows = attendSheet.data.values.slice(1);
    let txt = attendRows.map(r =>
      `${r[attendHeader.indexOf("姓名")]} ${r[attendHeader.indexOf("日期")]}：${r[attendHeader.indexOf("上班時間")] || "-"} ~ ${r[attendHeader.indexOf("下班時間")] || "-"} ${(r[attendHeader.indexOf("請假")] === "V") ? `(請假${r[attendHeader.indexOf("假別說明")]}：${r[attendHeader.indexOf("請假狀態")] || ""})` : ""}`
    ).join("\n");
    return client.replyMessage(event.replyToken, { type: "text", text: `出勤報表預覽：\n${txt.slice(0, 3800)}${txt.length > 3800 ? '\n（內容過長已截斷）' : ''}` });
  }

  // 其它不回應
  return;
}

// 啟動 server
app.listen(process.env.PORT || 3000, '0.0.0.0', () => {
  console.log("LINE AI 打卡BOT已啟動");
});
