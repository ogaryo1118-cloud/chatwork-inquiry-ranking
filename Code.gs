/**
 * ============================================================
 *  Chatwork 困りごとダッシュボード
 *  Google Apps Script — Code.gs
 * ============================================================
 *
 *  【初回セットアップ手順】
 *  1. 新しい Google スプレッドシートを作成する
 *  2. 拡張機能 → Apps Script を開く
 *  3. このファイルの内容を Code.gs に貼り付ける
 *  4. ＋ボタンで HTML ファイルを追加し、名前を「dashboard」にする
 *     → dashboard.html の内容を貼り付ける
 *  5. 「setup」関数を実行してシートとトリガーを初期化する
 *  6. シート「設定」に Chatwork APIキー・Claude APIキー・
 *     管理者メールアドレスを入力する
 *  7. デプロイ → 新しいデプロイ → ウェブアプリとして公開し
 *     URLを社内に共有する
 *
 *  【注意】
 *  - 初日はデータがないためランキングは空です
 *  - 毎日自動取得が蓄積されることで1週間後に週次、
 *    1ヶ月後に月次ランキングが揃います
 * ============================================================
 */

// ============================================================
// 定数・グローバル設定
// ============================================================
const CONFIG = {
  CLAUDE_MODEL: 'claude-sonnet-4-5',          // 使用するClaudeモデル
  CHATWORK_API_BASE: 'https://api.chatwork.com/v2',
  CLAUDE_API_URL: 'https://api.anthropic.com/v1/messages',
  TIMEZONE: 'Asia/Tokyo',
  API_RATE_LIMIT_MS: 1000,      // Chatwork APIレート制限対策（1秒）
  KEYWORD_BATCH_SIZE: 50,       // Claude APIへの1バッチのメッセージ数
  MAX_KEYWORDS_PER_CALL: 20,    // 1回のClaude API呼び出しで抽出する最大キーワード数
  RISING_THRESHOLD: 0.5,        // 急上昇判定閾値（50%増）
};

const SHEET = {
  DASHBOARD: 'ダッシュボード',
  SETTINGS: '設定',
  DATA: 'データ',
};

const KEY = {
  CHATWORK_API: 'Chatwork APIキー',
  CLAUDE_API: 'Claude APIキー',
  ADMIN_EMAIL: '管理者メールアドレス',
  DAILY_REPORT: '日次レポート',
  WEEKLY_REPORT: '週次レポート',
  MONTHLY_REPORT: '月次レポート',
  AUTO_FETCH: '自動取得',
};


// ============================================================
// Web App エントリポイント
// ============================================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('dashboard')
    .evaluate()
    .setTitle('困りごとダッシュボード')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}


// ============================================================
// 初回セットアップ
// ============================================================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setupDashboardSheet(ss);
  _setupSettingsSheet(ss);
  _setupDataSheet(ss);
  _removeDefaultSheet(ss);
  setupTriggers();

  ss.toast(
    'セットアップ完了！シート「設定」にAPIキーとメールアドレスを入力してください。',
    '✅ セットアップ完了', 10
  );
}

function _setupDashboardSheet(ss) {
  let sh = ss.getSheetByName(SHEET.DASHBOARD);
  if (!sh) sh = ss.insertSheet(SHEET.DASHBOARD, 0);

  sh.clear();
  sh.clearFormats();
  sh.setTabColor('#1a73e8');

  sh.getRange('A1')
    .setValue('🔍 困りごとダッシュボード')
    .setFontSize(22).setFontWeight('bold').setFontColor('#1a73e8');
  sh.getRange('A2')
    .setValue('上のメニュー「🔧 システム管理」→ ランキング表示を選択してください')
    .setFontSize(11).setFontColor('#888888');

  // 操作ガイドテーブル
  const guide = [
    ['期間', '操作方法', '説明'],
    ['📊 本日', 'メニュー → 📊 本日のランキング表示', '当日に集計されたキーワード'],
    ['📅 今週', 'メニュー → 📅 今週のランキング表示', '月曜日からの累計キーワード'],
    ['📆 今月', 'メニュー → 📆 今月のランキング表示', '月初からの累計キーワード'],
  ];
  sh.getRange(4, 1, 4, 3).setValues(guide);
  sh.getRange('A4:C4').setFontWeight('bold').setBackground('#e8f0fe').setFontColor('#1a73e8');
  sh.getRange(4, 1, 4, 3)
    .setBorder(true, true, true, true, true, true, '#cccccc',
               SpreadsheetApp.BorderStyle.SOLID);

  sh.setColumnWidth(1, 100); sh.setColumnWidth(2, 280); sh.setColumnWidth(3, 220);
  sh.setColumnWidth(4, 100); sh.setColumnWidth(5, 180);

  return sh;
}

function _setupSettingsSheet(ss) {
  let sh = ss.getSheetByName(SHEET.SETTINGS);
  if (!sh) sh = ss.insertSheet(SHEET.SETTINGS, 1);

  sh.clear(); sh.clearFormats();
  sh.setTabColor('#ea4335');

  const rows = [
    ['⚙️ 設定画面', '',  '※ 管理者のみ編集してください'],
    ['', '', ''],
    ['── API設定 ──', '', ''],
    [KEY.CHATWORK_API,  '', 'Chatwork の API Token を入力（Settings > API Token）'],
    [KEY.CLAUDE_API,    '', 'Anthropic の APIキー を入力（sk-ant-...）'],
    [KEY.ADMIN_EMAIL,   '', 'レポートの送信先メールアドレス'],
    ['', '', ''],
    ['── レポート設定 ──', '', ''],
    [KEY.DAILY_REPORT,  'ON', '毎日 8:00 に自動送信（ON / OFF）'],
    [KEY.WEEKLY_REPORT, 'ON', '毎週月曜 8:00 に自動送信（ON / OFF）'],
    [KEY.MONTHLY_REPORT,'ON', '毎月 1日 8:00 に自動送信（ON / OFF）'],
    ['', '', ''],
    ['── 自動取得設定 ──', '', ''],
    [KEY.AUTO_FETCH,    'ON', '毎日 3:00 にメッセージを自動取得（ON / OFF）'],
  ];

  sh.getRange(1, 1, rows.length, 3).setValues(rows);

  // スタイル
  sh.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#ea4335');
  sh.getRange('C1').setFontColor('#888888').setFontStyle('italic');

  [[3,'A3:C3'], [8,'A8:C8'], [13,'A13:C13']].forEach(([row, range]) => {
    sh.getRange(range).setBackground('#f1f3f4').setFontWeight('bold').setFontColor('#555555');
  });

  // 入力欄（黄色背景）
  [4, 5, 6, 9, 10, 11, 14].forEach(r => {
    sh.getRange(r, 1).setFontWeight('bold');
    sh.getRange(r, 2).setBackground('#fffde7')
      .setBorder(true, true, true, true, false, false, '#f0c000',
                 SpreadsheetApp.BorderStyle.SOLID);
  });

  // ON/OFF ドロップダウン
  const onOffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['ON', 'OFF'], true)
    .setAllowInvalid(false).build();
  sh.getRange('B9:B11').setDataValidation(onOffRule);
  sh.getRange('B14').setDataValidation(onOffRule);

  sh.setColumnWidth(1, 220); sh.setColumnWidth(2, 260); sh.setColumnWidth(3, 340);

  // シート保護（警告のみ）
  sh.protect().setDescription('設定シート — 管理者のみ').setWarningOnly(true);

  return sh;
}

function _setupDataSheet(ss) {
  let sh = ss.getSheetByName(SHEET.DATA);
  if (!sh) sh = ss.insertSheet(SHEET.DATA, 2);

  sh.clear(); sh.clearFormats();
  sh.setTabColor('#34a853');
  sh.hideSheet(); // 通常ユーザーには非表示

  sh.getRange(1, 1, 1, 3).setValues([['日付', 'キーワード', '件数']])
    .setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');

  sh.setColumnWidth(1, 130); sh.setColumnWidth(2, 220); sh.setColumnWidth(3, 80);
  sh.setFrozenRows(1);
  sh.getRange('A1').setNote('⚠️ このシートは自動管理されます。手動で編集しないでください。');

  return sh;
}

function _removeDefaultSheet(ss) {
  ['Sheet1', 'シート1'].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh && ss.getSheets().length > 3) {
      try { ss.deleteSheet(sh); } catch (_) {}
    }
  });
}


// ============================================================
// カスタムメニュー（スプレッドシートを開いたときに自動表示）
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔧 システム管理')
    .addItem('⚙️ セットアップ実行（初回のみ）', 'setup')
    .addItem('📌 ダッシュボードにボタン追加', 'setupDashboardButtons')
    .addSeparator()
    .addItem('▶ 今すぐデータ取得', 'manualFetch')
    .addItem('⏹ 自動取得を停止',   'stopAutoFetch')
    .addItem('▶ 自動取得を再開',   'resumeAutoFetch')
    .addSeparator()
    .addItem('📊 本日のランキング表示', 'showTodayRanking')
    .addItem('📅 今週のランキング表示', 'showWeekRanking')
    .addItem('📆 今月のランキング表示', 'showMonthRanking')
    .addSeparator()
    .addItem('📧 日次レポート テスト送信',  'sendDailyReport')
    .addItem('📧 週次レポート テスト送信',  'sendWeeklyReport')
    .addItem('📧 月次レポート テスト送信',  'sendMonthlyReport')
    .addSeparator()
    .addItem('🔒 設定を開く（パスワード認証）', 'openSettingsWithPassword')
    .addItem('🔓 設定を閉じる',               'closeSettings')
    .addItem('🔑 設定パスワードを変更',         'changeSettingsPassword')
    .addSeparator()
    .addItem('🗑 本日データを削除（テスト用）', 'deleteTodayData')
    .addToUi();
}


// ============================================================
// トリガー管理
// ============================================================
function setupTriggers() {
  // 既存トリガーをすべて削除してから再作成
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // 毎日 3:00 — メッセージ取得
  ScriptApp.newTrigger('dailyTask').timeBased().everyDays(1).atHour(3).create();

  // 毎日 8:00 — 日次レポート
  ScriptApp.newTrigger('sendDailyReport').timeBased().everyDays(1).atHour(8).create();

  // 毎週月曜 8:00 — 週次レポート
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();

  // 毎月 1日 8:00 — 月次レポート
  ScriptApp.newTrigger('sendMonthlyReport').timeBased().onMonthDay(1).atHour(8).create();

  console.log('トリガーを設定しました。');
}

function stopAutoFetch() {
  _setSettingValue(KEY.AUTO_FETCH, 'OFF');
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('自動取得を停止しました。', '⏹ 停止', 5);
}

function resumeAutoFetch() {
  _setSettingValue(KEY.AUTO_FETCH, 'ON');
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('自動取得を再開しました。', '▶ 再開', 5);
}


// ============================================================
// 設定シートの読み書き
// ============================================================
function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.SETTINGS);
  if (!sh) return {};

  const values = sh.getDataRange().getValues();
  const settings = {};

  values.forEach(([k, v]) => {
    const key = String(k).trim();
    // セクション行・空行を除外
    if (key && !key.startsWith('─') && !key.startsWith('⚙') && !key.startsWith('※')) {
      settings[key] = v;
    }
  });
  return settings;
}

function _setSettingValue(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.SETTINGS);
  if (!sh) return;

  const values = sh.getDataRange().getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === key) {
      sh.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
}


// ============================================================
// メイン日次タスク（毎日3時に自動実行）
// ============================================================
function dailyTask() {
  const settings = getSettings();

  if (settings[KEY.AUTO_FETCH] !== 'ON') {
    console.log('[dailyTask] 自動取得はOFFです。スキップします。');
    return;
  }

  const chatworkKey = settings[KEY.CHATWORK_API];
  const claudeKey   = settings[KEY.CLAUDE_API];

  if (!chatworkKey || !claudeKey) {
    console.error('[dailyTask] APIキーが未設定です。設定シートを確認してください。');
    return;
  }

  try {
    fetchAndProcessMessages(chatworkKey, claudeKey);
  } catch (e) {
    console.error('[dailyTask] エラー: ' + e.toString());
  }
}

// 手動実行（メニューから呼び出し）
function manualFetch() {
  const ui = SpreadsheetApp.getUi();
  const settings = getSettings();
  const chatworkKey = settings[KEY.CHATWORK_API];
  const claudeKey   = settings[KEY.CLAUDE_API];

  if (!chatworkKey || !claudeKey) {
    ui.alert(
      '⚠️ APIキー未設定',
      'シート「設定」にChatwork APIキーとClaude APIキーを入力してください。',
      ui.ButtonSet.OK
    );
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet()
    .toast('データ取得中…（数分かかる場合があります）', '⏳ 処理中', 120);

  try {
    fetchAndProcessMessages(chatworkKey, claudeKey);
    SpreadsheetApp.getActiveSpreadsheet()
      .toast('データ取得が完了しました！', '✅ 完了', 5);
  } catch (e) {
    ui.alert('エラー', 'データ取得中にエラーが発生しました：\n' + e.toString(), ui.ButtonSet.OK);
  }
}


// ============================================================
// データ取得＆処理メイン
// ============================================================
function fetchAndProcessMessages(chatworkKey, claudeKey) {
  const today   = new Date();
  const todayStr = Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd');

  console.log('==== データ取得開始: ' + todayStr + ' ====');

  // 1. 全ルーム取得
  const rooms = _getChatworkRooms(chatworkKey);
  if (!rooms || rooms.length === 0) {
    console.log('[fetch] ルームが見つかりません。');
    return;
  }
  console.log('[fetch] ルーム数: ' + rooms.length);

  let allMessages = [];

  // 2. 各ルームのメッセージを取得（本日分のみ）
  rooms.forEach((room, idx) => {
    try {
      const msgs = _getChatworkMessages(chatworkKey, room.room_id);

      const todayMsgs = msgs.filter(m => {
        const d = Utilities.formatDate(
          new Date(m.send_time * 1000), CONFIG.TIMEZONE, 'yyyy-MM-dd'
        );
        return d === todayStr;
      });

      const texts = todayMsgs
        .map(m => _sanitizePersonalInfo(m.body || '')) // Claude送信前にPIIをマスク
        .filter(t => t.trim().length >= 5); // 短すぎるメッセージを除外

      allMessages = allMessages.concat(texts);
    } catch (e) {
      console.error('[fetch] ルーム ' + room.room_id + ' エラー: ' + e);
    }

    // レート制限：1秒待機
    if (idx < rooms.length - 1) Utilities.sleep(CONFIG.API_RATE_LIMIT_MS);
  });

  console.log('[fetch] 本日のメッセージ数: ' + allMessages.length);

  if (allMessages.length === 0) {
    console.log('[fetch] 本日のメッセージがありません。');
    return;
  }

  // 3. Claude API でキーワード抽出
  const keywordCounts = _extractKeywordsWithClaude(claudeKey, allMessages);
  const kwCount = Object.keys(keywordCounts).length;

  if (kwCount === 0) {
    console.log('[fetch] 抽出されたキーワードがありません。');
    return;
  }

  // 4. シートに保存
  _saveKeywords(keywordCounts, todayStr);
  console.log('[fetch] キーワード保存完了: ' + kwCount + '種類');
  console.log('==== データ取得完了 ====');
}


// ============================================================
// Chatwork API
// ============================================================
function _getChatworkRooms(apiKey) {
  const res = UrlFetchApp.fetch(CONFIG.CHATWORK_API_BASE + '/rooms', {
    method: 'get',
    headers: { 'X-ChatWorkToken': apiKey },
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    console.error('[Chatwork] rooms API エラー: ' + res.getResponseCode()
                  + ' — ' + res.getContentText().substring(0, 200));
    return [];
  }

  try { return JSON.parse(res.getContentText()); }
  catch (e) { console.error('[Chatwork] rooms JSON解析エラー: ' + e); return []; }
}

function _getChatworkMessages(apiKey, roomId) {
  // force=1: 既読・未読に関わらず最新100件を取得
  const url = CONFIG.CHATWORK_API_BASE + '/rooms/' + roomId + '/messages?force=1';

  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'X-ChatWorkToken': apiKey },
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  if (code === 204) return []; // コンテンツなし
  if (code !== 200) {
    console.error('[Chatwork] messages エラー (room:' + roomId + '): ' + code);
    return [];
  }

  try {
    const data = JSON.parse(res.getContentText());
    return Array.isArray(data) ? data : [];
  } catch (e) {
    console.error('[Chatwork] messages JSON解析エラー: ' + e);
    return [];
  }
}


// ============================================================
// Claude API — キーワード抽出
// ============================================================
function _extractKeywordsWithClaude(apiKey, messages) {
  const keywordCounts = {};

  // バッチ処理でトークン制限を回避
  for (let i = 0; i < messages.length; i += CONFIG.KEYWORD_BATCH_SIZE) {
    const batch = messages.slice(i, i + CONFIG.KEYWORD_BATCH_SIZE);
    const batchText = batch.join('\n---\n');
    const batchNum = Math.floor(i / CONFIG.KEYWORD_BATCH_SIZE) + 1;

    try {
      const keywords = _callClaudeAPI(apiKey, batchText);
      keywords.forEach(kw => {
        const k = kw.trim();
        if (k && k.length >= 2) {
          keywordCounts[k] = (keywordCounts[k] || 0) + 1;
        }
      });
      console.log('[Claude] バッチ' + batchNum + '完了: ' + keywords.length + 'キーワード');
    } catch (e) {
      console.error('[Claude] バッチ' + batchNum + 'エラー: ' + e);
    }

    // Claude API レート制限対策
    if (i + CONFIG.KEYWORD_BATCH_SIZE < messages.length) Utilities.sleep(500);
  }

  return keywordCounts;
}

// ============================================================
// 個人情報マスキング（Claude API 送信前の前処理）
// ============================================================
// 【設計方針】
//   Chatwork から取得したメッセージは、Claude API へ渡す前に
//   このスプレッドシート上（Google サーバー内）で個人情報を
//   自動マスクします。外部 AI クラウドには、マスク済みテキスト
//   のみが送信されるため、個人情報がクラウド外に出ることはありません。
//
//   マスク対象：
//     - メールアドレス  → [メールアドレス]
//     - URL           → [URL]
//     - 電話番号       → [電話番号]
//     - 郵便番号       → [郵便番号]
//     - 敬称付き名前   → [お名前]（例：田中さん、山田様）
// ============================================================
function _sanitizePersonalInfo(text) {
  let s = String(text);
  // メールアドレス
  s = s.replace(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g, '[メールアドレス]');
  // URL (http / https)
  s = s.replace(/https?:\/\/[^\s\u3000]+/g, '[URL]');
  // 電話番号（固定・携帯・フリーダイヤル）
  s = s.replace(/0\d{1,4}[-\u30FB\s]?\d{1,4}[-\u30FB\s]?\d{4}/g, '[電話番号]');
  // 郵便番号
  s = s.replace(/\u3012?\d{3}[-\s]?\d{4}/g, '[郵便番号]');
  // 敬称付き人名（漢字1〜4文字 + さん／様／くん／ちゃん／氏）
  s = s.replace(/[\u4E00-\u9FFF]{1,4}(さん|様|くん|ちゃん|氏)/g, '[お名前]');
  return s;
}

function _callClaudeAPI(apiKey, messagesText) {
  const prompt =
`以下のチャットメッセージを分析して、「困りごと・質問・不満・問題点・要望」に関するキーワードを抽出してください。

【絶対に守るルール — セキュリティ要件】
- 個人名・会社名・電話番号・メールアドレス・URL・住所を一切含めない
- 固有名詞（人名・製品名・地名）は除外する
- 困りごとの「内容・種類・カテゴリ」を表す一般的なキーワードのみ抽出する
  例：「田中さんからログインできないと言われた」→「ログイン障害」
  例：「XXX社のシステムが遅い」→「動作遅延」
  例：「○○の使い方がわからない」→「操作方法不明」

【出力形式（JSON のみ、他の文章は不要）】
{"keywords": ["キーワード1", "キーワード2", ...]}
最大${CONFIG.MAX_KEYWORDS_PER_CALL}個。困りごとのないメッセージからは抽出しないこと。

【分析対象メッセージ】
${messagesText}`;

  const payload = {
    model: CONFIG.CLAUDE_MODEL,
    max_tokens: 1024,
    messages: [{ role: 'user', content: prompt }],
  };

  const res = UrlFetchApp.fetch(CONFIG.CLAUDE_API_URL, {
    method: 'post',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json',
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (res.getResponseCode() !== 200) {
    console.error('[Claude] API エラー: ' + res.getResponseCode()
                  + ' — ' + res.getContentText().substring(0, 300));
    return [];
  }

  try {
    const result  = JSON.parse(res.getContentText());
    const content = result.content[0].text;

    const jsonMatch = content.match(/\{[\s\S]*"keywords"[\s\S]*?\}/);
    if (!jsonMatch) {
      console.error('[Claude] JSON抽出失敗。レスポンス: ' + content.substring(0, 300));
      return [];
    }

    const parsed = JSON.parse(jsonMatch[0]);
    return Array.isArray(parsed.keywords) ? parsed.keywords : [];
  } catch (e) {
    console.error('[Claude] レスポンス解析エラー: ' + e);
    return [];
  }
}


// ============================================================
// データシート — 保存・取得
// ============================================================
function _saveKeywords(keywordCounts, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.DATA);
  if (!sh) return;

  const lastRow    = sh.getLastRow();
  const existing   = lastRow > 1
    ? sh.getRange(2, 1, lastRow - 1, 3).getValues()
    : [];

  // 既存データのインデックスマップ {date_keyword: sheetRowIndex(1-based)}
  const existMap = {};
  existing.forEach((row, i) => {
    existMap[row[0] + '__' + row[1]] = i + 2;
  });

  const newRows = [];

  Object.entries(keywordCounts).forEach(([keyword, count]) => {
    const mapKey = dateStr + '__' + keyword;
    if (existMap[mapKey]) {
      // 既存行のカウントを加算
      const existingCount = Number(existing[existMap[mapKey] - 2][2]) || 0;
      sh.getRange(existMap[mapKey], 3).setValue(existingCount + count);
    } else {
      newRows.push([dateStr, keyword, count]);
    }
  });

  if (newRows.length > 0) {
    sh.getRange(lastRow + 1, 1, newRows.length, 3).setValues(newRows);
  }
}

/**
 * 指定した期間のキーワード集計を返す
 * @param {string} startDateStr  'yyyy-MM-dd'
 * @param {string} endDateStr    'yyyy-MM-dd'
 * @returns {Object} { keyword: count, ... }
 */
function _getDataForPeriod(startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.DATA);
  if (!sh || sh.getLastRow() <= 1) return {};

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const counts = {};

  data.forEach(row => {
    const d = String(row[0]);
    const k = String(row[1]);
    const c = Number(row[2]) || 0;
    if (d >= startDateStr && d <= endDateStr && k && c > 0) {
      counts[k] = (counts[k] || 0) + c;
    }
  });

  return counts;
}

/**
 * 期間の開始・終了日付文字列を返す
 * @param {string} period  'today' | 'week' | 'month'
 */
function _getDateRange(period) {
  const now      = new Date();
  const todayStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  let   startStr;

  if (period === 'today') {
    startStr = todayStr;
  } else if (period === 'week') {
    const d   = now.getDay();
    const diff = d === 0 ? 6 : d - 1; // 月曜始まり
    const start = new Date(now);
    start.setDate(now.getDate() - diff);
    startStr = Utilities.formatDate(start, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  } else if (period === 'month') {
    startStr = Utilities.formatDate(
      new Date(now.getFullYear(), now.getMonth(), 1),
      CONFIG.TIMEZONE, 'yyyy-MM-dd'
    );
  }

  return { startStr, endStr: todayStr };
}

/**
 * 期間集計からTOP10ランキング配列を生成
 */
function _buildRanking(keywordCounts) {
  return Object.entries(keywordCounts)
    .map(([keyword, count]) => ({ keyword, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);
}

/**
 * 急上昇キーワードを計算（今週vs先週、+50%以上）
 */
function _getRisingKeywords() {
  const now      = new Date();
  const todayStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd');

  // 今週（過去7日）
  const thisStart = new Date(now);
  thisStart.setDate(now.getDate() - 7);
  const thisStartStr = Utilities.formatDate(thisStart, CONFIG.TIMEZONE, 'yyyy-MM-dd');

  // 先週（8〜14日前）
  const lastEnd   = new Date(thisStart);
  lastEnd.setDate(lastEnd.getDate() - 1);
  const lastStart = new Date(lastEnd);
  lastStart.setDate(lastEnd.getDate() - 6);
  const lastStartStr = Utilities.formatDate(lastStart, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const lastEndStr   = Utilities.formatDate(lastEnd,   CONFIG.TIMEZONE, 'yyyy-MM-dd');

  const thisWeek = _getDataForPeriod(thisStartStr, todayStr);
  const lastWeek = _getDataForPeriod(lastStartStr, lastEndStr);

  const rising = [];

  Object.entries(thisWeek).forEach(([kw, thisCount]) => {
    const lastCount = lastWeek[kw] || 0;
    let increase;

    if (lastCount === 0 && thisCount > 0) {
      increase = 100; // 新規出現
    } else if (lastCount > 0 && thisCount >= lastCount * (1 + CONFIG.RISING_THRESHOLD)) {
      increase = Math.round(((thisCount - lastCount) / lastCount) * 100);
    } else {
      return; // 対象外
    }

    rising.push({ keyword: kw, thisCount, lastCount, increase });
  });

  return rising.sort((a, b) => b.increase - a.increase).slice(0, 5);
}


// ============================================================
// Web App 用データ取得関数（dashboard.html から呼ばれる）
// ============================================================

/** ダッシュボード全データをJSON文字列で返す */
function getDashboardData(period) {
  try {
    const { startStr, endStr } = _getDateRange(period);
    const counts   = _getDataForPeriod(startStr, endStr);
    const ranking  = _buildRanking(counts);
    const rising   = _getRisingKeywords();
    const total    = ranking.reduce((s, item) => s + item.count, 0);

    return JSON.stringify({
      success: true,
      period,
      ranking,
      rising,
      summary: {
        totalCount:    total,
        categoryCount: Object.keys(counts).length,
        topKeyword:    ranking.length > 0 ? ranking[0].keyword : 'データなし',
      },
    });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.toString() });
  }
}


// ============================================================
// ダッシュボードシート更新（スプレッドシート内の表示）
// ============================================================
function showTodayRanking() { _updateDashboardSheet('today'); }
function showWeekRanking()  { _updateDashboardSheet('week');  }
function showMonthRanking() { _updateDashboardSheet('month'); }

function _updateDashboardSheet(period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.DASHBOARD);
  if (!sh) return;

  const { startStr, endStr } = _getDateRange(period);
  const counts  = _getDataForPeriod(startStr, endStr);
  const ranking = _buildRanking(counts);
  const rising  = _getRisingKeywords();
  const total   = ranking.reduce((s, i) => s + i.count, 0);
  const catCount = Object.keys(counts).length;

  const labels = { today: '本日', week: '今週', month: '今月' };
  const pLabel = labels[period];
  const now    = new Date();
  const nowStr = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy年MM月dd日 HH:mm');

  sh.clear(); sh.clearFormats();

  // ─── タイトル ───────────────────────────────────────────
  sh.getRange('A1').setValue('🔍 困りごとダッシュボード')
    .setFontSize(22).setFontWeight('bold').setFontColor('#1a73e8');
  sh.getRange('C1').setValue('更新: ' + nowStr + '　期間: ' + pLabel)
    .setFontColor('#888888').setFontSize(10).setHorizontalAlignment('right');

  // ─── サマリー ───────────────────────────────────────────
  sh.getRange('A3').setValue('📊 サマリー')
    .setFontSize(13).setFontWeight('bold').setFontColor('#1a73e8');

  const summaryRows = [
    ['総件数',      total + ' 件'],
    ['カテゴリ数',  catCount + ' 種類'],
    ['1位キーワード', ranking.length > 0 ? ranking[0].keyword : 'データなし'],
  ];
  summaryRows.forEach(([label, val], i) => {
    const row = 4 + i;
    sh.getRange(row, 1).setValue(label).setFontWeight('bold').setFontColor('#555555');
    sh.getRange(row, 2).setValue(val).setFontSize(13).setFontWeight('bold');
    sh.getRange(row, 1, 1, 3).setBackground(i % 2 === 0 ? '#e8f0fe' : '#f8f9fa');
  });

  // ─── ランキング ─────────────────────────────────────────
  sh.getRange('A8').setValue('🏆 困りごとランキング TOP10 (' + pLabel + ')')
    .setFontSize(13).setFontWeight('bold').setFontColor('#1a73e8');

  sh.getRange(9, 1, 1, 4).setValues([['順位', 'キーワード', '件数', 'バー']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff').setFontSize(11);

  if (ranking.length === 0) {
    sh.getRange('A10').setValue('データがありません（自動取得後に表示されます）')
      .setFontColor('#999999');
  } else {
    const max = ranking[0].count;
    const medals = ['🥇', '🥈', '🥉'];

    ranking.forEach((item, i) => {
      const row = 10 + i;
      const rank = medals[i] || String(i + 1);
      const bar  = '█'.repeat(Math.max(1, Math.round((item.count / max) * 18)));
      const bg   = i === 0 ? '#fff8e1' : (i % 2 === 0 ? '#ffffff' : '#f8f9fa');

      sh.getRange(row, 1).setValue(rank).setHorizontalAlignment('center');
      sh.getRange(row, 2).setValue(item.keyword);
      sh.getRange(row, 3).setValue(item.count).setHorizontalAlignment('center');
      sh.getRange(row, 4).setValue(bar).setFontColor('#1a73e8');

      sh.getRange(row, 1, 1, 4).setBackground(bg);
      if (i === 0) sh.getRange(row, 1, 1, 4).setFontWeight('bold');
    });
  }

  // ─── 急上昇キーワード ────────────────────────────────────
  const rRow = 22;
  sh.getRange('A' + rRow).setValue('⬆ 急上昇キーワード（先週比 +50% 以上）')
    .setFontSize(13).setFontWeight('bold').setFontColor('#e65100');

  if (rising.length === 0) {
    sh.getRange('A' + (rRow + 1)).setValue('急上昇キーワードはありません')
      .setFontColor('#999999');
  } else {
    rising.forEach((item, i) => {
      const row   = rRow + 1 + i;
      const isNew = item.lastCount === 0;
      const icon  = isNew ? '🆕' : '⬆';
      const pct   = '+' + item.increase + '%';
      const detail = '今週: ' + item.thisCount + '件 / 先週: ' + item.lastCount + '件';

      sh.getRange(row, 1).setValue(icon).setHorizontalAlignment('center');
      sh.getRange(row, 2).setValue(item.keyword).setFontWeight('bold');
      sh.getRange(row, 3).setValue(pct).setFontColor('#d32f2f').setFontWeight('bold');
      sh.getRange(row, 4).setValue(detail).setFontColor('#666666');
      sh.getRange(row, 1, 1, 4).setBackground('#fff3e0');
    });
  }

  // ─── 列幅 ───────────────────────────────────────────────
  sh.setColumnWidth(1, 65); sh.setColumnWidth(2, 210);
  sh.setColumnWidth(3, 80); sh.setColumnWidth(4, 200);

  ss.setActiveSheet(sh);
  ss.toast(pLabel + 'のランキングを更新しました。', '✅ 更新完了', 3);
}


// ============================================================
// メールレポート送信
// ============================================================
function sendDailyReport() {
  const settings = getSettings();
  if (settings[KEY.DAILY_REPORT] !== 'ON') {
    console.log('[report] 日次レポートはOFF。スキップ。');
    return;
  }
  _sendReport('today', '日次');
}

function sendWeeklyReport() {
  const settings = getSettings();
  if (settings[KEY.WEEKLY_REPORT] !== 'ON') {
    console.log('[report] 週次レポートはOFF。スキップ。');
    return;
  }
  _sendReport('week', '週次');
}

function sendMonthlyReport() {
  const settings = getSettings();
  if (settings[KEY.MONTHLY_REPORT] !== 'ON') {
    console.log('[report] 月次レポートはOFF。スキップ。');
    return;
  }
  _sendReport('month', '月次');
}

function _sendReport(period, reportType) {
  const settings = getSettings();
  const email    = settings[KEY.ADMIN_EMAIL];

  if (!email) {
    console.error('[report] 管理者メールアドレスが未設定です。');
    return;
  }

  const { startStr, endStr } = _getDateRange(period);
  const counts   = _getDataForPeriod(startStr, endStr);
  const ranking  = _buildRanking(counts);
  const rising   = _getRisingKeywords();
  const total    = ranking.reduce((s, i) => s + i.count, 0);
  const catCount = Object.keys(counts).length;
  const pLabel   = { today: '本日', week: '今週', month: '今月' }[period];
  const dateStr  = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy年MM月dd日');

  const subject = `【${reportType}レポート】困りごとランキング — ${dateStr}`;

  let body = '';
  body += '━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  body += `  困りごとキーワード ${reportType}レポート\n`;
  body += `  集計期間: ${pLabel}（${startStr} ～ ${endStr}）\n`;
  body += '━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';

  body += '■ サマリー\n';
  body += `  総件数　　: ${total} 件\n`;
  body += `  カテゴリ数: ${catCount} 種類\n`;
  body += `  1位　　　: ${ranking.length > 0 ? ranking[0].keyword : 'データなし'}\n\n`;

  body += '■ 困りごとランキング TOP10\n';
  if (ranking.length === 0) {
    body += '  データがありません。\n';
  } else {
    const medals = ['🥇', '🥈', '🥉'];
    ranking.forEach((item, i) => {
      const prefix = medals[i] ? medals[i] + ' ' : `  ${i + 1}位 `;
      body += `${prefix}${item.keyword}（${item.count}件）\n`;
    });
  }

  body += '\n■ 急上昇キーワード（先週比 +50% 以上）\n';
  if (rising.length === 0) {
    body += '  該当なし\n';
  } else {
    rising.forEach(item => {
      const tag = item.lastCount === 0 ? '🆕新規' : `⬆ +${item.increase}%`;
      body += `  ${tag} ${item.keyword}（今週: ${item.thisCount}件 / 先週: ${item.lastCount}件）\n`;
    });
  }

  body += '\n━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
  body += 'このメールは自動送信されています。\n';
  body += '配信停止はスプレッドシートの「設定」シートで変更できます。\n';

  GmailApp.sendEmail(email, subject, body);
  console.log('[report] ' + reportType + 'レポート送信完了 → ' + email);
}


// ============================================================
// ユーティリティ（テスト・メンテナンス用）
// ============================================================

/** 本日のデータを削除する（テスト用） */
function deleteTodayData() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.alert(
    '⚠ 確認',
    '本日のデータをすべて削除しますか？\nこの操作は元に戻せません。',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.DATA);
  if (!sh || sh.getLastRow() <= 1) {
    ui.alert('データがありません。'); return;
  }

  const todayStr = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  const data     = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const keep     = data.filter(row => String(row[0]) !== todayStr);

  sh.clearContents();
  sh.getRange(1, 1, 1, 3).setValues([['日付', 'キーワード', '件数']]);
  if (keep.length > 0) sh.getRange(2, 1, keep.length, 3).setValues(keep);

  ss.toast('本日のデータを削除しました。', '🗑 削除完了', 3);
}

// ============================================================
// ダッシュボード ランキングボタン（チェックボックス方式）
// ============================================================

/**
 * ダッシュボードシートに3つのランキングボタン（チェックボックス）を追加する。
 * 初回セットアップ後に1度だけ実行。
 */
function setupDashboardButtons() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.DASHBOARD);

  // 列幅
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 200);
  sh.setColumnWidth(4, 120);

  // セクションヘッダー (行9)
  sh.getRange(9, 1, 1, 4).clear().setBackground('#f1f3f4');
  sh.getRange('A9')
    .setValue('▼ ランキング表示（下のチェックボックスをクリック）')
    .setFontSize(10).setFontColor('#5f6368').setFontStyle('italic');
  sh.setRowHeight(9, 26);

  // ボタンラベル行 (行10)
  const BTNS = [
    ['A10', '📊 本日のランキング',  '#1a73e8', '#0d47a1'],
    ['B10', '📅 今週のランキング',  '#188038', '#0d652d'],
    ['C10', '📆 今月のランキング',  '#ea4335', '#b71c1c'],
  ];
  BTNS.forEach(([addr, label, bg, bdr]) => {
    const c = sh.getRange(addr);
    c.setValue(label)
      .setBackground(bg).setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(12)
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    c.setBorder(true, true, true, true, null, null,
                bdr, SpreadsheetApp.BorderStyle.SOLID_THICK);
  });
  sh.setRowHeight(10, 45);

  // チェックボックス行 (行11) — クリックでonEdit発火
  sh.getRange('A11:C11').clearContent().clearFormat();
  sh.getRange('A11:C11').insertCheckboxes();
  sh.getRange('A11').setBackground('#e8f0fe')
    .setNote('チェックで「本日のランキング」を表示します');
  sh.getRange('B11').setBackground('#e6f4ea')
    .setNote('チェックで「今週のランキング」を表示します');
  sh.getRange('C11').setBackground('#fce8e6')
    .setNote('チェックで「今月のランキング」を表示します');
  sh.getRange('A11:C11').setHorizontalAlignment('center');
  sh.setRowHeight(11, 32);

  // ヒントテキスト (行12)
  sh.getRange(12, 1, 1, 4).clear().setBackground('#ffffff');
  sh.getRange('A12')
    .setValue('↑ チェックボックスをクリックするとランキングがこの下に表示されます')
    .setFontColor('#9e9e9e').setFontSize(9).setFontStyle('italic');
  sh.setRowHeight(12, 20);

  ss.toast('ダッシュボードにランキングボタンを追加しました！', '✅ 完了', 5);
}

// ============================================================
// onEdit — チェックボックスボタン処理
// ============================================================
function onEdit(e) {
  const sheet = e.range.getSheet();
  const row   = e.range.getRow();
  const col   = e.range.getColumn();

  // ダッシュボード シート・行11・チェックONの場合のみ処理
  if (sheet.getName() !== SHEET.DASHBOARD) return;
  if (row !== 11) return;
  if (e.value !== 'TRUE') return;

  e.range.setValue(false); // 即リセット
  if      (col === 1) _showSampleRanking('today');
  else if (col === 2) _showSampleRanking('week');
  else if (col === 3) _showSampleRanking('month');
}

// ============================================================
// サンプルランキングをシート内に表示
// ============================================================
function _showSampleRanking(period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.DASHBOARD);

  const META = {
    today: { title: '📊 本日のランキング（サンプルデータ）', bg: '#e8f0fe', fg: '#1a73e8' },
    week:  { title: '📅 今週のランキング（サンプルデータ）', bg: '#e6f4ea', fg: '#188038' },
    month: { title: '📆 今月のランキング（サンプルデータ）', bg: '#fce8e6', fg: '#ea4335' },
  };

  const SAMPLE = {
    today: [
      ['エラーが出る', 15], ['接続できない', 12], ['ログインできない', 9],
      ['表示されない', 8],  ['動作が遅い', 7],    ['データが消えた', 5],
      ['印刷できない', 4],  ['音が出ない', 3],    ['画面が固まる', 2],
      ['メールが届かない', 1],
    ],
    week: [
      ['接続できない', 58],  ['エラーが出る', 47],      ['動作が遅い', 39],
      ['ログインできない', 32], ['表示されない', 28],   ['データが消えた', 19],
      ['印刷できない', 15],  ['メールが届かない', 12],  ['画面が固まる', 8],
      ['音が出ない', 5],
    ],
    month: [
      ['動作が遅い', 210],   ['接続できない', 185],     ['エラーが出る', 172],
      ['表示されない', 143], ['ログインできない', 128],  ['データが消えた', 95],
      ['印刷できない', 67],  ['メールが届かない', 54],   ['画面が固まる', 43],
      ['音が出ない', 28],
    ],
  };

  const { title, bg, fg } = META[period];
  const data  = SAMPLE[period];
  const START = 14; // ランキング表示開始行

  // 既存ランキングをクリア
  const lastRow = sh.getLastRow();
  if (lastRow >= START) {
    sh.getRange(START, 1, lastRow - START + 1, 4).clear();
  }

  // タイトル行
  sh.getRange(START, 1, 1, 3).merge()
    .setValue(title)
    .setFontSize(14).setFontWeight('bold').setFontColor(fg)
    .setBackground(bg).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(START, 38);

  // ヘッダー行
  const hRow = START + 1;
  sh.getRange(hRow, 1, 1, 3)
    .setValues([['順位', 'キーワード', '件数']])
    .setFontWeight('bold').setBackground(fg).setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(hRow, 28);

  // データ行
  const medals  = ['🥇', '🥈', '🥉'];
  const topBgs  = ['#fff8e1', '#f5f5f5', '#fafafa'];
  data.forEach((item, i) => {
    const r    = hRow + 1 + i;
    const rank = i < 3 ? (medals[i] + ' ' + (i + 1) + '位') : ((i + 1) + '位');
    sh.getRange(r, 1).setValue(rank).setHorizontalAlignment('center');
    sh.getRange(r, 2).setValue(item[0]).setHorizontalAlignment('left');
    sh.getRange(r, 3).setValue(item[1]).setHorizontalAlignment('center');
    const rowBg = i < 3 ? topBgs[i] : (i % 2 === 0 ? '#f8f9fa' : '#ffffff');
    sh.getRange(r, 1, 1, 3).setBackground(rowBg).setVerticalAlignment('middle');
    sh.setRowHeight(r, 24);
  });

  // 外枠・横線
  sh.getRange(hRow, 1, data.length + 1, 3)
    .setBorder(true, true, true, true, false, true,
               '#e0e0e0', SpreadsheetApp.BorderStyle.SOLID);

  // フッター（更新日時）
  const now  = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy/MM/dd HH:mm');
  const fRow = hRow + data.length + 1;
  sh.getRange(fRow, 1, 1, 3).merge()
    .setValue('※ サンプルデータ  |  ' + now + ' 表示')
    .setFontColor('#9e9e9e').setFontSize(9)
    .setHorizontalAlignment('right').setFontStyle('italic');

  // ランキング先頭へスクロール
  ss.setActiveSheet(sh);
  sh.setActiveCell(sh.getRange(START, 1));
  ss.toast(title + ' を表示しました', '📊 ランキング', 3);
}


// ============================================================
// 設定シート パスワード保護
// ============================================================

const _PW_KEY = 'SETTINGS_SHEET_PASSWORD';

/**
 * 設定シートにパスワードロックをかけて非表示にする（初回1度だけ実行）。
 * デフォルトパスワード: admin1234
 */
function setupSettingsProtection() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(_PW_KEY)) {
    props.setProperty(_PW_KEY, 'admin1234');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.SETTINGS);

  // 既存保護を削除→再設定
  sh.getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .forEach(p => p.remove());
  const prot = sh.protect().setDescription('設定シート — パスワード保護');
  prot.removeEditors(prot.getEditors());
  if (prot.canDomainEdit()) prot.setDomainEdit(false);

  // シートを非表示
  sh.hideSheet();

  ss.toast(
    '設定シートを保護しました。\nデフォルトパスワード: admin1234\n' +
    'メニュー「🔒 設定を開く」からアクセスしてください。',
    '🔒 保護完了', 12
  );
}

/** パスワード認証して設定シートを表示する */
function openSettingsWithPassword() {
  const stored = PropertiesService.getScriptProperties().getProperty(_PW_KEY);

  if (!stored) { _revealSettings(); return; }  // パスワード未設定時はそのまま開く

  const input = Browser.inputBox(
    '🔒 設定シート認証',
    'パスワードを入力してください：',
    Browser.Buttons.OK_CANCEL
  );
  if (input === 'cancel') return;

  if (input === stored) {
    _revealSettings();
    SpreadsheetApp.getActiveSpreadsheet()
      .toast('認証成功。編集後は必ず「🔓 設定を閉じる」を実行してください。', '🔓 認証成功', 8);
  } else {
    SpreadsheetApp.getUi().alert('❌ パスワードが違います。アクセスを拒否しました。');
  }
}

function _revealSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.SETTINGS);
  sh.showSheet();
  ss.setActiveSheet(sh);
}

/** 設定シートを非表示にして閉じる */
function closeSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheetByName(SHEET.SETTINGS).hideSheet();
  ss.setActiveSheet(ss.getSheetByName(SHEET.DASHBOARD));
  ss.toast('設定シートを閉じました。', '🔒 ロック完了', 3);
}

/** 設定パスワードを変更する */
function changeSettingsPassword() {
  const props  = PropertiesService.getScriptProperties();
  const stored = props.getProperty(_PW_KEY) || '';

  const cur = Browser.inputBox(
    '🔑 パスワード変更 (1/2)',
    '現在のパスワードを入力してください：',
    Browser.Buttons.OK_CANCEL
  );
  if (cur === 'cancel') return;
  if (cur !== stored) {
    SpreadsheetApp.getUi().alert('❌ 現在のパスワードが違います。');
    return;
  }

  const nw = Browser.inputBox(
    '🔑 パスワード変更 (2/2)',
    '新しいパスワードを入力してください（空欄でキャンセル）：',
    Browser.Buttons.OK_CANCEL
  );
  if (nw === 'cancel' || nw.trim() === '') return;

  props.setProperty(_PW_KEY, nw.trim());
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('パスワードを変更しました。', '✅ 変更完了', 5);
}
