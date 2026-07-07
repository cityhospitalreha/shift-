// ============================================================
// Code.gs — シフト希望収集システム（GAS側プログラム）
// ============================================================
// このコードをGoogleスプレッドシートの Apps Script に貼り付けて
// 「ウェブアプリ」としてデプロイすると、専用URLが発行されます。
//
// 役割は2つ:
//   doPost … ①入力アプリから届いた希望データをシートに保存する係
//   doGet  … 設定（対象月・締切）や、保存済みデータを渡す係
//
// 使用するシート（タブ）:
//   「設定」 A2=対象月(例 2026-08)  B2=締切日(例 2026-07-20)
//   「回答」 送信が1件届くたびに1行追記される（履歴として全部残る）
// ============================================================

const SHEET_SETTINGS = "設定";
const SHEET_RESPONSES = "回答";

// ------------------------------------------------------------
// 共通: JSONを返す（アプリ側のJavaScriptが読み取れる形式）
// ------------------------------------------------------------
function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// 共通: 日本時間の今日を "YYYY-MM-DD" 形式で返す
// （タイムゾーン設定に左右されないよう Asia/Tokyo を明示）
// ------------------------------------------------------------
function todayJST_() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
}

// ------------------------------------------------------------
// 共通: 設定シートから対象月と締切を読む
// ------------------------------------------------------------
function getSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sh) throw new Error("「設定」シートが見つかりません");

  // getDisplayValue = セルに見えているままの文字を取得
  // （日付セルが勝手にDate型になるトラブルを避けるため）
  const month = String(sh.getRange("A2").getDisplayValue()).trim();
  const deadline = String(sh.getRange("B2").getDisplayValue()).trim();
  return { month: month, deadline: deadline };
}

// ============================================================
// doGet — アプリからの「取得」リクエスト窓口
//   ?action=settings              → 対象月・締切・今日の日付
//   ?action=responses&month=YYYY-MM → その月の全員分の最新回答
// ============================================================
function doGet(e) {
  try {
    const action = (e.parameter.action || "settings");

    if (action === "settings") {
      const s = getSettings_();
      return jsonOut_({
        ok: true,
        month: s.month,
        deadline: s.deadline,
        today: todayJST_(),
        closed: todayJST_() > s.deadline   // 文字列比較でOK（同じ桁数のため）
      });
    }

    if (action === "responses") {
      const month = e.parameter.month || getSettings_().month;
      return jsonOut_({ ok: true, month: month, responses: getLatestResponses_(month) });
    }

    return jsonOut_({ ok: false, error: "不明なaction: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ============================================================
// doPost — アプリからの「送信」リクエスト窓口
//   ①入力アプリが {action:"submit", ...} を送ってくる
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === "submit") {
      return handleSubmit_(data);
    }

    return jsonOut_({ ok: false, error: "不明なaction" });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ------------------------------------------------------------
// 希望データの保存処理
// ------------------------------------------------------------
function handleSubmit_(data) {
  const s = getSettings_();

  // --- 締切チェック（サーバー側の二重ロック） ---
  // アプリ側でもロックするが、万一URLを直接叩かれても
  // ここで拒否されるので締切後のデータは絶対に入らない。
  if (todayJST_() > s.deadline) {
    return jsonOut_({ ok: false, error: "締切（" + s.deadline + "）を過ぎているため受付できません" });
  }

  // --- 対象月チェック ---
  if (data.month !== s.month) {
    return jsonOut_({ ok: false, error: "対象月が現在の募集（" + s.month + "）と一致しません。ページを再読み込みしてください" });
  }

  // --- 必須項目チェック ---
  if (!data.staffId || !data.name || !data.answers) {
    return jsonOut_({ ok: false, error: "データが不足しています" });
  }

  // --- 保存（1行追記） ---
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_RESPONSES);
  if (!sh) {
    // 「回答」シートがなければ自動作成して見出し行を付ける
    sh = ss.insertSheet(SHEET_RESPONSES);
    sh.appendRow(["受信日時", "対象月", "職員ID", "名前", "職種", "回答JSON", "コメント"]);
  }

  sh.appendRow([
    Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss"),
    data.month,
    data.staffId,
    data.name,
    data.job || "",
    JSON.stringify(data.answers),   // {"2026-08-01":"◎", ...} を文字列で保存
    data.comment || ""
  ]);

  return jsonOut_({ ok: true, message: "受付完了" });
}

// ------------------------------------------------------------
// 指定月の「各職員の最新回答」を集める
// （同じ人が複数回送信していたら、いちばん新しい行だけ採用）
// ------------------------------------------------------------
function getLatestResponses_(month) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESPONSES);
  if (!sh) return [];

  const rows = sh.getDataRange().getValues();  // シート全体を2次元配列で取得
  const latest = {};  // staffId → 行データ

  for (let i = 1; i < rows.length; i++) {      // i=0 は見出し行なので飛ばす
    const [ts, m, staffId, name, job, answersJson, comment] = rows[i];
    if (String(m) !== String(month)) continue; // 対象月以外は無視
    // 下の行ほど新しいので、単純に上書きしていけば最後に残るのが最新
    latest[staffId] = {
      staffId: String(staffId),
      name: String(name),
      job: String(job),
      answers: JSON.parse(answersJson),
      comment: String(comment),
      sentAt: String(ts)
    };
  }

  return Object.values(latest);
}
