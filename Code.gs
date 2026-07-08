// ============================================================
// Code.gs — シフト希望収集システム（GAS側プログラム）v2
// ============================================================
// v2での追加：入力ルール（◎の上限・×の上限）と、
//             設定ページ(settings.html)からの設定変更受付
//
// 使用するシート（タブ）:
//   「設定」 A2=対象月  B2=締切日  C2=◎上限  D2=×上限  E2=管理パスワード
//            （上限は 0 または空欄 = 無制限）
//   「回答」 送信が1件届くたびに1行追記される
// ============================================================

const SHEET_SETTINGS = "設定";
const SHEET_RESPONSES = "回答";

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function todayJST_() {
  return Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd");
}

// ------------------------------------------------------------
// 日付文字列の正規化（ここが今回のNaN/undefined対策の核心）
// ------------------------------------------------------------
// スプレッドシートのセルは、入力の仕方や表示形式次第で
//   "2026-09-01" / "2026/09/01" / "2026年9月1日" / Date型オブジェクト
// など、いろいろな見た目でGASに渡ってくる可能性があります。
// ここで一度すべて "YYYY-MM-DD" に統一してしまえば、
// 以降の比較・分割処理が絶対にNaN/undefinedにならなくなります。
//
// 戻り値: 正規化できたら "YYYY-MM-DD"、できなければ null
function normalizeDate_(raw) {
  if (raw === null || raw === undefined || raw === "") return null;

  // Date型オブジェクトのまま渡ってきた場合（セルが日付形式のとき）
  if (Object.prototype.toString.call(raw) === "[object Date]") {
    return Utilities.formatDate(raw, "Asia/Tokyo", "yyyy-MM-dd");
  }

  // 文字列化して、全角数字→半角、区切り文字（/ 年 月 日 .）を統一
  let s = String(raw).trim();
  s = s.replace(/[０-９]/g, ch => "0123456789"["０１２３４５６７８９".indexOf(ch)]);
  s = s.replace(/[年月.]/g, "-").replace(/日/g, "").replace(/\//g, "-");
  s = s.replace(/-+/g, "-").replace(/-$/, "");

  const m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return null;
  const y = m[1], mo = String(m[2]).padStart(2, "0"), d = String(m[3]).padStart(2, "0");
  return `${y}-${mo}-${d}`;
}

// "YYYY-MM-DD" または "YYYY-MM" を受け取り、"YYYY-MM" だけを返す
function normalizeMonth_(raw) {
  if (raw === null || raw === undefined || raw === "") return null;

  if (Object.prototype.toString.call(raw) === "[object Date]") {
    return Utilities.formatDate(raw, "Asia/Tokyo", "yyyy-MM");
  }

  let s = String(raw).trim();
  s = s.replace(/[０-９]/g, ch => "0123456789"["０１２３４５６７８９".indexOf(ch)]);
  s = s.replace(/[年月.]/g, "-").replace(/日/g, "").replace(/\//g, "-");
  s = s.replace(/-+/g, "-").replace(/-$/, "");

  // "2026-09-01"（日付まで入力）でも "2026-09"（月だけ）でもOKにする
  const m = s.match(/^(\d{4})-(\d{1,2})(?:-\d{1,2})?$/);
  if (!m) return null;
  return `${m[1]}-${String(m[2]).padStart(2, "0")}`;
}

// ------------------------------------------------------------
// 設定シートの読み取り（正規化＋不正値の検出つき）
// ------------------------------------------------------------
function getSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!sh) throw new Error("「設定」シートが見つかりません");

  // getDisplayValues ではなく getValues を使う。
  // 表示形式（例：セルが「日付」書式だと年月日が省略表示されることがある）
  // に影響されず、常に元の値（文字列 or Date型）を取得するため。
  const row = sh.getRange("A2:E2").getValues()[0];

  const month = normalizeMonth_(row[0]);
  const deadline = normalizeDate_(row[1]);

  // ここで検出しておくことで、「あとでNaNやundefinedとして
  // 画面に漏れ出る」のではなく「原因がわかるエラー」として弾ける。
  if (!month) {
    throw new Error(
      "設定シートA2（対象月）の形式が読み取れません。現在の値: 「" + row[0] +
      "」。A2には 2026-09 のように年-月の形式で入力してください（B1など別セルではなくA2に入れる必要があります）"
    );
  }
  if (!deadline) {
    throw new Error(
      "設定シートB2（締切日）の形式が読み取れません。現在の値: 「" + row[1] +
      "」。B2には 2026-08-20 のように年-月-日の形式で入力してください"
    );
  }

  return {
    month: month,       // 常に "YYYY-MM" 形式で返す
    deadline: deadline, // 常に "YYYY-MM-DD" 形式で返す
    maxMaru2: Number(row[2]) || 0,
    maxBatsu: Number(row[3]) || 0,
    password: String(row[4] || "").trim()
  };
}

// ============================================================
// doGet — 取得の窓口
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
        closed: todayJST_() > s.deadline,
        rules: { maxMaru2: s.maxMaru2, maxBatsu: s.maxBatsu }
        // ※passwordは絶対に返さない
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
// doPost — 送信の窓口
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === "submit")      return handleSubmit_(data);
    if (data.action === "setSettings") return handleSetSettings_(data);
    return jsonOut_({ ok: false, error: "不明なaction" });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ------------------------------------------------------------
// 設定の変更（settings.htmlから届く）
// パスワードが一致したときだけ書き換える
// ------------------------------------------------------------
function handleSetSettings_(data) {
  const s = getSettings_();

  if (!s.password) {
    return jsonOut_({ ok: false, error: "設定シートのE2に管理パスワードが設定されていません" });
  }
  if (String(data.password) !== s.password) {
    return jsonOut_({ ok: false, error: "パスワードが違います" });
  }

  // 形式チェック（正規化できない値は保存させない）
  const month = normalizeMonth_(data.month);
  const deadline = normalizeDate_(data.deadline);
  if (!month)    return jsonOut_({ ok:false, error:"対象月の形式が不正です（例: 2026-09）" });
  if (!deadline) return jsonOut_({ ok:false, error:"締切日の形式が不正です（例: 2026-08-20）" });

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  sh.getRange("A2").setValue(month);
  sh.getRange("B2").setValue(deadline);
  sh.getRange("C2").setValue(Number(data.maxMaru2) || 0);
  sh.getRange("D2").setValue(Number(data.maxBatsu) || 0);

  return jsonOut_({ ok: true, message: "設定を保存しました" });
}

// ------------------------------------------------------------
// 希望データの保存
// ------------------------------------------------------------
function handleSubmit_(data) {
  const s = getSettings_();

  if (todayJST_() > s.deadline) {
    return jsonOut_({ ok: false, error: "締切（" + s.deadline + "）を過ぎているため受付できません" });
  }
  if (data.month !== s.month) {
    return jsonOut_({ ok: false, error: "対象月が現在の募集（" + s.month + "）と一致しません。ページを再読み込みしてください" });
  }
  if (!data.staffId || !data.name || !data.answers) {
    return jsonOut_({ ok: false, error: "データが不足しています" });
  }

  // --- 入力ルールの検証（アプリ側と二重チェック） ---
  const marks = Object.values(data.answers);
  const nMaru2 = marks.filter(m => m === "◎").length;
  const nBatsu = marks.filter(m => m === "×").length;
  if (s.maxMaru2 > 0 && nMaru2 > s.maxMaru2) {
    return jsonOut_({ ok: false, error: "「◎」は1人 " + s.maxMaru2 + " 日までです（現在 " + nMaru2 + " 日）" });
  }
  if (s.maxBatsu > 0 && nBatsu > s.maxBatsu) {
    return jsonOut_({ ok: false, error: "「×」は1人 " + s.maxBatsu + " 日までです（現在 " + nBatsu + " 日）" });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_RESPONSES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_RESPONSES);
    sh.appendRow(["受信日時", "対象月", "職員ID", "名前", "職種", "回答JSON", "コメント"]);
  }

  sh.appendRow([
    Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm:ss"),
    data.month,
    data.staffId,
    data.name,
    data.job || "",
    JSON.stringify(data.answers),
    data.comment || ""
  ]);

  return jsonOut_({ ok: true, message: "受付完了" });
}

// ------------------------------------------------------------
// 指定月の「各職員の最新回答」
// ------------------------------------------------------------
function getLatestResponses_(month) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RESPONSES);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const latest = {};
  for (let i = 1; i < rows.length; i++) {
    const [ts, m, staffId, name, job, answersJson, comment] = rows[i];
    if (String(m) !== String(month)) continue;
    latest[staffId] = {
      staffId: String(staffId), name: String(name), job: String(job),
      answers: JSON.parse(answersJson), comment: String(comment), sentAt: String(ts)
    };
  }
  return Object.values(latest);
}
