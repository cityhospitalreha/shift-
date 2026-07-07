// ============================================================
// staff.js — 職員名簿
// ============================================================
// ここを書き換えると、①入力アプリ・②管理アプリの両方に反映されます。
//
// id  : 職員を区別するための記号。重複しなければ何でもOK。
//       （データはこのidで管理するので、名前を後で修正してもデータは壊れません）
// name: 画面に表示される名前（苗字だけでOK）
// job : "ST" / "OT" / "PT" のいずれか
//
// ★後で本物の苗字リストをもらったら、nameの部分だけ書き換えてください。
// ============================================================

const STAFF = [
  // ---- ST ----
  { id: "st01", name: "職員A",  job: "ST" },
  { id: "st02", name: "職員B",  job: "ST" },
  { id: "st03", name: "職員C",  job: "ST" },
  { id: "st04", name: "職員D",  job: "ST" },
  { id: "st05", name: "職員E",  job: "ST" },
  { id: "st06", name: "職員F",  job: "ST" },
  { id: "st07", name: "職員G",  job: "ST" },

  // ---- OT ----
  { id: "ot01", name: "職員H",  job: "OT" },
  { id: "ot02", name: "職員I",  job: "OT" },
  { id: "ot03", name: "職員J",  job: "OT" },
  { id: "ot04", name: "職員K",  job: "OT" },
  { id: "ot05", name: "職員L",  job: "OT" },
  { id: "ot06", name: "職員M",  job: "OT" },
  { id: "ot07", name: "職員N",  job: "OT" },

  // ---- PT ----
  { id: "pt01", name: "職員O",  job: "PT" },
  { id: "pt02", name: "職員P",  job: "PT" },
  { id: "pt03", name: "職員Q",  job: "PT" },
  { id: "pt04", name: "職員R",  job: "PT" },
  { id: "pt05", name: "職員S",  job: "PT" },
  { id: "pt06", name: "職員T",  job: "PT" },
  { id: "pt07", name: "職員U",  job: "PT" },
];
