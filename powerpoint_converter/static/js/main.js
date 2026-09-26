"use strict";
// 起動: サーバの情報を読んでから、URL の # に合わせて一覧か資料を開く
(async () => {
  try { S.info = await api("/api/info"); } catch { }
  route();
})();
