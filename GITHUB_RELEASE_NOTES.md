v1.1.18 — Excel 元件本機化

1. 專案新增 xlsx.full.min.js（SheetJS 0.20.3）。
2. index.html 改為從本站載入元件，不再使用 SheetJS CDN。
3. Service Worker 將此檔案加入離線快取清單。
4. version.json 更新為 v1.1.18。
5. 統計公式、資料欄位及本機儲存不變；第 10 項清理未進行。

離線使用前須至少成功連線開啟網站，讓資源載入及快取完成。首次完全離線、尚未載入過網站的裝置無法開啟 GitHub Pages。
