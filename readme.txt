介紹：一個將COPR19、COPR66、Richtek三個檔案整合成MTC_POsummary半成品的小工具。

使用說明：
1. 開啟"訂單整合小工具.exe"
2. 點選"選擇檔案"按鈕選擇三個檔案輸入
3. 點選"輸出結果"按鈕選擇輸出位置和檔名
4. 即可得到檔名"XXX.xlsx"的半成品輸出和檔名"XXX_error.xlsx"的Qty不相符清單

注意事項：
1. 不會輸出"Order Run Type"和"TTS"欄位
2. "Required Qty"欄位會直接填入COPR19和COPR66的資料
3. "XXX_error.xlsx"中，會計算同Customer PO和Customer Production訂單的Qty總額再與Richtek比較