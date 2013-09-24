使用方法：

1. 建立表單(form)
2. 新增「回覆內容」試算表(spreadsheet)
3. 新增一頁資料表，專門放真正的報名名單(會去掉重複報名者的資料)
4. 從「工具」選擇「指令碼編輯器」，將source code貼上並存檔
5. 找到變數URL，將spreadsheet的網址貼上存到變數URL中
6. 根據表單項目，修改e.values[x]的資料存到的變數名稱(依序)
7. 重複驗證依據是Email的重複
8. 找到 sheet.getRange(data.length, 7).setValue(i+1); 指令，將7改成(表單儲存欄位項目數量+1)
9. 更改 mailBody 變數(信件內容)
10. 從資源(Source)設定觸發器(Trigger)，條件為「當表單提交時」。
11. 給予授權。
12. 即可使用。


Use Google App Script
