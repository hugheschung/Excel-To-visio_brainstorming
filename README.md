# Excel-To-visio_brainstorming_v1.0
一個好用的 excel VBA 轉換 visio brainstorming(mind map)的程式碼
excel VBA to visio brainstorming(mind map)

---
## 使用方法：
1. 建立Excel，將資料按照階層排序，如下圖
![image](https://github.com/hugheschung/Excel-To-visio_brainstorming/blob/main/001.png)
2. 使用VBA巨集(ExcelToV_v1.0.bas)，進行轉換visio xml ，檔名預設存為 ExcelToV.xml
(註:如何使用Excel VBA這邊不再介紹了，請看官自行去學習匯入使用)
3. 使用note pad++(或可以轉換編碼的軟體)，將編碼轉為uft-8，如下圖
![image](https://github.com/hugheschung/Excel-To-visio_brainstorming/blob/main/003.png)
4. 切換到visio brainstorming 匯入方才轉換的 ExcelToV.xml檔案，如下圖
![image](https://github.com/hugheschung/Excel-To-visio_brainstorming/blob/main/002.png)

---
## 目前問題：
 a1.原本有設計自動轉存uft-8格式，但是中文字會亂碼，有想幾個方案，但礙於程式碼可能有非big-5編碼的使用者使用，所以還是請大家手段轉換好了


---
## 作者的廢話
hello!
因為使用visio這套軟體的時候，發現不能透過excel將階層化的文字直接轉入 visio，我就寫了這個VBA code，
希望可以幫助到有需要的人!!

一開始我也是四處爬文，發現國外的論壇有滿多人有這個需求，但是有需求，卻沒有人寫代碼，有可能是根本的需求可能不大，可能大家另尋其他mind map軟體也說不定...

---
