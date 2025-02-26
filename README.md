# Instagram Messages to Excel

這個專案的目的是將 Instagram 訊息資料轉換為 Excel 檔案，並支持根據發件人名稱創建不同的分頁。每個訊息包括發件人名稱、時間戳、內容等，並且可以將訊息資料以 Excel 格式下載。

## 功能
- 解析 JSON 格式的 Instagram 訊息資料
- 依據發件人名稱為每個 JSON 檔案創建一個分頁
- 設定 Excel 標題與內容的樣式
- 自動調整欄寬和行高，確保每一個單元格的內容都清晰可見
- 將結果下載為 Excel 檔案

## 使用方式

### 安裝

1. 將專案下載或克隆到本地：
   ```bash
   git clone https://github.com/yourusername/instagram-messages-to-excel.git
   cd instagram-messages-to-excel
   pnpm install

