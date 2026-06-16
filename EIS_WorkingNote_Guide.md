# WorkingNote 工作內容填寫規範

自動化程式會讀取 WorkingNote 的「工作項目」欄，並依據填寫格式將工作時數統計至 EIS。
**若填寫格式不正確，程式將無法自動對應，需人工補正。**

---

## 填寫格式

```
Project Name - EIS 工作項目名稱 -xN
```

| 元素 | 說明 | 範例 |
|------|------|------|
| `Project Name` | 專案名稱，放在最前面 | `Aymara` |
| `-` | 用 `-` 分隔各個欄位 | |
| EIS 工作項目名稱 | 必須對應 EIS 中 column A 的工作項目，越接近越好 | `Release MSA Report` |
| `-xN` | 次數，N 為數字；若省略則預設為 1 次 | `-x2`、`-x1` |

### 範例

```
Aymara - Release MSA Report -x1
Aymara - Review Logs (Fail / Incosistant / Mixed Logs) -x3
XYZ_Project - Provide Measurement Distribution Chart -x2
XYZ_Project - Review Logs (Fail / Incosistant / Mixed Logs)
```

> 最後一行沒有 `-xN`，程式自動計為 1 次。

---

## 特殊分類（無法對應到特定專案時使用）

若工作內容**無法歸入現有 EIS 工作項目**，請依下列規則選擇分類：

| 分類 | 使用時機 | 範例 |
|------|----------|------|
| `RFI_RFQ` | RFI / RFQ 相關工作，但**專案名稱尚未確定** | 回覆客戶 RFQ 詢問、準備報價資料 |
| `Kore` | **SWQA 相關工作**，但不屬於任何特定專案 | SWQA 內部流程改善、跨專案支援 |
| `Other` | **與 SWQA 無關**的其他工作 | 學習新技術、練習程式碼、行政事務 |

### 特殊分類範例

```
RFI_RFQ - Review RFQ document -x1
RFI_RFQ - Prepare quotation material -x2
Kore - SWQA internal process review -x1
Other - Study new testing framework -x1
Other - Coding practice -x2
```

> 特殊分類的工作項目**不需要對應 EIS column A**，程式會另外統整至 `Special_Category_Summary` sheet，不計入各專案欄位。

---

## 注意事項

1. **Project Name 和工作項目之間用 `-` 分隔**，前後可加空格。
2. **工作項目名稱盡量與 EIS column A 一致**，程式雖有模糊比對，但越接近越不容易出錯。
3. **次數 `-xN` 若省略，預設計為 1 次**。
4. 可在工作項目後面加狀態標籤如 `[Complete]`、`[In Progress]`，程式會自動忽略，不影響計算。
5. 若填寫後發現工作項目被歸入 `Unmatched_WorkingNote_Items`，代表程式找不到對應的 EIS 工作項目，請確認拼字是否正確，或告知負責人新增對應別名。

---

## 快速對照表：我的工作應該用哪個分類？

```
這個工作有對應的 EIS 工作項目嗎？
         │
         ├─► 有 → ProjectName - EIS 工作項目名稱 -xN
         │
         └─► 沒有
                   │
                   ├─► 是 RFI/RFQ 相關，且專案名稱未定 → RFI_RFQ - 工作描述 -xN
                   │
                   ├─► 是 SWQA 相關，但不屬於特定專案 → Kore - 工作描述 -xN
                   │
                   └─► 與 SWQA 無關（自我學習、行政等） → Other - 工作描述 -xN
```
