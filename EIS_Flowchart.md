# EIS 自動化流程圖

## 整體流程

```
WorkingNote_Cindy.xlsx ─┐
EIS.xlsx ───────────────┴─► [Step 1] monthly_eis_calculator.py ──► EIS_calculated.xlsx
                                                                          │
                                                                          ▼
                                                          [Step 2] eis_postprocess_extra_web.py
                                                                          │
                                                                          ▼
                                                                   EIS_final.xlsx
```

---

## Step 1 — monthly_eis_calculator.py

```
EIS.xlsx + WorkingNote.xlsx
         │
         ▼
讀取 WorkingNote 每一格「工作項目」欄
         │
         ▼
解析每一格內容：
  [ProjectName]          ← 標記目前 Project
  1. 任務名稱 *2         ← 任務 + 數量(qty)
         │
         ├─► Project 是 kore / other / rfi_rfq？
         │         └─► 寫入 Special_Category_Summary sheet（不填入 EIS）
         │
         └─► 一般 Project
                   │
                   ▼
            比對 EIS column A 任務名稱
            （依序：TASK_ALIASES → 完全符合 → 包含比對 → 模糊比對 cutoff=0.82）
                   │
                   ├─► 找不到 → 寫入 Unmatched_WorkingNote_Items sheet
                   │
                   └─► 找到 → 找 EIS 對應 Project 欄位
                                  │
                                  ├─► 欄位不存在 → 自動新增欄位（插在 Summary 左邊）
                                  └─► 累加 qty 到該儲存格
         │
         ▼
儲存 EIS_calculated.xlsx
```

---

## Step 2 — eis_postprocess_extra_web.py

```
EIS_calculated.xlsx
         │
         ▼
自動偵測 EIS 主工作表
（找 row1 含 Summary 且 row20 含 掛名人員）
         │
         ├─────────────────────────────────────────┐
         ▼                                         ▼
  【填寫 EIS_Extra sheet】                  【填寫 EIS_Web sheet】
         │                                         │
  掃描 row 20，找含「Extra」的欄位          EIS_Web 每一列（col A=員工, col D=專案）
         │                                         │
  ▸ col A ← EIS row 1 的 Project 名稱      比對員工是否在 EIS row 20
  ▸ col C ← EIS row 19 的值               且 Project 是否在 EIS row 1
              （空白 → 填 0）                       │
                                             ├─► 找到 → col E ← EIS row 19 的值
                                             └─► 找不到 → col E ← 0
         │                                         │
         └──────────────┬──────────────────────────┘
                        ▼
               儲存 EIS_final.xlsx
```
