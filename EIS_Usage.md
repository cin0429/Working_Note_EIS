# EIS 自動化工具使用說明

## 使用方式

### Step 1 — monthly_eis_calculator.py

```bash
python monthly_eis_calculator.py \
  --eis EIS.xlsx \
  --working-notes WorkingNote_Cindy.xlsx \
  --output EIS_calculated.xlsx
```

| 參數 | 說明 |
|------|------|
| `--eis` | EIS 模板或當月空白檔 |
| `--working-notes` | 一個或多個 WorkingNote 檔（多個用空格分隔） |
| `--output` | 輸出檔案名稱 |

### Step 2 — eis_postprocess_extra_web.py

```bash
python eis_postprocess_extra_web.py
# 預設讀 EIS_calculated.xlsx，輸出 EIS_final.xlsx

# 或指定路徑：
python eis_postprocess_extra_web.py \
  --input EIS_calculated.xlsx \
  --output EIS_final.xlsx \
  --eis-sheet EIS      # EIS 主工作表名稱無法自動偵測時才需要加
```

| 參數 | 說明 |
|------|------|
| `--input` | 輸入檔，預設 `EIS_calculated.xlsx` |
| `--output` | 輸出檔，預設 `EIS_final.xlsx` |
| `--eis-sheet` | EIS 主工作表名稱，可省略（自動偵測） |

---

## 每月操作 SOP

```
1. 準備好當月 EIS.xlsx 和 WorkingNote_Cindy.xlsx

2. 執行 Step 1 → 產生 EIS_calculated.xlsx
   python monthly_eis_calculator.py --eis EIS.xlsx --working-notes WorkingNote_Cindy.xlsx --output EIS_calculated.xlsx

3. 開啟 EIS_calculated.xlsx，人工確認：
   - Unmatched_WorkingNote_Items sheet（有沒有沒對到的任務）
   - Special_Category_Summary sheet（kore / other / rfi_rfq 的彙總）
   - 手動填入 EIS row 19（建議申報值）

4. 存檔後執行 Step 2 → 產生 EIS_final.xlsx
   python eis_postprocess_extra_web.py
```
