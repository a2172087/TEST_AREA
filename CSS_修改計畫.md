# AutoZ Wafer4P Aligner - CSS 修改計畫

## 修改範圍確認

**目標檔案**: `AutoZ_Wafer4P_Aligner_V13_0_8.py`  
**目標函數**: `generate_result_html(data)`  
**函數位置**: 第 1966 行 - 第 2611 行  
**修改目的**: 修正 `/result` 頁面的 Tab 視覺效果，使其與 Wafer Map Stack Analysis 一致

---

## 必須修改（影響視覺比例）

### 修改 1: `.tabs` 容器圓角
- **行數**: 第 2082 行
- **目的**: 消除黑色背景的視覺斷層

**修改前**:
```css
.tabs {{
    display: flex;
    background-color: #2D2D2D;
    padding: 10px 10px 0 10px;
    border-radius: 0;
}}
```

**修改後**:
```css
.tabs {{
    display: flex;
    background-color: #2D2D2D;
    padding: 10px 10px 0 10px;
    border-radius: 8px 8px 0 0;
}}
```

**變更內容**: 
```
border-radius: 0;  →  border-radius: 8px 8px 0 0;
```

---

### 修改 2: `.tab` 按鈕內距
- **行數**: 第 2086 行
- **目的**: 讓按鈕更精緻，不會顯得笨重

**修改前**:
```css
.tab {{
    padding: 12px 24px;
    cursor: pointer;
    background-color: #454545;
    color: #E0E0E0;
    margin-right: 5px;
    border-radius: 8px 8px 0 0;
    font-weight: bold;
    font-size: 15px;
    transition: all 0.3s ease;
}}
```

**修改後**:
```css
.tab {{
    padding: 10px 20px;
    cursor: pointer;
    background-color: #454545;
    color: #E0E0E0;
    margin-right: 5px;
    border-radius: 8px 8px 0 0;
    font-weight: bold;
    font-size: 15px;
    transition: all 0.3s ease;
}}
```

**變更內容**: 
```
padding: 12px 24px;  →  padding: 10px 20px;
```

---

### 修改 3: `.tab` 字體大小
- **行數**: 第 2093 行
- **目的**: 字體大小與參考專案一致

**修改前**:
```css
.tab {{
    padding: 10px 20px;
    cursor: pointer;
    background-color: #454545;
    color: #E0E0E0;
    margin-right: 5px;
    border-radius: 8px 8px 0 0;
    font-weight: bold;
    font-size: 15px;
    transition: all 0.3s ease;
}}
```

**修改後**:
```css
.tab {{
    padding: 10px 20px;
    cursor: pointer;
    background-color: #454545;
    color: #E0E0E0;
    margin-right: 5px;
    border-radius: 8px 8px 0 0;
    font-weight: bold;
    transition: all 0.3s ease;
}}
```

**變更內容**: 
```
刪除這一行: font-size: 15px;
```

---

## 可選修改（不影響主要視覺）

以下修改不影響靜態視覺效果，建議**保留不變**：

### 保留項目 1: `.tab` 的 transition
- **行數**: 第 2094 行
- **說明**: 動畫過渡效果，提升使用者體驗
- **建議**: **保留**

### 保留項目 2: `.tab:hover` 區塊
- **行數**: 第 2097-2099 行
- **說明**: 滑鼠懸停變色效果
- **建議**: **保留**

---

## 修改摘要

| 項目 | 行數 | 原始值 | 修改值 | 影響 |
|------|------|--------|--------|------|
| `.tabs` border-radius | 2082 | `0` | `8px 8px 0 0` | 消除視覺斷層 |
| `.tab` padding | 2086 | `12px 24px` | `10px 20px` | 按鈕更精緻 |
| `.tab` font-size | 2093 | `15px` | *刪除此行* | 使用預設 14px |

---

## 執行指令

### 使用 `str_replace` 工具進行修改

#### 修改 1: 更新 `.tabs` 的 border-radius
```python
str_replace(
    path="/mnt/project/AutoZ_Wafer4P_Aligner_V13_0_8.py",
    old_str="            .tabs {{\n                display: flex;\n                background-color: #2D2D2D;\n                padding: 10px 10px 0 10px;\n                border-radius: 0;\n            }}",
    new_str="            .tabs {{\n                display: flex;\n                background-color: #2D2D2D;\n                padding: 10px 10px 0 10px;\n                border-radius: 8px 8px 0 0;\n            }}"
)
```

#### 修改 2: 更新 `.tab` 的 padding
```python
str_replace(
    path="/mnt/project/AutoZ_Wafer4P_Aligner_V13_0_8.py",
    old_str="            .tab {{\n                padding: 12px 24px;",
    new_str="            .tab {{\n                padding: 10px 20px;"
)
```

#### 修改 3: 刪除 `.tab` 的 font-size
```python
str_replace(
    path="/mnt/project/AutoZ_Wafer4P_Aligner_V13_0_8.py",
    old_str="                border-radius: 8px 8px 0 0;\n                font-weight: bold;\n                font-size: 15px;\n                transition: all 0.3s ease;",
    new_str="                border-radius: 8px 8px 0 0;\n                font-weight: bold;\n                transition: all 0.3s ease;"
)
```

---

## 預期效果

修改完成後，AutoZ Wafer4P Aligner 的 `/result` 頁面應該會呈現：

1. ✅ Header 和 Tabs 的黑色背景完美融合，無斷層感
2. ✅ Tab 按鈕大小適中，視覺更精緻
3. ✅ 字體大小與 Wafer Map Stack Analysis 一致
4. ✅ 整體視覺比例協調統一

---

## 備註

- 所有修改僅涉及 CSS 樣式，不影響功能邏輯
- 建議在修改後測試頁面顯示效果
- 原始檔案建議先備份
- 修改時注意保持 Python f-string 的雙大括號 `{{` 語法

---

## 版本資訊

- **原始版本**: AutoZ_Wafer4P_Aligner_V13_0_8.py
- **修改日期**: 2025-10-31
- **修改內容**: Tab 系統 CSS 樣式優化
- **參考專案**: Wafer_Map_Stack_Analysis_V6_0.py
