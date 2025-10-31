# AutoZ Wafer4P Aligner - Axis Selection 功能擴展修改計畫

## 修改目標

將 Axis Selection 功能從僅影響 "AutoZ Values" 圖表，擴展至同時控制 "Value Anomaly Analysis" 圖表。

---

## 修改範圍

### 1. Python 後端函數

#### 1.1 `create_z_anomaly_chart()` → `create_anomaly_chart()`

**修改原因**：
- 當前函數硬編碼為 Z 軸專用
- 需重構為通用函數，支持 X/Y/Z 三軸

**修改內容**：
```python
# 原函數簽名
def create_z_anomaly_chart(wafer_data, z_standard, standard_point_data=None)

# 新函數簽名
def create_anomaly_chart(wafer_data, axis_type, standard_value, standard_point_data=None)
```

**主要變更**：
- 參數 `z_standard` → `standard_value`（通用標準值）
- 新增參數 `axis_type`（'x', 'y', 'z'）
- 數據提取邏輯改為 `data[f'{axis_type}_values']`
- 異常判斷邏輯改為 `value < standard_value`（適用所有軸）
- 圖表標題動態化：`f"{axis_type.upper()} Value Anomaly Analysis"`
- Y 軸標題動態化：`f"{axis_type.upper()} Value (µm)"`
- 標準線標籤動態化：`f"{axis_type.upper()} Standard ({standard_value} µm)"`

---

#### 1.2 `/api/regenerate_chart` 端點

**修改原因**：
- 當前僅返回單一圖表數據
- 需同時返回 AutoZ Values 和 Anomaly Analysis 兩個圖表

**修改內容**：

```python
@app.route('/api/regenerate_chart', methods=['POST'])
def regenerate_chart():
    # ... 前置檢查保持不變 ...
    
    # 獲取對應軸的標準值
    standard_map = {
        'x': x_standard,
        'y': y_standard,
        'z': z_standard
    }
    standard_value = standard_map.get(axis_type)
    
    # 生成主圖表（保持原邏輯）
    main_fig, stats = create_line_chart(
        wafer_data, 
        axis_type, 
        standard_value if axis_type == 'z' else None,  # 僅 Z 軸顯示標準線
        standard_point_data
    )
    
    # 【新增】生成異常分析圖表
    anomaly_fig, anomaly_stats = create_anomaly_chart(
        wafer_data,
        axis_type,
        standard_value,
        standard_point_data
    )
    
    # 【修改】返回結構
    return jsonify({
        'success': True,
        'main_chart': main_fig.to_dict(),
        'anomaly_chart': anomaly_fig.to_dict(),  # 新增
        'stats': stats,
        'anomaly_stats': anomaly_stats  # 新增（可選）
    })
```

---

#### 1.3 `generate_result_html()`

**修改原因**：
- 初始頁面需正確設置圖表容器
- JavaScript 需能識別並更新異常圖表

**修改內容**：

**HTML 結構調整**：
```html
<!-- 原結構 -->
<div class="chart-container">
    <div class="chart-title">Z Value Anomaly Analysis</div>
    {z_anomaly_html}
</div>

<!-- 新結構 -->
<div class="chart-container" id="anomalyChartContainer">
    <div class="chart-title" id="anomalyChartTitle">Z Value Anomaly Analysis</div>
    <div id="anomalyChart">{z_anomaly_html}</div>
</div>
```

**JavaScript `switchAxis()` 函數修改**：
```javascript
async function switchAxis(axisType) {
    // ... 按鈕狀態更新保持不變 ...
    
    if (result.success) {
        // ... 更新統計數據（保持不變）...
        
        // 更新主圖表（保持不變）
        const chartContainer = document.getElementById('autoZValuesChartContainer');
        chartContainer.innerHTML = '<div class="chart-title">' + axisType.toUpperCase() + ' AutoZ Values</div><div id="newChart"></div>';
        Plotly.newPlot('newChart', result.main_chart.data, result.main_chart.layout, {responsive: true});
        
        // 【新增】更新異常分析圖表
        const anomalyContainer = document.getElementById('anomalyChartContainer');
        anomalyContainer.innerHTML = `
            <div class="chart-title" id="anomalyChartTitle">${axisType.toUpperCase()} Value Anomaly Analysis</div>
            <div id="anomalyChart"></div>
        `;
        Plotly.newPlot('anomalyChart', result.anomaly_chart.data, result.anomaly_chart.layout, {responsive: true});
    }
}
```

---

## 修改步驟

### Step 1: 重構 `create_z_anomaly_chart()`
1. 複製原函數代碼
2. 重命名為 `create_anomaly_chart()`
3. 修改參數列表
4. 替換所有硬編碼的 `'z'` 為 `axis_type` 變量
5. 替換 `z_standard` 為 `standard_value`
6. 動態生成圖表標題和軸標籤
7. 測試 X/Y/Z 三軸生成

### Step 2: 修改 API 端點
1. 在 `/api/regenerate_chart` 中調用新函數
2. 構建 `standard_map` 映射表
3. 添加異常圖表生成邏輯
4. 修改返回的 JSON 結構
5. 測試 API 返回數據完整性

### Step 3: 更新前端
1. 修改 HTML 圖表容器結構
2. 添加可識別的 ID
3. 更新 `switchAxis()` JavaScript 函數
4. 添加第二個 Plotly 圖表渲染邏輯
5. 測試三軸切換流暢性

---

## 修改前後對比

### 當前行為
```
選擇 X 軸 → 上方顯示 X AutoZ Values + 下方仍顯示 Z Anomaly（錯誤）
選擇 Y 軸 → 上方顯示 Y AutoZ Values + 下方仍顯示 Z Anomaly（錯誤）
選擇 Z 軸 → 上方顯示 Z AutoZ Values + 下方顯示 Z Anomaly（正確）
```

### 修改後行為
```
選擇 X 軸 → 上方顯示 X AutoZ Values + 下方顯示 X Anomaly（正確）
選擇 Y 軸 → 上方顯示 Y AutoZ Values + 下方顯示 Y Anomaly（正確）
選擇 Z 軸 → 上方顯示 Z AutoZ Values + 下方顯示 Z Anomaly（正確）
```

---

## 風險評估

### 低風險
- ✅ 不影響現有 Z 軸功能（向下兼容）
- ✅ 不修改數據庫結構
- ✅ 不影響文件上傳邏輯

### 中風險
- ⚠️ 需確保 X/Y 標準值正確傳遞
- ⚠️ AutoZ Complete 點位在 X/Y 軸的識別邏輯

### 注意事項
1. **標準值判斷邏輯**：確認 X/Y 軸是否也使用 `< standard` 判斷異常（可能需要調整）
2. **圖表顏色一致性**：異常點配色在三軸保持統一
3. **統計數據**：`anomaly_stats` 是否需要在前端顯示

---

## 測試計畫

### 單元測試
- [ ] `create_anomaly_chart('x', ...)` 正確生成 X 軸圖表
- [ ] `create_anomaly_chart('y', ...)` 正確生成 Y 軸圖表
- [ ] `create_anomaly_chart('z', ...)` 與原函數結果一致

### 集成測試
- [ ] API `/api/regenerate_chart` 返回完整數據結構
- [ ] 前端切換 X/Y/Z 軸，兩個圖表同步更新
- [ ] AutoZ Complete 點位在三軸都正確標示

### 回歸測試
- [ ] 原有 Z 軸功能不受影響
- [ ] Wafer Status Dashboard 正常顯示
- [ ] 統計數據計算正確

---

## 預估工作量

- **代碼修改**：1-2 小時
- **測試驗證**：1 小時
- **總計**：2-3 小時

---

## 審核確認

- [ ] 修改計畫已審閱
- [ ] 邏輯正確性已確認
- [ ] 可開始實施修改

---

**備註**：修改過程中如發現其他需調整的邏輯（如 X/Y 軸異常判斷標準），將另行提出討論。
