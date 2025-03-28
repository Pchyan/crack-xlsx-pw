# Excel XLSX 密碼破解工具

這是一個用於破解受密碼保護的 Excel XLSX 檔案的 Python 工具。

## 安裝需求

1. Python 3.6 或更高版本
2. 安裝必要的套件：
   ```bash
   pip install -r requirements.txt
   ```

## 使用方法

### 1. 免安裝版本（推薦）

直接執行 `dist/Excel密碼破解工具.exe` 即可使用，無需安裝 Python 或其他套件。

### 2. 原始碼版本

#### 圖形介面版本
執行圖形介面程式：
```bash
python gui.py
```

圖形介面包含以下功能：
1. 密碼字典生成
   - 設定最小和最大密碼長度
   - 選擇輸出檔案位置
   - 顯示生成進度

2. Excel 檔案破解
   - 選擇要破解的 Excel 檔案
   - 選擇密碼字典檔案
   - 顯示破解進度和結果

3. 執行狀態顯示
   - 即時顯示操作進度
   - 顯示錯誤訊息
   - 支援捲動查看歷史訊息

#### 命令列版本

##### 生成密碼字典
```bash
python generate_dict.py
```

命令列參數：
- `-m, --min`：最小密碼長度（預設：4）
- `-M, --max`：最大密碼長度（預設：8）
- `-o, --output`：輸出檔案名稱（預設：password_dict.txt）

使用範例：
```bash
# 生成 6-10 位密碼
python generate_dict.py -m 6 -M 10

# 生成 4-8 位密碼並指定輸出檔案
python generate_dict.py -m 4 -M 8 -o my_passwords.txt
```

##### 破解 Excel 檔案
```bash
python crack_xlsx.py <xlsx檔案路徑> [密碼字典路徑]
```

### 參數說明
- `xlsx檔案路徑`：要破解的 Excel 檔案路徑（必填）
- `密碼字典路徑`：包含密碼列表的文字檔案路徑（選填）

### 密碼字典格式
密碼字典應該是一個文字檔案，每行一個密碼。例如：
```
password123
123456
admin
```

## 開發者說明

### 打包成執行檔

1. 安裝打包工具：
   ```bash
   pip install -r requirements.txt
   ```

2. 執行打包腳本：
   ```bash
   python build.py
   ```

3. 打包完成後，執行檔將位於：
   - 單一執行檔：`dist/Excel密碼破解工具.exe`
   - 完整程式目錄：`dist/Excel密碼破解工具/`

## 注意事項
1. 本工具僅供學習和研究使用
2. 請確保您有權限破解目標檔案
3. 如果沒有提供密碼字典，程式會使用預設的簡單密碼列表
4. 破解時間取決於密碼的複雜度和密碼列表的大小
5. 生成密碼字典時請注意硬碟空間，密碼數量可能非常大
6. 建議使用圖形介面版本，操作更直觀且可即時查看進度
7. 免安裝版本無需安裝 Python 或其他套件，可直接執行 