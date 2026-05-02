# SessionLogger - Windows 登入/登出追蹤系統

自動追蹤 Windows 電腦的所有帳號登入、登出時間，支援 Web 儀表板和每日 Email 報告。

## 功能

- **自動追蹤** — 每分鐘偵測目前線上使用者（支援所有登入方式：密碼、PIN、Hello）
- **每日報告** — 自動寄送 Excel 圖表到指定 Email
- **即時儀表板** — http://localhost:5000 查看使用時長圖表
- **安靜運行** — 在背景隱形執行，不彈出任何視窗

## 支援的帳號

yikai、ander、yeeli、judy0、spunk（可自行新增其他本機帳號）

## 安裝方式

### 1. 複製專案

```bash
git clone https://github.com/你的帳號/session-logger.git
cd session-logger
```

### 2. 安裝依賴

```cmd
pip install openpyxl
```

### 3. 執行安裝腳本

```powershell
powershell -ExecutionPolicy Bypass -File setup_scheduler.ps1
```

### 4. 啟動儀表板

```cmd
node dashboard_server.js
```

## 各檔案說明

| 檔案 | 用途 |
|------|------|
| `session_logger.ps1` | 主要追蹤腳本，每分鐘執行，偵測登入/登出 |
| `session_report.py` | 產生 Excel 報告並寄送 Email |
| `dashboard_server.js` | Web 儀表板（Node.js） |
| `run_logger.vbs` | 背景執行追蹤腳本（不起黑窗） |
| `run_dashboard.vbs` | 背景啟動儀表板 |
| `setup_scheduler.ps1` | 一次性設定排程任務 |

## 資料存放位置

- 登入記錄：`C:\Users\你的使用者名稱\session_log.csv`
- 執行狀態：`C:\Users\你的使用者名稱\AppData\Local\Temp\session_state.json`

## 設定每日 Email 報告

在 `session_report.py` 中設定你的 Email 和 App Password：

```python
GMAIL_USER = "your-email@gmail.com"
GMAIL_APP_PASSWORD = "xxxx xxxx xxxx xxxx"  # Gmail App Password
REPORT_TO = "your-email@gmail.com"
```

然後設定排程：
```powershell
# 每日 20:00 自動寄送報告
schtasks /create /tn SessionReport_Daily /tr "python session_report.py --email-only" /sc daily /st 20:00
```