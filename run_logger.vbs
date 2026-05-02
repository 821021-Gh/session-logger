Set WShell = CreateObject("WScript.Shell")
WShell.Run """powershell.exe"" -WindowStyle Hidden -ExecutionPolicy Bypass -File ""C:\Users\spunk\.qclaw-oversea\workspace\session_logger.ps1""", 0, False
Set WShell = Nothing