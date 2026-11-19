# Shows a Windows notification after report.py has run successfully.
# Claude Code sends details of the finished command as JSON on standard input.
# This script does not read, change or delete anything in data/ or output/.

[Console]::InputEncoding = [Text.Encoding]::UTF8
try { $data = [Console]::In.ReadToEnd() | ConvertFrom-Json } catch { exit 0 }

# 1. Only a command that runs report.py. The interpreter must be at the start of a command
#    (start of line, or after ; & | or "("), so commands that only read or search report.py
#    (Get-Content, cat, type, Select-String, grep ...) do not match. Accepted forms:
#      py report.py data/FILE.csv
#      python report.py ...        & "C:\...\python.exe" report.py ...      & 'C:\...\py.exe' report.py ...
$exe = '(?:py|python3?)(?:\.exe)?'
$interpreter = '(?:"(?:[^"]*[\\/])?' + $exe + '"|''(?:[^'']*[\\/])?' + $exe + '''|(?:[^\s''"]*[\\/])?' + $exe + ')'
$pattern = '(?i)(?:^|[;&|(])\s*(?:&\s*)?' + $interpreter + '\s+(?:-\S+\s+)*[''"]?(?:[^\s''"]*[\\/])?report\.py\b'
$command = [string]$data.tool_input.command
if ($command -notmatch $pattern) { exit 0 }

# 2. Only a successful run: report.py prints "Written:" only after every check passed
#    and the files were saved, and it never prints MISMATCH on success.
$result = $data.tool_response | ConvertTo-Json -Depth 5 -Compress
if ($result -notmatch 'Written:' -or $result -match 'MISMATCH') { exit 0 }

# --- Show the notification ---
$null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
$toast = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent('ToastText02')
$text = $toast.GetElementsByTagName('text')
$null = $text.Item(0).AppendChild($toast.CreateTextNode("Mama Tee's Bakery"))
$null = $text.Item(1).AppendChild($toast.CreateTextNode('Report ready in output/'))
$appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
[Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show([Windows.UI.Notifications.ToastNotification]::new($toast))
