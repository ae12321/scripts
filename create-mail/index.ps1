$_today = Get-Date

# address setting
$to = "aaa@gmail.com; bbb@gmail.com"
$cc = "ccc@gmail.com; ddd@gmail.com"

# subject setting
$_template = "start: {0}({1})"
$_ymd = "{0:yyyy/M/d}" -f $_today
$_wday = "{0:ddd}" -f $_today
$subject = $_template -f $_ymd, $_wday

# body setting
$_ymd2 = "{0:yyyy/MM/dd}" -f $_today
$body = "〇〇様

本日({0})、お休みします
" -f $_ymd2

try {
    $outlook = New-Object -ComObject Outlook.Application
    $Mail = $outlook.CreateItem(0)

    $Mail.To = $To
    $Mail.CC = $CC

    $Mail.Subject = $Subject

    $Mail.Body = $Body
    $Mail.BodyFormat = 1

    $inspector = $Mail.GetInspector
    $inspector.Display()
} catch {
    $_
} finally {
    # 解放
    Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable
    # GCを強制
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    # 自動変数をクリーンアップ
    1 | % { $_ } > $null
    [System.GC]::Collect()
}
