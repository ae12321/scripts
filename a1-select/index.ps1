$_today = Get-Date

Add-Type -assemblyName System.Windows.Forms
Add-Type -AssemblyName "Microsoft.VisualBasic"

function main([string]$filePath, [int]$zoomSize) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    $book = $excel.Workbooks.Open($filePath)

    # 1シートに対する処理
    foreach ($worksheet in $excel.Worksheets) {
        # 非表示シートは処理できない
        if ($worksheet.Visible -eq $false) {
            continue
        }
        
        # A1 select
        $worksheet.Activate()
        $worksheet.Cells.Item(1,1).Select() | Out-Null

        # zoom size
        if ($zoomSize -gt 0) {
            $excel.ActiveWindow.Zoom = $zoomSize
        }
    }

    # 1シート目選択
    $excel.Worksheets.Item(1).Activate()

    # 上書き保存
    $book.save()
    $book.close()
}

try {
    # ファイル複数選択
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "エクセル|*.xls;*.xlsx;*.xlsm|All Files|*.*"
    $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $dialog.Multiselect = $true
    if (-not($dialog.ShowDialog() -eq "OK")) {
        return
    }

    # 倍率指定
    $zoomSize = [Microsoft.VisualBasic.Interaction]::InputBox("倍率を変更しない：-1`n`n倍率を変更する：50～200", "倍率", 100)

    # ファイル個々に対する処理
    foreach ($filePath in $dialog.FileNames) {
        main $filePath $zoomSize
    }
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
