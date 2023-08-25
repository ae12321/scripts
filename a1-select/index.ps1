$_today = Get-Date

Add-Type -assemblyName System.Windows.Forms
Add-Type -AssemblyName "Microsoft.VisualBasic"

function main([string]$filePath, [int]$zoomSize) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    $book = $excel.Workbooks.Open($filePath)

    # 1�V�[�g�ɑ΂��鏈��
    foreach ($worksheet in $excel.Worksheets) {
        # ��\���V�[�g�͏����ł��Ȃ�
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

    # 1�V�[�g�ڑI��
    $excel.Worksheets.Item(1).Activate()

    # �㏑���ۑ�
    $book.save()
    $book.close()
}

try {
    # �t�@�C�������I��
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "�G�N�Z��|*.xls;*.xlsx;*.xlsm|All Files|*.*"
    $dialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $dialog.Multiselect = $true
    if (-not($dialog.ShowDialog() -eq "OK")) {
        return
    }

    # �{���w��
    $zoomSize = [Microsoft.VisualBasic.Interaction]::InputBox("�{����ύX���Ȃ��F-1`n`n�{����ύX����F50�`200", "�{��", 100)

    # �t�@�C���X�ɑ΂��鏈��
    foreach ($filePath in $dialog.FileNames) {
        main $filePath $zoomSize
    }
} catch {
    $_
} finally {
    # ���
    Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable
    # GC������
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    # �����ϐ����N���[���A�b�v
    1 | % { $_ } > $null
    [System.GC]::Collect()
}
