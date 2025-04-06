# export_vba.ps1
param (
    [string]$xlsmFile,
    [string]$exportFolder
)

# フォルダの存在チェック
if (-not (Test-Path $exportFolder)) {
    New-Item -ItemType Directory -Path $exportFolder
}

# Excelアプリケーションを作成
$excelApp = New-Object -ComObject Excel.Application
$excelApp.Visible = $false
$excelApp.DisplayAlerts = $false

# Excelファイルを開く
$workbook = $excelApp.Workbooks.Open($xlsmFile)

# VBA コンポーネントをエクスポート
foreach ($component in $workbook.VBProject.VBComponents) {
    $exportPath = Join-Path $exportFolder $component.Name
    
    # モジュール、クラスモジュール、フォームに分けてエクスポート
    switch ($component.Type) {
        1 { # 標準モジュール
            $exportPath += ".bas"
        }
        2 { # クラスモジュール（ThisWorkbook, Sheet1 など）
            $exportPath += ".cls"
        }
        3 { # フォーム
            $exportPath += ".frm"
        }
        default {
            $exportPath += ".txt"
        }
    }

    # エクスポート
    $component.Export($exportPath)
}

# ファイルを閉じてExcelを終了
$workbook.Close($false)
$excelApp.Quit()

# メモリ解放
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
