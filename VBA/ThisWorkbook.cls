VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' ■ 関数の存在チェック
Function FunctionExists(sheetCodeName As String, funcName As String) As Boolean
    
    ' 変数定義
    Dim exists As Boolean
    
    ' テストモードで呼び出してみることで存在を確認
    On Error Resume Next
    Dim test As Variant
    test = Application.Run(sheetCodeName & "." & funcName, True)
    exists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    
    ' 結果をセット
    FunctionExists = exists

End Function


' ■ ワークブック保存時に実行されるサブルーチン
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    ' 変数定義
    Dim ws As Worksheet           ' ワークシート
    Dim current_ws As Worksheet   ' アクティブなワークシート
    Dim hasError As Boolean       ' エラー有無
    Dim result As VbMsgBoxResult  ' メッセージボックス選択結果
    Dim functionResult As Variant ' 関数戻り値
    Dim sheetCodeName As String   ' ワークシートのコードネーム

    Set current_ws = ActiveSheet

    hasError = False

    ' すべてのシートをループ
    For Each ws In ThisWorkbook.Sheets
        ' シート名が "WBS" で始まるか確認
        If Left(ws.Name, 3) = "WBS" And ws.CodeName = current_ws.CodeName Then
        
            ' シートの CodeName を取得
            sheetCodeName = ws.CodeName
            
            ' PrepareForProcessing が存在するか確認して実行（引数 False）
            If FunctionExists(sheetCodeName, "PrepareForProcessing") Then
            
                Application.ScreenUpdating = False
                Application.Calculation = xlCalculationManual
                Application.EnableEvents = False
            
                ' CallByName を使ってメソッドを実行
                CallByName ws, "PrepareForProcessing", VbMethod, False
               
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                Application.EnableEvents = True
            End If
                        
            ' HasErrors が存在するか確認して実行（引数 False）
            If FunctionExists(sheetCodeName, "HasErrors") Then
                ' CallByName を使って関数を実行し、戻り値を取得
                functionResult = CallByName(ws, "HasErrors", VbMethod, False)
                ' 関数が正しく実行され、戻り値が取得できた場合
                If functionResult Then
                    hasError = True
                End If
            End If
            
        End If
    Next ws

    ' エラーがある場合は確認メッセージを出す
    If hasError Then
        ' 保存するか確認
        result = MsgBox("エラーがありますが保存しますか？", vbExclamation + vbYesNo, "エラーチェック")

        ' 「No」を押したら保存をキャンセル
        If result = vbNo Then
            Cancel = True
        End If
    End If
    
End Sub

