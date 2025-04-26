VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' ------------------------------------------------------------------------------
' Copyright 2025 Hiroki Chiba <h16k148@gmail.com>
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
' ------------------------------------------------------------------------------


' ■ 関数の存在チェック
Function FunctionExists(strSheetCodeName As String, strFuncName As String) As Boolean
    
    ' 変数定義
    Dim blnExists As Boolean
    
    ' テストモードで呼び出してみることで存在を確認
    On Error Resume Next
    Dim test As Variant
    test = Application.Run(strSheetCodeName & "." & strFuncName)
    blnExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    
    ' 結果をセット
    FunctionExists = blnExists

End Function


' ■ ワークブック保存時に実行されるサブルーチン
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    ' 変数定義
    Dim ws As Worksheet              ' ワークシート
    Dim wsCurrent As Worksheet       ' アクティブなワークシート
    Dim strSheetCodeName As String   ' ワークシートのコードネーム

    ' アクティブなシートを取得
    Set wsCurrent = ActiveSheet

    ' すべてのシートをループ
    For Each ws In ThisWorkbook.Sheets
    
        ' シート名が "WBS" で始まるか確認
        If Left(ws.Name, 3) = "WBS" And ws.CodeName = wsCurrent.CodeName Then
            ' シートの CodeName を取得
            strSheetCodeName = ws.CodeName
            ' IsWBS が存在するか確認してから実行
            If FunctionExists(strSheetCodeName, "IsWBS") Then
                wbsui.ExecBeforeSave ws
            End If
        End If
    
    Next ws
    
End Sub

