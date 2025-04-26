VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
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


' ◆ WBSシートであることを示す関数を定義
Public Function IsWBS() As Boolean
    IsWBS = True
End Function


' ◆ CHK と OPT のダブルクリックイベントを実装
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' 変数定義
    Dim lngClickedColumn As Long

    ' 列番号を取得
    lngClickedColumn = Target.Column

    ' ガード条件（対象列以外はデフォルト動作のままで終了）
    If lngClickedColumn <> cfg.COL_CHK And lngClickedColumn <> cfg.COL_OPT Then
        Exit Sub
    End If

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim lngClickedRow As Long
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varClicked As Variant
    ' 一時変数定義
    Dim rngFoundCell As Range

    ' 行番号を取得
    lngClickedRow = Target.row

    ' 開始行と終了行を取得
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' ガード条件（行番号が指定範囲外の場合は終了）
    If lngClickedRow < lngStartRow Or lngClickedRow > lngEndRow Then
        Exit Sub
    End If

    ' CHK 処理
    If lngClickedColumn = cfg.COL_CHK Then
        varClicked = ws.Cells(lngClickedRow, cfg.COL_CHK).value

        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False

        If varClicked = cfg.CHK_MARK_T Then
            ws.Cells(lngClickedRow, cfg.COL_CHK).value = cfg.CHK_MARK_F
        Else
            ws.Cells(lngClickedRow, cfg.COL_CHK).value = cfg.CHK_MARK_T
        End If

        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True

        Cancel = True
        Exit Sub
    End If

    ' OPT 処理 (高速化版)
    If lngClickedColumn = cfg.COL_OPT Then
        ' クリックした値を取得
        varClicked = ws.Cells(lngClickedRow, cfg.COL_OPT).value
        If varClicked <> cfg.OPT_MARK_T Then ' クリックされたセルが cfg.OPT_MARK_T でない場合のみ処理
            ' lngStartRow から lngEndRow の範囲で cfg.OPT_MARK_T を持つ最初のセルを検索
            On Error Resume Next
            Set rngFoundCell = ws.Range(ws.Cells(lngStartRow, cfg.COL_OPT), ws.Cells(lngEndRow, cfg.COL_OPT)).Find(What:=cfg.OPT_MARK_T, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
            On Error GoTo 0

            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False

            ' cfg.OPT_MARK_T を持つセルが見つかったら cfg.OPT_MARK_F に変更
            If Not rngFoundCell Is Nothing Then
                rngFoundCell.value = cfg.OPT_MARK_F
            End If

            ' クリックされたセルの値を cfg.OPT_MARK_T に変更
            ws.Cells(lngClickedRow, cfg.COL_OPT).value = cfg.OPT_MARK_T
            ws.Cells(lngStartRow - 1, cfg.COL_OPT).value = lngClickedRow

            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True

            Cancel = True
            Exit Sub
        End If
    End If

End Sub

