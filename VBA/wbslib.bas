Attribute VB_Name = "wbslib"
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


' ■ データ行範囲を取得する (高速化版)
Public Function FindDataRangeRows(ws As Worksheet) As Variant

    Dim startCell As Range
    Dim endCell As Range
    Dim lngStartRow As Long
    Dim lngEndRow As Long

    ' 1. KEY列で "@" を持つ最初のセルを高速に検索
    On Error Resume Next
    Set startCell = ws.Columns(cfg.COL_KEY).Find(What:="@", LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    On Error GoTo 0

    If startCell Is Nothing Then
        ' "@" が見つからない場合は、開始行を 0 として処理を終える
        FindDataRangeRows = Array(0, 0)
        
        MsgBox "KEY列（" & utils.ConvertColNumberToLetter(cfg.COL_KEY) & "）の開始行マーカー「@」が見つかりません。" & vbCrLf & _
               "（KEY列が非表示となっている場合は表示状態にしてください）", vbExclamation, "通知"
        
        Exit Function
    Else
        lngStartRow = startCell.row + 1 ' 実際のデータ開始行は "@" の次の行
    End If

    ' 2. KEY列で "$" を持つ最初のセルを "@" の次の行から高速に検索
    If lngStartRow > 1 Then ' "@" が見つかった場合のみ検索
        On Error Resume Next
        Set endCell = ws.Columns(cfg.COL_KEY).Find(What:="$", LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True, After:=ws.Cells(lngStartRow - 1, cfg.COL_KEY))
        On Error GoTo 0
        
        If endCell Is Nothing Then
            ' "$" が見つからない場合は、最終行をシートの最終行とするか、特定の値にするか検討
            lngEndRow = ws.Cells(ws.Rows.Count, cfg.COL_KEY).End(xlUp).row - 1 ' "$" がなくても最後まで
            
            MsgBox "KEY列（" & utils.ConvertColNumberToLetter(cfg.COL_KEY) & "）の終了行マーカー「$」が見つかりません。" & vbCrLf & _
                   "（KEY列が非表示となっている場合は表示状態にしてください）", vbExclamation, "通知"
        Else
            lngEndRow = endCell.row - 1 ' 実際のデータ終了行は "$" の前の行
        End If
    Else
        ' "@" が最初の行にある場合などの処理
        lngEndRow = ws.Cells(ws.Rows.Count, cfg.COL_KEY).End(xlUp).row - 1
    End If

    ' 結果を配列で返す
    FindDataRangeRows = Array(lngStartRow, lngEndRow)

End Function


' ■ エラーチェックを実施
Public Sub CheckWbsErrors(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim blnHasError As Boolean
    Dim intErrorCount As Integer
    Dim varData As Variant
    Dim colWbsId As New Collection         ' key=行番号文字列、value=WbsId文字列
    Dim colParentWbsId As New Collection   ' key=WbsId文字列、 value=WbsIdの親階層文字列
    Dim colError1Count As New Collection   ' key=行番号文字列、value=エラー数1
    Dim colError1Message As New Collection ' key=行番号文字列、value=エラーメッセージ1
    Dim colError2Count As New Collection   ' key=行番号文字列、value=エラー数2
    Dim colError2Message As New Collection ' key=行番号文字列、value=エラーメッセージ2
    Dim colError3Count As New Collection   ' key=行番号文字列、value=エラー数3
    Dim colError3Message As New Collection ' key=行番号文字列、value=エラーメッセージ3
    ' 一時変数定義
    Dim r As Long, c As Long
    Dim tmpWbsId As String
    Dim tmpCellValue As Variant
    Dim tmpPreCell As Variant
    Dim tmpErrorCount As Variant
    Dim tmpErrorMessage As Variant
    Dim tmpDotPosition As Long
    Dim tmpParentWbsId As String
    Dim tmpRecordError As Boolean
    Dim tmpWbsIdCount As Long
    Dim tmpRowIdx As String
    Dim tmpCol As Long
    Dim tmpRow As Long
    Dim tmpCount As Long
    Dim tmpCell As Range
    
    ' 初期化
    blnHasError = False

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' ERR 列の E をクリア
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ERR), ws.Cells(lngEndRow, cfg.COL_ERR)).ClearContents
    With ws.Range(ws.Cells(lngStartRow, cfg.COL_ERR), ws.Cells(lngEndRow, cfg.COL_ERR))
        For Each tmpCell In .Cells
            If Not tmpCell.Comment Is Nothing Then
                tmpCell.Comment.Delete
            End If
        Next tmpCell
    End With
    
    ' 指定範囲のデータを一度に取得
    varData = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).value

    ' 配列をループしてチェック＆必要なデータを収集
    For r = 1 To UBound(varData, 1)
        ' 実際の行番号を作成
        tmpRow = r + lngStartRow - 1
        tmpRowIdx = "IDX" & tmpRow
        tmpRecordError = False
        tmpPreCell = ""
        tmpWbsId = ""
        For c = 1 To UBound(varData, 2)
            ' 実際のカラム番号を作成
            tmpCol = c + cfg.COL_OPT
            ' 現在のセルの値を取得
            tmpCellValue = varData(r, c)
            If c = 1 Then
                ' # L1 の場合 #
                If Not IsEmpty(tmpCellValue) And tmpCellValue <> "" Then
                    ' セルが空ではない場合、WbsId に文字列を追加
                    tmpWbsId = tmpWbsId & tmpCellValue
                End If
            ElseIf c = 6 Then
                ' # TASK の場合 #
                If Not IsEmpty(tmpCellValue) And tmpCellValue <> "" Then
                    ' セルが空ではない場合、WbsId に文字列を追加
                    tmpWbsId = tmpWbsId & ".T" & tmpCellValue
                End If
                ' ここまで来たら正常終了
                colError1Count.Add 0, tmpRowIdx
                colError1Message.Add "", tmpRowIdx
            Else
                ' # L2～L5 の場合 #
                If Not IsEmpty(tmpCellValue) And tmpCellValue <> "" Then
                    ' # 現在のセルが空ではない場合 #
                    If Not IsEmpty(tmpPreCell) And tmpPreCell <> "" Then
                        ' # 直前のセルが空ではない場合、WbsId に文字列を追加 #
                        tmpWbsId = tmpWbsId & "." & tmpCellValue
                    Else
                        ' # 直前のセルが空の場合、エラーとして処理 #
                        blnHasError = True
                        ' エラー件数およびメッセージに追加して、コレクションに再セット
                        colError1Count.Add 1, tmpRowIdx
                        colError1Message.Add "・階層番号に問題（" & utils.ConvertColNumberToLetter(tmpCol - 1) & tmpRow & "が数値ではない）" & vbCrLf, tmpRowIdx
                        ' エラー行としてカラムのループを終了
                        tmpRecordError = True
                        Exit For
                    End If
                End If
            End If
            tmpPreCell = tmpCellValue
        Next c
        ' レコードエラーが発生してない場合、WbsId や ParentWbsId をコレクションに追加
        If tmpRecordError = False Then
            ' WbsId をコレクションに追加
            colWbsId.Add tmpWbsId, tmpRowIdx
            ' WbsId の親階層を作成し、コレクションに追加
            tmpDotPosition = InStrRev(tmpWbsId, ".")
            If tmpDotPosition > 0 Then
                tmpParentWbsId = Left(tmpWbsId, tmpDotPosition - 1)
                On Error Resume Next
                colParentWbsId.Add tmpParentWbsId, tmpWbsId
                On Error GoTo 0
            End If
        End If
    Next r
    
    ' すべての行を調査
    For r = lngStartRow To lngEndRow
        tmpRowIdx = "IDX" & r
        ' ここまでのエラー件数を取得する
        tmpErrorCount = 0
        If utils.ExistsColKey(colError1Count, tmpRowIdx) = True Then
            tmpErrorCount = colError1Count.item(tmpRowIdx)
        End If
        ' まだエラーが発生していない行で、WbsId が登録されているもののみ検査
        If tmpErrorCount = 0 And utils.ExistsColKey(colWbsId, tmpRowIdx) Then
            tmpWbsId = colWbsId.item(tmpRowIdx)
            ' 空文字列となっているWbsIdを除外して、エラーチェックを行う
            If tmpWbsId <> "" Then
                ' WbsId の数を取得して、1件以上ならエラーとする
                If utils.ExistsColItemCount(colWbsId, tmpWbsId) > 1 Then
                    blnHasError = True
                    colError2Count.Add 1, tmpRowIdx
                    colError2Message.Add "・同一階層番号が存在（Row=" & r & "）" & vbCrLf, tmpRowIdx
                End If
                ' L1 以外（WbsId のドットがある）で、親階層WbsIdがない場合エラーとする
                tmpDotPosition = InStrRev(tmpWbsId, ".")
                If tmpDotPosition > 0 And utils.ExistsColKey(colParentWbsId, tmpWbsId) = True Then
                    tmpParentWbsId = colParentWbsId.item(tmpWbsId)
                    If utils.ExistsColItem(colWbsId, tmpParentWbsId) = False Then
                        blnHasError = True
                        colError3Count.Add 1, tmpRowIdx
                        colError3Message.Add "・親階層が存在しない（Row=" & r & "）" & vbCrLf, tmpRowIdx
                    End If
                End If
            End If
        End If
    Next r
    
    ' 集計したエラーで表示を作成
    If blnHasError = True Then
        For r = lngStartRow To lngEndRow
            tmpRowIdx = "IDX" & r
            tmpErrorCount = 0
            tmpErrorMessage = ""
            If utils.ExistsColKey(colError1Count, tmpRowIdx) = True Then
                tmpErrorCount = tmpErrorCount + colError1Count.item(tmpRowIdx)
                If utils.ExistsColKey(colError1Message, tmpRowIdx) = True Then
                    tmpErrorMessage = tmpErrorMessage & colError1Message.item(tmpRowIdx)
                End If
            End If
            If utils.ExistsColKey(colError2Count, tmpRowIdx) = True Then
                tmpErrorCount = tmpErrorCount + colError2Count.item(tmpRowIdx)
                If utils.ExistsColKey(colError2Message, tmpRowIdx) = True Then
                    tmpErrorMessage = tmpErrorMessage & colError2Message.item(tmpRowIdx)
                End If
            End If
            If utils.ExistsColKey(colError3Count, tmpRowIdx) = True Then
                tmpErrorCount = tmpErrorCount + colError3Count.item(tmpRowIdx)
                If utils.ExistsColKey(colError3Message, tmpRowIdx) = True Then
                    tmpErrorMessage = tmpErrorMessage & colError3Message.item(tmpRowIdx)
                End If
            End If
            If tmpErrorCount > 0 Then
                ws.Cells(r, lngCOL_ERR).value = "E"
                If ws.Cells(r, lngCOL_ERR).Comment Is Nothing Then
                    ws.Cells(r, lngCOL_ERR).AddComment
                End If
                ws.Cells(r, lngCOL_ERR).Comment.Text Text:=tmpErrorMessage
                intErrorCount = intErrorCount + tmpErrorCount
                ' コメントの幅と高さを手動で設定
                With ws.Cells(r, lngCOL_ERR).Comment.Shape
                    .Width = 300   ' 幅を 300 に設定
                    .Height = 100  ' 高さを 100 に設定
                End With
            End If
        Next r
    End If
    
    ' エラーがあればメッセージ表示
    If intErrorCount > 0 Then
        MsgBox intErrorCount & " 件の異常を検出しました。", vbExclamation, "エラーチェック"
    End If
    
End Sub


' ■ ソート用カラムに数式をセット
Public Sub SetFormulaForWbsIdx(ws As Worksheet)
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' 数式を作成
    strFormula = "=IF(" & cfg.COL_ERR_LABEL & lngStartRow & "=""E"",""ERROR""," & _
                    "IF(" & cfg.COL_L1_LABEL & lngStartRow & "="""",""XXX.XXX.XXX.XXX.XXX.XXX"", CONCAT(TEXT(" & cfg.COL_L1_LABEL & lngStartRow & ",""000"")," & _
                    "IF(" & cfg.COL_L2_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_L2_LABEL & lngStartRow & ",""000""))," & _
                    "IF(" & cfg.COL_L3_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_L3_LABEL & lngStartRow & ",""000""))," & _
                    "IF(" & cfg.COL_L4_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_L4_LABEL & lngStartRow & ",""000""))," & _
                    "IF(" & cfg.COL_L5_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_L5_LABEL & lngStartRow & ",""000""))," & _
                    "IF(" & cfg.COL_TASK_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_TASK_LABEL & lngStartRow & ",""000"")))))"

    ' 一括で対象範囲を取得
    With ws.Range(cfg.COL_WBS_IDX_LABEL & lngStartRow & ":" & cfg.COL_WBS_IDX_LABEL & lngEndRow)
        ' 数値書式を一括で設定
        .NumberFormat = "General"
        ' 数式をセット
        .Formula = strFormula
    End With
    
End Sub


' ■ WBS-IDX数カラムに数式をセット
Public Sub SetFormulaForWbsCnt(ws As Worksheet)
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' 数式を作成
    strFormula = "=COUNTIF(" & _
                    cfg.COL_WBS_IDX_LABEL & "$" & lngStartRow & ":" & _
                    cfg.COL_WBS_IDX_LABEL & "$" & lngEndRow & "," & _
                    cfg.COL_WBS_IDX_LABEL & lngStartRow & ")"

    ' 一括で対象範囲を取得
    With ws.Range(cfg.COL_WBS_CNT_LABEL & lngStartRow & ":" & cfg.COL_WBS_CNT_LABEL & lngEndRow)
        ' 数値書式を一括で設定
        .NumberFormat = "General"
        ' 数式をセット
        .Formula = strFormula
    End With
    
End Sub


' ■ ID表示用カラムに数式をセット
Public Sub SetFormulaForWbsId(ws As Worksheet)
   
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' 数式を作成
    strFormula = "=IF(" & cfg.COL_ERR_LABEL & lngStartRow & "=""E"",""ERROR""," & _
                    "IF(" & cfg.COL_L1_LABEL & lngStartRow & "="""","""",CONCAT(" & cfg.COL_L1_LABEL & lngStartRow & "," & _
                    "IF(" & cfg.COL_L2_LABEL & lngStartRow & "="""","""","".""&" & cfg.COL_L2_LABEL & lngStartRow & " ), " & _
                    "IF(" & cfg.COL_L3_LABEL & lngStartRow & "="""","""","".""&" & cfg.COL_L3_LABEL & lngStartRow & " ), " & _
                    "IF(" & cfg.COL_L4_LABEL & lngStartRow & "="""","""","".""&" & cfg.COL_L4_LABEL & lngStartRow & " ), " & _
                    "IF(" & cfg.COL_L5_LABEL & lngStartRow & "="""","""","".""&" & cfg.COL_L5_LABEL & lngStartRow & " ), " & _
                    "IF(" & cfg.COL_TASK_LABEL & lngStartRow & "="""","""","".T""&" & cfg.COL_TASK_LABEL & lngStartRow & " ))))"

    ' 一括で対象範囲を取得
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_WBS_ID_LABEL & lngEndRow)
        ' 数値書式を一括で設定
        .NumberFormat = "General"
        ' 数式をセット
        .Formula = strFormula
    End With
    
End Sub


' ■ レベルカラムに数式をセット
Public Sub SetFormulaForLevel(ws As Worksheet)
   
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' 数式を作成
    strFormula = "=IF(AND(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & ")," & cfg.COL_L2_LABEL & lngStartRow & "=""""," & cfg.COL_L3_LABEL & lngStartRow & "=""""," & _
                    cfg.COL_L4_LABEL & lngStartRow & "=""""," & cfg.COL_L5_LABEL & lngStartRow & "=""""),1," & _
                    "IF(AND(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & "),ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & ")," & cfg.COL_L3_LABEL & lngStartRow & "=""""," & _
                    cfg.COL_L4_LABEL & lngStartRow & "=""""," & cfg.COL_L5_LABEL & lngStartRow & "=""""),2," & _
                    "IF(AND(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & "),ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & "),ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & ")," & _
                    cfg.COL_L4_LABEL & lngStartRow & "=""""," & cfg.COL_L5_LABEL & lngStartRow & "=""""),3," & _
                    "IF(AND(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & "),ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & "),ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & "),ISNUMBER(" & _
                    cfg.COL_L4_LABEL & lngStartRow & ")," & cfg.COL_L5_LABEL & lngStartRow & "=""""),4," & _
                    "IF(AND(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & "),ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & "),ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & "),ISNUMBER(" & _
                    cfg.COL_L4_LABEL & lngStartRow & "),ISNUMBER(" & cfg.COL_L5_LABEL & lngStartRow & ")),5,0)))))"

    ' 一括で対象範囲を取得
    With ws.Range(cfg.COL_LEVEL_LABEL & lngStartRow & ":" & cfg.COL_LEVEL_LABEL & lngEndRow)
        ' 数値書式を一括で設定
        .NumberFormat = "General"
        ' 数式をセット
        .Formula = strFormula
    End With
    
End Sub


' ■ フラグTカラムに数式をセット
Public Sub SetFormulaForFlgT(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' 数式を作成
    strFormula = "=IF(AND(" & cfg.COL_TASK_LABEL & lngStartRow & "<>"""",ISNUMBER(" & cfg.COL_TASK_LABEL & lngStartRow & ")),TRUE,FALSE)"

    ' 一括で対象範囲を取得
    With ws.Range(cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow)
        ' 数値書式を一括で設定
        .NumberFormat = "General"
        ' 数式をセット
        .Formula = strFormula
    End With

End Sub


' ■ フラグICカラムに数式をセット
Public Sub SetFormulaForFlgIC(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' 数式を作成
    strFormula = "=NOT(OR(" & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""移管済""," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""棚上げ""," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""却下""" & "))"

    ' 一括で対象範囲を取得
    With ws.Range(cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow)
        ' 数値書式を一括で設定
        .NumberFormat = "General"
        ' 数式をセット
        .Formula = strFormula
    End With

End Sub


' ■ フラグPEカラムに数式をセット
Public Sub SetFormulaForFlgPE(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' 数式を作成
    strFormula = "=AND(" & _
                    cfg.COL_LEVEL_LABEL & lngStartRow & ">0," & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & "<>"""",IFERROR(ISNUMBER(MATCH(IFERROR(LEFT(" & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & ",FIND(""~"",SUBSTITUTE(" & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & ",""."",""~"",LEN(" & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & ")-LEN(SUBSTITUTE(" & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & ",""."",""""))))-1)," & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & ")," & _
                    cfg.COL_WBS_ID_LABEL & "$" & lngStartRow & ":" & cfg.COL_WBS_ID_LABEL & "$" & lngEndRow & _
                    ",0)),FALSE))"

    ' 一括で対象範囲を取得
    With ws.Range(cfg.COL_FLG_PE_LABEL & lngStartRow & ":" & cfg.COL_FLG_PE_LABEL & lngEndRow)
        ' 数値書式を一括で設定
        .NumberFormat = "General"
        ' 数式をセット
        .Formula = strFormula
    End With

End Sub


' ■ フラグCEカラムに数式をセット
Public Sub SetFormulaForFlgCE(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' 数式を作成
    strFormula = "=AND(" & _
                    cfg.COL_LEVEL_LABEL & lngStartRow & ">0," & _
                    cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE," & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & "<>"""",IFERROR(SUMPRODUCT(--(LEFT(" & _
                    cfg.COL_WBS_ID_LABEL & "$" & lngStartRow & ":" & cfg.COL_WBS_ID_LABEL & "$" & lngEndRow & ",LEN(" & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & "&"".""))=" & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & "&"".""))>0,FALSE))"

    ' 一括で対象範囲を取得
    With ws.Range(cfg.COL_FLG_CE_LABEL & lngStartRow & ":" & cfg.COL_FLG_CE_LABEL & lngEndRow)
        ' 数値書式を一括で設定
        .NumberFormat = "General"
        ' 数式をセット
        .Formula = strFormula
    End With

End Sub


' ■ Exe1ボタンクリック時に実行される処理
Public Sub Exe1ButtonClick()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    ' 変数定義
    Dim lngSelectedIndex As Long
    Dim shpExe1ComboBox As Shape
    
    ' 実行コンボボックスを取得
    On Error Resume Next
    Set shpExe1ComboBox = ws.Shapes(cfg.NAME_EXE1_COMBOBOX)
    On Error GoTo 0
    
    ' 実行コンボボックスが存在しない場合、終了
    If shpExe1ComboBox Is Nothing Then
        Exit Sub
    End If
    
    ' 選択中のインデックスを取得
    lngSelectedIndex = shpExe1ComboBox.ControlFormat.ListIndex

    ' インデックスに対応する処理を実行
    Select Case lngSelectedIndex
        Case 1
            MsgBox "１つ目を選択（" & ws.Name & "）"
        Case 2
            MsgBox "２つ目を選択（" & ws.Name & "）"
        Case 3
            MsgBox "３つ目を選択（" & ws.Name & "）"
        Case 4
            MsgBox "４つ目を選択（" & ws.Name & "）"
        Case 5
            MsgBox "５つ目を選択（" & ws.Name & "）"
        Case 6
            MsgBox "６つ目を選択（" & ws.Name & "）"
        Case 7
            MsgBox "７つ目を選択（" & ws.Name & "）"
        Case 8
            MsgBox "８つ目を選択（" & ws.Name & "）"
        Case 9
            MsgBox "９つ目を選択（" & ws.Name & "）"
        Case Else
            MsgBox "項目が選択されていません。"
    End Select

End Sub


' ■ Reset1ボタンクリック時に実行される処理
Public Sub Reset1ButtonClick()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    ' 変数定義
    Dim lngSelectedIndex As Long
    Dim shpExe1ComboBox As Shape
    
    ' 実行1コンボボックスを取得
    On Error Resume Next
    Set shpExe1ComboBox = ws.Shapes(cfg.NAME_EXE1_COMBOBOX)
    On Error GoTo 0
    
    ' 実行1コンボボックスが存在しない場合、終了
    If shpExe1ComboBox Is Nothing Then
        Exit Sub
    End If
    
    ' 実行1コンボボックスのリストインデックスを先頭に
    With ws.DropDowns(cfg.NAME_EXE1_COMBOBOX)
        .ListIndex = 1
    End With
    
End Sub


' ■ Exe2ボタンクリック時に実行される処理
Public Sub Exe2ButtonClick()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    ' 変数定義
    Dim lngSelectedIndex As Long
    Dim shpExe2ComboBox As Shape
    
    ' 実行コンボボックスを取得
    On Error Resume Next
    Set shpExe2ComboBox = ws.Shapes(cfg.NAME_EXE2_COMBOBOX)
    On Error GoTo 0
    
    ' 実行コンボボックスが存在しない場合、終了
    If shpExe2ComboBox Is Nothing Then
        Exit Sub
    End If
    
    ' 選択中のインデックスを取得
    lngSelectedIndex = shpExe2ComboBox.ControlFormat.ListIndex

    ' インデックスに対応する処理を実行
    Select Case lngSelectedIndex
        Case 1
            MsgBox "１つ目を選択（" & ws.Name & "）"
        Case 2
            MsgBox "２つ目を選択（" & ws.Name & "）"
        Case 3
            MsgBox "３つ目を選択（" & ws.Name & "）"
        Case 4
            MsgBox "４つ目を選択（" & ws.Name & "）"
        Case 5
            MsgBox "５つ目を選択（" & ws.Name & "）"
        Case Else
            MsgBox "項目が選択されていません。"
    End Select

End Sub


' ■ Reset2ボタンクリック時に実行される処理
Public Sub Reset2ButtonClick()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    ' 変数定義
    Dim lngSelectedIndex As Long
    Dim shpExe2ComboBox As Shape
    
    ' 実行2コンボボックスを取得
    On Error Resume Next
    Set shpExe2ComboBox = ws.Shapes(cfg.NAME_EXE2_COMBOBOX)
    On Error GoTo 0
    
    ' 実行2コンボボックスが存在しない場合、終了
    If shpExe2ComboBox Is Nothing Then
        Exit Sub
    End If
    
    ' 実行2コンボボックスのリストインデックスを先頭に
    With ws.DropDowns(cfg.NAME_EXE2_COMBOBOX)
        .ListIndex = 1
    End With
    
End Sub

