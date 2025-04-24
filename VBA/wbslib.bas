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
Public Sub ExecCheckWbsErrors(ws As Worksheet)

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
                ws.Cells(r, cfg.COL_ERR).value = "E"
                If ws.Cells(r, cfg.COL_ERR).Comment Is Nothing Then
                    ws.Cells(r, cfg.COL_ERR).AddComment
                End If
                ws.Cells(r, cfg.COL_ERR).Comment.Text Text:=tmpErrorMessage
                intErrorCount = intErrorCount + tmpErrorCount
                ' コメントの幅と高さを手動で設定
                With ws.Cells(r, cfg.COL_ERR).Comment.Shape
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


' ■ データ範囲をソートする
Public Sub ExecSortWbsRange(ws As Worksheet)

    ' 変数定義
    Dim rngSortTarget As Range
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

   ' エラー列～最終列の範囲を指定（startRow～endRow）
    Set rngSortTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_ERR), ws.Cells(lngEndRow, cfg.COL_LAST))

    ' WBSインデックス列をキーとして昇順にソート
    rngSortTarget.Sort Key1:=ws.Range(cfg.COL_WBS_IDX_LABEL & lngStartRow), Order1:=xlAscending, Header:=xlNo

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
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_DELETED & """," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_TRANSFERRED & """," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_SHELVED & """," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_REJECTED & """" & "))"

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


' ■ 予定工数を集計する式をセット
Public Sub SetFormulaForPlannedEffort(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 現在のインデックスを取得
        i = r - lngStartRow + 1
        ' 現在のWBSレベルセルの値を取得
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' 現在のWBSタスクセルの値を取得
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = False Then
            ' # 行がタスク以外の場合 #
            If tmpVarLevelCell = 5 Then
                ' # 行がL5階層の場合 #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # 行がL4階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # 行がL3階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # 行がL2階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # 行がL1階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
    ' L1集計数式をセット
    tmpStrBoolArrayH = "(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "))*" & _
              "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
              "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
    tmpStrFormula = "=SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngEndRow + 2).Formula = tmpStrFormula

End Sub


' ■ 実績済工数を集計する式をセット
Public Sub SetFormulaForActualCompletedEffort(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 現在のインデックスを取得
        i = r - lngStartRow + 1
        ' 現在のWBSレベルセルの値を取得
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' 現在のWBSタスクセルの値を取得
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = False Then
            ' # 行がタスク以外の場合 #
            If tmpVarLevelCell = 5 Then
                ' # 行がL5階層の場合 #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # 行がL4階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # 行がL3階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # 行がL2階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # 行がL1階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
    ' L1集計数式をセット
    tmpStrBoolArrayH = "(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "))*" & _
              "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
              "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
    tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
    ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow + 2).Formula = tmpStrFormula

End Sub


' ■ 実績残工数を集計する式をセット
Public Sub SetFormulaForActualRemainingEffort(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 現在のインデックスを取得
        i = r - lngStartRow + 1
        ' 現在のWBSレベルセルの値を取得
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' 現在のWBSタスクセルの値を取得
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = False Then
            ' # 行がタスク以外の場合 #
            If tmpVarLevelCell = 5 Then
                ' # 行がL5階層の場合 #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # 行がL4階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # 行がL3階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # 行がL2階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # 行がL1階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
    ' L1集計数式をセット
    tmpStrBoolArrayH = "(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "))*" & _
              "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
              "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
    tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow + 2).Formula = tmpStrFormula

End Sub


' ■ タスク進捗率を集計する式をセット
Public Sub SetFormulaForTaskProgressRate(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String
    Dim tmpStrSumWeightH As String, tmpStrSumWeightT As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 現在のインデックスを取得
        i = r - lngStartRow + 1
        ' 現在のWBSレベルセルの値を取得
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' 現在のWBSタスクセルの値を取得
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = True Then
            ' # 行がタスクの場合 #
            ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "0.0%"
        Else
            ' # 行がタスク以外の場合 #
            If tmpVarLevelCell = 5 Then
                ' # 行がL5階層の場合 #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrSumWeightT = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
                          "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
                          "," & tmpStrBoolArrayT & ",0))" & _
                          "/IF(" & tmpStrSumWeightT & "=0,1," & tmpStrSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # 行がL4階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrSumWeightT = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                tmpStrSumWeightH = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
                tmpStrFormula = "=(SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
                          "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
                          "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
                          "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
                          "," & tmpStrBoolArrayT & ",0)))" & _
                          "/IF(" & tmpStrSumWeightH & "+" & tmpStrSumWeightT & "=0,1," & tmpStrSumWeightH & "+" & tmpStrSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # 行がL3階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrSumWeightT = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                tmpStrSumWeightH = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
                tmpStrFormula = "=(SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
                          "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
                          "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
                          "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
                          "," & tmpStrBoolArrayT & ",0)))" & _
                          "/IF(" & tmpStrSumWeightH & "+" & tmpStrSumWeightT & "=0,1," & tmpStrSumWeightH & "+" & tmpStrSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # 行がL2階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrSumWeightT = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                tmpStrSumWeightH = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
                tmpStrFormula = "=(SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
                          "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
                          "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
                          "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
                          "," & tmpStrBoolArrayT & ",0)))" & _
                          "/IF(" & tmpStrSumWeightH & "+" & tmpStrSumWeightT & "=0,1," & tmpStrSumWeightH & "+" & tmpStrSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # 行がL1階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrSumWeightT = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                tmpStrSumWeightH = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
                tmpStrFormula = "=(SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
                          "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
                          "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
                          "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
                          "," & tmpStrBoolArrayT & ",0)))" & _
                          "/IF(" & tmpStrSumWeightH & "+" & tmpStrSumWeightT & "=0,1," & tmpStrSumWeightH & "+" & tmpStrSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
    ' L1集計数式をセット
    tmpStrBoolArrayH = "(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "))*" & _
              "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
              "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
    tmpStrSumWeightH = "SUM(FILTER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
    tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow & _
              "*(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow & ")" & _
              "," & tmpStrBoolArrayH & ",0))" & _
              "/IF(" & tmpStrSumWeightH & "=0,1," & tmpStrSumWeightH & ")"
    ws.Range(cfg.COL_TASK_PROG_LABEL & lngEndRow + 2).Formula = tmpStrFormula

End Sub


' ■ 工数進捗率を集計する式をセット
Public Sub SetFormulaForEffortProgressRate(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String
    Dim tmpStrCountH As String, tmpStrCountT As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 現在のインデックスを取得
        i = r - lngStartRow + 1
        ' 現在のWBSレベルセルの値を取得
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' 現在のWBSタスクセルの値を取得
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = True Then
            ' # 行がタスクの場合 #
            tmpStrFormula = "=" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r & _
            "/IF(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r & "+" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r & "=0," & _
            "1," & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r & "+" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r & ")"
            ' 指定された列のセルに数式をセット
            ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
        Else
            ' # 行がタスク以外の場合 #
            If tmpVarLevelCell = 5 Then
                ' # 行がL5階層の場合 #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrCountT = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ")),0)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
                          "," & tmpStrBoolArrayT & ",0))" & _
                          "/IF(" & tmpStrCountT & "=0,1," & tmpStrCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # 行がL4階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrCountH = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ")),0)"
                tmpStrCountT = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ")),0)"
                tmpStrFormula = "=(SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
                          "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
                          "," & tmpStrBoolArrayT & ",0)))" & _
                          "/IF(" & tmpStrCountH & "+" & tmpStrCountT & "=0,1," & tmpStrCountH & "+" & tmpStrCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # 行がL3階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrCountH = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ")),0)"
                tmpStrCountT = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ")),0)"
                tmpStrFormula = "=(SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
                          "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
                          "," & tmpStrBoolArrayT & ",0)))" & _
                          "/IF(" & tmpStrCountH & "+" & tmpStrCountT & "=0,1," & tmpStrCountH & "+" & tmpStrCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # 行がL2階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrCountH = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ")),0)"
                tmpStrCountT = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ")),0)"
                tmpStrFormula = "=(SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
                          "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
                          "," & tmpStrBoolArrayT & ",0)))" & _
                          "/IF(" & tmpStrCountH & "+" & tmpStrCountT & "=0,1," & tmpStrCountH & "+" & tmpStrCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # 行がL1階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrCountH = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ")),0)"
                tmpStrCountT = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ")),0)"
                tmpStrFormula = "=(SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
                          "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
                          "," & tmpStrBoolArrayT & ",0)))" & _
                          "/IF(" & tmpStrCountH & "+" & tmpStrCountT & "=0,1," & tmpStrCountH & "+" & tmpStrCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
    ' L1集計数式をセット
    tmpStrBoolArrayH = "(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "))*" & _
              "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
              "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
    tmpStrCountH = "IFERROR(COUNT(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ")),0)"
    tmpStrFormula = "=SUM(FILTER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow & _
              "," & tmpStrBoolArrayH & ",0))" & _
              "/IF(" & tmpStrCountH & "=0,1," & tmpStrCountH & ")"
    ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2).Formula = tmpStrFormula

End Sub


' ■ タスク合計件数を集計する式をセット
Public Sub SetFormulaForTaskCount(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 現在のインデックスを取得
        i = r - lngStartRow + 1
        ' 現在のWBSレベルセルの値を取得
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' 現在のWBSタスクセルの値を取得
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = True Then
            ' # 行がタスクの場合 #
            ws.Range(cfg.COL_TASK_COUNT_LABEL & r).value = 1
        Else
            ' # 行がタスク以外の場合 #
            If tmpVarLevelCell = 5 Then
                ' # 行がL5階層の場合 #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # 行がL4階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # 行がL3階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # 行がL2階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # 行がL1階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
    ' L1集計数式をセット
    tmpStrBoolArrayH = "(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "))*" & _
              "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
              "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
    tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
    ws.Range(cfg.COL_TASK_COUNT_LABEL & lngEndRow + 2).Formula = tmpStrFormula

End Sub


' ■ タスク完了件数を集計する式をセット
Public Sub SetFormulaForTaskCompCount(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 現在のインデックスを取得
        i = r - lngStartRow + 1
        ' 現在のWBSレベルセルの値を取得
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' 現在のWBSタスクセルの値を取得
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = True Then
            ' # 行がタスクの場合 #
            tmpStrFormula = "=IF(" & cfg.COL_WBS_STATUS_LABEL & r & "=""" & cfg.WBS_STATUS_COMPLETED & """,1,0)"
            ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
        Else
            ' # 行がタスク以外の場合 #
            If tmpVarLevelCell = 5 Then
                ' # 行がL5階層の場合 #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # 行がL4階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # 行がL3階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # 行がL2階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # 行がL1階層の場合 #
                tmpStrBoolArrayH = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(ISNUMBER(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "))*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
    ' L1集計数式をセット
    tmpStrBoolArrayH = "(ISNUMBER(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "))*" & _
              "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "="""")*" & _
              "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=FALSE)*" & _
              "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
    tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayH & ",0))"
    ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow + 2).Formula = tmpStrFormula

End Sub


' ■ 選択中のオプションボタンから行番号を取得
Private Function GetCheckedOptSingleRow(ws As Worksheet) As Long
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim rngFoundCell As Range
    ' 一時変数定義
    Dim r As Long

    ' 開始行と終了行を取得
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then
        GetCheckedOptSingleRow = 0
        Exit Function
    End If
    
    ' lngStartRow から lngEndRow の範囲で cfg.OPT_MARK_T を持つ最初のセルを検索
    On Error Resume Next
    Set rngFoundCell = ws.Range(ws.Cells(lngStartRow, cfg.COL_OPT), ws.Cells(lngEndRow, cfg.COL_OPT)).Find( _
        What:=cfg.OPT_MARK_T, _
        LookAt:=xlWhole, _
        LookIn:=xlValues, _
        MatchCase:=True _
    )
    On Error GoTo 0
    
    ' セルが見つかった場合
    If Not rngFoundCell Is Nothing Then
        GetCheckedOptSingleRow = rngFoundCell.row
        Exit Function
    End If

    ' ここまで来たらチェックなし
    GetCheckedOptSingleRow = 0
End Function


' ■ 選択中のチェックボックスから行番号コレクションを取得
Private Function GetCheckedChkMultpleRows(ws As Worksheet) As Collection

    ' 変数定義
    Dim rowCollection As New Collection
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varData As Variant
    ' 一時変数定義
    Dim r As Long

    ' 開始行と終了行を取得
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then
        Set GetCheckedChkMultpleRows = rowCollection
        Exit Function
    End If

    ' 該当範囲のセルデータを一括で配列に格納
    varData = ws.Range(ws.Cells(lngStartRow, cfg.COL_CHK), ws.Cells(lngEndRow, cfg.COL_CHK)).value
    
    ' 配列をループして一致する行番号を収集
    For r = 1 To UBound(varData, 1)  ' 配列の行数分だけループ
        If varData(r, 1) = cfg.CHK_MARK_T Then
            rowCollection.Add lngStartRow + r - 1 ' 実際の行番号を追加
        End If
    Next r

    ' 結果として行番号のコレクションを返す
    Set GetCheckedChkMultpleRows = rowCollection
End Function


' ■ 選択した行の下に一行追加
Public Sub ExecInsertRowBelowSelection(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long
    Dim lngSelectedRow As Long
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    
    ' 行を追加
    lngSelectedRow = GetCheckedOptSingleRow(ws)
    If lngSelectedRow <> 0 Then
        ' 行を追加
        ws.Rows(lngSelectedRow + 1).Insert Shift:=xlDown
    Else
        MsgBox "選択してください（OPT)。", vbExclamation, "通知"
    End If

End Sub


' ■ 選択行の最終レベルIDをインクリメント
Public Sub ExecIncrementSelectedLastLevel(ws As Worksheet)

    ' 変数定義
    Dim lngSelectedRow As Long, intSelectedRowLevel As Integer, blnSelectedRowIsTask As Boolean
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colTargetIdx As New Collection
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpRngTarget As Range
    Dim tmpVarTargetArray As Variant
    Dim tmpVarLevelArray As Variant
    Dim tmpVarTaskArray As Variant
    Dim tmpLngSelectedRowL1 As Long, tmpLngSelectedRowL2 As Long, tmpLngSelectedRowL3 As Long, tmpLngSelectedRowL4 As Long, tmpLngSelectedRowL5 As Long, tmpLngSelectedRowTask As Long
    Dim tmpVarIdx As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 選択した行の番号を取得
    lngSelectedRow = GetCheckedOptSingleRow(ws)
    
    ' ガード条件（未選択の場合は、メッセージを出して終了）
    If lngSelectedRow = 0 Then
        MsgBox "選択してください（OPT)。", vbExclamation, "通知"
        Exit Sub
    End If
    
    ' あらかじめ更新対象範囲列のデータを取得
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    tmpVarTargetArray = tmpRngTarget.value
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' 選択した行のレベルを取得
    intSelectedRowLevel = tmpVarLevelArray(lngSelectedRow - lngStartRow + 1, 1)
    ' 選択した行がタスクかどうか取得
    blnSelectedRowIsTask = tmpVarTaskArray(lngSelectedRow - lngStartRow + 1, 1)
    
    ' 選択した行のデータを取得
    tmpLngSelectedRowL1 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 1)
    tmpLngSelectedRowL2 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 2)
    tmpLngSelectedRowL3 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 3)
    tmpLngSelectedRowL4 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 4)
    tmpLngSelectedRowL5 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 5)
    tmpLngSelectedRowTask = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 6)
    
    ' 更新対象範囲列のデータを更新
    If blnSelectedRowIsTask = True Then
        ' # 選択行がタスクの場合 #
        ' 対象となるデータインデックスをコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedRowLevel = 5 And _
                    tmpVarTargetArray(i, 6) >= tmpLngSelectedRowTask And _
                    tmpVarTargetArray(i, 5) = tmpLngSelectedRowL5 And _
                    tmpVarTargetArray(i, 4) = tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 4 And _
                    tmpVarTargetArray(i, 6) >= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    tmpVarTargetArray(i, 4) = tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 3 And _
                    tmpVarTargetArray(i, 6) >= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    IsEmpty(tmpVarTargetArray(i, 4)) And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 2 And _
                    tmpVarTargetArray(i, 6) >= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    IsEmpty(tmpVarTargetArray(i, 4)) And _
                    IsEmpty(tmpVarTargetArray(i, 3)) And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 1 And _
                    tmpVarTargetArray(i, 6) >= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    IsEmpty(tmpVarTargetArray(i, 4)) And _
                    IsEmpty(tmpVarTargetArray(i, 3)) And _
                    IsEmpty(tmpVarTargetArray(i, 2)) And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
        Next r
        ' 対象となるデータインデックスのみ値を更新する
        For Each tmpVarIdx In colTargetIdx
            tmpVarTargetArray(tmpVarIdx, 6) = tmpVarTargetArray(tmpVarIdx, 6) + 1
        Next tmpVarIdx
    Else
        ' # 選択行がタスクでない場合 #
        ' 対象となるデータインデックスをコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedRowLevel = 5 And _
                    tmpVarTargetArray(i, 5) >= tmpLngSelectedRowL5 And _
                    tmpVarTargetArray(i, 4) = tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 4 And _
                    tmpVarTargetArray(i, 4) >= tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 3 And _
                    tmpVarTargetArray(i, 3) >= tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 2 And _
                    tmpVarTargetArray(i, 2) >= tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 1 And _
                    tmpVarTargetArray(i, 1) >= tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
        Next r
        ' 対象となるデータインデックスのみ値を更新する
        For Each tmpVarIdx In colTargetIdx
            tmpVarTargetArray(tmpVarIdx, intSelectedRowLevel) = tmpVarTargetArray(tmpVarIdx, intSelectedRowLevel) + 1
        Next tmpVarIdx
    End If
    
    ' データの更新結果を反映
    tmpRngTarget.value = tmpVarTargetArray

End Sub


' ■ 選択行の最終レベルIDをデクリメント
Public Sub ExecDecrementSelectedLastLevel(ws As Worksheet)

    ' 変数定義
    Dim lngSelectedRow As Long, intSelectedRowLevel As Integer, blnSelectedRowIsTask As Boolean, lngSelectedRowLastValue As Long
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colTargetIdx As New Collection
    Dim lngFirstMissingFoundValue As Long
    ' 一時変数定義
    Dim r As Long, i As Long, v As Long
    Dim tmpRngTarget As Range
    Dim tmpVarTargetArray As Variant
    Dim tmpVarLevelArray As Variant
    Dim tmpVarTaskArray As Variant
    Dim tmpLngSelectedRowL1 As Long, tmpLngSelectedRowL2 As Long, tmpLngSelectedRowL3 As Long, tmpLngSelectedRowL4 As Long, tmpLngSelectedRowL5 As Long, tmpLngSelectedRowTask As Long
    Dim tmpVarIdx As Variant
    Dim tmpColTargetValue As New Collection
    Dim tmpVal As Variant, tmpBlnExist As Boolean
    
    ' 開始行と終了行を取得
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 選択した行の番号を取得
    lngSelectedRow = GetCheckedOptSingleRow(ws)
    
    ' ガード条件（未選択の場合は、メッセージを出して終了）
    If lngSelectedRow = 0 Then
        MsgBox "選択してください（OPT)。", vbExclamation, "通知"
        Exit Sub
    End If
    
    ' あらかじめ更新対象範囲列のデータを取得
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    tmpVarTargetArray = tmpRngTarget.value
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' 選択した行のレベルを取得
    intSelectedRowLevel = tmpVarLevelArray(lngSelectedRow - lngStartRow + 1, 1)
    ' 選択した行がタスクかどうか取得
    blnSelectedRowIsTask = tmpVarTaskArray(lngSelectedRow - lngStartRow + 1, 1)
    ' 選択した行の末尾の値を取得
    If blnSelectedRowIsTask Then
        lngSelectedRowLastValue = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 6)
    Else
        lngSelectedRowLastValue = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, intSelectedRowLevel)
    End If
    
    ' 選択した行のデータを取得
    tmpLngSelectedRowL1 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 1)
    tmpLngSelectedRowL2 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 2)
    tmpLngSelectedRowL3 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 3)
    tmpLngSelectedRowL4 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 4)
    tmpLngSelectedRowL5 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 5)
    tmpLngSelectedRowTask = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 6)
    
    ' 更新対象範囲列のデータを更新
    If blnSelectedRowIsTask = True Then
        ' # 選択行がタスクの場合 #
        ' 対象となる値をコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedRowLevel = 5 And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    tmpVarTargetArray(i, 5) = tmpLngSelectedRowL5 And _
                    tmpVarTargetArray(i, 4) = tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, 6), CStr(tmpVarTargetArray(i, 6))
                On Error GoTo 0
            End If
            If intSelectedRowLevel = 4 And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    tmpVarTargetArray(i, 4) = tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, 6), CStr(tmpVarTargetArray(i, 6))
                On Error GoTo 0
            End If
            If intSelectedRowLevel = 3 And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    IsEmpty(tmpVarTargetArray(i, 4)) And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, 6), CStr(tmpVarTargetArray(i, 6))
                On Error GoTo 0
            End If
            If intSelectedRowLevel = 2 And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    IsEmpty(tmpVarTargetArray(i, 4)) And _
                    IsEmpty(tmpVarTargetArray(i, 3)) And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, 6), CStr(tmpVarTargetArray(i, 6))
                On Error GoTo 0
            End If
            If intSelectedRowLevel = 1 And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    IsEmpty(tmpVarTargetArray(i, 4)) And _
                    IsEmpty(tmpVarTargetArray(i, 3)) And _
                    IsEmpty(tmpVarTargetArray(i, 2)) And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, 6), CStr(tmpVarTargetArray(i, 6))
                On Error GoTo 0
            End If
        Next r
        ' 値コレクションをから最初の存在しない値を取得
        lngFirstMissingFoundValue = 0
        For v = lngSelectedRowLastValue To 1 Step -1
            tmpBlnExist = False
            For Each tmpVal In tmpColTargetValue
                If v = tmpVal Then
                    tmpBlnExist = True
                    Exit For
                End If
            Next tmpVal
            If tmpBlnExist = False Then
                lngFirstMissingFoundValue = v
                Exit For
            End If
        Next v
        ' ガード条件（空き番号が存在しなかったら終了）
        If lngFirstMissingFoundValue = 0 Then
            MsgBox "空き番号がありません。", vbExclamation, "通知"
            Exit Sub
        End If
        ' 対象となるデータインデックスをコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedRowLevel = 5 And _
                    tmpVarTargetArray(i, 6) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    tmpVarTargetArray(i, 5) = tmpLngSelectedRowL5 And _
                    tmpVarTargetArray(i, 4) = tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 4 And _
                    tmpVarTargetArray(i, 6) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    tmpVarTargetArray(i, 4) = tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 3 And _
                    tmpVarTargetArray(i, 6) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    IsEmpty(tmpVarTargetArray(i, 4)) And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 2 And _
                    tmpVarTargetArray(i, 6) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    IsEmpty(tmpVarTargetArray(i, 4)) And _
                    IsEmpty(tmpVarTargetArray(i, 3)) And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 1 And _
                    tmpVarTargetArray(i, 6) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 6) <= tmpLngSelectedRowTask And _
                    IsEmpty(tmpVarTargetArray(i, 5)) And _
                    IsEmpty(tmpVarTargetArray(i, 4)) And _
                    IsEmpty(tmpVarTargetArray(i, 3)) And _
                    IsEmpty(tmpVarTargetArray(i, 2)) And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
        Next r
        ' 対象となるデータインデックスのみ値を更新する
        For Each tmpVarIdx In colTargetIdx
            tmpVarTargetArray(tmpVarIdx, 6) = tmpVarTargetArray(tmpVarIdx, 6) - 1
        Next tmpVarIdx
    Else
        ' # 選択行がタスクでない場合 #
        ' 対象となる値をコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedRowLevel = 5 And _
                    tmpVarTargetArray(i, 5) <= tmpLngSelectedRowL5 And _
                    tmpVarTargetArray(i, 4) = tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, intSelectedRowLevel), CStr(tmpVarTargetArray(i, intSelectedRowLevel))
                On Error GoTo 0
            End If
            If intSelectedRowLevel = 4 And _
                    tmpVarTargetArray(i, 4) <= tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, intSelectedRowLevel), CStr(tmpVarTargetArray(i, intSelectedRowLevel))
                On Error GoTo 0
            End If
            If intSelectedRowLevel = 3 And _
                    tmpVarTargetArray(i, 3) <= tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, intSelectedRowLevel), CStr(tmpVarTargetArray(i, intSelectedRowLevel))
                On Error GoTo 0
            End If
            If intSelectedRowLevel = 2 And _
                    tmpVarTargetArray(i, 2) <= tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, intSelectedRowLevel), CStr(tmpVarTargetArray(i, intSelectedRowLevel))
                On Error GoTo 0
            End If
            If intSelectedRowLevel = 1 And _
                    tmpVarTargetArray(i, 1) <= tmpLngSelectedRowL1 Then
                On Error Resume Next
                tmpColTargetValue.Add tmpVarTargetArray(i, intSelectedRowLevel), CStr(tmpVarTargetArray(i, intSelectedRowLevel))
                On Error GoTo 0
            End If
        Next r
        ' 値コレクションをから最初の存在しない値を取得
        lngFirstMissingFoundValue = 0
        For v = lngSelectedRowLastValue To 1 Step -1
            tmpBlnExist = False
            For Each tmpVal In tmpColTargetValue
                If v = tmpVal Then
                    tmpBlnExist = True
                    Exit For
                End If
            Next tmpVal
            If tmpBlnExist = False Then
                lngFirstMissingFoundValue = v
                Exit For
            End If
        Next v
        ' ガード条件（空き番号が存在しなかったら終了）
        If lngFirstMissingFoundValue = 0 Then
            MsgBox "空き番号がありません。", vbExclamation, "通知"
            Exit Sub
        End If
        ' 対象となるデータインデックスをコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedRowLevel = 5 And _
                    tmpVarTargetArray(i, 5) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 5) <= tmpLngSelectedRowL5 And _
                    tmpVarTargetArray(i, 4) = tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 4 And _
                    tmpVarTargetArray(i, 4) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 4) <= tmpLngSelectedRowL4 And _
                    tmpVarTargetArray(i, 3) = tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 3 And _
                    tmpVarTargetArray(i, 3) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 3) <= tmpLngSelectedRowL3 And _
                    tmpVarTargetArray(i, 2) = tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 2 And _
                    tmpVarTargetArray(i, 2) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 2) <= tmpLngSelectedRowL2 And _
                    tmpVarTargetArray(i, 1) = tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedRowLevel = 1 And _
                    tmpVarTargetArray(i, 1) > lngFirstMissingFoundValue And _
                    tmpVarTargetArray(i, 1) <= tmpLngSelectedRowL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
        Next r
        ' 対象となるデータインデックスのみ値を更新する
        For Each tmpVarIdx In colTargetIdx
            tmpVarTargetArray(tmpVarIdx, intSelectedRowLevel) = tmpVarTargetArray(tmpVarIdx, intSelectedRowLevel) - 1
        Next tmpVarIdx
    End If
    
    ' データの更新結果を反映
    tmpRngTarget.value = tmpVarTargetArray

End Sub



