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
Public Function ExecCheckWbsHasErrors(ws As Worksheet, _
                                        Optional ByVal blnShowMessage As Boolean = True) As Boolean

    ' 変数定義
    Dim blnResult As Boolean
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

    ' 開始行と終了行が見つからなければエラー終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then
        ExecCheckWbsHasErrors = True
        Exit Function
    End If

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
    varData = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value

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
            tmpErrorCount = colError1Count.Item(tmpRowIdx)
        End If
        ' まだエラーが発生していない行で、WbsId が登録されているもののみ検査
        If tmpErrorCount = 0 And utils.ExistsColKey(colWbsId, tmpRowIdx) Then
            tmpWbsId = colWbsId.Item(tmpRowIdx)
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
                    tmpParentWbsId = colParentWbsId.Item(tmpWbsId)
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
                tmpErrorCount = tmpErrorCount + colError1Count.Item(tmpRowIdx)
                If utils.ExistsColKey(colError1Message, tmpRowIdx) = True Then
                    tmpErrorMessage = tmpErrorMessage & colError1Message.Item(tmpRowIdx)
                End If
            End If
            If utils.ExistsColKey(colError2Count, tmpRowIdx) = True Then
                tmpErrorCount = tmpErrorCount + colError2Count.Item(tmpRowIdx)
                If utils.ExistsColKey(colError2Message, tmpRowIdx) = True Then
                    tmpErrorMessage = tmpErrorMessage & colError2Message.Item(tmpRowIdx)
                End If
            End If
            If utils.ExistsColKey(colError3Count, tmpRowIdx) = True Then
                tmpErrorCount = tmpErrorCount + colError3Count.Item(tmpRowIdx)
                If utils.ExistsColKey(colError3Message, tmpRowIdx) = True Then
                    tmpErrorMessage = tmpErrorMessage & colError3Message.Item(tmpRowIdx)
                End If
            End If
            If tmpErrorCount > 0 Then
                ws.Cells(r, cfg.COL_ERR).Value = "E"
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
    If blnShowMessage = True And intErrorCount > 0 Then
        MsgBox intErrorCount & " 件の異常を検出しました。", vbExclamation, "エラーチェック"
    End If
    
    ExecCheckWbsHasErrors = blnHasError
End Function


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
    strFormula = "=CustomFormatWbsIdx(" & _
                    cfg.COL_ERR_LABEL & lngStartRow & "," & _
                    cfg.COL_L1_LABEL & lngStartRow & "," & _
                    cfg.COL_L2_LABEL & lngStartRow & "," & _
                    cfg.COL_L3_LABEL & lngStartRow & "," & _
                    cfg.COL_L4_LABEL & lngStartRow & "," & _
                    cfg.COL_L5_LABEL & lngStartRow & "," & _
                    cfg.COL_TASK_LABEL & lngStartRow & ")"

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
    strFormula = "=CustomFormatWbsId(" & _
                    cfg.COL_ERR_LABEL & lngStartRow & "," & _
                    cfg.COL_L1_LABEL & lngStartRow & "," & _
                    cfg.COL_L2_LABEL & lngStartRow & "," & _
                    cfg.COL_L3_LABEL & lngStartRow & "," & _
                    cfg.COL_L4_LABEL & lngStartRow & "," & _
                    cfg.COL_L5_LABEL & lngStartRow & "," & _
                    cfg.COL_TASK_LABEL & lngStartRow & ")"

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
    strFormula = "=CustomFuncGetLevel(" & _
                    cfg.COL_L1_LABEL & lngStartRow & "," & _
                    cfg.COL_L2_LABEL & lngStartRow & "," & _
                    cfg.COL_L3_LABEL & lngStartRow & "," & _
                    cfg.COL_L4_LABEL & lngStartRow & "," & _
                    cfg.COL_L5_LABEL & lngStartRow & ")"

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
    Dim varFormulas() As Variant
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
    
    ' 数式をセットするデータを用意
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_PLANNED_EFF), ws.Cells(lngEndRow, cfg.COL_PLANNED_EFF)).Formula = varFormulas
    
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


' □ 再帰的に予定工数を合計してセットする
Private Sub SetValueRecursiveForPlannedEffort(ws As Worksheet, _
                                                varValues As Variant, _
                                                varHierarchyArray As Variant, _
                                                varFlgIcArray As Variant, _
                                                varPlannedEffortArray As Variant, _
                                                lngTargetIdx As Long)
    
    ' 変数定義
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumEffort As Double
    ' 一時変数定義
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' ガード条件（入力されたインデックスが0以下の場合は終了）
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の行数を越えたインデックスを指定された場合は終了）
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' ガード条件（既に値が求められている場合は終了）
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の列数が6でない場合は終了）
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' 指定インデックスの値を取得
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' タスク状態の取得
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' レベルの取得
    If IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsNumeric(varTargetL5) And Not IsNull(varTargetL5) And Not IsEmpty(varTargetL5) Then
        intTargetLevel = 5
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 4
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 3
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 2
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 1
    Else
        ' # 階層に問題がある場合 #
        Exit Sub
    End If
    
    ' メイン処理
    If blnTargetTask = True Then
        ' # タスクには子階層がないため、1をセット #
        If IsEmpty(varPlannedEffortArray(lngTargetIdx, 1)) Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = varPlannedEffortArray(lngTargetIdx, 1)
        End If
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # タスクでない場合、子階層を集計して値をセット #
        
        ' 子階層を取得
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' ガード条件（子階層が存在しない場合、0をセットして終了）
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' 階層の値をチェックし、未セットなら再帰的に関数を呼び出し、値を集計
        dblSumEffort = 0
        For Each tmpVarChildIdx In tmpColChildIdxs
            
            If Not IsEmpty(varFlgIcArray(tmpVarChildIdx, 1)) And varFlgIcArray(tmpVarChildIdx, 1) = True Then
                If IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                    SetValueRecursiveForPlannedEffort ws, varValues, varHierarchyArray, varFlgIcArray, varPlannedEffortArray, CLng(tmpVarChildIdx)
                    If Not IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                        dblSumEffort = dblSumEffort + varValues(tmpVarChildIdx, 1)
                    End If
                Else
                    dblSumEffort = dblSumEffort + varValues(tmpVarChildIdx, 1)
                End If
            End If
            
        Next tmpVarChildIdx
        varValues(lngTargetIdx, 1) = dblSumEffort
        varValues(lngTargetIdx, 2) = intTargetLevel
        
    End If
    
End Sub


' ■ 予定工数を集計した値をセット
Public Sub SetValueForPlannedEffort(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varPlannedEffortArray As Variant
    Dim dblSumEffort As Double
    ' 一時変数定義
    Dim r As Long, i As Long

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 値をセットするデータを用意
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' あらかじめチェック対象範囲列のデータを取得
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' あらかじめFLG_IC列のデータを取得
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' あらかじめ予定工数列のデータを取得
    varPlannedEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_PLANNED_EFF), ws.Cells(lngEndRow, cfg.COL_PLANNED_EFF)).Value
    
    ' 順番に集計を行う
    dblSumEffort = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForPlannedEffort ws, varValues, varHierarchyArray, varFlgIcArray, varPlannedEffortArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            dblSumEffort = dblSumEffort + varValues(i, 1)
        End If
    Next i
    
    ' 結果を反映する
    ws.Range(ws.Cells(lngStartRow, cfg.COL_PLANNED_EFF), ws.Cells(lngEndRow, cfg.COL_PLANNED_EFF)).Value = varValues
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngEndRow + 2).Value = dblSumEffort

End Sub


' ■ 実績済工数を集計する式をセット
Public Sub SetFormulaForActualCompletedEffort(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
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
    
    ' 数式をセットするデータを用意
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_COMPLETED_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_COMPLETED_EFF)).Formula = varFormulas
    
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


' □ 再帰的に実績済工数を合計してセットする
Private Sub SetValueRecursiveForActualCompletedEffort(ws As Worksheet, _
                                                        varValues As Variant, _
                                                        varHierarchyArray As Variant, _
                                                        varFlgIcArray As Variant, _
                                                        varActualCompletedEffortArray As Variant, _
                                                        lngTargetIdx As Long)
    
    ' 変数定義
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumEffort As Double
    ' 一時変数定義
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' ガード条件（入力されたインデックスが0以下の場合は終了）
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の行数を越えたインデックスを指定された場合は終了）
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' ガード条件（既に値が求められている場合は終了）
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の列数が6でない場合は終了）
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' 指定インデックスの値を取得
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' タスク状態の取得
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' レベルの取得
    If IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsNumeric(varTargetL5) And Not IsNull(varTargetL5) And Not IsEmpty(varTargetL5) Then
        intTargetLevel = 5
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 4
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 3
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 2
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 1
    Else
        ' # 階層に問題がある場合 #
        Exit Sub
    End If
    
    ' メイン処理
    If blnTargetTask = True Then
        ' # タスクには子階層がないため、1をセット #
        If IsEmpty(varActualCompletedEffortArray(lngTargetIdx, 1)) Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = varActualCompletedEffortArray(lngTargetIdx, 1)
        End If
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # タスクでない場合、子階層を集計して値をセット #
        
        ' 子階層を取得
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' ガード条件（子階層が存在しない場合、0をセットして終了）
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' 階層の値をチェックし、未セットなら再帰的に関数を呼び出し、値を集計
        dblSumEffort = 0
        For Each tmpVarChildIdx In tmpColChildIdxs
            
            If Not IsEmpty(varFlgIcArray(tmpVarChildIdx, 1)) And varFlgIcArray(tmpVarChildIdx, 1) = True Then
                If IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                    SetValueRecursiveForActualCompletedEffort ws, varValues, varHierarchyArray, varFlgIcArray, varActualCompletedEffortArray, CLng(tmpVarChildIdx)
                    If Not IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                        dblSumEffort = dblSumEffort + varValues(tmpVarChildIdx, 1)
                    End If
                Else
                    dblSumEffort = dblSumEffort + varValues(tmpVarChildIdx, 1)
                End If
            End If
            
        Next tmpVarChildIdx
        varValues(lngTargetIdx, 1) = dblSumEffort
        varValues(lngTargetIdx, 2) = intTargetLevel
        
    End If
    
End Sub


' ■ 実績済工数を集計した値をセット
Public Sub SetValueForActualCompletedEffort(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varActualCompletedEffortArray As Variant
    Dim dblSumEffort As Double
    ' 一時変数定義
    Dim r As Long, i As Long

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 値をセットするデータを用意
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' あらかじめチェック対象範囲列のデータを取得
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' あらかじめFLG_IC列のデータを取得
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' あらかじめ実績済工数のデータを取得
    varActualCompletedEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_COMPLETED_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_COMPLETED_EFF)).Value
    
    ' 順番に集計を行う
    dblSumEffort = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForActualCompletedEffort ws, varValues, varHierarchyArray, varFlgIcArray, varActualCompletedEffortArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            dblSumEffort = dblSumEffort + varValues(i, 1)
        End If
    Next i
    
    ' 結果を反映する
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_COMPLETED_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_COMPLETED_EFF)).Value = varValues
    ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow + 2).Value = dblSumEffort

End Sub


' ■ 実績残工数を集計する式をセット
Public Sub SetFormulaForActualRemainingEffort(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
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
    
    ' 数式をセットするデータを用意
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_REMAINING_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_REMAINING_EFF)).Formula = varFormulas
    
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


' □ 再帰的に実績残工数を合計してセットする
Private Sub SetValueRecursiveForActualRemainingEffort(ws As Worksheet, _
                                                        varValues As Variant, _
                                                        varHierarchyArray As Variant, _
                                                        varFlgIcArray As Variant, _
                                                        varActualRemainingEffortArray As Variant, _
                                                        lngTargetIdx As Long)
    
    ' 変数定義
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumEffort As Double
    ' 一時変数定義
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' ガード条件（入力されたインデックスが0以下の場合は終了）
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の行数を越えたインデックスを指定された場合は終了）
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' ガード条件（既に値が求められている場合は終了）
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の列数が6でない場合は終了）
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' 指定インデックスの値を取得
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' タスク状態の取得
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' レベルの取得
    If IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsNumeric(varTargetL5) And Not IsNull(varTargetL5) And Not IsEmpty(varTargetL5) Then
        intTargetLevel = 5
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 4
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 3
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 2
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 1
    Else
        ' # 階層に問題がある場合 #
        Exit Sub
    End If
    
    ' メイン処理
    If blnTargetTask = True Then
        ' # タスクには子階層がないため、1をセット #
        If IsEmpty(varActualRemainingEffortArray(lngTargetIdx, 1)) Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = varActualRemainingEffortArray(lngTargetIdx, 1)
        End If
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # タスクでない場合、子階層を集計して値をセット #
        
        ' 子階層を取得
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' ガード条件（子階層が存在しない場合、0をセットして終了）
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' 階層の値をチェックし、未セットなら再帰的に関数を呼び出し、値を集計
        dblSumEffort = 0
        For Each tmpVarChildIdx In tmpColChildIdxs
            
            If Not IsEmpty(varFlgIcArray(tmpVarChildIdx, 1)) And varFlgIcArray(tmpVarChildIdx, 1) = True Then
                If IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                    SetValueRecursiveForActualRemainingEffort ws, varValues, varHierarchyArray, varFlgIcArray, varActualRemainingEffortArray, CLng(tmpVarChildIdx)
                    If Not IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                        dblSumEffort = dblSumEffort + varValues(tmpVarChildIdx, 1)
                    End If
                Else
                    dblSumEffort = dblSumEffort + varValues(tmpVarChildIdx, 1)
                End If
            End If
            
        Next tmpVarChildIdx
        varValues(lngTargetIdx, 1) = dblSumEffort
        varValues(lngTargetIdx, 2) = intTargetLevel
        
    End If
    
End Sub


' ■ 実績残工数を集計した値をセット
Public Sub SetValueForActualRemainingEffort(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varActualRemainingEffortArray As Variant
    Dim dblSumEffort As Double
    ' 一時変数定義
    Dim r As Long, i As Long

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 値をセットするデータを用意
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' あらかじめチェック対象範囲列のデータを取得
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' あらかじめFLG_IC列のデータを取得
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' あらかじめ実績残工数のデータを取得
    varActualRemainingEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_REMAINING_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_REMAINING_EFF)).Value
    
    ' 順番に集計を行う
    dblSumEffort = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForActualRemainingEffort ws, varValues, varHierarchyArray, varFlgIcArray, varActualRemainingEffortArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            dblSumEffort = dblSumEffort + varValues(i, 1)
        End If
    Next i
    
    ' 結果を反映する
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_REMAINING_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_REMAINING_EFF)).Value = varValues
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow + 2).Value = dblSumEffort

End Sub


' ■ タスク進捗率を集計する式をセット
Public Sub SetFormulaForTaskProgressRate(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    Dim varNumberFormats() As Variant
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarTaskProgArray As Variant
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
    
    ' 数式をセットするデータを用意
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    ReDim varNumberFormats(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' あらかじめ項目消化率列のデータを取得
    tmpVarTaskProgArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).Value
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
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
            varNumberFormats(i, 1) = "0.0%"
            varFormulas(i, 1) = tmpVarTaskProgArray(i, 1)
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
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
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
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
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
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
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
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
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
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).NumberFormat = varNumberFormats
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).Formula = varFormulas
    
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


' □ 再帰的にタスク進捗率を集計してセットする
Private Sub SetValueRecursiveForTaskProgressRate(ws As Worksheet, _
                                                    varValues As Variant, _
                                                    varHierarchyArray As Variant, _
                                                    varFlgIcArray As Variant, _
                                                    varTaskProgressRateArray As Variant, _
                                                    varTaskWeightArray As Variant, _
                                                    lngTargetIdx As Long)
    
    ' 変数定義
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumProgressRate As Double
    Dim intSumWeight As Integer
    ' 一時変数定義
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    Dim tmpIntWeight As Integer
    
    ' ガード条件（入力されたインデックスが0以下の場合は終了）
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の行数を越えたインデックスを指定された場合は終了）
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' ガード条件（既に値が求められている場合は終了）
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の列数が6でない場合は終了）
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' 指定インデックスの値を取得
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' タスク状態の取得
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' レベルの取得
    If IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsNumeric(varTargetL5) And Not IsNull(varTargetL5) And Not IsEmpty(varTargetL5) Then
        intTargetLevel = 5
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 4
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 3
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 2
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 1
    Else
        ' # 階層に問題がある場合 #
        Exit Sub
    End If
    
    ' メイン処理
    If blnTargetTask = True Then
        ' # タスクには子階層がないため、1をセット #
        If IsEmpty(varTaskProgressRateArray(lngTargetIdx, 1)) Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = varTaskProgressRateArray(lngTargetIdx, 1)
        End If
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # タスクでない場合、子階層を集計して値をセット #
        
        ' 子階層を取得
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' ガード条件（子階層が存在しない場合、0をセットして終了）
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' 階層の値をチェックし、未セットなら再帰的に関数を呼び出し、値を集計
        dblSumProgressRate = 0
        intSumWeight = 0
        For Each tmpVarChildIdx In tmpColChildIdxs
        
            If IsEmpty(varTaskWeightArray(tmpVarChildIdx, 1)) Then
                tmpIntWeight = 0
            Else
                tmpIntWeight = varTaskWeightArray(tmpVarChildIdx, 1)
            End If
            If Not IsEmpty(varFlgIcArray(tmpVarChildIdx, 1)) And varFlgIcArray(tmpVarChildIdx, 1) = True Then
                If IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                    SetValueRecursiveForTaskProgressRate ws, varValues, varHierarchyArray, varFlgIcArray, varTaskProgressRateArray, varTaskWeightArray, CLng(tmpVarChildIdx)
                    If Not IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                        dblSumProgressRate = dblSumProgressRate + (varValues(tmpVarChildIdx, 1) * tmpIntWeight)
                    End If
                Else
                    dblSumProgressRate = dblSumProgressRate + (varValues(tmpVarChildIdx, 1) * tmpIntWeight)
                End If
                intSumWeight = intSumWeight + tmpIntWeight
            End If
            
        Next tmpVarChildIdx
        
        If intSumWeight = 0 Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = dblSumProgressRate / intSumWeight
        End If
        varValues(lngTargetIdx, 2) = intTargetLevel
        
    End If
    
End Sub


' ■ タスク進捗率を集計した値をセット
Public Sub SetValueForTaskProgressRate(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varTaskProgressRateArray As Variant
    Dim varTaskWeightArray As Variant
    Dim dblSumRate As Double
    Dim intSumWeight As Integer
    ' 一時変数定義
    Dim r As Long, i As Long
    Dim tmpIntWeight As Integer

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 値をセットするデータを用意（1:集計結果［工数進捗率］、2:レベル）
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' あらかじめチェック対象範囲列のデータを取得
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' あらかじめFLG_IC列のデータを取得
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' あらかじめ項目消化率列のデータを取得
    varTaskProgressRateArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).Value
    ' あらかじめ項目加重列のデータを取得
    varTaskWeightArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_WGT), ws.Cells(lngEndRow, cfg.COL_TASK_WGT)).Value
    
    ' 順番に集計を行う
    dblSumRate = 0
    intSumWeight = 0
    For i = 1 To UBound(varHierarchyArray, 1)
    
        If IsEmpty(varTaskWeightArray(i, 1)) Then
            tmpIntWeight = 0
        Else
            tmpIntWeight = varTaskWeightArray(i, 1)
        End If
        SetValueRecursiveForTaskProgressRate ws, varValues, varHierarchyArray, varFlgIcArray, varTaskProgressRateArray, varTaskWeightArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            dblSumRate = dblSumRate + (varValues(i, 1) * tmpIntWeight)
            intSumWeight = intSumWeight + tmpIntWeight
        End If
        
    Next i
    
    ' 結果を反映する
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).Value = varValues
    If intSumWeight = 0 Then
        ws.Range(cfg.COL_TASK_PROG_LABEL & lngEndRow + 2).Value = 0
    Else
        ws.Range(cfg.COL_TASK_PROG_LABEL & lngEndRow + 2).Value = dblSumRate / intSumWeight
    End If
    
End Sub


' ■ 工数進捗率を集計する式をセット
Public Sub SetFormulaForEffortProgressRate(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
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
    
    ' 数式をセットするデータを用意
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
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
            varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_EFFORT_PROG), ws.Cells(lngEndRow, cfg.COL_EFFORT_PROG)).Formula = varFormulas
    
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


' □ 再帰的に工数進捗率を集計してセットする
Private Sub SetValueRecursiveForEffortProgressRate(ws As Worksheet, _
                                                    varValues As Variant, _
                                                    varHierarchyArray As Variant, _
                                                    varFlgIcArray As Variant, _
                                                    varActualRemainingEffortArray As Variant, _
                                                    varActualCompletedEffortArray As Variant, _
                                                    lngTargetIdx As Long)
    
    ' 変数定義
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumProgressRate As Double
    Dim intSumCount As Integer
    ' 一時変数定義
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    Dim tmpDblActualRemainingEffort As Double
    Dim tmpDblActualCompletedEffort As Double
    
    ' ガード条件（入力されたインデックスが0以下の場合は終了）
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の行数を越えたインデックスを指定された場合は終了）
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' ガード条件（既に値が求められている場合は終了）
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の列数が6でない場合は終了）
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' 指定インデックスの値を取得
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' タスク状態の取得
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' レベルの取得
    If IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsNumeric(varTargetL5) And Not IsNull(varTargetL5) And Not IsEmpty(varTargetL5) Then
        intTargetLevel = 5
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 4
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 3
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 2
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 1
    Else
        ' # 階層に問題がある場合 #
        Exit Sub
    End If
    
    ' メイン処理
    If blnTargetTask = True Then
        ' # タスクには子階層がないため、1をセット #
        tmpDblActualRemainingEffort = 0
        If IsEmpty(varActualRemainingEffortArray(lngTargetIdx, 1)) Then
            tmpDblActualRemainingEffort = 0
        Else
            tmpDblActualRemainingEffort = varActualRemainingEffortArray(lngTargetIdx, 1)
        End If
        tmpDblActualCompletedEffort = 0
        If IsEmpty(varActualCompletedEffortArray(lngTargetIdx, 1)) Then
            tmpDblActualCompletedEffort = 0
        Else
            tmpDblActualCompletedEffort = varActualCompletedEffortArray(lngTargetIdx, 1)
        End If
        If tmpDblActualRemainingEffort = 0 And tmpDblActualCompletedEffort = 0 Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = tmpDblActualCompletedEffort / (tmpDblActualRemainingEffort + tmpDblActualCompletedEffort)
        End If
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # タスクでない場合、子階層を集計して値をセット #
        
        ' 子階層を取得
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' ガード条件（子階層が存在しない場合、0をセットして終了）
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' 階層の値をチェックし、未セットなら再帰的に関数を呼び出し、値を集計
        dblSumProgressRate = 0
        intSumCount = 0
        For Each tmpVarChildIdx In tmpColChildIdxs
            
            If Not IsEmpty(varFlgIcArray(tmpVarChildIdx, 1)) And varFlgIcArray(tmpVarChildIdx, 1) = True Then
                If IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                    SetValueRecursiveForEffortProgressRate ws, varValues, varHierarchyArray, varFlgIcArray, varActualRemainingEffortArray, varActualCompletedEffortArray, CLng(tmpVarChildIdx)
                    If Not IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                        dblSumProgressRate = dblSumProgressRate + varValues(tmpVarChildIdx, 1)
                    End If
                Else
                    dblSumProgressRate = dblSumProgressRate + varValues(tmpVarChildIdx, 1)
                End If
                intSumCount = intSumCount + 1
            End If
            
        Next tmpVarChildIdx
        
        varValues(lngTargetIdx, 1) = dblSumProgressRate / intSumCount
        varValues(lngTargetIdx, 2) = intTargetLevel
        
    End If
    
End Sub


' ■ 工数進捗率を集計した値をセット
Public Sub SetValueForEffortProgressRate(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varActualRemainingEffortArray As Variant
    Dim varActualCompletedEffortArray As Variant
    Dim dblSumRate As Double
    Dim intSumCount As Integer
    ' 一時変数定義
    Dim r As Long, i As Long

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 値をセットするデータを用意（1:集計結果［工数進捗率］、2:レベル）
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' あらかじめチェック対象範囲列のデータを取得
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' あらかじめFLG_IC列のデータを取得
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' あらかじめ実績残工数のデータを取得
    varActualRemainingEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_REMAINING_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_REMAINING_EFF)).Value
    ' あらかじめ実績済工数のデータを取得
    varActualCompletedEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_COMPLETED_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_COMPLETED_EFF)).Value
    
    ' 順番に集計を行う
    dblSumRate = 0
    intSumCount = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForEffortProgressRate ws, varValues, varHierarchyArray, varFlgIcArray, varActualRemainingEffortArray, varActualCompletedEffortArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            dblSumRate = dblSumRate + varValues(i, 1)
            intSumCount = intSumCount + 1
        End If
    Next i
    
    ' 結果を反映する
    ws.Range(ws.Cells(lngStartRow, cfg.COL_EFFORT_PROG), ws.Cells(lngEndRow, cfg.COL_EFFORT_PROG)).Value = varValues
    If intSumCount = 0 Then
        ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2).Value = 0
    Else
        ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2).Value = dblSumRate / intSumCount
    End If
    
End Sub


' ■ タスク合計件数を集計する式をセット
Public Sub SetFormulaForTaskCount(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
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
    
    ' 数式をセットするデータを用意
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
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
            varFormulas(i, 1) = 1
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_COUNT), ws.Cells(lngEndRow, cfg.COL_TASK_COUNT)).Formula = varFormulas
    
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


' □ 再帰的にタスク合計数をカウントしてセットする
Private Sub SetValueRecursiveForTaskCount(ws As Worksheet, _
                                            varValues As Variant, _
                                            varHierarchyArray As Variant, _
                                            varFlgIcArray As Variant, _
                                            lngTargetIdx As Long)
    
    ' 変数定義
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim lngSumCount As Long
    ' 一時変数定義
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' ガード条件（入力されたインデックスが0以下の場合は終了）
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の行数を越えたインデックスを指定された場合は終了）
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' ガード条件（既に値が求められている場合は終了）
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の列数が6でない場合は終了）
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' 指定インデックスの値を取得
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' タスク状態の取得
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' レベルの取得
    If IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsNumeric(varTargetL5) And Not IsNull(varTargetL5) And Not IsEmpty(varTargetL5) Then
        intTargetLevel = 5
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 4
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 3
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 2
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 1
    Else
        ' # 階層に問題がある場合 #
        Exit Sub
    End If
    
    ' メイン処理
    If blnTargetTask = True Then
        ' # タスクには子階層がないため、1をセット #
        varValues(lngTargetIdx, 1) = 1
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # タスクでない場合、子階層を集計して値をセット #
        
        ' 子階層を取得
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' ガード条件（子階層が存在しない場合、0をセットして終了）
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' 階層の値をチェックし、未セットなら再帰的に関数を呼び出し、値を集計
        lngSumCount = 0
        For Each tmpVarChildIdx In tmpColChildIdxs
            
            If Not IsEmpty(varFlgIcArray(tmpVarChildIdx, 1)) And varFlgIcArray(tmpVarChildIdx, 1) = True Then
                If IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                    SetValueRecursiveForTaskCount ws, varValues, varHierarchyArray, varFlgIcArray, CLng(tmpVarChildIdx)
                    If Not IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                        lngSumCount = lngSumCount + varValues(tmpVarChildIdx, 1)
                    End If
                Else
                    lngSumCount = lngSumCount + varValues(tmpVarChildIdx, 1)
                End If
            End If
            
        Next tmpVarChildIdx
        varValues(lngTargetIdx, 1) = lngSumCount
        varValues(lngTargetIdx, 2) = intTargetLevel
        
    End If
    
End Sub


' ■ タスク合計件数を集計する式をセット
Public Sub SetValueForTaskCount(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim lngSumCount As Long
    ' 一時変数定義
    Dim r As Long, i As Long

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 値をセットするデータを用意
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' あらかじめチェック対象範囲列のデータを取得
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' あらかじめFLG_IC列のデータを取得
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    
    ' 順番に集計を行う
    lngSumCount = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForTaskCount ws, varValues, varHierarchyArray, varFlgIcArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            lngSumCount = lngSumCount + varValues(i, 1)
        End If
    Next i
    
    ' 結果を反映する
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_COUNT), ws.Cells(lngEndRow, cfg.COL_TASK_COUNT)).Value = varValues
    ws.Range(cfg.COL_TASK_COUNT_LABEL & lngEndRow + 2).Value = lngSumCount

End Sub


' ■ タスク完了件数を集計する式をセット
Public Sub SetFormulaForTaskCompCount(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
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
    
    ' 数式をセットするデータを用意
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' あらかじめWBSレベル列のデータを取得
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
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
            varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
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
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_COMP_COUNT), ws.Cells(lngEndRow, cfg.COL_TASK_COMP_COUNT)).Formula = varFormulas
    
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


' □ 再帰的にタスク完了数をカウントしてセットする
Private Sub SetValueRecursiveForTaskCompCount(ws As Worksheet, _
                                                varValues As Variant, _
                                                varHierarchyArray As Variant, _
                                                varFlgIcArray As Variant, _
                                                varWbsStatusArray As Variant, _
                                                lngTargetIdx As Long)
    
    ' 変数定義
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim lngSumCount As Long
    ' 一時変数定義
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' ガード条件（入力されたインデックスが0以下の場合は終了）
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の行数を越えたインデックスを指定された場合は終了）
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' ガード条件（既に値が求められている場合は終了）
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' ガード条件（入力された階層配列の列数が6でない場合は終了）
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' 指定インデックスの値を取得
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' タスク状態の取得
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' レベルの取得
    If IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsNumeric(varTargetL5) And Not IsNull(varTargetL5) And Not IsEmpty(varTargetL5) Then
        intTargetLevel = 5
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 4
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 3
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 2
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 1
    Else
        ' # 階層に問題がある場合 #
        Exit Sub
    End If
    
    ' メイン処理
    If blnTargetTask = True Then
        ' # タスクには子階層がないため、1をセット #
        If varWbsStatusArray(lngTargetIdx, 1) = cfg.WBS_STATUS_COMPLETED Then
            varValues(lngTargetIdx, 1) = 1
            varValues(lngTargetIdx, 2) = 6
        Else
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = 6
        End If
    Else
        ' # タスクでない場合、子階層を集計して値をセット #
        
        ' 子階層を取得
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' ガード条件（子階層が存在しない場合、0をセットして終了）
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' 階層の値をチェックし、未セットなら再帰的に関数を呼び出し、値を集計
        lngSumCount = 0
        For Each tmpVarChildIdx In tmpColChildIdxs
            
            If Not IsEmpty(varFlgIcArray(tmpVarChildIdx, 1)) And varFlgIcArray(tmpVarChildIdx, 1) = True Then
                If IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                    SetValueRecursiveForTaskCompCount ws, varValues, varHierarchyArray, varFlgIcArray, varWbsStatusArray, CLng(tmpVarChildIdx)
                    If Not IsEmpty(varValues(tmpVarChildIdx, 1)) Then
                        lngSumCount = lngSumCount + varValues(tmpVarChildIdx, 1)
                    End If
                Else
                    lngSumCount = lngSumCount + varValues(tmpVarChildIdx, 1)
                End If
            End If
            
        Next tmpVarChildIdx
        varValues(lngTargetIdx, 1) = lngSumCount
        varValues(lngTargetIdx, 2) = intTargetLevel
        
    End If
    
End Sub


' ■ タスク完了件数を集計する式をセット
Public Sub SetValueForTaskCompCount(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varWbsStatusArray As Variant
    Dim lngSumCount As Long
    ' 一時変数定義
    Dim r As Long, i As Long

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 値をセットするデータを用意
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' あらかじめチェック対象範囲列のデータを取得
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' あらかじめFLG_IC列のデータを取得
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' あらかじめWBSステータス列のデータを取得
    varWbsStatusArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_WBS_STATUS), ws.Cells(lngEndRow, cfg.COL_WBS_STATUS)).Value
    
    ' 順番に集計を行う
    lngSumCount = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForTaskCompCount ws, varValues, varHierarchyArray, varFlgIcArray, varWbsStatusArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            lngSumCount = lngSumCount + varValues(i, 1)
        End If
    Next i
    
    ' 結果を反映する
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_COMP_COUNT), ws.Cells(lngEndRow, cfg.COL_TASK_COMP_COUNT)).Value = varValues
    ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow + 2).Value = lngSumCount

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
    varData = ws.Range(ws.Cells(lngStartRow, cfg.COL_CHK), ws.Cells(lngEndRow, cfg.COL_CHK)).Value
    
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
    ' 一時変数定義
    Dim tmpLngCol As Long
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    
    ' 行を追加
    lngSelectedRow = GetCheckedOptSingleRow(ws)
    If lngSelectedRow <> 0 Then
        ' 行を追加
        ws.Rows(lngSelectedRow + 1).Insert Shift:=xlDown
        ' 1列ずつチェックして、基本数式だけコピー
        For tmpLngCol = cfg.COL_WBS_IDX To cfg.COL_WBS_ID
            If ws.Cells(lngSelectedRow, tmpLngCol).HasFormula Then
                ws.Cells(lngSelectedRow + 1, tmpLngCol).Formula = ws.Cells(lngSelectedRow, tmpLngCol).Formula
            End If
        Next tmpLngCol
    Else
        MsgBox "選択してください（OPT)。", vbExclamation, "通知"
    End If

End Sub


' ■ 選択行の最終レベルIDをインクリメント
Public Sub ExecIncrementSelectedLastLevel(ws As Worksheet)

    ' 変数定義
    Dim lngSelectedRow As Long, intSelectedLevel As Integer, blnSelectedIsTask As Boolean
    Dim varSelectedL1 As Variant, varSelectedL2 As Variant, varSelectedL3 As Variant, varSelectedL4 As Variant, varSelectedL5 As Variant, varSelectedTask As Variant
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colTargetIdx As New Collection
    Dim rngHierarchy As Range
    Dim varHierarchyArray As Variant
    Dim varLevelArray As Variant
    Dim varTaskArray As Variant
    ' 一時変数定義
    Dim r As Long, i As Long
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
    Set rngHierarchy = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    varHierarchyArray = rngHierarchy.Value
    ' あらかじめWBSレベル列のデータを取得
    varLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    varTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' 選択した行のレベルを取得
    intSelectedLevel = varLevelArray(lngSelectedRow - lngStartRow + 1, 1)
    ' 選択した行がタスクかどうか取得
    blnSelectedIsTask = varTaskArray(lngSelectedRow - lngStartRow + 1, 1)
    ' 選択した行のデータを取得
    varSelectedL1 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 1)
    varSelectedL2 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 2)
    varSelectedL3 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 3)
    varSelectedL4 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 4)
    varSelectedL5 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 5)
    varSelectedTask = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 6)
    
    ' 更新対象範囲列のデータを更新
    If blnSelectedIsTask = True Then
        ' # 選択行がタスクの場合 #
        ' 対象となるデータインデックスをコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedLevel = 5 And _
                    varHierarchyArray(i, 6) >= varSelectedTask And _
                    varHierarchyArray(i, 5) = varSelectedL5 And _
                    varHierarchyArray(i, 4) = varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 4 And _
                    varHierarchyArray(i, 6) >= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    varHierarchyArray(i, 4) = varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 3 And _
                    varHierarchyArray(i, 6) >= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    IsEmpty(varHierarchyArray(i, 4)) And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 2 And _
                    varHierarchyArray(i, 6) >= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    IsEmpty(varHierarchyArray(i, 4)) And _
                    IsEmpty(varHierarchyArray(i, 3)) And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 1 And _
                    varHierarchyArray(i, 6) >= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    IsEmpty(varHierarchyArray(i, 4)) And _
                    IsEmpty(varHierarchyArray(i, 3)) And _
                    IsEmpty(varHierarchyArray(i, 2)) And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
        Next r
        ' 対象となるデータインデックスのみ値を更新する
        For Each tmpVarIdx In colTargetIdx
            varHierarchyArray(tmpVarIdx, 6) = varHierarchyArray(tmpVarIdx, 6) + 1
        Next tmpVarIdx
    Else
        ' # 選択行がタスクでない場合 #
        ' 対象となるデータインデックスをコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedLevel = 5 And _
                    varHierarchyArray(i, 5) >= varSelectedL5 And _
                    varHierarchyArray(i, 4) = varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 4 And _
                    varHierarchyArray(i, 4) >= varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 3 And _
                    varHierarchyArray(i, 3) >= varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 2 And _
                    varHierarchyArray(i, 2) >= varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 1 And _
                    varHierarchyArray(i, 1) >= varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
        Next r
        ' 対象となるデータインデックスのみ値を更新する
        For Each tmpVarIdx In colTargetIdx
            varHierarchyArray(tmpVarIdx, intSelectedLevel) = varHierarchyArray(tmpVarIdx, intSelectedLevel) + 1
        Next tmpVarIdx
    End If
    
    ' データの更新結果を反映
    rngHierarchy.Value = varHierarchyArray

End Sub


' ■ 選択行の最終レベルIDをデクリメント
Public Sub ExecDecrementSelectedLastLevel(ws As Worksheet)

    ' 変数定義
    Dim lngSelectedRow As Long, intSelectedLevel As Integer, blnSelectedIsTask As Boolean, varSelectedLastValue As Variant
    Dim varSelectedL1 As Variant, varSelectedL2 As Variant, varSelectedL3 As Variant, varSelectedL4 As Variant, varSelectedL5 As Variant, varSelectedTask As Variant
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colTargetIdx As New Collection
    Dim lngFirstMissingFoundValue As Long
    Dim rngHierarchy As Range
    Dim varHierarchyArray As Variant
    Dim varLevelArray As Variant
    Dim varTaskArray As Variant
    Dim colTargetValue As New Collection
    ' 一時変数定義
    Dim r As Long, i As Long, v As Long
    Dim tmpVarIdx As Variant
    Dim tmpVarValue As Variant, tmpBlnExist As Boolean
    
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
    Set rngHierarchy = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    varHierarchyArray = rngHierarchy.Value
    ' あらかじめWBSレベル列のデータを取得
    varLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    varTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' 選択した行のレベルを取得
    intSelectedLevel = varLevelArray(lngSelectedRow - lngStartRow + 1, 1)
    ' 選択した行がタスクかどうか取得
    blnSelectedIsTask = varTaskArray(lngSelectedRow - lngStartRow + 1, 1)
    ' 選択した行の末尾の値を取得
    If blnSelectedIsTask Then
        varSelectedLastValue = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 6)
    Else
        varSelectedLastValue = varHierarchyArray(lngSelectedRow - lngStartRow + 1, intSelectedLevel)
    End If
    
    ' 選択した行のデータを取得
    varSelectedL1 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 1)
    varSelectedL2 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 2)
    varSelectedL3 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 3)
    varSelectedL4 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 4)
    varSelectedL5 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 5)
    varSelectedTask = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 6)
    
    ' 更新対象範囲列のデータを更新
    If blnSelectedIsTask = True Then
        ' # 選択行がタスクの場合 #
        ' 対象となる値をコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedLevel = 5 And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    varHierarchyArray(i, 5) = varSelectedL5 And _
                    varHierarchyArray(i, 4) = varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, 6), CStr(varHierarchyArray(i, 6))
                On Error GoTo 0
            End If
            If intSelectedLevel = 4 And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    varHierarchyArray(i, 4) = varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, 6), CStr(varHierarchyArray(i, 6))
                On Error GoTo 0
            End If
            If intSelectedLevel = 3 And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    IsEmpty(varHierarchyArray(i, 4)) And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, 6), CStr(varHierarchyArray(i, 6))
                On Error GoTo 0
            End If
            If intSelectedLevel = 2 And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    IsEmpty(varHierarchyArray(i, 4)) And _
                    IsEmpty(varHierarchyArray(i, 3)) And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, 6), CStr(varHierarchyArray(i, 6))
                On Error GoTo 0
            End If
            If intSelectedLevel = 1 And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    IsEmpty(varHierarchyArray(i, 4)) And _
                    IsEmpty(varHierarchyArray(i, 3)) And _
                    IsEmpty(varHierarchyArray(i, 2)) And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, 6), CStr(varHierarchyArray(i, 6))
                On Error GoTo 0
            End If
        Next r
        ' 値コレクションをから最初の存在しない値を取得
        lngFirstMissingFoundValue = 0
        For v = varSelectedLastValue To 1 Step -1
            tmpBlnExist = False
            For Each tmpVarValue In colTargetValue
                If v = tmpVarValue Then
                    tmpBlnExist = True
                    Exit For
                End If
            Next tmpVarValue
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
            If intSelectedLevel = 5 And _
                    varHierarchyArray(i, 6) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    varHierarchyArray(i, 5) = varSelectedL5 And _
                    varHierarchyArray(i, 4) = varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 4 And _
                    varHierarchyArray(i, 6) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    varHierarchyArray(i, 4) = varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 3 And _
                    varHierarchyArray(i, 6) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    IsEmpty(varHierarchyArray(i, 4)) And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 2 And _
                    varHierarchyArray(i, 6) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    IsEmpty(varHierarchyArray(i, 4)) And _
                    IsEmpty(varHierarchyArray(i, 3)) And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 1 And _
                    varHierarchyArray(i, 6) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 6) <= varSelectedTask And _
                    IsEmpty(varHierarchyArray(i, 5)) And _
                    IsEmpty(varHierarchyArray(i, 4)) And _
                    IsEmpty(varHierarchyArray(i, 3)) And _
                    IsEmpty(varHierarchyArray(i, 2)) And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
        Next r
        ' 対象となるデータインデックスのみ値を更新する
        For Each tmpVarIdx In colTargetIdx
            varHierarchyArray(tmpVarIdx, 6) = varHierarchyArray(tmpVarIdx, 6) - 1
        Next tmpVarIdx
    Else
        ' # 選択行がタスクでない場合 #
        ' 対象となる値をコレクションに格納
        For r = lngStartRow To lngEndRow
            ' 現在のインデックスを取得
            i = r - lngStartRow + 1
            ' 対象行か判定してコレクションに格納
            If intSelectedLevel = 5 And _
                    varHierarchyArray(i, 5) <= varSelectedL5 And _
                    varHierarchyArray(i, 4) = varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, intSelectedLevel), CStr(varHierarchyArray(i, intSelectedLevel))
                On Error GoTo 0
            End If
            If intSelectedLevel = 4 And _
                    varHierarchyArray(i, 4) <= varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, intSelectedLevel), CStr(varHierarchyArray(i, intSelectedLevel))
                On Error GoTo 0
            End If
            If intSelectedLevel = 3 And _
                    varHierarchyArray(i, 3) <= varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, intSelectedLevel), CStr(varHierarchyArray(i, intSelectedLevel))
                On Error GoTo 0
            End If
            If intSelectedLevel = 2 And _
                    varHierarchyArray(i, 2) <= varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, intSelectedLevel), CStr(varHierarchyArray(i, intSelectedLevel))
                On Error GoTo 0
            End If
            If intSelectedLevel = 1 And _
                    varHierarchyArray(i, 1) <= varSelectedL1 Then
                On Error Resume Next
                colTargetValue.Add varHierarchyArray(i, intSelectedLevel), CStr(varHierarchyArray(i, intSelectedLevel))
                On Error GoTo 0
            End If
        Next r
        ' 値コレクションをから最初の存在しない値を取得
        lngFirstMissingFoundValue = 0
        For v = varSelectedLastValue To 1 Step -1
            tmpBlnExist = False
            For Each tmpVarValue In colTargetValue
                If v = tmpVarValue Then
                    tmpBlnExist = True
                    Exit For
                End If
            Next tmpVarValue
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
            If intSelectedLevel = 5 And _
                    varHierarchyArray(i, 5) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 5) <= varSelectedL5 And _
                    varHierarchyArray(i, 4) = varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 4 And _
                    varHierarchyArray(i, 4) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 4) <= varSelectedL4 And _
                    varHierarchyArray(i, 3) = varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 3 And _
                    varHierarchyArray(i, 3) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 3) <= varSelectedL3 And _
                    varHierarchyArray(i, 2) = varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 2 And _
                    varHierarchyArray(i, 2) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 2) <= varSelectedL2 And _
                    varHierarchyArray(i, 1) = varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
            If intSelectedLevel = 1 And _
                    varHierarchyArray(i, 1) > lngFirstMissingFoundValue And _
                    varHierarchyArray(i, 1) <= varSelectedL1 Then
                colTargetIdx.Add i, CStr(i)
            End If
        Next r
        ' 対象となるデータインデックスのみ値を更新する
        For Each tmpVarIdx In colTargetIdx
            varHierarchyArray(tmpVarIdx, intSelectedLevel) = varHierarchyArray(tmpVarIdx, intSelectedLevel) - 1
        Next tmpVarIdx
    End If
    
    ' データの更新結果を反映
    rngHierarchy.Value = varHierarchyArray

End Sub


' ■ チェックした２点の最終レベルIDを交換する
Public Sub ExecSwapCheckedLastLevel(ws As Worksheet)

    ' 変数定義
    Dim lngChecked1Row As Long, intChecked1Level As Integer, blnChecked1IsTask As Boolean, varChecked1LastValue As Variant, varChecked1Id As Variant
    Dim varChecked1L1 As Variant, varChecked1L2 As Variant, varChecked1L3 As Variant, varChecked1L4 As Variant, varChecked1L5 As Variant, varChecked1Task As Variant
    Dim lngChecked2Row As Long, intChecked2Level As Integer, blnChecked2IsTask As Boolean, varChecked2LastValue As Variant, varChecked2Id As Variant
    Dim varChecked2L1 As Variant, varChecked2L2 As Variant, varChecked2L3 As Variant, varChecked2L4 As Variant, varChecked2L5 As Variant, varChecked2Task As Variant
    Dim colCheckedRows As Collection
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim rngHierarchy As Range
    Dim varHierarchyArray As Variant
    Dim varLevelArray As Variant
    Dim varTaskArray As Variant
    Dim varIdArray As Variant
    ' 一時変数定義
    Dim r As Long, i As Long, v As Long
    
    ' 開始行と終了行を取得
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' チェックされている行番号を取得
    Set colCheckedRows = GetCheckedChkMultpleRows(ws)
    
    ' ガード条件（チェックが２つでなかった場合は終了）
    If colCheckedRows.Count <> 2 Then
        MsgBox "交換したい２つをチェックしてください（CHK)。" & vbCrLf & "（" & colCheckedRows.Count & " 箇所が選択されています）", vbExclamation, "通知"
        Exit Sub
    End If
    
    ' あらかじめ更新対象範囲列のデータを取得
    Set rngHierarchy = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    varHierarchyArray = rngHierarchy.Value
    ' あらかじめWBSレベル列のデータを取得
    varLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' あらかじめWBSタスク判定列のデータを取得
    varTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    ' あらかじめWBS-ID列のデータを取得
    varIdArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_WBS_ID), ws.Cells(lngEndRow, cfg.COL_WBS_ID)).Value

    ' ■ チェック1情報収集
    lngChecked1Row = colCheckedRows.Item(1)
    ' 選択した行のレベルを取得
    intChecked1Level = varLevelArray(lngChecked1Row - lngStartRow + 1, 1)
    ' 選択した行がタスクかどうか取得
    blnChecked1IsTask = varTaskArray(lngChecked1Row - lngStartRow + 1, 1)
    ' 選択した行の末尾の値を取得
    If blnChecked1IsTask Then
        varChecked1LastValue = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 6)
    Else
        varChecked1LastValue = varHierarchyArray(lngChecked1Row - lngStartRow + 1, intChecked1Level)
    End If
    ' 選択した行のIDを取得
    varChecked1Id = varIdArray(lngChecked1Row - lngStartRow + 1, 1)
    ' 選択した行のデータを取得
    varChecked1L1 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 1)
    varChecked1L2 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 2)
    varChecked1L3 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 3)
    varChecked1L4 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 4)
    varChecked1L5 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 5)
    varChecked1Task = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 6)

    ' ■ チェック2情報収集
    lngChecked2Row = colCheckedRows.Item(2)
    ' 選択した行のレベルを取得
    intChecked2Level = varLevelArray(lngChecked2Row - lngStartRow + 1, 1)
    ' 選択した行がタスクかどうか取得
    blnChecked2IsTask = varTaskArray(lngChecked2Row - lngStartRow + 1, 1)
    ' 選択した行の末尾の値を取得
    If blnChecked2IsTask Then
        varChecked2LastValue = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 6)
    Else
        varChecked2LastValue = varHierarchyArray(lngChecked2Row - lngStartRow + 1, intChecked2Level)
    End If
    ' 選択した行のIDを取得
    varChecked2Id = varIdArray(lngChecked2Row - lngStartRow + 1, 1)
    ' 選択した行のデータを取得
    varChecked2L1 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 1)
    varChecked2L2 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 2)
    varChecked2L3 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 3)
    varChecked2L4 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 4)
    varChecked2L5 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 5)
    varChecked2Task = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 6)
    
    ' ガード条件（２つの階層及びタスクか否かが一致しない場合、終了）
    If (intChecked1Level <> intChecked2Level) Or (blnChecked1IsTask <> blnChecked2IsTask) Then
        MsgBox "交換候補の階層およびタスクかどうかが一致しません（CHK)。" & vbCrLf & _
        vbCrLf & "チェック1: 階層=" & intChecked1Level & ", タスク=" & blnChecked1IsTask & _
        vbCrLf & "チェック2: 階層=" & intChecked2Level & ", タスク=" & blnChecked2IsTask & _
        "", vbExclamation, "通知"
        Exit Sub
    End If
    
    ' ガード条件（２つの末尾番号以外の階層番号が一致しない場合、終了）
    If blnChecked1IsTask = True Then
        If varChecked1L1 <> varChecked2L1 Or varChecked1L2 <> varChecked2L2 Or varChecked1L3 <> varChecked2L3 Or varChecked1L4 <> varChecked2L4 Or varChecked1L5 <> varChecked2L5 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varChecked1Id & _
            vbCrLf & "チェック2: " & varChecked2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 5 Then
        If varChecked1L1 <> varChecked2L1 Or varChecked1L2 <> varChecked2L2 Or varChecked1L3 <> varChecked2L3 Or varChecked1L4 <> varChecked2L4 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varChecked1Id & _
            vbCrLf & "チェック2: " & varChecked2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 4 Then
        If varChecked1L1 <> varChecked2L1 Or varChecked1L2 <> varChecked2L2 Or varChecked1L3 <> varChecked2L3 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varChecked1Id & _
            vbCrLf & "チェック2: " & varChecked2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 3 Then
        If varChecked1L1 <> varChecked2L1 Or varChecked1L2 <> varChecked2L2 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varChecked1Id & _
            vbCrLf & "チェック2: " & varChecked2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 2 Then
        If varChecked1L1 <> varChecked2L1 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varChecked1Id & _
            vbCrLf & "チェック2: " & varChecked2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    End If
    
    ' 値の交換を実施
    For r = lngStartRow To lngEndRow
        
        ' 現在のインデックスを取得
        i = r - lngStartRow + 1
        
        ' 交換した値をセット
        If blnChecked1IsTask = True Then
            If varChecked1L1 = varHierarchyArray(i, 1) And _
                    varChecked1L2 = varHierarchyArray(i, 2) And _
                    varChecked1L3 = varHierarchyArray(i, 3) And _
                    varChecked1L4 = varHierarchyArray(i, 4) And _
                    varChecked1L5 = varHierarchyArray(i, 5) Then
                If varHierarchyArray(i, 6) = varChecked1Task Then
                    varHierarchyArray(i, 6) = varChecked2Task
                ElseIf varHierarchyArray(i, 6) = varChecked2Task Then
                    varHierarchyArray(i, 6) = varChecked1Task
                End If
            End If
        ElseIf intChecked1Level = 5 Then
            If varChecked1L1 = varHierarchyArray(i, 1) And _
                    varChecked1L2 = varHierarchyArray(i, 2) And _
                    varChecked1L3 = varHierarchyArray(i, 3) And _
                    varChecked1L4 = varHierarchyArray(i, 4) Then
                If varHierarchyArray(i, 5) = varChecked1L5 Then
                    varHierarchyArray(i, 5) = varChecked2L5
                ElseIf varHierarchyArray(i, 5) = varChecked2L5 Then
                    varHierarchyArray(i, 5) = varChecked1L5
                End If
            End If
        ElseIf intChecked1Level = 4 Then
            If varChecked1L1 = varHierarchyArray(i, 1) And _
                    varChecked1L2 = varHierarchyArray(i, 2) And _
                    varChecked1L3 = varHierarchyArray(i, 3) Then
                If varHierarchyArray(i, 4) = varChecked1L4 Then
                    varHierarchyArray(i, 4) = varChecked2L4
                ElseIf varHierarchyArray(i, 4) = varChecked2L4 Then
                    varHierarchyArray(i, 4) = varChecked1L4
                End If
            End If
        ElseIf intChecked1Level = 3 Then
            If varChecked1L1 = varHierarchyArray(i, 1) And _
                    varChecked1L2 = varHierarchyArray(i, 2) Then
                If varHierarchyArray(i, 3) = varChecked1L3 Then
                    varHierarchyArray(i, 3) = varChecked2L3
                ElseIf varHierarchyArray(i, 3) = varChecked2L3 Then
                    varHierarchyArray(i, 3) = varChecked1L3
                End If
            End If
        ElseIf intChecked1Level = 2 Then
            If varChecked1L1 = varHierarchyArray(i, 1) Then
                If varHierarchyArray(i, 2) = varChecked1L2 Then
                    varHierarchyArray(i, 2) = varChecked2L2
                ElseIf varHierarchyArray(i, 2) = varChecked2L2 Then
                    varHierarchyArray(i, 2) = varChecked1L2
                End If
            End If
        ElseIf intChecked1Level = 1 Then
            If varHierarchyArray(i, 1) = varChecked1L1 Then
                varHierarchyArray(i, 1) = varChecked2L1
            ElseIf varHierarchyArray(i, 1) = varChecked2L1 Then
                varHierarchyArray(i, 1) = varChecked1L1
            End If
        End If
    Next r
    
    ' 値を反映
    rngHierarchy.Value = varHierarchyArray
    
End Sub


' ■ 指定の階層配列を対象に、指定インデックスにあるデータの子階層にあたるインデックスのコレクションを取得
Private Function GetTargetChildIdxs(varHierarchyArray As Variant, _
                                        lngTargetIdx As Long) As Collection
    
    ' 変数定義
    Dim colResultIdxs As New Collection
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    ' 一時変数定義
    Dim i As Long
    
    ' ガード条件（入力されたインデックスが0以下の場合は終了）
    If lngTargetIdx <= 0 Then
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' ガード条件（入力された階層配列の列数が6でない場合は終了）
    If UBound(varHierarchyArray, 2) <> 6 Then
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' ガード条件（入力された階層配列の行数を越えたインデックスを指定された場合は終了）
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' 指定インデックスの値を取得
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' タスク状態の取得
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' レベルの取得
    If IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsNumeric(varTargetL5) And Not IsNull(varTargetL5) And Not IsEmpty(varTargetL5) Then
        intTargetLevel = 5
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsNumeric(varTargetL4) And Not IsNull(varTargetL4) And Not IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 4
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsNumeric(varTargetL3) And Not IsNull(varTargetL3) And Not IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 3
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsNumeric(varTargetL2) And Not IsNull(varTargetL2) And Not IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 2
    ElseIf IsNumeric(varTargetL1) And Not IsNull(varTargetL1) And Not IsEmpty(varTargetL1) And _
            IsEmpty(varTargetL2) And _
            IsEmpty(varTargetL3) And _
            IsEmpty(varTargetL4) And _
            IsEmpty(varTargetL5) Then
        intTargetLevel = 1
    Else
        ' # 階層に問題がある場合 #
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' ガード条件（タスクの場合は終了）
    If blnTargetTask = True Then
        ' # タスクには子階層がないため #
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' 該当するインデックスを収集
    For i = 1 To UBound(varHierarchyArray, 1)
        If intTargetLevel = 5 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                varTargetL4 = varHierarchyArray(i, 4) And _
                varTargetL5 = varHierarchyArray(i, 5) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L5の場合、L5のタスクならば追加 #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 4 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                varTargetL4 = varHierarchyArray(i, 4) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L4の場合、L4のタスクならば追加 #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 4 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                varTargetL4 = varHierarchyArray(i, 4) And _
                IsNumeric(varHierarchyArray(i, 5)) And Not IsNull(varHierarchyArray(i, 5)) And Not IsEmpty(varHierarchyArray(i, 5)) And _
                IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L4の場合、L4の子であるL5ならば追加 #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 3 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L3の場合、L3のタスクならば追加 #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 3 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                IsNumeric(varHierarchyArray(i, 4)) And Not IsNull(varHierarchyArray(i, 4)) And Not IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L3の場合、L3の子であるL4ならば追加 #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 2 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                IsEmpty(varHierarchyArray(i, 3)) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L2の場合、L2のタスクならば追加 #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 2 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                IsNumeric(varHierarchyArray(i, 3)) And Not IsNull(varHierarchyArray(i, 3)) And Not IsEmpty(varHierarchyArray(i, 3)) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L2の場合、L2の子であるL3ならば追加 #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 1 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                IsEmpty(varHierarchyArray(i, 2)) And _
                IsEmpty(varHierarchyArray(i, 3)) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L1の場合、L1のタスクならば追加 #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 1 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                IsNumeric(varHierarchyArray(i, 2)) And Not IsNull(varHierarchyArray(i, 2)) And Not IsEmpty(varHierarchyArray(i, 2)) And _
                IsEmpty(varHierarchyArray(i, 3)) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L1の場合、L1の子であるL2ならば追加 #
            colResultIdxs.Add i, CStr(i)
        End If
    Next i
    
    Set GetTargetChildIdxs = colResultIdxs
End Function



' ■ チェックした行を削除する
Public Sub ExecRemoveCheckedRows(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colCheckedRows As Collection
    Dim varHierarchyArray As Variant
    Dim varChildExistArray As Variant
    Dim rngChk As Range
    Dim varChkArray As Variant
    Dim varIdArray As Variant
    Dim colRemoveRows As New Collection
    Dim rngRemoveTarget As Range
    ' 一時変数定義
    Dim tmpVarCheckedItem As Variant
    Dim tmpVarChildIdx As Variant
    Dim i As Long
    Dim tmpColChilds As Collection
    Dim tmpVar As Variant
    Dim answer As VbMsgBoxResult

    ' 開始行と終了行を取得
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' チェックされている行番号を取得
    Set colCheckedRows = GetCheckedChkMultpleRows(ws)
    
    ' あらかじめチェック対象範囲列のデータを取得
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' あらかじめWBS子有無判定列のデータを取得
    varChildExistArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_CE), ws.Cells(lngEndRow, cfg.COL_FLG_CE)).Value
    ' あらかじめチェック列のデータを取得
    Set rngChk = ws.Range(ws.Cells(lngStartRow, cfg.COL_CHK), ws.Cells(lngEndRow, cfg.COL_CHK))
    varChkArray = rngChk.Value
    ' あらかじめWBS-ID列のデータを取得
    varIdArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_WBS_ID), ws.Cells(lngEndRow, cfg.COL_WBS_ID)).Value
    
    ' チェックされた行ごとに削除可能かチェックを実施
    For Each tmpVarCheckedItem In colCheckedRows
        ' 現在のインデックスを取得
        i = tmpVarCheckedItem - lngStartRow + 1
        ' 子があるかどうか
        If varChildExistArray(i, 1) Then
            ' # 子が存在する場合 #
            Set tmpColChilds = GetTargetChildIdxs(varHierarchyArray, i)
            For Each tmpVarChildIdx In tmpColChilds
                tmpVar = varChildExistArray(tmpVarChildIdx, 1)
                If tmpVar = True Then
                    ' # 孫が存在する場合 #
                    MsgBox "孫階層が存在するため削除できません。" & vbCrLf & _
                    "", vbExclamation, "通知"
                    Exit Sub
                Else
                    ' # 孫が存在しない場合 #
                    On Error Resume Next
                    colRemoveRows.Add (tmpVarChildIdx + lngStartRow - 1), CStr(tmpVarChildIdx + lngStartRow - 1)
                    On Error GoTo 0
                End If
            Next tmpVarChildIdx
            On Error Resume Next
            colRemoveRows.Add tmpVarCheckedItem, CStr(tmpVarCheckedItem)
            On Error GoTo 0
        Else
            ' # 子が存在しない場合 #
            On Error Resume Next
            colRemoveRows.Add tmpVarCheckedItem, CStr(tmpVarCheckedItem)
            On Error GoTo 0
        End If
    Next tmpVarCheckedItem
    
    ' チェック列を更新する
    For Each tmpVar In colRemoveRows
        varChkArray(tmpVar - lngStartRow + 1, 1) = cfg.CHK_MARK_T
    Next tmpVar
    rngChk.Value = varChkArray
    
    ' 一時的に描画を再開
    If Application.ScreenUpdating = False And Application.EnableEvents = False Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    End If
    
    ' 確認の上、削除を実行
    answer = MsgBox("本当に削除してもよいですか？", vbOKCancel + vbQuestion, "確認")
    If answer = vbOK Then
        ' 削除対象範囲を用意
        For Each tmpVar In colRemoveRows
            If rngRemoveTarget Is Nothing Then
                Set rngRemoveTarget = Rows(tmpVar)
            Else
                Set rngRemoveTarget = Union(rngRemoveTarget, Rows(tmpVar))
            End If
        Next tmpVar
        ' 一括削除を実行
        If Not rngRemoveTarget Is Nothing Then rngRemoveTarget.Delete
    End If

End Sub


' ■ 基本数式を数値に変換する
Public Sub ExecConvertBasicFormulasToValues(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long
    Dim tmpRange As Range
    Dim tmpVariant As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ■ 全行にアクセスが必要なコストの高い数式を数値に変換
    ' WBS_CNTの式→値
    Set tmpRange = ws.Range(cfg.COL_WBS_CNT_LABEL & lngStartRow & ":" & cfg.COL_WBS_CNT_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.Value = tmpVariant
    
    ' FLG_PEの式→値
    Set tmpRange = ws.Range(cfg.COL_FLG_PE_LABEL & lngStartRow & ":" & cfg.COL_FLG_PE_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.Value = tmpVariant
    
    ' FLG_CEの式→値
    Set tmpRange = ws.Range(cfg.COL_FLG_CE_LABEL & lngStartRow & ":" & cfg.COL_FLG_CE_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.Value = tmpVariant

End Sub


' ■ 集計数式を数値に変換する
Public Sub ExecConvertAggregateFormulasToValues(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long
    Dim tmpRange As Range
    Dim tmpVariant As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
        
    ' タスク集計合計の式→値
    Set tmpRange = ws.Range(cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
    
    ' タスク集計完了の式→値
    Set tmpRange = ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
    
    ' 工数進捗率の式→値
    Set tmpRange = ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "0.0%"
    tmpRange.Value = tmpVariant
    
    ' 項目消化率の式→値
    Set tmpRange = ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "0.0%"
    tmpRange.Value = tmpVariant
    
    ' 予定工数の式→値
    Set tmpRange = ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
    
    ' 実績残工数の式→値
    Set tmpRange = ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
    
    ' 実績済工数の式→値
    Set tmpRange = ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
           
    ' 特定のセルの式を値に変換
    ws.Range(cfg.COL_TASK_COUNT_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_TASK_COUNT_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_TASK_PROG_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_TASK_PROG_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow + 2).Value

End Sub


' ■ カスタムフォーマット関数（WBS-IDX用）
Function CustomFormatWbsIdx(varB As Variant, _
                                varE As Variant, _
                                varF As Variant, _
                                varG As Variant, _
                                varH As Variant, _
                                varI As Variant, _
                                varJ As Variant) As String
    
    ' 変数定義
    Dim strResult As String
    Dim varValues As Variant
    ' 一時変数定義
    Dim parts(0 To 5) As String
    Dim i As Integer

    ' もしvarBが"E"なら"ERROR"を返す
    If varB = "E" Then
        CustomFormat = "ERROR"
        Exit Function
    End If

    ' もしvarEが空なら固定文字列を返す
    If varE = "" Then
        CustomFormatWbsIdx = "XXX.XXX.XXX.XXX.XXX.XXX"
        Exit Function
    End If

    ' 各値を配列にまとめる
    varValues = Array(varE, varF, varG, varH, varI, varJ)

    ' 各要素をループして処理
    For i = 0 To 5
        If varValues(i) = "" Then
            parts(i) = "---"
        Else
            parts(i) = Format(varValues(i), "000")
        End If
    Next i

    ' 結合して結果を作成
    strResult = parts(0) & "." & parts(1) & "." & parts(2) & "." & parts(3) & "." & parts(4) & "." & parts(5)

    CustomFormatWbsIdx = strResult
End Function


' ■ カスタムフォーマット関数（WBS-ID用）
Function CustomFormatWbsId(varB As Variant, _
                            varE As Variant, _
                            varF As Variant, _
                            varG As Variant, _
                            varH As Variant, _
                            varI As Variant, _
                            varJ As Variant) As String
    
    ' 変数定義
    Dim strResult As String

    ' もしvarBが"E"なら"ERROR"を返す
    If varB = "E" Then
        CustomFormatWbsId = "ERROR"
        Exit Function
    End If

    ' もしvarEが空なら空文字を返す
    If varE = "" Then
        CustomFormatWbsId = ""
        Exit Function
    End If

    ' 連結処理
    strResult = varE

    If varF <> "" Then strResult = strResult & "." & varF
    If varG <> "" Then strResult = strResult & "." & varG
    If varH <> "" Then strResult = strResult & "." & varH
    If varI <> "" Then strResult = strResult & "." & varI
    If varJ <> "" Then strResult = strResult & ".T" & varJ

    CustomFormatWbsId = strResult
End Function


' ■ カスタム関数（LEVEL）
Function CustomFuncGetLevel(varE As Variant, _
                                varF As Variant, _
                                varG As Variant, _
                                varH As Variant, _
                                varI As Variant) As Integer
    
    ' デフォルトは0
    CustomFuncGetLevel = 0
    
    ' 順番にチェックしていく
    If IsNumeric(varE) And Not IsEmpty(varE) And Not IsNull(varE) Then
        If varF = "" Then
            CustomFuncGetLevel = 1
        ElseIf IsNumeric(varF) Then
            If varG = "" Then
                CustomFuncGetLevel = 2
            ElseIf IsNumeric(varG) Then
                If varH = "" Then
                    CustomFuncGetLevel = 3
                ElseIf IsNumeric(varH) Then
                    If varI = "" Then
                        CustomFuncGetLevel = 4
                    ElseIf IsNumeric(varI) Then
                        CustomFuncGetLevel = 5
                    End If
                End If
            End If
        End If
    End If
End Function


' ■ 制御に使用する列に数式をまとめてセット
Public Sub SetFormulaToControlColumn(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    ' 一時変数を定義
    Dim i As Long, j As Long
    Dim tmpLngRow As Long

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 数式をセットするデータを用意
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To cfg.COL_WBS_ID - cfg.COL_WBS_IDX + 1)

    ' 数式をセット
    For i = 1 To cfg.COL_WBS_ID - cfg.COL_WBS_IDX + 1
        For j = 1 To lngEndRow - lngStartRow + 1
            tmpLngRow = lngStartRow + j - 1
            Select Case i
                Case 1
                    ' # WBS_IDX #
                    varFormulas(j, i) = "=CustomFormatWbsIdx(" & _
                                            cfg.COL_ERR_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L1_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L2_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L3_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L4_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L5_LABEL & tmpLngRow & "," & _
                                            cfg.COL_TASK_LABEL & tmpLngRow & ")"
                Case 2
                    ' # WBS_CNT #
                    varFormulas(j, i) = "=COUNTIF(" & _
                                            cfg.COL_WBS_IDX_LABEL & "$" & lngStartRow & ":" & _
                                            cfg.COL_WBS_IDX_LABEL & "$" & lngEndRow & "," & _
                                            cfg.COL_WBS_IDX_LABEL & tmpLngRow & ")"
                Case 3
                    ' # LEVEL #
                    varFormulas(j, i) = "=CustomFuncGetLevel(" & _
                                            cfg.COL_L1_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L2_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L3_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L4_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L5_LABEL & tmpLngRow & ")"
                Case 4
                    ' # FLG_T #
                    varFormulas(j, i) = "=AND(" & _
                                            cfg.COL_TASK_LABEL & tmpLngRow & "<>"""",ISNUMBER(" & _
                                            cfg.COL_TASK_LABEL & tmpLngRow & "))"
                Case 5
                    ' # FLG_IC #
                    varFormulas(j, i) = "=NOT(OR(" & _
                                            cfg.COL_WBS_STATUS_LABEL & tmpLngRow & "=""" & cfg.WBS_STATUS_DELETED & """," & _
                                            cfg.COL_WBS_STATUS_LABEL & tmpLngRow & "=""" & cfg.WBS_STATUS_TRANSFERRED & """," & _
                                            cfg.COL_WBS_STATUS_LABEL & tmpLngRow & "=""" & cfg.WBS_STATUS_SHELVED & """," & _
                                            cfg.COL_WBS_STATUS_LABEL & tmpLngRow & "=""" & cfg.WBS_STATUS_REJECTED & """" & "))"
                Case 6
                    ' # FLG_PE #
                    varFormulas(j, i) = "=AND(" & _
                                            cfg.COL_LEVEL_LABEL & tmpLngRow & ">0," & _
                                            cfg.COL_WBS_ID_LABEL & tmpLngRow & "<>"""",IFERROR(ISNUMBER(MATCH(IFERROR(LEFT(" & _
                                            cfg.COL_WBS_ID_LABEL & tmpLngRow & ",FIND(""~"",SUBSTITUTE(" & _
                                            cfg.COL_WBS_ID_LABEL & tmpLngRow & ",""."",""~"",LEN(" & _
                                            cfg.COL_WBS_ID_LABEL & tmpLngRow & ")-LEN(SUBSTITUTE(" & _
                                            cfg.COL_WBS_ID_LABEL & tmpLngRow & ",""."",""""))))-1)," & _
                                            cfg.COL_WBS_ID_LABEL & tmpLngRow & ")," & _
                                            cfg.COL_WBS_ID_LABEL & "$" & lngStartRow & ":" & _
                                            cfg.COL_WBS_ID_LABEL & "$" & lngEndRow & _
                                            ",0)),FALSE))"
                Case 7
                    ' # FLG_CE #
                    varFormulas(j, i) = "=AND(" & _
                                            cfg.COL_LEVEL_LABEL & tmpLngRow & ">0," & _
                                            cfg.COL_FLG_T_LABEL & tmpLngRow & "=FALSE," & _
                                            cfg.COL_WBS_ID_LABEL & tmpLngRow & "<>"""",IFERROR(SUMPRODUCT(--(LEFT(" & _
                                            cfg.COL_WBS_ID_LABEL & "$" & lngStartRow & ":" & _
                                            cfg.COL_WBS_ID_LABEL & "$" & lngEndRow & ",LEN(" & _
                                            cfg.COL_WBS_ID_LABEL & tmpLngRow & "&"".""))=" & _
                                            cfg.COL_WBS_ID_LABEL & tmpLngRow & "&"".""))>0,FALSE))"

                Case 8
                    ' # WBS_ID #
                    varFormulas(j, i) = "=CustomFormatWbsId(" & _
                                            cfg.COL_ERR_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L1_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L2_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L3_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L4_LABEL & tmpLngRow & "," & _
                                            cfg.COL_L5_LABEL & tmpLngRow & "," & _
                                            cfg.COL_TASK_LABEL & tmpLngRow & ")"
            End Select
        Next j
    Next i

    ' 一括で対象範囲に対し処理を行う
    With ws.Range(cfg.COL_WBS_IDX_LABEL & lngStartRow & ":" & cfg.COL_WBS_ID_LABEL & lngEndRow)
        ' 書式を一括で設定
        .NumberFormat = "General"
        ' 式をセット
        .Formula = varFormulas
    End With

End Sub

