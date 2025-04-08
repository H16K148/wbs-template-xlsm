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

' ■ 定数定義

Private Const strCHK_T As String = "■"
Private Const strCHK_F As String = "・"
Private Const strOPT_T As String = "●"
Private Const strOPT_F As String = "・"

' 列定義
Private Const lngCOL_KEY  As Long = 1                  '  A列：KEY カラム
Private Const lngCOL_ERR  As Long = 2                  '  B列：ERROR 情報が出力されるカラム
Private Const lngCOL_CHK  As Long = 3                  '  C列：チェックボックスのカラム
Private Const lngCOL_OPT  As Long = 4                  '  D列：オプションボックスのカラム
Private Const lngCOL_L1   As Long = 5                  '  E列：L1 番号を入力するカラム
Private Const lngCOL_L2   As Long = 6                  '  F列：L2 番号を入力するカラム
Private Const lngCOL_L3   As Long = 7                  '  G列：L3 番号を入力するカラム
Private Const lngCOL_L4   As Long = 8                  '  H列：L4 番号を入力するカラム
Private Const lngCOL_L5   As Long = 9                  '  I列：L5 番号を入力するカラム
Private Const lngCOL_TASK As Long = 10                 '  J列：TASK 番号を入力するカラム
Private Const lngCOL_WBS_IDX As Long = 11              '  K列：WBS_IDXを表示するカラム（普段は非表示）
Private Const lngCOL_WBS_ID As Long = 12               '  L列：WBS_IDを表示するカラム
Private Const lngCOL_TASK_COUNT As Long = 20           '  T列：タスク数のカラム
Private Const lngCOL_EFFORT_PROG As Long = 21          '  U列：工数進捗率のカラム
Private Const lngCOL_TASK_PROG As Long = 22            '  V列：項目消化率のカラム
Private Const lngCOL_TASK_WGT As Long = 23             '  W列：項目加重のカラム
Private Const lngCOL_WBS_STATUS As Long = 24           '  X列：WBSステータスのカラム
Private Const lngCOL_TEAM_SLCT As Long = 25            '  Y列：組織選択のカラム
Private Const lngCOL_PERSON_SLCT As Long = 26          '  Z列：担当選択のカラム
Private Const lngCOL_PLANNED_EFF As Long = 28          ' AB列：予定工数のカラム
Private Const lngCOL_ACTUAL_REMAINING_EFF As Long = 31 ' AE列：実績残工数のカラム
Private Const lngCOL_ACTUAL_COMPLETED_EFF As Long = 32 ' AF列：実績済工数のカラム
Private Const lngCOL_LAST As Long = 35                 ' AI列：（データの）最終カラム
Private Const lngCOL_DEFINE_TEAM As Long = 37          ' AK列：組織を定義するカラム
Private Const lngCOL_DEFINE_PERSON As Long = 38        ' AL列：担当を定義するカラム

' 行定義
Private Const lngROW_CTRL1 As Long = 3            ' コントロールを配置する行

' コントロール名関連定義
Private Const strNAME_EXE_COMBOBOX1 As String = "ExecuteComboBox_1"         '
Private Const strNAME_EXE_BUTTON1   As String = "ExecuteButton_1"           '
Private Const strNAME_RESET_BUTTON As String = "ResetButton_1"              '
Private Const strNAME_RECALCULATE_BUTTON As String = "RecalculateButton_1"  '

' コントロール関連定義
Private Const dblSHAPE_MARGIN As Double = 0.2
Private Const intEXE_COMBOBOX_WIDTH = 500
Private Const intRESET_BUTTON_WIDTH = 55
Private Const intRECALCULATE_BUTTON_WIDTH = 55


' ■ ユーティリティ：コレクション
Function ExistsColKey(objCol As Collection, strKey As String) As Boolean
     
    ' 戻り値の初期値：False
    ExistsColKey = False
     
    ' 変数にCollection未設定の場合は処理終了
    If objCol Is Nothing Then Exit Function
     
    ' Collectionのメンバー数が「0」の場合は処理終了
    If objCol.Count = 0 Then Exit Function
    
    ' エラー設定の変更
    On Error Resume Next
     
    ' Itemメソッドを実行
    Call objCol.Item(strKey)
         
    ' エラー値がない場合：キー検索はヒット（戻り値：True）
    If Err.Number = 0 Then ExistsColKey = True
    
    ' 処理の終了とともに、エラー設定が元に戻る
 
End Function


' ■ ユーティリティ：コレクション
Function ExistsColItem(objCol As Collection, varItem As Variant) As Boolean
    
    ' 変数定義
    Dim v As Variant
     
    ' 戻り値の初期値：False
    ExistsColItem = False
     
    ' 変数にCollection未設定の場合は処理終了
    If objCol Is Nothing Then Exit Function
     
    ' Collectionのメンバー数が「0」の場合は処理終了
    If objCol.Count = 0 Then Exit Function
     
    ' Collectionの各メンバーと突合
    For Each v In objCol
         
        ' 突合結果が一致した場合：戻り値「True」にループ抜け
        If v = varItem Then ExistsColItem = True: Exit For
         
    Next
     
End Function


' ■ ユーティリティ：コレクション
Function ExistsColItemCount(objCol As Collection, varItem As Variant) As Long

    ' 変数定義
    Dim v As Variant
    Dim lngMatchCount As Long ' 一致した数をカウントする変数

    ' 戻り値の初期化
    ExistsColItemCount = 0

    ' 変数にCollection未設定の場合は処理終了
    If objCol Is Nothing Then Exit Function

    ' Collectionのメンバー数が「0」の場合は処理終了
    If objCol.Count = 0 Then Exit Function

    ' Collectionの各メンバーと突合
    For Each v In objCol
        ' 突合結果が一致した場合：カウントをインクリメント
        If v = varItem Then
            lngMatchCount = lngMatchCount + 1
        End If
    Next

    ' 一致した数を関数の戻り値として設定
    ExistsColItemCount = lngMatchCount

End Function


' ◆ CHK と OPT のダブルクリックイベントを実装
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    ' 変数定義
    Dim lngClickedColumn As Long
    
    ' 列番号を取得
    lngClickedColumn = Target.Column
    
    ' ガード条件（対象列以外はデフォルト動作のままで終了）
    If lngClickedColumn <> lngCOL_CHK And lngClickedColumn <> lngCOL_OPT Then
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
    lngClickedRow = Target.Row
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ガード条件（行番号が指定範囲外の場合は終了）
    If lngClickedRow < lngStartRow Or lngClickedRow > lngEndRow Then
        Exit Sub
    End If
    
    ' CHK 処理
    If lngClickedColumn = lngCOL_CHK Then
        varClicked = ws.Cells(lngClickedRow, lngCOL_CHK).value
        
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
            
        If varClicked = strCHK_T Then
            ws.Cells(lngClickedRow, lngCOL_CHK).value = strCHK_F
        Else
            ws.Cells(lngClickedRow, lngCOL_CHK).value = strCHK_T
        End If
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        
        Cancel = True
        Exit Sub
    End If
    
    ' OPT 処理 (高速化版)
    If lngClickedColumn = lngCOL_OPT Then
        ' クリックした値を取得
        varClicked = ws.Cells(lngClickedRow, lngCOL_OPT).value
        If varClicked <> strOPT_T Then ' クリックされたセルが strOPT_T でない場合のみ処理
            ' lngStartRow から lngEndRow の範囲で strOPT_T を持つ最初のセルを検索
            On Error Resume Next
            Set rngFoundCell = ws.Range(ws.Cells(lngStartRow, lngCOL_OPT), ws.Cells(lngEndRow, lngCOL_OPT)).Find( _
                What:=strOPT_T, _
                LookAt:=xlWhole, _
                LookIn:=xlValues, _
                MatchCase:=True _
            )
            On Error GoTo 0
            
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' strOPT_T を持つセルが見つかったら strOPT_F に変更
            If Not rngFoundCell Is Nothing Then
                rngFoundCell.value = strOPT_F
            End If

            ' クリックされたセルの値を strOPT_T に変更
            ws.Cells(lngClickedRow, lngCOL_OPT).value = strOPT_T
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            
            Cancel = True
            Exit Sub
        End If
    End If
    
End Sub


' ■ 列番号を列文字列に変換する
Private Function ConvertColNumberToLetter(lngColNum As Long) As String

    ' 変数定義
    Dim lngDiv As Long
    Dim lngModVal As Long
    Dim strColLetter As String

    ' 初期化
    lngDiv = lngColNum
    strColLetter = ""

    ' 計算
    Do While lngDiv > 0
        lngModVal = (lngDiv - 1) Mod 26
        strColLetter = Chr(65 + lngModVal) & strColLetter
        lngDiv = Int((lngDiv - lngModVal) / 26)
    Loop

    ConvertColNumberToLetter = strColLetter
End Function


' ■ データ行範囲を取得する (高速化版)
Private Function FindDataRangeRows() As Variant

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    Dim startCell As Range
    Dim endCell As Range
    Dim lngStartRow As Long
    Dim lngEndRow As Long

    ' 1. KEY列で "@" を持つ最初のセルを高速に検索
    On Error Resume Next
    Set startCell = ws.Columns(lngCOL_KEY).Find(What:="@", LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    On Error GoTo 0

    If startCell Is Nothing Then
        ' "@" が見つからない場合は、開始行を 0 として処理を終える
        FindDataRangeRows = Array(0, 0)
        
        MsgBox "KEY列（" & ConvertColNumberToLetter(lngCOL_KEY) & "）の開始行マーカー「@」が見つかりません。" & vbCrLf & _
               "（KEY列が非表示となっている場合は表示状態にしてください）", vbExclamation, "通知"
        
        Exit Function
    Else
        lngStartRow = startCell.Row + 1 ' 実際のデータ開始行は "@" の次の行
    End If

    ' 2. KEY列で "$" を持つ最初のセルを "@" の次の行から高速に検索
    If lngStartRow > 1 Then ' "@" が見つかった場合のみ検索
        On Error Resume Next
        Set endCell = ws.Columns(lngCOL_KEY).Find(What:="$", LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True, After:=ws.Cells(lngStartRow - 1, lngCOL_KEY))
        On Error GoTo 0
        
        If endCell Is Nothing Then
            ' "$" が見つからない場合は、最終行をシートの最終行とするか、特定の値にするか検討
            lngEndRow = ws.Cells(ws.Rows.Count, lngCOL_KEY).End(xlUp).Row - 1 ' "$" がなくても最後まで
            
            MsgBox "KEY列（" & ConvertColNumberToLetter(lngCOL_KEY) & "）の終了行マーカー「$」が見つかりません。" & vbCrLf & _
                   "（KEY列が非表示となっている場合は表示状態にしてください）", vbExclamation, "通知"
        Else
            lngEndRow = endCell.Row - 1 ' 実際のデータ終了行は "$" の前の行
        End If
    Else
        ' "@" が最初の行にある場合などの処理
        lngEndRow = ws.Cells(ws.Rows.Count, lngCOL_KEY).End(xlUp).Row - 1
    End If

    ' 結果を配列で返す
    FindDataRangeRows = Array(lngStartRow, lngEndRow)

End Function


' ■ エラーが存在するかチェック（B列をチェック）
Public Function HasErrors(Optional ByVal blnCheckOnly As Boolean = False) As Boolean

    ' チェック専用の場合、即座にリターン
    If blnCheckOnly Then
        HasErrors = False
        Exit Function
    End If

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim lngLastRow As Long
    Dim varArray As Variant

    ' ERR 列の最終行を取得
    lngLastRow = ws.Cells(ws.Rows.Count, lngCOL_ERR).End(xlUp).Row

    ' B列の値を一括取得（配列として）
    varArray = ws.Range(ws.Cells(1, lngCOL_ERR), ws.Cells(lngLastRow, lngCOL_ERR)).value

    ' 配列をループして "E" を探す
    Dim i As Long
    For i = 1 To UBound(varArray, 1)
        If varArray(i, 1) = "E" Then
            HasErrors = True
            Exit Function
        End If
    Next i

    HasErrors = False
End Function


' ■ エラーチェックを実施
Private Sub CheckErrors()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
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
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' ERR 列の E をクリア
    ws.Range(ws.Cells(lngStartRow, lngCOL_ERR), ws.Cells(lngEndRow, lngCOL_ERR)).ClearContents
    With ws.Range(ws.Cells(lngStartRow, lngCOL_ERR), ws.Cells(lngEndRow, lngCOL_ERR))
        For Each tmpCell In .Cells
            If Not tmpCell.Comment Is Nothing Then
                tmpCell.Comment.Delete
            End If
        Next tmpCell
    End With
    
    ' 指定範囲のデータを一度に取得
    varData = ws.Range(ws.Cells(lngStartRow, lngCOL_L1), ws.Cells(lngEndRow, lngCOL_TASK)).value

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
            tmpCol = c + lngCOL_OPT
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
                        colError1Message.Add "・階層番号に問題（" & ConvertColNumberToLetter(tmpCol - 1) & tmpRow & "が数値ではない）" & vbCrLf, tmpRowIdx
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
        If ExistsColKey(colError1Count, tmpRowIdx) = True Then
            tmpErrorCount = colError1Count.Item(tmpRowIdx)
        End If
        ' まだエラーが発生していない行で、WbsId が登録されているもののみ検査
        If tmpErrorCount = 0 And ExistsColKey(colWbsId, tmpRowIdx) Then
            tmpWbsId = colWbsId.Item(tmpRowIdx)
            ' 空文字列となっているWbsIdを除外して、エラーチェックを行う
            If tmpWbsId <> "" Then
                ' WbsId の数を取得して、1件以上ならエラーとする
                If ExistsColItemCount(colWbsId, tmpWbsId) > 1 Then
                    blnHasError = True
                    colError2Count.Add 1, tmpRowIdx
                    colError2Message.Add "・同一階層番号が存在（Row=" & r & "）" & vbCrLf, tmpRowIdx
                End If
                ' L1 以外（WbsId のドットがある）で、親階層WbsIdがない場合エラーとする
                tmpDotPosition = InStrRev(tmpWbsId, ".")
                If tmpDotPosition > 0 And ExistsColKey(colParentWbsId, tmpWbsId) = True Then
                    tmpParentWbsId = colParentWbsId.Item(tmpWbsId)
                    If ExistsColItem(colWbsId, tmpParentWbsId) = False Then
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
            If ExistsColKey(colError1Count, tmpRowIdx) = True Then
                tmpErrorCount = tmpErrorCount + colError1Count.Item(tmpRowIdx)
                If ExistsColKey(colError1Message, tmpRowIdx) = True Then
                    tmpErrorMessage = tmpErrorMessage & colError1Message.Item(tmpRowIdx)
                End If
            End If
            If ExistsColKey(colError2Count, tmpRowIdx) = True Then
                tmpErrorCount = tmpErrorCount + colError2Count.Item(tmpRowIdx)
                If ExistsColKey(colError2Message, tmpRowIdx) = True Then
                    tmpErrorMessage = tmpErrorMessage & colError2Message.Item(tmpRowIdx)
                End If
            End If
            If ExistsColKey(colError3Count, tmpRowIdx) = True Then
                tmpErrorCount = tmpErrorCount + colError3Count.Item(tmpRowIdx)
                If ExistsColKey(colError3Message, tmpRowIdx) = True Then
                    tmpErrorMessage = tmpErrorMessage & colError3Message.Item(tmpRowIdx)
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
Private Sub SetFormulaForWbsIdx()
    
    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String
    Dim strTargetColLetter As String
    Dim rngTarget As Range

    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' lngCOL_WBS_IDX をアルファベットに変換（ループ外で一度だけ実行）
    strTargetColLetter = ConvertColNumberToLetter(lngCOL_WBS_IDX)

    ' 数式を作成
    strFormula = "=IF(B" & lngStartRow & "=""E"",""ERROR""," & _
              "IF(E" & lngStartRow & "="""",""XXX.XXX.XXX.XXX.XXX.XXX"", CONCAT(" & _
              "TEXT(E" & lngStartRow & ",""000"")," & _
              "IF(F" & lngStartRow & "="""","".---"", ""."" & TEXT(F" & lngStartRow & ",""000""))," & _
              "IF(G" & lngStartRow & "="""","".---"", ""."" & TEXT(G" & lngStartRow & ",""000""))," & _
              "IF(H" & lngStartRow & "="""","".---"", ""."" & TEXT(H" & lngStartRow & ",""000""))," & _
              "IF(I" & lngStartRow & "="""","".---"", ""."" & TEXT(I" & lngStartRow & ",""000""))," & _
              "IF(J" & lngStartRow & "="""","".---"", ""."" & TEXT(J" & lngStartRow & ",""000"")))))"

    ' 一括で対象範囲を取得
    Set rngTarget = ws.Range(strTargetColLetter & lngStartRow & ":" & strTargetColLetter & lngEndRow)
    
    ' 数値書式を一括で設定
    rngTarget.NumberFormat = "General"
    
    ' 数式をセット
    rngTarget.Formula = strFormula
    
End Sub


' ■ ID表示用カラムに数式をセット
Private Sub SetFormulaForWbsId()
   
    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String
    Dim strTargetColLetter As String
    Dim rngTarget As Range

    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' lngCOL_WBS_IDX をアルファベットに変換（ループ外で一度だけ実行）
    strTargetColLetter = ConvertColNumberToLetter(lngCOL_WBS_ID)

    ' 数式を作成
    strFormula = "=IF(B" & lngStartRow & "=""E"",""ERROR""," & _
              "IF(E" & lngStartRow & "="""","""",CONCAT(" & _
              "E" & lngStartRow & "," & _
              "IF(F" & lngStartRow & "="""","""","".""&F" & lngStartRow & " ), " & _
              "IF(G" & lngStartRow & "="""","""","".""&G" & lngStartRow & " ), " & _
              "IF(H" & lngStartRow & "="""","""","".""&H" & lngStartRow & " ), " & _
              "IF(I" & lngStartRow & "="""","""","".""&I" & lngStartRow & " ), " & _
              "IF(J" & lngStartRow & "="""","""","".T""&J" & lngStartRow & " ))))"

    ' 一括で対象範囲を取得
    Set rngTarget = ws.Range(strTargetColLetter & lngStartRow & ":" & strTargetColLetter & lngEndRow)

    ' 数値書式を一括で設定
    rngTarget.NumberFormat = "General"
    
    ' 数式をセット
    rngTarget.Formula = strFormula
    
End Sub


' ■ データ範囲をソートする
Private Sub SortWbsRange()

    ' ワークシート取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim rngSortTarget As Range
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long

    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' ソート列に値をセット
    SetFormulaForWbsIdx
    SetFormulaForWbsId

    ' エラー列～最終列の範囲を指定（startRow～endRow）
    Set rngSortTarget = ws.Range(ws.Cells(lngStartRow, ConvertColNumberToLetter(lngCOL_ERR)), ws.Cells(lngEndRow, ConvertColNumberToLetter(lngCOL_LAST)))

    ' WBSインデックス列をキーとして昇順にソート
    rngSortTarget.Sort Key1:=ws.Range(ConvertColNumberToLetter(lngCOL_WBS_IDX) & lngStartRow), Order1:=xlAscending, Header:=xlNo

End Sub


' ■ 処理実行準備
Public Sub PrepareForProcessing(Optional ByVal blnCheckOnly As Boolean = False)

    ' チェック専用の場合、即座にリターン
    If blnCheckOnly Then Exit Sub

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' タイトルを更新
    ws.Cells(2, 2).value = ws.Name
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colControlNames As Collection
    Dim lngRowCount As Long
    Dim varChkArray() As Variant
    Dim varOptArray() As Variant
    Dim varWgtArray() As Variant
    Dim varStatusArray() As Variant
    Dim varReadWgtArray As Variant
    Dim varReadStatusArray As Variant
    ' - 画面配置コントロール
    ' Dim shpGroupBox As Shape
    Dim shpExeComboBox As Shape
    Dim shpExeButton As Shape
    Dim shpResetButton As Shape
    Dim shpRecalculateButton As Shape
    ' - 実行コンボボックスの位置とサイズを計算するための変数
    Dim dblExeComboBoxLeft As Double
    Dim dblExeComboBoxTop As Double
    Dim dblExeComboBoxWidth As Double
    Dim dblExeComboBoxHeight As Double
    ' - 実行ボタンの位置とサイズを計算するための変数
    Dim dblExeButtonLeft As Double
    Dim dblExeButtonTop As Double
    Dim dblExeButtonWidth As Double
    Dim dblExeButtonHeight As Double
    ' - リセットボタンの位置とサイズを計算するための変数
    Dim dblResetButtonLeft As Double
    Dim dblResetButtonTop As Double
    Dim dblResetButtonWidth As Double
    Dim dblResetButtonHeight As Double
    ' - 再計算ボタンの位置とサイズを計算するための変数
    Dim dblRecalculateButtonLeft As Double
    Dim dblRecalculateButtonTop As Double
    Dim dblRecalculateButtonWidth As Double
    Dim dblRecalculateButtonHeight As Double
    ' 一時変数定義
    Dim r As Long
    Dim chkBox As Shape, optBtn As Shape
    Dim shp As Shape
    
    ' コントロール名を格納するコレクションを初期化
    Set colControlNames = New Collection

    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' 実行コンボボックスの位置とサイズを計算
    dblExeComboBoxLeft = ws.Cells(lngROW_CTRL1, lngCOL_L1).Left + dblSHAPE_MARGIN
    dblExeComboBoxTop = ws.Cells(lngROW_CTRL1, lngCOL_L1).Top + dblSHAPE_MARGIN
    dblExeComboBoxWidth = intEXE_COMBOBOX_WIDTH
    dblExeComboBoxHeight = ws.Cells(lngROW_CTRL1, lngCOL_L1).Height - (dblSHAPE_MARGIN * 2)
    ' 同じ名前の実行コンボボックスが存在するか確認
    On Error Resume Next
    Set shpExeComboBox = ws.Shapes(strNAME_EXE_COMBOBOX1)
    On Error GoTo 0
    ' 実行コンボボックスが存在する場合、削除
    If Not shpExeComboBox Is Nothing Then
        shpExeComboBox.Delete
    End If
    ' 実行コンボボックスを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpExeComboBox = ws.Shapes.AddFormControl(xlDropDown, dblExeComboBoxLeft, dblExeComboBoxTop, dblExeComboBoxWidth, dblExeComboBoxHeight)
    shpExeComboBox.Name = strNAME_EXE_COMBOBOX1
    With shpExeComboBox.ControlFormat
        .AddItem "【OPT】 選択した行の下に一行追加"
        .AddItem "【OPT】 選択した階層番号の末尾を＋１"
        .AddItem "【OPT】 選択した階層番号の末尾を－１"
        .AddItem "【CHK】 チェックした２箇所の階層番号の末尾番号を交換　※ チェックが同階層である階層行の２箇所でなかった場合は不可 ※"
        .AddItem "【CHK】 チェックした行を削除　※ 子階層や子タスクがある場合は不可 ※"
        .AddItem "階層番号をソート"
        .AddItem "L2以下を非表示"
        .AddItem "L3以下を非表示"
        .AddItem "L4以下を非表示"
        .AddItem "すべてを再表示"
    End With
    With ws.DropDowns(strNAME_EXE_COMBOBOX1)
        .ListIndex = 1
    End With
    colControlNames.Add shpExeComboBox.Name
    
    ' 実行ボタンの位置とサイズを計算
    dblExeButtonLeft = ws.Cells(lngROW_CTRL1, lngCOL_CHK).Left + dblSHAPE_MARGIN
    dblExeButtonTop = ws.Cells(lngROW_CTRL1, lngCOL_CHK).Top + dblSHAPE_MARGIN
    dblExeButtonWidth = ws.Cells(lngROW_CTRL1, lngCOL_CHK).Width + ws.Cells(lngROW_CTRL1, lngCOL_OPT).Width - (dblSHAPE_MARGIN * 2)
    dblExeButtonHeight = ws.Cells(lngROW_CTRL1, lngCOL_CHK).Height - (dblSHAPE_MARGIN * 2)
    ' 同じ名前の実行ボタンが存在するか確認
    On Error Resume Next
    Set shpExeButton = ws.Shapes(strNAME_EXE_BUTTON1)
    On Error GoTo 0
    ' 実行ボタンが存在する場合、削除
    If Not shpExeButton Is Nothing Then
        shpExeButton.Delete
    End If
    ' 実行ボタンを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpExeButton = ws.Shapes.AddFormControl(xlButtonControl, dblExeButtonLeft, dblExeButtonTop, dblExeButtonWidth, dblExeButtonHeight)
    shpExeButton.Name = strNAME_EXE_BUTTON1
    With ws.Buttons(strNAME_EXE_BUTTON1)
        .Characters.Text = "実行"
    End With
    shpExeButton.OnAction = ws.CodeName & ".ExeButtonClick"
    colControlNames.Add shpExeButton.Name

    ' リセットボタンの位置とサイズを計算
    dblResetButtonLeft = ws.Cells(lngROW_CTRL1, lngCOL_L1).Left + intEXE_COMBOBOX_WIDTH
    dblResetButtonTop = ws.Cells(lngROW_CTRL1, lngCOL_TASK_WGT).Top + dblSHAPE_MARGIN
    dblResetButtonWidth = intRESET_BUTTON_WIDTH
    dblResetButtonHeight = ws.Cells(lngROW_CTRL1, lngCOL_TASK_WGT).Height - (dblSHAPE_MARGIN * 2)
    ' 同じ名前のリセットボタンが存在するか確認
    On Error Resume Next
    Set shpResetButton = ws.Shapes(strNAME_RESET_BUTTON)
    On Error GoTo 0
    ' リセットボタンが存在する場合、削除
    If Not shpResetButton Is Nothing Then
        shpResetButton.Delete
    End If
    ' リセットボタンを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpResetButton = ws.Shapes.AddFormControl(xlButtonControl, dblResetButtonLeft, dblResetButtonTop, dblResetButtonWidth, dblResetButtonHeight)
    shpResetButton.Name = strNAME_RESET_BUTTON
    With ws.Buttons(strNAME_RESET_BUTTON)
        .Characters.Text = "リセット"
    End With
    shpResetButton.OnAction = ws.CodeName & ".PrepareSheet"
    colControlNames.Add shpResetButton.Name

    ' 再計算ボタンの位置とサイズを計算
    dblRecalculateButtonLeft = ws.Cells(lngROW_CTRL1, lngCOL_L1).Left + intEXE_COMBOBOX_WIDTH + intRESET_BUTTON_WIDTH
    dblRecalculateButtonTop = ws.Cells(lngROW_CTRL1, lngCOL_TASK_WGT).Top + dblSHAPE_MARGIN
    dblRecalculateButtonWidth = intRECALCULATE_BUTTON_WIDTH
    dblRecalculateButtonHeight = ws.Cells(lngROW_CTRL1, lngCOL_TASK_WGT).Height - (dblSHAPE_MARGIN * 2)
    ' 同じ名前の再計算ボタンが存在するか確認
    On Error Resume Next
    Set shpRecalculateButton = ws.Shapes(strNAME_RECALCULATE_BUTTON)
    On Error GoTo 0
    ' 再計算ボタンが存在する場合、削除
    If Not shpRecalculateButton Is Nothing Then
        shpRecalculateButton.Delete
    End If
    ' 再計算ボタンを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpRecalculateButton = ws.Shapes.AddFormControl(xlButtonControl, dblRecalculateButtonLeft, dblRecalculateButtonTop, dblRecalculateButtonWidth, dblRecalculateButtonHeight)
    shpRecalculateButton.Name = strNAME_RECALCULATE_BUTTON
    With ws.Buttons(strNAME_RECALCULATE_BUTTON)
        .Characters.Text = "再計算"
    End With
    shpRecalculateButton.OnAction = ws.CodeName & ".RecalculateSheet"
    colControlNames.Add shpRecalculateButton.Name

    ' 対象行数を取得
    lngRowCount = lngEndRow - lngStartRow + 1
    
    ' 一括書き込みのためのデータを用意
    ReDim varChkArray(1 To lngRowCount, 1 To 1)
    ReDim varOptArray(1 To lngRowCount, 1 To 1)
    ReDim varWgtArray(1 To lngRowCount, 1 To 1)
    ReDim varStatusArray(1 To lngRowCount, 1 To 1)
    
    ' 値を配列に格納（固定値）
    For r = 1 To lngRowCount
        varChkArray(r, 1) = strCHK_F                  ' CHK列
        varOptArray(r, 1) = strOPT_F                  ' OPT列
    Next r
    
    ' 値を配列に格納（値に応じて書き込みする値を変えたデータを用意）
    varReadWgtArray = ws.Range(ws.Cells(lngStartRow, lngCOL_TASK_WGT), ws.Cells(lngEndRow, lngCOL_TASK_WGT)).value
    varReadStatusArray = ws.Range(ws.Cells(lngStartRow, lngCOL_WBS_STATUS), ws.Cells(lngEndRow, lngCOL_WBS_STATUS)).value
    Set rngWgtRange = ws.Range(ws.Cells(lngStartRow, lngCOL_TASK_WGT), ws.Cells(lngEndRow, lngCOL_TASK_WGT))
    Set rngStatusRange = ws.Range(ws.Cells(lngStartRow, lngCOL_WBS_STATUS), ws.Cells(lngEndRow, lngCOL_WBS_STATUS))
    For r = 1 To lngRowCount
        If IsEmpty(varReadWgtArray(r, 1)) Then
            varWgtArray(r, 1) = 1
        Else
            varWgtArray(r, 1) = varReadWgtArray(r, 1)
        End If
        
        If IsEmpty(varReadStatusArray(r, 1)) Then
            varStatusArray(r, 1) = "-"
        Else
            varStatusArray(r, 1) = varReadStatusArray(r, 1)
        End If
    Next r
    
    ' 結果を書き込み
    ws.Range(ws.Cells(lngStartRow, lngCOL_CHK), ws.Cells(lngEndRow, lngCOL_CHK)).value = varChkArray
    ws.Range(ws.Cells(lngStartRow, lngCOL_OPT), ws.Cells(lngEndRow, lngCOL_OPT)).value = varOptArray
    ws.Range(ws.Cells(lngStartRow, lngCOL_TASK_WGT), ws.Cells(lngEndRow, lngCOL_TASK_WGT)).value = varWgtArray
    ws.Range(ws.Cells(lngStartRow, lngCOL_WBS_STATUS), ws.Cells(lngEndRow, lngCOL_WBS_STATUS)).value = varStatusArray

    ' 不要なコントロールを削除
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            If Not IsControlNameInCollection(shp.Name, colControlNames) Then
                shp.Delete
            End If
        End If
    Next shp
    
    ' 入力候補（プルダウンを設定）の設定
    SetDropDownListResponsiblePerson
    SetDropDownListResponsibleTeam
    
    ' エラーチェック
    CheckErrors
    
End Sub


' ■ ボタンクリック時に実行される処理
Public Sub ExeButtonClick()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim lngSelectedIndex As Long
    
    ' 選択中のインデックスを取得
    lngSelectedIndex = ws.Shapes(strNAME_EXE_COMBOBOX1).ControlFormat.ListIndex

    ' インデックスに対応する処理を実行
    Select Case lngSelectedIndex
        Case 1
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            InsertRowBelowSelection
            PrepareForProcessing
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 2
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            IncrementSelectedLastLevel
            PrepareForProcessing
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 3
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            DecrementSelectedLastLevel
            PrepareForProcessing
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 4
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
        
            SwapCheckedLastLevel
            PrepareForProcessing
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 5
            MsgBox "未実装です7。"
        Case 6
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            SortWbsRange
            PrepareForProcessing
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 7
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            HideLowerThanL2
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 8
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            HideLowerThanL3
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 9
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            HideLowerThanL4
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 10
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            UnhideAllLevels
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case Else
            MsgBox "項目が選択されていません。"
    End Select

End Sub


' ■ コントロール名がコレクションに含まれているか確認する関数
Private Function IsControlNameInCollection(strControlName As String, colControlNames As Collection) As Boolean

    ' 変数定義
    Dim i As Long
    IsControlNameInCollection = False
    
    ' コントロール名の有無をチェック
    For i = 1 To colControlNames.Count
        If colControlNames(i) = strControlName Then
            IsControlNameInCollection = True
            Exit Function
        End If
    Next i
    
End Function


' ■ 担当プルダウン設定
Private Sub SetDropDownListResponsiblePerson()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strSelectColLetter As String, strSelectRangeLetter As String
    Dim strDefineColLetter As String, strDefineRangeLetter As String
    Dim lngSelectEndRow As Long
    ' 一時変数定義
    Dim r As Long

    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
        
    ' 初期値セット
    lngSelectEndRow = 0

    ' lngCOL_DEFINE_PERSON 列を上から走査して値をセット
    For r = lngStartRow To ws.Cells(ws.Rows.Count, lngCOL_DEFINE_PERSON).End(xlUp).Row
        If IsEmpty(ws.Cells(r, lngCOL_DEFINE_PERSON).value) Or ws.Cells(r, lngCOL_DEFINE_PERSON).value = "" Then
            Exit For
        End If
        lngSelectEndRow = r
    Next r
    
    ' 適用対象範囲を作成
    strSelectColLetter = ConvertColNumberToLetter(lngCOL_PERSON_SLCT)
    strSelectRangeLetter = strSelectColLetter & lngStartRow & ":" & strSelectColLetter & lngEndRow
    
    ' 定義範囲を作成
    strDefineColLetter = ConvertColNumberToLetter(lngCOL_DEFINE_PERSON)
    strDefineRangeLetter = "$" & strDefineColLetter & "$" & lngStartRow & ":$" & strDefineColLetter & "$" & lngSelectEndRow

    ' 適用対象範囲のセルに設定
    With ws.Range(strSelectRangeLetter).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=" & strDefineRangeLetter
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
End Sub


' ■ 組織プルダウン設定
Private Sub SetDropDownListResponsibleTeam()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strSelectColLetter As String, strSelectRangeLetter As String
    Dim strDefineColLetter As String, strDefineRangeLetter As String
    Dim lngSelectEndRow As Long
    ' 一時変数定義
    Dim r As Long

    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 初期値セット
    lngSelectEndRow = 0

    ' lngCOL_DEFINE_TEAM 列を上から走査して値をセット
    For r = lngStartRow To ws.Cells(ws.Rows.Count, lngCOL_DEFINE_TEAM).End(xlUp).Row
        If IsEmpty(ws.Cells(r, lngCOL_DEFINE_TEAM).value) Or ws.Cells(r, lngCOL_DEFINE_TEAM).value = "" Then
            Exit For
        End If
        lngSelectEndRow = r
    Next r
    
    ' 適用対象範囲を作成
    strSelectColLetter = ConvertColNumberToLetter(lngCOL_TEAM_SLCT)
    strSelectRangeLetter = strSelectColLetter & lngStartRow & ":" & strSelectColLetter & lngEndRow
    
    ' 定義範囲を作成
    strDefineColLetter = ConvertColNumberToLetter(lngCOL_DEFINE_TEAM)
    strDefineRangeLetter = "$" & strDefineColLetter & "$" & lngStartRow & ":$" & strDefineColLetter & "$" & lngSelectEndRow

    ' 適用対象範囲のセルに設定
    With ws.Range(strSelectRangeLetter).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
            xlBetween, Formula1:="=" & strDefineRangeLetter
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
    
End Sub


' ■ 予定工数を集計する式をセット
Private Sub SetFormulaForPlannedEffort()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strColLabelPlannedEffort As String, strColLabelTask As String
    Dim strColLabelL1 As String, strColLabelL2 As String, strColLabelL3 As String, strColLabelL4 As String, strColLabelL5 As String
    Dim strBoolArrayH As String, strBoolArrayT As String
    Dim strFormula As String
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer, isTask As Boolean
    Dim val As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ラベル文字列事前用意
    strColLabelPlannedEffort = ConvertColNumberToLetter(lngCOL_PLANNED_EFF)
    strColLabelL1 = ConvertColNumberToLetter(lngCOL_L1)
    strColLabelL2 = ConvertColNumberToLetter(lngCOL_L2)
    strColLabelL3 = ConvertColNumberToLetter(lngCOL_L3)
    strColLabelL4 = ConvertColNumberToLetter(lngCOL_L4)
    strColLabelL5 = ConvertColNumberToLetter(lngCOL_L5)
    strColLabelTask = ConvertColNumberToLetter(lngCOL_TASK)
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(ws.Cells(r, lngCOL_TASK).value) Or ws.Cells(r, lngCOL_TASK).value <> ""
        
        ' タスク行でない場合に式をセット
        If Not isTask Then
            ' L5階層の数式をセット
            If level = 5 Then
                ' 数式を作成
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "=" & strColLabelL5 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelPlannedEffort & r).Formula = strFormula
            End If
            
            ' L4階層の数式をセット
            If level = 4 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "))*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelPlannedEffort & r).Formula = strFormula
            End If
            
            ' L3階層の数式をセット
            If level = 3 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "))*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelPlannedEffort & r).Formula = strFormula
            End If
            
            ' L2階層の数式をセット
            If level = 2 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "))*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelPlannedEffort & r).Formula = strFormula
            End If
            
            ' L1階層の数式をセット
            If level = 1 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "))*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelPlannedEffort & r).Formula = strFormula
            End If
        End If
    Next r
    
    ' L1 を集計する数式をセット
    strBoolArrayH = "(ISNUMBER(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "))*" & _
                   "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                   "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
    strFormula = "=SUM(FILTER(" & strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow & "," & strBoolArrayH & ",0))"
    ws.Range(strColLabelPlannedEffort & lngEndRow + 2).Formula = strFormula
    
End Sub


' ■ タスク件数を集計する式をセット
Private Sub SetFormulaForTaskCount()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strColLabelTaskCount As String, strColLabelTask As String
    Dim strColLabelL1 As String, strColLabelL2 As String, strColLabelL3 As String, strColLabelL4 As String, strColLabelL5 As String
    Dim strBoolArrayH As String, strBoolArrayT As String
    Dim strFormula As String
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer, isTask As Boolean
    Dim val As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ラベル文字列事前用意
    strColLabelTaskCount = ConvertColNumberToLetter(lngCOL_TASK_COUNT)
    strColLabelL1 = ConvertColNumberToLetter(lngCOL_L1)
    strColLabelL2 = ConvertColNumberToLetter(lngCOL_L2)
    strColLabelL3 = ConvertColNumberToLetter(lngCOL_L3)
    strColLabelL4 = ConvertColNumberToLetter(lngCOL_L4)
    strColLabelL5 = ConvertColNumberToLetter(lngCOL_L5)
    strColLabelTask = ConvertColNumberToLetter(lngCOL_TASK)
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(ws.Cells(r, lngCOL_TASK).value) Or ws.Cells(r, lngCOL_TASK).value <> ""
        
        If isTask Then
            ' # タスク行の場合は値をセット #
            ws.Range(strColLabelTaskCount & r).value = 1
        Else
            ' # タスク行でない場合に式をセット #
            ' L5階層の数式をセット
            If level = 5 Then
                ' 数式を作成
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "=" & strColLabelL5 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskCount & r).Formula = strFormula
            End If
            
            ' L4階層の数式をセット
            If level = 4 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "))*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskCount & r).Formula = strFormula
            End If
            
            ' L3階層の数式をセット
            If level = 3 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "))*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskCount & r).Formula = strFormula
            End If
            
            ' L2階層の数式をセット
            If level = 2 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "))*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskCount & r).Formula = strFormula
            End If
            
            ' L1階層の数式をセット
            If level = 1 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "))*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskCount & r).Formula = strFormula
            End If
        End If
    Next r
    
    ' L1 を集計する数式をセット
    strBoolArrayH = "(ISNUMBER(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "))*" & _
                   "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                   "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
    strFormula = "=SUM(FILTER(" & strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow & "," & strBoolArrayH & ",0))"
    ws.Range(strColLabelTaskCount & lngEndRow + 2).Formula = strFormula
    
End Sub


' ■ 実績残工数を集計する式をセット
Private Sub SetFormulaForActualRemainingEffort()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strColLabelActualRemainingEffort As String, strColLabelTask As String
    Dim strColLabelL1 As String, strColLabelL2 As String, strColLabelL3 As String, strColLabelL4 As String, strColLabelL5 As String
    Dim strBoolArrayH As String, strBoolArrayT As String
    Dim strFormula As String
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer, isTask As Boolean
    Dim val As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ラベル文字列事前用意
    strColLabelActualRemainingEffort = ConvertColNumberToLetter(lngCOL_ACTUAL_REMAINING_EFF)
    strColLabelL1 = ConvertColNumberToLetter(lngCOL_L1)
    strColLabelL2 = ConvertColNumberToLetter(lngCOL_L2)
    strColLabelL3 = ConvertColNumberToLetter(lngCOL_L3)
    strColLabelL4 = ConvertColNumberToLetter(lngCOL_L4)
    strColLabelL5 = ConvertColNumberToLetter(lngCOL_L5)
    strColLabelTask = ConvertColNumberToLetter(lngCOL_TASK)
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(ws.Cells(r, lngCOL_TASK).value) Or ws.Cells(r, lngCOL_TASK).value <> ""
        
        ' タスク行でない場合に式をセット
        If Not isTask Then
            ' L5階層の数式をセット
            If level = 5 Then
                ' 数式を作成
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "=" & strColLabelL5 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualRemainingEffort & r).Formula = strFormula
            End If
            
            ' L4階層の数式をセット
            If level = 4 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "))*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualRemainingEffort & r).Formula = strFormula
            End If
            
            ' L3階層の数式をセット
            If level = 3 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "))*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualRemainingEffort & r).Formula = strFormula
            End If
            
            ' L2階層の数式をセット
            If level = 2 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "))*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualRemainingEffort & r).Formula = strFormula
            End If
            
            ' L1階層の数式をセット
            If level = 1 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "))*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualRemainingEffort & r).Formula = strFormula
            End If
        End If
    Next r
    
    ' L1 を集計する数式をセット
    strBoolArrayH = "(ISNUMBER(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "))*" & _
                   "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                   "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
    strFormula = "=SUM(FILTER(" & strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow & "," & strBoolArrayH & ",0))"
    ws.Range(strColLabelActualRemainingEffort & lngEndRow + 2).Formula = strFormula
    
End Sub


' ■ 実績済工数を集計する式をセット
Private Sub SetFormulaForActualCompletedEffort()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strColLabelActualCompletedEffort As String, strColLabelTask As String
    Dim strColLabelL1 As String, strColLabelL2 As String, strColLabelL3 As String, strColLabelL4 As String, strColLabelL5 As String
    Dim strBoolArrayH As String, strBoolArrayT As String
    Dim strFormula As String
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer, isTask As Boolean
    Dim val As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ラベル文字列事前用意
    strColLabelActualCompletedEffort = ConvertColNumberToLetter(lngCOL_ACTUAL_COMPLETED_EFF)
    strColLabelL1 = ConvertColNumberToLetter(lngCOL_L1)
    strColLabelL2 = ConvertColNumberToLetter(lngCOL_L2)
    strColLabelL3 = ConvertColNumberToLetter(lngCOL_L3)
    strColLabelL4 = ConvertColNumberToLetter(lngCOL_L4)
    strColLabelL5 = ConvertColNumberToLetter(lngCOL_L5)
    strColLabelTask = ConvertColNumberToLetter(lngCOL_TASK)
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(ws.Cells(r, lngCOL_TASK).value) Or ws.Cells(r, lngCOL_TASK).value <> ""
        
        ' タスク行でない場合に式をセット
        If Not isTask Then
            ' L5階層の数式をセット
            If level = 5 Then
                ' 数式を作成
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "=" & strColLabelL5 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualCompletedEffort & r).Formula = strFormula
            End If
            
            ' L4階層の数式をセット
            If level = 4 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "))*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualCompletedEffort & r).Formula = strFormula
            End If
            
            ' L3階層の数式をセット
            If level = 3 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "))*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualCompletedEffort & r).Formula = strFormula
            End If
            
            ' L2階層の数式をセット
            If level = 2 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "))*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualCompletedEffort & r).Formula = strFormula
            End If
            
            ' L1階層の数式をセット
            If level = 1 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "))*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelActualCompletedEffort & r).Formula = strFormula
            End If
        End If
    Next r
    
    ' L1 を集計する数式をセット
    strBoolArrayH = "(ISNUMBER(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "))*" & _
                   "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                   "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
    strFormula = "=SUM(FILTER(" & strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow & "," & strBoolArrayH & ",0))"
    ws.Range(strColLabelActualCompletedEffort & lngEndRow + 2).Formula = strFormula
    
End Sub


' ■ タスク進捗率を集計する式をセット
Private Sub SetFormulaForTaskProgressRate()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strColLabelTaskProg As String, strColLabelTaskWeight As String, strColLabelTask As String
    Dim strColLabelL1 As String, strColLabelL2 As String, strColLabelL3 As String, strColLabelL4 As String, strColLabelL5 As String
    Dim strBoolArrayH As String, strBoolArrayT As String
    Dim strSumWeightH As String, strSumWeightT As String
    Dim strFormula As String
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer, isTask As Boolean
    Dim val As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ラベル文字列事前用意
    strColLabelTaskProg = ConvertColNumberToLetter(lngCOL_TASK_PROG)
    strColLabelTaskWeight = ConvertColNumberToLetter(lngCOL_TASK_WGT)
    strColLabelL1 = ConvertColNumberToLetter(lngCOL_L1)
    strColLabelL2 = ConvertColNumberToLetter(lngCOL_L2)
    strColLabelL3 = ConvertColNumberToLetter(lngCOL_L3)
    strColLabelL4 = ConvertColNumberToLetter(lngCOL_L4)
    strColLabelL5 = ConvertColNumberToLetter(lngCOL_L5)
    strColLabelTask = ConvertColNumberToLetter(lngCOL_TASK)
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(ws.Cells(r, lngCOL_TASK).value) Or ws.Cells(r, lngCOL_TASK).value <> ""
        
        ' タスク行でない場合に式をセット
        If Not isTask Then
        
            ' L5階層の数式をセット
            If level = 5 Then
                ' 数式を作成
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "=" & strColLabelL5 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strSumWeightT = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))"
                strFormula = "=SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
                          "," & strBoolArrayT & ",0))" & _
                          "/IF(" & strSumWeightT & "=0,1," & strSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).Formula = strFormula
            End If
            
            ' L4階層の数式をセット
            If level = 4 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "))*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strSumWeightH = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))"
                strSumWeightT = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))"
                strFormula = "=(SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
                          "," & strBoolArrayT & ",0)))" & _
                          "/IF(" & strSumWeightH & "+" & strSumWeightT & "=0,1," & strSumWeightH & "+" & strSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).Formula = strFormula
            End If
            
            ' L3階層の数式をセット
            If level = 3 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "))*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strSumWeightH = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))"
                strSumWeightT = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))"
                strFormula = "=(SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
                          "," & strBoolArrayT & ",0)))" & _
                          "/IF(" & strSumWeightH & "+" & strSumWeightT & "=0,1," & strSumWeightH & "+" & strSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).Formula = strFormula
            End If
            
            ' L2階層の数式をセット
            If level = 2 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "))*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strSumWeightH = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))"
                strSumWeightT = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))"
                strFormula = "=(SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
                          "," & strBoolArrayT & ",0)))" & _
                          "/IF(" & strSumWeightH & "+" & strSumWeightT & "=0,1," & strSumWeightH & "+" & strSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).Formula = strFormula
            End If
            
            ' L1階層の数式をセット
            If level = 1 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "))*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strSumWeightH = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))"
                strSumWeightT = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))"
                strFormula = "=(SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
                          "," & strBoolArrayT & ",0)))" & _
                          "/IF(" & strSumWeightH & "+" & strSumWeightT & "=0,1," & strSumWeightH & "+" & strSumWeightT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).Formula = strFormula
            End If
        End If
    Next r
    
    ' L1 を集計する数式をセット
    strBoolArrayH = "(ISNUMBER(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "))*" & _
                   "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                   "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
    strSumWeightH = "SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))"
    strFormula = "=SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
              "*(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & ")" & _
              "," & strBoolArrayH & ",0))" & _
              "/IF(" & strSumWeightH & "=0,1," & strSumWeightH & ")"
    ws.Range(strColLabelTaskProg & lngEndRow + 2).Formula = strFormula
    
End Sub


' ■ 工数進捗率を集計する式をセット
Private Sub SetFormulaForEffortProgressRate()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strColLabelEffortProg As String, strColLabelActualRemainingEffort As String, strColLabelActualCompletedEffort As String
    Dim strColLabelTask As String
    Dim strColLabelL1 As String, strColLabelL2 As String, strColLabelL3 As String, strColLabelL4 As String, strColLabelL5 As String
    Dim strBoolArrayH As String, strBoolArrayT As String
    Dim strCountH As String, strCountT As String
    Dim strFormula As String
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer, isTask As Boolean
    Dim val As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ラベル文字列事前用意
    strColLabelEffortProg = ConvertColNumberToLetter(lngCOL_EFFORT_PROG)
    strColLabelActualRemainingEffort = ConvertColNumberToLetter(lngCOL_ACTUAL_REMAINING_EFF)
    strColLabelActualCompletedEffort = ConvertColNumberToLetter(lngCOL_ACTUAL_COMPLETED_EFF)
    strColLabelL1 = ConvertColNumberToLetter(lngCOL_L1)
    strColLabelL2 = ConvertColNumberToLetter(lngCOL_L2)
    strColLabelL3 = ConvertColNumberToLetter(lngCOL_L3)
    strColLabelL4 = ConvertColNumberToLetter(lngCOL_L4)
    strColLabelL5 = ConvertColNumberToLetter(lngCOL_L5)
    strColLabelTask = ConvertColNumberToLetter(lngCOL_TASK)
        
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(ws.Cells(r, lngCOL_TASK).value) Or ws.Cells(r, lngCOL_TASK).value <> ""
        
        ' 式をセット
        If isTask Then
            ' # タスク行の処理 #
            strFormula = "=" & strColLabelActualCompletedEffort & r & _
            "/IF(" & strColLabelActualRemainingEffort & r & "+" & strColLabelActualCompletedEffort & r & "=0," & _
            "1," & strColLabelActualRemainingEffort & r & "+" & strColLabelActualCompletedEffort & r & ")"
            ' 指定された列のセルに数式をセット
            ws.Range(strColLabelEffortProg & r).Formula = strFormula
        Else
            ' # タスク行以外の処理 #
            ' L5階層の数式をセット
            If level = 5 Then
                ' 数式を作成
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "=" & strColLabelL5 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strCountT = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayT & ")),0)"
                strFormula = "=SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
                          "," & strBoolArrayT & ",0))" & _
                          "/IF(" & strCountT & "=0,1," & strCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelEffortProg & r).Formula = strFormula
            End If
            
            ' L4階層の数式をセット
            If level = 4 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "))*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strCountH = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayH & ")),0)"
                strCountT = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayT & ")),0)"
                strFormula = "=(SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
                          "," & strBoolArrayT & ",0)))" & _
                          "/IF(" & strCountH & "+" & strCountT & "=0,1," & strCountH & "+" & strCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelEffortProg & r).Formula = strFormula
            End If
            
            ' L3階層の数式をセット
            If level = 3 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "))*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strCountH = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayH & ")),0)"
                strCountT = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayT & ")),0)"
                strFormula = "=(SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
                          "," & strBoolArrayT & ",0)))" & _
                          "/IF(" & strCountH & "+" & strCountT & "=0,1," & strCountH & "+" & strCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelEffortProg & r).Formula = strFormula
            End If
            
            ' L2階層の数式をセット
            If level = 2 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "))*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strCountH = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayH & ")),0)"
                strCountT = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayT & ")),0)"
                strFormula = "=(SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
                          "," & strBoolArrayT & ",0)))" & _
                          "/IF(" & strCountH & "+" & strCountT & "=0,1," & strCountH & "+" & strCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelEffortProg & r).Formula = strFormula
            End If
            
            ' L1階層の数式をセット
            If level = 1 Then
                ' 数式を作成
                strBoolArrayH = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "))*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strCountH = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayH & ")),0)"
                strCountT = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayT & ")),0)"
                strFormula = "=(SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
                          "," & strBoolArrayT & ",0)))" & _
                          "/IF(" & strCountH & "+" & strCountT & "=0,1," & strCountH & "+" & strCountT & ")"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelEffortProg & r).Formula = strFormula
            End If
        End If
    Next r
    
    ' L1 を集計する数式をセット
    strBoolArrayH = "(ISNUMBER(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "))*" & _
                   "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "="""")*" & _
                   "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "="""")*" & _
                   "(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "="""")"
    strCountH = "IFERROR(COUNT(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & "," & strBoolArrayH & ")),0)"
    strFormula = "=SUM(FILTER(" & strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow & _
              "," & strBoolArrayH & ",0))" & _
                          "/IF(" & strCountH & "=0,1," & strCountH & ")"
    ws.Range(strColLabelEffortProg & lngEndRow + 2).Formula = strFormula
    
End Sub


' ■ L2階層以下を非表示（TODO: 高速化）
Private Sub HideLowerThanL2()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim rngToHide As Range
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer
    Dim val As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' L2階層以下ならば非表示とする
        If level >= 2 Then
            If rngToHide Is Nothing Then
                Set rngToHide = ws.Rows(r)
            Else
                Set rngToHide = Union(rngToHide, ws.Rows(r))
            End If
        End If
        
    Next r
    
    ' 非表示にする
    If Not rngToHide Is Nothing Then
        rngToHide.EntireRow.Hidden = True
    End If

End Sub


' ■ L3階層以下を非表示（TODO: 高速化）
Private Sub HideLowerThanL3()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim rngToHide As Range
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer
    Dim val As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' L3階層以下ならば非表示とする
        If level >= 3 Then
            If rngToHide Is Nothing Then
                Set rngToHide = ws.Rows(r)
            Else
                Set rngToHide = Union(rngToHide, ws.Rows(r))
            End If
        End If
        
    Next r
    
    ' 非表示にする
    If Not rngToHide Is Nothing Then
        rngToHide.EntireRow.Hidden = True
    End If

End Sub


' ■ L4階層以下を非表示（TODO: 高速化）
Private Sub HideLowerThanL4()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim rngToHide As Range
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer
    Dim val As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' L4階層以下ならば非表示とする
        If level >= 4 Then
            If rngToHide Is Nothing Then
                Set rngToHide = ws.Rows(r)
            Else
                Set rngToHide = Union(rngToHide, ws.Rows(r))
            End If
        End If
        
    Next r
    
    ' 非表示にする
    If Not rngToHide Is Nothing Then
        rngToHide.EntireRow.Hidden = True
    End If

End Sub


' ■ すべてのレベルを再表示（TODO: 高速化）
Private Sub UnhideAllLevels()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim rngToHide As Range
    ' 一時変数定義
    Dim r As Long
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' すべての行を追加
    For r = lngStartRow To lngEndRow
        
        If rngToHide Is Nothing Then
            Set rngToHide = ws.Rows(r)
        Else
            Set rngToHide = Union(rngToHide, ws.Rows(r))
        End If
        
    Next r
    
    ' 再表示する
    If Not rngToHide Is Nothing Then
        rngToHide.EntireRow.Hidden = False
    End If

End Sub


' ■ 選択中のオプションボタンから行番号を取得
Private Function GetCheckedOptSingleRow() As Long
    
    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim rngFoundCell As Range
    ' 一時変数定義
    Dim r As Long

    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then
        GetCheckedOptSingleRow = 0
        Exit Function
    End If
    
    ' lngStartRow から lngEndRow の範囲で strOPT_T を持つ最初のセルを検索
    On Error Resume Next
    Set rngFoundCell = ws.Range(ws.Cells(lngStartRow, lngCOL_OPT), ws.Cells(lngEndRow, lngCOL_OPT)).Find( _
        What:=strOPT_T, _
        LookAt:=xlWhole, _
        LookIn:=xlValues, _
        MatchCase:=True _
    )
    On Error GoTo 0
    
    ' セルが見つかった場合
    If Not rngFoundCell Is Nothing Then
        GetCheckedOptSingleRow = rngFoundCell.Row
        Exit Function
    End If

    ' ここまで来たらチェックなし
    GetCheckedOptSingleRow = 0
End Function


' ■ 選択中のチェックボックスから行番号コレクションを取得
Private Function GetCheckedChkMultpleRows() As Collection

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim rowCollection As New Collection
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varData As Variant
    ' 一時変数定義
    Dim r As Long

    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then
        Set GetCheckedChkMultpleRows = rowCollection
        Exit Function
    End If

    ' 該当範囲のセルデータを一括で配列に格納
    varData = ws.Range(ws.Cells(lngStartRow, lngCOL_CHK), ws.Cells(lngEndRow, lngCOL_CHK)).value
    
    ' 配列をループして一致する行番号を収集
    For r = 1 To UBound(varData, 1)  ' 配列の行数分だけループ
        If varData(r, 1) = strCHK_T Then
            rowCollection.Add lngStartRow + r - 1 ' 実際の行番号を追加
        End If
    Next r

    ' 結果として行番号のコレクションを返す
    Set GetCheckedChkMultpleRows = rowCollection
End Function


' ■ 指定列の指定行以降で空でないセルを上にシフトする。
Private Sub ShiftUpNonEmptyCells(ByVal intColumnNumber As Integer, ByVal lngStartRow As Long)

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim lngLastRow As Long
    Dim lngCurrentRow As Long
    Dim lngEmptyRow As Long
    Dim rngCurrentCell As Range

    ' 最終行を取得
    lngLastRow = ws.Cells(ws.Rows.Count, intColumnNumber).End(xlUp).Row

    ' 空白行のインデックスを格納する変数
    lngEmptyRow = lngStartRow

    ' lngStartRow から最終行までループ
    For lngCurrentRow = lngStartRow To lngLastRow
        ' 現在のセルを取得
        Set rngCurrentCell = ws.Cells(lngCurrentRow, intColumnNumber)

        ' 空でないセルが見つかった場合
        If Not IsEmpty(rngCurrentCell.value) Then
            ' 空白行の値を現在のセルの値に設定
            If lngCurrentRow > lngEmptyRow Then
                ws.Cells(lngEmptyRow, intColumnNumber).value = rngCurrentCell.value
                ws.Cells(lngCurrentRow, intColumnNumber).ClearContents
            End If
            ' 空白行を次の行に更新
            lngEmptyRow = lngEmptyRow + 1
        End If
    Next lngCurrentRow
End Sub


' ■ 選択した行の下に一行追加
Private Sub InsertRowBelowSelection()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long
    Dim lngSelectedRow As Long
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    
    ' 行を追加
    lngSelectedRow = GetCheckedOptSingleRow()
    If lngSelectedRow <> 0 Then
        ' 行を追加
        ws.Rows(lngSelectedRow + 1).Insert Shift:=xlDown
        ' 組織・担当定義列の空でないセルを上にシフト
        Call ShiftUpNonEmptyCells(lngCOL_DEFINE_TEAM, lngStartRow)
        Call ShiftUpNonEmptyCells(lngCOL_DEFINE_PERSON, lngStartRow)
    Else
        MsgBox "選択してください（OPT)。", vbExclamation, "通知"
    End If

End Sub


' ■ 選択行の最終レベルIDをインクリメント（TODO: 高速化）
Private Sub IncrementSelectedLastLevel()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim lngSelectedRow As Long
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varL1 As Variant, varL2 As Variant, varL3 As Variant, varL4 As Variant, varL5 As Variant, varTask As Variant
    Dim lngSelectedLastId As Long, blnSelectedTaskFlg As Boolean, intSelectedLevel As Integer
    Dim colTargetLastId As New Collection, colTargetRows As New Collection
    Dim lngFirstMissingFoundId As Long
    Dim lngTargetLastIdMax As Long
    ' 一時変数定義
    Dim r As Long
    Dim tmpL1 As Variant, tmpL2 As Variant, tmpL3 As Variant, tmpL4 As Variant, tmpL5 As Variant, tmpTask As Variant
    Dim tmpCurrentLevel As Integer, tmpCurrentTaskFlg As Boolean, tmpCurrentLastId As Long
    Dim isExist As Boolean, tmpVal As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 選択した行の番号を取得
    lngSelectedRow = GetCheckedOptSingleRow()
    
    ' ガード条件（未選択の場合は、メッセージを出して終了）
    If lngSelectedRow = 0 Then
        MsgBox "選択してください（OPT)。", vbExclamation, "通知"
        Exit Sub
    End If
    
    varL1 = ws.Cells(lngSelectedRow, lngCOL_L1).value
    varL2 = ws.Cells(lngSelectedRow, lngCOL_L2).value
    varL3 = ws.Cells(lngSelectedRow, lngCOL_L3).value
    varL4 = ws.Cells(lngSelectedRow, lngCOL_L4).value
    varL5 = ws.Cells(lngSelectedRow, lngCOL_L5).value
    varTask = ws.Cells(lngSelectedRow, lngCOL_TASK).value
    
    ' 選択した行のレベル判定
    intSelectedLevel = 0
    blnSelectedTaskFlg = False
    If Not IsEmpty(varL5) Then
        intSelectedLevel = 5
        lngSelectedLastId = varL5
    ElseIf Not IsEmpty(varL4) Then
        intSelectedLevel = 4
        lngSelectedLastId = varL4
    ElseIf Not IsEmpty(varL3) Then
        intSelectedLevel = 3
        lngSelectedLastId = varL3
    ElseIf Not IsEmpty(varL2) Then
        intSelectedLevel = 2
        lngSelectedLastId = varL2
    ElseIf Not IsEmpty(varL1) Then
        intSelectedLevel = 1
        lngSelectedLastId = varL1
    End If
    If Not IsEmpty(varTask) Then
        blnSelectedTaskFlg = True
        lngSelectedLastId = varTask
    End If
    
    ' 対象レベルに対象の値～最大値までの間で空きがあるかチェック
    For r = lngStartRow To lngEndRow
    
        ' 現在の行における各カラムの値を取得
        tmpL1 = ws.Cells(r, lngCOL_L1).value
        tmpL2 = ws.Cells(r, lngCOL_L2).value
        tmpL3 = ws.Cells(r, lngCOL_L3).value
        tmpL4 = ws.Cells(r, lngCOL_L4).value
        tmpL5 = ws.Cells(r, lngCOL_L5).value
        tmpTask = ws.Cells(r, lngCOL_TASK).value
        
        ' 現在の行のレベル判定
        tmpCurrentLevel = 0
        tmpCurrentTaskFlg = False
        If Not IsEmpty(tmpL5) Then
            tmpCurrentLevel = 5
        ElseIf Not IsEmpty(tmpL4) Then
            tmpCurrentLevel = 4
        ElseIf Not IsEmpty(tmpL3) Then
            tmpCurrentLevel = 3
        ElseIf Not IsEmpty(tmpL2) Then
            tmpCurrentLevel = 2
        ElseIf Not IsEmpty(tmpL1) Then
            tmpCurrentLevel = 1
        End If
        If Not IsEmpty(tmpTask) Then
           tmpCurrentTaskFlg = True
        End If
        
        ' 選択した行と同じ階層のIDを取得
        If intSelectedLevel = 5 Then
            tmpCurrentLastId = tmpL5
        ElseIf intSelectedLevel = 4 Then
            tmpCurrentLastId = tmpL4
        ElseIf intSelectedLevel = 3 Then
            tmpCurrentLastId = tmpL3
        ElseIf intSelectedLevel = 2 Then
            tmpCurrentLastId = tmpL2
        ElseIf intSelectedLevel = 1 Then
            tmpCurrentLastId = tmpL1
        End If
        If blnSelectedTaskFlg = True Then
            tmpCurrentLastId = tmpTask
        End If

        If intSelectedLevel <= tmpCurrentLevel Then
            ' 現在の階層レベルが選択された階層レベル以上でなければならない
            If blnSelectedTaskFlg = True Then
                ' 選択行がタスク行ならば、L1～L5まですべての階層が一致する必要あり
                If varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        varL4 = tmpL4 And _
                        varL5 = tmpL5 And _
                        Not IsEmpty(tmpTask) Then
                    If tmpTask >= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                End If
            Else
                ' 選択行が階層行ならば、一つ前の階層まで一致し、同階層は空であってはいけない
                If intSelectedLevel = 5 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        varL4 = tmpL4 And _
                        Not IsEmpty(tmpL5) Then
                    If tmpL5 >= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                ElseIf intSelectedLevel = 4 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        Not IsEmpty(tmpL4) Then
                    If tmpL4 >= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                ElseIf intSelectedLevel = 3 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        Not IsEmpty(tmpL3) Then
                    If tmpL3 >= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                ElseIf intSelectedLevel = 2 And _
                        varL1 = tmpL1 And _
                        Not IsEmpty(tmpL2) Then
                    If tmpL2 >= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                ElseIf intSelectedLevel = 1 And _
                        Not IsEmpty(tmpL1) Then
                    If tmpL1 >= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next r
    
    ' 最大値を取得
    lngTargetLastIdMax = 0
    For Each tmpVal In colTargetLastId
        If lngTargetLastIdMax < tmpVal Then
            lngTargetLastIdMax = tmpVal
        End If
    Next tmpVal
    
    ' 対象の値～最大値までの値で、存在しない最初の値を取得
    lngFirstMissingFoundId = 0
    For r = lngSelectedLastId To lngTargetLastIdMax
        isExist = False
        For Each tmpVal In colTargetLastId
            If r = tmpVal Then
                isExist = True
                Exit For
            End If
        Next tmpVal
        If isExist = False Then
            lngFirstMissingFoundId = r
            Exit For
        End If
    Next r
    
    ' 存在しない値がなかったら最大値+1をセット
    If lngFirstMissingFoundId = 0 Then
        lngFirstMissingFoundId = lngTargetLastIdMax + 1
    End If
    
    ' 対象行の更新
    For r = lngStartRow To lngEndRow
        ' 現在の行における各カラムの値を取得
        tmpL1 = ws.Cells(r, lngCOL_L1).value
        tmpL2 = ws.Cells(r, lngCOL_L2).value
        tmpL3 = ws.Cells(r, lngCOL_L3).value
        tmpL4 = ws.Cells(r, lngCOL_L4).value
        tmpL5 = ws.Cells(r, lngCOL_L5).value
        tmpTask = ws.Cells(r, lngCOL_TASK).value
        
        ' 現在の行のレベル判定
        tmpCurrentLevel = 0
        tmpCurrentTaskFlg = False
        If Not IsEmpty(tmpL5) Then
            tmpCurrentLevel = 5
        ElseIf Not IsEmpty(tmpL4) Then
            tmpCurrentLevel = 4
        ElseIf Not IsEmpty(tmpL3) Then
            tmpCurrentLevel = 3
        ElseIf Not IsEmpty(tmpL2) Then
            tmpCurrentLevel = 2
        ElseIf Not IsEmpty(tmpL1) Then
            tmpCurrentLevel = 1
        End If
        If Not IsEmpty(tmpTask) Then
            tmpCurrentTaskFlg = True
        End If
        
        ' 選択した行と同じ階層のIDを取得
        If intSelectedLevel = 5 Then
            tmpCurrentLastId = tmpL5
        ElseIf intSelectedLevel = 4 Then
            tmpCurrentLastId = tmpL4
        ElseIf intSelectedLevel = 3 Then
            tmpCurrentLastId = tmpL3
        ElseIf intSelectedLevel = 2 Then
            tmpCurrentLastId = tmpL2
        ElseIf intSelectedLevel = 1 Then
            tmpCurrentLastId = tmpL1
        End If
        If blnSelectedTaskFlg = True Then
            tmpCurrentLastId = tmpTask
        End If

        If intSelectedLevel <= tmpCurrentLevel Then
            ' 現在の階層レベルが選択された階層レベル以上でなければならない
            If blnSelectedTaskFlg = True Then
                ' 選択行がタスク行ならば、L1～L5まですべての階層が一致する必要あり
                If varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        varL4 = tmpL4 And _
                        varL5 = tmpL5 And _
                        Not IsEmpty(tmpTask) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpTask < lngFirstMissingFoundId And tmpTask >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_TASK).value = tmpTask + 1
                    End If
                End If
            Else
                ' 選択行が階層行ならば、一つ前の階層まで一致し、同階層は空であってはいけない
                If intSelectedLevel = 5 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        varL4 = tmpL4 And _
                        Not IsEmpty(tmpL5) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpL5 < lngFirstMissingFoundId And tmpL5 >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L5).value = tmpL5 + 1
                    End If
                ElseIf intSelectedLevel = 4 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        Not IsEmpty(tmpL4) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpL4 < lngFirstMissingFoundId And tmpL4 >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L4).value = tmpL4 + 1
                    End If
                ElseIf intSelectedLevel = 3 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        Not IsEmpty(tmpL3) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpL3 < lngFirstMissingFoundId And tmpL3 >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L3).value = tmpL3 + 1
                    End If
                ElseIf intSelectedLevel = 2 And _
                        varL1 = tmpL1 And _
                        Not IsEmpty(tmpL2) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpL2 < lngFirstMissingFoundId And tmpL2 >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L2).value = tmpL2 + 1
                    End If
                ElseIf intSelectedLevel = 1 And _
                        Not IsEmpty(tmpL1) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpL1 < lngFirstMissingFoundId And tmpL1 >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L1).value = tmpL1 + 1
                    End If
                End If
            End If
        End If
    Next r

End Sub


' ■ 選択行の最終レベルIDをデクリメント（TODO: 高速化）
Private Sub DecrementSelectedLastLevel()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim lngSelectedRow As Long
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varL1 As Variant, varL2 As Variant, varL3 As Variant, varL4 As Variant, varL5 As Variant, varTask As Variant
    Dim lngSelectedLastId As Long, blnSelectedTaskFlg As Boolean, intSelectedLevel As Integer
    Dim colTargetLastId As New Collection, colTargetRows As New Collection
    Dim lngFirstMissingFoundId As Long
    ' 一時変数定義
    Dim r As Long
    Dim tmpL1 As Variant, tmpL2 As Variant, tmpL3 As Variant, tmpL4 As Variant, tmpL5 As Variant, tmpTask As Variant
    Dim tmpCurrentLevel As Integer, tmpCurrentTaskFlg As Boolean, tmpCurrentLastId As Long
    Dim isExist As Boolean, tmpVal As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 選択した行の番号を取得
    lngSelectedRow = GetCheckedOptSingleRow()
    
    ' ガード条件（未選択の場合は、メッセージを出して終了）
    If lngSelectedRow = 0 Then
        MsgBox "選択してください（OPT)。", vbExclamation, "通知"
        Exit Sub
    End If
    
    varL1 = ws.Cells(lngSelectedRow, lngCOL_L1).value
    varL2 = ws.Cells(lngSelectedRow, lngCOL_L2).value
    varL3 = ws.Cells(lngSelectedRow, lngCOL_L3).value
    varL4 = ws.Cells(lngSelectedRow, lngCOL_L4).value
    varL5 = ws.Cells(lngSelectedRow, lngCOL_L5).value
    varTask = ws.Cells(lngSelectedRow, lngCOL_TASK).value
    
    ' 選択した行のレベル判定
    intSelectedLevel = 0
    blnSelectedTaskFlg = False
    If Not IsEmpty(varL5) Then
        intSelectedLevel = 5
        lngSelectedLastId = varL5
    ElseIf Not IsEmpty(varL4) Then
        intSelectedLevel = 4
        lngSelectedLastId = varL4
    ElseIf Not IsEmpty(varL3) Then
        intSelectedLevel = 3
        lngSelectedLastId = varL3
    ElseIf Not IsEmpty(varL2) Then
        intSelectedLevel = 2
        lngSelectedLastId = varL2
    ElseIf Not IsEmpty(varL1) Then
        intSelectedLevel = 1
        lngSelectedLastId = varL1
    End If
    If Not IsEmpty(varTask) Then
        blnSelectedTaskFlg = True
        lngSelectedLastId = varTask
    End If
    
    ' 対象レベルに対象の値以下に空きがあるかチェック
    For r = lngStartRow To lngEndRow
    
        ' 現在の行における各カラムの値を取得
        tmpL1 = ws.Cells(r, lngCOL_L1).value
        tmpL2 = ws.Cells(r, lngCOL_L2).value
        tmpL3 = ws.Cells(r, lngCOL_L3).value
        tmpL4 = ws.Cells(r, lngCOL_L4).value
        tmpL5 = ws.Cells(r, lngCOL_L5).value
        tmpTask = ws.Cells(r, lngCOL_TASK).value
        
        ' 現在の行のレベル判定
        tmpCurrentLevel = 0
        tmpCurrentTaskFlg = False
        If Not IsEmpty(tmpL5) Then
            tmpCurrentLevel = 5
        ElseIf Not IsEmpty(tmpL4) Then
            tmpCurrentLevel = 4
        ElseIf Not IsEmpty(tmpL3) Then
            tmpCurrentLevel = 3
        ElseIf Not IsEmpty(tmpL2) Then
            tmpCurrentLevel = 2
        ElseIf Not IsEmpty(tmpL1) Then
            tmpCurrentLevel = 1
        End If
        If Not IsEmpty(tmpTask) Then
           tmpCurrentTaskFlg = True
        End If
        
        ' 選択した行と同じ階層のIDを取得
        If intSelectedLevel = 5 Then
            tmpCurrentLastId = tmpL5
        ElseIf intSelectedLevel = 4 Then
            tmpCurrentLastId = tmpL4
        ElseIf intSelectedLevel = 3 Then
            tmpCurrentLastId = tmpL3
        ElseIf intSelectedLevel = 2 Then
            tmpCurrentLastId = tmpL2
        ElseIf intSelectedLevel = 1 Then
            tmpCurrentLastId = tmpL1
        End If
        If blnSelectedTaskFlg = True Then
            tmpCurrentLastId = tmpTask
        End If

        If intSelectedLevel <= tmpCurrentLevel Then
            ' 現在の階層レベルが選択された階層レベル以上でなければならない
            If blnSelectedTaskFlg = True Then
                ' 選択行がタスク行ならば、L1～L5まですべての階層が一致する必要あり
                If varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        varL4 = tmpL4 And _
                        varL5 = tmpL5 And _
                        Not IsEmpty(tmpTask) Then
                    If tmpTask <= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                End If
            Else
                ' 選択行が階層行ならば、一つ前の階層まで一致し、同階層は空であってはいけない
                If intSelectedLevel = 5 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        varL4 = tmpL4 And _
                        Not IsEmpty(tmpL5) Then
                    If tmpL5 <= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                ElseIf intSelectedLevel = 4 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        Not IsEmpty(tmpL4) Then
                    If tmpL4 <= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                ElseIf intSelectedLevel = 3 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        Not IsEmpty(tmpL3) Then
                    If tmpL3 <= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                ElseIf intSelectedLevel = 2 And _
                        varL1 = tmpL1 And _
                        Not IsEmpty(tmpL2) Then
                    If tmpL2 <= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                ElseIf intSelectedLevel = 1 And _
                        Not IsEmpty(tmpL1) Then
                    If tmpL1 <= lngSelectedLastId Then
                        On Error Resume Next
                        colTargetLastId.Add tmpCurrentLastId, CStr(tmpCurrentLastId)
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next r
    
    ' 対象の値～1までの値で、存在しない最初の値を取得
    lngFirstMissingFoundId = 0
    For r = lngSelectedLastId To 1 Step -1
        isExist = False
        For Each tmpVal In colTargetLastId
            If r = tmpVal Then
                isExist = True
                Exit For
            End If
        Next tmpVal
        If isExist = False Then
            lngFirstMissingFoundId = r
            Exit For
        End If
    Next r
    
    ' 存在しない値がなかったら終了
    If lngFirstMissingFoundId = 0 Then
        MsgBox "空き番号がありません。", vbExclamation, "通知"
        Exit Sub
    End If
    
    ' 対象行の更新
    For r = lngStartRow To lngEndRow
        ' 現在の行における各カラムの値を取得
        tmpL1 = ws.Cells(r, lngCOL_L1).value
        tmpL2 = ws.Cells(r, lngCOL_L2).value
        tmpL3 = ws.Cells(r, lngCOL_L3).value
        tmpL4 = ws.Cells(r, lngCOL_L4).value
        tmpL5 = ws.Cells(r, lngCOL_L5).value
        tmpTask = ws.Cells(r, lngCOL_TASK).value
        
        ' 現在の行のレベル判定
        tmpCurrentLevel = 0
        tmpCurrentTaskFlg = False
        If Not IsEmpty(tmpL5) Then
            tmpCurrentLevel = 5
        ElseIf Not IsEmpty(tmpL4) Then
            tmpCurrentLevel = 4
        ElseIf Not IsEmpty(tmpL3) Then
            tmpCurrentLevel = 3
        ElseIf Not IsEmpty(tmpL2) Then
            tmpCurrentLevel = 2
        ElseIf Not IsEmpty(tmpL1) Then
            tmpCurrentLevel = 1
        End If
        If Not IsEmpty(tmpTask) Then
            tmpCurrentTaskFlg = True
        End If
        
        ' 選択した行と同じ階層のIDを取得
        If intSelectedLevel = 5 Then
            tmpCurrentLastId = tmpL5
        ElseIf intSelectedLevel = 4 Then
            tmpCurrentLastId = tmpL4
        ElseIf intSelectedLevel = 3 Then
            tmpCurrentLastId = tmpL3
        ElseIf intSelectedLevel = 2 Then
            tmpCurrentLastId = tmpL2
        ElseIf intSelectedLevel = 1 Then
            tmpCurrentLastId = tmpL1
        End If
        If blnSelectedTaskFlg = True Then
            tmpCurrentLastId = tmpTask
        End If

        If intSelectedLevel <= tmpCurrentLevel Then
            ' 現在の階層レベルが選択された階層レベル以上でなければならない
            If blnSelectedTaskFlg = True Then
                ' 選択行がタスク行ならば、L1～L5まですべての階層が一致する必要あり
                If varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        varL4 = tmpL4 And _
                        varL5 = tmpL5 And _
                        Not IsEmpty(tmpTask) Then
                    ' 値を -1 する条件：現在行IDが開いているIDより大きく、選択行のID以下
                    If tmpTask > lngFirstMissingFoundId And tmpTask <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_TASK).value = tmpTask - 1
                    End If
                End If
            Else
                ' 選択行が階層行ならば、一つ前の階層まで一致し、同階層は空であってはいけない
                If intSelectedLevel = 5 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        varL4 = tmpL4 And _
                        Not IsEmpty(tmpL5) Then
                    ' 値を -1 する条件：現在行IDが空いているIDより大きく、選択行のID以下
                    If tmpL5 > lngFirstMissingFoundId And tmpL5 <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L5).value = tmpL5 - 1
                    End If
                ElseIf intSelectedLevel = 4 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        Not IsEmpty(tmpL4) Then
                    ' 値を -1 する条件：現在行IDが空いているIDより大きく、選択行のID以下
                    If tmpL4 > lngFirstMissingFoundId And tmpL4 <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L4).value = tmpL4 - 1
                    End If
                ElseIf intSelectedLevel = 3 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        Not IsEmpty(tmpL3) Then
                    ' 値を -1 する条件：現在行IDが空いているIDより大きく、選択行のID以下
                    If tmpL3 > lngFirstMissingFoundId And tmpL3 <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L3).value = tmpL3 - 1
                    End If
                ElseIf intSelectedLevel = 2 And _
                        varL1 = tmpL1 And _
                        Not IsEmpty(tmpL2) Then
                    ' 値を -1 する条件：現在行IDが空いているIDより大きく、選択行のID以下
                    If tmpL2 > lngFirstMissingFoundId And tmpL2 <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L2).value = tmpL2 - 1
                    End If
                ElseIf intSelectedLevel = 1 And _
                        Not IsEmpty(tmpL1) Then
                    ' 値を -1 する条件：現在行IDが空いているIDより大きく、選択行のID以下
                    If tmpL1 > lngFirstMissingFoundId And tmpL1 <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L1).value = tmpL1 - 1
                    End If
                End If
            End If
        End If
    Next r
    
End Sub


' ■ チェックした２点の最終レベルIDを交換する（TODO: 高速化）
Private Sub SwapCheckedLastLevel()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colCheckedRows As Collection
    Dim lngRow1 As Long, intRowLevel1 As Integer, blnIsTask1 As Boolean, varRow1Id As Variant
    Dim varRow1L1 As Variant, varRow1L2 As Variant, varRow1L3 As Variant, varRow1L4 As Variant, varRow1L5 As Variant, varRow1Task As Variant
    Dim lngRow2 As Long, intRowLevel2 As Integer, blnIsTask2 As Boolean, varRow2Id As Variant
    Dim varRow2L1 As Variant, varRow2L2 As Variant, varRow2L3 As Variant, varRow2L4 As Variant, varRow2L5 As Variant, varRow2Task As Variant
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer, isTask As Boolean
    Dim val As Variant
    Dim tmpL1 As Variant, tmpL2 As Variant, tmpL3 As Variant, tmpL4 As Variant, tmpL5 As Variant, tmpTask As Variant
    
    ' チェックされている行番号を取得
    Set colCheckedRows = GetCheckedChkMultpleRows()
    
    ' ガード条件（チェックが２つでなかった場合は終了）
    If colCheckedRows.Count <> 2 Then
        MsgBox "交換したい２つをチェックしてください（CHK)。" & vbCrLf & "（" & colCheckedRows.Count & " 箇所が選択されています）", vbExclamation, "通知"
        Exit Sub
    End If
    
    ' 対象レコード１を抽出
    lngRow1 = colCheckedRows.Item(1)
    intRowLevel1 = 5
    blnIsTask1 = Not IsEmpty(ws.Cells(lngRow1, lngCOL_TASK).value) Or ws.Cells(lngRow1, lngCOL_TASK).value <> ""
    For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
        val = ws.Cells(lngRow1, levelCol).value
        If Not IsNumeric(val) Or val = "" Then
            intRowLevel1 = intRowLevel1 - 1
        Else
            Exit For
        End If
    Next levelCol
    varRow1L1 = ws.Cells(lngRow1, lngCOL_L1).value
    varRow1L2 = ws.Cells(lngRow1, lngCOL_L2).value
    varRow1L3 = ws.Cells(lngRow1, lngCOL_L3).value
    varRow1L4 = ws.Cells(lngRow1, lngCOL_L4).value
    varRow1L5 = ws.Cells(lngRow1, lngCOL_L5).value
    varRow1Task = ws.Cells(lngRow1, lngCOL_TASK).value
    varRow1Id = ws.Cells(lngRow1, lngCOL_WBS_ID).value
    
    
    ' 対象レコード２を抽出
    lngRow2 = colCheckedRows.Item(2)
    intRowLevel2 = 5
    blnIsTask2 = Not IsEmpty(ws.Cells(lngRow2, lngCOL_TASK).value) Or ws.Cells(lngRow2, lngCOL_TASK).value <> ""
    For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
        val = ws.Cells(lngRow2, levelCol).value
        If Not IsNumeric(val) Or val = "" Then
            intRowLevel2 = intRowLevel2 - 1
        Else
            Exit For
        End If
    Next levelCol
    varRow2L1 = ws.Cells(lngRow2, lngCOL_L1).value
    varRow2L2 = ws.Cells(lngRow2, lngCOL_L2).value
    varRow2L3 = ws.Cells(lngRow2, lngCOL_L3).value
    varRow2L4 = ws.Cells(lngRow2, lngCOL_L4).value
    varRow2L5 = ws.Cells(lngRow2, lngCOL_L5).value
    varRow2Task = ws.Cells(lngRow2, lngCOL_TASK).value
    varRow2Id = ws.Cells(lngRow2, lngCOL_WBS_ID).value
    
    ' ガード条件（２つの階層及びタスクか否かが一致しない場合、終了）
    If (intRowLevel1 <> intRowLevel2) Or (blnIsTask1 <> blnIsTask2) Then
        MsgBox "交換候補の階層およびタスクかどうかが一致しません（CHK)。" & vbCrLf & _
        vbCrLf & "チェック1: 階層=" & intRowLevel1 & ", タスク=" & blnIsTask1 & _
        vbCrLf & "チェック2: 階層=" & intRowLevel2 & ", タスク=" & blnIsTask2 & _
        "", vbExclamation, "通知"
        Exit Sub
    End If
    
    ' ガード条件（２つの末尾番号以外の階層番号が一致しない場合、終了）
    If blnIsTask1 = True Then
        If varRow1L1 <> varRow2L1 Or varRow1L2 <> varRow2L2 Or varRow1L3 <> varRow2L3 Or varRow1L4 <> varRow2L4 Or varRow1L5 <> varRow2L5 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varRow1Id & _
            vbCrLf & "チェック2: " & varRow2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 5 Then
        If varRow1L1 <> varRow2L1 Or varRow1L2 <> varRow2L2 Or varRow1L3 <> varRow2L3 Or varRow1L4 <> varRow2L4 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varRow1Id & _
            vbCrLf & "チェック2: " & varRow2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 4 Then
        If varRow1L1 <> varRow2L1 Or varRow1L2 <> varRow2L2 Or varRow1L3 <> varRow2L3 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varRow1Id & _
            vbCrLf & "チェック2: " & varRow2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 3 Then
        If varRow1L1 <> varRow2L1 Or varRow1L2 <> varRow2L2 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varRow1Id & _
            vbCrLf & "チェック2: " & varRow2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 2 Then
        If varRow1L1 <> varRow2L1 Then
            MsgBox "交換候補の末尾番号以外の階層番号が一致しません（CHK)。" & vbCrLf & _
            vbCrLf & "チェック1: " & varRow1Id & _
            vbCrLf & "チェック2: " & varRow2Id & _
            "", vbExclamation, "通知"
            Exit Sub
        End If
    End If
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' 現在の行における各カラムの値を取得
        tmpL1 = ws.Cells(r, lngCOL_L1).value
        tmpL2 = ws.Cells(r, lngCOL_L2).value
        tmpL3 = ws.Cells(r, lngCOL_L3).value
        tmpL4 = ws.Cells(r, lngCOL_L4).value
        tmpL5 = ws.Cells(r, lngCOL_L5).value
        tmpTask = ws.Cells(r, lngCOL_TASK).value
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(tmpTask) Or tmpTask <> ""
        
        ' 交換を実行
        If isTask = True And isTask = blnIsTask1 Then
            If tmpL1 = varRow1L1 And tmpL2 = varRow1L2 And tmpL3 = varRow1L3 And tmpL4 = varRow1L4 And tmpL5 = varRow1L5 And tmpTask = varRow1Task Then
                ws.Cells(r, lngCOL_TASK).value = varRow2Task
            ElseIf tmpL1 = varRow2L1 And tmpL2 = varRow2L2 And tmpL3 = varRow2L3 And tmpL4 = varRow2L4 And tmpL5 = varRow2L5 And tmpTask = varRow2Task Then
                ws.Cells(r, lngCOL_TASK).value = varRow1Task
            End If
        ElseIf intRowLevel1 = 5 Then
            If tmpL1 = varRow1L1 And tmpL2 = varRow1L2 And tmpL3 = varRow1L3 And tmpL4 = varRow1L4 And tmpL5 = varRow1L5 Then
                ws.Cells(r, lngCOL_L5).value = varRow2L5
            ElseIf tmpL1 = varRow2L1 And tmpL2 = varRow2L2 And tmpL3 = varRow2L3 And tmpL4 = varRow2L4 And tmpL5 = varRow2L5 Then
                ws.Cells(r, lngCOL_L5).value = varRow1L5
            End If
        ElseIf intRowLevel1 = 4 Then
            If tmpL1 = varRow1L1 And tmpL2 = varRow1L2 And tmpL3 = varRow1L3 And tmpL4 = varRow1L4 Then
                ws.Cells(r, lngCOL_L4).value = varRow2L4
            ElseIf tmpL1 = varRow2L1 And tmpL2 = varRow2L2 And tmpL3 = varRow2L3 And tmpL4 = varRow2L4 Then
                ws.Cells(r, lngCOL_L4).value = varRow1L4
            End If
        ElseIf intRowLevel1 = 3 Then
            If tmpL1 = varRow1L1 And tmpL2 = varRow1L2 And tmpL3 = varRow1L3 Then
                ws.Cells(r, lngCOL_L3).value = varRow2L3
            ElseIf tmpL1 = varRow2L1 And tmpL2 = varRow2L2 And tmpL3 = varRow2L3 Then
                ws.Cells(r, lngCOL_L3).value = varRow1L3
            End If
        ElseIf intRowLevel1 = 2 Then
            If tmpL1 = varRow1L1 And tmpL2 = varRow1L2 Then
                ws.Cells(r, lngCOL_L2).value = varRow2L2
            ElseIf tmpL1 = varRow2L1 And tmpL2 = varRow2L2 Then
                ws.Cells(r, lngCOL_L2).value = varRow1L2
            End If
        ElseIf intRowLevel1 = 1 Then
            If tmpL1 = varRow1L1 Then
                ws.Cells(r, lngCOL_L1).value = varRow2L1
            ElseIf tmpL1 = varRow2L1 Then
                ws.Cells(r, lngCOL_L1).value = varRow1L1
            End If
        End If
    
    Next r
    
End Sub


' ■ 数式を数値に変換する
Private Sub ConvertFormulasToValues()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strColLabelActualRemainingEffort As String
    Dim strColLabelActualCompletedEffort As String
    Dim strColLabelTaskCount As String
    Dim strColLabelPlannedEffort As String
    Dim strColLabelEffortProg As String
    Dim strColLabelTaskProg As String
    Dim strColLabelWbsId As String
    Dim strColLabelWbsIdx As String
    ' 一時変数定義
    Dim r As Long
    Dim tmpString As String
    Dim tmpRange As Range
    Dim tmpVariant As Variant
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ラベル文字列事前用意
    strColLabelActualRemainingEffort = ConvertColNumberToLetter(lngCOL_ACTUAL_REMAINING_EFF)
    strColLabelActualCompletedEffort = ConvertColNumberToLetter(lngCOL_ACTUAL_COMPLETED_EFF)
    strColLabelTaskCount = ConvertColNumberToLetter(lngCOL_TASK_COUNT)
    strColLabelPlannedEffort = ConvertColNumberToLetter(lngCOL_PLANNED_EFF)
    strColLabelEffortProg = ConvertColNumberToLetter(lngCOL_EFFORT_PROG)
    strColLabelTaskProg = ConvertColNumberToLetter(lngCOL_TASK_PROG)
    strColLabelWbsId = ConvertColNumberToLetter(lngCOL_WBS_ID)
    strColLabelWbsIdx = ConvertColNumberToLetter(lngCOL_WBS_IDX)
    
    ' 実績残工数の式→値
    Set tmpRange = ws.Range(strColLabelActualRemainingEffort & lngStartRow & ":" & strColLabelActualRemainingEffort & lngEndRow)
    tmpVariant = tmpRange.value
    tmpRange.NumberFormat = "General"
    tmpRange.value = tmpVariant
    
    ' 実績済工数の式→値
    Set tmpRange = ws.Range(strColLabelActualCompletedEffort & lngStartRow & ":" & strColLabelActualCompletedEffort & lngEndRow)
    tmpVariant = tmpRange.value
    tmpRange.NumberFormat = "General"
    tmpRange.value = tmpVariant
    
    ' タスク件数の式→値
    Set tmpRange = ws.Range(strColLabelTaskCount & lngStartRow & ":" & strColLabelTaskCount & lngEndRow)
    tmpVariant = tmpRange.value
    tmpRange.NumberFormat = "General"
    tmpRange.value = tmpVariant
    
    ' 予定工数の式→値
    Set tmpRange = ws.Range(strColLabelPlannedEffort & lngStartRow & ":" & strColLabelPlannedEffort & lngEndRow)
    tmpVariant = tmpRange.value
    tmpRange.NumberFormat = "General"
    tmpRange.value = tmpVariant
    
    ' タスク消化率の式→値
    Set tmpRange = ws.Range(strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow)
    tmpVariant = tmpRange.value
    tmpRange.NumberFormat = "0.0%"
    tmpRange.value = tmpVariant
    
    ' 工数進捗率の式→値
    Set tmpRange = ws.Range(strColLabelEffortProg & lngStartRow & ":" & strColLabelEffortProg & lngEndRow)
    tmpVariant = tmpRange.value
    tmpRange.NumberFormat = "0.0%"
    tmpRange.value = tmpVariant
    
    ' WbsId の式→値（文字列化）
    Set tmpRange = ws.Range(strColLabelWbsId & lngStartRow & ":" & strColLabelWbsId & lngEndRow)
    tmpVariant = tmpRange.value
    tmpRange.NumberFormat = "@"
    tmpRange.value = tmpVariant
    
    ' WbsIdx の式→値（文字列化）
    Set tmpRange = ws.Range(strColLabelWbsIdx & lngStartRow & ":" & strColLabelWbsIdx & lngEndRow)
    tmpVariant = tmpRange.value
    tmpRange.NumberFormat = "@"
    tmpRange.value = tmpVariant
        
    ' 特定のセルの式を値に変換
    ws.Range(strColLabelActualRemainingEffort & lngEndRow + 2).value = ws.Range(strColLabelActualRemainingEffort & lngEndRow + 2).value
    ws.Range(strColLabelActualCompletedEffort & lngEndRow + 2).value = ws.Range(strColLabelActualCompletedEffort & lngEndRow + 2).value
    ws.Range(strColLabelTaskCount & lngEndRow + 2).value = ws.Range(strColLabelTaskCount & lngEndRow + 2).value
    ws.Range(strColLabelPlannedEffort & lngEndRow + 2).value = ws.Range(strColLabelPlannedEffort & lngEndRow + 2).value
    ws.Range(strColLabelEffortProg & lngEndRow + 2).value = ws.Range(strColLabelEffortProg & lngEndRow + 2).value
    ws.Range(strColLabelTaskProg & lngEndRow + 2).value = ws.Range(strColLabelTaskProg & lngEndRow + 2).value

End Sub


' ■ シート操作の準備をする
Private Sub PrepareSheet()
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    PrepareForProcessing
    
    ' 集計式の挿入
    SetFormulaForTaskCount
    SetFormulaForPlannedEffort
    SetFormulaForActualRemainingEffort
    SetFormulaForActualCompletedEffort
    SetFormulaForEffortProgressRate
    SetFormulaForTaskProgressRate
    SetFormulaForWbsId
    SetFormulaForWbsIdx
    
    ' 式のマニュアル更新時、一時的自動計算を行う
    If Application.Calculation = xlCalculationManual Then
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
    End If
    
    ' 式を値に変換
    ConvertFormulasToValues
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub


' ■ シートを再計算する
Private Sub RecalculateSheet()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' 集計式の挿入
    SetFormulaForPlannedEffort
    SetFormulaForActualRemainingEffort
    SetFormulaForActualCompletedEffort
    SetFormulaForEffortProgressRate
    SetFormulaForTaskProgressRate
    SetFormulaForWbsId
    SetFormulaForWbsIdx
    
    ' 式のマニュアル更新時、一時的自動計算を行う
    If Application.Calculation = xlCalculationManual Then
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual
    End If
    
    ' 式を値に変換
    ConvertFormulasToValues
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Sub


' ■ チェックした行を削除する
Private Sub RemoveCheckedRows()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colCheckedRows As Collection
    Dim lngRow1 As Long, intRowLevel1 As Integer, blnIsTask1 As Boolean, varRow1Id As Variant
    Dim varRow1L1 As Variant, varRow1L2 As Variant, varRow1L3 As Variant, varRow1L4 As Variant, varRow1L5 As Variant, varRow1Task As Variant
    Dim lngRow2 As Long, intRowLevel2 As Integer, blnIsTask2 As Boolean, varRow2Id As Variant
    Dim varRow2L1 As Variant, varRow2L2 As Variant, varRow2L3 As Variant, varRow2L4 As Variant, varRow2L5 As Variant, varRow2Task As Variant
    ' 一時変数定義
    Dim r As Long, levelCol As Integer, level As Integer, isTask As Boolean
    Dim val As Variant
    Dim tmpL1 As Variant, tmpL2 As Variant, tmpL3 As Variant, tmpL4 As Variant, tmpL5 As Variant, tmpTask As Variant
    
    ' チェックされている行番号を取得
    Set colCheckedRows = GetCheckedChkMultpleRows()
    
    ' 開始行と終了行を取得
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
    Next r

End Sub


' ◇ 利用イメージ↓
' Dim persons As New Collection
' Call SavePersonList(persons)
' Dim orgs As Collection
' Set orgs = LoadOrganizationList()

Private Sub ClearPersonList()

End Sub

Private Sub SavePersonList(ByVal colPersons As Collection)

End Sub

Private Function LoadPersonList() As Collection

End Function

Private Sub ClearOrganizationList()

End Sub

Private Sub SaveOrganizationList(ByVal colOrgs As Collection)

End Sub

Private Function LoadOrganizationList() As Collection

End Function
