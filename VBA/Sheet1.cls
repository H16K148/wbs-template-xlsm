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

' 列定義
Private Const lngCOL_KEY  As Long = 1             '  A列：KEY カラム
Private Const lngCOL_ERR  As Long = 2             '  B列：ERROR 情報が出力されるカラム
Private Const lngCOL_CHK  As Long = 3             '  C列：チェックボックスのカラム
Private Const lngCOL_OPT  As Long = 4             '  D列：オプションボックスのカラム
Private Const lngCOL_L1   As Long = 5             '  E列：L1 番号を入力するカラム
Private Const lngCOL_L2   As Long = 6             '  F列：L2 番号を入力するカラム
Private Const lngCOL_L3   As Long = 7             '  G列：L3 番号を入力するカラム
Private Const lngCOL_L4   As Long = 8             '  H列：L4 番号を入力するカラム
Private Const lngCOL_L5   As Long = 9             '  I列：L5 番号を入力するカラム
Private Const lngCOL_TASK As Long = 10            '  J列：TASK 番号を入力するカラム
Private Const lngCOL_WBS_IDX As Long = 11         '  K列：WBS_IDXを表示するカラム（普段は非表示）
Private Const lngCOL_WBS_ID As Long = 12          '  L列：WBS_IDを表示するカラム
Private Const lngCOL_EFF_PROG As Long = 20        '  T列：工数進捗率のカラム
Private Const lngCOL_TASK_PROG As Long = 21       '  U列：項目消化率のカラム
Private Const lngCOL_TASK_WGT As Long = 22        '  V列：項目加重のカラム
Private Const lngCOL_WBS_STATUS As Long = 23      '  W列：WBSステータスのカラム
Private Const lngCOL_TEAM_SLCT As Long = 24       '  X列：組織選択のカラム
Private Const lngCOL_PERSON_SLCT As Long = 25     '  Y列：担当選択のカラム
Private Const lngCOL_PLANNED_EFF As Long = 27     ' AA列：予定工数のカラム
Private Const lngCOL_LAST As Long = 34            ' AH列：（データの）最終カラム
Private Const lngCOL_DEFINE_TEAM As Long = 36     ' AJ列：組織を定義するカラム
Private Const lngCOL_DEFINE_PERSON As Long = 37   ' AK列：担当を定義するカラム

' 行定義
Private Const lngROW_CTRL1 As Long = 3            ' コントロールを配置する行

' コントロール名関連定義
Private Const strNAME_OPT_GROUPBOX1 As String = "OptionGroupBox_1"   ' オプションボタングループ名
Private Const strNAME_EXE_COMBOBOX1 As String = "ExecuteComboBox_1"  '
Private Const strNAME_EXE_BUTTON1   As String = "ExecuteButton_1"    '
Private Const strNAME_RESET_BUTTON As String = "ResetButton_1"      '
Private Const strNAME_PREFIX_CHK As String = "CheckBox_"
Private Const strNAME_PREFIX_OPT As String = "OptionButton_"

' コントロール関連定義
Private Const dblSHAPE_MARGIN As Double = 0.2
Private Const intEXE_COMBOBOX_WIDTH = 500
Private Const intRESET_BUTTON_WIDTH = 55


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


' ■ データ行範囲を取得する
Private Function FindDataRangeRows() As Variant

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim r As Long

    ' 初期値セット
    lngStartRow = 0
    lngEndRow = 0

    ' KEY 列を上から走査して値をセット
    For r = 1 To ws.Cells(ws.Rows.Count, lngCOL_KEY).End(xlUp).Row
        If ws.Cells(r, lngCOL_KEY).Value = "@" Then
            lngStartRow = r
        ElseIf ws.Cells(r, lngCOL_KEY).Value = "$" Then
            lngEndRow = r
            Exit For
        End If
    Next r

    ' lngStartRow + 1 と lngEndRow - 1 を使って実際の範囲を決定
    If lngStartRow > 0 Then
        lngStartRow = lngStartRow + 1
    End If
    If lngEndRow > 0 Then
        lngEndRow = lngEndRow - 1
    End If

    '結果を配列で返す
    FindDataRangeRows = Array(lngStartRow, lngEndRow)
End Function


' ■ 階層エラーチェック１
Private Function CheckHierarchyError1(rngTargetCell As Range, ByRef brfErrorMsg As String) As Boolean
    ' 【チェック内容】
    ' ・L2～TASK 列で、対象セルに値が入っている場合にチェックを行う。
    ' ・対象セルの右側セルに値が入っている場合は、右側セルでチェックを行えばいいため、チェックを終了する。
    ' ・対象セルの親階層となる行（対象セル 1.2.3.4 なら 1, 1.2, 1.2.3 の定義）が存在するかチェックする。

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varL1 As Variant, varL2 As Variant, varL3 As Variant, varL4 As Variant, varL5 As Variant, varTask As Variant
    Dim varL1Check As Variant, varL2Check As Variant, varL3Check As Variant, varL4Check As Variant, varL5Check As Variant, varTaskCheck As Variant
    Dim blnL1Result As Boolean, blnL2Result As Boolean, blnL3Result As Boolean, blnL4Result As Boolean, blnL5Result As Boolean
    ' 一時変数定義
    Dim r As Long
    
    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)
    
    ' 初期状態ではエラーなし
    CheckHierarchyError1 = False
    ' エラーメッセージの初期化
    brfErrorMsg = ""
    
    ' 対象セルが L2～TASK 列ではない、または、対象セルの値が Empty の場合 → チェックしない
    If rngTargetCell.Column < lngCOL_L2 Or rngTargetCell.Column > lngCOL_TASK Or IsEmpty(rngTargetCell.Value) Then
        CheckHierarchyError1 = False
        Exit Function
    End If

    ' 対象セルが TASK 列ではなく、かつ、対象セルの右側が空でない場合はチェックから外す
    If rngTargetCell.Column <> lngCOL_TASK And Not IsEmpty(rngTargetCell.Offset(0, 1).Value) Then
        CheckHierarchyError1 = False
        Exit Function
    End If
    
    ' 対象セルの行の値を取り出す
    varL1 = rngTargetCell.Worksheet.Cells(rngTargetCell.Row, lngCOL_L1).Value
    varL2 = rngTargetCell.Worksheet.Cells(rngTargetCell.Row, lngCOL_L2).Value
    varL3 = rngTargetCell.Worksheet.Cells(rngTargetCell.Row, lngCOL_L3).Value
    varL4 = rngTargetCell.Worksheet.Cells(rngTargetCell.Row, lngCOL_L4).Value
    varL5 = rngTargetCell.Worksheet.Cells(rngTargetCell.Row, lngCOL_L5).Value
    
    ' チェック結果変数に値をセットすることでチェック対象を
    blnL1Result = False
    blnL2Result = False
    blnL3Result = False
    blnL4Result = False
    blnL5Result = False
    If rngTargetCell.Column = lngCOL_L2 Then
        blnL2Result = True
        blnL3Result = True
        blnL4Result = True
        blnL5Result = True
    End If
    If rngTargetCell.Column = lngCOL_L3 Then
        blnL3Result = True
        blnL4Result = True
        blnL5Result = True
    End If
    If rngTargetCell.Column = lngCOL_L4 Then
        blnL4Result = True
        blnL5Result = True
    End If
    If rngTargetCell.Column = lngCOL_L5 Then
        blnL5Result = True
    End If

    ' 行範囲内の各行に対してチェック処理を実施
    For r = lngStartRow To lngEndRow
        ' L1, L2, L3, L4, L5, Task 項目を取得
        varL1Check = rngTargetCell.Worksheet.Cells(r, lngCOL_L1).Value
        varL2Check = rngTargetCell.Worksheet.Cells(r, lngCOL_L2).Value
        varL3Check = rngTargetCell.Worksheet.Cells(r, lngCOL_L3).Value
        varL4Check = rngTargetCell.Worksheet.Cells(r, lngCOL_L4).Value
        varL5Check = rngTargetCell.Worksheet.Cells(r, lngCOL_L5).Value
        varTaskCheck = rngTargetCell.Worksheet.Cells(r, lngCOL_TASK).Value
        
        ' 対象セルに対して L1 定義の行があるかチェック
        If blnL1Result = False Then
            If varL1 = varL1Check And _
                    IsEmpty(varL2Check) And _
                    IsEmpty(varL3Check) And _
                    IsEmpty(varL4Check) And _
                    IsEmpty(varL5Check) And _
                    IsEmpty(varTaskCheck) Then
                blnL1Result = True
            End If
        End If
        
        ' 対象セルに対して L2 定義の行があるかチェック
        If blnL2Result = False Then
            If varL1 = varL1Check And _
                    varL2 = varL2Check And _
                    IsEmpty(varL3Check) And _
                    IsEmpty(varL4Check) And _
                    IsEmpty(varL5Check) And _
                    IsEmpty(varTaskCheck) Then
                blnL2Result = True
            End If
        End If
        
        ' 対象セルに対して L3 定義の行があるかチェック
        If blnL3Result = False Then
            If varL1 = varL1Check And _
                    varL2 = varL2Check And _
                    varL3 = varL3Check And _
                    IsEmpty(varL4Check) And _
                    IsEmpty(varL5Check) And _
                    IsEmpty(varTaskCheck) Then
                blnL3Result = True
            End If
        End If
    
        ' 対象セルに対して L4 定義の行があるかチェック
        If blnL4Result = False Then
            If varL1 = varL1Check And _
                    varL2 = varL2Check And _
                    varL3 = varL3Check And _
                    varL4 = varL4Check And _
                    IsEmpty(varL5Check) And _
                    IsEmpty(varTaskCheck) Then
                blnL4Result = True
            End If
        End If
    
        ' 対象セルに対して L5 定義の行があるかチェック
        If blnL5Result = False Then
            If varL1 = varL1Check And _
                    varL2 = varL2Check And _
                    varL3 = varL3Check And _
                    varL4 = varL4Check And _
                    varL5 = varL5Check And _
                    IsEmpty(varTaskCheck) Then
                blnL5Result = True
            End If
        End If
    Next r
    
    ' エラーが見つかった場合
    If blnL1Result = False Or _
            blnL2Result = False Or _
            blnL3Result = False Or _
            blnL4Result = False Or _
            blnL5Result = False Then
        CheckHierarchyError1 = True
        brfErrorMsg = brfErrorMsg & "親階層が定義されていません。" & vbNewLine
        Exit Function
    End If

    CheckHierarchyError1 = False
End Function


' ■ 階層エラーチェック２
Private Function CheckHierarchyError2(rngTargetCell As Range, ByRef brfErrorMsg As String) As Boolean
    ' 【チェック内容】
    ' ・L2～L5 列で、対象セルが Empty ではなく数値である場合にチェックを行う。
    ' ・対象セルの左側の値が数値であるかチェックする。
    
    ' 変数定義
    Dim rngLeftCell As Range

    ' 対象セルが L2～L5 列ではない → チェックしない
    If rngTargetCell.Column < lngCOL_L2 Or rngTargetCell.Column > lngCOL_L5 Then
        CheckHierarchyError2 = False
        Exit Function
    End If

    ' チェック処理を実施
    If IsNumeric(rngTargetCell.Value) And Not IsEmpty(rngTargetCell.Value) Then
        
        ' 対象セルの左隣のセルを取得
        On Error Resume Next
        Set rngLeftCell = rngTargetCell.Offset(0, -1) ' 左隣のセル
        On Error GoTo 0

        ' 左隣のセルが数値でない場合、エラーメッセージをセット
        If Not rngLeftCell Is Nothing Then
            If Not IsNumeric(rngLeftCell.Value) Or IsEmpty(rngLeftCell.Value) Then
                CheckHierarchyError2 = True
                brfErrorMsg = "階層番号に問題があります。 (" & rngLeftCell.Address(False, False) & " の値が数値でありません)"
                Exit Function
            End If
        End If
    End If

    CheckHierarchyError2 = False
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
    ' 一時変数定義
    Dim r As Long

    ' ERR 列の最終行を取得
    lngLastRow = ws.Cells(ws.Rows.Count, lngCOL_ERR).End(xlUp).Row

    ' ERR 列をチェックして「E」があるか確認
    For r = 1 To lngLastRow
        If ws.Cells(r, lngCOL_ERR).Value = "E" Then
            HasErrors = True
            Exit Function
        End If
    Next r
    
    HasErrors = False
End Function


' ■ エラーチェックを実施
Private Sub CheckErrors()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strErrorMsg As String
    Dim blnHasError As Boolean
    Dim intErrorCount As Integer
    ' 一時変数定義
    Dim r As Long, c As Long
    Dim cell As Range
    Dim tempMsg As String

    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' ERR 列の E をクリア
    ws.Range(ws.Cells(lngStartRow, lngCOL_ERR), ws.Cells(lngEndRow, lngCOL_ERR)).ClearContents

    ' セルを走査
    intErrorCount = 0
    For r = lngStartRow To lngEndRow
        
        blnHasError = False
        strErrorMsg = ""
        
        For c = lngCOL_L1 To lngCOL_LAST
            Set cell = ws.Cells(r, c)

            ' 各エラーチェック関数を実行
            tempMsg = "" ' クリア
            If CheckHierarchyError1(cell, tempMsg) Then
                blnHasError = True
                strErrorMsg = strErrorMsg & vbNewLine & tempMsg & " (" & cell.Address(False, False) & ");"
            End If

            tempMsg = "" ' クリア
            If CheckHierarchyError2(cell, tempMsg) Then
                blnHasError = True
                strErrorMsg = strErrorMsg & vbNewLine & tempMsg & " (" & cell.Address(False, False) & ");"
            End If
        Next c

        ' エラーがあれば ERR 列に「E」、コメント追加
        If blnHasError Then
            ws.Cells(r, lngCOL_ERR).Value = "E"
            If ws.Cells(r, lngCOL_ERR).Comment Is Nothing Then
                ws.Cells(r, lngCOL_ERR).AddComment
            End If
            ws.Cells(r, lngCOL_ERR).Comment.Text Text:=strErrorMsg
            intErrorCount = intErrorCount + 1
            ' コメントの幅と高さを手動で設定
            With ws.Cells(r, lngCOL_ERR).Comment.Shape
                .Width = 300   ' 幅を 300 に設定
                .Height = 100  ' 高さを 100 に設定
            End With
        Else
            ' 既存のコメントがあれば削除
            If Not ws.Cells(r, lngCOL_ERR).Comment Is Nothing Then
                ws.Cells(r, lngCOL_ERR).Comment.Delete
            End If
        End If
    Next r

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
    ' 一時変数定義
    Dim r As Long

    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' lngCOL_WBS_IDX をアルファベットに変換（ループ外で一度だけ実行）
    strTargetColLetter = ConvertColNumberToLetter(lngCOL_WBS_IDX)

    ' 数式をセット
    For r = lngStartRow To lngEndRow
        ' 数式を作成
        strFormula = "=IF(B" & r & "=""E"",""ERROR""," & _
                  "IF(E" & r & "="""",""XXX.XXX.XXX.XXX.XXX.XXX"", CONCAT(" & _
                  "TEXT(E" & r & ",""000"")," & _
                  "IF(F" & r & "="""","".---"", ""."" & TEXT(F" & r & ",""000""))," & _
                  "IF(G" & r & "="""","".---"", ""."" & TEXT(G" & r & ",""000""))," & _
                  "IF(H" & r & "="""","".---"", ""."" & TEXT(H" & r & ",""000""))," & _
                  "IF(I" & r & "="""","".---"", ""."" & TEXT(I" & r & ",""000""))," & _
                  "IF(J" & r & "="""","".---"", ""."" & TEXT(J" & r & ",""000"")))))"
        ' 指定された列のセルに数式をセット
        ws.Range(strTargetColLetter & r).formula = strFormula
    Next r
    
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
    ' 一時変数定義
    Dim r As Long

    ' 開始行と終了行に値をセット
    varRangeRows = FindDataRangeRows()
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' lngCOL_WBS_IDX をアルファベットに変換（ループ外で一度だけ実行）
    strTargetColLetter = ConvertColNumberToLetter(lngCOL_WBS_ID)

    ' 数式をセット
    For r = lngStartRow To lngEndRow
        ' 数式を作成
        strFormula = "=IF(B" & r & "=""E"",""ERROR""," & _
                  "IF(E" & r & "="""","""",CONCAT(" & _
                  "E" & r & "," & _
                  "IF(F" & r & "="""","""","".""&F" & r & " ), " & _
                  "IF(G" & r & "="""","""","".""&G" & r & " ), " & _
                  "IF(H" & r & "="""","""","".""&H" & r & " ), " & _
                  "IF(I" & r & "="""","""","".""&I" & r & " ), " & _
                  "IF(J" & r & "="""","""","".T""&J" & r & " ))))"
        ' 指定された列のセルに数式をセット
        ws.Range(strTargetColLetter & r).formula = strFormula
    Next r
    
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
    ws.Cells(2, 2).Value = ws.Name

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colControlNames As Collection
    ' - 画面配置コントロール
    Dim shpGroupBox As Shape
    Dim shpExeComboBox As Shape
    Dim shpExeButton As Shape
    Dim shpResetButton As Shape
    ' - オプショングループボックスの位置とサイズを計算するための変数
    Dim dblOptGroupBoxLeft As Double
    Dim dblOptGroupBoxTop As Double
    Dim dblOptGroupBoxWidth As Double
    Dim dblOptGroupBoxHeight As Double
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

    ' オプショングループボックスの位置とサイズを計算
    dblOptGroupBoxLeft = ws.Cells(lngStartRow, lngCOL_CHK).Left + dblSHAPE_MARGIN
    dblOptGroupBoxTop = ws.Cells(lngStartRow, lngCOL_CHK).Top + dblSHAPE_MARGIN
    dblOptGroupBoxWidth = ws.Cells(lngEndRow, lngCOL_CHK).Width + ws.Cells(lngEndRow, lngCOL_OPT).Width - (dblSHAPE_MARGIN * 2)
    dblOptGroupBoxHeight = ws.Cells(lngEndRow, lngCOL_OPT).Top + ws.Cells(lngEndRow, lngCOL_OPT).Height - ws.Cells(lngStartRow, lngCOL_OPT).Top - (dblSHAPE_MARGIN * 4)
    ' 同じ名前のオプショングループボックスが存在するか確認
    On Error Resume Next
    Set shpGroupBox = ws.Shapes(strNAME_OPT_GROUPBOX1)
    On Error GoTo 0
    ' オプショングループボックスが存在する場合、削除
    If Not shpGroupBox Is Nothing Then
        shpGroupBox.Delete
    End If
    ' オプショングループボックスを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpGroupBox = ws.Shapes.AddFormControl(xlGroupBox, dblOptGroupBoxLeft, dblOptGroupBoxTop, dblOptGroupBoxWidth, dblOptGroupBoxHeight)
    shpGroupBox.Name = strNAME_OPT_GROUPBOX1
    ' キャプションを削除
    With ws.GroupBoxes(strNAME_OPT_GROUPBOX1)
        .Width = dblOptGroupBoxWidth
        .Caption = ""
        .Visible = False
    End With
    colControlNames.Add shpGroupBox.Name
    
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
        .AddItem "【OPT】 選択した行の上に一行追加"
        .AddItem "【OPT】 選択した階層番号の末尾を＋１"
        .AddItem "【OPT】 選択した階層番号の末尾を－１"
        .AddItem "【CHK】 チェックした２箇所の階層番号の末尾番号を交換　※ チェックが同階層である階層行の２箇所でなかった場合は不可 ※"
        .AddItem "【CHK】 チェックしたタスク行を削除　※ 階層行は不可 ※"
        .AddItem "【CHK】 チェックした階層行を削除　※ タスクがある場合や子階層がある場合は不可 ※"
        .AddItem "階層番号をソート"
        .AddItem "項目消化率、予定工数を集計"
        .AddItem "L2以下を非表示"
        .AddItem "L3以下を非表示"
        .AddItem "L4以下を非表示"
        .AddItem "すべてを再表示"
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
    shpResetButton.OnAction = ws.CodeName & ".PrepareForProcessing"
    colControlNames.Add shpResetButton.Name

    ' 数式をセット
    For r = lngStartRow To lngEndRow
        ' CHK列にチェックボックスを設置または更新
        Set chkBox = CreateOrUpdateCheckBox(r)
        colControlNames.Add chkBox.Name

        ' OPT列に同じグループのラジオボタンを設置または更新
        Set optBtn = CreateOrUpdateOptionButton(r)
        colControlNames.Add optBtn.Name

        ' 項目加重列が Empty の場合、数値の 1 をセット
        If IsEmpty(ws.Cells(r, lngCOL_TASK_WGT).Value) Then
            ws.Cells(r, lngCOL_TASK_WGT).Value = 1
        End If

        ' WBSステータス列が Empty の場合、文字列「-」をセット
        If IsEmpty(ws.Cells(r, lngCOL_WBS_STATUS).Value) Then
            ws.Cells(r, lngCOL_WBS_STATUS).Value = "-"
        End If
    Next r

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
    
    ' 集計式の挿入
    SetFormulaForPlannedEffort
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

    Select Case lngSelectedIndex
        Case 1
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            InsertRowAboveSelection
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
            MsgBox "未実装です6。"
        Case 5
            MsgBox "未実装です7。"
        Case 6
            MsgBox "未実装です8。"
        Case 7
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            SortWbsRange
            PrepareForProcessing
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 8
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
        
            SetFormulaForPlannedEffort
            SetFormulaForTaskProgressRate
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 9
            HideLowerThanL2
        Case 10
            HideLowerThanL3
        Case 11
            HideLowerThanL4
        Case 12
            UnhideAllLevels
        Case Else
            MsgBox "項目が選択されていません。"
    End Select

End Sub


' ■ チェックボックスを設置または更新する関数
Private Function CreateOrUpdateCheckBox(lngRowNum As Long) As Shape

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 名前を作成
    Dim strChkName As String
    strChkName = strNAME_PREFIX_CHK & lngRowNum

    ' 同じ名前のチェックボックスが存在するか確認
    On Error Resume Next
    Set CreateOrUpdateCheckBox = ws.Shapes(strChkName)
    On Error GoTo 0

    ' チェックボックスが存在しない場合、作成
    If CreateOrUpdateCheckBox Is Nothing Then
        Set CreateOrUpdateCheckBox = ws.Shapes.AddFormControl(xlCheckBox, _
            ws.Cells(lngRowNum, lngCOL_CHK).Left + dblSHAPE_MARGIN, _
            ws.Cells(lngRowNum, lngCOL_CHK).Top + dblSHAPE_MARGIN, _
            ws.Cells(lngRowNum, lngCOL_CHK).Width - (dblSHAPE_MARGIN * 2), _
            ws.Cells(lngRowNum, lngCOL_CHK).Height - (dblSHAPE_MARGIN * 2))
        CreateOrUpdateCheckBox.Name = strChkName
    End If

    ' チェックボックス独自の操作
    With ws.CheckBoxes(strChkName)
        .Left = ws.Cells(lngRowNum, lngCOL_CHK).Left + dblSHAPE_MARGIN
        .Top = ws.Cells(lngRowNum, lngCOL_CHK).Top + dblSHAPE_MARGIN
        .Width = ws.Cells(lngRowNum, lngCOL_CHK).Width - (dblSHAPE_MARGIN * 2)
        .Height = ws.Cells(lngRowNum, lngCOL_CHK).Height - (dblSHAPE_MARGIN * 2)
        .Caption = ""
        .Value = False
    End With

End Function


' ■ ラジオボタンを設置または更新する関数
Private Function CreateOrUpdateOptionButton(lngRowNum As Long) As Shape

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 名前を作成
    Dim strOptName As String
    strOptName = strNAME_PREFIX_OPT & lngRowNum

    ' 同じ名前のラジオボタンが存在するか確認
    On Error Resume Next
    Set CreateOrUpdateOptionButton = ws.Shapes(strOptName)
    On Error GoTo 0

    ' ラジオボタンが存在しない場合、作成
    If CreateOrUpdateOptionButton Is Nothing Then
        Set CreateOrUpdateOptionButton = ws.Shapes.AddFormControl(xlOptionButton, _
            ws.Cells(lngRowNum, lngCOL_OPT).Left + dblSHAPE_MARGIN, _
            ws.Cells(lngRowNum, lngCOL_OPT).Top + dblSHAPE_MARGIN, _
            ws.Cells(lngRowNum, lngCOL_OPT).Width - (dblSHAPE_MARGIN * 2), _
            ws.Cells(lngRowNum, lngCOL_OPT).Height - (dblSHAPE_MARGIN * 2))
        CreateOrUpdateOptionButton.Name = strOptName
    End If

    ' ラジオボタン独自の操作
    With ws.OptionButtons(strOptName)
        .Left = ws.Cells(lngRowNum, lngCOL_OPT).Left + dblSHAPE_MARGIN
        .Top = ws.Cells(lngRowNum, lngCOL_OPT).Top + dblSHAPE_MARGIN
        .Width = ws.Cells(lngRowNum, lngCOL_OPT).Width - (dblSHAPE_MARGIN * 2)
        .Height = ws.Cells(lngRowNum, lngCOL_OPT).Height - (dblSHAPE_MARGIN * 2)
        .Caption = ""
        .Value = False
    End With

End Function


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
        If IsEmpty(ws.Cells(r, lngCOL_DEFINE_PERSON).Value) Or ws.Cells(r, lngCOL_DEFINE_PERSON).Value = "" Then
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
        If IsEmpty(ws.Cells(r, lngCOL_DEFINE_TEAM).Value) Or ws.Cells(r, lngCOL_DEFINE_TEAM).Value = "" Then
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
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).Value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(ws.Cells(r, lngCOL_TASK).Value) Or ws.Cells(r, lngCOL_TASK).Value <> ""
        
        ' タスク行でない場合に式をセット
        If Not isTask Then
        
            ' ラベル文字列事前用意
            strColLabelPlannedEffort = ConvertColNumberToLetter(lngCOL_PLANNED_EFF)
            strColLabelL1 = ConvertColNumberToLetter(lngCOL_L1)
            strColLabelL2 = ConvertColNumberToLetter(lngCOL_L2)
            strColLabelL3 = ConvertColNumberToLetter(lngCOL_L3)
            strColLabelL4 = ConvertColNumberToLetter(lngCOL_L4)
            strColLabelL5 = ConvertColNumberToLetter(lngCOL_L5)
            strColLabelTask = ConvertColNumberToLetter(lngCOL_TASK)
            
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
                ws.Range(strColLabelPlannedEffort & r).formula = strFormula
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
                ws.Range(strColLabelPlannedEffort & r).formula = strFormula
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
                ws.Range(strColLabelPlannedEffort & r).formula = strFormula
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
                ws.Range(strColLabelPlannedEffort & r).formula = strFormula
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
                ws.Range(strColLabelPlannedEffort & r).formula = strFormula
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
    ws.Range(strColLabelPlannedEffort & lngEndRow + 2).formula = strFormula
    
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
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).Value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(ws.Cells(r, lngCOL_TASK).Value) Or ws.Cells(r, lngCOL_TASK).Value <> ""
        
        ' タスク行でない場合に式をセット
        If Not isTask Then
        
            ' ラベル文字列事前用意
            strColLabelTaskProg = ConvertColNumberToLetter(lngCOL_TASK_PROG)
            strColLabelTaskWeight = ConvertColNumberToLetter(lngCOL_TASK_WGT)
            strColLabelL1 = ConvertColNumberToLetter(lngCOL_L1)
            strColLabelL2 = ConvertColNumberToLetter(lngCOL_L2)
            strColLabelL3 = ConvertColNumberToLetter(lngCOL_L3)
            strColLabelL4 = ConvertColNumberToLetter(lngCOL_L4)
            strColLabelL5 = ConvertColNumberToLetter(lngCOL_L5)
            strColLabelTask = ConvertColNumberToLetter(lngCOL_TASK)
            
            ' L5階層の数式をセット
            If level = 5 Then
                ' 数式を作成
                strBoolArrayT = "(" & strColLabelL1 & lngStartRow & ":" & strColLabelL1 & lngEndRow & "=" & strColLabelL1 & r & ")*" & _
                          "(" & strColLabelL2 & lngStartRow & ":" & strColLabelL2 & lngEndRow & "=" & strColLabelL2 & r & ")*" & _
                          "(" & strColLabelL3 & lngStartRow & ":" & strColLabelL3 & lngEndRow & "=" & strColLabelL3 & r & ")*" & _
                          "(" & strColLabelL4 & lngStartRow & ":" & strColLabelL4 & lngEndRow & "=" & strColLabelL4 & r & ")*" & _
                          "(" & strColLabelL5 & lngStartRow & ":" & strColLabelL5 & lngEndRow & "=" & strColLabelL5 & r & ")*" & _
                          "(ISNUMBER(" & strColLabelTask & lngStartRow & ":" & strColLabelTask & lngEndRow & "))"
                strFormula = "=SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))" & _
                          "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0)))" & _
                          "," & strBoolArrayT & ",0))"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).formula = strFormula
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
                strFormula = "=(SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0)))" & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))" & _
                          "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0)))" & _
                          "," & strBoolArrayT & ",0)))/2"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).formula = strFormula
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
                strFormula = "=(SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0)))" & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))" & _
                          "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0)))" & _
                          "," & strBoolArrayT & ",0)))/2"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).formula = strFormula
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
                strFormula = "=(SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0)))" & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))" & _
                          "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0)))" & _
                          "," & strBoolArrayT & ",0)))/2"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).formula = strFormula
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
                strFormula = "=(SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))" & _
                          "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0)))" & _
                          "," & strBoolArrayH & ",0))" & _
                          "+SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
                          "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0))" & _
                          "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayT & ",0)))" & _
                          "," & strBoolArrayT & ",0)))/2"
                ' 指定された列のセルに数式をセット
                ws.Range(strColLabelTaskProg & r).formula = strFormula
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
    strFormula = "=SUM(FILTER(" & strColLabelTaskProg & lngStartRow & ":" & strColLabelTaskProg & lngEndRow & _
              "*(COUNT(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0))" & _
              "/SUM(FILTER(" & strColLabelTaskWeight & lngStartRow & ":" & strColLabelTaskWeight & lngEndRow & "," & strBoolArrayH & ",0)))" & _
              "," & strBoolArrayH & ",0))"
    ws.Range(strColLabelTaskProg & lngEndRow + 2).formula = strFormula
    
End Sub


' ■ L2階層以下を非表示
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
            val = ws.Cells(r, levelCol).Value
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
    
    ' 一気に非表示とする
    If Not rngToHide Is Nothing Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        rngToHide.EntireRow.Hidden = True
    
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If

End Sub


' ■ L3階層以下を非表示
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
            val = ws.Cells(r, levelCol).Value
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
    
    ' 一気に非表示とする
    If Not rngToHide Is Nothing Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        rngToHide.EntireRow.Hidden = True
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If

End Sub


' ■ L4階層以下を非表示
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
    startRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).Value
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
    
    ' 一気に非表示とする
    If Not rngToHide Is Nothing Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False

        rngToHide.EntireRow.Hidden = True
    
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If

End Sub


' ■ すべてのレベルを再表示
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
    
    ' 一気に再表示とする
    If Not rngToHide Is Nothing Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        rngToHide.EntireRow.Hidden = False
        PrepareForProcessing
        
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If

End Sub


' ■ 選択中のオプションボタンから行番号を取得
Private Function GetCheckedOptSingleRow() As Long
    
    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
        
    ' 一時変数定義
    Dim optBtn As OptionButton
    Dim btnName As String
    Dim parts() As String

    ' 行番号を取得
    For Each optBtn In ws.OptionButtons
        If optBtn.Value = xlOn Then
            btnName = optBtn.Name
            parts = Split(btnName, "_")
            If UBound(parts) = 1 Then
                GetCheckedOptSingleRow = CLng(parts(1))
                Exit Function
            End If
        End If
    Next optBtn

    ' ここまで来たらチェックなし
    GetCheckedOptSingleRow = 0
End Function


' ■ 選択中のチェックボックスから行番号コレクションを取得
Private Function GetCheckedChkMultpleiRows() As Collection

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 一時変数定義
    Dim cb As CheckBox
    Dim rowCollection As New Collection
    Dim parts() As String
    Dim rowNum As Long

    ' 行番号を収集
    For Each cb In ws.CheckBoxes
        If cb.Value = xlOn Then
            parts = Split(cb.Name, "_")
            If UBound(parts) = 1 Then
                rowNum = CLng(parts(1))
                rowCollection.Add rowNum
            End If
        End If
    Next cb

    Set GetCheckedChkMultpleiRows = rowCollection
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
        If Not IsEmpty(rngCurrentCell.Value) Then
            ' 空白行の値を現在のセルの値に設定
            If lngCurrentRow > lngEmptyRow Then
                ws.Cells(lngEmptyRow, intColumnNumber).Value = rngCurrentCell.Value
                ws.Cells(lngCurrentRow, intColumnNumber).ClearContents
            End If
            ' 空白行を次の行に更新
            lngEmptyRow = lngEmptyRow + 1
        End If
    Next lngCurrentRow
End Sub


' ■ 選択した行の上に一行追加
Private Sub InsertRowAboveSelection()

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
        ws.Rows(lngSelectedRow).Insert Shift:=xlDown
    End If
    
    ' 組織・担当定義列の空でないセルを上にシフト
    Call ShiftUpNonEmptyCells(lngCOL_DEFINE_TEAM, lngStartRow)
    Call ShiftUpNonEmptyCells(lngCOL_DEFINE_PERSON, lngStartRow)

End Sub


' ■ 選択行の最終レベルIDをインクリメント
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
    varL1 = ws.Cells(lngSelectedRow, lngCOL_L1).Value
    varL2 = ws.Cells(lngSelectedRow, lngCOL_L2).Value
    varL3 = ws.Cells(lngSelectedRow, lngCOL_L3).Value
    varL4 = ws.Cells(lngSelectedRow, lngCOL_L4).Value
    varL5 = ws.Cells(lngSelectedRow, lngCOL_L5).Value
    varTask = ws.Cells(lngSelectedRow, lngCOL_TASK).Value
    
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
        tmpL1 = ws.Cells(r, lngCOL_L1).Value
        tmpL2 = ws.Cells(r, lngCOL_L2).Value
        tmpL3 = ws.Cells(r, lngCOL_L3).Value
        tmpL4 = ws.Cells(r, lngCOL_L4).Value
        tmpL5 = ws.Cells(r, lngCOL_L5).Value
        tmpTask = ws.Cells(r, lngCOL_TASK).Value
        
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
        tmpL1 = ws.Cells(r, lngCOL_L1).Value
        tmpL2 = ws.Cells(r, lngCOL_L2).Value
        tmpL3 = ws.Cells(r, lngCOL_L3).Value
        tmpL4 = ws.Cells(r, lngCOL_L4).Value
        tmpL5 = ws.Cells(r, lngCOL_L5).Value
        tmpTask = ws.Cells(r, lngCOL_TASK).Value
        
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
                        ws.Cells(r, lngCOL_TASK).Value = tmpTask + 1
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
                        ws.Cells(r, lngCOL_L5).Value = tmpL5 + 1
                    End If
                ElseIf intSelectedLevel = 4 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        Not IsEmpty(tmpL4) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpL4 < lngFirstMissingFoundId And tmpL4 >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L4).Value = tmpL4 + 1
                    End If
                ElseIf intSelectedLevel = 3 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        Not IsEmpty(tmpL3) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpL3 < lngFirstMissingFoundId And tmpL3 >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L3).Value = tmpL3 + 1
                    End If
                ElseIf intSelectedLevel = 2 And _
                        varL1 = tmpL1 And _
                        Not IsEmpty(tmpL2) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpL2 < lngFirstMissingFoundId And tmpL2 >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L2).Value = tmpL2 + 1
                    End If
                ElseIf intSelectedLevel = 1 And _
                        Not IsEmpty(tmpL1) Then
                    ' 値を +1 する条件：現在行IDが空いているIDより小さく、選択行のID以上
                    If tmpL1 < lngFirstMissingFoundId And tmpL1 >= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L1).Value = tmpL1 + 1
                    End If
                End If
            End If
        End If
    Next r

End Sub


' ■ 選択行の最終レベルIDをデクリメント
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
    varL1 = ws.Cells(lngSelectedRow, lngCOL_L1).Value
    varL2 = ws.Cells(lngSelectedRow, lngCOL_L2).Value
    varL3 = ws.Cells(lngSelectedRow, lngCOL_L3).Value
    varL4 = ws.Cells(lngSelectedRow, lngCOL_L4).Value
    varL5 = ws.Cells(lngSelectedRow, lngCOL_L5).Value
    varTask = ws.Cells(lngSelectedRow, lngCOL_TASK).Value
    
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
        tmpL1 = ws.Cells(r, lngCOL_L1).Value
        tmpL2 = ws.Cells(r, lngCOL_L2).Value
        tmpL3 = ws.Cells(r, lngCOL_L3).Value
        tmpL4 = ws.Cells(r, lngCOL_L4).Value
        tmpL5 = ws.Cells(r, lngCOL_L5).Value
        tmpTask = ws.Cells(r, lngCOL_TASK).Value
        
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
        tmpL1 = ws.Cells(r, lngCOL_L1).Value
        tmpL2 = ws.Cells(r, lngCOL_L2).Value
        tmpL3 = ws.Cells(r, lngCOL_L3).Value
        tmpL4 = ws.Cells(r, lngCOL_L4).Value
        tmpL5 = ws.Cells(r, lngCOL_L5).Value
        tmpTask = ws.Cells(r, lngCOL_TASK).Value
        
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
                        ws.Cells(r, lngCOL_TASK).Value = tmpTask - 1
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
                        ws.Cells(r, lngCOL_L5).Value = tmpL5 - 1
                    End If
                ElseIf intSelectedLevel = 4 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        varL3 = tmpL3 And _
                        Not IsEmpty(tmpL4) Then
                    ' 値を -1 する条件：現在行IDが空いているIDより大きく、選択行のID以下
                    If tmpL4 > lngFirstMissingFoundId And tmpL4 <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L4).Value = tmpL4 - 1
                    End If
                ElseIf intSelectedLevel = 3 And _
                        varL1 = tmpL1 And _
                        varL2 = tmpL2 And _
                        Not IsEmpty(tmpL3) Then
                    ' 値を -1 する条件：現在行IDが空いているIDより大きく、選択行のID以下
                    If tmpL3 > lngFirstMissingFoundId And tmpL3 <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L3).Value = tmpL3 - 1
                    End If
                ElseIf intSelectedLevel = 2 And _
                        varL1 = tmpL1 And _
                        Not IsEmpty(tmpL2) Then
                    ' 値を -1 する条件：現在行IDが空いているIDより大きく、選択行のID以下
                    If tmpL2 > lngFirstMissingFoundId And tmpL2 <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L2).Value = tmpL2 - 1
                    End If
                ElseIf intSelectedLevel = 1 And _
                        Not IsEmpty(tmpL1) Then
                    ' 値を -1 する条件：現在行IDが空いているIDより大きく、選択行のID以下
                    If tmpL1 > lngFirstMissingFoundId And tmpL1 <= lngSelectedLastId Then
                        ws.Cells(r, lngCOL_L1).Value = tmpL1 - 1
                    End If
                End If
            End If
        End If
    Next r
    
End Sub


' ■ チェックした２点の最終レベルIDを交換する
Private Sub SwapCheckedLastLevel()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me
    
    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    
    
End Sub


' ■ 数式を数値に変換する
Private Sub ConvertFormulasToValues()

    ' 現在のシートを取得
    Dim ws As Worksheet
    Set ws = Me

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strColLabelPlannedEffort As String, strColLabelTaskProg As String
    Dim strColLabelWbsId As String, strColLabelWbsIdx As String
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
    strColLabelTaskProg = ConvertColNumberToLetter(lngCOL_TASK_PROG)
    strColLabelWbsId = ConvertColNumberToLetter(lngCOL_WBS_ID)
    strColLabelWbsIdx = ConvertColNumberToLetter(lngCOL_WBS_IDX)
    
    ' すべてのタスクと階層のキーを作成
    For r = lngStartRow To lngEndRow
        
        ' 初期化
        level = 5
        isTask = False
        
        ' 階層レベルの判定
        For levelCol = lngCOL_L5 To lngCOL_L1 Step -1
            val = ws.Cells(r, levelCol).Value
            If Not IsNumeric(val) Or val = "" Then
                level = level - 1
            Else
                Exit For
            End If
        Next levelCol
        
        ' タスクかどうかの判定
        isTask = Not IsEmpty(ws.Cells(r, lngCOL_TASK).Value) Or ws.Cells(r, lngCOL_TASK).Value <> ""
        
        ' タスク行でない場合に式を値に変換
        If Not isTask Then
            ws.Range(strColLabelPlannedEffort & r).Value = ws.Range(strColLabelPlannedEffort & r).Value
            ws.Range(strColLabelTaskProg & r).Value = ws.Range(strColLabelTaskProg & r).Value
        End If
        
        ' すべての行で式を値に変換
        ws.Range(strColLabelWbsId & r).Value = ws.Range(strColLabelWbsId & r).Value
        ws.Range(strColLabelWbsIdx & r).Value = ws.Range(strColLabelWbsIdx & r).Value
        
    Next r
    
    ' 特定のセルの式を値に変換
    ws.Range(strColLabelPlannedEffort & lngEndRow + 2).Value = ws.Range(strColLabelPlannedEffort & lngEndRow + 2).Value
    ws.Range(strColLabelTaskProg & lngEndRow + 2).Value = ws.Range(strColLabelTaskProg & lngEndRow + 2).Value

End Sub
