Attribute VB_Name = "wbsui"
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


' ◆ シート初期化ユーザーフォーム表示用マクロ
Sub ShowInitWBS()
    InitWBS.Show vbModeless
End Sub


' ■ シートを初期化します
Public Sub InitSheet(ws As Worksheet)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' ベースデザインを反映
    InitSheetBaseDesign ws
    
    ' ダミーデータを投入
    InputDammyData ws
    
    ' タイトルセルのリセット
    ResetTitleRow ws
    
    ' 基本数式のリセット
    ResetBasicFormulas ws
    
    ' 集計数式のリセット
    ResetAggregateFormulas ws
    
    ' 条件付き書式のリセット
    ResetConditionalFormatting ws
    
    ' セル配置のリセット
    ResetHorizontalAlignment ws
    
    ' データ入力規則のリセット
    ResetDataValidation ws
        
    ' フォーム関連のリセット
    ResetExecuteForm ws
    
    ' 初期値をセット
    SetInitialValue ws
    
    ' オートフィルターのリセット
    ResetAutoFilter ws
    
    ' シートにイベントコードを追加（ダブルクリックイベント）
    InitDoubleClickHandlerToSheet ws
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub


' □ ダミーデータを登録します（※ シート初期化時以外に使用してはいけません ※）
Private Sub InputDammyData(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long
    Dim level1 As Long, level2 As Long, level3 As Long, level4 As Long, level5 As Long, taskCount As Long
    Dim tmpCurrentRow As Long
    Dim tmpEndFlg As Boolean
    
    level1 = 3
    level2 = 2    ' 各階層1の下に作成する数
    level3 = 2    ' 各階層2の下に作成する数
    level4 = 2    ' 各階層3の下に作成する数
    level5 = 2    ' 各階層4の下に作成する数
    taskCount = 3 ' 各階層に作成するタスク数
    
    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 組織名の入力
    ws.Cells(cfg.ROW_CTRL1, cfg.COL_EFFORT_PROG).value = "A社,B社,C社"
    
    ' 担当名の入力
    ws.Cells(cfg.ROW_CTRL2, cfg.COL_EFFORT_PROG).value = "佐藤一郎,鈴木二郎,高橋三郎"
    
    ' カテゴリ1の入力
    ws.Cells(cfg.ROW_CTRL1, cfg.COL_CATEGORY2).value = "A,B,C"
    
    ' カテゴリ2の入力
    ws.Cells(cfg.ROW_CTRL2, cfg.COL_CATEGORY2).value = "D,E,F"
    
    ' タスク番号を入力
    tmpEndFlg = False            ' 強制終了フラグ
    tmpCurrentRow = lngStartRow  ' 開始行を現在行とする
    ' 階層1
    For i = 1 To level1
        ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
        If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
            tmpEndFlg = True
            Exit For
        End If
        ' 階層1の行入力
        ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
        If i = 1 Then
            ws.Cells(tmpCurrentRow, cfg.COL_L1_TEXT).value = "階層1 テキスト"
        End If
        tmpCurrentRow = tmpCurrentRow + 1
        ' 階層1のタスク行入力
        For n = 1 To taskCount
            ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
            If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                tmpEndFlg = True
                Exit For
            End If
            ' 階層1タスクの行入力
            ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
            ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
            If n = 1 Then
                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層1タスク テキスト1"
                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_DELETED
            ElseIf n = 2 Then
                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層1タスク テキスト2"
                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_REJECTED
            ElseIf n = 3 Then
                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層1タスク テキスト3"
                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
            End If
            tmpCurrentRow = tmpCurrentRow + 1
        Next n
        ' 階層2
        For j = 1 To level2
            ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
            If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                tmpEndFlg = True
                Exit For
            End If
            ' 階層2の行入力
            ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
            ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
            If j = 1 Then
                ws.Cells(tmpCurrentRow, cfg.COL_L2_TEXT).value = "階層2 テキスト"
            End If
            tmpCurrentRow = tmpCurrentRow + 1
            ' 階層2のタスク行入力
            For n = 1 To taskCount
                ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
                If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                    tmpEndFlg = True
                    Exit For
                End If
                ' 階層2タスクの行入力
                ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
                If n = 1 Then
                    ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層2タスク テキスト1"
                    ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_ON_HOLD
                ElseIf n = 2 Then
                    ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層2タスク テキスト2"
                    ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_SHELVED
                ElseIf n = 3 Then
                    ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層2タスク テキスト3"
                    ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                End If
                tmpCurrentRow = tmpCurrentRow + 1
            Next n
            ' 階層3
            For k = 1 To level3
                ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
                If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                    tmpEndFlg = True
                    Exit For
                End If
                ' 階層3の行入力
                ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                If k = 1 Then
                    ws.Cells(tmpCurrentRow, cfg.COL_L3_TEXT).value = "階層3 テキスト"
                End If
                tmpCurrentRow = tmpCurrentRow + 1
                ' 階層3のタスク行入力
                For n = 1 To taskCount
                    ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
                    If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                        tmpEndFlg = True
                        Exit For
                    End If
                    ' 階層3タスクの行入力
                    ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                    ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                    ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                    ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
                    If n = 1 Then
                        ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層3タスク テキスト1"
                        ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_TRANSFERRED
                    ElseIf n = 2 Then
                        ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層3タスク テキスト2"
                        ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_NOT_STARTED
                    ElseIf n = 3 Then
                        ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層3タスク テキスト3"
                        ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                    End If
                    tmpCurrentRow = tmpCurrentRow + 1
                Next n
                ' 階層4
                For l = 1 To level4
                    ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
                    If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                        tmpEndFlg = True
                        Exit For
                    End If
                    ' 階層4の行入力
                    ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                    ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                    ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                    ws.Cells(tmpCurrentRow, cfg.COL_L4).value = l
                    If l = 1 Then
                        ws.Cells(tmpCurrentRow, cfg.COL_L4_TEXT).value = "階層4 テキスト"
                    End If
                    tmpCurrentRow = tmpCurrentRow + 1
                    ' 階層4のタスク行入力
                    For n = 1 To taskCount
                        ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
                        If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                            tmpEndFlg = True
                            Exit For
                        End If
                        ' 階層4タスクの行入力
                        ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                        ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                        ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                        ws.Cells(tmpCurrentRow, cfg.COL_L4).value = l
                        ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
                        If n = 1 Then
                            ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層4タスク テキスト1"
                            ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_COMPLETED
                        ElseIf n = 2 Then
                            ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層4タスク テキスト2"
                            ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_IN_PROGRESS
                        ElseIf n = 3 Then
                            ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層4タスク テキスト3"
                            ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                        End If
                        tmpCurrentRow = tmpCurrentRow + 1
                    Next n
                    ' 階層5
                    For m = 1 To level5
                        ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
                        If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                            tmpEndFlg = True
                            Exit For
                        End If
                        ' 階層5の行入力
                        ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                        ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                        ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                        ws.Cells(tmpCurrentRow, cfg.COL_L4).value = l
                        ws.Cells(tmpCurrentRow, cfg.COL_L5).value = m
                        If m = 1 Then
                            ws.Cells(tmpCurrentRow, cfg.COL_L5_TEXT).value = "階層5 テキスト"
                        End If
                        tmpCurrentRow = tmpCurrentRow + 1
                        ' 階層5のタスク行入力
                        For n = 1 To taskCount
                            ' ガード条件（強制終了フラグが立っているか、最終行の場合、終了）
                            If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                                tmpEndFlg = True
                                Exit For
                            End If
                            ' 階層5タスクの行入力
                            ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                            ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                            ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                            ws.Cells(tmpCurrentRow, cfg.COL_L4).value = l
                            ws.Cells(tmpCurrentRow, cfg.COL_L5).value = m
                            ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
                            If n = 1 Then
                                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層5タスク テキスト1"
                                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                            ElseIf n = 2 Then
                                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "階層5タスク テキスト2"
                                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                            ElseIf n = 3 Then
                                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                            End If
                            tmpCurrentRow = tmpCurrentRow + 1
                        Next n
                    Next m
                Next l
            Next k
        Next j
    Next i

End Sub


' ■ シート初期化 - ベースデザイン（セルのサイズ、全体フォント、など）
Public Sub InitSheetBaseDesign(ws As Worksheet)
    
    ' 変数定義
    Dim win As Window
    Dim lngTitleRow As Long
    Dim lngDataStartRow As Long
    Dim lngDataEndRow As Long
    ' 一時変数定義
    Dim tmpWorksheet As Worksheet
    Dim tmpCharLength As Integer
    
    ' 初期化
    lngTitleRow = 2
    lngDataStartRow = 7
    lngDataEndRow = 219

    ' シート全体のフォントを変更
    With ws.Cells
        .Font.Name = "Yu Gothic"                ' フォント名
        .Font.Size = 9                          ' フォントサイズ
        .Font.Bold = False                      ' 太字（Trueで太字、Falseで通常）
        .Font.Italic = False                    ' 斜体（Trueで斜体、Falseで通常）
        .Font.Underline = xlUnderlineStyleNone  ' 下線（なしに設定）
        .VerticalAlignment = xlVAlignCenter     ' セルの縦方向中央揃え
        .RowHeight = 18
    End With
    
    ' シートの行幅を編集
    ws.Rows(1).RowHeight = 3.75
    ws.Rows(2).RowHeight = 30
    
    ' シートの列幅を編集
    ws.Columns("A").ColumnWidth = 0.08
    ws.Columns("B:R").ColumnWidth = 3
    ws.Columns("S:X").ColumnWidth = 1.5
    ws.Columns("Y").ColumnWidth = 60
    ws.Columns("Z:AA").ColumnWidth = 5
    ws.Columns("AB").ColumnWidth = 8
    ws.Columns("AC:AD").ColumnWidth = 7.5
    ws.Columns("AE").ColumnWidth = 5
    ws.Columns("AF:AH").ColumnWidth = 12
    ws.Columns("AI").ColumnWidth = 8
    ws.Columns("AJ:AK").ColumnWidth = 10
    ws.Columns("AL:AM").ColumnWidth = 8
    ws.Columns("AN:AO").ColumnWidth = 10
    ws.Columns("AP:AQ").ColumnWidth = 12
    ws.Columns("AR").ColumnWidth = 50
    
    ' シート名を装飾
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").IndentLevel = 1
    
    ' コントロール行入力
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL1).value = "全体："
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL2).value = "選択："
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL1).value = "【選択候補定義】組織："
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL2).value = "【選択候補定義】担当："
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL1).value = "【選択候補定義】カテゴリ1："
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL2).value = "【選択候補定義】カテゴリ2："
    
    ' コントロール行入力文字列の装飾
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    
    ' ヘッダー１文字列入力
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1).value = "CHK"
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_HEADER1).value = "OPT"
    ws.Range(cfg.COL_L1_LABEL & cfg.ROW_HEADER1).value = "L1"
    ws.Range(cfg.COL_L2_LABEL & cfg.ROW_HEADER1).value = "L2"
    ws.Range(cfg.COL_L3_LABEL & cfg.ROW_HEADER1).value = "L3"
    ws.Range(cfg.COL_L4_LABEL & cfg.ROW_HEADER1).value = "L4"
    ws.Range(cfg.COL_L5_LABEL & cfg.ROW_HEADER1).value = "L5"
    ws.Range(cfg.COL_TASK_LABEL & cfg.ROW_HEADER1).value = "TASK"
    ws.Range(cfg.COL_WBS_IDX_LABEL & cfg.ROW_HEADER1).value = "IDX"
    ws.Range(cfg.COL_WBS_CNT_LABEL & cfg.ROW_HEADER1).value = "CNT"
    ws.Range(cfg.COL_LEVEL_LABEL & cfg.ROW_HEADER1).value = "LV"
    ws.Range(cfg.COL_FLG_T_LABEL & cfg.ROW_HEADER1).value = "T"
    ws.Range(cfg.COL_FLG_IC_LABEL & cfg.ROW_HEADER1).value = "IC"
    ws.Range(cfg.COL_FLG_PE_LABEL & cfg.ROW_HEADER1).value = "PE"
    ws.Range(cfg.COL_FLG_CE_LABEL & cfg.ROW_HEADER1).value = "CE"
    ws.Range(cfg.COL_WBS_ID_LABEL & cfg.ROW_HEADER1).value = "ID"
    ws.Range(cfg.COL_L1_TEXT_LABEL & cfg.ROW_HEADER1).value = "L1"
    ws.Range(cfg.COL_L2_TEXT_LABEL & cfg.ROW_HEADER1).value = "L2"
    ws.Range(cfg.COL_L3_TEXT_LABEL & cfg.ROW_HEADER1).value = "L3"
    ws.Range(cfg.COL_L4_TEXT_LABEL & cfg.ROW_HEADER1).value = "L4"
    ws.Range(cfg.COL_L5_TEXT_LABEL & cfg.ROW_HEADER1).value = "L5"
    ws.Range(cfg.COL_TASK_TEXT_LABEL & cfg.ROW_HEADER1).value = "TASK"
    ws.Range(cfg.COL_TASK_COUNT_LABEL & cfg.ROW_HEADER1).value = "TASK集計"
    ws.Range(cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_HEADER1).value = "工数"
    ws.Range(cfg.COL_TASK_PROG_LABEL & cfg.ROW_HEADER1).value = "項目"
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER1).value = "予定"
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER1).value = "実績"
    
    ' ヘッダー１文字列の装飾
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenter
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER1).VerticalAlignment = xlBottom
    ws.Range(cfg.COL_TASK_COUNT_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_TASK_PROG_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_WGT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_PLANNED_END_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_ACTUAL_END_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & cfg.ROW_HEADER1).Font.Size = 7
    
    ' ヘッダー２文字列入力
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER2).value = "D-Click!"
    ws.Range(cfg.COL_L1_LABEL & cfg.ROW_HEADER2).value = "階層番号"
    ws.Range(cfg.COL_WBS_ID_LABEL & cfg.ROW_HEADER2).value = "WBS項目名"
    ws.Range(cfg.COL_TASK_COUNT_LABEL & cfg.ROW_HEADER2).value = "合計"
    ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & cfg.ROW_HEADER2).value = "完了"
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_HEADER2).value = "ステータス"
    ws.Range(cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_HEADER2).value = "進捗率"
    ws.Range(cfg.COL_TASK_PROG_LABEL & cfg.ROW_HEADER2).value = "消化率"
    ws.Range(cfg.COL_TASK_WGT_LABEL & cfg.ROW_HEADER2).value = "加重"
    ws.Range(cfg.COL_TEAM_SLCT_LABEL & cfg.ROW_HEADER2).value = "組織"
    ws.Range(cfg.COL_PERSON_SLCT_LABEL & cfg.ROW_HEADER2).value = "担当"
    ws.Range(cfg.COL_OUTPUT_LABEL & cfg.ROW_HEADER2).value = "成果物"
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER2).value = "工数(人日)"
    ws.Range(cfg.COL_PLANNED_START_LABEL & cfg.ROW_HEADER2).value = "開始日"
    ws.Range(cfg.COL_PLANNED_END_LABEL & cfg.ROW_HEADER2).value = "終了日"
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER2).value = "残工数(人日)"
    ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & cfg.ROW_HEADER2).value = "済工数(人日)"
    ws.Range(cfg.COL_ACTUAL_START_LABEL & cfg.ROW_HEADER2).value = "開始日"
    ws.Range(cfg.COL_ACTUAL_END_LABEL & cfg.ROW_HEADER2).value = "終了日"
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_HEADER2).value = "カテゴリ1"
    ws.Range(cfg.COL_CATEGORY2_LABEL & cfg.ROW_HEADER2).value = "カテゴリ2"
    ws.Range(cfg.COL_LAST_LABEL & cfg.ROW_HEADER2).value = "備考"
    
    ' ヘッダー２文字列の装飾
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenter
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_OPT_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_L1_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_TASK_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_WBS_ID_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_TEXT_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenterAcrossSelection
    
    With ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER2)
        tmpCharLength = Len(.value)
        If tmpCharLength > 2 Then
            .Characters(Start:=3, Length:=tmpCharLength - 2).Font.Size = 6
        End If
    End With
    With ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER2)
        tmpCharLength = Len(.value)
        If tmpCharLength > 3 Then
            .Characters(Start:=4, Length:=tmpCharLength - 3).Font.Size = 6
        End If
    End With
    With ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & cfg.ROW_HEADER2)
        tmpCharLength = Len(.value)
        If tmpCharLength > 3 Then
            .Characters(Start:=4, Length:=tmpCharLength - 3).Font.Size = 6
        End If
    End With
    
    ' データ開始キー文字列入力
    ws.Range(cfg.COL_KEY_LABEL & lngDataStartRow).value = "@"
    
    ' データ終了キー文字列入力
    ws.Range(cfg.COL_KEY_LABEL & lngDataEndRow).value = "$"
    
    ' データ終了行文字列入力
    ws.Range(cfg.COL_TASK_COUNT_LABEL & lngDataEndRow).value = "合計"
    ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngDataEndRow).value = "合計"
    ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngDataEndRow).value = "全体%"
    ws.Range(cfg.COL_TASK_PROG_LABEL & lngDataEndRow).value = "全体%"
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngDataEndRow).value = "合計人日"
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngDataEndRow).value = "合計人日"
    ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngDataEndRow).value = "合計人日"
    
    ' データ終了行文字列の装飾
    With ws.Range(cfg.COL_CHK_LABEL & lngDataEndRow & ":" & cfg.COL_LAST_LABEL & lngDataEndRow)
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(255, 255, 255)
    End With
    With ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngDataEndRow)
        tmpCharLength = Len(.value)
        If tmpCharLength > 2 Then
            .Characters(Start:=3, Length:=tmpCharLength - 2).Font.Size = 6
        End If
    End With
    With ws.Range(cfg.COL_TASK_PROG_LABEL & lngDataEndRow)
        tmpCharLength = Len(.value)
        If tmpCharLength > 2 Then
            .Characters(Start:=3, Length:=tmpCharLength - 2).Font.Size = 6
        End If
    End With
    With ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngDataEndRow)
        tmpCharLength = Len(.value)
        If tmpCharLength > 2 Then
            .Characters(Start:=3, Length:=tmpCharLength - 2).Font.Size = 6
        End If
    End With
    With ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngDataEndRow)
        tmpCharLength = Len(.value)
        If tmpCharLength > 2 Then
            .Characters(Start:=3, Length:=tmpCharLength - 2).Font.Size = 6
        End If
    End With
    With ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngDataEndRow)
        tmpCharLength = Len(.value)
        If tmpCharLength > 2 Then
            .Characters(Start:=3, Length:=tmpCharLength - 2).Font.Size = 6
        End If
    End With
    
    ' 非表示カラム設定
    ws.Range(cfg.COL_WBS_IDX_LABEL & 1 & ":" & cfg.COL_FLG_CE_LABEL & 1).EntireColumn.Hidden = True
    
    ' ウィンドウ枠の固定
    Set win = ThisWorkbook.Windows(1)
    win.FreezePanes = False
    ws.Activate
    ws.Cells(lngDataStartRow + 1, cfg.COL_TEAM_SLCT).Select
    win.FreezePanes = True
    ws.Cells(1, 1).Select

End Sub


' ■ シート初期化 - 条件付書式
Public Sub ResetConditionalFormatting(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim tmpFc As FormatCondition
    Dim tmpDataBar As Databar
    
    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' シート全体の条件付書式を削除
    ws.Cells.FormatConditions.Delete
    
    ' ■ 開始・終了行の装飾
    ' 開始行の背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_DATA_START & ":" & cfg.COL_LAST_LABEL & cfg.ROW_DATA_START)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Interior.Color = RGB(0, 0, 0)
            .StopIfTrue = False
        End With
    End With
    ' 終了行の背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & (lngEndRow + 1) & ":" & cfg.COL_LAST_LABEL & (lngEndRow + 1))
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Interior.Color = RGB(0, 0, 0)
            .Font.Bold = True
            .StopIfTrue = False
        End With
    End With
    
    ' ■ ヘッダー行の装飾
    ' ヘッダー１の背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER1)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Interior.Color = RGB(255, 230, 153)
            .Font.Color = RGB(128, 128, 132)
            .StopIfTrue = False
        End With
    End With
    ' ヘッダー２の背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Interior.Color = RGB(255, 230, 153)
            .Font.Bold = True
            .StopIfTrue = False
        End With
    End With
    ' ヘッダー行の格子を設定
    With ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_CTRL1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_CTRL1)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Font.Bold = True
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_CTRL2 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_CTRL2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Font.Bold = True
            With .Borders(xlBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_CHK_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L1_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_L1_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_TASK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_TEAM_SLCT_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_OUTPUT_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_TASK_WGT_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_ACTUAL_END_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    
    ' ■ 値による装飾
    With ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow)
        Set tmpDataBar = .FormatConditions.AddDatabar
        With tmpDataBar
            .MinPoint.Modify xlConditionValueNumber, 0
            .MaxPoint.Modify xlConditionValueNumber, 1
            .BarFillType = xlDataBarFillSolid
            .BarColor.Color = RGB(51, 204, 51)
            .ShowValue = True
        End With
    End With
    With ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow)
        Set tmpDataBar = .FormatConditions.AddDatabar
        With tmpDataBar
            .MinPoint.Modify xlConditionValueNumber, 0
            .MaxPoint.Modify xlConditionValueNumber, 1
            .BarFillType = xlDataBarFillSolid
            .BarColor.Color = RGB(51, 204, 51)
            .ShowValue = True
        End With
    End With
    With ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & "=0")
        With tmpFc
            .Font.Color = RGB(114, 114, 114)
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_TASK_PROG_LABEL & lngStartRow & "=0")
        With tmpFc
            .Font.Color = RGB(114, 114, 114)
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & "=0")
        With tmpFc
            .Font.Color = RGB(114, 114, 114)
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & "=0")
        With tmpFc
            .Font.Color = RGB(114, 114, 114)
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & "=0")
        With tmpFc
            .Font.Color = RGB(114, 114, 114)
            .StopIfTrue = False
        End With
    End With
    
    ' ■ 表示形式の装飾
    With ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2 & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow + 2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        With tmpFc
            .NumberFormat = "0.0%"
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngEndRow + 2 & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow + 2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        With tmpFc
            .NumberFormat = "0.0"
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow + 2 & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow + 2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        With tmpFc
            .NumberFormat = "0.0"
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_PLANNED_START_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_END_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        With tmpFc
            .NumberFormat = "yyyy/mm/dd"
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_ACTUAL_START_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_END_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        With tmpFc
            .NumberFormat = "yyyy/mm/dd"
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        With tmpFc
            .NumberFormat = "(0.0%)"
            .Font.Italic = True
            .Font.Bold = True
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE")
        With tmpFc
            .NumberFormat = "0.0%"
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE")
        With tmpFc
            .NumberFormat = "(0.0%)"
            .Font.Italic = True
            .Font.Bold = True
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE")
        With tmpFc
            .NumberFormat = "0.0"
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE")
        With tmpFc
            .NumberFormat = "(0.0)"
            .Font.Italic = True
            .Font.Bold = True
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE")
        With tmpFc
            .NumberFormat = "0.0"
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE")
        With tmpFc
            .NumberFormat = "(0.0)"
            .Font.Italic = True
            .Font.Bold = True
            .StopIfTrue = False
        End With
    End With
    
    ' ■ データ行のエラー装飾
    With ws.Range(cfg.COL_ERR_LABEL & lngStartRow & ":" & cfg.COL_ERR_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_ERR_LABEL & lngStartRow & "=""E""")
        With tmpFc
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
            .Interior.Color = RGB(255, 204, 204)
            .StopIfTrue = False
        End With
    End With
        
    ' ■ 必ず表示するデータ行の警告装飾
    ' テキストカラムに入力されていない警告
    With ws.Range(cfg.COL_TASK_TEXT_LABEL & lngStartRow & ":" & cfg.COL_TASK_TEXT_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_TASK_TEXT_LABEL & lngStartRow & "="""",$" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 0)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L1_TEXT_LABEL & lngStartRow & ":" & cfg.COL_L1_TEXT_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_LEVEL_LABEL & lngStartRow & "=1,$" & _
                        cfg.COL_L1_TEXT_LABEL & lngStartRow & "="""",$" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 0)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L2_TEXT_LABEL & lngStartRow & ":" & cfg.COL_L2_TEXT_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_LEVEL_LABEL & lngStartRow & "=2,$" & _
                        cfg.COL_L2_TEXT_LABEL & lngStartRow & "="""",$" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 0)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L3_TEXT_LABEL & lngStartRow & ":" & cfg.COL_L3_TEXT_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_LEVEL_LABEL & lngStartRow & "=3,$" & _
                        cfg.COL_L3_TEXT_LABEL & lngStartRow & "="""",$" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 0)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L4_TEXT_LABEL & lngStartRow & ":" & cfg.COL_L4_TEXT_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_LEVEL_LABEL & lngStartRow & "=4,$" & _
                        cfg.COL_L4_TEXT_LABEL & lngStartRow & "="""",$" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 0)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L5_TEXT_LABEL & lngStartRow & ":" & cfg.COL_L5_TEXT_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_LEVEL_LABEL & lngStartRow & "=5,$" & _
                        cfg.COL_L5_TEXT_LABEL & lngStartRow & "="""",$" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 0)
        .StopIfTrue = False
        End With
    End With

    ' ■ データ行の警告装飾
    ' タスク行で入力が必要なカラムに入力されていない警告
    With ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_TASK_PROG_LABEL & lngStartRow & "="""",$" & _
                        cfg.COL_FLG_IC_LABEL & lngStartRow & "=TRUE,$" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 0)
        .StopIfTrue = False
        End With
    End With

    ' ■ データ行の通常装飾
    ' データ行のステータス背景色を設定
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_IC_LABEL & lngStartRow & "=FALSE")
        With tmpFc
            .Font.Strikethrough = True
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_IC_LABEL & lngStartRow & "=FALSE")
        With tmpFc
            .Font.Strikethrough = True
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE,$" & cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_REJECTED & """)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 255)
            .Interior.PatternColor = RGB(217, 217, 217)
            .Interior.Pattern = xlPatternCrissCross
            .Font.Color = RGB(114, 114, 114)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE,$" & cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_SHELVED & """)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 255)
            .Interior.PatternColor = RGB(153, 153, 255)
            .Interior.Pattern = xlPatternDown
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE,$" & cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_TRANSFERRED & """)")
        With tmpFc
            .Interior.Color = RGB(255, 255, 255)
            .Interior.PatternColor = RGB(255, 192, 0)
            .Interior.Pattern = xlPatternUp
            .Font.Color = RGB(114, 114, 114)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE,$" & cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_DELETED & """)")
        With tmpFc
            .Interior.Color = RGB(217, 217, 217)
            .Font.Color = RGB(255, 0, 0)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE,$" & cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_COMPLETED & """)")
        With tmpFc
            .Interior.Color = RGB(217, 217, 217)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE,$" & cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_ON_HOLD & """)")
        With tmpFc
            .Interior.Color = RGB(255, 242, 204)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE,$" & cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_IN_PROGRESS & """)")
        With tmpFc
            .Interior.Color = RGB(226, 239, 218)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE,$" & cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_NOT_STARTED & """)")
        With tmpFc
            .Interior.Color = RGB(252, 228, 214)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=TRUE")
        With tmpFc
            .Font.Color = RGB(114, 114, 114)
        .StopIfTrue = False
        End With
    End With
    ' データ行のレベルカラムの装飾
    With ws.Range(cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_LEVEL_LABEL & lngStartRow & ">=1")
        With tmpFc
            .Interior.Color = RGB(48, 84, 150)
            .Font.Color = RGB(255, 255, 255)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_LEVEL_LABEL & lngStartRow & ">=2")
        With tmpFc
            .Interior.Color = RGB(180, 198, 231)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_LEVEL_LABEL & lngStartRow & ">=3")
        With tmpFc
            .Interior.Color = RGB(217, 225, 242)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_LEVEL_LABEL & lngStartRow & ">=4")
        With tmpFc
            .Interior.Color = RGB(236, 240, 248)
        .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_LEVEL_LABEL & lngStartRow & ">=5")
        With tmpFc
            .Interior.Color = RGB(245, 247, 251)
        .StopIfTrue = False
        End With
    End With
    ' データ行のタスク階層以外文字装飾を設定
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=False")
        With tmpFc
            .Font.Bold = True
        .StopIfTrue = False
        End With
    End With
    ' データ行のタスク階層背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE")
        With tmpFc
            .Interior.Color = RGB(255, 255, 255)
        .StopIfTrue = False
        End With
    End With
    ' データ行のL1階層背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=1,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(48, 84, 150)
            .Font.Color = RGB(255, 255, 255)
        .StopIfTrue = False
        End With
    End With
    ' データ行のL2階層背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=2,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(180, 198, 231)
        .StopIfTrue = False
        End With
    End With
    ' データ行のL3階層背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=3,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(217, 225, 242)
        .StopIfTrue = False
        End With
    End With
    ' データ行のL4階層背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=4,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(236, 240, 248)
        .StopIfTrue = False
        End With
    End With
    ' データ行のL5階層背景色を設定
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=5,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(245, 247, 251)
        .StopIfTrue = False
        End With
    End With
    ' データ行の格子を設定
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlRight)
                .LineStyle = xlDash
                .Weight = xlThin
                .Color = RGB(136, 136, 136)
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlLeft)
                .LineStyle = xlDash
                .Weight = xlThin
                .Color = RGB(136, 136, 136)
            End With
            With .Borders(xlRight)
                .LineStyle = xlDash
                .Weight = xlThin
                .Color = RGB(136, 136, 136)
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_TASK_LABEL & lngStartRow & ":" & cfg.COL_TASK_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlLeft)
                .LineStyle = xlDash
                .Weight = xlThin
                .Color = RGB(136, 136, 136)
            End With
            With .Borders(xlRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_IDX_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    With ws.Range(cfg.COL_WBS_STATUS_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            With .Borders(xlTop)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlBottom)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlLeft)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With .Borders(xlRight)
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            .StopIfTrue = False
        End With
    End With
    
End Sub


' ■ シート初期化 - データ入力規則
Public Sub ResetDataValidation(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim tmpRange As Range
    Dim tmpRuleList As String
       
    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ■ WBSステータス列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_WBS_STATUS_LABEL & lngStartRow & ":" & cfg.COL_WBS_STATUS_LABEL & lngEndRow)
    ' ルールを取得
    tmpRuleList = "-," & cfg.WBS_STATUS_LIST
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' ■ 組織列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_TEAM_SLCT_LABEL & lngStartRow & ":" & cfg.COL_TEAM_SLCT_LABEL & lngEndRow)
    ' ルールを取得
    tmpRuleList = CreateValidationListString(ws, ws.Range(cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_CTRL1), tmpRange)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-," & tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    ' ■ 担当列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_PERSON_SLCT_LABEL & lngStartRow & ":" & cfg.COL_PERSON_SLCT_LABEL & lngEndRow)
    ' ルールを取得
    tmpRuleList = CreateValidationListString(ws, ws.Range(cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_CTRL2), tmpRange)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-," & tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' ■ カテゴリ1列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_CATEGORY1_LABEL & lngStartRow & ":" & cfg.COL_CATEGORY1_LABEL & lngEndRow)
    ' ルールを取得
    tmpRuleList = CreateValidationListString(ws, ws.Range(cfg.COL_CATEGORY2_LABEL & cfg.ROW_CTRL1), tmpRange)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-," & tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' ■ カテゴリ2列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_CATEGORY2_LABEL & lngStartRow & ":" & cfg.COL_CATEGORY2_LABEL & lngEndRow)
    ' ルールを取得
    tmpRuleList = CreateValidationListString(ws, ws.Range(cfg.COL_CATEGORY2_LABEL & cfg.ROW_CTRL2), tmpRange)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-," & tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' ■ 階層番号列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_TASK_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateWholeNumber, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="1", Formula2:="999"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "1～999 の整数"
        .ErrorTitle = "入力エラー"
        .InputMessage = "1～999 の整数を入力してください（空白も可）"
        .ErrorMessage = "1～999 の整数のみ入力可能です。"
        .ShowInput = True
        .ShowError = True
    End With

    ' ■ レベル列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_LEVEL_LABEL & lngStartRow & ":" & cfg.COL_LEVEL_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateWholeNumber, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="0", Formula2:="5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "0～5 の整数のみ入力可能です。"
        .ShowInput = False
        .ShowError = True
    End With

    ' ■ 各種フラグ列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_CE_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_FLG_T_LABEL & lngStartRow & ")," & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE," & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)"
        .IgnoreBlank = True
        .InCellDropdown = False
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "TRUE または FALSE のみ入力可能です。"
        .ShowInput = False
        .ShowError = True
    End With

    ' ■ タスク合計列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ")," & _
                        cfg.COL_TASK_COUNT_LABEL & lngStartRow & ">=0,INT(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ")=" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & "))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "0以上の整数のみ入力可能です。"
        .ShowInput = False
        .ShowError = True
    End With

    ' ■ タスク完了列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ")," & _
                        cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ">=0,INT(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ")=" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & "))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .ErrorTitle = "入力エラー"
        .ErrorMessage = "0以上の整数のみ入力可能です。"
        .ShowInput = False
        .ShowError = True
    End With

    ' ■ 工数進捗率列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ")," & _
                        cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ">=0," & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & "<=1))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0～100% の値のみ"
        .ErrorTitle = "入力エラー"
        .InputMessage = "0～100%（= 0～1）の値を入力してください（空白可）"
        .ErrorMessage = "0～100%の間の数値のみ入力可能です。"
        .ShowInput = True
        .ShowError = True
    End With

    ' ■ 項目消化率列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ")," & _
                        cfg.COL_TASK_PROG_LABEL & lngStartRow & ">=0," & cfg.COL_TASK_PROG_LABEL & lngStartRow & "<=1))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0～100% の値のみ"
        .ErrorTitle = "入力エラー"
        .InputMessage = "0～100%（= 0～1）の値を入力してください（空白可）"
        .ErrorMessage = "0～100%の間の数値のみ入力可能です。"
        .ShowInput = True
        .ShowError = True
    End With

    ' ■ 項目加重列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ")," & _
             cfg.COL_TASK_WGT_LABEL & lngStartRow & ">=1,INT(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ")=" & cfg.COL_TASK_WGT_LABEL & lngStartRow & "))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "1以上の整数のみ"
        .ErrorTitle = "入力エラー"
        .InputMessage = "1以上の整数を入力してください（空白可）"
        .ErrorMessage = "1以上の整数のみ入力可能です。"
        .ShowInput = True
        .ShowError = True
    End With

    ' ■ 予定工数列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ")," & _
             cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ">=0))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0以上の数値"
        .ErrorTitle = "入力エラー"
        .InputMessage = "0以上の数値を入力してください（空白可）"
        .ErrorMessage = "0以上の数値のみ入力できます。"
        .ShowInput = True
        .ShowError = True
    End With

    ' ■ 実績残工数列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ")," & _
             cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ">=0))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0以上の数値"
        .ErrorTitle = "入力エラー"
        .InputMessage = "0以上の数値を入力してください（空白可）"
        .ErrorMessage = "0以上の数値のみ入力できます。"
        .ShowInput = True
        .ShowError = True
    End With

    ' ■ 実績済工数列の入力規則を更新
    ' データの範囲を指定
    Set tmpRange = ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ルールを削除
        .Delete
        ' ルールを設定
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ")," & _
             cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ">=0))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0以上の数値"
        .ErrorTitle = "入力エラー"
        .InputMessage = "0以上の数値を入力してください（空白可）"
        .ErrorMessage = "0以上の数値のみ入力できます。"
        .ShowInput = True
        .ShowError = True
    End With
    
    ' ■ パーセンテージの入力のためのセル書式設定
    ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow).NumberFormat = "0.0%"

End Sub


' □ データ入力規則文字列を作成
Private Function CreateValidationListString(ws As Worksheet, defineRange As Range, dataRange As Range) As String

    ' 変数定義
    Dim colUniqueList As Collection
    Dim strDefine As String
    ' 一時変数定義
    Dim i As Long
    Dim tmpCell As Range
    Dim tmpCellValue As String
    Dim tmpTrimmedValue As String
    Dim tmpDelimiter As String
    Dim tmpDefineArray As Variant
    Dim tmpDataArray As Variant
    Dim tmpVal As String
    Dim r As Long, c As Long
    Dim tmpItem As Variant

    ' Collection オブジェクトを生成
    Set colUniqueList = New Collection

    ' defineRange の文字列を処理
    strDefine = Trim(defineRange.value)
    If Len(strDefine) > 0 Then
        tmpDefineArray = Split(strDefine, ",")
        For i = LBound(tmpDefineArray) To UBound(tmpDefineArray)
            tmpTrimmedValue = Trim(tmpDefineArray(i))
            If Len(tmpTrimmedValue) > 0 Then
                ' Collection のキーに値を追加（重複はエラーになるため On Error Resume Next で無視）
                On Error Resume Next
                colUniqueList.Add tmpTrimmedValue, tmpTrimmedValue
                On Error GoTo 0 ' エラー処理を通常に戻す
            End If
        Next i
    End If

    ' dataRange のセル値を配列として一括取得・処理
    tmpDataArray = dataRange.value
    If IsArray(tmpDataArray) Then
        For r = LBound(tmpDataArray, 1) To UBound(tmpDataArray, 1)
            For c = LBound(tmpDataArray, 2) To UBound(tmpDataArray, 2)
                tmpVal = Trim(CStr(tmpDataArray(r, c)))
                If Len(tmpVal) > 0 Then
                    On Error Resume Next
                    colUniqueList.Add tmpVal, tmpVal
                    On Error GoTo 0
                End If
            Next c
        Next r
    Else
        ' 単一セル（配列でない）対応
        tmpVal = Trim(CStr(tmpDataArray))
        If Len(tmpVal) > 0 Then
            On Error Resume Next
            colUniqueList.Add tmpVal, tmpVal
            On Error GoTo 0
        End If
    End If

    ' Collection のアイテムをカンマ区切りの文字列として組み立てる
    tmpDelimiter = ""
    For Each tmpItem In colUniqueList
        CreateValidationListString = CreateValidationListString & tmpDelimiter & tmpItem
        tmpDelimiter = ","
    Next tmpItem

End Function


' ■ シート初期化 - セル配置
Public Sub ResetHorizontalAlignment(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    
    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' シート全体のセル配置をリセット
    With ws.Cells
        .HorizontalAlignment = xlGeneral
    End With
    
    ' インデントを再設定
    ws.Range("B2").IndentLevel = 1
    
    ' タイトル行の配置
    ws.Range(cfg.COL_LAST_LABEL & cfg.ROW_TITLE).HorizontalAlignment = xlRight
    
    ' コントロール行入力文字列の配置
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    
    ' ヘッダー１の配置
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenter
    ws.Range(cfg.COL_TASK_TEXT_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_TEXT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlLeft
    ws.Range(cfg.COL_TASK_COUNT_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_TASK_PROG_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_WGT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_PLANNED_END_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_ACTUAL_END_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    
    ' ヘッダー２の配置
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenter
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_OPT_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_L1_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_TASK_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_WBS_ID_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_TEXT_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenterAcrossSelection
    
    ' データ終了行とその次の行の配置
    ws.Range(cfg.COL_CHK_LABEL & (lngEndRow + 1) & ":" & cfg.COL_LAST_LABEL & (lngEndRow + 2)).HorizontalAlignment = xlCenter
    
    ' 先頭～レベル・タスク列までの配置
    ws.Range(cfg.COL_ERR_LABEL & lngStartRow & ":" & cfg.COL_TASK_LABEL & lngEndRow).HorizontalAlignment = xlCenter
    
    ' タスク集計合計～カテゴリ2の配置
    ws.Range(cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_CATEGORY2_LABEL & lngEndRow).HorizontalAlignment = xlCenter

    ' 成果物の配置
    ws.Range(cfg.COL_OUTPUT_LABEL & lngStartRow & ":" & cfg.COL_OUTPUT_LABEL & lngEndRow).HorizontalAlignment = xlLeft

End Sub


' ■ シート初期化 - フォーム
Public Sub ResetExecuteForm(ws As Worksheet, Optional blnShouldClearOptMemory As Boolean = False)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim lngRowCount As Long
    Dim varChkArray() As Variant
    Dim varOptArray() As Variant
    ' - 画面配置コントロール
    Dim shpExe1ComboBox As Shape
    Dim shpExe1Button As Shape
    Dim shpReset1Button As Shape
    Dim shpExe2ComboBox As Shape
    Dim shpExe2Button As Shape
    Dim shpReset2Button As Shape
    ' - 実行1コンボボックスの位置とサイズを計算するための変数
    Dim dblExe1ComboBoxLeft As Double
    Dim dblExe1ComboBoxTop As Double
    Dim dblExe1ComboBoxWidth As Double
    Dim dblExe1ComboBoxHeight As Double
    ' - 実行1ボタンの位置とサイズを計算するための変数
    Dim dblExe1ButtonLeft As Double
    Dim dblExe1ButtonTop As Double
    Dim dblExe1ButtonWidth As Double
    Dim dblExe1ButtonHeight As Double
    ' - リセット1ボタンの位置とサイズを計算するための変数
    Dim dblReset1ButtonLeft As Double
    Dim dblReset1ButtonTop As Double
    Dim dblReset1ButtonWidth As Double
    Dim dblReset1ButtonHeight As Double
    ' - 実行2コンボボックスの位置とサイズを計算するための変数
    Dim dblExe2ComboBoxLeft As Double
    Dim dblExe2ComboBoxTop As Double
    Dim dblExe2ComboBoxWidth As Double
    Dim dblExe2ComboBoxHeight As Double
    ' - 実行2ボタンの位置とサイズを計算するための変数
    Dim dblExe2ButtonLeft As Double
    Dim dblExe2ButtonTop As Double
    Dim dblExe2ButtonWidth As Double
    Dim dblExe2ButtonHeight As Double
    ' 一時変数定義
    Dim r As Long
    Dim tmpVar As Variant

    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' 実行1コンボボックスの位置とサイズを計算
    dblExe1ComboBoxLeft = ws.Cells(cfg.ROW_CTRL1, cfg.COL_L1).Left
    dblExe1ComboBoxTop = ws.Cells(cfg.ROW_CTRL1, cfg.COL_L1).Top
    dblExe1ComboBoxWidth = cfg.WIDTH_EXE1_COMBOBOX
    dblExe1ComboBoxHeight = ws.Cells(cfg.ROW_CTRL1, cfg.COL_L1).Height
    ' 同じ名前の実行1コンボボックスが存在するか確認
    On Error Resume Next
    Set shpExe1ComboBox = ws.Shapes(cfg.NAME_EXE1_COMBOBOX)
    On Error GoTo 0
    ' 実行1コンボボックスが存在する場合、削除
    If Not shpExe1ComboBox Is Nothing Then
        shpExe1ComboBox.Delete
    End If
    ' 実行1コンボボックスを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpExe1ComboBox = ws.Shapes.AddFormControl(xlDropDown, dblExe1ComboBoxLeft, dblExe1ComboBoxTop, dblExe1ComboBoxWidth, dblExe1ComboBoxHeight)
    shpExe1ComboBox.Name = cfg.NAME_EXE1_COMBOBOX
    With shpExe1ComboBox.ControlFormat
        .AddItem "すべて再計算"
        .AddItem "オートフィルターをリセット"
        .AddItem "階層番号で全体をソート"
        .AddItem "書式・入力規則をリセット"
        .AddItem "入力フォームをリセット"
        .AddItem "エラーチェック"
    End With
    With ws.DropDowns(cfg.NAME_EXE1_COMBOBOX)
        .ListIndex = 1
    End With
    
    ' 実行1ボタンの位置とサイズを計算
    dblExe1ButtonLeft = dblExe1ComboBoxLeft + dblExe1ComboBoxWidth
    dblExe1ButtonTop = dblExe1ComboBoxTop
    dblExe1ButtonWidth = cfg.WIDTH_EXE1_BUTTON
    dblExe1ButtonHeight = dblExe1ComboBoxHeight
    ' 同じ名前の実行1ボタンが存在するか確認
    On Error Resume Next
    Set shpExe1Button = ws.Shapes(cfg.NAME_EXE1_BUTTON)
    On Error GoTo 0
    ' 実行1ボタンが存在する場合、削除
    If Not shpExe1Button Is Nothing Then
        shpExe1Button.Delete
    End If
    ' 実行1ボタンを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpExe1Button = ws.Shapes.AddFormControl(xlButtonControl, dblExe1ButtonLeft, dblExe1ButtonTop, dblExe1ButtonWidth, dblExe1ButtonHeight)
    shpExe1Button.Name = cfg.NAME_EXE1_BUTTON
    With ws.Buttons(cfg.NAME_EXE1_BUTTON)
        .Characters.Text = "実行"
    End With
    shpExe1Button.OnAction = "wbsui.Exe1ButtonClick"
    
     ' リセット1ボタンの位置とサイズを計算
    dblReset1ButtonLeft = dblExe1ComboBoxLeft + dblExe1ComboBoxWidth + dblExe1ButtonWidth
    dblReset1ButtonTop = dblExe1ComboBoxTop
    dblReset1ButtonWidth = cfg.WIDTH_RESET1_BUTTON
    dblReset1ButtonHeight = dblExe1ComboBoxHeight
    ' 同じ名前のリセット1ボタンが存在するか確認
    On Error Resume Next
    Set shpReset1Button = ws.Shapes(cfg.NAME_RESET1_BUTTON)
    On Error GoTo 0
    ' リセット1ボタンが存在する場合、削除
    If Not shpReset1Button Is Nothing Then
        shpReset1Button.Delete
    End If
    ' リセット1ボタンを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpReset1Button = ws.Shapes.AddFormControl(xlButtonControl, dblReset1ButtonLeft, dblReset1ButtonTop, dblReset1ButtonWidth, dblReset1ButtonHeight)
    shpReset1Button.Name = cfg.NAME_RESET1_BUTTON
    With ws.Buttons(cfg.NAME_RESET1_BUTTON)
        .Characters.Text = "リセット"
    End With
    shpReset1Button.OnAction = "wbsui.Reset1ButtonClick"
    
    ' 実行2コンボボックスの位置とサイズを計算
    dblExe2ComboBoxLeft = ws.Cells(cfg.ROW_CTRL2, cfg.COL_L1).Left
    dblExe2ComboBoxTop = ws.Cells(cfg.ROW_CTRL2, cfg.COL_L1).Top
    dblExe2ComboBoxWidth = cfg.WIDTH_EXE2_COMBOBOX
    dblExe2ComboBoxHeight = ws.Cells(cfg.ROW_CTRL2, cfg.COL_L1).Height
    ' 同じ名前の実行2コンボボックスが存在するか確認
    On Error Resume Next
    Set shpExe2ComboBox = ws.Shapes(cfg.NAME_EXE2_COMBOBOX)
    On Error GoTo 0
    ' 実行2コンボボックスが存在する場合、削除
    If Not shpExe2ComboBox Is Nothing Then
        shpExe2ComboBox.Delete
    End If
    ' 実行2コンボボックスを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpExe2ComboBox = ws.Shapes.AddFormControl(xlDropDown, dblExe2ComboBoxLeft, dblExe2ComboBoxTop, dblExe2ComboBoxWidth, dblExe2ComboBoxHeight)
    shpExe2ComboBox.Name = cfg.NAME_EXE2_COMBOBOX
    With shpExe2ComboBox.ControlFormat
        .AddItem "【OPT】 選択した行の下に一行追加"
        .AddItem "【OPT】 選択した階層番号の末尾を＋１"
        .AddItem "【OPT】 選択した階層番号の末尾を－１"
        .AddItem "【CHK】 チェックした２箇所の階層番号の末尾番号を交換　※ チェックが同階層である階層行の２箇所でなかった場合は不可 ※"
        .AddItem "【CHK】 チェックした行を削除　※ 子階層や子タスクがある場合は不可 ※"
    End With
    With ws.DropDowns(cfg.NAME_EXE2_COMBOBOX)
        .ListIndex = 1
    End With
    
    ' 実行2ボタンの位置とサイズを計算
    dblExe2ButtonLeft = dblExe2ComboBoxLeft + dblExe2ComboBoxWidth
    dblExe2ButtonTop = dblExe2ComboBoxTop
    dblExe2ButtonWidth = cfg.WIDTH_EXE2_BUTTON
    dblExe2ButtonHeight = dblExe2ComboBoxHeight
    ' 同じ名前の実行2ボタンが存在するか確認
    On Error Resume Next
    Set shpExe2Button = ws.Shapes(cfg.NAME_EXE2_BUTTON)
    On Error GoTo 0
    ' 実行2ボタンが存在する場合、削除
    If Not shpExe2Button Is Nothing Then
        shpExe2Button.Delete
    End If
    ' 実行2ボタンを新たに作成（サイズが変わる可能性があるため、毎回、作り直す）
    Set shpExe2Button = ws.Shapes.AddFormControl(xlButtonControl, dblExe2ButtonLeft, dblExe2ButtonTop, dblExe2ButtonWidth, dblExe2ButtonHeight)
    shpExe2Button.Name = cfg.NAME_EXE2_BUTTON
    With ws.Buttons(cfg.NAME_EXE2_BUTTON)
        .Characters.Text = "実行"
    End With
    shpExe2Button.OnAction = "wbsui.Exe2ButtonClick"
    
    ' リセット2ボタンの位置とサイズを計算
    dblReset2ButtonLeft = dblExe2ComboBoxLeft + dblExe2ComboBoxWidth + dblExe2ButtonWidth
    dblReset2ButtonTop = dblExe2ComboBoxTop
    dblReset2ButtonWidth = cfg.WIDTH_RESET2_BUTTON
    dblReset2ButtonHeight = dblExe2ComboBoxHeight
    ' 同じ名前のリセット2ボタンが存在するか確認
    On Error Resume Next
    Set shpReset2Button = ws.Shapes(cfg.NAME_RESET2_BUTTON)
    On Error GoTo 0
    ' リセット2ボタンが存在する場合、削除
    If Not shpReset2Button Is Nothing Then
        shpReset2Button.Delete
    End If
    ' リセット2ボタンを新たに作成
    Set shpReset2Button = ws.Shapes.AddFormControl(xlButtonControl, dblReset2ButtonLeft, dblReset2ButtonTop, dblReset2ButtonWidth, dblReset2ButtonHeight)
    shpReset2Button.Name = cfg.NAME_RESET2_BUTTON
    With ws.Buttons(cfg.NAME_RESET2_BUTTON)
        .Characters.Text = "リセット"
    End With
    shpReset2Button.OnAction = "wbsui.Reset2ButtonClick"
    
     ' 対象行数を取得
    lngRowCount = lngEndRow - lngStartRow + 1
    
    ' 一括書き込みのためのデータを用意
    ReDim varChkArray(1 To lngRowCount, 1 To 1)
    ReDim varOptArray(1 To lngRowCount, 1 To 1)
    
    ' 値を配列に格納（固定値）
    For r = 1 To lngRowCount
        varChkArray(r, 1) = cfg.CHK_MARK_F
        varOptArray(r, 1) = cfg.OPT_MARK_F
    Next r
    
    ' 結果を書き込み
    ws.Range(ws.Cells(lngStartRow, cfg.COL_CHK), ws.Cells(lngEndRow, cfg.COL_CHK)).value = varChkArray
    ws.Range(ws.Cells(lngStartRow, cfg.COL_OPT), ws.Cells(lngEndRow, cfg.COL_OPT)).value = varOptArray
    
    ' 引数の指定でOPTのメモリセルをクリアする必要がある場合
    If blnShouldClearOptMemory = True Then
        ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_DATA_START).ClearContents
    End If
    
    ' 最後に選択したOPTを反映
    tmpVar = ws.Cells(cfg.ROW_DATA_START, cfg.COL_OPT).value
    If tmpVar <> "" And _
            IsNumeric(tmpVar) And _
            tmpVar >= lngStartRow And _
            tmpVar <= lngEndRow Then
        ws.Cells(tmpVar, cfg.COL_OPT).value = cfg.OPT_MARK_T
        
    End If

End Sub


' ■ シート初期化 - タイトル行
Public Sub ResetTitleRow(ws As Worksheet)

    ' 変数定義
    Dim rngTargetCell As Range
    Dim strGitHubURL As String
    Dim strCommentText As String

    ' タイトル行のデータを一旦削除
    ws.Rows(cfg.ROW_TITLE).ClearContents

    ' シート名をセット
    ws.Range(cfg.COL_ERR_LABEL & cfg.ROW_TITLE).value = ws.Name
    
    ' 対象のセルを設定
    Set rngTargetCell = ws.Range(cfg.COL_LAST_LABEL & cfg.ROW_TITLE)

    ' GitHub の URL
    strGitHubURL = "https://github.com/H16K148/wbs-template-xlsm"

    ' セルに文字列を入力
    rngTargetCell.value = strGitHubURL

    ' ハイパーリンクを設定
    ws.Hyperlinks.Add Anchor:=rngTargetCell, Address:=strGitHubURL, TextToDisplay:=strGitHubURL

    ' 文字サイズを 8 に設定
    rngTargetCell.Font.Size = 8

    ' 配置を右下寄せに設定
    rngTargetCell.HorizontalAlignment = xlRight
    rngTargetCell.VerticalAlignment = xlBottom
    
    ' コメントの内容
    strCommentText = "バージョン：v" & cfg.APP_VERSION & vbCrLf & "参考情報：" & vbCrLf & "シートロックの解除に関する重要な情報は、リンク先のドキュメントに記載されています。"
    
    ' セルにコメントを追加または編集
    On Error Resume Next
    rngTargetCell.Comment.Delete
    On Error GoTo 0
    rngTargetCell.AddComment strCommentText
    With rngTargetCell.Comment
        .Visible = False
        .Shape.Width = 120
        .Shape.Height = 75
    End With

End Sub


' ■ 基本数式のリセット
Public Sub ResetBasicFormulas(ws As Worksheet)
    
    wbslib.SetFormulaForWbsIdx ws
    wbslib.SetFormulaForWbsCnt ws
    wbslib.SetFormulaForWbsId ws
    wbslib.SetFormulaForLevel ws
    wbslib.SetFormulaForFlgT ws
    wbslib.SetFormulaForFlgIC ws
    wbslib.SetFormulaForFlgPE ws
    wbslib.SetFormulaForFlgCE ws
    
End Sub


' ■ 集計数式のリセット
Public Sub ResetAggregateFormulas(ws As Worksheet)

    wbslib.SetFormulaForPlannedEffort ws
    wbslib.SetFormulaForActualCompletedEffort ws
    wbslib.SetFormulaForActualRemainingEffort ws
    wbslib.SetFormulaForTaskProgressRate ws
    wbslib.SetFormulaForEffortProgressRate ws
    wbslib.SetFormulaForTaskCount ws
    wbslib.SetFormulaForTaskCompCount ws

End Sub


' ■ オートフィルターのリセット
Public Sub ResetAutoFilter(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    
    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' いったん解除
    ws.AutoFilterMode = False
    
    ' 設定
    ws.Range(cfg.COL_L1_LABEL & cfg.ROW_DATA_START & ":" & cfg.COL_CATEGORY2_LABEL & lngEndRow).AutoFilter

End Sub


' ■ 初期値のセット
Public Sub SetInitialValue(ws As Worksheet)

    ' 変数定義
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' 一時変数定義
    Dim tmpRngTarget As Range
    Dim tmpVarTarget As Variant
    
    ' 開始行と終了行に値をセット
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' 開始行と終了行が見つからなければ終了
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' WBSステータス行
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_WBS_STATUS), ws.Cells(lngEndRow, cfg.COL_WBS_STATUS))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget
    
    ' 項目加重行
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_WGT), ws.Cells(lngEndRow, cfg.COL_TASK_WGT))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = 1
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

    ' 組織行
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_TEAM_SLCT), ws.Cells(lngEndRow, cfg.COL_TEAM_SLCT))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

    ' 担当行
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_PERSON_SLCT), ws.Cells(lngEndRow, cfg.COL_PERSON_SLCT))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

    ' カテゴリ1行
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_CATEGORY1), ws.Cells(lngEndRow, cfg.COL_CATEGORY1))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

    ' カテゴリ2行
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_CATEGORY2), ws.Cells(lngEndRow, cfg.COL_CATEGORY2))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

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
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' タイトル、基本数式のリセット
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' 集計数式のリセット
            wbsui.ResetAggregateFormulas ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 2
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' タイトル、基本数式のリセット
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' オートフィルターをリセット
            wbsui.ResetAutoFilter ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 3
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' タイトル、基本数式のリセット
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' ソートを実施
            wbslib.ExecSortWbsRange ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 4
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' タイトル、基本数式のリセット
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' 書式・入力規則をリセット
            wbsui.ResetConditionalFormatting ws
            wbsui.ResetDataValidation ws
            wbsui.ResetHorizontalAlignment ws
            wbsui.ResetAutoFilter ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 5
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' タイトル、基本数式のリセット
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' 入力フォームをリセット
            wbsui.ResetExecuteForm ws, True
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 6
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' タイトル、基本数式のリセット
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' エラーチェック
            wbslib.ExecCheckWbsErrors ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
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
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' タイトル、基本数式のリセット
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' 選択した行の下に一行追加
            wbslib.ExecInsertRowBelowSelection ws
            
            ' 初期値入力
            wbsui.SetInitialValue ws
            
            ' 基本数式リセット
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' 集計数式リセット
            wbsui.ResetAggregateFormulas ws
            
            ' エラーチェック
            wbslib.ExecCheckWbsErrors ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 2
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' タイトル、基本数式のリセット
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' 選択した行の末尾のインデックスを+1
            wbslib.ExecIncrementSelectedLastLevel ws
            
            ' 初期値入力
            wbsui.SetInitialValue ws
            
            ' 基本数式リセット
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' 集計数式リセット
            wbsui.ResetAggregateFormulas ws
            
            ' エラーチェック
            wbslib.ExecCheckWbsErrors ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 3
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' タイトル、基本数式のリセット
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' 選択した行の末尾のインデックスを-1
            wbslib.ExecDecrementSelectedLastLevel ws
            
            ' 初期値入力
            wbsui.SetInitialValue ws
            
            ' 基本数式リセット
            wbsui.ResetBasicFormulas ws
            
            ' 式の更新時、一時的自動計算を行う
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' 集計数式リセット
            wbsui.ResetAggregateFormulas ws
            
            ' エラーチェック
            wbslib.ExecCheckWbsErrors ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
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


' □ イベントコードを挿入する
' 　 次の設定が必要か
' 　 ・ツール > 参照設定 > Microsoft Visual Basic for Applications Extensibillity 5.3 にチェック（VBIDE への参照追加）
' 　 ・セキュリティセンター > マクロの設定 > 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」
Private Sub InitDoubleClickHandlerToSheet(ws As Worksheet)

    ' 変数定義
    Dim vbComp As VBIDE.VBComponent
    Dim codeLines As New Collection

    ' 対象のシートモジュールを取得
    Set vbComp = ThisWorkbook.VBProject.VBComponents(ws.CodeName)

    codeLines.Add "' ------------------------------------------------------------------------------"
    codeLines.Add "' Copyright 2025 Hiroki Chiba <h16k148@gmail.com>"
    codeLines.Add "'"
    codeLines.Add "' Licensed under the Apache License, Version 2.0 (the ""License"");"
    codeLines.Add "' you may not use this file except in compliance with the License."
    codeLines.Add "' You may obtain a copy of the License at"
    codeLines.Add "'"
    codeLines.Add "'     http://www.apache.org/licenses/LICENSE-2.0"
    codeLines.Add "'"
    codeLines.Add "' Unless required by applicable law or agreed to in writing, software"
    codeLines.Add "' distributed under the License is distributed on an ""AS IS"" BASIS,"
    codeLines.Add "' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied."
    codeLines.Add "' See the License for the specific language governing permissions and"
    codeLines.Add "' limitations under the License."
    codeLines.Add "' ------------------------------------------------------------------------------"
    codeLines.Add ""
    codeLines.Add ""
    codeLines.Add "' ◆ CHK と OPT のダブルクリックイベントを実装"
    codeLines.Add "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)"
    codeLines.Add "    ' 変数定義"
    codeLines.Add "    Dim lngClickedColumn As Long"
    codeLines.Add ""
    codeLines.Add "    ' 列番号を取得"
    codeLines.Add "    lngClickedColumn = Target.Column"
    codeLines.Add ""
    codeLines.Add "    ' ガード条件（対象列以外はデフォルト動作のままで終了）"
    codeLines.Add "    If lngClickedColumn <> cfg.COL_CHK And lngClickedColumn <> cfg.COL_OPT Then"
    codeLines.Add "        Exit Sub"
    codeLines.Add "    End If"
    codeLines.Add ""
    codeLines.Add "    ' 現在のシートを取得"
    codeLines.Add "    Dim ws As Worksheet"
    codeLines.Add "    Set ws = Me"
    codeLines.Add ""
    codeLines.Add "    ' 変数定義"
    codeLines.Add "    Dim lngClickedRow As Long"
    codeLines.Add "    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long"
    codeLines.Add "    Dim varClicked As Variant"
    codeLines.Add "    ' 一時変数定義"
    codeLines.Add "    Dim rngFoundCell As Range"
    codeLines.Add ""
    codeLines.Add "    ' 行番号を取得"
    codeLines.Add "    lngClickedRow = Target.row"
    codeLines.Add ""
    codeLines.Add "    ' 開始行と終了行を取得"
    codeLines.Add "    varRangeRows = wbslib.FindDataRangeRows(ws)"
    codeLines.Add "    lngStartRow = varRangeRows(0)"
    codeLines.Add "    lngEndRow = varRangeRows(1)"
    codeLines.Add ""
    codeLines.Add "    ' 開始行と終了行が見つからなければ終了"
    codeLines.Add "    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub"
    codeLines.Add ""
    codeLines.Add "    ' ガード条件（行番号が指定範囲外の場合は終了）"
    codeLines.Add "    If lngClickedRow < lngStartRow Or lngClickedRow > lngEndRow Then"
    codeLines.Add "        Exit Sub"
    codeLines.Add "    End If"
    codeLines.Add ""
    codeLines.Add "    ' CHK 処理"
    codeLines.Add "    If lngClickedColumn = cfg.COL_CHK Then"
    codeLines.Add "        varClicked = ws.Cells(lngClickedRow, cfg.COL_CHK).value"
    codeLines.Add ""
    codeLines.Add "        Application.ScreenUpdating = False"
    codeLines.Add "        Application.Calculation = xlCalculationManual"
    codeLines.Add "        Application.EnableEvents = False"
    codeLines.Add ""
    codeLines.Add "        If varClicked = cfg.CHK_MARK_T Then"
    codeLines.Add "            ws.Cells(lngClickedRow, cfg.COL_CHK).value = cfg.CHK_MARK_F"
    codeLines.Add "        Else"
    codeLines.Add "            ws.Cells(lngClickedRow, cfg.COL_CHK).value = cfg.CHK_MARK_T"
    codeLines.Add "        End If"
    codeLines.Add ""
    codeLines.Add "        Application.ScreenUpdating = True"
    codeLines.Add "        Application.Calculation = xlCalculationAutomatic"
    codeLines.Add "        Application.EnableEvents = True"
    codeLines.Add ""
    codeLines.Add "        Cancel = True"
    codeLines.Add "        Exit Sub"
    codeLines.Add "    End If"
    codeLines.Add ""
    codeLines.Add "    ' OPT 処理 (高速化版)"
    codeLines.Add "    If lngClickedColumn = cfg.COL_OPT Then"
    codeLines.Add "        ' クリックした値を取得"
    codeLines.Add "        varClicked = ws.Cells(lngClickedRow, cfg.COL_OPT).value"
    codeLines.Add "        If varClicked <> cfg.OPT_MARK_T Then ' クリックされたセルが cfg.OPT_MARK_T でない場合のみ処理"
    codeLines.Add "            ' lngStartRow から lngEndRow の範囲で cfg.OPT_MARK_T を持つ最初のセルを検索"
    codeLines.Add "            On Error Resume Next"
    codeLines.Add "            Set rngFoundCell = ws.Range(ws.Cells(lngStartRow, cfg.COL_OPT), ws.Cells(lngEndRow, cfg.COL_OPT)).Find(What:=cfg.OPT_MARK_T, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)"
    codeLines.Add "            On Error GoTo 0"
    codeLines.Add ""
    codeLines.Add "            Application.ScreenUpdating = False"
    codeLines.Add "            Application.Calculation = xlCalculationManual"
    codeLines.Add "            Application.EnableEvents = False"
    codeLines.Add ""
    codeLines.Add "            ' cfg.OPT_MARK_T を持つセルが見つかったら cfg.OPT_MARK_F に変更"
    codeLines.Add "            If Not rngFoundCell Is Nothing Then"
    codeLines.Add "                rngFoundCell.value = cfg.OPT_MARK_F"
    codeLines.Add "            End If"
    codeLines.Add ""
    codeLines.Add "            ' クリックされたセルの値を cfg.OPT_MARK_T に変更"
    codeLines.Add "            ws.Cells(lngClickedRow, cfg.COL_OPT).value = cfg.OPT_MARK_T"
    codeLines.Add "            ws.Cells(lngStartRow -1, cfg.COL_OPT).value = lngClickedRow"
    codeLines.Add ""
    codeLines.Add "            Application.ScreenUpdating = True"
    codeLines.Add "            Application.Calculation = xlCalculationAutomatic"
    codeLines.Add "            Application.EnableEvents = True"
    codeLines.Add ""
    codeLines.Add "            Cancel = True"
    codeLines.Add "            Exit Sub"
    codeLines.Add "        End If"
    codeLines.Add "    End If"
    codeLines.Add ""
    codeLines.Add "End Sub"
    codeLines.Add ""

    ' コードを挿入
    With vbComp.CodeModule
        For i = 1 To codeLines.Count
            .InsertLines .CountOfLines + 1, codeLines(i)
        Next i
    End With

End Sub

