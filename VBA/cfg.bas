Attribute VB_Name = "cfg"
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

' 解除パスワード
Public Const APP_PASSWORD As String = "h16k148"

' バージョン
Public Const APP_VERSION As String = "0.1.2"

' マーク文字列定義
Public Const CHK_MARK_T As String = "■"
Public Const CHK_MARK_F As String = "・"
Public Const OPT_MARK_T As String = "●"
Public Const OPT_MARK_F As String = "・"


' ステータス文字列定義（未着手,着手中,保留,移管済,棚上げ,却下,完了,削除）
Public Const WBS_STATUS_NOT_STARTED As String = "未着手"
Public Const WBS_STATUS_IN_PROGRESS As String = "着手中"
Public Const WBS_STATUS_ON_HOLD     As String = "保留"
Public Const WBS_STATUS_TRANSFERRED As String = "移管済"
Public Const WBS_STATUS_SHELVED     As String = "棚上げ"
Public Const WBS_STATUS_REJECTED    As String = "却下"
Public Const WBS_STATUS_COMPLETED   As String = "完了"
Public Const WBS_STATUS_DELETED     As String = "削除"


' ステータス一覧文字列（定数の組み合わせ）
Public Const WBS_STATUS_LIST As String = _
    WBS_STATUS_NOT_STARTED & "," & _
    WBS_STATUS_IN_PROGRESS & "," & _
    WBS_STATUS_ON_HOLD & "," & _
    WBS_STATUS_TRANSFERRED & "," & _
    WBS_STATUS_SHELVED & "," & _
    WBS_STATUS_REJECTED & "," & _
    WBS_STATUS_COMPLETED & "," & _
    WBS_STATUS_DELETED


' 列定義（数値）
Public Const COL_KEY  As Long = 1                  '  A列：　表示：KEY カラム
Public Const COL_KEY_LABEL As String = "A"

Public Const COL_ERR  As Long = 2                  '  B列：　表示：ERROR情報カラム
Public Const COL_ERR_LABEL As String = "B"

Public Const COL_CHK  As Long = 3                  '  C列：　表示：CHKカラム
Public Const COL_CHK_LABEL As String = "C"

Public Const COL_OPT  As Long = 4                  '  D列：　表示：OPTカラム
Public Const COL_OPT_LABEL As String = "D"

Public Const COL_L1   As Long = 5                  '  E列：　表示：L1  番号の入力カラム
Public Const COL_L1_LABEL As String = "E"

Public Const COL_L2   As Long = 6                  '  F列：　表示：L2  番号の入力カラム
Public Const COL_L2_LABEL As String = "F"

Public Const COL_L3   As Long = 7                  '  G列：　表示：L3  番号の入力カラム
Public Const COL_L3_LABEL As String = "G"

Public Const COL_L4   As Long = 8                  '  H列：　表示：L4  番号の入力カラム
Public Const COL_L4_LABEL As String = "H"

Public Const COL_L5   As Long = 9                  '  I列：　表示：L5  番号の入力カラム
Public Const COL_L5_LABEL As String = "I"

Public Const COL_TASK As Long = 10                 '  J列：　表示：TASK番号の入力カラム
Public Const COL_TASK_LABEL As String = "J"

Public Const COL_WBS_IDX As Long = 11              '  K列：非表示：WBS_IDX用カラム：式
Public Const COL_WBS_IDX_LABEL As String = "K"

Public Const COL_WBS_CNT As Long = 12              '  L列：非表示：WBS_CNT用カラム：式
Public Const COL_WBS_CNT_LABEL As String = "L"

Public Const COL_LEVEL As Long = 13                '  M列：非表示：WBSレベルカラム：式
Public Const COL_LEVEL_LABEL As String = "M"

Public Const COL_FLG_T As Long = 14                '  N列：非表示：WBSタスク判定カラム：式
Public Const COL_FLG_T_LABEL As String = "N"

Public Const COL_FLG_IC As Long = 15               '  O列：非表示：計算対象判定（IncludeInCalculation）：式
Public Const COL_FLG_IC_LABEL As String = "O"

Public Const COL_FLG_PE As Long = 16               '  P列：非表示：WBS親有無判定（Parent Exist）：式→固定値
Public Const COL_FLG_PE_LABEL As String = "P"

Public Const COL_FLG_CE As Long = 17               '  Q列：非表示：WBS子有無判定（Child Exist）：式→固定値
Public Const COL_FLG_CE_LABEL As String = "Q"

Public Const COL_WBS_ID As Long = 18               '  R列：　表示：WBS_IDカラム：式
Public Const COL_WBS_ID_LABEL As String = "R"

Public Const COL_L1_TEXT As Long = 19              '  S列：　表示：L1  テキストカラム
Public Const COL_L1_TEXT_LABEL As String = "S"

Public Const COL_L2_TEXT As Long = 20              '  T列：　表示：L2  テキストカラム
Public Const COL_L2_TEXT_LABEL As String = "T"

Public Const COL_L3_TEXT As Long = 21              '  U列：　表示：L3  テキストカラム
Public Const COL_L3_TEXT_LABEL As String = "U"

Public Const COL_L4_TEXT As Long = 22              '  V列：　表示：L4  テキストカラム
Public Const COL_L4_TEXT_LABEL As String = "V"

Public Const COL_L5_TEXT As Long = 23              '  W列：　表示：L5  テキストカラム
Public Const COL_L5_TEXT_LABEL As String = "W"

Public Const COL_TASK_TEXT As Long = 24            '  X列：　表示：TASKテキストカラム
Public Const COL_TASK_TEXT_LABEL As String = "X"

Public Const COL_TEXT As Long = 25                 '  Y列：　表示：テキストカラム
Public Const COL_TEXT_LABEL As String = "Y"

Public Const COL_TASK_COUNT As Long = 26           '  Z列：　表示：TASK計カラム：式→固定値
Public Const COL_TASK_COUNT_LABEL As String = "Z"

Public Const COL_TASK_COMP_COUNT As Long = 27      ' AA列：　表示：TASK完カラム：式→固定値
Public Const COL_TASK_COMP_COUNT_LABEL As String = "AA"

Public Const COL_WBS_STATUS As Long = 28           ' AB列：　表示：WBS状態カラム
Public Const COL_WBS_STATUS_LABEL As String = "AB"

Public Const COL_EFFORT_PROG As Long = 29          ' AC列：　表示：工数進捗率カラム：式→固定値
Public Const COL_EFFORT_PROG_LABEL As String = "AC"

Public Const COL_TASK_PROG As Long = 30            ' AD列：　表示：項目消化率カラム：式→固定値
Public Const COL_TASK_PROG_LABEL As String = "AD"

Public Const COL_TASK_WGT As Long = 31             ' AE列：　表示：項目加重カラム
Public Const COL_TASK_WGT_LABEL As String = "AE"

Public Const COL_TEAM_SLCT As Long = 32            ' AF列：　表示：組織選択カラム
Public Const COL_TEAM_SLCT_LABEL As String = "AF"

Public Const COL_PERSON_SLCT As Long = 33          ' AG列：　表示：担当選択カラム
Public Const COL_PERSON_SLCT_LABEL As String = "AG"

Public Const COL_OUTPUT As Long = 34               ' AH列：　表示：成果物
Public Const COL_OUTPUT_LABEL As String = "AH"

Public Const COL_PLANNED_EFF As Long = 35          ' AI列：　表示：予定工数カラム：式→固定値
Public Const COL_PLANNED_EFF_LABEL As String = "AI"

Public Const COL_PLANNED_START As Long = 36        ' AJ列：　表示：予定開始カラム
Public Const COL_PLANNED_START_LABEL As String = "AJ"

Public Const COL_PLANNED_END As Long = 37          ' AK列：　表示：予定終了カラム
Public Const COL_PLANNED_END_LABEL As String = "AK"

Public Const COL_ACTUAL_REMAINING_EFF As Long = 38 ' AL列：　表示：実績残工数カラム：式→固定値
Public Const COL_ACTUAL_REMAINING_EFF_LABEL As String = "AL"

Public Const COL_ACTUAL_COMPLETED_EFF As Long = 39 ' AM列：　表示：実績済工数カラム：式→固定値
Public Const COL_ACTUAL_COMPLETED_EFF_LABEL As String = "AM"

Public Const COL_ACTUAL_START As Long = 40         ' AN列：　表示：実績開始カラム
Public Const COL_ACTUAL_START_LABEL As String = "AN"

Public Const COL_ACTUAL_END As Long = 41           ' AO列：　表示：実績終了カラム
Public Const COL_ACTUAL_END_LABEL As String = "AO"

Public Const COL_CATEGORY1 As Long = 42            ' AP列：　表示：カテゴリ1カラム
Public Const COL_CATEGORY1_LABEL As String = "AP"

Public Const COL_CATEGORY2 As Long = 43            ' AQ列：　表示：カテゴリ2カラム
Public Const COL_CATEGORY2_LABEL As String = "AQ"

Public Const COL_LAST As Long = 44                 ' AR列：　表示：備考カラム（最終）
Public Const COL_LAST_LABEL As String = "AR"


' 行定義
Public Const ROW_TITLE As Long = 2
Public Const ROW_CTRL1 As Long = 3                 ' コントロールを配置する行1
Public Const ROW_CTRL2 As Long = 4                 ' コントロールを配置する行2
Public Const ROW_HEADER1 As Long = 5
Public Const ROW_HEADER2 As Long = 6
Public Const ROW_DATA_START As Long = 7

' コントロール名関連定義
Public Const NAME_EXE1_COMBOBOX As String = "Execute1ComboBox"
Public Const NAME_EXE1_BUTTON   As String = "Execute1Button"
Public Const NAME_RESET1_BUTTON   As String = "Reset1Button"
Public Const NAME_EXE2_COMBOBOX As String = "Execute2ComboBox"
Public Const NAME_EXE2_BUTTON   As String = "Execute2Button"
Public Const NAME_RESET2_BUTTON   As String = "Reset2Button"

' コントロール関連定義
Public Const WIDTH_EXE1_COMBOBOX = 250
Public Const WIDTH_EXE1_BUTTON = 40
Public Const WIDTH_RESET1_BUTTON = 55
Public Const WIDTH_EXE2_COMBOBOX = 400
Public Const WIDTH_EXE2_BUTTON = 40
Public Const WIDTH_RESET2_BUTTON = 55


