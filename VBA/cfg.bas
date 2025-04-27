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

' �� �萔��`

' �����p�X���[�h
Public Const APP_PASSWORD As String = "h16k148"

' �o�[�W����
Public Const APP_VERSION As String = "0.1.2"

' �}�[�N�������`
Public Const CHK_MARK_T As String = "��"
Public Const CHK_MARK_F As String = "�E"
Public Const OPT_MARK_T As String = "��"
Public Const OPT_MARK_F As String = "�E"


' �X�e�[�^�X�������`�i������,���蒆,�ۗ�,�ڊǍ�,�I�グ,�p��,����,�폜�j
Public Const WBS_STATUS_NOT_STARTED As String = "������"
Public Const WBS_STATUS_IN_PROGRESS As String = "���蒆"
Public Const WBS_STATUS_ON_HOLD     As String = "�ۗ�"
Public Const WBS_STATUS_TRANSFERRED As String = "�ڊǍ�"
Public Const WBS_STATUS_SHELVED     As String = "�I�グ"
Public Const WBS_STATUS_REJECTED    As String = "�p��"
Public Const WBS_STATUS_COMPLETED   As String = "����"
Public Const WBS_STATUS_DELETED     As String = "�폜"


' �X�e�[�^�X�ꗗ������i�萔�̑g�ݍ��킹�j
Public Const WBS_STATUS_LIST As String = _
    WBS_STATUS_NOT_STARTED & "," & _
    WBS_STATUS_IN_PROGRESS & "," & _
    WBS_STATUS_ON_HOLD & "," & _
    WBS_STATUS_TRANSFERRED & "," & _
    WBS_STATUS_SHELVED & "," & _
    WBS_STATUS_REJECTED & "," & _
    WBS_STATUS_COMPLETED & "," & _
    WBS_STATUS_DELETED


' ���`�i���l�j
Public Const COL_KEY  As Long = 1                  '  A��F�@�\���FKEY �J����
Public Const COL_KEY_LABEL As String = "A"

Public Const COL_ERR  As Long = 2                  '  B��F�@�\���FERROR���J����
Public Const COL_ERR_LABEL As String = "B"

Public Const COL_CHK  As Long = 3                  '  C��F�@�\���FCHK�J����
Public Const COL_CHK_LABEL As String = "C"

Public Const COL_OPT  As Long = 4                  '  D��F�@�\���FOPT�J����
Public Const COL_OPT_LABEL As String = "D"

Public Const COL_L1   As Long = 5                  '  E��F�@�\���FL1  �ԍ��̓��̓J����
Public Const COL_L1_LABEL As String = "E"

Public Const COL_L2   As Long = 6                  '  F��F�@�\���FL2  �ԍ��̓��̓J����
Public Const COL_L2_LABEL As String = "F"

Public Const COL_L3   As Long = 7                  '  G��F�@�\���FL3  �ԍ��̓��̓J����
Public Const COL_L3_LABEL As String = "G"

Public Const COL_L4   As Long = 8                  '  H��F�@�\���FL4  �ԍ��̓��̓J����
Public Const COL_L4_LABEL As String = "H"

Public Const COL_L5   As Long = 9                  '  I��F�@�\���FL5  �ԍ��̓��̓J����
Public Const COL_L5_LABEL As String = "I"

Public Const COL_TASK As Long = 10                 '  J��F�@�\���FTASK�ԍ��̓��̓J����
Public Const COL_TASK_LABEL As String = "J"

Public Const COL_WBS_IDX As Long = 11              '  K��F��\���FWBS_IDX�p�J�����F��
Public Const COL_WBS_IDX_LABEL As String = "K"

Public Const COL_WBS_CNT As Long = 12              '  L��F��\���FWBS_CNT�p�J�����F��
Public Const COL_WBS_CNT_LABEL As String = "L"

Public Const COL_LEVEL As Long = 13                '  M��F��\���FWBS���x���J�����F��
Public Const COL_LEVEL_LABEL As String = "M"

Public Const COL_FLG_T As Long = 14                '  N��F��\���FWBS�^�X�N����J�����F��
Public Const COL_FLG_T_LABEL As String = "N"

Public Const COL_FLG_IC As Long = 15               '  O��F��\���F�v�Z�Ώ۔���iIncludeInCalculation�j�F��
Public Const COL_FLG_IC_LABEL As String = "O"

Public Const COL_FLG_PE As Long = 16               '  P��F��\���FWBS�e�L������iParent Exist�j�F�����Œ�l
Public Const COL_FLG_PE_LABEL As String = "P"

Public Const COL_FLG_CE As Long = 17               '  Q��F��\���FWBS�q�L������iChild Exist�j�F�����Œ�l
Public Const COL_FLG_CE_LABEL As String = "Q"

Public Const COL_WBS_ID As Long = 18               '  R��F�@�\���FWBS_ID�J�����F��
Public Const COL_WBS_ID_LABEL As String = "R"

Public Const COL_L1_TEXT As Long = 19              '  S��F�@�\���FL1  �e�L�X�g�J����
Public Const COL_L1_TEXT_LABEL As String = "S"

Public Const COL_L2_TEXT As Long = 20              '  T��F�@�\���FL2  �e�L�X�g�J����
Public Const COL_L2_TEXT_LABEL As String = "T"

Public Const COL_L3_TEXT As Long = 21              '  U��F�@�\���FL3  �e�L�X�g�J����
Public Const COL_L3_TEXT_LABEL As String = "U"

Public Const COL_L4_TEXT As Long = 22              '  V��F�@�\���FL4  �e�L�X�g�J����
Public Const COL_L4_TEXT_LABEL As String = "V"

Public Const COL_L5_TEXT As Long = 23              '  W��F�@�\���FL5  �e�L�X�g�J����
Public Const COL_L5_TEXT_LABEL As String = "W"

Public Const COL_TASK_TEXT As Long = 24            '  X��F�@�\���FTASK�e�L�X�g�J����
Public Const COL_TASK_TEXT_LABEL As String = "X"

Public Const COL_TEXT As Long = 25                 '  Y��F�@�\���F�e�L�X�g�J����
Public Const COL_TEXT_LABEL As String = "Y"

Public Const COL_TASK_COUNT As Long = 26           '  Z��F�@�\���FTASK�v�J�����F�����Œ�l
Public Const COL_TASK_COUNT_LABEL As String = "Z"

Public Const COL_TASK_COMP_COUNT As Long = 27      ' AA��F�@�\���FTASK���J�����F�����Œ�l
Public Const COL_TASK_COMP_COUNT_LABEL As String = "AA"

Public Const COL_WBS_STATUS As Long = 28           ' AB��F�@�\���FWBS��ԃJ����
Public Const COL_WBS_STATUS_LABEL As String = "AB"

Public Const COL_EFFORT_PROG As Long = 29          ' AC��F�@�\���F�H���i�����J�����F�����Œ�l
Public Const COL_EFFORT_PROG_LABEL As String = "AC"

Public Const COL_TASK_PROG As Long = 30            ' AD��F�@�\���F���ڏ������J�����F�����Œ�l
Public Const COL_TASK_PROG_LABEL As String = "AD"

Public Const COL_TASK_WGT As Long = 31             ' AE��F�@�\���F���ډ��d�J����
Public Const COL_TASK_WGT_LABEL As String = "AE"

Public Const COL_TEAM_SLCT As Long = 32            ' AF��F�@�\���F�g�D�I���J����
Public Const COL_TEAM_SLCT_LABEL As String = "AF"

Public Const COL_PERSON_SLCT As Long = 33          ' AG��F�@�\���F�S���I���J����
Public Const COL_PERSON_SLCT_LABEL As String = "AG"

Public Const COL_OUTPUT As Long = 34               ' AH��F�@�\���F���ʕ�
Public Const COL_OUTPUT_LABEL As String = "AH"

Public Const COL_PLANNED_EFF As Long = 35          ' AI��F�@�\���F�\��H���J�����F�����Œ�l
Public Const COL_PLANNED_EFF_LABEL As String = "AI"

Public Const COL_PLANNED_START As Long = 36        ' AJ��F�@�\���F�\��J�n�J����
Public Const COL_PLANNED_START_LABEL As String = "AJ"

Public Const COL_PLANNED_END As Long = 37          ' AK��F�@�\���F�\��I���J����
Public Const COL_PLANNED_END_LABEL As String = "AK"

Public Const COL_ACTUAL_REMAINING_EFF As Long = 38 ' AL��F�@�\���F���юc�H���J�����F�����Œ�l
Public Const COL_ACTUAL_REMAINING_EFF_LABEL As String = "AL"

Public Const COL_ACTUAL_COMPLETED_EFF As Long = 39 ' AM��F�@�\���F���эύH���J�����F�����Œ�l
Public Const COL_ACTUAL_COMPLETED_EFF_LABEL As String = "AM"

Public Const COL_ACTUAL_START As Long = 40         ' AN��F�@�\���F���ъJ�n�J����
Public Const COL_ACTUAL_START_LABEL As String = "AN"

Public Const COL_ACTUAL_END As Long = 41           ' AO��F�@�\���F���яI���J����
Public Const COL_ACTUAL_END_LABEL As String = "AO"

Public Const COL_CATEGORY1 As Long = 42            ' AP��F�@�\���F�J�e�S��1�J����
Public Const COL_CATEGORY1_LABEL As String = "AP"

Public Const COL_CATEGORY2 As Long = 43            ' AQ��F�@�\���F�J�e�S��2�J����
Public Const COL_CATEGORY2_LABEL As String = "AQ"

Public Const COL_LAST As Long = 44                 ' AR��F�@�\���F���l�J�����i�ŏI�j
Public Const COL_LAST_LABEL As String = "AR"


' �s��`
Public Const ROW_TITLE As Long = 2
Public Const ROW_CTRL1 As Long = 3                 ' �R���g���[����z�u����s1
Public Const ROW_CTRL2 As Long = 4                 ' �R���g���[����z�u����s2
Public Const ROW_HEADER1 As Long = 5
Public Const ROW_HEADER2 As Long = 6
Public Const ROW_DATA_START As Long = 7

' �R���g���[�����֘A��`
Public Const NAME_EXE1_COMBOBOX As String = "Execute1ComboBox"
Public Const NAME_EXE1_BUTTON   As String = "Execute1Button"
Public Const NAME_RESET1_BUTTON   As String = "Reset1Button"
Public Const NAME_EXE2_COMBOBOX As String = "Execute2ComboBox"
Public Const NAME_EXE2_BUTTON   As String = "Execute2Button"
Public Const NAME_RESET2_BUTTON   As String = "Reset2Button"

' �R���g���[���֘A��`
Public Const WIDTH_EXE1_COMBOBOX = 250
Public Const WIDTH_EXE1_BUTTON = 40
Public Const WIDTH_RESET1_BUTTON = 55
Public Const WIDTH_EXE2_COMBOBOX = 400
Public Const WIDTH_EXE2_BUTTON = 40
Public Const WIDTH_RESET2_BUTTON = 55


