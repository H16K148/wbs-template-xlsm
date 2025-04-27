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


' �� �V�[�g���������[�U�[�t�H�[���\���p�}�N��
Sub ShowInitWBS()
    InitWBS.Show vbModeless
End Sub


' �� �V�[�g�����������܂�
Public Sub InitSheet(ws As Worksheet)
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' �x�[�X�f�U�C���𔽉f
    InitSheetBaseDesign ws
    
    ' �_�~�[�f�[�^�𓊓�
    InputDammyData ws
    
    ' �^�C�g���Z���̃��Z�b�g
    ResetTitleRow ws
    
    ' ��{�����̃��Z�b�g
    ResetBasicFormulas ws
    
    ' �W�v�����̃��Z�b�g
    ResetAggregateFormulas ws
    
    ' �����t�������̃��Z�b�g
    ResetConditionalFormatting ws
    
    ' �Z���z�u�̃��Z�b�g
    ResetHorizontalAlignment ws
    
    ' �f�[�^���͋K���̃��Z�b�g
    ResetDataValidation ws
        
    ' �t�H�[���֘A�̃��Z�b�g
    ResetExecuteForm ws
    
    ' �����l���Z�b�g
    SetInitialValue ws
    
    ' �I�[�g�t�B���^�[�̃��Z�b�g
    ResetAutoFilter ws
    
    ' �V�[�g�ɃC�x���g�R�[�h��ǉ��i�_�u���N���b�N�C�x���g�j
    InitDoubleClickHandlerToSheet ws
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub


' �� �_�~�[�f�[�^��o�^���܂��i�� �V�[�g���������ȊO�Ɏg�p���Ă͂����܂��� ���j
Private Sub InputDammyData(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' �ꎞ�ϐ���`
    Dim i As Long, j As Long, k As Long, l As Long, m As Long, n As Long
    Dim level1 As Long, level2 As Long, level3 As Long, level4 As Long, level5 As Long, taskCount As Long
    Dim tmpCurrentRow As Long
    Dim tmpEndFlg As Boolean
    
    level1 = 3
    level2 = 2    ' �e�K�w1�̉��ɍ쐬���鐔
    level3 = 2    ' �e�K�w2�̉��ɍ쐬���鐔
    level4 = 2    ' �e�K�w3�̉��ɍ쐬���鐔
    level5 = 2    ' �e�K�w4�̉��ɍ쐬���鐔
    taskCount = 3 ' �e�K�w�ɍ쐬����^�X�N��
    
    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �g�D���̓���
    ws.Cells(cfg.ROW_CTRL1, cfg.COL_EFFORT_PROG).value = "A��,B��,C��"
    
    ' �S�����̓���
    ws.Cells(cfg.ROW_CTRL2, cfg.COL_EFFORT_PROG).value = "������Y,��ؓ�Y,�����O�Y"
    
    ' �J�e�S��1�̓���
    ws.Cells(cfg.ROW_CTRL1, cfg.COL_CATEGORY2).value = "A,B,C"
    
    ' �J�e�S��2�̓���
    ws.Cells(cfg.ROW_CTRL2, cfg.COL_CATEGORY2).value = "D,E,F"
    
    ' �^�X�N�ԍ������
    tmpEndFlg = False            ' �����I���t���O
    tmpCurrentRow = lngStartRow  ' �J�n�s�����ݍs�Ƃ���
    ' �K�w1
    For i = 1 To level1
        ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
        If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
            tmpEndFlg = True
            Exit For
        End If
        ' �K�w1�̍s����
        ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
        If i = 1 Then
            ws.Cells(tmpCurrentRow, cfg.COL_L1_TEXT).value = "�K�w1 �e�L�X�g"
        End If
        tmpCurrentRow = tmpCurrentRow + 1
        ' �K�w1�̃^�X�N�s����
        For n = 1 To taskCount
            ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
            If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                tmpEndFlg = True
                Exit For
            End If
            ' �K�w1�^�X�N�̍s����
            ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
            ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
            If n = 1 Then
                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w1�^�X�N �e�L�X�g1"
                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_DELETED
            ElseIf n = 2 Then
                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w1�^�X�N �e�L�X�g2"
                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_REJECTED
            ElseIf n = 3 Then
                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w1�^�X�N �e�L�X�g3"
                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
            End If
            tmpCurrentRow = tmpCurrentRow + 1
        Next n
        ' �K�w2
        For j = 1 To level2
            ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
            If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                tmpEndFlg = True
                Exit For
            End If
            ' �K�w2�̍s����
            ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
            ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
            If j = 1 Then
                ws.Cells(tmpCurrentRow, cfg.COL_L2_TEXT).value = "�K�w2 �e�L�X�g"
            End If
            tmpCurrentRow = tmpCurrentRow + 1
            ' �K�w2�̃^�X�N�s����
            For n = 1 To taskCount
                ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
                If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                    tmpEndFlg = True
                    Exit For
                End If
                ' �K�w2�^�X�N�̍s����
                ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
                If n = 1 Then
                    ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w2�^�X�N �e�L�X�g1"
                    ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_ON_HOLD
                ElseIf n = 2 Then
                    ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w2�^�X�N �e�L�X�g2"
                    ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_SHELVED
                ElseIf n = 3 Then
                    ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w2�^�X�N �e�L�X�g3"
                    ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                End If
                tmpCurrentRow = tmpCurrentRow + 1
            Next n
            ' �K�w3
            For k = 1 To level3
                ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
                If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                    tmpEndFlg = True
                    Exit For
                End If
                ' �K�w3�̍s����
                ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                If k = 1 Then
                    ws.Cells(tmpCurrentRow, cfg.COL_L3_TEXT).value = "�K�w3 �e�L�X�g"
                End If
                tmpCurrentRow = tmpCurrentRow + 1
                ' �K�w3�̃^�X�N�s����
                For n = 1 To taskCount
                    ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
                    If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                        tmpEndFlg = True
                        Exit For
                    End If
                    ' �K�w3�^�X�N�̍s����
                    ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                    ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                    ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                    ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
                    If n = 1 Then
                        ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w3�^�X�N �e�L�X�g1"
                        ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_TRANSFERRED
                    ElseIf n = 2 Then
                        ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w3�^�X�N �e�L�X�g2"
                        ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_NOT_STARTED
                    ElseIf n = 3 Then
                        ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w3�^�X�N �e�L�X�g3"
                        ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                    End If
                    tmpCurrentRow = tmpCurrentRow + 1
                Next n
                ' �K�w4
                For l = 1 To level4
                    ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
                    If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                        tmpEndFlg = True
                        Exit For
                    End If
                    ' �K�w4�̍s����
                    ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                    ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                    ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                    ws.Cells(tmpCurrentRow, cfg.COL_L4).value = l
                    If l = 1 Then
                        ws.Cells(tmpCurrentRow, cfg.COL_L4_TEXT).value = "�K�w4 �e�L�X�g"
                    End If
                    tmpCurrentRow = tmpCurrentRow + 1
                    ' �K�w4�̃^�X�N�s����
                    For n = 1 To taskCount
                        ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
                        If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                            tmpEndFlg = True
                            Exit For
                        End If
                        ' �K�w4�^�X�N�̍s����
                        ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                        ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                        ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                        ws.Cells(tmpCurrentRow, cfg.COL_L4).value = l
                        ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
                        If n = 1 Then
                            ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w4�^�X�N �e�L�X�g1"
                            ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_COMPLETED
                        ElseIf n = 2 Then
                            ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w4�^�X�N �e�L�X�g2"
                            ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = cfg.WBS_STATUS_IN_PROGRESS
                        ElseIf n = 3 Then
                            ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w4�^�X�N �e�L�X�g3"
                            ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                        End If
                        tmpCurrentRow = tmpCurrentRow + 1
                    Next n
                    ' �K�w5
                    For m = 1 To level5
                        ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
                        If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                            tmpEndFlg = True
                            Exit For
                        End If
                        ' �K�w5�̍s����
                        ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                        ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                        ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                        ws.Cells(tmpCurrentRow, cfg.COL_L4).value = l
                        ws.Cells(tmpCurrentRow, cfg.COL_L5).value = m
                        If m = 1 Then
                            ws.Cells(tmpCurrentRow, cfg.COL_L5_TEXT).value = "�K�w5 �e�L�X�g"
                        End If
                        tmpCurrentRow = tmpCurrentRow + 1
                        ' �K�w5�̃^�X�N�s����
                        For n = 1 To taskCount
                            ' �K�[�h�����i�����I���t���O�������Ă��邩�A�ŏI�s�̏ꍇ�A�I���j
                            If tmpEndFlg = True Or tmpCurrentRow = lngEndRow Then
                                tmpEndFlg = True
                                Exit For
                            End If
                            ' �K�w5�^�X�N�̍s����
                            ws.Cells(tmpCurrentRow, cfg.COL_L1).value = i
                            ws.Cells(tmpCurrentRow, cfg.COL_L2).value = j
                            ws.Cells(tmpCurrentRow, cfg.COL_L3).value = k
                            ws.Cells(tmpCurrentRow, cfg.COL_L4).value = l
                            ws.Cells(tmpCurrentRow, cfg.COL_L5).value = m
                            ws.Cells(tmpCurrentRow, cfg.COL_TASK).value = n
                            If n = 1 Then
                                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w5�^�X�N �e�L�X�g1"
                                ws.Cells(tmpCurrentRow, cfg.COL_WBS_STATUS).value = "-"
                            ElseIf n = 2 Then
                                ws.Cells(tmpCurrentRow, cfg.COL_TASK_TEXT).value = "�K�w5�^�X�N �e�L�X�g2"
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


' �� �V�[�g������ - �x�[�X�f�U�C���i�Z���̃T�C�Y�A�S�̃t�H���g�A�Ȃǁj
Public Sub InitSheetBaseDesign(ws As Worksheet)
    
    ' �ϐ���`
    Dim win As Window
    Dim lngTitleRow As Long
    Dim lngDataStartRow As Long
    Dim lngDataEndRow As Long
    ' �ꎞ�ϐ���`
    Dim tmpWorksheet As Worksheet
    Dim tmpCharLength As Integer
    
    ' ������
    lngTitleRow = 2
    lngDataStartRow = 7
    lngDataEndRow = 219

    ' �V�[�g�S�̂̃t�H���g��ύX
    With ws.Cells
        .Font.Name = "Yu Gothic"                ' �t�H���g��
        .Font.Size = 9                          ' �t�H���g�T�C�Y
        .Font.Bold = False                      ' �����iTrue�ő����AFalse�Œʏ�j
        .Font.Italic = False                    ' �ΆiTrue�ŎΆAFalse�Œʏ�j
        .Font.Underline = xlUnderlineStyleNone  ' �����i�Ȃ��ɐݒ�j
        .VerticalAlignment = xlVAlignCenter     ' �Z���̏c������������
        .RowHeight = 18
    End With
    
    ' �V�[�g�̍s����ҏW
    ws.Rows(1).RowHeight = 3.75
    ws.Rows(2).RowHeight = 30
    
    ' �V�[�g�̗񕝂�ҏW
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
    
    ' �V�[�g���𑕏�
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").IndentLevel = 1
    
    ' �R���g���[���s����
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL1).value = "�S�́F"
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL2).value = "�I���F"
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL1).value = "�y�I������`�z�g�D�F"
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL2).value = "�y�I������`�z�S���F"
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL1).value = "�y�I������`�z�J�e�S��1�F"
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL2).value = "�y�I������`�z�J�e�S��2�F"
    
    ' �R���g���[���s���͕�����̑���
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    
    ' �w�b�_�[�P���������
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
    ws.Range(cfg.COL_TASK_COUNT_LABEL & cfg.ROW_HEADER1).value = "TASK�W�v"
    ws.Range(cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_HEADER1).value = "�H��"
    ws.Range(cfg.COL_TASK_PROG_LABEL & cfg.ROW_HEADER1).value = "����"
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER1).value = "�\��"
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER1).value = "����"
    
    ' �w�b�_�[�P������̑���
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenter
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER1).VerticalAlignment = xlBottom
    ws.Range(cfg.COL_TASK_COUNT_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_TASK_PROG_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_WGT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_PLANNED_END_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_ACTUAL_END_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & cfg.ROW_HEADER1).Font.Size = 7
    
    ' �w�b�_�[�Q���������
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER2).value = "D-Click!"
    ws.Range(cfg.COL_L1_LABEL & cfg.ROW_HEADER2).value = "�K�w�ԍ�"
    ws.Range(cfg.COL_WBS_ID_LABEL & cfg.ROW_HEADER2).value = "WBS���ږ�"
    ws.Range(cfg.COL_TASK_COUNT_LABEL & cfg.ROW_HEADER2).value = "���v"
    ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & cfg.ROW_HEADER2).value = "����"
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_HEADER2).value = "�X�e�[�^�X"
    ws.Range(cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_HEADER2).value = "�i����"
    ws.Range(cfg.COL_TASK_PROG_LABEL & cfg.ROW_HEADER2).value = "������"
    ws.Range(cfg.COL_TASK_WGT_LABEL & cfg.ROW_HEADER2).value = "���d"
    ws.Range(cfg.COL_TEAM_SLCT_LABEL & cfg.ROW_HEADER2).value = "�g�D"
    ws.Range(cfg.COL_PERSON_SLCT_LABEL & cfg.ROW_HEADER2).value = "�S��"
    ws.Range(cfg.COL_OUTPUT_LABEL & cfg.ROW_HEADER2).value = "���ʕ�"
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER2).value = "�H��(�l��)"
    ws.Range(cfg.COL_PLANNED_START_LABEL & cfg.ROW_HEADER2).value = "�J�n��"
    ws.Range(cfg.COL_PLANNED_END_LABEL & cfg.ROW_HEADER2).value = "�I����"
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER2).value = "�c�H��(�l��)"
    ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & cfg.ROW_HEADER2).value = "�ύH��(�l��)"
    ws.Range(cfg.COL_ACTUAL_START_LABEL & cfg.ROW_HEADER2).value = "�J�n��"
    ws.Range(cfg.COL_ACTUAL_END_LABEL & cfg.ROW_HEADER2).value = "�I����"
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_HEADER2).value = "�J�e�S��1"
    ws.Range(cfg.COL_CATEGORY2_LABEL & cfg.ROW_HEADER2).value = "�J�e�S��2"
    ws.Range(cfg.COL_LAST_LABEL & cfg.ROW_HEADER2).value = "���l"
    
    ' �w�b�_�[�Q������̑���
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
    
    ' �f�[�^�J�n�L�[���������
    ws.Range(cfg.COL_KEY_LABEL & lngDataStartRow).value = "@"
    
    ' �f�[�^�I���L�[���������
    ws.Range(cfg.COL_KEY_LABEL & lngDataEndRow).value = "$"
    
    ' �f�[�^�I���s���������
    ws.Range(cfg.COL_TASK_COUNT_LABEL & lngDataEndRow).value = "���v"
    ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngDataEndRow).value = "���v"
    ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngDataEndRow).value = "�S��%"
    ws.Range(cfg.COL_TASK_PROG_LABEL & lngDataEndRow).value = "�S��%"
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngDataEndRow).value = "���v�l��"
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngDataEndRow).value = "���v�l��"
    ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngDataEndRow).value = "���v�l��"
    
    ' �f�[�^�I���s������̑���
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
    
    ' ��\���J�����ݒ�
    ws.Range(cfg.COL_WBS_IDX_LABEL & 1 & ":" & cfg.COL_FLG_CE_LABEL & 1).EntireColumn.Hidden = True
    
    ' �E�B���h�E�g�̌Œ�
    Set win = ThisWorkbook.Windows(1)
    win.FreezePanes = False
    ws.Activate
    ws.Cells(lngDataStartRow + 1, cfg.COL_TEAM_SLCT).Select
    win.FreezePanes = True
    ws.Cells(1, 1).Select

End Sub


' �� �V�[�g������ - �����t����
Public Sub ResetConditionalFormatting(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' �ꎞ�ϐ���`
    Dim tmpFc As FormatCondition
    Dim tmpDataBar As Databar
    
    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �V�[�g�S�̂̏����t�������폜
    ws.Cells.FormatConditions.Delete
    
    ' �� �J�n�E�I���s�̑���
    ' �J�n�s�̔w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_DATA_START & ":" & cfg.COL_LAST_LABEL & cfg.ROW_DATA_START)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Interior.Color = RGB(0, 0, 0)
            .StopIfTrue = False
        End With
    End With
    ' �I���s�̔w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & (lngEndRow + 1) & ":" & cfg.COL_LAST_LABEL & (lngEndRow + 1))
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Interior.Color = RGB(0, 0, 0)
            .Font.Bold = True
            .StopIfTrue = False
        End With
    End With
    
    ' �� �w�b�_�[�s�̑���
    ' �w�b�_�[�P�̔w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER1)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Interior.Color = RGB(255, 230, 153)
            .Font.Color = RGB(128, 128, 132)
            .StopIfTrue = False
        End With
    End With
    ' �w�b�_�[�Q�̔w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER2)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=True")
        With tmpFc
            .Interior.Color = RGB(255, 230, 153)
            .Font.Bold = True
            .StopIfTrue = False
        End With
    End With
    ' �w�b�_�[�s�̊i�q��ݒ�
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
    
    ' �� �l�ɂ�鑕��
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
    
    ' �� �\���`���̑���
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
    
    ' �� �f�[�^�s�̃G���[����
    With ws.Range(cfg.COL_ERR_LABEL & lngStartRow & ":" & cfg.COL_ERR_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_ERR_LABEL & lngStartRow & "=""E""")
        With tmpFc
            .Font.Color = RGB(255, 0, 0)
            .Font.Bold = True
            .Interior.Color = RGB(255, 204, 204)
            .StopIfTrue = False
        End With
    End With
        
    ' �� �K���\������f�[�^�s�̌x������
    ' �e�L�X�g�J�����ɓ��͂���Ă��Ȃ��x��
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

    ' �� �f�[�^�s�̌x������
    ' �^�X�N�s�œ��͂��K�v�ȃJ�����ɓ��͂���Ă��Ȃ��x��
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

    ' �� �f�[�^�s�̒ʏ푕��
    ' �f�[�^�s�̃X�e�[�^�X�w�i�F��ݒ�
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
    ' �f�[�^�s�̃��x���J�����̑���
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
    ' �f�[�^�s�̃^�X�N�K�w�ȊO����������ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=False")
        With tmpFc
            .Font.Bold = True
        .StopIfTrue = False
        End With
    End With
    ' �f�[�^�s�̃^�X�N�K�w�w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE")
        With tmpFc
            .Interior.Color = RGB(255, 255, 255)
        .StopIfTrue = False
        End With
    End With
    ' �f�[�^�s��L1�K�w�w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=1,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(48, 84, 150)
            .Font.Color = RGB(255, 255, 255)
        .StopIfTrue = False
        End With
    End With
    ' �f�[�^�s��L2�K�w�w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=2,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(180, 198, 231)
        .StopIfTrue = False
        End With
    End With
    ' �f�[�^�s��L3�K�w�w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=3,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(217, 225, 242)
        .StopIfTrue = False
        End With
    End With
    ' �f�[�^�s��L4�K�w�w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=4,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(236, 240, 248)
        .StopIfTrue = False
        End With
    End With
    ' �f�[�^�s��L5�K�w�w�i�F��ݒ�
    With ws.Range(cfg.COL_CHK_LABEL & lngStartRow & ":" & cfg.COL_LAST_LABEL & lngEndRow)
        Set tmpFc = .FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($" & cfg.COL_LEVEL_LABEL & lngStartRow & "=5,$" & cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)")
        With tmpFc
            .Interior.Color = RGB(245, 247, 251)
        .StopIfTrue = False
        End With
    End With
    ' �f�[�^�s�̊i�q��ݒ�
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


' �� �V�[�g������ - �f�[�^���͋K��
Public Sub ResetDataValidation(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' �ꎞ�ϐ���`
    Dim tmpRange As Range
    Dim tmpRuleList As String
       
    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �� WBS�X�e�[�^�X��̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_WBS_STATUS_LABEL & lngStartRow & ":" & cfg.COL_WBS_STATUS_LABEL & lngEndRow)
    ' ���[�����擾
    tmpRuleList = "-," & cfg.WBS_STATUS_LIST
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' �� �g�D��̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_TEAM_SLCT_LABEL & lngStartRow & ":" & cfg.COL_TEAM_SLCT_LABEL & lngEndRow)
    ' ���[�����擾
    tmpRuleList = CreateValidationListString(ws, ws.Range(cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_CTRL1), tmpRange)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-," & tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With

    ' �� �S����̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_PERSON_SLCT_LABEL & lngStartRow & ":" & cfg.COL_PERSON_SLCT_LABEL & lngEndRow)
    ' ���[�����擾
    tmpRuleList = CreateValidationListString(ws, ws.Range(cfg.COL_EFFORT_PROG_LABEL & cfg.ROW_CTRL2), tmpRange)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-," & tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' �� �J�e�S��1��̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_CATEGORY1_LABEL & lngStartRow & ":" & cfg.COL_CATEGORY1_LABEL & lngEndRow)
    ' ���[�����擾
    tmpRuleList = CreateValidationListString(ws, ws.Range(cfg.COL_CATEGORY2_LABEL & cfg.ROW_CTRL1), tmpRange)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-," & tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' �� �J�e�S��2��̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_CATEGORY2_LABEL & lngStartRow & ":" & cfg.COL_CATEGORY2_LABEL & lngEndRow)
    ' ���[�����擾
    tmpRuleList = CreateValidationListString(ws, ws.Range(cfg.COL_CATEGORY2_LABEL & cfg.ROW_CTRL2), tmpRange)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-," & tmpRuleList
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
    
    ' �� �K�w�ԍ���̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_TASK_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateWholeNumber, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="1", Formula2:="999"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "1�`999 �̐���"
        .ErrorTitle = "���̓G���["
        .InputMessage = "1�`999 �̐�������͂��Ă��������i�󔒂��j"
        .ErrorMessage = "1�`999 �̐����̂ݓ��͉\�ł��B"
        .ShowInput = True
        .ShowError = True
    End With

    ' �� ���x����̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_LEVEL_LABEL & lngStartRow & ":" & cfg.COL_LEVEL_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateWholeNumber, _
             AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, _
             Formula1:="0", Formula2:="5"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ErrorTitle = "���̓G���["
        .ErrorMessage = "0�`5 �̐����̂ݓ��͉\�ł��B"
        .ShowInput = False
        .ShowError = True
    End With

    ' �� �e��t���O��̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_CE_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_FLG_T_LABEL & lngStartRow & ")," & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=TRUE," & _
                        cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE)"
        .IgnoreBlank = True
        .InCellDropdown = False
        .ErrorTitle = "���̓G���["
        .ErrorMessage = "TRUE �܂��� FALSE �̂ݓ��͉\�ł��B"
        .ShowInput = False
        .ShowError = True
    End With

    ' �� �^�X�N���v��̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ")," & _
                        cfg.COL_TASK_COUNT_LABEL & lngStartRow & ">=0,INT(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ")=" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & "))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .ErrorTitle = "���̓G���["
        .ErrorMessage = "0�ȏ�̐����̂ݓ��͉\�ł��B"
        .ShowInput = False
        .ShowError = True
    End With

    ' �� �^�X�N������̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ")," & _
                        cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ">=0,INT(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ")=" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & "))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .ErrorTitle = "���̓G���["
        .ErrorMessage = "0�ȏ�̐����̂ݓ��͉\�ł��B"
        .ShowInput = False
        .ShowError = True
    End With

    ' �� �H���i������̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ")," & _
                        cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ">=0," & cfg.COL_EFFORT_PROG_LABEL & lngStartRow & "<=1))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0�`100% �̒l�̂�"
        .ErrorTitle = "���̓G���["
        .InputMessage = "0�`100%�i= 0�`1�j�̒l����͂��Ă��������i�󔒉j"
        .ErrorMessage = "0�`100%�̊Ԃ̐��l�̂ݓ��͉\�ł��B"
        .ShowInput = True
        .ShowError = True
    End With

    ' �� ���ڏ�������̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_TASK_PROG_LABEL & lngStartRow & ")," & _
                        cfg.COL_TASK_PROG_LABEL & lngStartRow & ">=0," & cfg.COL_TASK_PROG_LABEL & lngStartRow & "<=1))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0�`100% �̒l�̂�"
        .ErrorTitle = "���̓G���["
        .InputMessage = "0�`100%�i= 0�`1�j�̒l����͂��Ă��������i�󔒉j"
        .ErrorMessage = "0�`100%�̊Ԃ̐��l�̂ݓ��͉\�ł��B"
        .ShowInput = True
        .ShowError = True
    End With

    ' �� ���ډ��d��̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_TASK_WGT_LABEL & lngStartRow & ":" & cfg.COL_TASK_WGT_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ")," & _
             cfg.COL_TASK_WGT_LABEL & lngStartRow & ">=1,INT(" & cfg.COL_TASK_WGT_LABEL & lngStartRow & ")=" & cfg.COL_TASK_WGT_LABEL & lngStartRow & "))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "1�ȏ�̐����̂�"
        .ErrorTitle = "���̓G���["
        .InputMessage = "1�ȏ�̐�������͂��Ă��������i�󔒉j"
        .ErrorMessage = "1�ȏ�̐����̂ݓ��͉\�ł��B"
        .ShowInput = True
        .ShowError = True
    End With

    ' �� �\��H����̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ")," & _
             cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ">=0))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0�ȏ�̐��l"
        .ErrorTitle = "���̓G���["
        .InputMessage = "0�ȏ�̐��l����͂��Ă��������i�󔒉j"
        .ErrorMessage = "0�ȏ�̐��l�̂ݓ��͂ł��܂��B"
        .ShowInput = True
        .ShowError = True
    End With

    ' �� ���юc�H����̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ")," & _
             cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ">=0))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0�ȏ�̐��l"
        .ErrorTitle = "���̓G���["
        .InputMessage = "0�ȏ�̐��l����͂��Ă��������i�󔒉j"
        .ErrorMessage = "0�ȏ�̐��l�̂ݓ��͂ł��܂��B"
        .ShowInput = True
        .ShowError = True
    End With

    ' �� ���эύH����̓��͋K�����X�V
    ' �f�[�^�͈̔͂��w��
    Set tmpRange = ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow)
    With tmpRange.Validation
        ' ���[�����폜
        .Delete
        ' ���[����ݒ�
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=OR(ISBLANK(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & "),AND(ISNUMBER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ")," & _
             cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ">=0))"
        .IgnoreBlank = True
        .InCellDropdown = False
        .InputTitle = "0�ȏ�̐��l"
        .ErrorTitle = "���̓G���["
        .InputMessage = "0�ȏ�̐��l����͂��Ă��������i�󔒉j"
        .ErrorMessage = "0�ȏ�̐��l�̂ݓ��͂ł��܂��B"
        .ShowInput = True
        .ShowError = True
    End With
    
    ' �� �p�[�Z���e�[�W�̓��͂̂��߂̃Z�������ݒ�
    ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow).NumberFormat = "0.0%"

End Sub


' �� �f�[�^���͋K����������쐬
Private Function CreateValidationListString(ws As Worksheet, defineRange As Range, dataRange As Range) As String

    ' �ϐ���`
    Dim colUniqueList As Collection
    Dim strDefine As String
    ' �ꎞ�ϐ���`
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

    ' Collection �I�u�W�F�N�g�𐶐�
    Set colUniqueList = New Collection

    ' defineRange �̕����������
    strDefine = Trim(defineRange.value)
    If Len(strDefine) > 0 Then
        tmpDefineArray = Split(strDefine, ",")
        For i = LBound(tmpDefineArray) To UBound(tmpDefineArray)
            tmpTrimmedValue = Trim(tmpDefineArray(i))
            If Len(tmpTrimmedValue) > 0 Then
                ' Collection �̃L�[�ɒl��ǉ��i�d���̓G���[�ɂȂ邽�� On Error Resume Next �Ŗ����j
                On Error Resume Next
                colUniqueList.Add tmpTrimmedValue, tmpTrimmedValue
                On Error GoTo 0 ' �G���[������ʏ�ɖ߂�
            End If
        Next i
    End If

    ' dataRange �̃Z���l��z��Ƃ��Ĉꊇ�擾�E����
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
        ' �P��Z���i�z��łȂ��j�Ή�
        tmpVal = Trim(CStr(tmpDataArray))
        If Len(tmpVal) > 0 Then
            On Error Resume Next
            colUniqueList.Add tmpVal, tmpVal
            On Error GoTo 0
        End If
    End If

    ' Collection �̃A�C�e�����J���}��؂�̕�����Ƃ��đg�ݗ��Ă�
    tmpDelimiter = ""
    For Each tmpItem In colUniqueList
        CreateValidationListString = CreateValidationListString & tmpDelimiter & tmpItem
        tmpDelimiter = ","
    Next tmpItem

End Function


' �� �V�[�g������ - �Z���z�u
Public Sub ResetHorizontalAlignment(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    
    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �V�[�g�S�̂̃Z���z�u�����Z�b�g
    With ws.Cells
        .HorizontalAlignment = xlGeneral
    End With
    
    ' �C���f���g���Đݒ�
    ws.Range("B2").IndentLevel = 1
    
    ' �^�C�g���s�̔z�u
    ws.Range(cfg.COL_LAST_LABEL & cfg.ROW_TITLE).HorizontalAlignment = xlRight
    
    ' �R���g���[���s���͕�����̔z�u
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_WBS_STATUS_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL1).HorizontalAlignment = xlRight
    ws.Range(cfg.COL_CATEGORY1_LABEL & cfg.ROW_CTRL2).HorizontalAlignment = xlRight
    
    ' �w�b�_�[�P�̔z�u
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenter
    ws.Range(cfg.COL_TASK_TEXT_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_TEXT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlLeft
    ws.Range(cfg.COL_TASK_COUNT_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_TASK_PROG_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_TASK_WGT_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_PLANNED_END_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & cfg.ROW_HEADER1 & ":" & cfg.COL_ACTUAL_END_LABEL & cfg.ROW_HEADER1).HorizontalAlignment = xlCenterAcrossSelection
    
    ' �w�b�_�[�Q�̔z�u
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_LAST_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenter
    ws.Range(cfg.COL_CHK_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_OPT_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_L1_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_TASK_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenterAcrossSelection
    ws.Range(cfg.COL_WBS_ID_LABEL & cfg.ROW_HEADER2 & ":" & cfg.COL_TEXT_LABEL & cfg.ROW_HEADER2).HorizontalAlignment = xlCenterAcrossSelection
    
    ' �f�[�^�I���s�Ƃ��̎��̍s�̔z�u
    ws.Range(cfg.COL_CHK_LABEL & (lngEndRow + 1) & ":" & cfg.COL_LAST_LABEL & (lngEndRow + 2)).HorizontalAlignment = xlCenter
    
    ' �擪�`���x���E�^�X�N��܂ł̔z�u
    ws.Range(cfg.COL_ERR_LABEL & lngStartRow & ":" & cfg.COL_TASK_LABEL & lngEndRow).HorizontalAlignment = xlCenter
    
    ' �^�X�N�W�v���v�`�J�e�S��2�̔z�u
    ws.Range(cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_CATEGORY2_LABEL & lngEndRow).HorizontalAlignment = xlCenter

    ' ���ʕ��̔z�u
    ws.Range(cfg.COL_OUTPUT_LABEL & lngStartRow & ":" & cfg.COL_OUTPUT_LABEL & lngEndRow).HorizontalAlignment = xlLeft

End Sub


' �� �V�[�g������ - �t�H�[��
Public Sub ResetExecuteForm(ws As Worksheet, Optional blnShouldClearOptMemory As Boolean = False)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim lngRowCount As Long
    Dim varChkArray() As Variant
    Dim varOptArray() As Variant
    ' - ��ʔz�u�R���g���[��
    Dim shpExe1ComboBox As Shape
    Dim shpExe1Button As Shape
    Dim shpReset1Button As Shape
    Dim shpExe2ComboBox As Shape
    Dim shpExe2Button As Shape
    Dim shpReset2Button As Shape
    ' - ���s1�R���{�{�b�N�X�̈ʒu�ƃT�C�Y���v�Z���邽�߂̕ϐ�
    Dim dblExe1ComboBoxLeft As Double
    Dim dblExe1ComboBoxTop As Double
    Dim dblExe1ComboBoxWidth As Double
    Dim dblExe1ComboBoxHeight As Double
    ' - ���s1�{�^���̈ʒu�ƃT�C�Y���v�Z���邽�߂̕ϐ�
    Dim dblExe1ButtonLeft As Double
    Dim dblExe1ButtonTop As Double
    Dim dblExe1ButtonWidth As Double
    Dim dblExe1ButtonHeight As Double
    ' - ���Z�b�g1�{�^���̈ʒu�ƃT�C�Y���v�Z���邽�߂̕ϐ�
    Dim dblReset1ButtonLeft As Double
    Dim dblReset1ButtonTop As Double
    Dim dblReset1ButtonWidth As Double
    Dim dblReset1ButtonHeight As Double
    ' - ���s2�R���{�{�b�N�X�̈ʒu�ƃT�C�Y���v�Z���邽�߂̕ϐ�
    Dim dblExe2ComboBoxLeft As Double
    Dim dblExe2ComboBoxTop As Double
    Dim dblExe2ComboBoxWidth As Double
    Dim dblExe2ComboBoxHeight As Double
    ' - ���s2�{�^���̈ʒu�ƃT�C�Y���v�Z���邽�߂̕ϐ�
    Dim dblExe2ButtonLeft As Double
    Dim dblExe2ButtonTop As Double
    Dim dblExe2ButtonWidth As Double
    Dim dblExe2ButtonHeight As Double
    ' �ꎞ�ϐ���`
    Dim r As Long
    Dim tmpVar As Variant

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' ���s1�R���{�{�b�N�X�̈ʒu�ƃT�C�Y���v�Z
    dblExe1ComboBoxLeft = ws.Cells(cfg.ROW_CTRL1, cfg.COL_L1).Left
    dblExe1ComboBoxTop = ws.Cells(cfg.ROW_CTRL1, cfg.COL_L1).Top
    dblExe1ComboBoxWidth = cfg.WIDTH_EXE1_COMBOBOX
    dblExe1ComboBoxHeight = ws.Cells(cfg.ROW_CTRL1, cfg.COL_L1).Height
    ' �������O�̎��s1�R���{�{�b�N�X�����݂��邩�m�F
    On Error Resume Next
    Set shpExe1ComboBox = ws.Shapes(cfg.NAME_EXE1_COMBOBOX)
    On Error GoTo 0
    ' ���s1�R���{�{�b�N�X�����݂���ꍇ�A�폜
    If Not shpExe1ComboBox Is Nothing Then
        shpExe1ComboBox.Delete
    End If
    ' ���s1�R���{�{�b�N�X��V���ɍ쐬�i�T�C�Y���ς��\�������邽�߁A����A��蒼���j
    Set shpExe1ComboBox = ws.Shapes.AddFormControl(xlDropDown, dblExe1ComboBoxLeft, dblExe1ComboBoxTop, dblExe1ComboBoxWidth, dblExe1ComboBoxHeight)
    shpExe1ComboBox.Name = cfg.NAME_EXE1_COMBOBOX
    With shpExe1ComboBox.ControlFormat
        .AddItem "���ׂčČv�Z"
        .AddItem "�I�[�g�t�B���^�[�����Z�b�g"
        .AddItem "�K�w�ԍ��őS�̂��\�[�g"
        .AddItem "�����E���͋K�������Z�b�g"
        .AddItem "���̓t�H�[�������Z�b�g"
        .AddItem "�G���[�`�F�b�N"
    End With
    With ws.DropDowns(cfg.NAME_EXE1_COMBOBOX)
        .ListIndex = 1
    End With
    
    ' ���s1�{�^���̈ʒu�ƃT�C�Y���v�Z
    dblExe1ButtonLeft = dblExe1ComboBoxLeft + dblExe1ComboBoxWidth
    dblExe1ButtonTop = dblExe1ComboBoxTop
    dblExe1ButtonWidth = cfg.WIDTH_EXE1_BUTTON
    dblExe1ButtonHeight = dblExe1ComboBoxHeight
    ' �������O�̎��s1�{�^�������݂��邩�m�F
    On Error Resume Next
    Set shpExe1Button = ws.Shapes(cfg.NAME_EXE1_BUTTON)
    On Error GoTo 0
    ' ���s1�{�^�������݂���ꍇ�A�폜
    If Not shpExe1Button Is Nothing Then
        shpExe1Button.Delete
    End If
    ' ���s1�{�^����V���ɍ쐬�i�T�C�Y���ς��\�������邽�߁A����A��蒼���j
    Set shpExe1Button = ws.Shapes.AddFormControl(xlButtonControl, dblExe1ButtonLeft, dblExe1ButtonTop, dblExe1ButtonWidth, dblExe1ButtonHeight)
    shpExe1Button.Name = cfg.NAME_EXE1_BUTTON
    With ws.Buttons(cfg.NAME_EXE1_BUTTON)
        .Characters.Text = "���s"
    End With
    shpExe1Button.OnAction = "wbsui.Exe1ButtonClick"
    
     ' ���Z�b�g1�{�^���̈ʒu�ƃT�C�Y���v�Z
    dblReset1ButtonLeft = dblExe1ComboBoxLeft + dblExe1ComboBoxWidth + dblExe1ButtonWidth
    dblReset1ButtonTop = dblExe1ComboBoxTop
    dblReset1ButtonWidth = cfg.WIDTH_RESET1_BUTTON
    dblReset1ButtonHeight = dblExe1ComboBoxHeight
    ' �������O�̃��Z�b�g1�{�^�������݂��邩�m�F
    On Error Resume Next
    Set shpReset1Button = ws.Shapes(cfg.NAME_RESET1_BUTTON)
    On Error GoTo 0
    ' ���Z�b�g1�{�^�������݂���ꍇ�A�폜
    If Not shpReset1Button Is Nothing Then
        shpReset1Button.Delete
    End If
    ' ���Z�b�g1�{�^����V���ɍ쐬�i�T�C�Y���ς��\�������邽�߁A����A��蒼���j
    Set shpReset1Button = ws.Shapes.AddFormControl(xlButtonControl, dblReset1ButtonLeft, dblReset1ButtonTop, dblReset1ButtonWidth, dblReset1ButtonHeight)
    shpReset1Button.Name = cfg.NAME_RESET1_BUTTON
    With ws.Buttons(cfg.NAME_RESET1_BUTTON)
        .Characters.Text = "���Z�b�g"
    End With
    shpReset1Button.OnAction = "wbsui.Reset1ButtonClick"
    
    ' ���s2�R���{�{�b�N�X�̈ʒu�ƃT�C�Y���v�Z
    dblExe2ComboBoxLeft = ws.Cells(cfg.ROW_CTRL2, cfg.COL_L1).Left
    dblExe2ComboBoxTop = ws.Cells(cfg.ROW_CTRL2, cfg.COL_L1).Top
    dblExe2ComboBoxWidth = cfg.WIDTH_EXE2_COMBOBOX
    dblExe2ComboBoxHeight = ws.Cells(cfg.ROW_CTRL2, cfg.COL_L1).Height
    ' �������O�̎��s2�R���{�{�b�N�X�����݂��邩�m�F
    On Error Resume Next
    Set shpExe2ComboBox = ws.Shapes(cfg.NAME_EXE2_COMBOBOX)
    On Error GoTo 0
    ' ���s2�R���{�{�b�N�X�����݂���ꍇ�A�폜
    If Not shpExe2ComboBox Is Nothing Then
        shpExe2ComboBox.Delete
    End If
    ' ���s2�R���{�{�b�N�X��V���ɍ쐬�i�T�C�Y���ς��\�������邽�߁A����A��蒼���j
    Set shpExe2ComboBox = ws.Shapes.AddFormControl(xlDropDown, dblExe2ComboBoxLeft, dblExe2ComboBoxTop, dblExe2ComboBoxWidth, dblExe2ComboBoxHeight)
    shpExe2ComboBox.Name = cfg.NAME_EXE2_COMBOBOX
    With shpExe2ComboBox.ControlFormat
        .AddItem "�yOPT�z �I�������s�̉��Ɉ�s�ǉ�"
        .AddItem "�yOPT�z �I�������K�w�ԍ��̖������{�P"
        .AddItem "�yOPT�z �I�������K�w�ԍ��̖������|�P"
        .AddItem "�yCHK�z �`�F�b�N�����Q�ӏ��̊K�w�ԍ��̖����ԍ��������@�� �`�F�b�N�����K�w�ł���K�w�s�̂Q�ӏ��łȂ������ꍇ�͕s�� ��"
        .AddItem "�yCHK�z �`�F�b�N�����s���폜�@�� �q�K�w��q�^�X�N������ꍇ�͕s�� ��"
    End With
    With ws.DropDowns(cfg.NAME_EXE2_COMBOBOX)
        .ListIndex = 1
    End With
    
    ' ���s2�{�^���̈ʒu�ƃT�C�Y���v�Z
    dblExe2ButtonLeft = dblExe2ComboBoxLeft + dblExe2ComboBoxWidth
    dblExe2ButtonTop = dblExe2ComboBoxTop
    dblExe2ButtonWidth = cfg.WIDTH_EXE2_BUTTON
    dblExe2ButtonHeight = dblExe2ComboBoxHeight
    ' �������O�̎��s2�{�^�������݂��邩�m�F
    On Error Resume Next
    Set shpExe2Button = ws.Shapes(cfg.NAME_EXE2_BUTTON)
    On Error GoTo 0
    ' ���s2�{�^�������݂���ꍇ�A�폜
    If Not shpExe2Button Is Nothing Then
        shpExe2Button.Delete
    End If
    ' ���s2�{�^����V���ɍ쐬�i�T�C�Y���ς��\�������邽�߁A����A��蒼���j
    Set shpExe2Button = ws.Shapes.AddFormControl(xlButtonControl, dblExe2ButtonLeft, dblExe2ButtonTop, dblExe2ButtonWidth, dblExe2ButtonHeight)
    shpExe2Button.Name = cfg.NAME_EXE2_BUTTON
    With ws.Buttons(cfg.NAME_EXE2_BUTTON)
        .Characters.Text = "���s"
    End With
    shpExe2Button.OnAction = "wbsui.Exe2ButtonClick"
    
    ' ���Z�b�g2�{�^���̈ʒu�ƃT�C�Y���v�Z
    dblReset2ButtonLeft = dblExe2ComboBoxLeft + dblExe2ComboBoxWidth + dblExe2ButtonWidth
    dblReset2ButtonTop = dblExe2ComboBoxTop
    dblReset2ButtonWidth = cfg.WIDTH_RESET2_BUTTON
    dblReset2ButtonHeight = dblExe2ComboBoxHeight
    ' �������O�̃��Z�b�g2�{�^�������݂��邩�m�F
    On Error Resume Next
    Set shpReset2Button = ws.Shapes(cfg.NAME_RESET2_BUTTON)
    On Error GoTo 0
    ' ���Z�b�g2�{�^�������݂���ꍇ�A�폜
    If Not shpReset2Button Is Nothing Then
        shpReset2Button.Delete
    End If
    ' ���Z�b�g2�{�^����V���ɍ쐬
    Set shpReset2Button = ws.Shapes.AddFormControl(xlButtonControl, dblReset2ButtonLeft, dblReset2ButtonTop, dblReset2ButtonWidth, dblReset2ButtonHeight)
    shpReset2Button.Name = cfg.NAME_RESET2_BUTTON
    With ws.Buttons(cfg.NAME_RESET2_BUTTON)
        .Characters.Text = "���Z�b�g"
    End With
    shpReset2Button.OnAction = "wbsui.Reset2ButtonClick"
    
     ' �Ώۍs�����擾
    lngRowCount = lngEndRow - lngStartRow + 1
    
    ' �ꊇ�������݂̂��߂̃f�[�^��p��
    ReDim varChkArray(1 To lngRowCount, 1 To 1)
    ReDim varOptArray(1 To lngRowCount, 1 To 1)
    
    ' �l��z��Ɋi�[�i�Œ�l�j
    For r = 1 To lngRowCount
        varChkArray(r, 1) = cfg.CHK_MARK_F
        varOptArray(r, 1) = cfg.OPT_MARK_F
    Next r
    
    ' ���ʂ���������
    ws.Range(ws.Cells(lngStartRow, cfg.COL_CHK), ws.Cells(lngEndRow, cfg.COL_CHK)).value = varChkArray
    ws.Range(ws.Cells(lngStartRow, cfg.COL_OPT), ws.Cells(lngEndRow, cfg.COL_OPT)).value = varOptArray
    
    ' �����̎w���OPT�̃������Z�����N���A����K�v������ꍇ
    If blnShouldClearOptMemory = True Then
        ws.Range(cfg.COL_OPT_LABEL & cfg.ROW_DATA_START).ClearContents
    End If
    
    ' �Ō�ɑI������OPT�𔽉f
    tmpVar = ws.Cells(cfg.ROW_DATA_START, cfg.COL_OPT).value
    If tmpVar <> "" And _
            IsNumeric(tmpVar) And _
            tmpVar >= lngStartRow And _
            tmpVar <= lngEndRow Then
        ws.Cells(tmpVar, cfg.COL_OPT).value = cfg.OPT_MARK_T
        
    End If

End Sub


' �� �V�[�g������ - �^�C�g���s
Public Sub ResetTitleRow(ws As Worksheet)

    ' �ϐ���`
    Dim rngTargetCell As Range
    Dim strGitHubURL As String
    Dim strCommentText As String

    ' �^�C�g���s�̃f�[�^����U�폜
    ws.Rows(cfg.ROW_TITLE).ClearContents

    ' �V�[�g�����Z�b�g
    ws.Range(cfg.COL_ERR_LABEL & cfg.ROW_TITLE).value = ws.Name
    
    ' �Ώۂ̃Z����ݒ�
    Set rngTargetCell = ws.Range(cfg.COL_LAST_LABEL & cfg.ROW_TITLE)

    ' GitHub �� URL
    strGitHubURL = "https://github.com/H16K148/wbs-template-xlsm"

    ' �Z���ɕ���������
    rngTargetCell.value = strGitHubURL

    ' �n�C�p�[�����N��ݒ�
    ws.Hyperlinks.Add Anchor:=rngTargetCell, Address:=strGitHubURL, TextToDisplay:=strGitHubURL

    ' �����T�C�Y�� 8 �ɐݒ�
    rngTargetCell.Font.Size = 8

    ' �z�u���E���񂹂ɐݒ�
    rngTargetCell.HorizontalAlignment = xlRight
    rngTargetCell.VerticalAlignment = xlBottom
    
    ' �R�����g�̓��e
    strCommentText = "�o�[�W�����Fv" & cfg.APP_VERSION & vbCrLf & "�Q�l���F" & vbCrLf & "�V�[�g���b�N�̉����Ɋւ���d�v�ȏ��́A�����N��̃h�L�������g�ɋL�ڂ���Ă��܂��B"
    
    ' �Z���ɃR�����g��ǉ��܂��͕ҏW
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


' �� ��{�����̃��Z�b�g
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


' �� �W�v�����̃��Z�b�g
Public Sub ResetAggregateFormulas(ws As Worksheet)

    wbslib.SetFormulaForPlannedEffort ws
    wbslib.SetFormulaForActualCompletedEffort ws
    wbslib.SetFormulaForActualRemainingEffort ws
    wbslib.SetFormulaForTaskProgressRate ws
    wbslib.SetFormulaForEffortProgressRate ws
    wbslib.SetFormulaForTaskCount ws
    wbslib.SetFormulaForTaskCompCount ws

End Sub


' �� �I�[�g�t�B���^�[�̃��Z�b�g
Public Sub ResetAutoFilter(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    
    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �����������
    ws.AutoFilterMode = False
    
    ' �ݒ�
    ws.Range(cfg.COL_L1_LABEL & cfg.ROW_DATA_START & ":" & cfg.COL_CATEGORY2_LABEL & lngEndRow).AutoFilter

End Sub


' �� �����l�̃Z�b�g
Public Sub SetInitialValue(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' �ꎞ�ϐ���`
    Dim tmpRngTarget As Range
    Dim tmpVarTarget As Variant
    
    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' WBS�X�e�[�^�X�s
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_WBS_STATUS), ws.Cells(lngEndRow, cfg.COL_WBS_STATUS))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget
    
    ' ���ډ��d�s
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_WGT), ws.Cells(lngEndRow, cfg.COL_TASK_WGT))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = 1
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

    ' �g�D�s
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_TEAM_SLCT), ws.Cells(lngEndRow, cfg.COL_TEAM_SLCT))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

    ' �S���s
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_PERSON_SLCT), ws.Cells(lngEndRow, cfg.COL_PERSON_SLCT))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

    ' �J�e�S��1�s
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_CATEGORY1), ws.Cells(lngEndRow, cfg.COL_CATEGORY1))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

    ' �J�e�S��2�s
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_CATEGORY2), ws.Cells(lngEndRow, cfg.COL_CATEGORY2))
    tmpVarTarget = tmpRngTarget.value
    For i = LBound(tmpVarTarget, 1) To UBound(tmpVarTarget, 1)
      If IsEmpty(tmpVarTarget(i, 1)) Then
        tmpVarTarget(i, 1) = "-"
      End If
    Next i
    tmpRngTarget.value = tmpVarTarget

End Sub


' �� Exe1�{�^���N���b�N���Ɏ��s����鏈��
Public Sub Exe1ButtonClick()

    ' ���݂̃V�[�g���擾
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    ' �ϐ���`
    Dim lngSelectedIndex As Long
    Dim shpExe1ComboBox As Shape
    
    ' ���s�R���{�{�b�N�X���擾
    On Error Resume Next
    Set shpExe1ComboBox = ws.Shapes(cfg.NAME_EXE1_COMBOBOX)
    On Error GoTo 0
    
    ' ���s�R���{�{�b�N�X�����݂��Ȃ��ꍇ�A�I��
    If shpExe1ComboBox Is Nothing Then
        Exit Sub
    End If
    
    ' �I�𒆂̃C���f�b�N�X���擾
    lngSelectedIndex = shpExe1ComboBox.ControlFormat.ListIndex

    ' �C���f�b�N�X�ɑΉ����鏈�������s
    Select Case lngSelectedIndex
        Case 1
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' �^�C�g���A��{�����̃��Z�b�g
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �W�v�����̃��Z�b�g
            wbsui.ResetAggregateFormulas ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 2
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' �^�C�g���A��{�����̃��Z�b�g
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �I�[�g�t�B���^�[�����Z�b�g
            wbsui.ResetAutoFilter ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 3
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' �^�C�g���A��{�����̃��Z�b�g
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �\�[�g�����{
            wbslib.ExecSortWbsRange ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 4
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' �^�C�g���A��{�����̃��Z�b�g
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �����E���͋K�������Z�b�g
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
            
            ' �^�C�g���A��{�����̃��Z�b�g
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' ���̓t�H�[�������Z�b�g
            wbsui.ResetExecuteForm ws, True
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 6
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' �^�C�g���A��{�����̃��Z�b�g
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �G���[�`�F�b�N
            wbslib.ExecCheckWbsErrors ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case Else
            MsgBox "���ڂ��I������Ă��܂���B"
    End Select

End Sub


' �� Reset1�{�^���N���b�N���Ɏ��s����鏈��
Public Sub Reset1ButtonClick()

    ' ���݂̃V�[�g���擾
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    ' �ϐ���`
    Dim lngSelectedIndex As Long
    Dim shpExe1ComboBox As Shape
    
    ' ���s1�R���{�{�b�N�X���擾
    On Error Resume Next
    Set shpExe1ComboBox = ws.Shapes(cfg.NAME_EXE1_COMBOBOX)
    On Error GoTo 0
    
    ' ���s1�R���{�{�b�N�X�����݂��Ȃ��ꍇ�A�I��
    If shpExe1ComboBox Is Nothing Then
        Exit Sub
    End If
    
    ' ���s1�R���{�{�b�N�X�̃��X�g�C���f�b�N�X��擪��
    With ws.DropDowns(cfg.NAME_EXE1_COMBOBOX)
        .ListIndex = 1
    End With
    
End Sub


' �� Exe2�{�^���N���b�N���Ɏ��s����鏈��
Public Sub Exe2ButtonClick()

    ' ���݂̃V�[�g���擾
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    ' �ϐ���`
    Dim lngSelectedIndex As Long
    Dim shpExe2ComboBox As Shape
    
    ' ���s�R���{�{�b�N�X���擾
    On Error Resume Next
    Set shpExe2ComboBox = ws.Shapes(cfg.NAME_EXE2_COMBOBOX)
    On Error GoTo 0
    
    ' ���s�R���{�{�b�N�X�����݂��Ȃ��ꍇ�A�I��
    If shpExe2ComboBox Is Nothing Then
        Exit Sub
    End If
    
    ' �I�𒆂̃C���f�b�N�X���擾
    lngSelectedIndex = shpExe2ComboBox.ControlFormat.ListIndex

    ' �C���f�b�N�X�ɑΉ����鏈�������s
    Select Case lngSelectedIndex
        Case 1
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' �^�C�g���A��{�����̃��Z�b�g
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �I�������s�̉��Ɉ�s�ǉ�
            wbslib.ExecInsertRowBelowSelection ws
            
            ' �����l����
            wbsui.SetInitialValue ws
            
            ' ��{�������Z�b�g
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �W�v�������Z�b�g
            wbsui.ResetAggregateFormulas ws
            
            ' �G���[�`�F�b�N
            wbslib.ExecCheckWbsErrors ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 2
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' �^�C�g���A��{�����̃��Z�b�g
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �I�������s�̖����̃C���f�b�N�X��+1
            wbslib.ExecIncrementSelectedLastLevel ws
            
            ' �����l����
            wbsui.SetInitialValue ws
            
            ' ��{�������Z�b�g
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �W�v�������Z�b�g
            wbsui.ResetAggregateFormulas ws
            
            ' �G���[�`�F�b�N
            wbslib.ExecCheckWbsErrors ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 3
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
            Application.EnableEvents = False
            
            ' �^�C�g���A��{�����̃��Z�b�g
            wbsui.ResetTitleRow ws
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �I�������s�̖����̃C���f�b�N�X��-1
            wbslib.ExecDecrementSelectedLastLevel ws
            
            ' �����l����
            wbsui.SetInitialValue ws
            
            ' ��{�������Z�b�g
            wbsui.ResetBasicFormulas ws
            
            ' ���̍X�V���A�ꎞ�I�����v�Z���s��
            If Application.Calculation = xlCalculationManual Then
                Application.Calculation = xlCalculationAutomatic
                Application.Calculation = xlCalculationManual
            End If
            
            ' �W�v�������Z�b�g
            wbsui.ResetAggregateFormulas ws
            
            ' �G���[�`�F�b�N
            wbslib.ExecCheckWbsErrors ws
            
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
        Case 4
            MsgBox "�S�ڂ�I���i" & ws.Name & "�j"
        Case 5
            MsgBox "�T�ڂ�I���i" & ws.Name & "�j"
        Case Else
            MsgBox "���ڂ��I������Ă��܂���B"
    End Select

End Sub


' �� Reset2�{�^���N���b�N���Ɏ��s����鏈��
Public Sub Reset2ButtonClick()

    ' ���݂̃V�[�g���擾
    Dim ws As Worksheet
    Set ws = Application.ActiveSheet

    ' �ϐ���`
    Dim lngSelectedIndex As Long
    Dim shpExe2ComboBox As Shape
    
    ' ���s2�R���{�{�b�N�X���擾
    On Error Resume Next
    Set shpExe2ComboBox = ws.Shapes(cfg.NAME_EXE2_COMBOBOX)
    On Error GoTo 0
    
    ' ���s2�R���{�{�b�N�X�����݂��Ȃ��ꍇ�A�I��
    If shpExe2ComboBox Is Nothing Then
        Exit Sub
    End If
    
    ' ���s2�R���{�{�b�N�X�̃��X�g�C���f�b�N�X��擪��
    With ws.DropDowns(cfg.NAME_EXE2_COMBOBOX)
        .ListIndex = 1
    End With
    
End Sub


' �� �C�x���g�R�[�h��}������
' �@ ���̐ݒ肪�K�v��
' �@ �E�c�[�� > �Q�Ɛݒ� > Microsoft Visual Basic for Applications Extensibillity 5.3 �Ƀ`�F�b�N�iVBIDE �ւ̎Q�ƒǉ��j
' �@ �E�Z�L�����e�B�Z���^�[ > �}�N���̐ݒ� > �uVBA �v���W�F�N�g �I�u�W�F�N�g ���f���ւ̃A�N�Z�X��M������v
Private Sub InitDoubleClickHandlerToSheet(ws As Worksheet)

    ' �ϐ���`
    Dim vbComp As VBIDE.VBComponent
    Dim codeLines As New Collection

    ' �Ώۂ̃V�[�g���W���[�����擾
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
    codeLines.Add "' �� CHK �� OPT �̃_�u���N���b�N�C�x���g������"
    codeLines.Add "Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)"
    codeLines.Add "    ' �ϐ���`"
    codeLines.Add "    Dim lngClickedColumn As Long"
    codeLines.Add ""
    codeLines.Add "    ' ��ԍ����擾"
    codeLines.Add "    lngClickedColumn = Target.Column"
    codeLines.Add ""
    codeLines.Add "    ' �K�[�h�����i�Ώۗ�ȊO�̓f�t�H���g����̂܂܂ŏI���j"
    codeLines.Add "    If lngClickedColumn <> cfg.COL_CHK And lngClickedColumn <> cfg.COL_OPT Then"
    codeLines.Add "        Exit Sub"
    codeLines.Add "    End If"
    codeLines.Add ""
    codeLines.Add "    ' ���݂̃V�[�g���擾"
    codeLines.Add "    Dim ws As Worksheet"
    codeLines.Add "    Set ws = Me"
    codeLines.Add ""
    codeLines.Add "    ' �ϐ���`"
    codeLines.Add "    Dim lngClickedRow As Long"
    codeLines.Add "    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long"
    codeLines.Add "    Dim varClicked As Variant"
    codeLines.Add "    ' �ꎞ�ϐ���`"
    codeLines.Add "    Dim rngFoundCell As Range"
    codeLines.Add ""
    codeLines.Add "    ' �s�ԍ����擾"
    codeLines.Add "    lngClickedRow = Target.row"
    codeLines.Add ""
    codeLines.Add "    ' �J�n�s�ƏI���s���擾"
    codeLines.Add "    varRangeRows = wbslib.FindDataRangeRows(ws)"
    codeLines.Add "    lngStartRow = varRangeRows(0)"
    codeLines.Add "    lngEndRow = varRangeRows(1)"
    codeLines.Add ""
    codeLines.Add "    ' �J�n�s�ƏI���s��������Ȃ���ΏI��"
    codeLines.Add "    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub"
    codeLines.Add ""
    codeLines.Add "    ' �K�[�h�����i�s�ԍ����w��͈͊O�̏ꍇ�͏I���j"
    codeLines.Add "    If lngClickedRow < lngStartRow Or lngClickedRow > lngEndRow Then"
    codeLines.Add "        Exit Sub"
    codeLines.Add "    End If"
    codeLines.Add ""
    codeLines.Add "    ' CHK ����"
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
    codeLines.Add "    ' OPT ���� (��������)"
    codeLines.Add "    If lngClickedColumn = cfg.COL_OPT Then"
    codeLines.Add "        ' �N���b�N�����l���擾"
    codeLines.Add "        varClicked = ws.Cells(lngClickedRow, cfg.COL_OPT).value"
    codeLines.Add "        If varClicked <> cfg.OPT_MARK_T Then ' �N���b�N���ꂽ�Z���� cfg.OPT_MARK_T �łȂ��ꍇ�̂ݏ���"
    codeLines.Add "            ' lngStartRow ���� lngEndRow �͈̔͂� cfg.OPT_MARK_T �����ŏ��̃Z��������"
    codeLines.Add "            On Error Resume Next"
    codeLines.Add "            Set rngFoundCell = ws.Range(ws.Cells(lngStartRow, cfg.COL_OPT), ws.Cells(lngEndRow, cfg.COL_OPT)).Find(What:=cfg.OPT_MARK_T, LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)"
    codeLines.Add "            On Error GoTo 0"
    codeLines.Add ""
    codeLines.Add "            Application.ScreenUpdating = False"
    codeLines.Add "            Application.Calculation = xlCalculationManual"
    codeLines.Add "            Application.EnableEvents = False"
    codeLines.Add ""
    codeLines.Add "            ' cfg.OPT_MARK_T �����Z�������������� cfg.OPT_MARK_F �ɕύX"
    codeLines.Add "            If Not rngFoundCell Is Nothing Then"
    codeLines.Add "                rngFoundCell.value = cfg.OPT_MARK_F"
    codeLines.Add "            End If"
    codeLines.Add ""
    codeLines.Add "            ' �N���b�N���ꂽ�Z���̒l�� cfg.OPT_MARK_T �ɕύX"
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

    ' �R�[�h��}��
    With vbComp.CodeModule
        For i = 1 To codeLines.Count
            .InsertLines .CountOfLines + 1, codeLines(i)
        Next i
    End With

End Sub

