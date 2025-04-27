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


' �� �f�[�^�s�͈͂��擾���� (��������)
Public Function FindDataRangeRows(ws As Worksheet) As Variant

    Dim startCell As Range
    Dim endCell As Range
    Dim lngStartRow As Long
    Dim lngEndRow As Long

    ' 1. KEY��� "@" �����ŏ��̃Z���������Ɍ���
    On Error Resume Next
    Set startCell = ws.Columns(cfg.COL_KEY).Find(What:="@", LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True)
    On Error GoTo 0

    If startCell Is Nothing Then
        ' "@" ��������Ȃ��ꍇ�́A�J�n�s�� 0 �Ƃ��ď������I����
        FindDataRangeRows = Array(0, 0)
        
        MsgBox "KEY��i" & utils.ConvertColNumberToLetter(cfg.COL_KEY) & "�j�̊J�n�s�}�[�J�[�u@�v��������܂���B" & vbCrLf & _
               "�iKEY�񂪔�\���ƂȂ��Ă���ꍇ�͕\����Ԃɂ��Ă��������j", vbExclamation, "�ʒm"
        
        Exit Function
    Else
        lngStartRow = startCell.row + 1 ' ���ۂ̃f�[�^�J�n�s�� "@" �̎��̍s
    End If

    ' 2. KEY��� "$" �����ŏ��̃Z���� "@" �̎��̍s���獂���Ɍ���
    If lngStartRow > 1 Then ' "@" �����������ꍇ�̂݌���
        On Error Resume Next
        Set endCell = ws.Columns(cfg.COL_KEY).Find(What:="$", LookAt:=xlWhole, LookIn:=xlValues, MatchCase:=True, After:=ws.Cells(lngStartRow - 1, cfg.COL_KEY))
        On Error GoTo 0
        
        If endCell Is Nothing Then
            ' "$" ��������Ȃ��ꍇ�́A�ŏI�s���V�[�g�̍ŏI�s�Ƃ��邩�A����̒l�ɂ��邩����
            lngEndRow = ws.Cells(ws.Rows.Count, cfg.COL_KEY).End(xlUp).row - 1 ' "$" ���Ȃ��Ă��Ō�܂�
            
            MsgBox "KEY��i" & utils.ConvertColNumberToLetter(cfg.COL_KEY) & "�j�̏I���s�}�[�J�[�u$�v��������܂���B" & vbCrLf & _
                   "�iKEY�񂪔�\���ƂȂ��Ă���ꍇ�͕\����Ԃɂ��Ă��������j", vbExclamation, "�ʒm"
        Else
            lngEndRow = endCell.row - 1 ' ���ۂ̃f�[�^�I���s�� "$" �̑O�̍s
        End If
    Else
        ' "@" ���ŏ��̍s�ɂ���ꍇ�Ȃǂ̏���
        lngEndRow = ws.Cells(ws.Rows.Count, cfg.COL_KEY).End(xlUp).row - 1
    End If

    ' ���ʂ�z��ŕԂ�
    FindDataRangeRows = Array(lngStartRow, lngEndRow)

End Function


' �� �G���[�`�F�b�N�����{
Public Sub CheckWbsErrors(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim blnHasError As Boolean
    Dim intErrorCount As Integer
    Dim varData As Variant
    Dim colWbsId As New Collection         ' key=�s�ԍ�������Avalue=WbsId������
    Dim colParentWbsId As New Collection   ' key=WbsId������A value=WbsId�̐e�K�w������
    Dim colError1Count As New Collection   ' key=�s�ԍ�������Avalue=�G���[��1
    Dim colError1Message As New Collection ' key=�s�ԍ�������Avalue=�G���[���b�Z�[�W1
    Dim colError2Count As New Collection   ' key=�s�ԍ�������Avalue=�G���[��2
    Dim colError2Message As New Collection ' key=�s�ԍ�������Avalue=�G���[���b�Z�[�W2
    Dim colError3Count As New Collection   ' key=�s�ԍ�������Avalue=�G���[��3
    Dim colError3Message As New Collection ' key=�s�ԍ�������Avalue=�G���[���b�Z�[�W3
    ' �ꎞ�ϐ���`
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
    
    ' ������
    blnHasError = False

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' ERR ��� E ���N���A
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ERR), ws.Cells(lngEndRow, cfg.COL_ERR)).ClearContents
    With ws.Range(ws.Cells(lngStartRow, cfg.COL_ERR), ws.Cells(lngEndRow, cfg.COL_ERR))
        For Each tmpCell In .Cells
            If Not tmpCell.Comment Is Nothing Then
                tmpCell.Comment.Delete
            End If
        Next tmpCell
    End With
    
    ' �w��͈͂̃f�[�^����x�Ɏ擾
    varData = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).value

    ' �z������[�v���ă`�F�b�N���K�v�ȃf�[�^�����W
    For r = 1 To UBound(varData, 1)
        ' ���ۂ̍s�ԍ����쐬
        tmpRow = r + lngStartRow - 1
        tmpRowIdx = "IDX" & tmpRow
        tmpRecordError = False
        tmpPreCell = ""
        tmpWbsId = ""
        For c = 1 To UBound(varData, 2)
            ' ���ۂ̃J�����ԍ����쐬
            tmpCol = c + cfg.COL_OPT
            ' ���݂̃Z���̒l���擾
            tmpCellValue = varData(r, c)
            If c = 1 Then
                ' # L1 �̏ꍇ #
                If Not IsEmpty(tmpCellValue) And tmpCellValue <> "" Then
                    ' �Z������ł͂Ȃ��ꍇ�AWbsId �ɕ������ǉ�
                    tmpWbsId = tmpWbsId & tmpCellValue
                End If
            ElseIf c = 6 Then
                ' # TASK �̏ꍇ #
                If Not IsEmpty(tmpCellValue) And tmpCellValue <> "" Then
                    ' �Z������ł͂Ȃ��ꍇ�AWbsId �ɕ������ǉ�
                    tmpWbsId = tmpWbsId & ".T" & tmpCellValue
                End If
                ' �����܂ŗ����琳��I��
                colError1Count.Add 0, tmpRowIdx
                colError1Message.Add "", tmpRowIdx
            Else
                ' # L2�`L5 �̏ꍇ #
                If Not IsEmpty(tmpCellValue) And tmpCellValue <> "" Then
                    ' # ���݂̃Z������ł͂Ȃ��ꍇ #
                    If Not IsEmpty(tmpPreCell) And tmpPreCell <> "" Then
                        ' # ���O�̃Z������ł͂Ȃ��ꍇ�AWbsId �ɕ������ǉ� #
                        tmpWbsId = tmpWbsId & "." & tmpCellValue
                    Else
                        ' # ���O�̃Z������̏ꍇ�A�G���[�Ƃ��ď��� #
                        blnHasError = True
                        ' �G���[��������у��b�Z�[�W�ɒǉ����āA�R���N�V�����ɍăZ�b�g
                        colError1Count.Add 1, tmpRowIdx
                        colError1Message.Add "�E�K�w�ԍ��ɖ��i" & utils.ConvertColNumberToLetter(tmpCol - 1) & tmpRow & "�����l�ł͂Ȃ��j" & vbCrLf, tmpRowIdx
                        ' �G���[�s�Ƃ��ăJ�����̃��[�v���I��
                        tmpRecordError = True
                        Exit For
                    End If
                End If
            End If
            tmpPreCell = tmpCellValue
        Next c
        ' ���R�[�h�G���[���������ĂȂ��ꍇ�AWbsId �� ParentWbsId ���R���N�V�����ɒǉ�
        If tmpRecordError = False Then
            ' WbsId ���R���N�V�����ɒǉ�
            colWbsId.Add tmpWbsId, tmpRowIdx
            ' WbsId �̐e�K�w���쐬���A�R���N�V�����ɒǉ�
            tmpDotPosition = InStrRev(tmpWbsId, ".")
            If tmpDotPosition > 0 Then
                tmpParentWbsId = Left(tmpWbsId, tmpDotPosition - 1)
                On Error Resume Next
                colParentWbsId.Add tmpParentWbsId, tmpWbsId
                On Error GoTo 0
            End If
        End If
    Next r
    
    ' ���ׂĂ̍s�𒲍�
    For r = lngStartRow To lngEndRow
        tmpRowIdx = "IDX" & r
        ' �����܂ł̃G���[�������擾����
        tmpErrorCount = 0
        If utils.ExistsColKey(colError1Count, tmpRowIdx) = True Then
            tmpErrorCount = colError1Count.item(tmpRowIdx)
        End If
        ' �܂��G���[���������Ă��Ȃ��s�ŁAWbsId ���o�^����Ă�����̂̂݌���
        If tmpErrorCount = 0 And utils.ExistsColKey(colWbsId, tmpRowIdx) Then
            tmpWbsId = colWbsId.item(tmpRowIdx)
            ' �󕶎���ƂȂ��Ă���WbsId�����O���āA�G���[�`�F�b�N���s��
            If tmpWbsId <> "" Then
                ' WbsId �̐����擾���āA1���ȏ�Ȃ�G���[�Ƃ���
                If utils.ExistsColItemCount(colWbsId, tmpWbsId) > 1 Then
                    blnHasError = True
                    colError2Count.Add 1, tmpRowIdx
                    colError2Message.Add "�E����K�w�ԍ������݁iRow=" & r & "�j" & vbCrLf, tmpRowIdx
                End If
                ' L1 �ȊO�iWbsId �̃h�b�g������j�ŁA�e�K�wWbsId���Ȃ��ꍇ�G���[�Ƃ���
                tmpDotPosition = InStrRev(tmpWbsId, ".")
                If tmpDotPosition > 0 And utils.ExistsColKey(colParentWbsId, tmpWbsId) = True Then
                    tmpParentWbsId = colParentWbsId.item(tmpWbsId)
                    If utils.ExistsColItem(colWbsId, tmpParentWbsId) = False Then
                        blnHasError = True
                        colError3Count.Add 1, tmpRowIdx
                        colError3Message.Add "�E�e�K�w�����݂��Ȃ��iRow=" & r & "�j" & vbCrLf, tmpRowIdx
                    End If
                End If
            End If
        End If
    Next r
    
    ' �W�v�����G���[�ŕ\�����쐬
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
                ' �R�����g�̕��ƍ������蓮�Őݒ�
                With ws.Cells(r, lngCOL_ERR).Comment.Shape
                    .Width = 300   ' ���� 300 �ɐݒ�
                    .Height = 100  ' ������ 100 �ɐݒ�
                End With
            End If
        Next r
    End If
    
    ' �G���[������΃��b�Z�[�W�\��
    If intErrorCount > 0 Then
        MsgBox intErrorCount & " ���ُ̈�����o���܂����B", vbExclamation, "�G���[�`�F�b�N"
    End If
    
End Sub


' �� �\�[�g�p�J�����ɐ������Z�b�g
Public Sub SetFormulaForWbsIdx(ws As Worksheet)
    
    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �������쐬
    strFormula = "=IF(" & cfg.COL_ERR_LABEL & lngStartRow & "=""E"",""ERROR""," & _
                    "IF(" & cfg.COL_L1_LABEL & lngStartRow & "="""",""XXX.XXX.XXX.XXX.XXX.XXX"", CONCAT(TEXT(" & cfg.COL_L1_LABEL & lngStartRow & ",""000"")," & _
                    "IF(" & cfg.COL_L2_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_L2_LABEL & lngStartRow & ",""000""))," & _
                    "IF(" & cfg.COL_L3_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_L3_LABEL & lngStartRow & ",""000""))," & _
                    "IF(" & cfg.COL_L4_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_L4_LABEL & lngStartRow & ",""000""))," & _
                    "IF(" & cfg.COL_L5_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_L5_LABEL & lngStartRow & ",""000""))," & _
                    "IF(" & cfg.COL_TASK_LABEL & lngStartRow & "="""","".---"", ""."" & TEXT(" & cfg.COL_TASK_LABEL & lngStartRow & ",""000"")))))"

    ' �ꊇ�őΏ۔͈͂��擾
    With ws.Range(cfg.COL_WBS_IDX_LABEL & lngStartRow & ":" & cfg.COL_WBS_IDX_LABEL & lngEndRow)
        ' ���l�������ꊇ�Őݒ�
        .NumberFormat = "General"
        ' �������Z�b�g
        .Formula = strFormula
    End With
    
End Sub


' �� WBS-IDX���J�����ɐ������Z�b�g
Public Sub SetFormulaForWbsCnt(ws As Worksheet)
    
    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �������쐬
    strFormula = "=COUNTIF(" & _
                    cfg.COL_WBS_IDX_LABEL & "$" & lngStartRow & ":" & _
                    cfg.COL_WBS_IDX_LABEL & "$" & lngEndRow & "," & _
                    cfg.COL_WBS_IDX_LABEL & lngStartRow & ")"

    ' �ꊇ�őΏ۔͈͂��擾
    With ws.Range(cfg.COL_WBS_CNT_LABEL & lngStartRow & ":" & cfg.COL_WBS_CNT_LABEL & lngEndRow)
        ' ���l�������ꊇ�Őݒ�
        .NumberFormat = "General"
        ' �������Z�b�g
        .Formula = strFormula
    End With
    
End Sub


' �� ID�\���p�J�����ɐ������Z�b�g
Public Sub SetFormulaForWbsId(ws As Worksheet)
   
    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �������쐬
    strFormula = "=IF(" & cfg.COL_ERR_LABEL & lngStartRow & "=""E"",""ERROR""," & _
                    "IF(" & cfg.COL_L1_LABEL & lngStartRow & "="""","""",CONCAT(" & cfg.COL_L1_LABEL & lngStartRow & "," & _
                    "IF(" & cfg.COL_L2_LABEL & lngStartRow & "="""","""","".""&" & cfg.COL_L2_LABEL & lngStartRow & " ), " & _
                    "IF(" & cfg.COL_L3_LABEL & lngStartRow & "="""","""","".""&" & cfg.COL_L3_LABEL & lngStartRow & " ), " & _
                    "IF(" & cfg.COL_L4_LABEL & lngStartRow & "="""","""","".""&" & cfg.COL_L4_LABEL & lngStartRow & " ), " & _
                    "IF(" & cfg.COL_L5_LABEL & lngStartRow & "="""","""","".""&" & cfg.COL_L5_LABEL & lngStartRow & " ), " & _
                    "IF(" & cfg.COL_TASK_LABEL & lngStartRow & "="""","""","".T""&" & cfg.COL_TASK_LABEL & lngStartRow & " ))))"

    ' �ꊇ�őΏ۔͈͂��擾
    With ws.Range(cfg.COL_WBS_ID_LABEL & lngStartRow & ":" & cfg.COL_WBS_ID_LABEL & lngEndRow)
        ' ���l�������ꊇ�Őݒ�
        .NumberFormat = "General"
        ' �������Z�b�g
        .Formula = strFormula
    End With
    
End Sub


' �� ���x���J�����ɐ������Z�b�g
Public Sub SetFormulaForLevel(ws As Worksheet)
   
    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �������쐬
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

    ' �ꊇ�őΏ۔͈͂��擾
    With ws.Range(cfg.COL_LEVEL_LABEL & lngStartRow & ":" & cfg.COL_LEVEL_LABEL & lngEndRow)
        ' ���l�������ꊇ�Őݒ�
        .NumberFormat = "General"
        ' �������Z�b�g
        .Formula = strFormula
    End With
    
End Sub


' �� �t���OT�J�����ɐ������Z�b�g
Public Sub SetFormulaForFlgT(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �������쐬
    strFormula = "=IF(AND(" & cfg.COL_TASK_LABEL & lngStartRow & "<>"""",ISNUMBER(" & cfg.COL_TASK_LABEL & lngStartRow & ")),TRUE,FALSE)"

    ' �ꊇ�őΏ۔͈͂��擾
    With ws.Range(cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow)
        ' ���l�������ꊇ�Őݒ�
        .NumberFormat = "General"
        ' �������Z�b�g
        .Formula = strFormula
    End With

End Sub


' �� �t���OIC�J�����ɐ������Z�b�g
Public Sub SetFormulaForFlgIC(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �������쐬
    strFormula = "=NOT(OR(" & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""�ڊǍ�""," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""�I�グ""," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""�p��""" & "))"

    ' �ꊇ�őΏ۔͈͂��擾
    With ws.Range(cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow)
        ' ���l�������ꊇ�Őݒ�
        .NumberFormat = "General"
        ' �������Z�b�g
        .Formula = strFormula
    End With

End Sub


' �� �t���OPE�J�����ɐ������Z�b�g
Public Sub SetFormulaForFlgPE(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �������쐬
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

    ' �ꊇ�őΏ۔͈͂��擾
    With ws.Range(cfg.COL_FLG_PE_LABEL & lngStartRow & ":" & cfg.COL_FLG_PE_LABEL & lngEndRow)
        ' ���l�������ꊇ�Őݒ�
        .NumberFormat = "General"
        ' �������Z�b�g
        .Formula = strFormula
    End With

End Sub


' �� �t���OCE�J�����ɐ������Z�b�g
Public Sub SetFormulaForFlgCE(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim strFormula As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �������쐬
    strFormula = "=AND(" & _
                    cfg.COL_LEVEL_LABEL & lngStartRow & ">0," & _
                    cfg.COL_FLG_T_LABEL & lngStartRow & "=FALSE," & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & "<>"""",IFERROR(SUMPRODUCT(--(LEFT(" & _
                    cfg.COL_WBS_ID_LABEL & "$" & lngStartRow & ":" & cfg.COL_WBS_ID_LABEL & "$" & lngEndRow & ",LEN(" & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & "&"".""))=" & _
                    cfg.COL_WBS_ID_LABEL & lngStartRow & "&"".""))>0,FALSE))"

    ' �ꊇ�őΏ۔͈͂��擾
    With ws.Range(cfg.COL_FLG_CE_LABEL & lngStartRow & ":" & cfg.COL_FLG_CE_LABEL & lngEndRow)
        ' ���l�������ꊇ�Őݒ�
        .NumberFormat = "General"
        ' �������Z�b�g
        .Formula = strFormula
    End With

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
            MsgBox "�P�ڂ�I���i" & ws.Name & "�j"
        Case 2
            MsgBox "�Q�ڂ�I���i" & ws.Name & "�j"
        Case 3
            MsgBox "�R�ڂ�I���i" & ws.Name & "�j"
        Case 4
            MsgBox "�S�ڂ�I���i" & ws.Name & "�j"
        Case 5
            MsgBox "�T�ڂ�I���i" & ws.Name & "�j"
        Case 6
            MsgBox "�U�ڂ�I���i" & ws.Name & "�j"
        Case 7
            MsgBox "�V�ڂ�I���i" & ws.Name & "�j"
        Case 8
            MsgBox "�W�ڂ�I���i" & ws.Name & "�j"
        Case 9
            MsgBox "�X�ڂ�I���i" & ws.Name & "�j"
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
            MsgBox "�P�ڂ�I���i" & ws.Name & "�j"
        Case 2
            MsgBox "�Q�ڂ�I���i" & ws.Name & "�j"
        Case 3
            MsgBox "�R�ڂ�I���i" & ws.Name & "�j"
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

