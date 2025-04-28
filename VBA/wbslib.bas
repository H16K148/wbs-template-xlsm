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
Public Function ExecCheckWbsHasErrors(ws As Worksheet, _
                                        Optional ByVal blnShowMessage As Boolean = True) As Boolean

    ' �ϐ���`
    Dim blnResult As Boolean
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

    ' �J�n�s�ƏI���s��������Ȃ���΃G���[�I��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then
        ExecCheckWbsHasErrors = True
        Exit Function
    End If

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
    varData = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value

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
            tmpErrorCount = colError1Count.Item(tmpRowIdx)
        End If
        ' �܂��G���[���������Ă��Ȃ��s�ŁAWbsId ���o�^����Ă�����̂̂݌���
        If tmpErrorCount = 0 And utils.ExistsColKey(colWbsId, tmpRowIdx) Then
            tmpWbsId = colWbsId.Item(tmpRowIdx)
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
                    tmpParentWbsId = colParentWbsId.Item(tmpWbsId)
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
                ' �R�����g�̕��ƍ������蓮�Őݒ�
                With ws.Cells(r, cfg.COL_ERR).Comment.Shape
                    .Width = 300   ' ���� 300 �ɐݒ�
                    .Height = 100  ' ������ 100 �ɐݒ�
                End With
            End If
        Next r
    End If
    
    ' �G���[������΃��b�Z�[�W�\��
    If blnShowMessage = True And intErrorCount > 0 Then
        MsgBox intErrorCount & " ���ُ̈�����o���܂����B", vbExclamation, "�G���[�`�F�b�N"
    End If
    
    ExecCheckWbsHasErrors = blnHasError
End Function


' �� �f�[�^�͈͂��\�[�g����
Public Sub ExecSortWbsRange(ws As Worksheet)

    ' �ϐ���`
    Dim rngSortTarget As Range
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

   ' �G���[��`�ŏI��͈̔͂��w��istartRow�`endRow�j
    Set rngSortTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_ERR), ws.Cells(lngEndRow, cfg.COL_LAST))

    ' WBS�C���f�b�N�X����L�[�Ƃ��ď����Ƀ\�[�g
    rngSortTarget.Sort Key1:=ws.Range(cfg.COL_WBS_IDX_LABEL & lngStartRow), Order1:=xlAscending, Header:=xlNo

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
    strFormula = "=CustomFormatWbsIdx(" & _
                    cfg.COL_ERR_LABEL & lngStartRow & "," & _
                    cfg.COL_L1_LABEL & lngStartRow & "," & _
                    cfg.COL_L2_LABEL & lngStartRow & "," & _
                    cfg.COL_L3_LABEL & lngStartRow & "," & _
                    cfg.COL_L4_LABEL & lngStartRow & "," & _
                    cfg.COL_L5_LABEL & lngStartRow & "," & _
                    cfg.COL_TASK_LABEL & lngStartRow & ")"

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
    strFormula = "=CustomFormatWbsId(" & _
                    cfg.COL_ERR_LABEL & lngStartRow & "," & _
                    cfg.COL_L1_LABEL & lngStartRow & "," & _
                    cfg.COL_L2_LABEL & lngStartRow & "," & _
                    cfg.COL_L3_LABEL & lngStartRow & "," & _
                    cfg.COL_L4_LABEL & lngStartRow & "," & _
                    cfg.COL_L5_LABEL & lngStartRow & "," & _
                    cfg.COL_TASK_LABEL & lngStartRow & ")"

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
    strFormula = "=CustomFuncGetLevel(" & _
                    cfg.COL_L1_LABEL & lngStartRow & "," & _
                    cfg.COL_L2_LABEL & lngStartRow & "," & _
                    cfg.COL_L3_LABEL & lngStartRow & "," & _
                    cfg.COL_L4_LABEL & lngStartRow & "," & _
                    cfg.COL_L5_LABEL & lngStartRow & ")"

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
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_DELETED & """," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_TRANSFERRED & """," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_SHELVED & """," & _
                    cfg.COL_WBS_STATUS_LABEL & lngStartRow & "=""" & cfg.WBS_STATUS_REJECTED & """" & "))"

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


' �� �\��H�����W�v���鎮���Z�b�g
Public Sub SetFormulaForPlannedEffort(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �������Z�b�g����f�[�^��p��
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' ���ׂẴ^�X�N�ƊK�w�̃L�[���쐬
    For r = lngStartRow To lngEndRow
        
        ' ���݂̃C���f�b�N�X���擾
        i = r - lngStartRow + 1
        ' ���݂�WBS���x���Z���̒l���擾
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' ���݂�WBS�^�X�N�Z���̒l���擾
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = False Then
            ' # �s���^�X�N�ȊO�̏ꍇ #
            If tmpVarLevelCell = 5 Then
                ' # �s��L5�K�w�̏ꍇ #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # �s��L4�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # �s��L3�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # �s��L2�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # �s��L1�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_PLANNED_EFF), ws.Cells(lngEndRow, cfg.COL_PLANNED_EFF)).Formula = varFormulas
    
    ' L1�W�v�������Z�b�g
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


' �� �ċA�I�ɗ\��H�������v���ăZ�b�g����
Private Sub SetValueRecursiveForPlannedEffort(ws As Worksheet, _
                                                varValues As Variant, _
                                                varHierarchyArray As Variant, _
                                                varFlgIcArray As Variant, _
                                                varPlannedEffortArray As Variant, _
                                                lngTargetIdx As Long)
    
    ' �ϐ���`
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumEffort As Double
    ' �ꎞ�ϐ���`
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' �K�[�h�����i���͂��ꂽ�C���f�b�N�X��0�ȉ��̏ꍇ�͏I���j
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̍s�����z�����C���f�b�N�X���w�肳�ꂽ�ꍇ�͏I���j
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���ɒl�����߂��Ă���ꍇ�͏I���j
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̗񐔂�6�łȂ��ꍇ�͏I���j
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' �w��C���f�b�N�X�̒l���擾
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' �^�X�N��Ԃ̎擾
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' ���x���̎擾
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
        ' # �K�w�ɖ�肪����ꍇ #
        Exit Sub
    End If
    
    ' ���C������
    If blnTargetTask = True Then
        ' # �^�X�N�ɂ͎q�K�w���Ȃ����߁A1���Z�b�g #
        If IsEmpty(varPlannedEffortArray(lngTargetIdx, 1)) Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = varPlannedEffortArray(lngTargetIdx, 1)
        End If
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # �^�X�N�łȂ��ꍇ�A�q�K�w���W�v���Ēl���Z�b�g #
        
        ' �q�K�w���擾
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' �K�[�h�����i�q�K�w�����݂��Ȃ��ꍇ�A0���Z�b�g���ďI���j
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' �K�w�̒l���`�F�b�N���A���Z�b�g�Ȃ�ċA�I�Ɋ֐����Ăяo���A�l���W�v
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


' �� �\��H�����W�v�����l���Z�b�g
Public Sub SetValueForPlannedEffort(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varPlannedEffortArray As Variant
    Dim dblSumEffort As Double
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �l���Z�b�g����f�[�^��p��
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' ���炩���߃`�F�b�N�Ώ۔͈͗�̃f�[�^���擾
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' ���炩����FLG_IC��̃f�[�^���擾
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' ���炩���ߗ\��H����̃f�[�^���擾
    varPlannedEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_PLANNED_EFF), ws.Cells(lngEndRow, cfg.COL_PLANNED_EFF)).Value
    
    ' ���ԂɏW�v���s��
    dblSumEffort = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForPlannedEffort ws, varValues, varHierarchyArray, varFlgIcArray, varPlannedEffortArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            dblSumEffort = dblSumEffort + varValues(i, 1)
        End If
    Next i
    
    ' ���ʂ𔽉f����
    ws.Range(ws.Cells(lngStartRow, cfg.COL_PLANNED_EFF), ws.Cells(lngEndRow, cfg.COL_PLANNED_EFF)).Value = varValues
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngEndRow + 2).Value = dblSumEffort

End Sub


' �� ���эύH�����W�v���鎮���Z�b�g
Public Sub SetFormulaForActualCompletedEffort(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �������Z�b�g����f�[�^��p��
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' ���ׂẴ^�X�N�ƊK�w�̃L�[���쐬
    For r = lngStartRow To lngEndRow
        
        ' ���݂̃C���f�b�N�X���擾
        i = r - lngStartRow + 1
        ' ���݂�WBS���x���Z���̒l���擾
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' ���݂�WBS�^�X�N�Z���̒l���擾
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = False Then
            ' # �s���^�X�N�ȊO�̏ꍇ #
            If tmpVarLevelCell = 5 Then
                ' # �s��L5�K�w�̏ꍇ #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # �s��L4�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # �s��L3�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # �s��L2�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # �s��L1�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_COMPLETED_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_COMPLETED_EFF)).Formula = varFormulas
    
    ' L1�W�v�������Z�b�g
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


' �� �ċA�I�Ɏ��эύH�������v���ăZ�b�g����
Private Sub SetValueRecursiveForActualCompletedEffort(ws As Worksheet, _
                                                        varValues As Variant, _
                                                        varHierarchyArray As Variant, _
                                                        varFlgIcArray As Variant, _
                                                        varActualCompletedEffortArray As Variant, _
                                                        lngTargetIdx As Long)
    
    ' �ϐ���`
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumEffort As Double
    ' �ꎞ�ϐ���`
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' �K�[�h�����i���͂��ꂽ�C���f�b�N�X��0�ȉ��̏ꍇ�͏I���j
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̍s�����z�����C���f�b�N�X���w�肳�ꂽ�ꍇ�͏I���j
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���ɒl�����߂��Ă���ꍇ�͏I���j
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̗񐔂�6�łȂ��ꍇ�͏I���j
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' �w��C���f�b�N�X�̒l���擾
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' �^�X�N��Ԃ̎擾
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' ���x���̎擾
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
        ' # �K�w�ɖ�肪����ꍇ #
        Exit Sub
    End If
    
    ' ���C������
    If blnTargetTask = True Then
        ' # �^�X�N�ɂ͎q�K�w���Ȃ����߁A1���Z�b�g #
        If IsEmpty(varActualCompletedEffortArray(lngTargetIdx, 1)) Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = varActualCompletedEffortArray(lngTargetIdx, 1)
        End If
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # �^�X�N�łȂ��ꍇ�A�q�K�w���W�v���Ēl���Z�b�g #
        
        ' �q�K�w���擾
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' �K�[�h�����i�q�K�w�����݂��Ȃ��ꍇ�A0���Z�b�g���ďI���j
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' �K�w�̒l���`�F�b�N���A���Z�b�g�Ȃ�ċA�I�Ɋ֐����Ăяo���A�l���W�v
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


' �� ���эύH�����W�v�����l���Z�b�g
Public Sub SetValueForActualCompletedEffort(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varActualCompletedEffortArray As Variant
    Dim dblSumEffort As Double
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �l���Z�b�g����f�[�^��p��
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' ���炩���߃`�F�b�N�Ώ۔͈͗�̃f�[�^���擾
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' ���炩����FLG_IC��̃f�[�^���擾
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' ���炩���ߎ��эύH���̃f�[�^���擾
    varActualCompletedEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_COMPLETED_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_COMPLETED_EFF)).Value
    
    ' ���ԂɏW�v���s��
    dblSumEffort = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForActualCompletedEffort ws, varValues, varHierarchyArray, varFlgIcArray, varActualCompletedEffortArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            dblSumEffort = dblSumEffort + varValues(i, 1)
        End If
    Next i
    
    ' ���ʂ𔽉f����
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_COMPLETED_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_COMPLETED_EFF)).Value = varValues
    ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow + 2).Value = dblSumEffort

End Sub


' �� ���юc�H�����W�v���鎮���Z�b�g
Public Sub SetFormulaForActualRemainingEffort(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �������Z�b�g����f�[�^��p��
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' ���ׂẴ^�X�N�ƊK�w�̃L�[���쐬
    For r = lngStartRow To lngEndRow
        
        ' ���݂̃C���f�b�N�X���擾
        i = r - lngStartRow + 1
        ' ���݂�WBS���x���Z���̒l���擾
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' ���݂�WBS�^�X�N�Z���̒l���擾
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = False Then
            ' # �s���^�X�N�ȊO�̏ꍇ #
            If tmpVarLevelCell = 5 Then
                ' # �s��L5�K�w�̏ꍇ #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # �s��L4�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # �s��L3�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # �s��L2�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # �s��L1�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_REMAINING_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_REMAINING_EFF)).Formula = varFormulas
    
    ' L1�W�v�������Z�b�g
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


' �� �ċA�I�Ɏ��юc�H�������v���ăZ�b�g����
Private Sub SetValueRecursiveForActualRemainingEffort(ws As Worksheet, _
                                                        varValues As Variant, _
                                                        varHierarchyArray As Variant, _
                                                        varFlgIcArray As Variant, _
                                                        varActualRemainingEffortArray As Variant, _
                                                        lngTargetIdx As Long)
    
    ' �ϐ���`
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumEffort As Double
    ' �ꎞ�ϐ���`
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' �K�[�h�����i���͂��ꂽ�C���f�b�N�X��0�ȉ��̏ꍇ�͏I���j
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̍s�����z�����C���f�b�N�X���w�肳�ꂽ�ꍇ�͏I���j
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���ɒl�����߂��Ă���ꍇ�͏I���j
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̗񐔂�6�łȂ��ꍇ�͏I���j
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' �w��C���f�b�N�X�̒l���擾
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' �^�X�N��Ԃ̎擾
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' ���x���̎擾
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
        ' # �K�w�ɖ�肪����ꍇ #
        Exit Sub
    End If
    
    ' ���C������
    If blnTargetTask = True Then
        ' # �^�X�N�ɂ͎q�K�w���Ȃ����߁A1���Z�b�g #
        If IsEmpty(varActualRemainingEffortArray(lngTargetIdx, 1)) Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = varActualRemainingEffortArray(lngTargetIdx, 1)
        End If
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # �^�X�N�łȂ��ꍇ�A�q�K�w���W�v���Ēl���Z�b�g #
        
        ' �q�K�w���擾
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' �K�[�h�����i�q�K�w�����݂��Ȃ��ꍇ�A0���Z�b�g���ďI���j
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' �K�w�̒l���`�F�b�N���A���Z�b�g�Ȃ�ċA�I�Ɋ֐����Ăяo���A�l���W�v
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


' �� ���юc�H�����W�v�����l���Z�b�g
Public Sub SetValueForActualRemainingEffort(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varActualRemainingEffortArray As Variant
    Dim dblSumEffort As Double
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �l���Z�b�g����f�[�^��p��
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' ���炩���߃`�F�b�N�Ώ۔͈͗�̃f�[�^���擾
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' ���炩����FLG_IC��̃f�[�^���擾
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' ���炩���ߎ��юc�H���̃f�[�^���擾
    varActualRemainingEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_REMAINING_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_REMAINING_EFF)).Value
    
    ' ���ԂɏW�v���s��
    dblSumEffort = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForActualRemainingEffort ws, varValues, varHierarchyArray, varFlgIcArray, varActualRemainingEffortArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            dblSumEffort = dblSumEffort + varValues(i, 1)
        End If
    Next i
    
    ' ���ʂ𔽉f����
    ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_REMAINING_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_REMAINING_EFF)).Value = varValues
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow + 2).Value = dblSumEffort

End Sub


' �� �^�X�N�i�������W�v���鎮���Z�b�g
Public Sub SetFormulaForTaskProgressRate(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    Dim varNumberFormats() As Variant
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarTaskProgArray As Variant
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String
    Dim tmpStrSumWeightH As String, tmpStrSumWeightT As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �������Z�b�g����f�[�^��p��
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    ReDim varNumberFormats(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' ���炩���ߍ��ڏ�������̃f�[�^���擾
    tmpVarTaskProgArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).Value
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' ���ׂẴ^�X�N�ƊK�w�̃L�[���쐬
    For r = lngStartRow To lngEndRow
        
        ' ���݂̃C���f�b�N�X���擾
        i = r - lngStartRow + 1
        ' ���݂�WBS���x���Z���̒l���擾
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' ���݂�WBS�^�X�N�Z���̒l���擾
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = True Then
            ' # �s���^�X�N�̏ꍇ #
            varNumberFormats(i, 1) = "0.0%"
            varFormulas(i, 1) = tmpVarTaskProgArray(i, 1)
        Else
            ' # �s���^�X�N�ȊO�̏ꍇ #
            If tmpVarLevelCell = 5 Then
                ' # �s��L5�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # �s��L4�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # �s��L3�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # �s��L2�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # �s��L1�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varNumberFormats(i, 1) = "General"
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).NumberFormat = varNumberFormats
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).Formula = varFormulas
    
    ' L1�W�v�������Z�b�g
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


' �� �ċA�I�Ƀ^�X�N�i�������W�v���ăZ�b�g����
Private Sub SetValueRecursiveForTaskProgressRate(ws As Worksheet, _
                                                    varValues As Variant, _
                                                    varHierarchyArray As Variant, _
                                                    varFlgIcArray As Variant, _
                                                    varTaskProgressRateArray As Variant, _
                                                    varTaskWeightArray As Variant, _
                                                    lngTargetIdx As Long)
    
    ' �ϐ���`
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumProgressRate As Double
    Dim intSumWeight As Integer
    ' �ꎞ�ϐ���`
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    Dim tmpIntWeight As Integer
    
    ' �K�[�h�����i���͂��ꂽ�C���f�b�N�X��0�ȉ��̏ꍇ�͏I���j
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̍s�����z�����C���f�b�N�X���w�肳�ꂽ�ꍇ�͏I���j
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���ɒl�����߂��Ă���ꍇ�͏I���j
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̗񐔂�6�łȂ��ꍇ�͏I���j
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' �w��C���f�b�N�X�̒l���擾
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' �^�X�N��Ԃ̎擾
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' ���x���̎擾
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
        ' # �K�w�ɖ�肪����ꍇ #
        Exit Sub
    End If
    
    ' ���C������
    If blnTargetTask = True Then
        ' # �^�X�N�ɂ͎q�K�w���Ȃ����߁A1���Z�b�g #
        If IsEmpty(varTaskProgressRateArray(lngTargetIdx, 1)) Then
            varValues(lngTargetIdx, 1) = 0
        Else
            varValues(lngTargetIdx, 1) = varTaskProgressRateArray(lngTargetIdx, 1)
        End If
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # �^�X�N�łȂ��ꍇ�A�q�K�w���W�v���Ēl���Z�b�g #
        
        ' �q�K�w���擾
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' �K�[�h�����i�q�K�w�����݂��Ȃ��ꍇ�A0���Z�b�g���ďI���j
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' �K�w�̒l���`�F�b�N���A���Z�b�g�Ȃ�ċA�I�Ɋ֐����Ăяo���A�l���W�v
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


' �� �^�X�N�i�������W�v�����l���Z�b�g
Public Sub SetValueForTaskProgressRate(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varTaskProgressRateArray As Variant
    Dim varTaskWeightArray As Variant
    Dim dblSumRate As Double
    Dim intSumWeight As Integer
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpIntWeight As Integer

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �l���Z�b�g����f�[�^��p�Ӂi1:�W�v���ʁm�H���i�����n�A2:���x���j
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' ���炩���߃`�F�b�N�Ώ۔͈͗�̃f�[�^���擾
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' ���炩����FLG_IC��̃f�[�^���擾
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' ���炩���ߍ��ڏ�������̃f�[�^���擾
    varTaskProgressRateArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).Value
    ' ���炩���ߍ��ډ��d��̃f�[�^���擾
    varTaskWeightArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_WGT), ws.Cells(lngEndRow, cfg.COL_TASK_WGT)).Value
    
    ' ���ԂɏW�v���s��
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
    
    ' ���ʂ𔽉f����
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_PROG), ws.Cells(lngEndRow, cfg.COL_TASK_PROG)).Value = varValues
    If intSumWeight = 0 Then
        ws.Range(cfg.COL_TASK_PROG_LABEL & lngEndRow + 2).Value = 0
    Else
        ws.Range(cfg.COL_TASK_PROG_LABEL & lngEndRow + 2).Value = dblSumRate / intSumWeight
    End If
    
End Sub


' �� �H���i�������W�v���鎮���Z�b�g
Public Sub SetFormulaForEffortProgressRate(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String
    Dim tmpStrCountH As String, tmpStrCountT As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �������Z�b�g����f�[�^��p��
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' ���ׂẴ^�X�N�ƊK�w�̃L�[���쐬
    For r = lngStartRow To lngEndRow
        
        ' ���݂̃C���f�b�N�X���擾
        i = r - lngStartRow + 1
        ' ���݂�WBS���x���Z���̒l���擾
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' ���݂�WBS�^�X�N�Z���̒l���擾
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = True Then
            ' # �s���^�X�N�̏ꍇ #
            tmpStrFormula = "=" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r & _
            "/IF(" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r & "+" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r & "=0," & _
            "1," & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r & "+" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r & ")"
            ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
            varFormulas(i, 1) = tmpStrFormula
        Else
            ' # �s���^�X�N�ȊO�̏ꍇ #
            If tmpVarLevelCell = 5 Then
                ' # �s��L5�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # �s��L4�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # �s��L3�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # �s��L2�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # �s��L1�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_EFFORT_PROG), ws.Cells(lngEndRow, cfg.COL_EFFORT_PROG)).Formula = varFormulas
    
    ' L1�W�v�������Z�b�g
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


' �� �ċA�I�ɍH���i�������W�v���ăZ�b�g����
Private Sub SetValueRecursiveForEffortProgressRate(ws As Worksheet, _
                                                    varValues As Variant, _
                                                    varHierarchyArray As Variant, _
                                                    varFlgIcArray As Variant, _
                                                    varActualRemainingEffortArray As Variant, _
                                                    varActualCompletedEffortArray As Variant, _
                                                    lngTargetIdx As Long)
    
    ' �ϐ���`
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim dblSumProgressRate As Double
    Dim intSumCount As Integer
    ' �ꎞ�ϐ���`
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    Dim tmpDblActualRemainingEffort As Double
    Dim tmpDblActualCompletedEffort As Double
    
    ' �K�[�h�����i���͂��ꂽ�C���f�b�N�X��0�ȉ��̏ꍇ�͏I���j
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̍s�����z�����C���f�b�N�X���w�肳�ꂽ�ꍇ�͏I���j
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���ɒl�����߂��Ă���ꍇ�͏I���j
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̗񐔂�6�łȂ��ꍇ�͏I���j
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' �w��C���f�b�N�X�̒l���擾
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' �^�X�N��Ԃ̎擾
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' ���x���̎擾
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
        ' # �K�w�ɖ�肪����ꍇ #
        Exit Sub
    End If
    
    ' ���C������
    If blnTargetTask = True Then
        ' # �^�X�N�ɂ͎q�K�w���Ȃ����߁A1���Z�b�g #
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
        ' # �^�X�N�łȂ��ꍇ�A�q�K�w���W�v���Ēl���Z�b�g #
        
        ' �q�K�w���擾
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' �K�[�h�����i�q�K�w�����݂��Ȃ��ꍇ�A0���Z�b�g���ďI���j
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' �K�w�̒l���`�F�b�N���A���Z�b�g�Ȃ�ċA�I�Ɋ֐����Ăяo���A�l���W�v
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


' �� �H���i�������W�v�����l���Z�b�g
Public Sub SetValueForEffortProgressRate(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varActualRemainingEffortArray As Variant
    Dim varActualCompletedEffortArray As Variant
    Dim dblSumRate As Double
    Dim intSumCount As Integer
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �l���Z�b�g����f�[�^��p�Ӂi1:�W�v���ʁm�H���i�����n�A2:���x���j
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' ���炩���߃`�F�b�N�Ώ۔͈͗�̃f�[�^���擾
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' ���炩����FLG_IC��̃f�[�^���擾
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' ���炩���ߎ��юc�H���̃f�[�^���擾
    varActualRemainingEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_REMAINING_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_REMAINING_EFF)).Value
    ' ���炩���ߎ��эύH���̃f�[�^���擾
    varActualCompletedEffortArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_ACTUAL_COMPLETED_EFF), ws.Cells(lngEndRow, cfg.COL_ACTUAL_COMPLETED_EFF)).Value
    
    ' ���ԂɏW�v���s��
    dblSumRate = 0
    intSumCount = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForEffortProgressRate ws, varValues, varHierarchyArray, varFlgIcArray, varActualRemainingEffortArray, varActualCompletedEffortArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            dblSumRate = dblSumRate + varValues(i, 1)
            intSumCount = intSumCount + 1
        End If
    Next i
    
    ' ���ʂ𔽉f����
    ws.Range(ws.Cells(lngStartRow, cfg.COL_EFFORT_PROG), ws.Cells(lngEndRow, cfg.COL_EFFORT_PROG)).Value = varValues
    If intSumCount = 0 Then
        ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2).Value = 0
    Else
        ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2).Value = dblSumRate / intSumCount
    End If
    
End Sub


' �� �^�X�N���v�������W�v���鎮���Z�b�g
Public Sub SetFormulaForTaskCount(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �������Z�b�g����f�[�^��p��
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' ���ׂẴ^�X�N�ƊK�w�̃L�[���쐬
    For r = lngStartRow To lngEndRow
        
        ' ���݂̃C���f�b�N�X���擾
        i = r - lngStartRow + 1
        ' ���݂�WBS���x���Z���̒l���擾
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' ���݂�WBS�^�X�N�Z���̒l���擾
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = True Then
            ' # �s���^�X�N�̏ꍇ #
            varFormulas(i, 1) = 1
        Else
            ' # �s���^�X�N�ȊO�̏ꍇ #
            If tmpVarLevelCell = 5 Then
                ' # �s��L5�K�w�̏ꍇ #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # �s��L4�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # �s��L3�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # �s��L2�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # �s��L1�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_COUNT), ws.Cells(lngEndRow, cfg.COL_TASK_COUNT)).Formula = varFormulas
    
    ' L1�W�v�������Z�b�g
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


' �� �ċA�I�Ƀ^�X�N���v�����J�E���g���ăZ�b�g����
Private Sub SetValueRecursiveForTaskCount(ws As Worksheet, _
                                            varValues As Variant, _
                                            varHierarchyArray As Variant, _
                                            varFlgIcArray As Variant, _
                                            lngTargetIdx As Long)
    
    ' �ϐ���`
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim lngSumCount As Long
    ' �ꎞ�ϐ���`
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' �K�[�h�����i���͂��ꂽ�C���f�b�N�X��0�ȉ��̏ꍇ�͏I���j
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̍s�����z�����C���f�b�N�X���w�肳�ꂽ�ꍇ�͏I���j
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���ɒl�����߂��Ă���ꍇ�͏I���j
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̗񐔂�6�łȂ��ꍇ�͏I���j
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' �w��C���f�b�N�X�̒l���擾
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' �^�X�N��Ԃ̎擾
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' ���x���̎擾
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
        ' # �K�w�ɖ�肪����ꍇ #
        Exit Sub
    End If
    
    ' ���C������
    If blnTargetTask = True Then
        ' # �^�X�N�ɂ͎q�K�w���Ȃ����߁A1���Z�b�g #
        varValues(lngTargetIdx, 1) = 1
        varValues(lngTargetIdx, 2) = 6
    Else
        ' # �^�X�N�łȂ��ꍇ�A�q�K�w���W�v���Ēl���Z�b�g #
        
        ' �q�K�w���擾
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' �K�[�h�����i�q�K�w�����݂��Ȃ��ꍇ�A0���Z�b�g���ďI���j
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' �K�w�̒l���`�F�b�N���A���Z�b�g�Ȃ�ċA�I�Ɋ֐����Ăяo���A�l���W�v
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


' �� �^�X�N���v�������W�v���鎮���Z�b�g
Public Sub SetValueForTaskCount(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim lngSumCount As Long
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �l���Z�b�g����f�[�^��p��
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' ���炩���߃`�F�b�N�Ώ۔͈͗�̃f�[�^���擾
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' ���炩����FLG_IC��̃f�[�^���擾
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    
    ' ���ԂɏW�v���s��
    lngSumCount = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForTaskCount ws, varValues, varHierarchyArray, varFlgIcArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            lngSumCount = lngSumCount + varValues(i, 1)
        End If
    Next i
    
    ' ���ʂ𔽉f����
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_COUNT), ws.Cells(lngEndRow, cfg.COL_TASK_COUNT)).Value = varValues
    ws.Range(cfg.COL_TASK_COUNT_LABEL & lngEndRow + 2).Value = lngSumCount

End Sub


' �� �^�X�N�����������W�v���鎮���Z�b�g
Public Sub SetFormulaForTaskCompCount(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
    Dim tmpVarLevelArray As Variant, tmpVarLevelCell As Variant
    Dim tmpVarTaskArray As Variant, tmpVarTaskCell As Variant
    Dim tmpStrBoolArrayH As String, tmpStrBoolArrayT As String

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �������Z�b�g����f�[�^��p��
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To 1)
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' ���ׂẴ^�X�N�ƊK�w�̃L�[���쐬
    For r = lngStartRow To lngEndRow
        
        ' ���݂̃C���f�b�N�X���擾
        i = r - lngStartRow + 1
        ' ���݂�WBS���x���Z���̒l���擾
        tmpVarLevelCell = tmpVarLevelArray(i, 1)
        ' ���݂�WBS�^�X�N�Z���̒l���擾
        tmpVarTaskCell = tmpVarTaskArray(i, 1)
        
        If tmpVarTaskCell = True Then
            ' # �s���^�X�N�̏ꍇ #
            tmpStrFormula = "=IF(" & cfg.COL_WBS_STATUS_LABEL & r & "=""" & cfg.WBS_STATUS_COMPLETED & """,1,0)"
            varFormulas(i, 1) = tmpStrFormula
        Else
            ' # �s���^�X�N�ȊO�̏ꍇ #
            If tmpVarLevelCell = 5 Then
                ' # �s��L5�K�w�̏ꍇ #
                tmpStrBoolArrayT = "(" & cfg.COL_L1_LABEL & lngStartRow & ":" & cfg.COL_L1_LABEL & lngEndRow & "=" & cfg.COL_L1_LABEL & r & ")*" & _
                          "(" & cfg.COL_L2_LABEL & lngStartRow & ":" & cfg.COL_L2_LABEL & lngEndRow & "=" & cfg.COL_L2_LABEL & r & ")*" & _
                          "(" & cfg.COL_L3_LABEL & lngStartRow & ":" & cfg.COL_L3_LABEL & lngEndRow & "=" & cfg.COL_L3_LABEL & r & ")*" & _
                          "(" & cfg.COL_L4_LABEL & lngStartRow & ":" & cfg.COL_L4_LABEL & lngEndRow & "=" & cfg.COL_L4_LABEL & r & ")*" & _
                          "(" & cfg.COL_L5_LABEL & lngStartRow & ":" & cfg.COL_L5_LABEL & lngEndRow & "=" & cfg.COL_L5_LABEL & r & ")*" & _
                          "(" & cfg.COL_FLG_T_LABEL & lngStartRow & ":" & cfg.COL_FLG_T_LABEL & lngEndRow & "=TRUE)*" & _
                          "(" & cfg.COL_FLG_IC_LABEL & lngStartRow & ":" & cfg.COL_FLG_IC_LABEL & lngEndRow & "=TRUE)"
                tmpStrFormula = "=SUM(FILTER(" & cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow & "," & tmpStrBoolArrayT & ",0))"
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 4 Then
                ' # �s��L4�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 3 Then
                ' # �s��L3�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 2 Then
                ' # �s��L2�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
            If tmpVarLevelCell = 1 Then
                ' # �s��L1�K�w�̏ꍇ #
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
                ' �w�肳�ꂽ��̃Z���ɐ������Z�b�g
                varFormulas(i, 1) = tmpStrFormula
            End If
        End If
    Next r
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_COMP_COUNT), ws.Cells(lngEndRow, cfg.COL_TASK_COMP_COUNT)).Formula = varFormulas
    
    ' L1�W�v�������Z�b�g
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


' �� �ċA�I�Ƀ^�X�N���������J�E���g���ăZ�b�g����
Private Sub SetValueRecursiveForTaskCompCount(ws As Worksheet, _
                                                varValues As Variant, _
                                                varHierarchyArray As Variant, _
                                                varFlgIcArray As Variant, _
                                                varWbsStatusArray As Variant, _
                                                lngTargetIdx As Long)
    
    ' �ϐ���`
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    Dim lngSumCount As Long
    ' �ꎞ�ϐ���`
    Dim tmpVar As Variant
    Dim tmpColChildIdxs As New Collection
    Dim tmpVarChildIdx As Variant
    
    ' �K�[�h�����i���͂��ꂽ�C���f�b�N�X��0�ȉ��̏ꍇ�͏I���j
    If lngTargetIdx <= 0 Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̍s�����z�����C���f�b�N�X���w�肳�ꂽ�ꍇ�͏I���j
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���ɒl�����߂��Ă���ꍇ�͏I���j
    If Not IsEmpty(varValues(lngTargetIdx, 1)) Then
        Exit Sub
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̗񐔂�6�łȂ��ꍇ�͏I���j
    If UBound(varHierarchyArray, 2) <> 6 Then
        Exit Sub
    End If
    
    ' �w��C���f�b�N�X�̒l���擾
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' �^�X�N��Ԃ̎擾
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' ���x���̎擾
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
        ' # �K�w�ɖ�肪����ꍇ #
        Exit Sub
    End If
    
    ' ���C������
    If blnTargetTask = True Then
        ' # �^�X�N�ɂ͎q�K�w���Ȃ����߁A1���Z�b�g #
        If varWbsStatusArray(lngTargetIdx, 1) = cfg.WBS_STATUS_COMPLETED Then
            varValues(lngTargetIdx, 1) = 1
            varValues(lngTargetIdx, 2) = 6
        Else
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = 6
        End If
    Else
        ' # �^�X�N�łȂ��ꍇ�A�q�K�w���W�v���Ēl���Z�b�g #
        
        ' �q�K�w���擾
        Set tmpColChildIdxs = GetTargetChildIdxs(varHierarchyArray, lngTargetIdx)
        
        ' �K�[�h�����i�q�K�w�����݂��Ȃ��ꍇ�A0���Z�b�g���ďI���j
        If tmpColChildIdxs.Count = 0 Then
            varValues(lngTargetIdx, 1) = 0
            varValues(lngTargetIdx, 2) = intTargetLevel
            Exit Sub
        End If
        
        ' �K�w�̒l���`�F�b�N���A���Z�b�g�Ȃ�ċA�I�Ɋ֐����Ăяo���A�l���W�v
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


' �� �^�X�N�����������W�v���鎮���Z�b�g
Public Sub SetValueForTaskCompCount(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varValues() As Variant
    Dim varHierarchyArray As Variant
    Dim varFlgIcArray As Variant
    Dim varWbsStatusArray As Variant
    Dim lngSumCount As Long
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �l���Z�b�g����f�[�^��p��
    ReDim varValues(1 To lngEndRow - lngStartRow + 1, 1 To 2)
    
    ' ���炩���߃`�F�b�N�Ώ۔͈͗�̃f�[�^���擾
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' ���炩����FLG_IC��̃f�[�^���擾
    varFlgIcArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_IC), ws.Cells(lngEndRow, cfg.COL_FLG_IC)).Value
    ' ���炩����WBS�X�e�[�^�X��̃f�[�^���擾
    varWbsStatusArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_WBS_STATUS), ws.Cells(lngEndRow, cfg.COL_WBS_STATUS)).Value
    
    ' ���ԂɏW�v���s��
    lngSumCount = 0
    For i = 1 To UBound(varHierarchyArray, 1)
        SetValueRecursiveForTaskCompCount ws, varValues, varHierarchyArray, varFlgIcArray, varWbsStatusArray, i
        If Not IsEmpty(varFlgIcArray(i, 1)) And varFlgIcArray(i, 1) = True And varValues(i, 2) = 1 Then
            lngSumCount = lngSumCount + varValues(i, 1)
        End If
    Next i
    
    ' ���ʂ𔽉f����
    ws.Range(ws.Cells(lngStartRow, cfg.COL_TASK_COMP_COUNT), ws.Cells(lngEndRow, cfg.COL_TASK_COMP_COUNT)).Value = varValues
    ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow + 2).Value = lngSumCount

End Sub


' �� �I�𒆂̃I�v�V�����{�^������s�ԍ����擾
Private Function GetCheckedOptSingleRow(ws As Worksheet) As Long
    
    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim rngFoundCell As Range
    ' �ꎞ�ϐ���`
    Dim r As Long

    ' �J�n�s�ƏI���s���擾
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then
        GetCheckedOptSingleRow = 0
        Exit Function
    End If
    
    ' lngStartRow ���� lngEndRow �͈̔͂� cfg.OPT_MARK_T �����ŏ��̃Z��������
    On Error Resume Next
    Set rngFoundCell = ws.Range(ws.Cells(lngStartRow, cfg.COL_OPT), ws.Cells(lngEndRow, cfg.COL_OPT)).Find( _
        What:=cfg.OPT_MARK_T, _
        LookAt:=xlWhole, _
        LookIn:=xlValues, _
        MatchCase:=True _
    )
    On Error GoTo 0
    
    ' �Z�������������ꍇ
    If Not rngFoundCell Is Nothing Then
        GetCheckedOptSingleRow = rngFoundCell.row
        Exit Function
    End If

    ' �����܂ŗ�����`�F�b�N�Ȃ�
    GetCheckedOptSingleRow = 0
End Function


' �� �I�𒆂̃`�F�b�N�{�b�N�X����s�ԍ��R���N�V�������擾
Private Function GetCheckedChkMultpleRows(ws As Worksheet) As Collection

    ' �ϐ���`
    Dim rowCollection As New Collection
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varData As Variant
    ' �ꎞ�ϐ���`
    Dim r As Long

    ' �J�n�s�ƏI���s���擾
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then
        Set GetCheckedChkMultpleRows = rowCollection
        Exit Function
    End If

    ' �Y���͈͂̃Z���f�[�^���ꊇ�Ŕz��Ɋi�[
    varData = ws.Range(ws.Cells(lngStartRow, cfg.COL_CHK), ws.Cells(lngEndRow, cfg.COL_CHK)).Value
    
    ' �z������[�v���Ĉ�v����s�ԍ������W
    For r = 1 To UBound(varData, 1)  ' �z��̍s�����������[�v
        If varData(r, 1) = cfg.CHK_MARK_T Then
            rowCollection.Add lngStartRow + r - 1 ' ���ۂ̍s�ԍ���ǉ�
        End If
    Next r

    ' ���ʂƂ��čs�ԍ��̃R���N�V������Ԃ�
    Set GetCheckedChkMultpleRows = rowCollection
End Function


' �� �I�������s�̉��Ɉ�s�ǉ�
Public Sub ExecInsertRowBelowSelection(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long
    Dim lngSelectedRow As Long
    ' �ꎞ�ϐ���`
    Dim tmpLngCol As Long
    
    ' �J�n�s�ƏI���s���擾
    varRangeRows = FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    
    ' �s��ǉ�
    lngSelectedRow = GetCheckedOptSingleRow(ws)
    If lngSelectedRow <> 0 Then
        ' �s��ǉ�
        ws.Rows(lngSelectedRow + 1).Insert Shift:=xlDown
        ' 1�񂸂`�F�b�N���āA��{���������R�s�[
        For tmpLngCol = cfg.COL_WBS_IDX To cfg.COL_WBS_ID
            If ws.Cells(lngSelectedRow, tmpLngCol).HasFormula Then
                ws.Cells(lngSelectedRow + 1, tmpLngCol).Formula = ws.Cells(lngSelectedRow, tmpLngCol).Formula
            End If
        Next tmpLngCol
    Else
        MsgBox "�I�����Ă��������iOPT)�B", vbExclamation, "�ʒm"
    End If

End Sub


' �� �I���s�̍ŏI���x��ID���C���N�������g
Public Sub ExecIncrementSelectedLastLevel(ws As Worksheet)

    ' �ϐ���`
    Dim lngSelectedRow As Long, intSelectedLevel As Integer, blnSelectedIsTask As Boolean
    Dim varSelectedL1 As Variant, varSelectedL2 As Variant, varSelectedL3 As Variant, varSelectedL4 As Variant, varSelectedL5 As Variant, varSelectedTask As Variant
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colTargetIdx As New Collection
    Dim rngHierarchy As Range
    Dim varHierarchyArray As Variant
    Dim varLevelArray As Variant
    Dim varTaskArray As Variant
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpVarIdx As Variant
    
    ' �J�n�s�ƏI���s���擾
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �I�������s�̔ԍ����擾
    lngSelectedRow = GetCheckedOptSingleRow(ws)
    
    ' �K�[�h�����i���I���̏ꍇ�́A���b�Z�[�W���o���ďI���j
    If lngSelectedRow = 0 Then
        MsgBox "�I�����Ă��������iOPT)�B", vbExclamation, "�ʒm"
        Exit Sub
    End If
    
    ' ���炩���ߍX�V�Ώ۔͈͗�̃f�[�^���擾
    Set rngHierarchy = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    varHierarchyArray = rngHierarchy.Value
    ' ���炩����WBS���x����̃f�[�^���擾
    varLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    varTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' �I�������s�̃��x�����擾
    intSelectedLevel = varLevelArray(lngSelectedRow - lngStartRow + 1, 1)
    ' �I�������s���^�X�N���ǂ����擾
    blnSelectedIsTask = varTaskArray(lngSelectedRow - lngStartRow + 1, 1)
    ' �I�������s�̃f�[�^���擾
    varSelectedL1 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 1)
    varSelectedL2 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 2)
    varSelectedL3 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 3)
    varSelectedL4 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 4)
    varSelectedL5 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 5)
    varSelectedTask = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 6)
    
    ' �X�V�Ώ۔͈͗�̃f�[�^���X�V
    If blnSelectedIsTask = True Then
        ' # �I���s���^�X�N�̏ꍇ #
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X�̂ݒl���X�V����
        For Each tmpVarIdx In colTargetIdx
            varHierarchyArray(tmpVarIdx, 6) = varHierarchyArray(tmpVarIdx, 6) + 1
        Next tmpVarIdx
    Else
        ' # �I���s���^�X�N�łȂ��ꍇ #
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X�̂ݒl���X�V����
        For Each tmpVarIdx In colTargetIdx
            varHierarchyArray(tmpVarIdx, intSelectedLevel) = varHierarchyArray(tmpVarIdx, intSelectedLevel) + 1
        Next tmpVarIdx
    End If
    
    ' �f�[�^�̍X�V���ʂ𔽉f
    rngHierarchy.Value = varHierarchyArray

End Sub


' �� �I���s�̍ŏI���x��ID���f�N�������g
Public Sub ExecDecrementSelectedLastLevel(ws As Worksheet)

    ' �ϐ���`
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
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long, v As Long
    Dim tmpVarIdx As Variant
    Dim tmpVarValue As Variant, tmpBlnExist As Boolean
    
    ' �J�n�s�ƏI���s���擾
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �I�������s�̔ԍ����擾
    lngSelectedRow = GetCheckedOptSingleRow(ws)
    
    ' �K�[�h�����i���I���̏ꍇ�́A���b�Z�[�W���o���ďI���j
    If lngSelectedRow = 0 Then
        MsgBox "�I�����Ă��������iOPT)�B", vbExclamation, "�ʒm"
        Exit Sub
    End If
    
    ' ���炩���ߍX�V�Ώ۔͈͗�̃f�[�^���擾
    Set rngHierarchy = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    varHierarchyArray = rngHierarchy.Value
    ' ���炩����WBS���x����̃f�[�^���擾
    varLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    varTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    
    ' �I�������s�̃��x�����擾
    intSelectedLevel = varLevelArray(lngSelectedRow - lngStartRow + 1, 1)
    ' �I�������s���^�X�N���ǂ����擾
    blnSelectedIsTask = varTaskArray(lngSelectedRow - lngStartRow + 1, 1)
    ' �I�������s�̖����̒l���擾
    If blnSelectedIsTask Then
        varSelectedLastValue = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 6)
    Else
        varSelectedLastValue = varHierarchyArray(lngSelectedRow - lngStartRow + 1, intSelectedLevel)
    End If
    
    ' �I�������s�̃f�[�^���擾
    varSelectedL1 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 1)
    varSelectedL2 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 2)
    varSelectedL3 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 3)
    varSelectedL4 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 4)
    varSelectedL5 = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 5)
    varSelectedTask = varHierarchyArray(lngSelectedRow - lngStartRow + 1, 6)
    
    ' �X�V�Ώ۔͈͗�̃f�[�^���X�V
    If blnSelectedIsTask = True Then
        ' # �I���s���^�X�N�̏ꍇ #
        ' �ΏۂƂȂ�l���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �l�R���N�V����������ŏ��̑��݂��Ȃ��l���擾
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
        ' �K�[�h�����i�󂫔ԍ������݂��Ȃ�������I���j
        If lngFirstMissingFoundValue = 0 Then
            MsgBox "�󂫔ԍ�������܂���B", vbExclamation, "�ʒm"
            Exit Sub
        End If
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X�̂ݒl���X�V����
        For Each tmpVarIdx In colTargetIdx
            varHierarchyArray(tmpVarIdx, 6) = varHierarchyArray(tmpVarIdx, 6) - 1
        Next tmpVarIdx
    Else
        ' # �I���s���^�X�N�łȂ��ꍇ #
        ' �ΏۂƂȂ�l���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �l�R���N�V����������ŏ��̑��݂��Ȃ��l���擾
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
        ' �K�[�h�����i�󂫔ԍ������݂��Ȃ�������I���j
        If lngFirstMissingFoundValue = 0 Then
            MsgBox "�󂫔ԍ�������܂���B", vbExclamation, "�ʒm"
            Exit Sub
        End If
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X�̂ݒl���X�V����
        For Each tmpVarIdx In colTargetIdx
            varHierarchyArray(tmpVarIdx, intSelectedLevel) = varHierarchyArray(tmpVarIdx, intSelectedLevel) - 1
        Next tmpVarIdx
    End If
    
    ' �f�[�^�̍X�V���ʂ𔽉f
    rngHierarchy.Value = varHierarchyArray

End Sub


' �� �`�F�b�N�����Q�_�̍ŏI���x��ID����������
Public Sub ExecSwapCheckedLastLevel(ws As Worksheet)

    ' �ϐ���`
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
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long, v As Long
    
    ' �J�n�s�ƏI���s���擾
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �`�F�b�N����Ă���s�ԍ����擾
    Set colCheckedRows = GetCheckedChkMultpleRows(ws)
    
    ' �K�[�h�����i�`�F�b�N���Q�łȂ������ꍇ�͏I���j
    If colCheckedRows.Count <> 2 Then
        MsgBox "�����������Q���`�F�b�N���Ă��������iCHK)�B" & vbCrLf & "�i" & colCheckedRows.Count & " �ӏ����I������Ă��܂��j", vbExclamation, "�ʒm"
        Exit Sub
    End If
    
    ' ���炩���ߍX�V�Ώ۔͈͗�̃f�[�^���擾
    Set rngHierarchy = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    varHierarchyArray = rngHierarchy.Value
    ' ���炩����WBS���x����̃f�[�^���擾
    varLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).Value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    varTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).Value
    ' ���炩����WBS-ID��̃f�[�^���擾
    varIdArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_WBS_ID), ws.Cells(lngEndRow, cfg.COL_WBS_ID)).Value

    ' �� �`�F�b�N1�����W
    lngChecked1Row = colCheckedRows.Item(1)
    ' �I�������s�̃��x�����擾
    intChecked1Level = varLevelArray(lngChecked1Row - lngStartRow + 1, 1)
    ' �I�������s���^�X�N���ǂ����擾
    blnChecked1IsTask = varTaskArray(lngChecked1Row - lngStartRow + 1, 1)
    ' �I�������s�̖����̒l���擾
    If blnChecked1IsTask Then
        varChecked1LastValue = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 6)
    Else
        varChecked1LastValue = varHierarchyArray(lngChecked1Row - lngStartRow + 1, intChecked1Level)
    End If
    ' �I�������s��ID���擾
    varChecked1Id = varIdArray(lngChecked1Row - lngStartRow + 1, 1)
    ' �I�������s�̃f�[�^���擾
    varChecked1L1 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 1)
    varChecked1L2 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 2)
    varChecked1L3 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 3)
    varChecked1L4 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 4)
    varChecked1L5 = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 5)
    varChecked1Task = varHierarchyArray(lngChecked1Row - lngStartRow + 1, 6)

    ' �� �`�F�b�N2�����W
    lngChecked2Row = colCheckedRows.Item(2)
    ' �I�������s�̃��x�����擾
    intChecked2Level = varLevelArray(lngChecked2Row - lngStartRow + 1, 1)
    ' �I�������s���^�X�N���ǂ����擾
    blnChecked2IsTask = varTaskArray(lngChecked2Row - lngStartRow + 1, 1)
    ' �I�������s�̖����̒l���擾
    If blnChecked2IsTask Then
        varChecked2LastValue = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 6)
    Else
        varChecked2LastValue = varHierarchyArray(lngChecked2Row - lngStartRow + 1, intChecked2Level)
    End If
    ' �I�������s��ID���擾
    varChecked2Id = varIdArray(lngChecked2Row - lngStartRow + 1, 1)
    ' �I�������s�̃f�[�^���擾
    varChecked2L1 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 1)
    varChecked2L2 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 2)
    varChecked2L3 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 3)
    varChecked2L4 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 4)
    varChecked2L5 = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 5)
    varChecked2Task = varHierarchyArray(lngChecked2Row - lngStartRow + 1, 6)
    
    ' �K�[�h�����i�Q�̊K�w�y�у^�X�N���ۂ�����v���Ȃ��ꍇ�A�I���j
    If (intChecked1Level <> intChecked2Level) Or (blnChecked1IsTask <> blnChecked2IsTask) Then
        MsgBox "�������̊K�w����у^�X�N���ǂ�������v���܂���iCHK)�B" & vbCrLf & _
        vbCrLf & "�`�F�b�N1: �K�w=" & intChecked1Level & ", �^�X�N=" & blnChecked1IsTask & _
        vbCrLf & "�`�F�b�N2: �K�w=" & intChecked2Level & ", �^�X�N=" & blnChecked2IsTask & _
        "", vbExclamation, "�ʒm"
        Exit Sub
    End If
    
    ' �K�[�h�����i�Q�̖����ԍ��ȊO�̊K�w�ԍ�����v���Ȃ��ꍇ�A�I���j
    If blnChecked1IsTask = True Then
        If varChecked1L1 <> varChecked2L1 Or varChecked1L2 <> varChecked2L2 Or varChecked1L3 <> varChecked2L3 Or varChecked1L4 <> varChecked2L4 Or varChecked1L5 <> varChecked2L5 Then
            MsgBox "�������̖����ԍ��ȊO�̊K�w�ԍ�����v���܂���iCHK)�B" & vbCrLf & _
            vbCrLf & "�`�F�b�N1: " & varChecked1Id & _
            vbCrLf & "�`�F�b�N2: " & varChecked2Id & _
            "", vbExclamation, "�ʒm"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 5 Then
        If varChecked1L1 <> varChecked2L1 Or varChecked1L2 <> varChecked2L2 Or varChecked1L3 <> varChecked2L3 Or varChecked1L4 <> varChecked2L4 Then
            MsgBox "�������̖����ԍ��ȊO�̊K�w�ԍ�����v���܂���iCHK)�B" & vbCrLf & _
            vbCrLf & "�`�F�b�N1: " & varChecked1Id & _
            vbCrLf & "�`�F�b�N2: " & varChecked2Id & _
            "", vbExclamation, "�ʒm"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 4 Then
        If varChecked1L1 <> varChecked2L1 Or varChecked1L2 <> varChecked2L2 Or varChecked1L3 <> varChecked2L3 Then
            MsgBox "�������̖����ԍ��ȊO�̊K�w�ԍ�����v���܂���iCHK)�B" & vbCrLf & _
            vbCrLf & "�`�F�b�N1: " & varChecked1Id & _
            vbCrLf & "�`�F�b�N2: " & varChecked2Id & _
            "", vbExclamation, "�ʒm"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 3 Then
        If varChecked1L1 <> varChecked2L1 Or varChecked1L2 <> varChecked2L2 Then
            MsgBox "�������̖����ԍ��ȊO�̊K�w�ԍ�����v���܂���iCHK)�B" & vbCrLf & _
            vbCrLf & "�`�F�b�N1: " & varChecked1Id & _
            vbCrLf & "�`�F�b�N2: " & varChecked2Id & _
            "", vbExclamation, "�ʒm"
            Exit Sub
        End If
    ElseIf intRowLevel1 = 2 Then
        If varChecked1L1 <> varChecked2L1 Then
            MsgBox "�������̖����ԍ��ȊO�̊K�w�ԍ�����v���܂���iCHK)�B" & vbCrLf & _
            vbCrLf & "�`�F�b�N1: " & varChecked1Id & _
            vbCrLf & "�`�F�b�N2: " & varChecked2Id & _
            "", vbExclamation, "�ʒm"
            Exit Sub
        End If
    End If
    
    ' �l�̌��������{
    For r = lngStartRow To lngEndRow
        
        ' ���݂̃C���f�b�N�X���擾
        i = r - lngStartRow + 1
        
        ' ���������l���Z�b�g
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
    
    ' �l�𔽉f
    rngHierarchy.Value = varHierarchyArray
    
End Sub


' �� �w��̊K�w�z���ΏۂɁA�w��C���f�b�N�X�ɂ���f�[�^�̎q�K�w�ɂ�����C���f�b�N�X�̃R���N�V�������擾
Private Function GetTargetChildIdxs(varHierarchyArray As Variant, _
                                        lngTargetIdx As Long) As Collection
    
    ' �ϐ���`
    Dim colResultIdxs As New Collection
    Dim intTargetLevel As Integer, blnTargetTask As Boolean
    Dim varTargetL1 As Variant, varTargetL2 As Variant, varTargetL3 As Variant, varTargetL4 As Variant, varTargetL5 As Variant, varTargetTask As Variant
    ' �ꎞ�ϐ���`
    Dim i As Long
    
    ' �K�[�h�����i���͂��ꂽ�C���f�b�N�X��0�ȉ��̏ꍇ�͏I���j
    If lngTargetIdx <= 0 Then
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̗񐔂�6�łȂ��ꍇ�͏I���j
    If UBound(varHierarchyArray, 2) <> 6 Then
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' �K�[�h�����i���͂��ꂽ�K�w�z��̍s�����z�����C���f�b�N�X���w�肳�ꂽ�ꍇ�͏I���j
    If UBound(varHierarchyArray, 1) < lngTargetIdx Then
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' �w��C���f�b�N�X�̒l���擾
    varTargetL1 = varHierarchyArray(lngTargetIdx, 1)
    varTargetL2 = varHierarchyArray(lngTargetIdx, 2)
    varTargetL3 = varHierarchyArray(lngTargetIdx, 3)
    varTargetL4 = varHierarchyArray(lngTargetIdx, 4)
    varTargetL5 = varHierarchyArray(lngTargetIdx, 5)
    varTargetTask = varHierarchyArray(lngTargetIdx, 6)
    ' �^�X�N��Ԃ̎擾
    If IsEmpty(varTargetTask) Then
        blnTargetTask = False
    Else
        blnTargetTask = True
    End If
    ' ���x���̎擾
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
        ' # �K�w�ɖ�肪����ꍇ #
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' �K�[�h�����i�^�X�N�̏ꍇ�͏I���j
    If blnTargetTask = True Then
        ' # �^�X�N�ɂ͎q�K�w���Ȃ����� #
        Set GetTargetChildIdxs = colResultIdxs
        Exit Function
    End If
    
    ' �Y������C���f�b�N�X�����W
    For i = 1 To UBound(varHierarchyArray, 1)
        If intTargetLevel = 5 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                varTargetL4 = varHierarchyArray(i, 4) And _
                varTargetL5 = varHierarchyArray(i, 5) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L5�̏ꍇ�AL5�̃^�X�N�Ȃ�Βǉ� #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 4 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                varTargetL4 = varHierarchyArray(i, 4) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L4�̏ꍇ�AL4�̃^�X�N�Ȃ�Βǉ� #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 4 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                varTargetL4 = varHierarchyArray(i, 4) And _
                IsNumeric(varHierarchyArray(i, 5)) And Not IsNull(varHierarchyArray(i, 5)) And Not IsEmpty(varHierarchyArray(i, 5)) And _
                IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L4�̏ꍇ�AL4�̎q�ł���L5�Ȃ�Βǉ� #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 3 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L3�̏ꍇ�AL3�̃^�X�N�Ȃ�Βǉ� #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 3 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                varTargetL3 = varHierarchyArray(i, 3) And _
                IsNumeric(varHierarchyArray(i, 4)) And Not IsNull(varHierarchyArray(i, 4)) And Not IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L3�̏ꍇ�AL3�̎q�ł���L4�Ȃ�Βǉ� #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 2 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                IsEmpty(varHierarchyArray(i, 3)) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L2�̏ꍇ�AL2�̃^�X�N�Ȃ�Βǉ� #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 2 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                varTargetL2 = varHierarchyArray(i, 2) And _
                IsNumeric(varHierarchyArray(i, 3)) And Not IsNull(varHierarchyArray(i, 3)) And Not IsEmpty(varHierarchyArray(i, 3)) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L2�̏ꍇ�AL2�̎q�ł���L3�Ȃ�Βǉ� #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 1 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                IsEmpty(varHierarchyArray(i, 2)) And _
                IsEmpty(varHierarchyArray(i, 3)) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsNumeric(varHierarchyArray(i, 6)) And Not IsNull(varHierarchyArray(i, 6)) And Not IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L1�̏ꍇ�AL1�̃^�X�N�Ȃ�Βǉ� #
            colResultIdxs.Add i, CStr(i)
        ElseIf intTargetLevel = 1 And _
                varTargetL1 = varHierarchyArray(i, 1) And _
                IsNumeric(varHierarchyArray(i, 2)) And Not IsNull(varHierarchyArray(i, 2)) And Not IsEmpty(varHierarchyArray(i, 2)) And _
                IsEmpty(varHierarchyArray(i, 3)) And _
                IsEmpty(varHierarchyArray(i, 4)) And _
                IsEmpty(varHierarchyArray(i, 5)) And _
                IsEmpty(varHierarchyArray(i, 6)) Then
            ' # L1�̏ꍇ�AL1�̎q�ł���L2�Ȃ�Βǉ� #
            colResultIdxs.Add i, CStr(i)
        End If
    Next i
    
    Set GetTargetChildIdxs = colResultIdxs
End Function



' �� �`�F�b�N�����s���폜����
Public Sub ExecRemoveCheckedRows(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colCheckedRows As Collection
    Dim varHierarchyArray As Variant
    Dim varChildExistArray As Variant
    Dim rngChk As Range
    Dim varChkArray As Variant
    Dim varIdArray As Variant
    Dim colRemoveRows As New Collection
    Dim rngRemoveTarget As Range
    ' �ꎞ�ϐ���`
    Dim tmpVarCheckedItem As Variant
    Dim tmpVarChildIdx As Variant
    Dim i As Long
    Dim tmpColChilds As Collection
    Dim tmpVar As Variant
    Dim answer As VbMsgBoxResult

    ' �J�n�s�ƏI���s���擾
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub

    ' �`�F�b�N����Ă���s�ԍ����擾
    Set colCheckedRows = GetCheckedChkMultpleRows(ws)
    
    ' ���炩���߃`�F�b�N�Ώ۔͈͗�̃f�[�^���擾
    varHierarchyArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK)).Value
    ' ���炩����WBS�q�L�������̃f�[�^���擾
    varChildExistArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_CE), ws.Cells(lngEndRow, cfg.COL_FLG_CE)).Value
    ' ���炩���߃`�F�b�N��̃f�[�^���擾
    Set rngChk = ws.Range(ws.Cells(lngStartRow, cfg.COL_CHK), ws.Cells(lngEndRow, cfg.COL_CHK))
    varChkArray = rngChk.Value
    ' ���炩����WBS-ID��̃f�[�^���擾
    varIdArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_WBS_ID), ws.Cells(lngEndRow, cfg.COL_WBS_ID)).Value
    
    ' �`�F�b�N���ꂽ�s���Ƃɍ폜�\���`�F�b�N�����{
    For Each tmpVarCheckedItem In colCheckedRows
        ' ���݂̃C���f�b�N�X���擾
        i = tmpVarCheckedItem - lngStartRow + 1
        ' �q�����邩�ǂ���
        If varChildExistArray(i, 1) Then
            ' # �q�����݂���ꍇ #
            Set tmpColChilds = GetTargetChildIdxs(varHierarchyArray, i)
            For Each tmpVarChildIdx In tmpColChilds
                tmpVar = varChildExistArray(tmpVarChildIdx, 1)
                If tmpVar = True Then
                    ' # �������݂���ꍇ #
                    MsgBox "���K�w�����݂��邽�ߍ폜�ł��܂���B" & vbCrLf & _
                    "", vbExclamation, "�ʒm"
                    Exit Sub
                Else
                    ' # �������݂��Ȃ��ꍇ #
                    On Error Resume Next
                    colRemoveRows.Add (tmpVarChildIdx + lngStartRow - 1), CStr(tmpVarChildIdx + lngStartRow - 1)
                    On Error GoTo 0
                End If
            Next tmpVarChildIdx
            On Error Resume Next
            colRemoveRows.Add tmpVarCheckedItem, CStr(tmpVarCheckedItem)
            On Error GoTo 0
        Else
            ' # �q�����݂��Ȃ��ꍇ #
            On Error Resume Next
            colRemoveRows.Add tmpVarCheckedItem, CStr(tmpVarCheckedItem)
            On Error GoTo 0
        End If
    Next tmpVarCheckedItem
    
    ' �`�F�b�N����X�V����
    For Each tmpVar In colRemoveRows
        varChkArray(tmpVar - lngStartRow + 1, 1) = cfg.CHK_MARK_T
    Next tmpVar
    rngChk.Value = varChkArray
    
    ' �ꎞ�I�ɕ`����ĊJ
    If Application.ScreenUpdating = False And Application.EnableEvents = False Then
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.Wait (Now + TimeValue("00:00:01"))
        Application.ScreenUpdating = False
        Application.EnableEvents = False
    End If
    
    ' �m�F�̏�A�폜�����s
    answer = MsgBox("�{���ɍ폜���Ă��悢�ł����H", vbOKCancel + vbQuestion, "�m�F")
    If answer = vbOK Then
        ' �폜�Ώ۔͈͂�p��
        For Each tmpVar In colRemoveRows
            If rngRemoveTarget Is Nothing Then
                Set rngRemoveTarget = Rows(tmpVar)
            Else
                Set rngRemoveTarget = Union(rngRemoveTarget, Rows(tmpVar))
            End If
        Next tmpVar
        ' �ꊇ�폜�����s
        If Not rngRemoveTarget Is Nothing Then rngRemoveTarget.Delete
    End If

End Sub


' �� ��{�����𐔒l�ɕϊ�����
Public Sub ExecConvertBasicFormulasToValues(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' �ꎞ�ϐ���`
    Dim r As Long
    Dim tmpRange As Range
    Dim tmpVariant As Variant
    
    ' �J�n�s�ƏI���s���擾
    varRangeRows = FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �� �S�s�ɃA�N�Z�X���K�v�ȃR�X�g�̍��������𐔒l�ɕϊ�
    ' WBS_CNT�̎����l
    Set tmpRange = ws.Range(cfg.COL_WBS_CNT_LABEL & lngStartRow & ":" & cfg.COL_WBS_CNT_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.Value = tmpVariant
    
    ' FLG_PE�̎����l
    Set tmpRange = ws.Range(cfg.COL_FLG_PE_LABEL & lngStartRow & ":" & cfg.COL_FLG_PE_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.Value = tmpVariant
    
    ' FLG_CE�̎����l
    Set tmpRange = ws.Range(cfg.COL_FLG_CE_LABEL & lngStartRow & ":" & cfg.COL_FLG_CE_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.Value = tmpVariant

End Sub


' �� �W�v�����𐔒l�ɕϊ�����
Public Sub ExecConvertAggregateFormulasToValues(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' �ꎞ�ϐ���`
    Dim r As Long
    Dim tmpRange As Range
    Dim tmpVariant As Variant
    
    ' �J�n�s�ƏI���s���擾
    varRangeRows = FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
        
    ' �^�X�N�W�v���v�̎����l
    Set tmpRange = ws.Range(cfg.COL_TASK_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COUNT_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
    
    ' �^�X�N�W�v�����̎����l
    Set tmpRange = ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngStartRow & ":" & cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
    
    ' �H���i�����̎����l
    Set tmpRange = ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngStartRow & ":" & cfg.COL_EFFORT_PROG_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "0.0%"
    tmpRange.Value = tmpVariant
    
    ' ���ڏ������̎����l
    Set tmpRange = ws.Range(cfg.COL_TASK_PROG_LABEL & lngStartRow & ":" & cfg.COL_TASK_PROG_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "0.0%"
    tmpRange.Value = tmpVariant
    
    ' �\��H���̎����l
    Set tmpRange = ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngStartRow & ":" & cfg.COL_PLANNED_EFF_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
    
    ' ���юc�H���̎����l
    Set tmpRange = ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
    
    ' ���эύH���̎����l
    Set tmpRange = ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngStartRow & ":" & cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow)
    tmpVariant = tmpRange.Value
    tmpRange.NumberFormat = "General"
    tmpRange.Value = tmpVariant
           
    ' ����̃Z���̎���l�ɕϊ�
    ws.Range(cfg.COL_TASK_COUNT_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_TASK_COUNT_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_EFFORT_PROG_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_TASK_PROG_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_TASK_PROG_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_PLANNED_EFF_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & lngEndRow + 2).Value
    ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow + 2).Value = ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & lngEndRow + 2).Value

End Sub


' �� �J�X�^���t�H�[�}�b�g�֐��iWBS-IDX�p�j
Function CustomFormatWbsIdx(varB As Variant, _
                                varE As Variant, _
                                varF As Variant, _
                                varG As Variant, _
                                varH As Variant, _
                                varI As Variant, _
                                varJ As Variant) As String
    
    ' �ϐ���`
    Dim strResult As String
    Dim varValues As Variant
    ' �ꎞ�ϐ���`
    Dim parts(0 To 5) As String
    Dim i As Integer

    ' ����varB��"E"�Ȃ�"ERROR"��Ԃ�
    If varB = "E" Then
        CustomFormat = "ERROR"
        Exit Function
    End If

    ' ����varE����Ȃ�Œ蕶�����Ԃ�
    If varE = "" Then
        CustomFormatWbsIdx = "XXX.XXX.XXX.XXX.XXX.XXX"
        Exit Function
    End If

    ' �e�l��z��ɂ܂Ƃ߂�
    varValues = Array(varE, varF, varG, varH, varI, varJ)

    ' �e�v�f�����[�v���ď���
    For i = 0 To 5
        If varValues(i) = "" Then
            parts(i) = "---"
        Else
            parts(i) = Format(varValues(i), "000")
        End If
    Next i

    ' �������Č��ʂ��쐬
    strResult = parts(0) & "." & parts(1) & "." & parts(2) & "." & parts(3) & "." & parts(4) & "." & parts(5)

    CustomFormatWbsIdx = strResult
End Function


' �� �J�X�^���t�H�[�}�b�g�֐��iWBS-ID�p�j
Function CustomFormatWbsId(varB As Variant, _
                            varE As Variant, _
                            varF As Variant, _
                            varG As Variant, _
                            varH As Variant, _
                            varI As Variant, _
                            varJ As Variant) As String
    
    ' �ϐ���`
    Dim strResult As String

    ' ����varB��"E"�Ȃ�"ERROR"��Ԃ�
    If varB = "E" Then
        CustomFormatWbsId = "ERROR"
        Exit Function
    End If

    ' ����varE����Ȃ�󕶎���Ԃ�
    If varE = "" Then
        CustomFormatWbsId = ""
        Exit Function
    End If

    ' �A������
    strResult = varE

    If varF <> "" Then strResult = strResult & "." & varF
    If varG <> "" Then strResult = strResult & "." & varG
    If varH <> "" Then strResult = strResult & "." & varH
    If varI <> "" Then strResult = strResult & "." & varI
    If varJ <> "" Then strResult = strResult & ".T" & varJ

    CustomFormatWbsId = strResult
End Function


' �� �J�X�^���֐��iLEVEL�j
Function CustomFuncGetLevel(varE As Variant, _
                                varF As Variant, _
                                varG As Variant, _
                                varH As Variant, _
                                varI As Variant) As Integer
    
    ' �f�t�H���g��0
    CustomFuncGetLevel = 0
    
    ' ���ԂɃ`�F�b�N���Ă���
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


' �� ����Ɏg�p�����ɐ������܂Ƃ߂ăZ�b�g
Public Sub SetFormulaToControlColumn(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim varFormulas() As Variant
    ' �ꎞ�ϐ����`
    Dim i As Long, j As Long
    Dim tmpLngRow As Long

    ' �J�n�s�ƏI���s�ɒl���Z�b�g
    varRangeRows = wbslib.FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    lngEndRow = varRangeRows(1)

    ' �J�n�s�ƏI���s��������Ȃ���ΏI��
    If lngStartRow = 0 Or lngEndRow = 0 Or lngStartRow >= lngEndRow Then Exit Sub
    
    ' �������Z�b�g����f�[�^��p��
    ReDim varFormulas(1 To lngEndRow - lngStartRow + 1, 1 To cfg.COL_WBS_ID - cfg.COL_WBS_IDX + 1)

    ' �������Z�b�g
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

    ' �ꊇ�őΏ۔͈͂ɑ΂��������s��
    With ws.Range(cfg.COL_WBS_IDX_LABEL & lngStartRow & ":" & cfg.COL_WBS_ID_LABEL & lngEndRow)
        ' �������ꊇ�Őݒ�
        .NumberFormat = "General"
        ' �����Z�b�g
        .Formula = varFormulas
    End With

End Sub

