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
Public Sub ExecCheckWbsErrors(ws As Worksheet)

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
                ws.Cells(r, cfg.COL_ERR).value = "E"
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
    If intErrorCount > 0 Then
        MsgBox intErrorCount & " ���ُ̈�����o���܂����B", vbExclamation, "�G���[�`�F�b�N"
    End If
    
End Sub


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
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
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
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_PLANNED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
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


' �� ���эύH�����W�v���鎮���Z�b�g
Public Sub SetFormulaForActualCompletedEffort(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
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
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
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
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_ACTUAL_COMPLETED_EFF_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
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


' �� ���юc�H�����W�v���鎮���Z�b�g
Public Sub SetFormulaForActualRemainingEffort(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
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
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
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
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_ACTUAL_REMAINING_EFF_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
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


' �� �^�X�N�i�������W�v���鎮���Z�b�g
Public Sub SetFormulaForTaskProgressRate(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpStrFormula As String
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
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
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
            ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "0.0%"
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
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).NumberFormat = "General"
                ws.Range(cfg.COL_TASK_PROG_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
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


' �� �H���i�������W�v���鎮���Z�b�g
Public Sub SetFormulaForEffortProgressRate(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
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
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
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
            ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_EFFORT_PROG_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
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


' �� �^�X�N���v�������W�v���鎮���Z�b�g
Public Sub SetFormulaForTaskCount(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
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
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
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
            ws.Range(cfg.COL_TASK_COUNT_LABEL & r).value = 1
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
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
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


' �� �^�X�N�����������W�v���鎮���Z�b�g
Public Sub SetFormulaForTaskCompCount(ws As Worksheet)

    ' �ϐ���`
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
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
    
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
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
            ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
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
                ws.Range(cfg.COL_TASK_COMP_COUNT_LABEL & r).Formula = tmpStrFormula
            End If
        End If
    Next r
    
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
    varData = ws.Range(ws.Cells(lngStartRow, cfg.COL_CHK), ws.Cells(lngEndRow, cfg.COL_CHK)).value
    
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
    
    ' �J�n�s�ƏI���s���擾
    varRangeRows = FindDataRangeRows(ws)
    lngStartRow = varRangeRows(0)
    
    ' �s��ǉ�
    lngSelectedRow = GetCheckedOptSingleRow(ws)
    If lngSelectedRow <> 0 Then
        ' �s��ǉ�
        ws.Rows(lngSelectedRow + 1).Insert Shift:=xlDown
    Else
        MsgBox "�I�����Ă��������iOPT)�B", vbExclamation, "�ʒm"
    End If

End Sub


' �� �I���s�̍ŏI���x��ID���C���N�������g
Public Sub ExecIncrementSelectedLastLevel(ws As Worksheet)

    ' �ϐ���`
    Dim lngSelectedRow As Long, intSelectedRowLevel As Integer, blnSelectedRowIsTask As Boolean
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colTargetIdx As New Collection
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long
    Dim tmpRngTarget As Range
    Dim tmpVarTargetArray As Variant
    Dim tmpVarLevelArray As Variant
    Dim tmpVarTaskArray As Variant
    Dim tmpLngSelectedRowL1 As Long, tmpLngSelectedRowL2 As Long, tmpLngSelectedRowL3 As Long, tmpLngSelectedRowL4 As Long, tmpLngSelectedRowL5 As Long, tmpLngSelectedRowTask As Long
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
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    tmpVarTargetArray = tmpRngTarget.value
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' �I�������s�̃��x�����擾
    intSelectedRowLevel = tmpVarLevelArray(lngSelectedRow - lngStartRow + 1, 1)
    ' �I�������s���^�X�N���ǂ����擾
    blnSelectedRowIsTask = tmpVarTaskArray(lngSelectedRow - lngStartRow + 1, 1)
    
    ' �I�������s�̃f�[�^���擾
    tmpLngSelectedRowL1 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 1)
    tmpLngSelectedRowL2 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 2)
    tmpLngSelectedRowL3 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 3)
    tmpLngSelectedRowL4 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 4)
    tmpLngSelectedRowL5 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 5)
    tmpLngSelectedRowTask = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 6)
    
    ' �X�V�Ώ۔͈͗�̃f�[�^���X�V
    If blnSelectedRowIsTask = True Then
        ' # �I���s���^�X�N�̏ꍇ #
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X�̂ݒl���X�V����
        For Each tmpVarIdx In colTargetIdx
            tmpVarTargetArray(tmpVarIdx, 6) = tmpVarTargetArray(tmpVarIdx, 6) + 1
        Next tmpVarIdx
    Else
        ' # �I���s���^�X�N�łȂ��ꍇ #
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X�̂ݒl���X�V����
        For Each tmpVarIdx In colTargetIdx
            tmpVarTargetArray(tmpVarIdx, intSelectedRowLevel) = tmpVarTargetArray(tmpVarIdx, intSelectedRowLevel) + 1
        Next tmpVarIdx
    End If
    
    ' �f�[�^�̍X�V���ʂ𔽉f
    tmpRngTarget.value = tmpVarTargetArray

End Sub


' �� �I���s�̍ŏI���x��ID���f�N�������g
Public Sub ExecDecrementSelectedLastLevel(ws As Worksheet)

    ' �ϐ���`
    Dim lngSelectedRow As Long, intSelectedRowLevel As Integer, blnSelectedRowIsTask As Boolean, lngSelectedRowLastValue As Long
    Dim varRangeRows As Variant, lngStartRow As Long, lngEndRow As Long
    Dim colTargetIdx As New Collection
    Dim lngFirstMissingFoundValue As Long
    ' �ꎞ�ϐ���`
    Dim r As Long, i As Long, v As Long
    Dim tmpRngTarget As Range
    Dim tmpVarTargetArray As Variant
    Dim tmpVarLevelArray As Variant
    Dim tmpVarTaskArray As Variant
    Dim tmpLngSelectedRowL1 As Long, tmpLngSelectedRowL2 As Long, tmpLngSelectedRowL3 As Long, tmpLngSelectedRowL4 As Long, tmpLngSelectedRowL5 As Long, tmpLngSelectedRowTask As Long
    Dim tmpVarIdx As Variant
    Dim tmpColTargetValue As New Collection
    Dim tmpVal As Variant, tmpBlnExist As Boolean
    
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
    Set tmpRngTarget = ws.Range(ws.Cells(lngStartRow, cfg.COL_L1), ws.Cells(lngEndRow, cfg.COL_TASK))
    tmpVarTargetArray = tmpRngTarget.value
    ' ���炩����WBS���x����̃f�[�^���擾
    tmpVarLevelArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_LEVEL), ws.Cells(lngEndRow, cfg.COL_LEVEL)).value
    ' ���炩����WBS�^�X�N�����̃f�[�^���擾
    tmpVarTaskArray = ws.Range(ws.Cells(lngStartRow, cfg.COL_FLG_T), ws.Cells(lngEndRow, cfg.COL_FLG_T)).value
    
    ' �I�������s�̃��x�����擾
    intSelectedRowLevel = tmpVarLevelArray(lngSelectedRow - lngStartRow + 1, 1)
    ' �I�������s���^�X�N���ǂ����擾
    blnSelectedRowIsTask = tmpVarTaskArray(lngSelectedRow - lngStartRow + 1, 1)
    ' �I�������s�̖����̒l���擾
    If blnSelectedRowIsTask Then
        lngSelectedRowLastValue = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 6)
    Else
        lngSelectedRowLastValue = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, intSelectedRowLevel)
    End If
    
    ' �I�������s�̃f�[�^���擾
    tmpLngSelectedRowL1 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 1)
    tmpLngSelectedRowL2 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 2)
    tmpLngSelectedRowL3 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 3)
    tmpLngSelectedRowL4 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 4)
    tmpLngSelectedRowL5 = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 5)
    tmpLngSelectedRowTask = tmpVarTargetArray(lngSelectedRow - lngStartRow + 1, 6)
    
    ' �X�V�Ώ۔͈͗�̃f�[�^���X�V
    If blnSelectedRowIsTask = True Then
        ' # �I���s���^�X�N�̏ꍇ #
        ' �ΏۂƂȂ�l���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �l�R���N�V����������ŏ��̑��݂��Ȃ��l���擾
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
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X�̂ݒl���X�V����
        For Each tmpVarIdx In colTargetIdx
            tmpVarTargetArray(tmpVarIdx, 6) = tmpVarTargetArray(tmpVarIdx, 6) - 1
        Next tmpVarIdx
    Else
        ' # �I���s���^�X�N�łȂ��ꍇ #
        ' �ΏۂƂȂ�l���R���N�V�����Ɋi�[
        For r = lngStartRow To lngEndRow
            ' ���݂̃C���f�b�N�X���擾
            i = r - lngStartRow + 1
            ' �Ώۍs�����肵�ăR���N�V�����Ɋi�[
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
        ' �l�R���N�V����������ŏ��̑��݂��Ȃ��l���擾
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
        ' �ΏۂƂȂ�f�[�^�C���f�b�N�X�̂ݒl���X�V����
        For Each tmpVarIdx In colTargetIdx
            tmpVarTargetArray(tmpVarIdx, intSelectedRowLevel) = tmpVarTargetArray(tmpVarIdx, intSelectedRowLevel) - 1
        Next tmpVarIdx
    End If
    
    ' �f�[�^�̍X�V���ʂ𔽉f
    tmpRngTarget.value = tmpVarTargetArray

End Sub



