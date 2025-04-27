Attribute VB_Name = "utils"
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


' �� ���[�e�B���e�B�F�R���N�V����
Public Function ExistsColKey(objCol As Collection, strKey As String) As Boolean
     
    ' �߂�l�̏����l�FFalse
    ExistsColKey = False
     
    ' �ϐ���Collection���ݒ�̏ꍇ�͏����I��
    If objCol Is Nothing Then Exit Function
     
    ' Collection�̃����o�[�����u0�v�̏ꍇ�͏����I��
    If objCol.Count = 0 Then Exit Function
    
    ' �G���[�ݒ�̕ύX
    On Error Resume Next
     
    ' Item���\�b�h�����s
    Call objCol.item(strKey)
         
    ' �G���[�l���Ȃ��ꍇ�F�L�[�����̓q�b�g�i�߂�l�FTrue�j
    If Err.Number = 0 Then ExistsColKey = True
    
    ' �����̏I���ƂƂ��ɁA�G���[�ݒ肪���ɖ߂�
 
End Function


' �� ���[�e�B���e�B�F�R���N�V����
Public Function ExistsColItem(objCol As Collection, varItem As Variant) As Boolean
    
    ' �ϐ���`
    Dim v As Variant
     
    ' �߂�l�̏����l�FFalse
    ExistsColItem = False
     
    ' �ϐ���Collection���ݒ�̏ꍇ�͏����I��
    If objCol Is Nothing Then Exit Function
     
    ' Collection�̃����o�[�����u0�v�̏ꍇ�͏����I��
    If objCol.Count = 0 Then Exit Function
     
    ' Collection�̊e�����o�[�Ɠˍ�
    For Each v In objCol
         
        ' �ˍ����ʂ���v�����ꍇ�F�߂�l�uTrue�v�Ƀ��[�v����
        If v = varItem Then ExistsColItem = True: Exit For
         
    Next
     
End Function


' �� ���[�e�B���e�B�F�R���N�V����
Public Function ExistsColItemCount(objCol As Collection, varItem As Variant) As Long

    ' �ϐ���`
    Dim v As Variant
    Dim lngMatchCount As Long ' ��v���������J�E���g����ϐ�

    ' �߂�l�̏�����
    ExistsColItemCount = 0

    ' �ϐ���Collection���ݒ�̏ꍇ�͏����I��
    If objCol Is Nothing Then Exit Function

    ' Collection�̃����o�[�����u0�v�̏ꍇ�͏����I��
    If objCol.Count = 0 Then Exit Function

    ' Collection�̊e�����o�[�Ɠˍ�
    For Each v In objCol
        ' �ˍ����ʂ���v�����ꍇ�F�J�E���g���C���N�������g
        If v = varItem Then
            lngMatchCount = lngMatchCount + 1
        End If
    Next

    ' ��v���������֐��̖߂�l�Ƃ��Đݒ�
    ExistsColItemCount = lngMatchCount

End Function


' �� ��ԍ���񕶎���ɕϊ�����
Public Function ConvertColNumberToLetter(lngColNum As Long) As String

    ' �ϐ���`
    Dim lngDiv As Long
    Dim lngModVal As Long
    Dim strColLetter As String

    ' ������
    lngDiv = lngColNum
    strColLetter = ""

    ' �v�Z
    Do While lngDiv > 0
        lngModVal = (lngDiv - 1) Mod 26
        strColLetter = Chr(65 + lngModVal) & strColLetter
        lngDiv = Int((lngDiv - lngModVal) / 26)
    Loop

    ConvertColNumberToLetter = strColLetter
End Function

