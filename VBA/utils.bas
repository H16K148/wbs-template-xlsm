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


' ■ ユーティリティ：コレクション
Public Function ExistsColKey(objCol As Collection, strKey As String) As Boolean
     
    ' 戻り値の初期値：False
    ExistsColKey = False
     
    ' 変数にCollection未設定の場合は処理終了
    If objCol Is Nothing Then Exit Function
     
    ' Collectionのメンバー数が「0」の場合は処理終了
    If objCol.Count = 0 Then Exit Function
    
    ' エラー設定の変更
    On Error Resume Next
     
    ' Itemメソッドを実行
    Call objCol.item(strKey)
         
    ' エラー値がない場合：キー検索はヒット（戻り値：True）
    If Err.Number = 0 Then ExistsColKey = True
    
    ' 処理の終了とともに、エラー設定が元に戻る
 
End Function


' ■ ユーティリティ：コレクション
Public Function ExistsColItem(objCol As Collection, varItem As Variant) As Boolean
    
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
Public Function ExistsColItemCount(objCol As Collection, varItem As Variant) As Long

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


' ■ 列番号を列文字列に変換する
Public Function ConvertColNumberToLetter(lngColNum As Long) As String

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

