VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InitWBS 
   Caption         =   "WBSシート作成"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3990
   OleObjectBlob   =   "InitWBS.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "InitWBS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


' 新規作成ボタンを押した時の処理
Private Sub CommandButton1_Click()

    ' 変数定義
    Dim ws As Worksheet
    
    ' ワークシートを新規作成
    Set ws = Worksheets.Add
    ws.Name = "WBS-" & InitWBS.TextBox1.Value
    
    ' ワークシートを初期化
    wbsui.InitSheet ws
    
End Sub

