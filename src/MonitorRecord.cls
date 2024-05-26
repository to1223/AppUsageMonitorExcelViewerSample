VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MonitorRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' 監視記録の1行の情報を保持するクラス

Option Explicit


Private dateTime_ As Date
Private isDisplayLocked_ As Boolean
Private isProcessActive_ As Boolean


Public Property Get DateTime() As Date
    
    DateTime = dateTime_

End Property


Public Property Get IsDisplayLocked() As Boolean

    IsDisplayLocked = isDisplayLocked_
    
End Property


Public Property Get IsProcessActive() As Boolean

    IsProcessActive = isProcessActive_

End Property


' テキスト行から読み取って初期化
' 行サンプル: 2024-05-03 11:11:30 Takuya Unlocked Up
Public Sub Parse(record_line As String)

    Const DateColumn As Integer = 0
    Const TimeColumn As Integer = 1
    Const UserNameColumn As Integer = 2
    Const DisplayStateColumn As Integer = 3
    Const ProcessStateColumn As Integer = 4

    Dim splitted As Variant
    splitted = Split(record_line, Delimiter:=" ")
    
    ' 記録時間の取得
    dateTime_ = CDate(splitted(DateColumn) & " " & splitted(TimeColumn))
    
    ' 画面ロック状態の取得
    If splitted(DisplayStateColumn) = "Locked" Then
        isDisplayLocked_ = True
    Else
        isDisplayLocked_ = False
    End If
    
    ' プロセス起動状態の取得
    If splitted(ProcessStateColumn) = "Up" Then
        isProcessActive_ = True
    Else
        isProcessActive_ = False
    End If
    
End Sub
