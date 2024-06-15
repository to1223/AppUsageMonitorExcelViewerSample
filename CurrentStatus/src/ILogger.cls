VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' Logger interface

Option Explicit


' ファイル等のリソース開放
Public Sub Dispose()
End Sub


Public Sub Trace(message As String)
End Sub


' 名前を Debug にすると Syntax Error になる
Public Sub DebugLog(message As String)
End Sub


Public Sub Info(message As String)
End Sub


Public Sub Error(message As String)
End Sub


Public Sub Warn(message As String)
End Sub


Public Sub Fatal(message As String)
End Sub
