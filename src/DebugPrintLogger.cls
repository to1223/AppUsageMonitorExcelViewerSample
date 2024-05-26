VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugPrintLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' Logger that uses Debug.Print

Option Explicit

Implements ILogger


Private Function BuildRecordLine(log_level As String, message As String) As String

    BuildRecordLine = Now & " " & log_level & " " & message

End Function


Public Sub ILogger_Dispose()
    
    ' Nothing

End Sub


Public Sub ILogger_Trace(message As String)

    Debug.Print BuildRecordLine("TRACE", message)
    
End Sub


Public Sub ILogger_DebugLog(message As String)

    Debug.Print BuildRecordLine("DEBUG", message)
    
End Sub


Public Sub ILogger_Info(message As String)

    Debug.Print BuildRecordLine("INFO", message)
    
End Sub


Public Sub ILogger_Error(message As String)

    Debug.Print BuildRecordLine("ERROR", message)
    
End Sub


Public Sub ILogger_Warn(message As String)

    Debug.Print BuildRecordLine("WARN", message)
    
End Sub


Public Sub ILogger_Fatal(message As String)

    Debug.Print BuildRecordLine("FATAL", message)
    
End Sub

