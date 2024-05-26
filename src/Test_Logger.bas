Attribute VB_Name = "Test_Logger"
Option Explicit

Public Sub Test_Debug()

    Dim logger_object As ILogger
    Set logger_object = New DebugPrintLogger
    logger_object.DebugLog "test"

    Debug.Print "OK: Test_Debug"
    
End Sub
