Attribute VB_Name = "Test_MonitorRecord"
''' MonitorRecord クラスの単体テスト

Option Explicit


Public Sub Test_Parse()

    Dim record As MonitorRecord
    Set record = New MonitorRecord
    record.Parse "2024-05-03 11:11:30 Takuya Unlocked Up"

    Debug.Assert record.DateTime = CDate("2024-05-03 11:11:30")
    Debug.Assert record.IsDisplayLocked = False
    Debug.Assert record.IsProcessActive = True
    
    Debug.Print "OK: Test_Parse"

End Sub
