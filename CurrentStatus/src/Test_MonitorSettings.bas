Attribute VB_Name = "Test_MonitorSettings"
''' MonitorSettings クラスの単体テスト

Option Explicit

Public Sub Test_Init()

    ' Setup
    Const SettingsFolderPath As String = "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\settings"
    Const MonitorOutputFolderPath As String = "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\monitor_out"
    
    Dim monitor_settings As MonitorSettings
    Set monitor_settings = New MonitorSettings
    monitor_settings.Init SettingsFolderPath, MonitorOutputFolderPath
    
    ' Test
    Debug.Print "IntervalSeconds: " & monitor_settings.IntervalSeconds
    Debug.Assert monitor_settings.IntervalSeconds = 60
    
    Dim i As Integer
    For i = LBound(monitor_settings.TargetPCs) To UBound(monitor_settings.TargetPCs)
        Debug.Print "TargetPCs(" & i & "): " & monitor_settings.TargetPCs(i)
    Next
    
    Debug.Print "MonitorOutputFolder: " & monitor_settings.MonitorOutputFolderPath
    Debug.Assert monitor_settings.MonitorOutputFolderPath = MonitorOutputFolderPath

    Debug.Print "OK: Test_MonitorSettings.Test_Init"

End Sub
