Attribute VB_Name = "Test_ViewController"
Option Explicit


Public Sub Test_GetCellNameFromPCName()

    Debug.Assert GetCellNameFromPCName("SURFACE-PRO-9") = "SURFACE_PRO_9"
    
    Debug.Print "OK: Test_GetCellNameFromPCName"

End Sub


Public Sub Test_UpdateLastUpdateCell()

    UpdateLastUpdateCell Now()

End Sub


Public Sub Test_ApplyFormatFromLegend()

    Dim status As AppUsageStatusType
'    status = AppUsageStatusActive
'    status = AppUsageStatusLogOff
'    status = AppUsageStatusInactive
    status = AppUsageStatusNotTarget

    ApplyFormatFromLegend _
        Range(GetCellNameFromPCName("SURFACE-PRO-9")), _
        Range(GetLegendCellNameFromStatus(status))

End Sub


Public Sub Test_UpdateAsIs()

    Dim as_is As Date
    
    Dim monitor_settings As MonitorSettings
    Set monitor_settings = New MonitorSettings
    
    Dim status_calculator As AppUsageStatusCalculator
    Set status_calculator = New AppUsageStatusCalculator

    Dim settings_folder_path As String
    Dim monitor_output_folder_path As String

    ' èåè1: LogOFf
    as_is = CDate("2024-05-24 00:20:00")
    settings_folder_path = "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\settings"
    monitor_output_folder_path = "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\monitor_out\LogOff_MonitorInactive_2024-05-24_00-20-00"
    
    monitor_settings.Init settings_folder_path, monitor_output_folder_path
    
    status_calculator.Init monitor_output_folder_path

    UpdateAsIs _
        as_is, monitor_settings.TargetPCs, monitor_settings.IntervalSeconds, _
        3, 30 * 60, status_calculator

End Sub
