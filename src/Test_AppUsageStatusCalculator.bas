Attribute VB_Name = "Test_AppUsageStatusCalculator"
Option Explicit


Public Sub Test_BuildLogFilePath()

    Debug.Print "OK: Test_AppUsageStatusCalculator.Test_BuildLogFilePath"

End Sub


Public Sub Test_IsActiveRecord()
    
    Debug.Print "OK: Test_AppUsageStatusCalculator.Test_IsActiveRecord"

End Sub


Public Sub Test_FindRecordWithinInterval()

    ' TODO: Write test

End Sub


Public Sub Test_CalcAppUsageStatus()

    ' ïœêîèÄîı
    Dim calculator As AppUsageStatusCalculator
    Set calculator = New AppUsageStatusCalculator

    Dim status As AppUsageStatusType
    
    ' èåè1
    calculator.Init "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\monitor_out\LogOff_MonitorInactive_2024-05-24_00-20-00"
    status = calculator.CalcAppUsageStatus( _
        "SURFACE-PRO-9", CDate("2024-05-24 00:20:00"), _
        60, 3, 30 * 60 _
    )
    Debug.Assert status = AppUsageStatusLogOff
    
    ' èåè2
    calculator.Init "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\monitor_out\Active_AllRecordIsActive_2024-05-24_00-10-00"
    status = calculator.CalcAppUsageStatus( _
        "SURFACE-PRO-9", CDate("2024-05-24 00:10:00"), _
        60, 3, 30 * 60 _
    )
    Debug.Assert status = AppUsageStatusActive
    
    ' èåè3
    calculator.Init "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\monitor_out\Active_ExistsGap_2024-05-24_00-10-00"
    status = calculator.CalcAppUsageStatus( _
        "SURFACE-PRO-9", CDate("2024-05-24 00:10:00"), _
        60, 3, 30 * 60 _
    )
    Debug.Assert status = AppUsageStatusActive
    
    ' èåè4
    calculator.Init "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\monitor_out\Inactive_AllRecordIsInactive_2024-05-24_00-10-00"
    status = calculator.CalcAppUsageStatus( _
        "SURFACE-PRO-9", CDate("2024-05-24 00:10:00"), _
        60, 3, 30 * 60 _
    )
    Debug.Assert status = AppUsageStatusInactive
        
        
    Debug.Print "OK: Test_AppUsageStatusCalculator.Test_CalcAppUsageStatus"

End Sub
