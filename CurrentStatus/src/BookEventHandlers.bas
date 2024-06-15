Attribute VB_Name = "BookEventHandlers"
Option Explicit


''' <summary>
''' ブック起動時に実行する処理。
''' </summary>
Public Sub OnBookOpen()

    ' Logger の設定
    Set Logger = New DebugPrintLogger
    
    ' モニター設定の読み取り
    Set MonitorSettingsObject = New MonitorSettings
    MonitorSettingsObject.Init MonitorSettingsFolderPath, MonitorOutputFolderPath

End Sub
