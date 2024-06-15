Attribute VB_Name = "UserSettings"
Option Explicit

''' ** 設定を変更したら、一度ブックを開きなおしてください。
' 相対パスの扱いが面倒なのと、想定する使い方から、絶対パスで指定する

' AppUsageMonitor の設定フォルダのパス
Public Const MonitorSettingsFolderPath = "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\settings"

' AppUsageMonitor の出力先フォルダのパス
Public Const MonitorOutputFolderPath = "C:\Users\Takuya\Projects\AppUsageMonitorExcelViewerSample\TestData\monitor_out\LogOff_MonitorInactive_2024-05-24_00-20-00"

