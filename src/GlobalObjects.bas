Attribute VB_Name = "GlobalObjects"
''' グローバルな各種オブジェクト変数を定義
'
' ブックオープン時に実体を生成して設定する。
' クラスの中からは使わずに、例外をクラスを使う側で補足して、そこからLoggerを使う方がいいのかも

Option Explicit

' AppUsageMonitor の設定ファイルの名前
Public Const IntervalSecondsFileName = "interval_seconds"
Public Const TargetPCsFileName = "target_pc"

' Logger
Public Logger As ILogger

' MonitorSettings
Public MonitorSettingsObject As MonitorSettings

