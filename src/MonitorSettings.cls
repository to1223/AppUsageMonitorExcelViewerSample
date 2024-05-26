VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MonitorSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' モニターアプリの設定を保持するクラス

Option Explicit


' 監視時間間隔（秒）
Private intervalSeconds_ As Integer
' 監視対象PC名の配列
Private targetPCs_ As Variant
' 監視記録の出力先フォルダ
Private monitorOutputFolderPath_ As String


''' <summary>
''' 設定ファイルを読み込んでインスタンスを初期化する
''' </summary>
''' <remarks>
''' 引数は絶対パスでの指定のみに対応。
''' </remarks>
Public Sub Init(settings_folder_path As String, monitor_output_folder_path As String)
    
'    Const adModeRead As Integer = 1
    Const adReadAll As Integer = -1
    Const adReadLine As Integer = -2

On Error GoTo ErrorHandling
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' TestStream を作成して操作
'    Dim ts As Object
    
    ' Read interval_seconds
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile fso.BuildPath(settings_folder_path, "interval_seconds")
        intervalSeconds_ = Val(.ReadText(adReadLine))
        .Close
    End With

    ' Read target_pc
    With CreateObject("ADODB.Stream")
        .Charset = "UTF-8"
        .Open
        .LoadFromFile fso.BuildPath(settings_folder_path, "target_pc")
        targetPCs_ = Split(.ReadText(adReadAll), vbCrLf)
        .Close
    End With
    ' TODO: eliminate blank lines
     
    ' Read monitor_output_folder
    ' TODO: パスに日本語が含まれる場合、文字コードの問題が生じるか確認し、その対策を検討
    monitorOutputFolderPath_ = monitor_output_folder_path
    
    GoTo Closing

ErrorHandling:
    ' Logging "Fatal 設定ファイル読み取り時にエラーが発生しました"
    Debug.Print "Fatal 設定ファイル読み取り時にエラーが発生しました"
    
Closing:
    Set fso = Nothing

End Sub


Public Property Get IntervalSeconds() As Integer

    IntervalSeconds = intervalSeconds_
    
End Property


' TODO: コピーを返すようにした方が安全であろう
Public Property Get TargetPCs() As Variant

    TargetPCs = targetPCs_

End Property


Public Property Get MonitorOutputFolderPath() As String
    
    ' GlobalSettings とぶつかってエラーになるかも？
    MonitorOutputFolderPath = monitorOutputFolderPath_

End Property
