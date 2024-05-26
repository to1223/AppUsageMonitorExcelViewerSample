VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppUsageStatusCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' <summary>
'''　監視アプリの出力ファイルを読み取り、アプリの使用状況を判定するクラス。
''' </summary>
' インターフェースにして、アルゴリズムを差し替えられるようにしてもいいだろう。

Option Explicit

Private monitorOutputFolderPath_ As String

''' <summary>
''' インスタンスを初期化する
''' </summary>
' グローバル変数を参照するのも考えたが、テストをしやすくするため、引数で渡して初期化する
Public Sub Init(monitor_output_folder_path As String)

    monitorOutputFolderPath_ = monitor_output_folder_path

End Sub

''' <summary>
''' アプリの使用状況を判定して返す。
''' </summary>
Public Function CalcAppUsageStatus( _
    pc_name As String, as_is As Date, interval_seconds As Long, _
    interval_margin_seconds As Long, inspection_timespan_seconds As Long _
) As AppUsageStatusType

On Error GoTo ErrorHandling

    ' 直近の記録の調査
    Dim recent_record As MonitorRecord
    Dim i As Long
    
    If Not FindRecordWithinInterval( _
        pc_name, as_is, interval_seconds, interval_margin_seconds, recent_record, i _
    ) Then
        GoTo StatusIsLogOff
    End If
    
    ' 調査すべき時間間隔の間に存在するレコードに対するループ

    Dim buf As Variant
    buf = CreateTextBufferArray(BuildLogFilePath(pc_name, recent_record.DateTime))
    
    Dim current_record As MonitorRecord
    Set current_record = recent_record

    Dim previous_record As MonitorRecord
    
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Do While True
        ' current_record が Active か
        If IsActiveRecord(current_record) Then
            GoTo StatusIsActive
        End If
        
        ' 調査すべき時間枠の調査を完了したか
        If DateAdd("s", -interval_seconds, current_record.DateTime) _
            < DateAdd("s", -inspection_timespan_seconds, recent_record.DateTime) _
        Then
            GoTo StatusIsInactive
        End If
        
        ' ひとつ前のレコードを探す。場合分けを増やせば、既に調査済みの行は見ないようにできるが、可読性とコード量を減らす目的で
        ' FindRecordWithinInterval を current_record.DateTime に対して呼び出す。
        ' もし処理速度が問題になったときには、修正を検討してみてもいいかもしれない。
        ' FindRecordWithinInterval で、bufを返すようにしたり、同じ日付のファイルをスキップするオプションつける？
        Set previous_record = New MonitorRecord
        If i = LBound(buf) Then
            ' bufを読み切っていた場合、関数でしかるべき記録を探す
            ' ひとつ前の記録が入っているべきファイルが存在しない場合　→　Active (m)
            If Not FindRecordWithinInterval( _
                pc_name, current_record.DateTime, interval_seconds, interval_margin_seconds, _
                previous_record, i _
            ) Then
                GoTo StatusIsActive
            End If
            
            ' ひとつ前の記録が入っているべきファイルが存在する場合
            buf = CreateTextBufferArray(BuildLogFilePath(pc_name, previous_record.DateTime))
            ' ループ前に減らしたとき、iになるように調整
            i = i + 1
        Else
            ' bufが残っている場合、bufのひとつ前のレコードを取得
            ' TODO: ここでカッコを外すとエラーになるが、理由がよく理解できていない。
            previous_record.Parse (buf(i - 1))
        End If
        
        ' ひとつ前の記録との間に、gapがある　→　Active (モニターが新たに起動した = ログインしたと想定)
        If DateDiff("s", previous_record.DateTime, current_record.DateTime) _
            > interval_seconds + interval_margin_seconds _
        Then
            GoTo StatusIsActive
        End If
        
        ' 次のレコードへ
        Set current_record = previous_record
        i = i - 1
    Loop

StatusIsActive:
    CalcAppUsageStatus = AppUsageStatusActive
    GoTo Closing

StatusIsLogOff:
    CalcAppUsageStatus = AppUsageStatusLogOff
    GoTo Closing

StatusIsInactive:
    CalcAppUsageStatus = AppUsageStatusInactive
    GoTo Closing

StatusIsNotTarget:
    CalcAppUsageStatus = AppUsageStatusNotTarget
    GoTo Closing

ErrorHandling:
    CalcAppUsageStatus = AppUsageStatusError
    GoTo Closing

Closing:
    Set fso = Nothing
    
End Function


''' <summary>
''' 指定した時点における直近のレコードを、指定した時間間隔だけ遡って探す。
''' </summary>
Public Function FindRecordWithinInterval( _
    pc_name As String, as_is As Date, _
    interval_seconds As Long, interval_margin_soconds As Long, _
    ByRef recent_record_out As MonitorRecord, ByRef i_out As Long _
) As Boolean

    ' 遡って直近の記録を探すべき日数
    Dim possible_date_back As Long
    possible_date_back = DateDiff("d", _
        DateValue(DateAdd("s", -(interval_seconds + interval_margin_soconds), as_is)), _
        DateValue(as_is) _
    )
    
    ' ファイルを遡る必要がなくなるまでループ
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    Dim date_back As Long
    Dim target_date As Date
    Dim log_file_path_of_the_day As String
    Dim buf As Variant
    Dim i As Long
    Dim record As MonitorRecord
    Set record = New MonitorRecord
    
    For date_back = 0 To possible_date_back
        ' 確認する日付
        target_date = DateAdd("d", -date_back, as_is)

        ' その日のファイルがある場合、その日のファイルを調査
        ' TODO: ここは別関数にできるかも？ fso使うのはここだけ
        log_file_path_of_the_day = BuildLogFilePath(pc_name, target_date)
        
        If fso.FileExists(log_file_path_of_the_day) Then
            
            ' ファイルの内容をテキストバッファ配列へ読み込み
            buf = CreateTextBufferArray(log_file_path_of_the_day)

            ' テキストバッファの後ろからループし、直近の記録を探す
            ' TODO: ここは別関数に抜き出せるかも？ 直近の記録を見つけるまで
            For i = UBound(buf) To LBound(buf) Step -1
                ' 未来の行はスキップ
                record.Parse (buf(i))
                ' 境界は含まない (時間が一致するレコードは未来の行としてスキップ)
                If record.DateTime >= as_is Then
                    GoTo Continue
                End If
                
                ' 直近の記録がみつかった
                
                ' interval を超えるギャップがある場合 → モニターは起動していなかった
                If DateDiff("s", record.DateTime, as_is) > interval_seconds + interval_margin_soconds Then
                    FindRecordWithinInterval = False
                    GoTo Closing
                End If
                
                ' モニターが起動していた
                FindRecordWithinInterval = True
                Set recent_record_out = record
                i_out = i
                GoTo Closing
                
Continue:
            Next
            ' このファイルに、直近の記録は無い
        End If
            
        ' その日のファイルがないか、あっても直近の行が見つからなかった
    Next
    
    ' 調べるべきファイルの中に、直近の記録が無かった
    FindRecordWithinInterval = False
    GoTo Closing

ErrorHandling:
Closing:
    Set fso = Nothing

End Function

''' <summary>
''' 指定したモニター出力ファイルを読み取り、改行で分割した配列を返す。空行は削除される。
''' </summary>
Public Function CreateTextBufferArray(file_path As String) As Variant
    ' 引数はファイル名がいいか、pc名と日付を渡して、内部でファイル名を作るのがいいのか？
    
    Dim buf As Variant
    
    With New ADODB.Stream
        .Charset = "UTF-8"
        .Open
        .LoadFromFile file_path
        buf = Split(.ReadText(adReadAll), vbCrLf)
        .Close
    End With
    
    ' 空行を除外する
    ' 以下では全ての要素を調査しているが、最終行にしか空行が無い想定なので、最後だけ確認して削除すればいいかも
    Dim eliminated_buf As Variant
    Dim i As Long
    Dim j As Long

    j = 0
    ReDim eliminated_buf(UBound(buf))
    
    For i = LBound(buf) To UBound(buf)
        If buf(i) = "" Then
            GoTo Continue
        End If
        
        eliminated_buf(j) = buf(i)
        j = j + 1
Continue:
    Next
    
    ReDim Preserve eliminated_buf(j - 1)
    
    CreateTextBufferArray = eliminated_buf

End Function

''' <summary>
''' レコードが Active と判断されるものかどうかを判定する。
''' </summary>
Private Function IsActiveRecord(record As MonitorRecord) As Boolean
    
    If Not record.IsProcessActive Then
        IsActiveRecord = False
        Exit Function
    End If
    
    If record.IsDisplayLocked Then
        IsActiveRecord = False
        Exit Function
    End If
    
    IsActiveRecord = True
    
End Function

''' <summary>
''' PC名と日付からログファイルのパスを生成する。
''' </summary>
Private Function BuildLogFilePath(pc_name As String, as_is As Date) As String
    
    Dim fso As Object
    Set fso = New Scripting.FileSystemObject

    Dim log_file_name As String
    log_file_name = Format(as_is, "yyyy-MM-dd") & "_" & pc_name & ".txt"

    BuildLogFilePath = fso.BuildPath(monitorOutputFolderPath_, log_file_name)

ErrorHandling:
Closing:
    Set fso = Nothing

End Function
