Attribute VB_Name = "ViewController"
''' <summary>
''' View (Worksheet) を操作する関数をまとめたモジュール
''' </summary>

Option Explicit


''' <summary>
''' PCのステータスの分類
''' </summary>
Public Enum AppUsageStatusType

    AppUsageStatusActive
    AppUsageStatusLogOff
    AppUsageStatusInactive
    AppUsageStatusNotTarget
    AppUsageStatusError

End Enum


''' <summary>
''' 現在の条件でシートを更新する。
''' </summary>
''' <remarks>
''' Update ボタンを押したときに実行されるべきプロシージャ
''' </remarks>
Public Sub Update()

    Dim status_calculator As AppUsageStatusCalculator
    Set status_calculator = New AppUsageStatusCalculator
    
    status_calculator.Init MonitorOutputFolderPath

    UpdateAsIs _
        Now(), MonitorSettingsObject.TargetPCs, MonitorSettingsObject.IntervalSeconds, _
        3, 30 * 60, _
        status_calculator
    
End Sub


''' <summary>
''' 指定した条件でシートを更新する。
''' </summary>
''' <remarks>
''' テストのために Public にしておく
''' </remarks>
Public Sub UpdateAsIs( _
    target_datetime As Date, target_pcs As Variant, interval_seconds As Long, _
    interval_margin_seconds As Long, inspection_timespan_seconds As Long, _
    calculator As AppUsageStatusCalculator _
)

    ' TODO: Active sheet の設定をした方がいいかな？

    ' ターゲットPCのループ
    Dim i As Long
    Dim status As AppUsageStatusType
    Dim pc_name As String
    Dim named_cell As Variant
    
    For i = LBound(target_pcs) To UBound(target_pcs)
        pc_name = CStr(target_pcs(i))
        
        ' pc_name に対応するセル名が無ければスキップ
        If Not IsExistsNamedCell(GetCellNameFromPCName(pc_name)) Then
            GoTo Continue:
        End If

        ' ステータスの判定
        status = calculator.CalcAppUsageStatus( _
            pc_name, target_datetime, interval_seconds, interval_margin_seconds, inspection_timespan_seconds _
        )
        
        ' ステータスに基づく書式の適用
        ApplyFormatFromLegend _
            Range(GetCellNameFromPCName(pc_name)), Range(GetLegendCellNameFromStatus(status))
Continue:
    Next
    
    ' 更新日セルの更新
    UpdateLastUpdateCell target_datetime
    
End Sub

''' <summary>
''' 指定した名前付きセルがワークシートに存在するかどうかを調べる。
''' </summary>
Public Function IsExistsNamedCell(cell_name As String) As Boolean

    Dim named_cell As Variant
    ' シート名を外して、名前だけ取り出すための配列
    Dim buf As Variant
    
    For Each named_cell In ActiveSheet.Names
        buf = Split(named_cell.Name, "!")
        
        If cell_name = buf(UBound(buf)) Then
            IsExistsNamedCell = True
            Exit Function
        End If
    Next
    
    IsExistsNamedCell = False

End Function


''' <summary>
''' PC名から、対応するセル名を返す
''' </summary>
Public Function GetCellNameFromPCName(pc_name As String) As String
    Dim cell_name As String

    ' アンスコ化
    cell_name = Replace(pc_name, "-", "_")
    
    GetCellNameFromPCName = cell_name
    
End Function


''' <summary>
''' Status から凡例セルの名前（番地）を返す
''' </summary>
Public Function GetLegendCellNameFromStatus(status As AppUsageStatusType) As String

    Select Case status
        Case AppUsageStatusActive
            GetLegendCellNameFromStatus = "LegendActive"
            
        Case AppUsageStatusLogOff
            GetLegendCellNameFromStatus = "LegendLogOff"
            
        Case AppUsageStatusInactive
            GetLegendCellNameFromStatus = "LegendInactive"

        Case AppUsageStatusNotTarget
            GetLegendCellNameFromStatus = "LegendNotTarget"

        Case AppUsageStatusError
            GetLegendCellNameFromStatus = "LegendError"
    End Select

End Function


''' <summary>
''' 更新日セルを更新する。
''' </summary>
Public Sub UpdateLastUpdateCell(target_datetime As Date)

    Range("LastUpdate") = target_datetime

End Sub


''' <summary>
''' 凡例セルの書式を対象セルへ適用する
''' </summary>
Public Sub ApplyFormatFromLegend(targetCell As Range, legendCell As Range)

    targetCell.Interior.Color = legendCell.Interior.Color
    ' Font 全体を直接渡せる？
    targetCell.Font.Color = legendCell.Font.Color
    targetCell.Font.Bold = legendCell.Font.Bold

End Sub

