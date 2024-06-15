Attribute VB_Name = "ViewController"
''' <summary>
''' View (Worksheet) �𑀍삷��֐����܂Ƃ߂����W���[��
''' </summary>

Option Explicit


''' <summary>
''' PC�̃X�e�[�^�X�̕���
''' </summary>
Public Enum AppUsageStatusType

    AppUsageStatusActive
    AppUsageStatusLogOff
    AppUsageStatusInactive
    AppUsageStatusNotTarget
    AppUsageStatusError

End Enum


''' <summary>
''' ���݂̏����ŃV�[�g���X�V����B
''' </summary>
''' <remarks>
''' Update �{�^�����������Ƃ��Ɏ��s�����ׂ��v���V�[�W��
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
''' �w�肵�������ŃV�[�g���X�V����B
''' </summary>
''' <remarks>
''' �e�X�g�̂��߂� Public �ɂ��Ă���
''' </remarks>
Public Sub UpdateAsIs( _
    target_datetime As Date, target_pcs As Variant, interval_seconds As Long, _
    interval_margin_seconds As Long, inspection_timespan_seconds As Long, _
    calculator As AppUsageStatusCalculator _
)

    ' TODO: Active sheet �̐ݒ�����������������ȁH

    ' �^�[�Q�b�gPC�̃��[�v
    Dim i As Long
    Dim status As AppUsageStatusType
    Dim pc_name As String
    Dim named_cell As Variant
    
    For i = LBound(target_pcs) To UBound(target_pcs)
        pc_name = CStr(target_pcs(i))
        
        ' pc_name �ɑΉ�����Z������������΃X�L�b�v
        If Not IsExistsNamedCell(GetCellNameFromPCName(pc_name)) Then
            GoTo Continue:
        End If

        ' �X�e�[�^�X�̔���
        status = calculator.CalcAppUsageStatus( _
            pc_name, target_datetime, interval_seconds, interval_margin_seconds, inspection_timespan_seconds _
        )
        
        ' �X�e�[�^�X�Ɋ�Â������̓K�p
        ApplyFormatFromLegend _
            Range(GetCellNameFromPCName(pc_name)), Range(GetLegendCellNameFromStatus(status))
Continue:
    Next
    
    ' �X�V���Z���̍X�V
    UpdateLastUpdateCell target_datetime
    
End Sub

''' <summary>
''' �w�肵�����O�t���Z�������[�N�V�[�g�ɑ��݂��邩�ǂ����𒲂ׂ�B
''' </summary>
Public Function IsExistsNamedCell(cell_name As String) As Boolean

    Dim named_cell As Variant
    ' �V�[�g�����O���āA���O�������o�����߂̔z��
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
''' PC������A�Ή�����Z������Ԃ�
''' </summary>
Public Function GetCellNameFromPCName(pc_name As String) As String
    Dim cell_name As String

    ' �A���X�R��
    cell_name = Replace(pc_name, "-", "_")
    
    GetCellNameFromPCName = cell_name
    
End Function


''' <summary>
''' Status ����}��Z���̖��O�i�Ԓn�j��Ԃ�
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
''' �X�V���Z�����X�V����B
''' </summary>
Public Sub UpdateLastUpdateCell(target_datetime As Date)

    Range("LastUpdate") = target_datetime

End Sub


''' <summary>
''' �}��Z���̏�����ΏۃZ���֓K�p����
''' </summary>
Public Sub ApplyFormatFromLegend(targetCell As Range, legendCell As Range)

    targetCell.Interior.Color = legendCell.Interior.Color
    ' Font �S�̂𒼐ړn����H
    targetCell.Font.Color = legendCell.Font.Color
    targetCell.Font.Bold = legendCell.Font.Bold

End Sub

