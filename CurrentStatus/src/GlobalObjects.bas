Attribute VB_Name = "GlobalObjects"
''' �O���[�o���Ȋe��I�u�W�F�N�g�ϐ����`
'
' �u�b�N�I�[�v�����Ɏ��̂𐶐����Đݒ肷��B
' �N���X�̒�����͎g�킸�ɁA��O���N���X���g�����ŕ⑫���āA��������Logger���g�����������̂���

Option Explicit

' AppUsageMonitor �̐ݒ�t�@�C���̖��O
Public Const IntervalSecondsFileName = "interval_seconds"
Public Const TargetPCsFileName = "target_pc"

' Logger
Public Logger As ILogger

' MonitorSettings
Public MonitorSettingsObject As MonitorSettings

