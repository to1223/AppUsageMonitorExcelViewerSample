Attribute VB_Name = "BookEventHandlers"
Option Explicit


''' <summary>
''' �u�b�N�N�����Ɏ��s���鏈���B
''' </summary>
Public Sub OnBookOpen()

    ' Logger �̐ݒ�
    Set Logger = New DebugPrintLogger
    
    ' ���j�^�[�ݒ�̓ǂݎ��
    Set MonitorSettingsObject = New MonitorSettings
    MonitorSettingsObject.Init MonitorSettingsFolderPath, MonitorOutputFolderPath

End Sub
