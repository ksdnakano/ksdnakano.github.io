Attribute VB_Name = "PCTK"
'
'���[�h�ǂݏグ�}�N��
'���̃v���O�����͒��쌠�ŕی삳��Ă��܂��B
'�i���j���m�V�X�e���J��

Option Explicit
Private Declare Function PCTKSETSTATUS Lib "PCTKUSR.dll" (ByVal para1 As Long, ByVal para2 As Long, ByVal para3 As Long) As Long
Private Declare Sub PCTKPREAD Lib "PCTKUSR.dll" (ByVal vString As String, ByVal vMode As Long, ByVal vFlag As Long)
Private Declare Sub PCTKCGUIDE Lib "PCTKUSR.dll" (ByVal vString As String, ByVal vMode As Long)
Private Declare Sub PCTKVRESET Lib "PCTKUSR.dll" ()
Private Declare Function PCTKGETVSTATUS Lib "PCTKUSR.dll" () As Long
Private Declare Sub PCTKBEEP Lib "PCTKUSR.dll" (ByVal BeepType As Long, ByVal UINT As Long, ByVal MachinType As Long)
'PCTalker���[�U�[�C���^�[�t�F�C�X�w�b�_
Public Const PS_AGSSTATUS = &H20000
Public Const PS_VOSSTATUS = &H10000

Public Const PS_APREADCALLBACK = &H20001       '   �ǂݏグ���R�[���o�b�N�֐��ݒ�
Public Const PS_APNOREAD = &H20002     '   �ǂݏグ�֎~�ݒ�
Public Const PS_APNOACTION = &H20003       '   PCTalker����֎~�ݒ�
Public Const PS_APSCUTCALLBACK = &H20004       '   �V���[�g�J�b�g�R�[���o�b�N�ݒ�
Public Const PS_STATUS = &H10001       '   ���ڐݒ�
Public Const PS_STATUSNEXT = &H10002       '   ���ڐݒ�g�O����
Public Const PS_STATUSBACK = &H10003       '   ���ڐݒ�g�O���O

'
'   PCTalker�ǂݏグ���쎞�R�[���o�b�N�֐��̈��� act
'   �A�v���P�[�V�����̓ǂݏグ����
Const ACV_MENUSELECT = 0                ' ���j���[�I�����̓ǂݏグ
Const ACV_MENUCLOSE = 1                 ' ���j���[�N���[�Y���̓ǂݏグ
Const ACV_CBOPENCLOSE = 2           ' �R���{�{�b�N�X�I�[�v���N���[�Y���̓ǂݏグ
Const ACV_WINACTIVATE = 3           ' �E�B���h�E�A�N�e�B�u���̓ǂݏグ
Const ACV_WINCREATE = 4             ' �A�v���P�[�V�����̋N�������̓ǂݏグ
Const ACV_WINCLOSE = 5          ' �E�B���h�E�N���[�Y���̓ǂݏグ
Const ACV_LTSELECT = 6              ' ���X�g�n�A�^�u�V�[�g�I��ؑ֎��̓ǂݏグ
Const ACV_TRACKBAR = 7          ' �g���b�N�o�[�̈ړ����̓ǂݏグ
Const ACV_FOCUS = 8             ' �t�H�[�J�X�ݒ莞�̓ǂݏグ
Const ACV_CHKBUTTON = 9             ' �`�F�b�N�{�^���ؑ֎��̓ǂݏグ
Const ACV_EDITMOVECUR = 10              ' ��ި�ĺ��۰كJ�[�\���ړ����̓ǂݏグ

'   �V�����R�[���o�b�N�֐��t���O�l(PC-Talker Ver 3.0�ȏ�)
Const ACV_IMM_ONOFF = 11                ' ���{��ϊ�ON/OFF���ǂݏグ
Const ACV_IMM_INPUT = 12                ' ���{����͎��ǂݏグ
Const ACV_IMM_MODECHANGE = 13               ' ���{��ϊ����[�h�ؑ֎��ǂݏグ
Const ACV_CHAR_INPUT = 14               ' ����(���p)���͎��ǂݏグ


'   PCTKVoiceGuide����  quick
'   �N�B�b�N�ݒ�l
Const PCTKQUICK_NONE = 0                '   �V�X�e���ɔC��
Const PCTKQUICK_OFF = 1             '   �N�B�b�NOFF
Const PCTKQUICK_ON = 2              '   �N�B�b�NON

Const PCTKERR_NOERROR = 0           '   ����
Const PCTKERR_OLEINIT = &H10001     '   OLE�̏������Ɏ��s����
Const PCTKERR_SERVEREXEC = &H10002      '   Server�N���G���[
Const PCTKERR_INTERFACE = &H10003       '   �C���^�[�t�F�C�X�擾�G���[
Public WaitFlag As Boolean
'���{��ϊ�
Public Declare Function ImmGetContext Lib "imm32.dll" (ByVal hwnd As Long) As Long
Public Declare Function ImmGetCompositionString Lib "imm32.dll" Alias "ImmGetCompositionStringA" (ByVal himc As Long, ByVal dw As Long, lpv As Any, ByVal dw2 As Long) As Long
Public Const GCS_COMPSTR = &H8
Public Const IME_CMODE_SYMBOL = &H400
Public Declare Function GetFocus Lib "user32" () As Long


Public Sub Voice(str As String)
''    Voice1.Go str
''    Exit Sub
'    If WaitFlag = True Then
'        Exit Sub
'    End If
    BackVoice = False
    DeleteVoice = False
    Debug.Print "VOICE<"; str;
    On Error GoTo ss1
    If Len(str) < 1 Then Exit Sub
    Call PCTKPREAD(str, 5, 1)
    Debug.Print ">"
ss1:
End Sub
Public Sub SVoice(str As String)
    Debug.Print "SVOICE<"; Hex$(Asc(Mid$(str, 1, 1)));
    On Error GoTo ss1
    BackVoice = False
    DeleteVoice = False
    If Len(str) > 1 Then
        Debug.Print "**"; Hex$(Asc(Mid$(str, 2, 1)))
    End If
    If str = Chr$(&HB) Then
        Voice "�s����"
    ElseIf str = Chr$(1) + Chr$(21) Then
        Voice "�s�N�`���["
    ElseIf str = Chr$(13) + Chr$(7) Then
        Voice "�Z�����s"
    ElseIf str = Chr$(&H8140) Then
        Call ExBeep(2, 10, 0)
    ElseIf str = Chr$(&H20) Then
        Call ExBeep(3, 10, 0)
    Else
        If Len(str) < 1 Then Exit Sub
        str = Left$(str, 1)
        Call PCTKCGUIDE(str, 0)
'        PCTK.GUID str, 0
    End If
    Debug.Print ">"

ss1:
End Sub
Public Sub SSVoice(str As String)
    Debug.Print "SVOICE<"; Hex$(Asc(Mid$(str, 1, 1)));
    On Error GoTo ss1
    BackVoice = False
    DeleteVoice = False
    If Len(str) > 1 Then
        Debug.Print "**"; Hex$(Asc(Mid$(str, 2, 1)))
    End If
    If str = Chr$(&HB) Then
        Voice "�s����"
    ElseIf str = Chr$(1) + Chr$(21) Then
        Voice "�s�N�`���["
    ElseIf str = Chr$(13) + Chr$(7) Then
        Voice "�Z�����s"
    ElseIf str = Chr$(&H8140) Then
        Call ExBeep(2, 10, 0)
    ElseIf str = Chr$(&H20) Then
        Call ExBeep(3, 10, 0)
    Else
        If Len(str) < 1 Then Exit Sub
            str = Left$(str, 1)
            Call PCTKCGUIDE(str, 6)
    '        PCTK.GUID str, 0
        End If
    Debug.Print ">"

ss1:
End Sub
Public Sub QVoice(str As String)
    Debug.Print "QVoice<";
    On Error GoTo ss1
    Call PCTKVRESET
    Call PCTKPREAD(str, 1, 1)
ss1:
    Debug.Print ">"
End Sub
Public Sub Quick()
'    If WaitFlag = True Then
'        Exit Sub
'    End If
    
    Debug.Print "Quick<";
    On Error GoTo ss1
    Call PCTKVRESET
    Debug.Print ">"
ss1:
End Sub
Public Sub VoicePause()
    On Error GoTo ss1
    Call PCTKSETSTATUS(PS_APNOACTION, -1, vbNull)
ss1:
End Sub
Public Sub VoiceActive()
    On Error GoTo ss1
    Call PCTKSETSTATUS(PS_APNOACTION, 0, vbNull)
ss1:
End Sub

Public Sub VoiceWait()
    WaitFlag = True
    Debug.Print "VoiceWait<";
    On Error GoTo ss1
        DoEvents
    While PCTKGETVSTATUS() <> 0
        DoEvents
        DoEvents
    Wend
    Debug.Print ">"
ss1:
    WaitFlag = False
End Sub
Public Sub ExBeep(para1 As Long, para2 As Long, para3 As Long)
On Error GoTo ss1
    Call PCTKBEEP(para1, para2, para3)
ss1:
End Sub

Public Function IMMuse() As Boolean
    Dim inHC, usecnv, setcnv, hWindMe As Long
    IMMuse = False
 
    hWindMe = GetFocus()
    
    If hWindMe = vbNull Then Exit Function
    inHC = ImmGetContext(hWindMe)
    If inHC = vbNull Then Exit Function
    If ImmGetCompositionString(inHC, GCS_COMPSTR, vbNull, 0) > 0 Then
            IMMuse = True
            Exit Function
    End If
    

End Function
