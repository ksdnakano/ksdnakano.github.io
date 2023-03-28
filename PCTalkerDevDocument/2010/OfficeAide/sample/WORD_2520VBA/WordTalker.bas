Attribute VB_Name = "WordTalker"
'
'���[�h�ǂݏグ�}�N��
'���̃v���O�����͒��쌠�ŕی삳��Ă��܂��B
'�i���j���m�V�X�e���J��
'
Option Explicit
'Public PCTK As PCTKSpeech
Public TateYokoMode As Long
'Public OnPctkF As Boolean
Public LineVoiceFlag As Boolean
Public LeftVoiceFlag As Boolean
Public RightVoiceFlag As Boolean
Public TopVoiceFlag As Boolean
Public DownVoiceFlag As Boolean
Public MacroGo As Boolean
Public BackVoice As Boolean
Public DeleteVoice As Boolean

Dim x As New Class1
Public AllReadFlag As Boolean
Public KeyBack As VbPCTK.Callback
Public KeyTimer As VbPCTK.VbTimer

Public dum As Long


'
'�� �L�[�{�[�h�}�N���̐ݒ� ��
'
Sub SetKeyEvent()
    Quick
    MenuYomiAdd
    CheckPcTalker
    KeyBindings.Add KeyCode:=BuildKeyCode(38), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curUp" '��
    KeyBindings.Add KeyCode:=BuildKeyCode(38, wdKeyShift), KeyCategory:=wdKeyCategoryMacro _
        , Command:="SHcurUp" 'S+��
    KeyBindings.Add KeyCode:=BuildKeyCode(40), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curDown" '��
    KeyBindings.Add KeyCode:=BuildKeyCode(40, wdKeyShift), KeyCategory:=wdKeyCategoryMacro _
        , Command:="SHcurDown" 'S+��
    KeyBindings.Add KeyCode:=BuildKeyCode(37), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curLeft" '��
    KeyBindings.Add KeyCode:=BuildKeyCode(37, wdKeyShift), KeyCategory:=wdKeyCategoryMacro _
        , Command:="SHcurLeft" 'S+��
    KeyBindings.Add KeyCode:=BuildKeyCode(39), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curRight"  '��
     KeyBindings.Add KeyCode:=BuildKeyCode(39, wdKeyShift), KeyCategory:=wdKeyCategoryMacro _
        , Command:="SHcurRight"  'S+��
       
     KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyHome, wdKeyControl), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curTopPage"  'C+HOME
     KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyEnd, wdKeyControl), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curLastPage"  'C+End
        
       
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyBackspace), KeyCategory:= _
    wdKeyCategoryMacro, Command:="BsChar"  '[BackSpace]
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyDelete), KeyCategory:= _
    wdKeyCategoryMacro, Command:="DelChar" '[Delete]
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyInsert), KeyCategory:= _
    wdKeyCategoryMacro, Command:="InsDel" '[Insert]
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyPageUp), KeyCategory:= _
    wdKeyCategoryMacro, Command:="curPageUp" '[PageUp]
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyPageDown), KeyCategory:= _
    wdKeyCategoryMacro, Command:="curPageDown" '[PageDown]
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyHome), KeyCategory:= _
    wdKeyCategoryMacro, Command:="curHome" '[Home]
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyEnd), KeyCategory:= _
    wdKeyCategoryMacro, Command:="curEnd" '[End]
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyEsc), KeyCategory:= _
    wdKeyCategoryMacro, Command:="Escape" '[Esc]

    'Alt+F?
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF1, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="FileRead"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF3, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="PageSetRead"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF8, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="LineReadAll"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF9, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="ParaRead"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF10, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="AllRead"
    'F?
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF6), KeyCategory:= _
        wdKeyCategoryMacro, Command:="KakudaiChange"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF9), KeyCategory:= _
        wdKeyCategoryMacro, Command:="PosVoice"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF10), KeyCategory:= _
        wdKeyCategoryMacro, Command:="LineReadAll"
    'Ctrl+?
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyK, wdKeyAlt, wdKeyShift), KeyCategory:= _
        wdKeyCategoryMacro, Command:="TableInfo"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF12, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="SetKeyEvent"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF11, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="ClearKeyEvent"
End Sub
'
'�� �L�[�{�[�h�}�N���̃N���A ��
'
Sub ClearKeyEvent()
    Quick
        Dim menuObj As Object

    Voice "�ǂݏグ ��~"
    KeyBindings.ClearAll
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF12, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="SetKeyEvent"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF11, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="ClearKeyEvent"
    CommandBars.ActiveMenuBar.Reset
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
        If menuObj.Caption = "�ǂ�(&R)" Then
            menuObj.Delete
            DoEvents
        End If
    Next
'    MenuYomiAdd2
End Sub


'
'�� �ǂ��ƭ���ǉ����� ��
'
Sub MenuYomiAdd()
    '�������ƭ�������΍폜
    Dim menuObj As Object
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
        If menuObj.Caption = "�ǂ�(&R)" Then
            menuObj.Delete
            DoEvents
        End If
    Next
    '�ƭ���ǉ�
    
    Dim myMenuBar, newMenu, ctrl1
    Set myMenuBar = CommandBars.ActiveMenuBar
    Set newMenu = myMenuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True, Before:=2)
    newMenu.Caption = "�ǂ�(&R)"
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton, ID:=1)
    With ctrl1
        .Caption = "���وʒu����S���ǂ�(&O) Alt+F10"
        .TooltipText = "���وʒu����S���ǂ�"
        .Style = msoButtonCaption
        .OnAction = "AllRead"
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "�i���ǂ�(&S) Alt+F9"
        .TooltipText = "�i���ǂ�"
        .Style = msoButtonCaption
        .OnAction = "ParaRead"
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "�P�s�ǂ�(&L) Alt+F8"
        .TooltipText = "�P�s�ǂ�"
        .Style = msoButtonCaption
        .OnAction = "LineReadAll"
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "̧�ٖ� �ǂݏグ(&F)"
        .TooltipText = "̧�ٖ� �ǂݏグ"
        .Style = msoButtonCaption
        .OnAction = "FileRead"
        .BeginGroup = True
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "�߰�ސݒ� �ǂݏグ(&P)"
        .TooltipText = "̧�ٖ��̓ǂݏグ"
        .Style = msoButtonCaption
        .OnAction = "PageSetRead"
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "Word�ǂݏグ�w���v(&H)"
        .TooltipText = "ܰ�ޓǂݏグ�w���v"
        .Style = msoButtonCaption
        .OnAction = "WordReadHelp"
        .BeginGroup = True
    End With
    DoEvents
End Sub
'
'�� �ǂ��ƭ���ǉ����� ��
'
Sub MenuYomiAdd2()
    '�������ƭ�������΍폜
    Dim menuObj As Object
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
        If menuObj.Caption = "�ǂ�(&R)" Then
            menuObj.Delete
        End If
    Next
    '�ƭ���ǉ�
    
    Dim myMenuBar, newMenu, ctrl1
    Set myMenuBar = CommandBars.ActiveMenuBar
    Set newMenu = myMenuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True, Before:=2)
    newMenu.Caption = "�ǂ�(&R)"
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton, ID:=1)
    With ctrl1
        .Caption = "�ǂ݂�����~(&C) Esc"
        .TooltipText = "�ǂ݂������~���܂�"
        .Style = msoButtonCaption
        .OnAction = "Escape"
    End With
    
    DoEvents
End Sub

'
'�� ܰ�ޓǂݏグ�w���v ��
'
Sub WordReadHelp()
    On Error GoTo ss1:
    Dim i
    Dim objDoc As Object
    For Each objDoc In Windows
        If objDoc.Caption = "Word�ǂݏグ�w���v" Then
            objDoc.Activate
            Quick
            Voice "Word�ǂݏグ�w���v "
            Exit Sub
        End If
    Next
    Set objDoc = Documents.Add
    For i = 0 To 100
        DoEvents
    Next
    Quick
    ActiveDocument.ActiveWindow.Caption = "Word�ǂݏグ�w���v"
    Voice "Word�ǂݏグ�w���v "
    ActiveDocument.Range.Font.Bold = True
    ActiveDocument.Range.Font.Size = 13

    ActiveDocument.Range.InsertAfter "��Word�ǂݏグ�w���v��    "
    ActiveDocument.Range.InsertAfter "�g�p�Җ�:" + System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\KSD\License", "OwnerString") + vbCr + vbCr
    ActiveDocument.Range.InsertAfter "�@�@�t�@�C�����ǂݏグ Alt+F1" + vbCr
    ActiveDocument.Range.InsertAfter "�@�@�y�[�W�ݒ�ǂݏグ Alt+F3" + vbCr
    ActiveDocument.Range.InsertAfter "�@�@�P�s�ǂ� Alt+F8" + vbCr
    ActiveDocument.Range.InsertAfter "�@�@�i���ǂ� Alt+F9" + vbCr
    ActiveDocument.Range.InsertAfter "�@�@�J�[�\���ʒu����S���ǂ� Alt+F10" + vbCr
    ActiveDocument.Range.InsertAfter "�@�@�ʏ�^�g��؊� F6" + vbCr
    ActiveDocument.Range.InsertAfter "�@�@�J�[�\���ʒu�̓ǂݏグ F9" + vbCr
    ActiveDocument.Range.InsertAfter "�@�@�r�����ǂ� Shift+Alt+K" + vbCr
    If Model = 0 Then
        ActiveDocument.Range.InsertAfter vbCr + "���ȉ���PC-Talker�̺���ނƓ�������̂��̂ł��B" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�������ǂ� Ctrl+Alt+G" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@���ʓǂ� Ctrl+Alt+M" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�s������J�[�\����O�ǂ� Ctrl+Alt+H" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@JIS�R�[�h�ǂ݁@Ctrl+Alt+I" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�P�s�ǂ� Ctrl+Alt+J" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�J�[�\���ʒu����s���ǂ� Ctrl+Alt+K" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�J�[�\���ʒu�ǂ� Ctrl+Alt+,(�R���})" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�J�[�\���ʒu����S���ǂ� Ctrl+Alt+A" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�I��͈͂̓ǂ݂��� Shift" + vbCr
        ActiveDocument.Range.InsertAfter vbCr + "�@�@Word�ǂݏグ�@�\�̒�~ Alt+F11" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@Word�ǂݏグ�@�\�̊J�n Alt+F12" + vbCr
        ActiveDocument.Range.InsertAfter vbCr
    Else
        ActiveDocument.Range.InsertAfter vbCr + "���ȉ���VDM100W�̺���ނƓ�������̂��̂ł��B" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�������ǂ� Ctrl+Alt+-(ϲŽ)" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@���ʓǂ� Ctrl+Alt+H" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�s������J�[�\����O�ǂ� Ctrl+Alt+J" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@JIS�R�[�h�ǂ݁@Ctrl+Alt+N" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�P�s�ǂ� Ctrl+Alt+K" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�J�[�\���ʒu����s���ǂ� Ctrl+Alt+L" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�J�[�\���ʒu�ǂ� Ctrl+Alt+,(���)" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�J�[�\���ʒu����S���ǂ� Ctrl+Alt+:(���)" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@�I��͈͂̓ǂ݂��� Shift" + vbCr
        ActiveDocument.Range.InsertAfter vbCr + "�@�@Word�ǂݏグ�@�\�̒�~ Alt+F11" + vbCr
        ActiveDocument.Range.InsertAfter "�@�@Word�ǂݏグ�@�\�̊J�n Alt+F12" + vbCr
        ActiveDocument.Range.InsertAfter vbCr
    End If
    ActiveDocument.Range.InsertAfter "��Word�ǂ݂����@�\�ɂ���" + vbCr
    ActiveDocument.Range.InsertAfter "����E����@�����m�V�X�e���J��" + vbCr
    ActiveDocument.Range.InsertAfter "���{�i�I�ȉ������[�v���̂����p�́A�O��I�ɑ��쐫��ǋy����" + vbCr + "����Ɏg���₷���Ȃ����A���{�ꃏ�[�v��" + vbCr _
    + "�u�l�x�v�n�q�c�V���[�Y�v�����Ђ����p���������B" + vbCr
    ActiveDocument.Range.InsertAfter "URL http://www.aok-unet.ocn.ne.jp/text" + vbCr
ss1:
End Sub
'
'�� ̧�ٖ��ǂݏグ ��
'
Sub FileRead()
    On Error GoTo ee1
    Quick
    Voice ActiveDocument.Name
ee1: Exit Sub



End Sub


'
'�� �߰�ސݒ�ǂݏグ ��
'
Sub PageSetRead()
    On Error GoTo ee1
    Quick
'
'�@�p��
    Select Case ActiveDocument.PageSetup.PaperSize

    Case wdPaperA3
        Voice "�T�C�Y A3"
    Case wdPaperA4
        Voice "�T�C�Y A4"
    Case wdPaperA4Small
        Voice "�T�C�Y A4 �X���[��"
    Case wdPaperA5
        Voice "�T�C�Y A5"
    Case wdPaperB4
        Voice "�T�C�Y B4"
    Case wdPaperB5
        Voice "�T�C�Y B5"
    Case wdPaperLetter
        Voice "�T�C�Y�@���^�["
    Case Else
        Voice "�p������" + CStr(ActiveDocument.PageSetup.PaperSize)
    End Select
'
' �u����
    If ActiveDocument.PageSetup.Orientation = wdOrientLandscape Then
        Voice " ������"
    Else
        Voice " �c����"
    End If
'
' �s�E������
    Voice " " + CStr(ActiveDocument.PageSetup.LinesPage) + "�s " + _
        CStr(ActiveDocument.PageSetup.CharsLine) + "�� "
'
'�@������
    If ActiveDocument.Range.Orientation = wdTextOrientationHorizontal Then
        Voice " ������"
    ElseIf ActiveDocument.Range.Orientation = wdTextOrientationVerticalFarEast Then
        Voice " �c����"
    End If
    Voice "�v�����^�� " + Application.ActivePrinter
ee1:
End Sub

'�� �ʏ�^�g��؊� ��
'
Sub KakudaiChange()
    On Error GoTo ss1
    Dim ss As Long
    ss = Selection.Start
        If Options.BlueScreen = False Then
            Quick
            Voice "�g����"
            Options.BlueScreen = True
            ActiveWindow.ActivePane.View.Zoom.Percentage = 500
            dum = Selection.MoveLeft
            If ss <> Selection.Start Then
                dum = Selection.MoveRight
            End If
        Else
            Quick
            Voice "�ʏ���"
            Options.BlueScreen = False
            ActiveWindow.ActivePane.View.Zoom.Percentage = 100
        
        End If
ss1:
End Sub



'
'���@���[�h�ǂݏグ�C�j�V�����C�Y�@��
'
Sub Vinit()
'    If OnPctkF = True Then
'        Exit Sub
'    End If
    CheckPcTalker
End Sub
'
'���@ܰ�ޓǂݏグ�����̏������@��
'
Sub CheckPcTalker()

'Application ������ׯ�߂̐ݒ�
Set x.App = ThisDocument.Application

On Error GoTo ee1
Set x.AppVB = New VbPCTK.Callback
Set x.AppVBtim = New VbPCTK.VbTimer


'COM
'On Error GoTo aa1:
'Set PCTK = New PCTKSpeech
'    If OnPctkF = False Then
'        OnPctkF = True
'        Quick
'        'Voice "�����J�n"
'    End If
'    Exit Sub
'
'aa1:
'    OnPctkF = False
    Exit Sub
ee1: MsgBox ("Word�ǂݏグ�ɕK�v�ȃv���O������������܂���")


End Sub
'
' ���@JIS�R�[�h�ǂ݁@��
'
Sub JisVoice()
    On Error GoTo ee1
    Dim s As String
    Dim i As Integer
    s = Selection.Text
    s = Right$("000" + Hex$(Asc(s)), 4)
    Quick
    Voice "�V�t�g�W�X�R�[�h "
    For i = 1 To 4: SVoice (Mid$(s, i, 1)): VoiceWait: Next
    
ee1: Exit Sub
End Sub

'
' ���@�ʒu�ǂ݁@��
'
Sub PosVoice()
    On Error GoTo ee1
    Dim Page, Keta, Gyo, Sect, Xmm, Ymm As Long
    Page = Selection.Information(wdActiveEndPageNumber)
    Keta = Selection.Information(wdFirstCharacterColumnNumber)
    Gyo = Selection.Information(wdFirstCharacterLineNumber)
    Sect = Selection.Information(wdActiveEndSectionNumber)
    Ymm = Selection.Information(wdVerticalPositionRelativeToPage)
    Xmm = Selection.Information(wdHorizontalPositionRelativeToPage)
    If Gyo = -1 Then
        If Sect = 1 Then
            Voice CStr(Page) + "�� " + CStr(Keta) + "��" + Chr$(&HFE)
        Else
            Voice CStr(Page) + "�� " + CStr(Sect) + "�Z�N�V���� " + CStr(Keta) + "��" + Chr$(&HFE)
        End If
    Else
        If Sect = 1 Then
            Voice CStr(Page) + "�� " + CStr(Gyo) + "�s " + CStr(Keta) + "��" + Chr$(&HFE)
        Else
            Voice CStr(Page) + "�� " + CStr(Sect) + "�Z�N�V���� " + CStr(Gyo) + "�s " + CStr(Keta) + "��" + Chr$(&HFE)
        End If
    End If
    Voice "�c " + CStr(Int(Ymm / 72 * 25.4 * 100#) / 100) + "�~��"
    Voice "�� " + CStr(Int(Xmm / 72 * 25.4 * 100#) / 100) + "�~��"
ee1: Exit Sub
    

End Sub

'
' ���@�ʒu�ǂ݂Q�@��
'
Sub PosVoice2()
    On Error GoTo ss1
    Dim Page, Keta, Gyo, Sect As Long
    Page = Selection.Information(wdActiveEndPageNumber)
    Keta = Selection.Information(wdFirstCharacterColumnNumber)
    Gyo = Selection.Information(wdFirstCharacterLineNumber)
    Sect = Selection.Information(wdActiveEndSectionNumber)
    If Sect = 1 Then
        Voice CStr(Page) + "�� " + CStr(Gyo) + "�s " + CStr(Keta) + "��"
    Else
        Voice CStr(Page) + "�� " + CStr(Sect) + "�Z�N�V���� " + CStr(Gyo) + "�s " + CStr(Keta) + "��"
    End If
ss1:
End Sub

'
' ���S���ǂ݁�
'
Sub AllRead()
    On Error GoTo ee1
'
    Dim menuObj As Object
    MenuYomiAdd2
    DoEvents
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
         If menuObj.Caption <> "�ǂ�(&R)" Then
            menuObj.Enabled = False
            DoEvents
         End If
    Next
    DoEvents
'
    Quick
    AllReadFlag = True
    Dim a As String
    Dim aa As String
    Dim StartPos As Long
    Dim EndPos As Long
    Dim dum As Long
    Dim b As Long
    Dim WaitTmg As Boolean
    Dim SelNext As Boolean
    Dim SelEnd As Boolean
    Dim MyObj As Object
    Dim SujiFlag As Boolean
    Dim YomiFlag As Boolean
    Set MyObj = Application.Selection
    aa = ""
    a = Selection.Text
    StartPos = MyObj.Start
    b = True
    With Selection
    While b <> 0 And AllReadFlag = True
        WaitTmg = False
        SelNext = False
        YomiFlag = False
            '�s���̃`�F�b�N
        If Selection.Paragraphs(1).Range.ListParagraphs.Count > 0 Then
            If (Len(Selection.Paragraphs(1).Range.ListFormat.ListString) + 2) _
            = Selection.Information(wdFirstCharacterColumnNumber) Then
                Voice Selection.Paragraphs(1).Range.ListFormat.ListString + " "
                VoiceWait
            End If
        End If
'        Call ExBeep(2, 10, 0)

        
        
        If (a = "," Or a = "�C" Or a = "." Or a = "�D") And SujiFlag = False Then
            YomiFlag = True
            aa = aa + a
            b = .MoveRight
            WaitTmg = True
            
            EndPos = MyObj.Start
            .Start = StartPos
            .End = EndPos
            Voice aa
            VoiceWait
            .Start = EndPos
            .End = EndPos
            StartPos = EndPos
            aa = ""
        ElseIf a = "�A" Or a = "�B" Or a = "��" Or a = "�" Or a = "�" _
         Or a = "?" Or a = "�H" Or a = "�" Or a = "�E" Then
            YomiFlag = True
            aa = aa + a
            b = .MoveRight
            WaitTmg = True
            
            EndPos = MyObj.Start
            .Start = StartPos
            .End = EndPos
            Voice aa
            VoiceWait
            .Start = EndPos
            .End = EndPos
            StartPos = EndPos
            aa = ""
        ElseIf a = " " Or a = "�@" Then
            YomiFlag = True
            aa = aa + a
            EndPos = MyObj.Start
            .Start = StartPos
            .End = EndPos

            Voice aa
            VoiceWait
            SelEnd = True
            .Start = EndPos
            .End = EndPos
            aa = ""
            StartPos = EndPos
            SelNext = True
        ElseIf InStr(a, Chr$(13)) Then
            YomiFlag = True
            aa = aa + a
          
            EndPos = Selection.Start
            .Start = StartPos
            .End = EndPos
 
            Voice aa
            VoiceWait
            .Start = EndPos
            .End = EndPos
            SelEnd = True
            aa = ""
            StartPos = EndPos
            SelNext = True
        Else
            aa = aa + a
        End If
        If WaitTmg = False Then
            b = MyObj.MoveRight
        End If
        If SelNext = True Then
            StartPos = .Start
        End If
        If Val(a) > 0 Or a = "0" Or a = "�O" Or (a >= "�P" And a <= "�X") Then
            SujiFlag = True
        Else
            SujiFlag = False
        End If
        a = Selection.Text
        DoEvents
        
    Wend
    End With
    Quick
    If aa <> "" Then
        Voice aa
    End If
    Voice "�ǂ� �I���"
    VoiceWait
    MyObj.Collapse (wdCollapseEnd)
    AllReadFlag = False
ee1:
    MenuYomiAdd
    DoEvents
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
'        If menuObj.Caption = "�ǂ�(&R)" Then
            menuObj.Enabled = True
'        End If
    Next

    Exit Sub
End Sub
Sub curPoint()
    On Error GoTo ss1
    Selection.MoveRight
    SVoice ThisDocument.Application.Selection.Text
ss1:
End Sub
'
'�� �G�X�P�[�v�L�[ ��
'
Sub Escape()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Selection.EscapeKey
    AllReadFlag = False
End Sub
'
'�� �C���T�[�g�L�[ ��
'
Sub InsDel()
    On Error GoTo ee1
    Quick
    Options.Overtype = Not Options.Overtype
    If Options.Overtype = True Then
        Voice "�㏑��"
    Else
        Voice "�}��"
    End If
ee1:    Exit Sub
End Sub
'
'�� �f���[�g�L�[ ��
'
Sub DelChar()
    On Error GoTo ee1
    Selection.Delete Unit:=wdCharacter, Count:=1
        Quick
        '���وʒu�������޲��
        If DeleteVoice = False Then
            Voice "�폜 "
        End If
        SVoice Selection.Text
        DeleteVoice = True
        
ee1:        Exit Sub
End Sub
'�� BS�L�[ ��
'
Sub BsChar()
    On Error GoTo ee1
    Dim objRng As Object
    Static sX As String
    Set objRng = Selection.Characters(1)
    dum = objRng.Move(wdCharacter, -2)
    If dum = -1 Then
        Quick
        Voice "�o�b�N �g�b�v"
        Exit Sub
        
    End If
    sX = objRng.Characters(1).Text
    Selection.TypeBackspace
        Quick
        '���وʒu�������޲��
        If BackVoice = False Then
            Voice "�o�b�N "
        End If
        SVoice sX
        BackVoice = True
ee1:    Exit Sub
End Sub
'
'�� CTRL+[Home] ��
'
Sub curTopPage()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Selection.HomeKey Unit:=wdStory
    Voice "�g�b�v�y�[�W"
End Sub
'
'�� CTRL+[END] ��
'
Sub curLastPage()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Selection.EndKey Unit:=wdStory
    Voice "���X�g�y�[�W"
End Sub
'
'�� [PageUp] ��
'
Sub curPageUp()
    Static Sainyu As Long
    If Sainyu = -1 Then
        Exit Sub
    End If
    Sainyu = -1
    Dim i
    On Error GoTo ee1
    Dim Page, Keta, Gyo As Long
    Quick
    Selection.MoveUp Unit:=wdScreen, Count:=1
    Page = Selection.Information(wdActiveEndPageNumber)
    Keta = Selection.Information(wdFirstCharacterColumnNumber)
    Gyo = Selection.Information(wdFirstCharacterLineNumber)
    If Gyo = -1 Then
        Voice "�o�b�N " + CStr(Page) + "�� " + CStr(Keta) + "�� "
    Else
        Voice "�o�b�N " + CStr(Page) + "��" + CStr(Gyo) + "��" + CStr(Keta) + " "
    End If
    For i = 0 To 10: DoEvents: Next
    SVoice Selection.Text

ee1:    Sainyu = 0
        Exit Sub
End Sub

'
'�� [PageDown] ��
'
Sub curPageDown()
    Static Sainyu As Long
    If Sainyu = -1 Then
        Exit Sub
    End If
    Sainyu = -1
    Dim i
    On Error GoTo ee1
    Dim Page, Keta, Gyo As Long
    Quick
    Selection.MoveDown Unit:=wdScreen, Count:=1
    Page = Selection.Information(wdActiveEndPageNumber)
    Keta = Selection.Information(wdFirstCharacterColumnNumber)
    Gyo = Selection.Information(wdFirstCharacterLineNumber)
    If Gyo = -1 Then
        Voice "�l�N�X�g " + CStr(Page) + "�� " + CStr(Keta) + "�� "
    Else
        Voice "�l�N�X�g " + CStr(Page) + "��" + CStr(Gyo) + "��" + CStr(Keta) + " "
    End If
    For i = 0 To 10: DoEvents: Next
    SVoice Selection.Text
ee1: Sainyu = 0

    Exit Sub
End Sub
'
'�� [Home] ��
'
Sub curHome()
    Dim i
    On Error GoTo ee1
    Quick
    Selection.HomeKey Unit:=wdLine
    If GetTateYoko() = 0 Then
        Voice "�Ђ���[ "
    End If
    For i = 0 To 10: DoEvents: Next
    SVoice Selection.Text
ee1:
    Exit Sub
End Sub
'
'�� [End] ��
'
Sub curEnd()
    Dim i
    On Error GoTo ee1
    Quick
    Selection.EndKey Unit:=wdLine
    If GetTateYoko() = 0 Then
        Voice "�݂��[ "
    End If
    For i = 0 To 10: DoEvents: Next
    SVoice Selection.Text

ee1:
    Exit Sub
End Sub
'
' ���I���ړ��ǂ݁�
'
Private Sub SelMoveVoice(st0 As Long, en0 As Long)

    Dim Xlen As Long
    If Selection.Start <> st0 Then
        Xlen = st0 - Selection.Start
        Quick         '���وʒu�������޲��
        If Xlen <= 1 Then
            SVoice Left$(ThisDocument.Application.Selection.Text, 1)
        Else
            Voice Left$(ThisDocument.Application.Selection.Text, Xlen)
        End If
    ElseIf Selection.End <> en0 Then
        Xlen = Selection.End - en0
        Quick         '���وʒu�������޲��
        If Xlen <= 1 Then
            SVoice Right$(ThisDocument.Application.Selection.Text, 1)
        Else
            Voice Right$(ThisDocument.Application.Selection.Text, Xlen)
        End If
    Else
        Quick
        Voice "�I�������"
    End If
End Sub


'
'�� ���ق��E�Ɉړ� ��
'
Sub curRight()
    '���ق��E�Ɉړ�
    If Documents.Count = 0 Then
        Exit Sub
    End If
    LineVoiceFlag = False
    LeftVoiceFlag = False
    RightVoiceFlag = False
    Select Case GetTateYoko()
        Case 0
            Selection.MoveRight
            RightVoiceFlag = True
        Case 10
            Selection.MoveUp
            LineVoiceFlag = True

        Case 1
            Selection.MoveDown
            LineVoiceFlag = True
        Case 2
            Selection.MoveUp
            LineVoiceFlag = True
        Case 11
            Selection.MoveRight
    End Select
    Quick
    If RightVoiceFlag = True Then
        RightGuide
    End If
    '���وʒu�������޲��
    If LineVoiceFlag = False Then
        SVoice ThisDocument.Application.Selection.Text
    Else
        LineReadRight
    End If
    dum = GetTateYoko()
End Sub
'
'�� Shift+���ق��E�Ɉړ� ��
'
Sub SHcurRight()
    Dim st0 As Long
    Dim en0 As Long
    If Documents.Count = 0 Then
        Exit Sub
    End If
    st0 = Selection.Start
    en0 = Selection.End
    Select Case GetTateYoko()
        Case 0
            Selection.MoveRight , , wdExtend
        Case 10
            Selection.MoveUp , , wdExtend
        Case 1
            Selection.MoveDown , , wdExtend
        Case 2
            Selection.MoveUp , , wdExtend
        Case 11
            Selection.MoveRight , , wdExtend
    End Select
    SelMoveVoice st0, en0
End Sub

'�� �I��͈͂�ǂ� ��
Sub SelectionVoice()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    SelVoice ThisDocument.Application.Selection.Text

End Sub
Private Sub SelVoice(str As String)
    Dim ss1 As String
    Dim ss2 As String
    
    If Len(str) > 60 Then
        ss1 = Mid$(str, 1, 30)
        While InStr(ss1, Chr$(13)) > 5
            ss1 = Mid$(ss1, 1, InStr(ss1, Chr$(13)) - 1)
        Wend
        While InStr(ss1, "�A") > 5
            ss1 = Mid$(ss1, 1, InStr(ss1, "�A") - 1)
        Wend

        ss2 = Mid$(str, Len(str) - 30)
        While (Len(ss2) - InStr(ss2, Chr$(13))) > 5
            If InStr(ss2, Chr$(13)) = 0 Then GoTo ww1
            ss2 = Mid$(ss2, InStr(ss2, Chr$(13)) + 1)
        Wend
ww1:
        While (Len(ss2) - InStr(ss2, "�A")) > 5
            If InStr(ss2, "�A") = 0 Then GoTo ww2
            ss2 = Mid$(ss2, InStr(ss2, "�A") + 1)
        Wend
ww2:
        Voice ss1 + " ���� " + ss2 + " �܂Ł@" + CStr(Len(str)) + "������I��"
    Else
        Voice str
    End If
 End Sub


'�� ���ق����Ɉړ� ��
Sub curLeft()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    '���ق����Ɉړ�
    LineVoiceFlag = False
    LeftVoiceFlag = False
    RightVoiceFlag = False
    Select Case GetTateYoko()
        Case 0
            Selection.MoveLeft
            LeftVoiceFlag = True
        Case 10
            Selection.MoveDown
            LineVoiceFlag = True
        Case 1
            Selection.MoveUp
            LineVoiceFlag = True
        Case 2
            Selection.MoveDown
            LineVoiceFlag = True
        Case 11
            Selection.MoveLeft
    End Select
    Quick
    If LeftVoiceFlag = True Then
        LeftGuide
    End If
    If LineVoiceFlag = False Then
        SVoice ThisDocument.Application.Selection.Text
    Else
        LineReadRight
    End If
    dum = GetTateYoko()
End Sub
'
'�� Shift+���ق����Ɉړ� ��
'
Sub SHcurLeft()
    Dim st0 As Long
    Dim en0 As Long
    If Documents.Count = 0 Then
        Exit Sub
    End If
    st0 = Selection.Start
    en0 = Selection.End
    Select Case GetTateYoko()
        Case 0
            Selection.MoveLeft , , wdExtend
        Case 10
            Selection.MoveDown , , wdExtend
        Case 1
            Selection.MoveUp , , wdExtend
        Case 2
            Selection.MoveDown , , wdExtend
        Case 11
            Selection.MoveLeft , , wdExtend
    End Select
    SelMoveVoice st0, en0

End Sub
'�� ���ق���Ɉړ� ��
Sub curUp()
    
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Static Sainyu As Long
    If Sainyu <> 0 Then
        Quick
        Debug.Print "�ē�"
        Exit Sub
    End If
    Sainyu = -1
    LineVoiceFlag = False
    TopVoiceFlag = False
    '���ق���Ɉړ�
    dum = Selection.Start
    Select Case GetTateYoko()
        Case 0
            Selection.MoveUp
            LineVoiceFlag = True
            TopVoiceFlag = True
        Case 10
            Selection.MoveLeft
        Case 1
            Selection.MoveRight
        Case 2
            Selection.MoveLeft
        Case 11
            Selection.MoveUp
            LineVoiceFlag = True
    End Select
    '���وʒu�������޲��
    Quick
    If TopVoiceFlag = True Then
        TopGuide
    End If
    If dum = Selection.Start Then
        Voice "�g�b�v"
        VoiceWait
         Quick
    End If
    If LineVoiceFlag = False Then
        SVoice ThisDocument.Application.Selection.Text
    Else
        LineReadRight
    End If
    dum = GetTateYoko()
    Sainyu = 0
End Sub
'
'�� Shift+���ق���Ɉړ� ��
'
Sub SHcurUp()
    Dim st0 As Long
    Dim en0 As Long
    If Documents.Count = 0 Then
        Exit Sub
    End If
    st0 = Selection.Start
    en0 = Selection.End
    Select Case GetTateYoko()
        Case 0
            Selection.MoveUp , , wdExtend
        Case 10
            Selection.MoveLeft , , wdExtend
        Case 1
            Selection.MoveRight , , wdExtend
        Case 2
            Selection.MoveLeft , , wdExtend
        Case 11
            Selection.MoveUp , , wdExtend
    End Select
    SelMoveVoice st0, en0


End Sub
'�� ���ق����Ɉړ� ��
Sub curDown()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Static Sainyu As Long
    If Sainyu <> 0 Then
        Exit Sub
    End If
    Sainyu = -1
    '���ق����Ɉړ�
    dum = Selection.Start
    LineVoiceFlag = False
    DownVoiceFlag = False
    Select Case GetTateYoko()
        Case 0
            Selection.MoveDown
            LineVoiceFlag = True
            DownVoiceFlag = True
        Case 10
            Selection.MoveRight
        Case 1
            Selection.MoveLeft
        Case 2
            Selection.MoveRight
        Case 11
            Selection.MoveDown
            LineVoiceFlag = True
    End Select
    '���وʒu�������޲��
    Quick
    If DownVoiceFlag = True Then
        DownGuide
    End If
    If dum = Selection.Start Then
        Voice "�����[ ���X�g"
        VoiceWait
         Quick
    End If
    If LineVoiceFlag = False Then
        SVoice ThisDocument.Application.Selection.Text
    Else
        LineReadRight
    End If
    dum = GetTateYoko()
    Sainyu = 0
End Sub
'
'�� Shift+���ق����Ɉړ� ��
'
Sub SHcurDown()
    Dim st0 As Long
    Dim en0 As Long
    If Documents.Count = 0 Then
        Exit Sub
    End If
    st0 = Selection.Start
    en0 = Selection.End
    Select Case GetTateYoko()
        Case 0
            Selection.MoveDown , , wdExtend
        Case 10
            Selection.MoveRight , , wdExtend
        Case 1
            Selection.MoveLeft , , wdExtend
        Case 2
            Selection.MoveRight , , wdExtend
        Case 11
            Selection.MoveDown , , wdExtend
    End Select
    SelMoveVoice st0, en0

End Sub
Private Sub TopGuide()
    If Selection.Information(wdFirstCharacterLineNumber) = 1 Then
           Quick
           Voice ("�����[" + Chr$(&HFE))
    End If
    
End Sub
Private Sub DownGuide()
    On Error GoTo ee1
    Static Pos1, Pos2 As Long
    Pos1 = Selection.Bookmarks.Item("\Page").End
    Pos2 = Selection.Bookmarks.Item("\Line").End
    If Pos1 <= Pos2 Then
           Quick
           Voice ("�����[" + Chr$(&HFE))
    End If
ee1:    Exit Sub
End Sub

Private Sub LeftGuide()
    If Selection.Information(wdFirstCharacterColumnNumber) = 1 Then
        If Left$(Selection.Text, 1) <> Chr$(13) Then
           Quick
           Voice ("�Ђ���[" + Chr$(&HFE))
        End If
    End If
    
End Sub
Private Sub RightGuide()
    On Error GoTo ee1
    dum = Len(Selection.Paragraphs(1).Range.Bookmarks.Item("\Line").Range.Text)
    If Selection.Information(wdFirstCharacterColumnNumber) = dum Then
        If Left$(Selection.Text, 1) <> Chr$(13) Then
           VoiceWait
           Quick
            Voice ("�݂��[" + Chr$(&HFE))
        End If
    End If
ee1:    Exit Sub
    
End Sub
'�� �͈͓ǂݻ�� ��
Private Sub RangeGo(ByVal st1 As Long, ByVal en1 As Long)
    On Error GoTo ee1
    If st1 >= en1 Then
        Voice "���ޮ�"
        Exit Sub
    End If
    Voice Mid$(Application.Selection.ShapeRange.TextFrame.TextRange.Text, st1 + 1, en1 - st1)
    Exit Sub
ee1: Resume ee2
ee2:
    Voice Application.ActiveDocument.Range(st1, en1).Text

End Sub
'
' ���@�i���ǂ݁@��
'
Sub ParaRead()
    On Error GoTo ee1
    Quick
    Voice Selection.Paragraphs(1).Range.Text
ee1:    Exit Sub
End Sub

'�@���P�s�ǂ݁�
Sub LineReadAll()
    On Error GoTo ee1
    dum = Selection.Start
    Quick
    On Error GoTo ss1
    Dim ss, x, y As Long
    Voice KajoStr + Selection.Paragraphs(1).Range.Bookmarks.Item("\Line").Range.Text
    Exit Sub
ss1: ss = Selection.Start
    Selection.EndKey Unit:=wdLine
    DoEvents
    x = Selection.Start

    Selection.HomeKey Unit:=wdLine

    DoEvents
    y = Selection.Start
    Selection.End = 0
    Selection.Start = ss

    RangeGo y, x
    
ee1:    Exit Sub

End Sub
'�@�����ق܂łP�s�ǂ݁�
Sub LineReadLeft()
    On Error GoTo ee1
    dum = Selection.Start
   Dim ss, x, y As Long
    On Error GoTo ss1
    x = Selection.Information(wdFirstCharacterColumnNumber)
    If x = 1 Then
        Voice "�s��"
    Else
        Voice Mid$(KajoStr + Selection.Paragraphs(1).Range.Bookmarks.Item("\Line").Range.Text, 1, x - 1)
    End If
    Exit Sub
ss1:
    ss = Selection.Start
    Selection.HomeKey Unit:=wdLine
    DoEvents
    x = Selection.Start

       Selection.End = 0
    Selection.Start = ss
    RangeGo x, ss
    
ee1:   Exit Sub
End Sub
'�@�����وȍ~�P�s�ǂ݁�
Sub LineReadRight()
    On Error GoTo ee1:
    Static ss, x, y As Long
    If Left$(Selection.Text, 1) = Chr$(13) Then
        Voice "���s"
        Exit Sub
    End If
    If Left$(Selection.Text, 1) = Chr$(11) Then
        Voice "�s����"
        Exit Sub
    End If
    On Error GoTo ss1
    x = Selection.Information(wdFirstCharacterColumnNumber)
    Voice Mid$(KajoStr + Selection.Paragraphs(1).Range.Bookmarks.Item("\Line").Range.Text, x)
    Exit Sub
ss1:
    
    
    ss = Selection.Start
    Selection.EndKey Unit:=wdLine
    DoEvents
    x = Selection.Start
    If ss = x Then
            Voice "�s��"
        Exit Sub
    End If
    
    Selection.HomeKey Unit:=wdLine
    DoEvents
    y = Selection.Start
    If y <> ss Then
        Selection.End = 0
        DoEvents
        Selection.Start = ss
        DoEvents
    End If
    If ss >= x Then
        If ss = y Then
            Voice "��s"
        Else
            Voice "�s��"
        End If
    Else
        RangeGo ss, x
    
    End If
    
ee1:    Exit Sub

End Sub
'�������ǂ݁�
Sub FormInfo()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Quick
    Select Case Selection.Paragraphs(1).Alignment
    
    Case wdAlignParagraphDistribute
        Voice "�ϓ� "
    Case wdAlignParagraphRight
        Voice "�E�l "
    Case wdAlignParagraphLeft
        Voice "���l "
    Case wdAlignParagraphCenter
        Voice "�Z���^�����O "
    Case wdAlignParagraphJustify
        Voice "���l "
    End Select
    Select Case Selection.Paragraphs(1).LineSpacingRule
    
    Case wdLineSpaceSingle
        Voice "�s�ԂP�s "
    Case wdLineSpace1pt5
        Voice "�s�Ԃ����Ă�T�s "
    Case wdLineSpaceDouble
        Voice "�s�ԂQ�s "
    End Select
    If Selection.Paragraphs(1).FirstLineIndent + Selection.Paragraphs(1).LeftIndent <> 0 Then
        If Selection.Paragraphs(1).FirstLineIndent <> 0 Then
            Voice "�P�s�ڂ̃C���f���g" + CStr(Selection.Paragraphs(1).FirstLineIndent _
            + Selection.Paragraphs(1).LeftIndent) + "�|�C���g"
        End If
    End If
    If Selection.Paragraphs(1).LeftIndent <> 0 Then
        If Selection.Paragraphs(1).FirstLineIndent < 0 Then
            Voice "�Ԃ炳���C���f���g" + CStr(Selection.Paragraphs(1).LeftIndent) + "�|�C���g"
        Else
            Voice "���C���f���g" + CStr(Selection.Paragraphs(1).LeftIndent) + "�|�C���g"
        End If
    End If
    If Selection.Paragraphs(1).RightIndent <> 0 Then
        Voice "�E�C���f���g" + CStr(Selection.Paragraphs(1).RightIndent) + "�|�C���g"
       
    End If
    If KajoStr() <> "" Then
        Voice "�ӏ����� " + KajoStr()
    End If
     
End Sub



'�����ǂ݁�
Sub CurInfo()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    '    TableInfo
    CharInfo
    ShapeInfo
End Sub
'�� �I�u�W�F�N�g����ǂ� ��
Public Sub ShapeInfo()
    On Error GoTo ee1

    Voice "�I�u�W�F�N�g��" + Chr$(&HFE) + Application.Selection.ShapeRange.Name + Chr$(&HFE)

    Exit Sub
ee1:    Voice "�}�`�ȊO�̃I�u�W�F�N�g"
        Exit Sub
End Sub
'�� �r������ǂ� ��
Sub TableInfo()
    On Error GoTo ee1
    If keisen2 = True Then
    With Application.Selection.Cells
        Voice "�r���@����" + Chr$(&HFE)
        keisen (wdBorderTop)
        Voice "�r���@����" + Chr$(&HFE)
        keisen (wdBorderBottom)
        Voice "�r���@�Ђ���" + Chr$(&HFE)
        keisen (wdBorderLeft)
        Voice "�r���@�݂�" + Chr$(&HFE)
        keisen (wdBorderRight)
        
    End With
    Else
        Voice "�r�� �Ȃ�"
    End If
ee1:    Exit Sub
End Sub
'�� �r������ǂށE���2 ��
Private Function keisen2() As Boolean
    Dim Retsu, Gyo
    keisen2 = False
    On Error GoTo ee1
    Retsu = Selection.Information(wdStartOfRangeColumnNumber)
    Gyo = Selection.Information(wdStartOfRangeRowNumber)
    If Gyo > 0 And Retsu > 0 Then
        Voice "�Z�� " + CStr(Gyo) + "�s" + CStr(Retsu) + "�� "
        keisen2 = True
    End If
ee1: Exit Function
End Function
'�� �r������ǂށE��� ��
Private Sub keisen(s As Long)
    On Error GoTo ee1
    With Application.Selection.Cells

    If .Borders(s).Visible = False Then
        Voice "�Ȃ�"
    Else
        Select Case .Borders(s).LineStyle
        Case wdLineStyleNone
            Voice "�r�� �Ȃ� "
        Case wdLineStyleSingle
            Voice "�ʏ�r�� "
        Case wdLineStyleDot
            Voice "�_�� "
        Case wdLineStyleDashSmallGap
            Voice "�j�� �� "
        Case wdLineStyleDashLargeGap
            Voice "�j�� �� "
        Case wdLineStyleDashDot
            Voice "�P�_���� "
        Case wdLineStyleDashDotDot
            Voice "�Q�_���� "
        Case wdLineStyleDouble
            Voice "�Q�d�r�� "
        Case wdLineStyleTriple
            Voice "�R�d�r�� "
        Case Else
 
        Voice "����" + CStr(.Borders(s).LineStyle) + Chr$(&HFE)
        End Select
        Select Case .Borders(s).LineWidth
            Case wdLineWidth025pt
                Voice "0.25�|�C���g" + Chr$(&HFE)
            Case wdLineWidth050pt
                Voice "0.5�|�C���g" + Chr$(&HFE)
            Case wdLineWidth075pt
                Voice "0.75�|�C���g" + Chr$(&HFE)
            Case wdLineWidth100pt
                Voice "1�|�C���g" + Chr$(&HFE)
            Case wdLineWidth150pt
                Voice "1.5�|�C���g" + Chr$(&HFE)
            Case wdLineWidth225pt
                Voice "2.25�|�C���g" + Chr$(&HFE)
            Case wdLineWidth300pt
                Voice "3�|�C���g" + Chr$(&HFE)
            Case wdLineWidth450pt
                Voice "4.5�|�C���g" + Chr$(&HFE)
            Case wdLineWidth600pt
                Voice "6�|�C���g" + Chr$(&HFE)
        End Select
        Select Case .Borders(s).ColorIndex
            Case wdAuto
                Voice "����" + Chr$(&HFE)
            Case wdBlack
                Voice "�۲�" + Chr$(&HFE)
            Case wdBlue
                Voice "�F" + Chr$(&HFE)
            Case wdBrightGreen
                Voice "�Z���ΐF" + Chr$(&HFE)
            Case wdByAuthor
                Voice "ByAuthor" + Chr$(&HFE)
            Case wdDarkBlue
                Voice "�Z���F" + Chr$(&HFE)
            Case wdDarkRed
                Voice "�Z���ԐF" + Chr$(&HFE)
            Case wdDarkYellow
                Voice "�Z�����F" + Chr$(&HFE)
            Case wdGray25
                Voice "�Z���D�F" + Chr$(&HFE)
            Case wdGray50
                Voice "�D�F" + Chr$(&HFE)
            Case wdGreen
                Voice "�ΐF" + Chr$(&HFE)
            Case wdNoHighlight
                Voice "NoHighlight" + Chr$(&HFE)
            Case wdYellow
                Voice "���F" + Chr$(&HFE)
            Case wdWhite
                Voice "���F" + Chr$(&HFE)
            Case wdViolet
                Voice "���F" + Chr$(&HFE)
            Case wdTurquoise
                Voice "���F" + Chr$(&HFE)
            Case wdTeal
                Voice "��" + Chr$(&HFE)
            Case wdRed
                Voice "�ԐF" + Chr$(&HFE)
            Case wdPink
                Voice "�s���N�F" + Chr$(&HFE)
            
        End Select
    End If
    End With
ee1:    Exit Sub
End Sub
'�� ��������ǂ� ��
Sub CharInfo()
    On Error GoTo eee1
    Dim str As String
    str = Left$(Application.Selection.Text, 1)
    SSVoice str
    VoiceWait
    If Asc(str) < 256 And Asc(str) >= 0 Then
        Voice "���p"
    Else
        Voice "�S�p"
    End If
    Select Case str
        Case "��" To "��"
            Voice "�Ђ炪��" + Chr$(&HFE)
        Case "�" To "�"
            Voice "��������" + Chr$(&HFE)
        Case "�@" To "��"
            Voice "��������" + Chr$(&HFE)
        Case "a" To "z"
            Voice "�p�� ������" + Chr$(&HFE)
        Case "A" To "Z"
            Voice "�p�� �啶��" + Chr$(&HFE)
        Case "�`" To "�y"
            Voice "�p�� �啶��" + Chr$(&HFE)
        Case "��" To "��"
            Voice "�p�� ������" + Chr$(&HFE)
    End Select
    If Application.Selection.Style <> "�W��" Then
        Voice Application.Selection.Style + "�̽���" + Chr$(&HFE)
    End If
    Voice Application.Selection.Font.Name + Chr$(&HFE)
    Voice CStr(Application.Selection.Font.Size) + "�|�C���g" + Chr$(&HFE)
    If Application.Selection.Font.Bold = True Then
        Voice "����" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Underline = True Then
        Voice "���ްײ�" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Italic = True Then
        Voice "�Α�" + Chr$(&HFE)
    End If
        Select Case Application.Selection.Font.ColorIndex
            Case wdAuto
                Voice "����" + Chr$(&HFE)
            Case wdBlack
                Voice "�۲�" + Chr$(&HFE)
                             
           
            Case wdBlue
                Voice "�F" + Chr$(&HFE)
            Case wdBrightGreen
                Voice "�Z���ΐF" + Chr$(&HFE)
            Case wdByAuthor
                Voice "ByAuthor" + Chr$(&HFE)
            Case wdDarkBlue
                Voice "�Z���F" + Chr$(&HFE)
            Case wdDarkRed
                Voice "�Z���ԐF" + Chr$(&HFE)
            Case wdDarkYellow
                Voice "�Z�����F" + Chr$(&HFE)
            Case wdGray25
                Voice "�Z���D�F" + Chr$(&HFE)
            Case wdGray50
                Voice "�D�F" + Chr$(&HFE)
            Case wdGreen
                Voice "�ΐF" + Chr$(&HFE)
            Case wdNoHighlight
                Voice "NoHighlight" + Chr$(&HFE)
            Case wdYellow
                Voice "���F" + Chr$(&HFE)
            Case wdWhite
                Voice "���F" + Chr$(&HFE)
            Case wdViolet
                Voice "���F" + Chr$(&HFE)
            Case wdTurquoise
                Voice "���F" + Chr$(&HFE)
            Case wdTeal
                Voice "��" + Chr$(&HFE)
            Case wdRed
                Voice "�ԐF" + Chr$(&HFE)
            Case wdPink
                Voice "�s���N�F" + Chr$(&HFE)
             
        End Select
    If Application.Selection.Font.DoubleStrikeThrough = True Then
        Voice "��d��������" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Emboss = True Then
        Voice "��������" + Chr$(&HFE)
    End If
    If Application.Selection.Font.EmphasisMark = True Then
        Voice "�T�_" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Engrave = True Then
        Voice "��������" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Hidden = True Then
        Voice "�B������" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Outline = True Then
        Voice "������" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Shadow = True Then
        Voice "�e�t��" + Chr$(&HFE)
    End If
    If Application.Selection.Font.StrikeThrough = True Then
        Voice "�������" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Subscript = True Then
        Voice "���t������" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Superscript = True Then
        Voice "��t������" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Engrave = True Then
        Voice "��������" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Scaling <> 100 Then
        Voice "�����{��" + CStr(Application.Selection.Font.Scaling) + Chr$(&HFE)
    End If
    AmiKake
eee1:    Exit Sub
End Sub

'�� ��������ǂށE��ށ|�Ԋ|�� ��

Private Sub AmiKake()
    On Error GoTo ee1
    If Application.Selection.Font.Shading.Texture <> 0 Then
        Voice "�Ԋ|������" + CStr(Application.Selection.Font.Shading.Texture) + Chr$(&HFE)
    End If
ee1:    Exit Sub
End Sub

'���@���c�������̎擾�@��
'
Private Function GetTateYoko() As Long
    On Error GoTo eeee1
    
    
    Select Case Application.Selection.Orientation
        Case wdTextOrientationHorizontal '�ʏ퉡����
            GetTateYoko = 0
        Case wdTextOrientationVerticalFarEast '�ʏ�c����
           GetTateYoko = 10
        Case wdTextOrientationUpward  '�������@���X�O�x��]�@������͏������
            GetTateYoko = 1
        Case wdTextOrientationDownward  '�������@�E�X�O�x��]�@������͉�������
            GetTateYoko = 2
        Case wdTextOrientationHorizontalRotatedFarEast '�c�����@���X�O�x��]�@������͉E������
            GetTateYoko = 11
    End Select
    If TateYokoMode <> GetTateYoko Then
    Select Case Application.Selection.Orientation
        Case wdTextOrientationHorizontal '�ʏ퉡����
            Voice "���������[�h"
        Case wdTextOrientationVerticalFarEast '�ʏ�c����
            Voice "�c�������[�h"
        Case wdTextOrientationUpward  '�������@���X�O�x��]�@������͏������
            Voice "�������@���X�O�x��]���[�h"
        Case wdTextOrientationDownward  '�������@�E�X�O�x��]�@������͉�������
            Voice "�������@�E�X�O�x��]���[�h"
        Case wdTextOrientationHorizontalRotatedFarEast '�c�����@���X�O�x��]�@������͉E������
            Voice "�c�����@���X�O�x��]���[�h"
    End Select
    End If
    TateYokoMode = GetTateYoko
    If ActiveWindow.View.Type <> wdPrintView Then
        GetTateYoko = 0
    End If
    
    Exit Function
eeee1: GetTateYoko = 0
    Exit Function
End Function


Function KajoStr() As String
    KajoStr = ""
    If Selection.Paragraphs(1).Range.ListParagraphs.Count > 0 Then
        KajoStr = Selection.Paragraphs(1).Range.ListFormat.ListString + " "
        Call ExBeep(1, 10, 0)

    End If
    
End Function

Function Model() As Long
    If System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\KSD\PCTalker\Default", "Model") = "" Then
        Model = 0
    Else
        Model = 1
    End If
End Function
