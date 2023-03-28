Attribute VB_Name = "WordTalker"
'
'ワード読み上げマクロ
'このプログラムは著作権で保護されています。
'（株）高知システム開発
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
'■ キーボードマクロの設定 ■
'
Sub SetKeyEvent()
    Quick
    MenuYomiAdd
    CheckPcTalker
    KeyBindings.Add KeyCode:=BuildKeyCode(38), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curUp" '↑
    KeyBindings.Add KeyCode:=BuildKeyCode(38, wdKeyShift), KeyCategory:=wdKeyCategoryMacro _
        , Command:="SHcurUp" 'S+↑
    KeyBindings.Add KeyCode:=BuildKeyCode(40), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curDown" '↓
    KeyBindings.Add KeyCode:=BuildKeyCode(40, wdKeyShift), KeyCategory:=wdKeyCategoryMacro _
        , Command:="SHcurDown" 'S+↓
    KeyBindings.Add KeyCode:=BuildKeyCode(37), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curLeft" '←
    KeyBindings.Add KeyCode:=BuildKeyCode(37, wdKeyShift), KeyCategory:=wdKeyCategoryMacro _
        , Command:="SHcurLeft" 'S+←
    KeyBindings.Add KeyCode:=BuildKeyCode(39), KeyCategory:=wdKeyCategoryMacro _
        , Command:="curRight"  '→
     KeyBindings.Add KeyCode:=BuildKeyCode(39, wdKeyShift), KeyCategory:=wdKeyCategoryMacro _
        , Command:="SHcurRight"  'S+→
       
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
'■ キーボードマクロのクリア ■
'
Sub ClearKeyEvent()
    Quick
        Dim menuObj As Object

    Voice "読み上げ 停止"
    KeyBindings.ClearAll
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF12, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="SetKeyEvent"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyF11, wdKeyAlt), KeyCategory:= _
        wdKeyCategoryMacro, Command:="ClearKeyEvent"
    CommandBars.ActiveMenuBar.Reset
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
        If menuObj.Caption = "読み(&R)" Then
            menuObj.Delete
            DoEvents
        End If
    Next
'    MenuYomiAdd2
End Sub


'
'■ 読みﾒﾆｭｰを追加する ■
'
Sub MenuYomiAdd()
    '既存のﾒﾆｭｰがあれば削除
    Dim menuObj As Object
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
        If menuObj.Caption = "読み(&R)" Then
            menuObj.Delete
            DoEvents
        End If
    Next
    'ﾒﾆｭｰを追加
    
    Dim myMenuBar, newMenu, ctrl1
    Set myMenuBar = CommandBars.ActiveMenuBar
    Set newMenu = myMenuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True, Before:=2)
    newMenu.Caption = "読み(&R)"
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton, ID:=1)
    With ctrl1
        .Caption = "ｶｰｿﾙ位置から全文読み(&O) Alt+F10"
        .TooltipText = "ｶｰｿﾙ位置から全文読み"
        .Style = msoButtonCaption
        .OnAction = "AllRead"
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "段落読み(&S) Alt+F9"
        .TooltipText = "段落読み"
        .Style = msoButtonCaption
        .OnAction = "ParaRead"
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "１行読み(&L) Alt+F8"
        .TooltipText = "１行読み"
        .Style = msoButtonCaption
        .OnAction = "LineReadAll"
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "ﾌｧｲﾙ名 読み上げ(&F)"
        .TooltipText = "ﾌｧｲﾙ名 読み上げ"
        .Style = msoButtonCaption
        .OnAction = "FileRead"
        .BeginGroup = True
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "ﾍﾟｰｼﾞ設定 読み上げ(&P)"
        .TooltipText = "ﾌｧｲﾙ名の読み上げ"
        .Style = msoButtonCaption
        .OnAction = "PageSetRead"
    End With
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton)
    With ctrl1
        .Caption = "Word読み上げヘルプ(&H)"
        .TooltipText = "ﾜｰﾄﾞ読み上げヘルプ"
        .Style = msoButtonCaption
        .OnAction = "WordReadHelp"
        .BeginGroup = True
    End With
    DoEvents
End Sub
'
'■ 読みﾒﾆｭｰを追加する ■
'
Sub MenuYomiAdd2()
    '既存のﾒﾆｭｰがあれば削除
    Dim menuObj As Object
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
        If menuObj.Caption = "読み(&R)" Then
            menuObj.Delete
        End If
    Next
    'ﾒﾆｭｰを追加
    
    Dim myMenuBar, newMenu, ctrl1
    Set myMenuBar = CommandBars.ActiveMenuBar
    Set newMenu = myMenuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True, Before:=2)
    newMenu.Caption = "読み(&R)"
    Set ctrl1 = newMenu.CommandBar.Controls _
    .Add(Type:=msoControlButton, ID:=1)
    With ctrl1
        .Caption = "読みあげ停止(&C) Esc"
        .TooltipText = "読みあげを停止します"
        .Style = msoButtonCaption
        .OnAction = "Escape"
    End With
    
    DoEvents
End Sub

'
'■ ﾜｰﾄﾞ読み上げヘルプ ■
'
Sub WordReadHelp()
    On Error GoTo ss1:
    Dim i
    Dim objDoc As Object
    For Each objDoc In Windows
        If objDoc.Caption = "Word読み上げヘルプ" Then
            objDoc.Activate
            Quick
            Voice "Word読み上げヘルプ "
            Exit Sub
        End If
    Next
    Set objDoc = Documents.Add
    For i = 0 To 100
        DoEvents
    Next
    Quick
    ActiveDocument.ActiveWindow.Caption = "Word読み上げヘルプ"
    Voice "Word読み上げヘルプ "
    ActiveDocument.Range.Font.Bold = True
    ActiveDocument.Range.Font.Size = 13

    ActiveDocument.Range.InsertAfter "■Word読み上げヘルプ■    "
    ActiveDocument.Range.InsertAfter "使用者名:" + System.PrivateProfileString("", "HKEY_CURRENT_USER\Software\KSD\License", "OwnerString") + vbCr + vbCr
    ActiveDocument.Range.InsertAfter "　　ファイル名読み上げ Alt+F1" + vbCr
    ActiveDocument.Range.InsertAfter "　　ページ設定読み上げ Alt+F3" + vbCr
    ActiveDocument.Range.InsertAfter "　　１行読み Alt+F8" + vbCr
    ActiveDocument.Range.InsertAfter "　　段落読み Alt+F9" + vbCr
    ActiveDocument.Range.InsertAfter "　　カーソル位置から全文読み Alt+F10" + vbCr
    ActiveDocument.Range.InsertAfter "　　通常／拡大切換 F6" + vbCr
    ActiveDocument.Range.InsertAfter "　　カーソル位置の読み上げ F9" + vbCr
    ActiveDocument.Range.InsertAfter "　　罫線情報読み Shift+Alt+K" + vbCr
    If Model = 0 Then
        ActiveDocument.Range.InsertAfter vbCr + "●以下はPC-Talkerのｺﾏﾝﾄﾞと同じ操作のものです。" + vbCr
        ActiveDocument.Range.InsertAfter "　　書式情報読み Ctrl+Alt+G" + vbCr
        ActiveDocument.Range.InsertAfter "　　識別読み Ctrl+Alt+M" + vbCr
        ActiveDocument.Range.InsertAfter "　　行頭からカーソル手前読み Ctrl+Alt+H" + vbCr
        ActiveDocument.Range.InsertAfter "　　JISコード読み　Ctrl+Alt+I" + vbCr
        ActiveDocument.Range.InsertAfter "　　１行読み Ctrl+Alt+J" + vbCr
        ActiveDocument.Range.InsertAfter "　　カーソル位置から行末読み Ctrl+Alt+K" + vbCr
        ActiveDocument.Range.InsertAfter "　　カーソル位置読み Ctrl+Alt+,(コンマ)" + vbCr
        ActiveDocument.Range.InsertAfter "　　カーソル位置から全文読み Ctrl+Alt+A" + vbCr
        ActiveDocument.Range.InsertAfter "　　選択範囲の読みあげ Shift" + vbCr
        ActiveDocument.Range.InsertAfter vbCr + "　　Word読み上げ機能の停止 Alt+F11" + vbCr
        ActiveDocument.Range.InsertAfter "　　Word読み上げ機能の開始 Alt+F12" + vbCr
        ActiveDocument.Range.InsertAfter vbCr
    Else
        ActiveDocument.Range.InsertAfter vbCr + "●以下はVDM100Wのｺﾏﾝﾄﾞと同じ操作のものです。" + vbCr
        ActiveDocument.Range.InsertAfter "　　書式情報読み Ctrl+Alt+-(ﾏｲﾅｽ)" + vbCr
        ActiveDocument.Range.InsertAfter "　　識別読み Ctrl+Alt+H" + vbCr
        ActiveDocument.Range.InsertAfter "　　行頭からカーソル手前読み Ctrl+Alt+J" + vbCr
        ActiveDocument.Range.InsertAfter "　　JISコード読み　Ctrl+Alt+N" + vbCr
        ActiveDocument.Range.InsertAfter "　　１行読み Ctrl+Alt+K" + vbCr
        ActiveDocument.Range.InsertAfter "　　カーソル位置から行末読み Ctrl+Alt+L" + vbCr
        ActiveDocument.Range.InsertAfter "　　カーソル位置読み Ctrl+Alt+,(ｺﾝﾏ)" + vbCr
        ActiveDocument.Range.InsertAfter "　　カーソル位置から全文読み Ctrl+Alt+:(ｺﾛﾝ)" + vbCr
        ActiveDocument.Range.InsertAfter "　　選択範囲の読みあげ Shift" + vbCr
        ActiveDocument.Range.InsertAfter vbCr + "　　Word読み上げ機能の停止 Alt+F11" + vbCr
        ActiveDocument.Range.InsertAfter "　　Word読み上げ機能の開始 Alt+F12" + vbCr
        ActiveDocument.Range.InsertAfter vbCr
    End If
    ActiveDocument.Range.InsertAfter "●Word読みあげ機能について" + vbCr
    ActiveDocument.Range.InsertAfter "制作・著作　㈱高知システム開発" + vbCr
    ActiveDocument.Range.InsertAfter "※本格的な音声ワープロのご利用は、徹底的に操作性を追及して" + vbCr + "さらに使いやすくなった、日本語ワープロ" + vbCr _
    + "「ＭＹＷＯＲＤシリーズ」をぜひご利用ください。" + vbCr
    ActiveDocument.Range.InsertAfter "URL http://www.aok-unet.ocn.ne.jp/text" + vbCr
ss1:
End Sub
'
'■ ﾌｧｲﾙ名読み上げ ■
'
Sub FileRead()
    On Error GoTo ee1
    Quick
    Voice ActiveDocument.Name
ee1: Exit Sub



End Sub


'
'■ ﾍﾟｰｼﾞ設定読み上げ ■
'
Sub PageSetRead()
    On Error GoTo ee1
    Quick
'
'　用紙
    Select Case ActiveDocument.PageSetup.PaperSize

    Case wdPaperA3
        Voice "サイズ A3"
    Case wdPaperA4
        Voice "サイズ A4"
    Case wdPaperA4Small
        Voice "サイズ A4 スモール"
    Case wdPaperA5
        Voice "サイズ A5"
    Case wdPaperB4
        Voice "サイズ B4"
    Case wdPaperB5
        Voice "サイズ B5"
    Case wdPaperLetter
        Voice "サイズ　レター"
    Case Else
        Voice "用紙ｽﾀｲﾙ" + CStr(ActiveDocument.PageSetup.PaperSize)
    End Select
'
' 置き方
    If ActiveDocument.PageSetup.Orientation = wdOrientLandscape Then
        Voice " 横方向"
    Else
        Voice " 縦方向"
    End If
'
' 行・文字数
    Voice " " + CStr(ActiveDocument.PageSetup.LinesPage) + "行 " + _
        CStr(ActiveDocument.PageSetup.CharsLine) + "桁 "
'
'　書き方
    If ActiveDocument.Range.Orientation = wdTextOrientationHorizontal Then
        Voice " 横書き"
    ElseIf ActiveDocument.Range.Orientation = wdTextOrientationVerticalFarEast Then
        Voice " 縦書き"
    End If
    Voice "プリンタ名 " + Application.ActivePrinter
ee1:
End Sub

'■ 通常／拡大切換 ■
'
Sub KakudaiChange()
    On Error GoTo ss1
    Dim ss As Long
    ss = Selection.Start
        If Options.BlueScreen = False Then
            Quick
            Voice "拡大画面"
            Options.BlueScreen = True
            ActiveWindow.ActivePane.View.Zoom.Percentage = 500
            dum = Selection.MoveLeft
            If ss <> Selection.Start Then
                dum = Selection.MoveRight
            End If
        Else
            Quick
            Voice "通常画面"
            Options.BlueScreen = False
            ActiveWindow.ActivePane.View.Zoom.Percentage = 100
        
        End If
ss1:
End Sub



'
'■　ワード読み上げイニシャライズ　■
'
Sub Vinit()
'    If OnPctkF = True Then
'        Exit Sub
'    End If
    CheckPcTalker
End Sub
'
'■　ﾜｰﾄﾞ読み上げ処理の初期化　■
'
Sub CheckPcTalker()

'Application ｲﾍﾞﾝﾄﾄﾗｯﾌﾟの設定
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
'        'Voice "音声開始"
'    End If
'    Exit Sub
'
'aa1:
'    OnPctkF = False
    Exit Sub
ee1: MsgBox ("Word読み上げに必要なプログラムが見つかりません")


End Sub
'
' ■　JISコード読み　■
'
Sub JisVoice()
    On Error GoTo ee1
    Dim s As String
    Dim i As Integer
    s = Selection.Text
    s = Right$("000" + Hex$(Asc(s)), 4)
    Quick
    Voice "シフトジスコード "
    For i = 1 To 4: SVoice (Mid$(s, i, 1)): VoiceWait: Next
    
ee1: Exit Sub
End Sub

'
' ■　位置読み　■
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
            Voice CStr(Page) + "頁 " + CStr(Keta) + "桁" + Chr$(&HFE)
        Else
            Voice CStr(Page) + "頁 " + CStr(Sect) + "セクション " + CStr(Keta) + "桁" + Chr$(&HFE)
        End If
    Else
        If Sect = 1 Then
            Voice CStr(Page) + "頁 " + CStr(Gyo) + "行 " + CStr(Keta) + "桁" + Chr$(&HFE)
        Else
            Voice CStr(Page) + "頁 " + CStr(Sect) + "セクション " + CStr(Gyo) + "行 " + CStr(Keta) + "桁" + Chr$(&HFE)
        End If
    End If
    Voice "縦 " + CStr(Int(Ymm / 72 * 25.4 * 100#) / 100) + "ミリ"
    Voice "横 " + CStr(Int(Xmm / 72 * 25.4 * 100#) / 100) + "ミリ"
ee1: Exit Sub
    

End Sub

'
' ■　位置読み２　■
'
Sub PosVoice2()
    On Error GoTo ss1
    Dim Page, Keta, Gyo, Sect As Long
    Page = Selection.Information(wdActiveEndPageNumber)
    Keta = Selection.Information(wdFirstCharacterColumnNumber)
    Gyo = Selection.Information(wdFirstCharacterLineNumber)
    Sect = Selection.Information(wdActiveEndSectionNumber)
    If Sect = 1 Then
        Voice CStr(Page) + "頁 " + CStr(Gyo) + "行 " + CStr(Keta) + "桁"
    Else
        Voice CStr(Page) + "頁 " + CStr(Sect) + "セクション " + CStr(Gyo) + "行 " + CStr(Keta) + "桁"
    End If
ss1:
End Sub

'
' ■全文読み■
'
Sub AllRead()
    On Error GoTo ee1
'
    Dim menuObj As Object
    MenuYomiAdd2
    DoEvents
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
         If menuObj.Caption <> "読み(&R)" Then
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
            '行頭のチェック
        If Selection.Paragraphs(1).Range.ListParagraphs.Count > 0 Then
            If (Len(Selection.Paragraphs(1).Range.ListFormat.ListString) + 2) _
            = Selection.Information(wdFirstCharacterColumnNumber) Then
                Voice Selection.Paragraphs(1).Range.ListFormat.ListString + " "
                VoiceWait
            End If
        End If
'        Call ExBeep(2, 10, 0)

        
        
        If (a = "," Or a = "，" Or a = "." Or a = "．") And SujiFlag = False Then
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
        ElseIf a = "、" Or a = "。" Or a = "を" Or a = "､" Or a = "｡" _
         Or a = "?" Or a = "？" Or a = "･" Or a = "・" Then
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
        ElseIf a = " " Or a = "　" Then
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
        If Val(a) > 0 Or a = "0" Or a = "０" Or (a >= "１" And a <= "９") Then
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
    Voice "読み 終わり"
    VoiceWait
    MyObj.Collapse (wdCollapseEnd)
    AllReadFlag = False
ee1:
    MenuYomiAdd
    DoEvents
    For Each menuObj In CommandBars.ActiveMenuBar.Controls
'        If menuObj.Caption = "読み(&R)" Then
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
'■ エスケープキー ■
'
Sub Escape()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Selection.EscapeKey
    AllReadFlag = False
End Sub
'
'■ インサートキー ■
'
Sub InsDel()
    On Error GoTo ee1
    Quick
    Options.Overtype = Not Options.Overtype
    If Options.Overtype = True Then
        Voice "上書き"
    Else
        Voice "挿入"
    End If
ee1:    Exit Sub
End Sub
'
'■ デリートキー ■
'
Sub DelChar()
    On Error GoTo ee1
    Selection.Delete Unit:=wdCharacter, Count:=1
        Quick
        'ｶｰｿﾙ位置を音声ｶﾞｲﾄﾞ
        If DeleteVoice = False Then
            Voice "削除 "
        End If
        SVoice Selection.Text
        DeleteVoice = True
        
ee1:        Exit Sub
End Sub
'■ BSキー ■
'
Sub BsChar()
    On Error GoTo ee1
    Dim objRng As Object
    Static sX As String
    Set objRng = Selection.Characters(1)
    dum = objRng.Move(wdCharacter, -2)
    If dum = -1 Then
        Quick
        Voice "バック トップ"
        Exit Sub
        
    End If
    sX = objRng.Characters(1).Text
    Selection.TypeBackspace
        Quick
        'ｶｰｿﾙ位置を音声ｶﾞｲﾄﾞ
        If BackVoice = False Then
            Voice "バック "
        End If
        SVoice sX
        BackVoice = True
ee1:    Exit Sub
End Sub
'
'■ CTRL+[Home] ■
'
Sub curTopPage()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Selection.HomeKey Unit:=wdStory
    Voice "トップページ"
End Sub
'
'■ CTRL+[END] ■
'
Sub curLastPage()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Selection.EndKey Unit:=wdStory
    Voice "ラストページ"
End Sub
'
'■ [PageUp] ■
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
        Voice "バック " + CStr(Page) + "頁 " + CStr(Keta) + "桁 "
    Else
        Voice "バック " + CStr(Page) + "の" + CStr(Gyo) + "の" + CStr(Keta) + " "
    End If
    For i = 0 To 10: DoEvents: Next
    SVoice Selection.Text

ee1:    Sainyu = 0
        Exit Sub
End Sub

'
'■ [PageDown] ■
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
        Voice "ネクスト " + CStr(Page) + "頁 " + CStr(Keta) + "桁 "
    Else
        Voice "ネクスト " + CStr(Page) + "の" + CStr(Gyo) + "の" + CStr(Keta) + " "
    End If
    For i = 0 To 10: DoEvents: Next
    SVoice Selection.Text
ee1: Sainyu = 0

    Exit Sub
End Sub
'
'■ [Home] ■
'
Sub curHome()
    Dim i
    On Error GoTo ee1
    Quick
    Selection.HomeKey Unit:=wdLine
    If GetTateYoko() = 0 Then
        Voice "ひだり端 "
    End If
    For i = 0 To 10: DoEvents: Next
    SVoice Selection.Text
ee1:
    Exit Sub
End Sub
'
'■ [End] ■
'
Sub curEnd()
    Dim i
    On Error GoTo ee1
    Quick
    Selection.EndKey Unit:=wdLine
    If GetTateYoko() = 0 Then
        Voice "みぎ端 "
    End If
    For i = 0 To 10: DoEvents: Next
    SVoice Selection.Text

ee1:
    Exit Sub
End Sub
'
' ■選択移動読み■
'
Private Sub SelMoveVoice(st0 As Long, en0 As Long)

    Dim Xlen As Long
    If Selection.Start <> st0 Then
        Xlen = st0 - Selection.Start
        Quick         'ｶｰｿﾙ位置を音声ｶﾞｲﾄﾞ
        If Xlen <= 1 Then
            SVoice Left$(ThisDocument.Application.Selection.Text, 1)
        Else
            Voice Left$(ThisDocument.Application.Selection.Text, Xlen)
        End If
    ElseIf Selection.End <> en0 Then
        Xlen = Selection.End - en0
        Quick         'ｶｰｿﾙ位置を音声ｶﾞｲﾄﾞ
        If Xlen <= 1 Then
            SVoice Right$(ThisDocument.Application.Selection.Text, 1)
        Else
            Voice Right$(ThisDocument.Application.Selection.Text, Xlen)
        End If
    Else
        Quick
        Voice "選択おわり"
    End If
End Sub


'
'■ ｶｰｿﾙを右に移動 ■
'
Sub curRight()
    'ｶｰｿﾙを右に移動
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
    'ｶｰｿﾙ位置を音声ｶﾞｲﾄﾞ
    If LineVoiceFlag = False Then
        SVoice ThisDocument.Application.Selection.Text
    Else
        LineReadRight
    End If
    dum = GetTateYoko()
End Sub
'
'■ Shift+ｶｰｿﾙを右に移動 ■
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

'■ 選択範囲を読む ■
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
        While InStr(ss1, "、") > 5
            ss1 = Mid$(ss1, 1, InStr(ss1, "、") - 1)
        Wend

        ss2 = Mid$(str, Len(str) - 30)
        While (Len(ss2) - InStr(ss2, Chr$(13))) > 5
            If InStr(ss2, Chr$(13)) = 0 Then GoTo ww1
            ss2 = Mid$(ss2, InStr(ss2, Chr$(13)) + 1)
        Wend
ww1:
        While (Len(ss2) - InStr(ss2, "、")) > 5
            If InStr(ss2, "、") = 0 Then GoTo ww2
            ss2 = Mid$(ss2, InStr(ss2, "、") + 1)
        Wend
ww2:
        Voice ss1 + " から " + ss2 + " まで　" + CStr(Len(str)) + "文字を選択"
    Else
        Voice str
    End If
 End Sub


'■ ｶｰｿﾙを左に移動 ■
Sub curLeft()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    'ｶｰｿﾙを左に移動
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
'■ Shift+ｶｰｿﾙを左に移動 ■
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
'■ ｶｰｿﾙを上に移動 ■
Sub curUp()
    
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Static Sainyu As Long
    If Sainyu <> 0 Then
        Quick
        Debug.Print "再入"
        Exit Sub
    End If
    Sainyu = -1
    LineVoiceFlag = False
    TopVoiceFlag = False
    'ｶｰｿﾙを上に移動
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
    'ｶｰｿﾙ位置を音声ｶﾞｲﾄﾞ
    Quick
    If TopVoiceFlag = True Then
        TopGuide
    End If
    If dum = Selection.Start Then
        Voice "トップ"
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
'■ Shift+ｶｰｿﾙを上に移動 ■
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
'■ ｶｰｿﾙを下に移動 ■
Sub curDown()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Static Sainyu As Long
    If Sainyu <> 0 Then
        Exit Sub
    End If
    Sainyu = -1
    'ｶｰｿﾙを下に移動
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
    'ｶｰｿﾙ位置を音声ｶﾞｲﾄﾞ
    Quick
    If DownVoiceFlag = True Then
        DownGuide
    End If
    If dum = Selection.Start Then
        Voice "した端 ラスト"
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
'■ Shift+ｶｰｿﾙを下に移動 ■
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
           Voice ("うえ端" + Chr$(&HFE))
    End If
    
End Sub
Private Sub DownGuide()
    On Error GoTo ee1
    Static Pos1, Pos2 As Long
    Pos1 = Selection.Bookmarks.Item("\Page").End
    Pos2 = Selection.Bookmarks.Item("\Line").End
    If Pos1 <= Pos2 Then
           Quick
           Voice ("した端" + Chr$(&HFE))
    End If
ee1:    Exit Sub
End Sub

Private Sub LeftGuide()
    If Selection.Information(wdFirstCharacterColumnNumber) = 1 Then
        If Left$(Selection.Text, 1) <> Chr$(13) Then
           Quick
           Voice ("ひだり端" + Chr$(&HFE))
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
            Voice ("みぎ端" + Chr$(&HFE))
        End If
    End If
ee1:    Exit Sub
    
End Sub
'■ 範囲読みｻﾌﾞ ■
Private Sub RangeGo(ByVal st1 As Long, ByVal en1 As Long)
    On Error GoTo ee1
    If st1 >= en1 Then
        Voice "ｸｳｷﾞｮｰ"
        Exit Sub
    End If
    Voice Mid$(Application.Selection.ShapeRange.TextFrame.TextRange.Text, st1 + 1, en1 - st1)
    Exit Sub
ee1: Resume ee2
ee2:
    Voice Application.ActiveDocument.Range(st1, en1).Text

End Sub
'
' ■　段落読み　■
'
Sub ParaRead()
    On Error GoTo ee1
    Quick
    Voice Selection.Paragraphs(1).Range.Text
ee1:    Exit Sub
End Sub

'　■１行読み■
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
'　■ｶｰｿﾙまで１行読み■
Sub LineReadLeft()
    On Error GoTo ee1
    dum = Selection.Start
   Dim ss, x, y As Long
    On Error GoTo ss1
    x = Selection.Information(wdFirstCharacterColumnNumber)
    If x = 1 Then
        Voice "行頭"
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
'　■ｶｰｿﾙ以降１行読み■
Sub LineReadRight()
    On Error GoTo ee1:
    Static ss, x, y As Long
    If Left$(Selection.Text, 1) = Chr$(13) Then
        Voice "改行"
        Exit Sub
    End If
    If Left$(Selection.Text, 1) = Chr$(11) Then
        Voice "行分割"
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
            Voice "行末"
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
            Voice "空行"
        Else
            Voice "行末"
        End If
    Else
        RangeGo ss, x
    
    End If
    
ee1:    Exit Sub

End Sub
'■書式読み■
Sub FormInfo()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    Quick
    Select Case Selection.Paragraphs(1).Alignment
    
    Case wdAlignParagraphDistribute
        Voice "均等 "
    Case wdAlignParagraphRight
        Voice "右詰 "
    Case wdAlignParagraphLeft
        Voice "左詰 "
    Case wdAlignParagraphCenter
        Voice "センタリング "
    Case wdAlignParagraphJustify
        Voice "左詰 "
    End Select
    Select Case Selection.Paragraphs(1).LineSpacingRule
    
    Case wdLineSpaceSingle
        Voice "行間１行 "
    Case wdLineSpace1pt5
        Voice "行間いってん５行 "
    Case wdLineSpaceDouble
        Voice "行間２行 "
    End Select
    If Selection.Paragraphs(1).FirstLineIndent + Selection.Paragraphs(1).LeftIndent <> 0 Then
        If Selection.Paragraphs(1).FirstLineIndent <> 0 Then
            Voice "１行目のインデント" + CStr(Selection.Paragraphs(1).FirstLineIndent _
            + Selection.Paragraphs(1).LeftIndent) + "ポイント"
        End If
    End If
    If Selection.Paragraphs(1).LeftIndent <> 0 Then
        If Selection.Paragraphs(1).FirstLineIndent < 0 Then
            Voice "ぶらさげインデント" + CStr(Selection.Paragraphs(1).LeftIndent) + "ポイント"
        Else
            Voice "左インデント" + CStr(Selection.Paragraphs(1).LeftIndent) + "ポイント"
        End If
    End If
    If Selection.Paragraphs(1).RightIndent <> 0 Then
        Voice "右インデント" + CStr(Selection.Paragraphs(1).RightIndent) + "ポイント"
       
    End If
    If KajoStr() <> "" Then
        Voice "箇条書き " + KajoStr()
    End If
     
End Sub



'■情報読み■
Sub CurInfo()
    If Documents.Count = 0 Then
        Exit Sub
    End If
    '    TableInfo
    CharInfo
    ShapeInfo
End Sub
'■ オブジェクト情報を読む ■
Public Sub ShapeInfo()
    On Error GoTo ee1

    Voice "オブジェクト名" + Chr$(&HFE) + Application.Selection.ShapeRange.Name + Chr$(&HFE)

    Exit Sub
ee1:    Voice "図形以外のオブジェクト"
        Exit Sub
End Sub
'■ 罫線情報を読む ■
Sub TableInfo()
    On Error GoTo ee1
    If keisen2 = True Then
    With Application.Selection.Cells
        Voice "罫線　うえ" + Chr$(&HFE)
        keisen (wdBorderTop)
        Voice "罫線　した" + Chr$(&HFE)
        keisen (wdBorderBottom)
        Voice "罫線　ひだり" + Chr$(&HFE)
        keisen (wdBorderLeft)
        Voice "罫線　みぎ" + Chr$(&HFE)
        keisen (wdBorderRight)
        
    End With
    Else
        Voice "罫線 なし"
    End If
ee1:    Exit Sub
End Sub
'■ 罫線情報を読む・ｻﾌﾞ2 ■
Private Function keisen2() As Boolean
    Dim Retsu, Gyo
    keisen2 = False
    On Error GoTo ee1
    Retsu = Selection.Information(wdStartOfRangeColumnNumber)
    Gyo = Selection.Information(wdStartOfRangeRowNumber)
    If Gyo > 0 And Retsu > 0 Then
        Voice "セル " + CStr(Gyo) + "行" + CStr(Retsu) + "列 "
        keisen2 = True
    End If
ee1: Exit Function
End Function
'■ 罫線情報を読む・ｻﾌﾞ ■
Private Sub keisen(s As Long)
    On Error GoTo ee1
    With Application.Selection.Cells

    If .Borders(s).Visible = False Then
        Voice "なし"
    Else
        Select Case .Borders(s).LineStyle
        Case wdLineStyleNone
            Voice "罫線 なし "
        Case wdLineStyleSingle
            Voice "通常罫線 "
        Case wdLineStyleDot
            Voice "点線 "
        Case wdLineStyleDashSmallGap
            Voice "破線 小 "
        Case wdLineStyleDashLargeGap
            Voice "破線 大 "
        Case wdLineStyleDashDot
            Voice "１点鎖線 "
        Case wdLineStyleDashDotDot
            Voice "２点鎖線 "
        Case wdLineStyleDouble
            Voice "２重罫線 "
        Case wdLineStyleTriple
            Voice "３重罫線 "
        Case Else
 
        Voice "ｽﾀｲﾙ" + CStr(.Borders(s).LineStyle) + Chr$(&HFE)
        End Select
        Select Case .Borders(s).LineWidth
            Case wdLineWidth025pt
                Voice "0.25ポイント" + Chr$(&HFE)
            Case wdLineWidth050pt
                Voice "0.5ポイント" + Chr$(&HFE)
            Case wdLineWidth075pt
                Voice "0.75ポイント" + Chr$(&HFE)
            Case wdLineWidth100pt
                Voice "1ポイント" + Chr$(&HFE)
            Case wdLineWidth150pt
                Voice "1.5ポイント" + Chr$(&HFE)
            Case wdLineWidth225pt
                Voice "2.25ポイント" + Chr$(&HFE)
            Case wdLineWidth300pt
                Voice "3ポイント" + Chr$(&HFE)
            Case wdLineWidth450pt
                Voice "4.5ポイント" + Chr$(&HFE)
            Case wdLineWidth600pt
                Voice "6ポイント" + Chr$(&HFE)
        End Select
        Select Case .Borders(s).ColorIndex
            Case wdAuto
                Voice "自動" + Chr$(&HFE)
            Case wdBlack
                Voice "ｸﾛｲﾛ" + Chr$(&HFE)
            Case wdBlue
                Voice "青色" + Chr$(&HFE)
            Case wdBrightGreen
                Voice "濃い緑色" + Chr$(&HFE)
            Case wdByAuthor
                Voice "ByAuthor" + Chr$(&HFE)
            Case wdDarkBlue
                Voice "濃い青色" + Chr$(&HFE)
            Case wdDarkRed
                Voice "濃い赤色" + Chr$(&HFE)
            Case wdDarkYellow
                Voice "濃い黄色" + Chr$(&HFE)
            Case wdGray25
                Voice "濃い灰色" + Chr$(&HFE)
            Case wdGray50
                Voice "灰色" + Chr$(&HFE)
            Case wdGreen
                Voice "緑色" + Chr$(&HFE)
            Case wdNoHighlight
                Voice "NoHighlight" + Chr$(&HFE)
            Case wdYellow
                Voice "黄色" + Chr$(&HFE)
            Case wdWhite
                Voice "白色" + Chr$(&HFE)
            Case wdViolet
                Voice "紫色" + Chr$(&HFE)
            Case wdTurquoise
                Voice "水色" + Chr$(&HFE)
            Case wdTeal
                Voice "青緑" + Chr$(&HFE)
            Case wdRed
                Voice "赤色" + Chr$(&HFE)
            Case wdPink
                Voice "ピンク色" + Chr$(&HFE)
            
        End Select
    End If
    End With
ee1:    Exit Sub
End Sub
'■ 文字情報を読む ■
Sub CharInfo()
    On Error GoTo eee1
    Dim str As String
    str = Left$(Application.Selection.Text, 1)
    SSVoice str
    VoiceWait
    If Asc(str) < 256 And Asc(str) >= 0 Then
        Voice "半角"
    Else
        Voice "全角"
    End If
    Select Case str
        Case "ぁ" To "ん"
            Voice "ひらがな" + Chr$(&HFE)
        Case "ｱ" To "ﾝ"
            Voice "かたかな" + Chr$(&HFE)
        Case "ァ" To "ヶ"
            Voice "かたかな" + Chr$(&HFE)
        Case "a" To "z"
            Voice "英数 小文字" + Chr$(&HFE)
        Case "A" To "Z"
            Voice "英数 大文字" + Chr$(&HFE)
        Case "Ａ" To "Ｚ"
            Voice "英数 大文字" + Chr$(&HFE)
        Case "ａ" To "ｚ"
            Voice "英数 小文字" + Chr$(&HFE)
    End Select
    If Application.Selection.Style <> "標準" Then
        Voice Application.Selection.Style + "のｽﾀｲﾙ" + Chr$(&HFE)
    End If
    Voice Application.Selection.Font.Name + Chr$(&HFE)
    Voice CStr(Application.Selection.Font.Size) + "ポイント" + Chr$(&HFE)
    If Application.Selection.Font.Bold = True Then
        Voice "太字" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Underline = True Then
        Voice "ｱﾝﾀﾞｰﾗｲﾝ" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Italic = True Then
        Voice "斜体" + Chr$(&HFE)
    End If
        Select Case Application.Selection.Font.ColorIndex
            Case wdAuto
                Voice "自動" + Chr$(&HFE)
            Case wdBlack
                Voice "ｸﾛｲﾛ" + Chr$(&HFE)
                             
           
            Case wdBlue
                Voice "青色" + Chr$(&HFE)
            Case wdBrightGreen
                Voice "濃い緑色" + Chr$(&HFE)
            Case wdByAuthor
                Voice "ByAuthor" + Chr$(&HFE)
            Case wdDarkBlue
                Voice "濃い青色" + Chr$(&HFE)
            Case wdDarkRed
                Voice "濃い赤色" + Chr$(&HFE)
            Case wdDarkYellow
                Voice "濃い黄色" + Chr$(&HFE)
            Case wdGray25
                Voice "濃い灰色" + Chr$(&HFE)
            Case wdGray50
                Voice "灰色" + Chr$(&HFE)
            Case wdGreen
                Voice "緑色" + Chr$(&HFE)
            Case wdNoHighlight
                Voice "NoHighlight" + Chr$(&HFE)
            Case wdYellow
                Voice "黄色" + Chr$(&HFE)
            Case wdWhite
                Voice "白色" + Chr$(&HFE)
            Case wdViolet
                Voice "紫色" + Chr$(&HFE)
            Case wdTurquoise
                Voice "水色" + Chr$(&HFE)
            Case wdTeal
                Voice "青緑" + Chr$(&HFE)
            Case wdRed
                Voice "赤色" + Chr$(&HFE)
            Case wdPink
                Voice "ピンク色" + Chr$(&HFE)
             
        End Select
    If Application.Selection.Font.DoubleStrikeThrough = True Then
        Voice "二重取り消し線" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Emboss = True Then
        Voice "浮きだし" + Chr$(&HFE)
    End If
    If Application.Selection.Font.EmphasisMark = True Then
        Voice "傍点" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Engrave = True Then
        Voice "浮き彫り" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Hidden = True Then
        Voice "隠し文字" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Outline = True Then
        Voice "中抜き" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Shadow = True Then
        Voice "影付き" + Chr$(&HFE)
    End If
    If Application.Selection.Font.StrikeThrough = True Then
        Voice "取消し線" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Subscript = True Then
        Voice "下付き文字" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Superscript = True Then
        Voice "上付き文字" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Engrave = True Then
        Voice "浮き彫り" + Chr$(&HFE)
    End If
    If Application.Selection.Font.Scaling <> 100 Then
        Voice "水平倍率" + CStr(Application.Selection.Font.Scaling) + Chr$(&HFE)
    End If
    AmiKake
eee1:    Exit Sub
End Sub

'■ 文字情報を読む・ｻﾌﾞ－網掛け ■

Private Sub AmiKake()
    On Error GoTo ee1
    If Application.Selection.Font.Shading.Texture <> 0 Then
        Voice "網掛けｽﾀｲﾙ" + CStr(Application.Selection.Font.Shading.Texture) + Chr$(&HFE)
    End If
ee1:    Exit Sub
End Sub

'■　横縦書き情報の取得　■
'
Private Function GetTateYoko() As Long
    On Error GoTo eeee1
    
    
    Select Case Application.Selection.Orientation
        Case wdTextOrientationHorizontal '通常横書き
            GetTateYoko = 0
        Case wdTextOrientationVerticalFarEast '通常縦書き
           GetTateYoko = 10
        Case wdTextOrientationUpward  '横書き　左９０度回転　文字列は上方向に
            GetTateYoko = 1
        Case wdTextOrientationDownward  '横書き　右９０度回転　文字列は下方向に
            GetTateYoko = 2
        Case wdTextOrientationHorizontalRotatedFarEast '縦書き　左９０度回転　文字列は右方向に
            GetTateYoko = 11
    End Select
    If TateYokoMode <> GetTateYoko Then
    Select Case Application.Selection.Orientation
        Case wdTextOrientationHorizontal '通常横書き
            Voice "横書きモード"
        Case wdTextOrientationVerticalFarEast '通常縦書き
            Voice "縦書きモード"
        Case wdTextOrientationUpward  '横書き　左９０度回転　文字列は上方向に
            Voice "横書き　左９０度回転モード"
        Case wdTextOrientationDownward  '横書き　右９０度回転　文字列は下方向に
            Voice "横書き　右９０度回転モード"
        Case wdTextOrientationHorizontalRotatedFarEast '縦書き　左９０度回転　文字列は右方向に
            Voice "縦書き　左９０度回転モード"
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
