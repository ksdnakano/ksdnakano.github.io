Attribute VB_Name = "PCTK"
'
'ワード読み上げマクロ
'このプログラムは著作権で保護されています。
'（株）高知システム開発

Option Explicit
Private Declare Function PCTKSETSTATUS Lib "PCTKUSR.dll" (ByVal para1 As Long, ByVal para2 As Long, ByVal para3 As Long) As Long
Private Declare Sub PCTKPREAD Lib "PCTKUSR.dll" (ByVal vString As String, ByVal vMode As Long, ByVal vFlag As Long)
Private Declare Sub PCTKCGUIDE Lib "PCTKUSR.dll" (ByVal vString As String, ByVal vMode As Long)
Private Declare Sub PCTKVRESET Lib "PCTKUSR.dll" ()
Private Declare Function PCTKGETVSTATUS Lib "PCTKUSR.dll" () As Long
Private Declare Sub PCTKBEEP Lib "PCTKUSR.dll" (ByVal BeepType As Long, ByVal UINT As Long, ByVal MachinType As Long)
'PCTalkerユーザーインターフェイスヘッダ
Public Const PS_AGSSTATUS = &H20000
Public Const PS_VOSSTATUS = &H10000

Public Const PS_APREADCALLBACK = &H20001       '   読み上げ時コールバック関数設定
Public Const PS_APNOREAD = &H20002     '   読み上げ禁止設定
Public Const PS_APNOACTION = &H20003       '   PCTalker動作禁止設定
Public Const PS_APSCUTCALLBACK = &H20004       '   ショートカットコールバック設定
Public Const PS_STATUS = &H10001       '   項目設定
Public Const PS_STATUSNEXT = &H10002       '   項目設定トグル次
Public Const PS_STATUSBACK = &H10003       '   項目設定トグル前

'
'   PCTalker読み上げ動作時コールバック関数の引数 act
'   アプリケーションの読み上げ動作
Const ACV_MENUSELECT = 0                ' メニュー選択時の読み上げ
Const ACV_MENUCLOSE = 1                 ' メニュークローズ時の読み上げ
Const ACV_CBOPENCLOSE = 2           ' コンボボックスオープンクローズ時の読み上げ
Const ACV_WINACTIVATE = 3           ' ウィンドウアクティブ時の読み上げ
Const ACV_WINCREATE = 4             ' アプリケーションの起動時時の読み上げ
Const ACV_WINCLOSE = 5          ' ウィンドウクローズ時の読み上げ
Const ACV_LTSELECT = 6              ' リスト系、タブシート選択切替時の読み上げ
Const ACV_TRACKBAR = 7          ' トラックバーの移動時の読み上げ
Const ACV_FOCUS = 8             ' フォーカス設定時の読み上げ
Const ACV_CHKBUTTON = 9             ' チェックボタン切替時の読み上げ
Const ACV_EDITMOVECUR = 10              ' ｴﾃﾞｨｯﾄｺﾝﾄﾛｰﾙカーソル移動時の読み上げ

'   新しいコールバック関数フラグ値(PC-Talker Ver 3.0以上)
Const ACV_IMM_ONOFF = 11                ' 日本語変換ON/OFF時読み上げ
Const ACV_IMM_INPUT = 12                ' 日本語入力時読み上げ
Const ACV_IMM_MODECHANGE = 13               ' 日本語変換モード切替時読み上げ
Const ACV_CHAR_INPUT = 14               ' 文字(半角)入力時読み上げ


'   PCTKVoiceGuide引数  quick
'   クィック設定値
Const PCTKQUICK_NONE = 0                '   システムに任す
Const PCTKQUICK_OFF = 1             '   クィックOFF
Const PCTKQUICK_ON = 2              '   クィックON

Const PCTKERR_NOERROR = 0           '   正常
Const PCTKERR_OLEINIT = &H10001     '   OLEの初期化に失敗した
Const PCTKERR_SERVEREXEC = &H10002      '   Server起動エラー
Const PCTKERR_INTERFACE = &H10003       '   インターフェイス取得エラー
Public WaitFlag As Boolean
'日本語変換
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
        Voice "行分割"
    ElseIf str = Chr$(1) + Chr$(21) Then
        Voice "ピクチャー"
    ElseIf str = Chr$(13) + Chr$(7) Then
        Voice "セル改行"
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
        Voice "行分割"
    ElseIf str = Chr$(1) + Chr$(21) Then
        Voice "ピクチャー"
    ElseIf str = Chr$(13) + Chr$(7) Then
        Voice "セル改行"
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
