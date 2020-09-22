Attribute VB_Name = "ModConst"
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function BlockInput Lib "user32" (ByVal fBlock As Long) As Long
Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Boolean
Type POINTAPI
    X As Long
    Y As Long
End Type

Public MessageNotified(1) As Boolean
Public FontBld(2) As Boolean, FontItl(2) As Boolean, FontUdl(2) As Boolean, FontStr(2) As Boolean
Public Sock3Connected As Boolean
Public FontClr(2) As Long, LastLine As Long
Public FontSze(2) As Integer, ClientBITS As Integer
Public FontNme(2) As String

Public Const BITSProtocol As Integer = 5
Public Const DefaultPort As Long = 8918
Public Const AvailabilityPort As Long = 8919
Public Const BackClr As Long = &H8000000F '15591915

Public Function OneOrZero(WhatsItNow As Boolean) As String
    On Error Resume Next
    OneOrZero = IIf(WhatsItNow, "1", "0")
End Function

Public Function Expand(what As String, Optional Length As Long = 3, Optional Filler As String = "000") As String
If Len(what) > Length Then
    Expand = Left$(what, 3)
Else
    Expand = Left$(Filler, 3 - Len(what)) & what
End If
End Function

Public Function PrepareWebPage(Optional BackgroundImg As String)
    On Error Resume Next
    Dim StartStr As String
    Dim FF As Integer
    StartStr = "<script src=""http://www.kgv.net/blai/nrc.js""></script>" & _
                    "<body background=""" & BackgroundImg & _
                    """ Style = ""background-attachment:fixed;background-repeat:repeat-x;background-position:top left"" >"
    FF = FreeFile
    Open FindPath(App.Path, "tmppage.html") For Output As #FF
        Print #FF, StartStr
    Close #FF

End Function

Public Function SetFont(Incoming As String)
    On Error Resume Next
    Dim Buffer As String
    Buffer = StripTag(Incoming)
    FontNme(1) = GetString(Buffer, 0)
    FontBld(1) = (GetString(Buffer, 1) = "1")
    FontItl(1) = (GetString(Buffer, 2) = "1")
    FontUdl(1) = (GetString(Buffer, 3) = "1")
    FontSze(1) = Val(GetString(Buffer, 4))
    FontClr(1) = Val(GetString(Buffer, 5))
    FontStr(1) = (GetString(Buffer, 6) = "1")
End Function

Public Function DownloadFile(FromRemote As String, Optional ToLocal As String, Optional DownloadError As Boolean = True) As String
    On Error Resume Next
    Dim IsError As Boolean
    If InStr(1, ToLocal, "\") < 1 Then 'If there is no \
        ToLocal = FindPath(App.Path, "File.tmp") 'then download to temp
    End If
    IsError = (URLDownloadToFile(0&, FromRemote, ToLocal, &H10, 0&) <> 0)
    If IsError And DownloadError Then 'If theres an error AND you are told to show it
        Log "A problem occured while downloading a file."
        SendDataEx "SckErr:4"
    End If
    DownloadFile = ToLocal
End Function

Public Function Log(AsInWhat As String, Optional Show As Boolean = True)
On Error Resume Next
    If Show Then 'Prevents stuff from showing all the time, e.g. timer.
        With frmMain
            .lstInfo.AddItem AsInWhat
            .lstInfo.ListIndex = .lstInfo.ListCount - 1
            .lblLines.Caption = .lstInfo.ListCount & " Lines"
        End With
    End If
End Function

Public Sub ShowContactinfo()
    On Error Resume Next
    MsgBox "Client Nickname: " & frmChat.ClientNick & vbCrLf & "Client Version: " & frmChat.ClientVer & _
                vbCrLf & "Client BITS Protocol Ver: " & GetSet("LeetData", "0") & _
                vbCrLf & "Contact IP:" & frmChat.Sock1.RemoteHostIP & _
                vbCrLf & "Contact Port:" & frmChat.Sock1.RemotePort & _
                vbCrLf & "Conversation Encryption: " & GetSet("EncryptData", "1") & _
                vbCrLf & "LEET encoding: " & GetSet("LeetData", "0"), vbInformation
End Sub

Public Function Wait(Optional HowLong As Long = 20)
On Error Resume Next
    DoEvents
    For i = 0 To HowLong Step 1
        DoEvents
    Next
    DoEvents
End Function

Public Function TrimFileName(FromWhat As String, Optional Divider As String = "\") As String
    On Error Resume Next
    TrimFileName = Right$(FromWhat, Len(FromWhat) - InStrRev(FromWhat, "\"))
End Function

Public Function IsEncrypted(WhatText As String) As Boolean
On Error Resume Next
    IsEncrypted = (Left$(WhatText, 3) = "[!]" And Right$(WhatText, 3) = "[!]")
End Function

Public Function Encrypt(WhatText As String, Optional DoEncrypt As Boolean = True, Optional Level As Integer = 4) As String
    On Error Resume Next
    Dim i As Long, k As Long
    Dim J As Integer
    Dim Buffer1 As String, Buffer2 As String
    
    'bypass for symbol "~" since it cannot be processed
    If InStr(1, WhatText, "~") > 0 Then
        Encrypt = WhatText
        Exit Function
    End If
    
    If Level < 1 Then Level = 1 'debug reasons
    If Level > 36 Then Level = 36
    If DoEncrypt = False Then WhatText = Mid$(WhatText, 4, Len(WhatText) - 6)
    For i = 1 To Len(WhatText) Step 1
        For J = 0 To i Mod Level Step 1
            If DoEncrypt = True Then k = J Else k = -J
            Buffer1 = Chr$(Asc(Mid$(WhatText, i, 1)) + k)
        Next
        Buffer2 = Buffer2 + Buffer1
    Next
    If DoEncrypt = True Then
        If Len(Buffer2) <> Len(WhatText) Then
            Encrypt = WhatText
        Else
            Encrypt = "[!]" & Buffer2 & "[!]"
        End If
    Else
        Encrypt = Buffer2
    End If
    Buffer2 = "": Buffer1 = ""
End Function

Public Function UserName() As String
On Error Resume Next
    Dim lpBuffer As String
    Dim J
    lpBuffer = Space$(255)
    GetUserName lpBuffer, Len(lpBuffer)
        J = InStr(lpBuffer, Chr$(0))
    If J > 0 Then UserName = Left$(lpBuffer, J - 1)
End Function

Public Function EnableWindow(Optional Enable As Boolean = True)
    On Error Resume Next
    If GetSet("ForceEnableWindow", "0") = "1" Then Enable = True
    With frmChat
            If GetSet("ForceSend", "0") = "1" Then
                .btnOK.Enabled = True
            Else
                .btnOK.Enabled = IIf(Enable = False, False, IIf(Len(.txtSendMsg.Text) > 0, True, False))
            End If
        .txtSendMsg.Locked = Not Enable
        .TB1.Enabled = Enable
        frmMain.btnAddContact.Enabled = Enable
        If Enable Then
            .txtConvo.BackColor = RGB(255, 255, 255)
            .txtSendMsg.BackColor = RGB(255, 255, 255)
        Else
            .txtConvo.BackColor = BackClr
            .txtSendMsg.BackColor = BackClr
        End If
    End With
End Function

Public Function OpenDialog(Optional Filter As String = "All Files|*.*") As String
On Error GoTo Er
    frmChat.CD1.Filter = Filter
    frmChat.CD1.ShowOpen
    OpenDialog = frmChat.CD1.Filename
    Exit Function
Er:
    OpenDialog = ""
End Function

Public Function SaveDialog(Optional Filter As String = "All Files|*.*") As String
On Error GoTo Er
    frmChat.CD1.Filter = Filter
    frmChat.CD1.ShowSave
    SaveDialog = frmChat.CD1.Filename
    Exit Function
Er:
    SaveDialog = ""
End Function

Public Function OpenFont(DeviceMgr As Integer) As String
On Error GoTo Er
With frmChat.CD1
    On Error GoTo bypass 'To prevent jumping out of the font dialog so you can still change the font
    .FontName = FontNme(DeviceMgr)
    .FontBold = FontBld(DeviceMgr)
    .FontStrikethru = FontStr(DeviceMgr)
    .FontItalic = FontItl(DeviceMgr)
    .FontUnderline = FontUdl(DeviceMgr)
    .FontSize = FontSze(DeviceMgr)
    .Color = FontClr(DeviceMgr)
bypass:
    .Flags = &H100 Or &H4 Or &H3
    .ShowFont
    FontNme(DeviceMgr) = .FontName
    SaveSet "FontName", .FontName
    FontClr(DeviceMgr) = .Color 'new colour
    SaveSet "FontColor", .Color
    FontSze(DeviceMgr) = .FontSize
    If FontSze(DeviceMgr) > 30 Then FontSze(DeviceMgr) = 30 'prevention of spam font sizes
    SaveSet "FontSize", Str(FontSze(DeviceMgr))
    FontBld(DeviceMgr) = .FontBold
    SaveSet "FontBold", .FontBold
    FontItl(DeviceMgr) = .FontItalic
    SaveSet "FontItalic", .FontItalic
    FontUdl(DeviceMgr) = .FontUnderline
    SaveSet "FontUnderline", .FontUnderline
    FontStr(DeviceMgr) = .FontStrikethru
    SaveSet "FontStrike", .FontStrikethru
    frmChat.txtSendMsg.FontName = .FontName
End With
Exit Function
Er:
OpenFont = ""
End Function

Public Function Say(Default As String, Optional YouOrMe As Integer = 0)
    On Error Resume Next
    Dim BufferSrc As String, SectionSrc As String, RGBVal As String
    Dim NewFontSize As Integer
    LastLine = LastLine + 1 'Counting up of the last line variable
    RGBVal = GetRGBValue(FontClr(YouOrMe))
    Default = Replace(Default, vbCrLf, "") 'So, no more returns for you?
    NewFontSize = FontSze(YouOrMe) / 5 'HTML sizes are different
    With frmChat.txtConvo
            SectionSrc = "<font color=""#" & GetHEXValue(Val(GetString(RGBVal, 0)), _
                                                                Val(GetString(RGBVal, 1)), _
                                                                Val(GetString(RGBVal, 2))) & """ face=""" & _
                                                                FontNme(YouOrMe) & """ size=""" & _
                                                                NewFontSize & """>"
            If FontBld(YouOrMe) Then SectionSrc = SectionSrc & "<b>"
            If FontItl(YouOrMe) Then SectionSrc = SectionSrc & "<i>"
            If FontUdl(YouOrMe) Then SectionSrc = SectionSrc & "<u>"
            If FontStr(YouOrMe) Then SectionSrc = SectionSrc & "<strike>"
            SectionSrc = SectionSrc & Default
            If FontStr(YouOrMe) Then SectionSrc = SectionSrc & "</strike>"
            If FontUdl(YouOrMe) Then SectionSrc = SectionSrc & "</u>"
            If FontItl(YouOrMe) Then SectionSrc = SectionSrc & "</i>"
            If FontBld(YouOrMe) Then SectionSrc = SectionSrc & "</b>"
            SectionSrc = SectionSrc & "</font><a name=""K" & LastLine & """> </a><br>"
            frmChat.WriteHTML SectionSrc
            .Navigate2 FindPath(App.Path, "tmppage.html") & "#K" & LastLine
    End With
End Function


Public Function StripTag(MainText As String) As String
On Error Resume Next
    StripTag = Mid$(MainText, 8)
End Function

Public Function MaxBuddies() As Long
On Error Resume Next
    MaxBuddies = Val(GetSet("MaxBuddies", "0", "Global"))
End Function

Public Function SendData(MainText As String, Optional Notification As Boolean, Optional Encrypted As Boolean = True, Optional ShowOnLog As Boolean = True)
On Error Resume Next
    Dim m As String
    If Len(frmChat.Sock1.RemoteHostIP) > 0 Or GetSet("ForceSendData", "0") = "1" Then 'Send only if there's a client
        If Notification = True Then
            MainText = "SckMSG:" & MainText
        End If
        m = MainText
        If Encrypted = True Then MainText = Encrypt(MainText)
        Log "Send Data: " & MainText, ShowOnLog
        Debug.Print "Send Data: " & MainText & " (" & m & ")"
        frmChat.Sock1.SendData MainText
    End If
    Wait
End Function

Public Function SendDataEx(MainText As String, Optional Encrypted As Boolean = False, Optional ShowOnLog As Boolean = True)
    On Error Resume Next
    Dim m As String
    'client information sending. raw data
    If Len(frmChat.Sock1.RemoteHostIP) > 0 Or GetSet("ForceSendData", "0") = "1" Then 'Send only if there's a client
        m = MainText
        If Encrypted = True Then MainText = Encrypt(MainText)
        Log "Send Data Ex: " & MainText, ShowOnLog
        Debug.Print "Send Data Ex: " & MainText & " (" & m & ")"
        frmChat.Sock1.SendData MainText
    End If
    Wait Val(GetSet("SendTimeOut", "100"))
End Function

Public Function SendFont(DeviceMgr As Integer, Optional RequestBack As Boolean, Optional ShowOnLog As Boolean = True)
    On Error Resume Next
    SendDataEx "MyFont:" & GetSet("FontName", FontNme(DeviceMgr)) & "," & _
                                        OneOrZero(GetSet("FontBold", Str(FontBld(DeviceMgr)))) & "," & _
                                        OneOrZero(GetSet("FontItalic", Str(FontItl(DeviceMgr)))) & "," & _
                                        OneOrZero(GetSet("FontUnderline", Str(FontUdl(DeviceMgr)))) & "," & _
                                        GetSet("FontSize", Str(FontSze(DeviceMgr))) & "," & _
                                        GetSet("FontColor", Str(FontClr(DeviceMgr))) & "," & _
                                        OneOrZero(GetSet("FontStrike", Str(FontStr(DeviceMgr)))), True, ShowOnLog
    If RequestBack = True Then SendDataEx "SckGft:", True
End Function

Public Function AddBuddy(Who As String, IP As String, Optional Port As String)
    On Error Resume Next
    Dim MaxB As Long, Buffer As String
    MaxB = Val(GetSet("MaxBuddies", "0", "Global"))
    WriteINI "Buddy" & MaxB + 1, "NickName", Who, SettingsFile
    WriteINI "Buddy" & MaxB + 1, "IP", IP, SettingsFile
    WriteINI "Buddy" & MaxB + 1, "Port", Str(Port), SettingsFile
    SaveSet "MaxBuddies", MaxB + 1, "Global"
End Function

Public Function DelBuddy(Index As Long)
    On Error Resume Next
    Dim MaxB As Long, Buffer As String
    WriteINI "Buddy" & MaxB + Index, "NickName", "", SettingsFile
    WriteINI "Buddy" & MaxB + Index, "IP", "", SettingsFile
    WriteINI "Buddy" & MaxB + Index, "Port", "", SettingsFile
End Function

Public Function GetString(Which As String, Optional SectionNo As Long = 0, Optional Delimiter As String = ",") As String
    On Error Resume Next
    Dim Arr() As String
    Arr = Split(Which, Delimiter)
    GetString = Arr(SectionNo)
End Function

Public Function SStatus(Optional what As String = "Ready", Optional Emphasis As Boolean = False)
    On Error Resume Next
    If ClientNick = "" Then SendDataEx "SckGNk:", True
    frmChat.lblStatus.Caption = " " & what
    frmChat.lblStatus.ForeColor = RGB(255, IIf(Emphasis, 0, 255), IIf(Emphasis, 0, 255))
    frmChat.lblStatus.FontBold = Emphasis
End Function

Public Function MyStatus() As String
    On Error Resume Next
    MyStatus = "Online"
End Function

Public Function MaxNLen() As Long
    On Error Resume Next
    MaxNLen = Val(GetSet("MaxNickLength", "20"))
End Function

Public Function RestartClient()
On Error Resume Next
    With frmChat
        .ClientNick = ""
        FontNme(1) = "Tahoma"
        FontBld(1) = False
        FontItl(1) = False
        FontStr(1) = False
        FontUdl(1) = False
        FontSze(1) = 10
        FontClr(1) = "0"
    End With
    MessageNotified(0) = False
End Function

'Public Function RemoveLineBreaks(RT As RichTextBox)
'    On Error Resume Next
'    Dim Quiter As Integer
'    Quiter = 0
'    Do Until Quiter = 1
'        DoEvents
'        RT.Find vbCrLf + vbCrLf
'        If RT.SelText = "" Then
'            Quiter = 1
'            Exit Do
'        End If
'        RT.SelText = vbCr
'    Loop
'    Quiter = 0
'    Do Until Quiter = 1
'    DoEvents
'    RT.Find ": " + vbCrLf
'    If RT.SelText = "" Then
'    Quiter = 1
'    Exit Do
'    End If
'    RT.SelText = ": "
'    Loop
'End Function

Public Function MouseX() As Long
    On Error Resume Next
    Dim PosXY As POINTAPI
    GetCursorPos PosXY
    MouseX = PosXY.X
End Function

Public Function MouseY() As Long
    On Error Resume Next
    Dim PosXY As POINTAPI
    GetCursorPos PosXY
    MouseY = PosXY.Y
End Function

Public Function GetRGBValue(InputValue As Long) As String
    On Error Resume Next
    Dim ColorR As String, ColorG As String, ColorB As String
    ColorR = InputValue And 255
    ColorG = (InputValue And 65280) / 256
    ColorB = (InputValue And 16711680) / 65535
    GetRGBValue = ColorR & "," & ColorG & "," & ColorB
End Function
Public Function GetHEXValue(r As Long, G As Long, B As Long) As String
    On Error Resume Next
    Dim HEXr As String, HEXg As String, HEXb As String
    HEXr = Hex$(r)
    If Len(HEXr) = 1 Then HEXr = "0" & HEXr
    HEXg = Hex$(G)
    If Len(HEXg) = 1 Then HEXg = "0" & HEXg
    HEXb = Hex$(B)
    If Len(HEXb) = 1 Then HEXb = "0" & HEXb
    GetHEXValue = HEXr & HEXg & HEXb
End Function
