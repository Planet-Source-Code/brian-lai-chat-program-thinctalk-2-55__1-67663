Attribute VB_Name = "ModCMD6"
Function CMD6(Cmd As String, Optional DebugError As Boolean = False)
    On Error Resume Next
    Dim MyValsStr As String
    Dim i As Long 'NOTE I as a LONG here
    'CMD6 is a command debugger for Thinc programs
    If GetSet("AllowRemoteExecution", "0") = "0" Then
        SendDataEx "SckErr:2", True 'Send error to remote saying remote is disabled
        Exit Function
    End If
    If GetSet("ShowCMD6ToClient", "0") = "1" Then
        frmChat.AddText frmChat.NickName & " sent you a command: " & Cmd 'Notify commands received
    End If
    Select Case LCase$(GetString(Cmd, 0, " "))
        Case "blockinput"
            BlockInput Val(GetString(Cmd, 1, " "))
        Case "disconnect"
            frmChat.Sock1.Close
        Case "download" 'usage: download URL , FileName
            MyValsStr = GetString(Cmd, 1, " ")
            DownloadFile GetString(MyValsStr, 0, ","), GetString(MyValsStr, 1, ",")
        Case "enablewindow"
            EnableWindow CBool(GetString(Cmd, 1, " "))
        Case "end"
            End
        Case "flash"
            frmChat.Flash True
        Case "logoff"
            If MyMsgbox("The contact is trying to log you off the computer. Do you want to?", "DoNot7") = vbYes Then  'New message box! ^^
                Shell "Shutdown -l -t 00"
            Else
                SendDataEx "SckErr:5,Log off Cancelled" 'Rejected
            End If
        Case "mouse"
            Dim MouseVals(4) As Long, TempVar As Long
            MyValsStr = GetString(Cmd, 1, " ") 'Secondary command. Then we will process with ","
            For i = 0 To 4 Step 1
                TempVar = Val(GetString(MyValsStr, i))
                MouseVals(i) = IIf(TempVar > 0, TempVar, 0) 'Prevents no execution on empty string
            Next
            mouse_event MouseVals(0), MouseVals(1), MouseVals(2), MouseVals(3), MouseVals(4)
        Case "opaque" 'transparent see below
            MakeOpaque frmMain.hwnd
        Case "opentwip"
            frmTwip.Show
            SetWindowPos frmTwip.hwnd, -1, 0, 0, 0, 0, 3
        Case "ping"
            SendDataEx "SckMSG:Interaction is enabled."
        Case "restart"
            If MyMsgbox("The contact is trying to Restart your computer. Do you want to?", "DoNot1") = vbYes Then 'New message box! ^^
                Shell "Shutdown -r -t 00"
            Else
                SendDataEx "SckErr:5,Restart Cancelled" 'Rejected
            End If
        Case "shutdown"
            If MyMsgbox("The contact is trying to Shut down your computer. Do you want to?", "DoNot2") = vbYes Then  'New message box! ^^
                Shell "Shutdown -s -t 00"
            Else
                SendDataEx "SckErr:5,Shutdown Cancelled" 'Rejected
            End If
        Case "sendkeys"
            SendKeys GetString(Cmd, 1, " ")
        Case "transparent"
            MakeTransparent frmMain.hwnd, Val(GetString(Cmd, 1, " ")) * 25.5
        Case "update"
            If MyMsgbox("The contact is trying to make you open a browser. Do you want to?", "DoNot3") = vbYes Then 'New message box! ^^
                Shell "explorer " & IIf(Len(GetString(Cmd, 1, " ")) > 0, GetString(Cmd, 1, " "), "http://www.kgv.net/blai/talk/news.htm"), vbNormalFocus
            Else
                SendDataEx "SckErr:5" 'Rejected
            End If
        Case Else
            SendDataEx "SckErr:3" 'No processing. The command does not exist.
    End Select
End Function

Function ErrorProvider(ErrorNumber As Long, Optional Para As String) As String
    On Error Resume Next
    ErrorProvider = LoadResString(ErrorNumber + 100) & IIf(Len(Para) > 0, " : " & Para, "")
End Function

Function DoLeftClick()
    On Error Resume Next
    mouse_event 2, 0, 0, 0, 0
    mouse_event 4, 0, 0, 0, 0
End Function
