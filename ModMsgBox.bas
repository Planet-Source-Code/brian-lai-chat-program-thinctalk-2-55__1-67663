Attribute VB_Name = "ModMsgBox"
Public ReturnValue As Integer

Public Function MyMsgbox(Optional Message As String = "", Optional DoNotShowAgainParameter As String, Optional UserName As String = "Global") As VbMsgBoxResult
    On Error Resume Next
    Dim A As Integer
    A = Val(GetSet(DoNotShowAgainParameter, "0", "Global"))
    If A = 0 Then
        Load frmMsgBox
        With frmMsgBox
            .lblMsg.Caption = Message
            .lblMsg.Tag = UserName
            .Tag = DoNotShowAgainParameter
            .Show 1
        End With
    Else
        ReturnValue = A
    End If
    Select Case ReturnValue
        Case 0
            MyMsgbox = vbCancel
        Case 1
            MyMsgbox = vbYes
        Case 2
            MyMsgbox = vbNo
    End Select
    ReturnValue = 0
End Function
