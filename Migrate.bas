Attribute VB_Name = "Module1"
'This Module is in public use with:
'HKID, Revert, ROT Decryp, FileA, QuickMolar, SockMSG


Public Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Const ICC_USEREX_CLASSES = &H200

Public Function XPVB() As Boolean
    On Error Resume Next
    If Dir(MyManifestFile) <> "" Then GoTo Written
    Dim XPStr As String
    Dim FF As Integer
    XPStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
            "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf & _
            "<assemblyIdentity version=""1.0.0.0"" processorArchitecture=""X86"" name=""Microsoft.VB6.VBnetStyles"" type=""win32""/>" & vbCrLf & _
            "<description>Windows XP manifest file</description>" & vbCrLf & "<dependency>" & vbCrLf & _
            "<dependentAssembly>" & vbCrLf & "<assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0"" processorArchitecture=""X86"" publicKeyToken=""6595b64144ccf1df"" language=""*""/>" & vbCrLf & _
            "</dependentAssembly>" & vbCrLf & "</dependency>" & vbCrLf & "</assembly>"
    FF = FreeFile
    Open MyManifestFile For Output As #FF
        Print #FF, XPStr
    Close #FF
Written:
    SetAttr MyManifestFile, 34 'vbHidden 'Hide the manifest
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    XPVB = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function FindPath(Parent As String, Optional Child As String, Optional Divider As String = "\") As String
    On Error Resume Next
    If Right$(Parent, 1) = Divider Then Parent = Left$(Parent, Len(Parent) - 1)
    If Left$(Child, 1) = Divider Then Child = Mid$(Child, 2)
    FindPath = Parent & Divider & Child
End Function

Public Function MyManifestFile() As String
    On Error Resume Next
    MyManifestFile = FindPath(App.Path, App.EXEName & ".exe.manifest")
End Function

Public Function MyVer() As String
    On Error Resume Next
    Dim Buffer2 As String
    Dim PreVer As Integer
    PreVer = App.Minor
    If App.Revision >= 1 Then
        PreVer = PreVer + 1
        Buffer2 = Trim$(Str$(PreVer) & " BETA")
    Else
        Buffer2 = Trim$(Str$(PreVer))
    End If
    MyVer = "V." & App.Major & "." & Buffer2
End Function

Sub Main()
    On Error Resume Next
    XPVB
    PrepareWebPage GetSet("BackgroundImg", "http://www.kgv.net/blai/Images/Bkg.bmp") ' top-priority task. Now added the background image! xD
    frmMain.Show
End Sub
