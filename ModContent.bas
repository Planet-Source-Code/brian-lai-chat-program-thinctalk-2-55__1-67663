Attribute VB_Name = "ModCOntent"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Dim sTime#


Public Function StripHTML(sHTML As String, ReplStr As String) As String
'ThincTalk: This is an external module I got from somewhere else

'Strip tag information from HTML
'this example was submitted by Tom Pydeski, BitWise Industrial Automation, Inc.
'modified 4/7/2005 to use the split.
'the original method concatenates stemp, which takes a lonnnnngggg time
'This one creates an array of lines and joins them at the end, but
'uses a byte array of the text instead of instr to find the tag characters
'I also utilized copy memory to move the text to the array, instead of
'the VB Left$ string manipulation function
Dim sTemp As String
Dim lngTagBeg As Long, lngTagEnd As Long, l As Long, TagsFound As Long, sLen As Long
Dim aryHTML$()
Dim ChArray() As Byte
sTemp$ = sHTML$
sTime# = Timer
ChArray() = StrConv(sTemp$, vbFromUnicode)
lngTagBeg = 0
lngTagEnd = 0
For l = 0 To UBound(ChArray) - 1
    If ChArray(l) = 60 Then
        Exit For
    End If
Next l
lngTagBeg = l + 1
If lngTagBeg >= UBound(ChArray) - 1 Then
    StripHTML = sHTML
    Exit Function
End If
ReDim Preserve aryHTML$(TagsFound + 1)
aryHTML$(TagsFound) = Left$(sTemp$, lngTagBeg - 1)
TagsFound = 1
Do
    ReDim Preserve aryHTML$(TagsFound + 1)
    For l = lngTagBeg To UBound(ChArray) - 1
        If ChArray(l) = 62 Then
            Exit For
        End If
    Next l
    lngTagEnd = l + 1
    For l = lngTagEnd To UBound(ChArray) - 1
        If ChArray(l) = 60 Then
            Exit For
        End If
    Next l
    lngTagBeg = l + 1
    If lngTagBeg >= UBound(ChArray) - 1 Then
        Exit Do
    End If
    sLen = (lngTagBeg - lngTagEnd) - 1
    aryHTML$(TagsFound) = Space(sLen)
    CopyMemory ByVal aryHTML$(TagsFound), ByVal VarPtr(ChArray(lngTagEnd)), sLen
    TagsFound = TagsFound + 1
Loop
Debug.Print TagsFound; "Tags Removed in "; FormatNumber(Timer - sTime#, 3)
StripHTML = Join(aryHTML$(), ReplStr)
End Function

Public Function AddImg(Src As String, Optional YouOrMe As Integer = 0, _
                                    Optional Width As Long = 120, Optional Height As Long = 90) As String
    On Error Resume Next
    Dim K As String
    LastLine = LastLine + 1 'Counting up of the last line variable
    K = "<font face=""arial"" size=""2"" color=""#7F7F7F""><i>" & IIf(YouOrMe = 0, frmChat.NickName, frmChat.ClientNick) & _
            " Sends:</i><br><a name=""" & LastLine & """ target=""top"" href=""" & Src & """>" & _
            "<img border=""0"" src=""" & Src & """ width=""" & Width & """ height=""" & Height & """></a></font><br>"
    frmChat.WriteHTML K
    frmChat.txtConvo.Navigate2 FindPath(App.Path, "tmppage.html") & "#K" & LastLine
End Function



