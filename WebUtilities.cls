VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebUtilities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function URLDecode(ByVal strIn As String) As String
    On Error Resume Next
    Dim sl&, tl&, Key$, kl&
    sl = 1:    tl = 1: Key = "%": kl = Len(Key)
    sl = InStr(sl, strIn, Key, 1)
    Do While sl > 0
        If (tl = 1 And sl <> 1) Or tl < sl Then
            URLDecode = URLDecode & Mid(strIn, tl, sl - tl)
        End If
        Dim hh$, hi$, hl$, a$
        Select Case UCase(Mid(strIn, sl + kl, 1))
            Case "U"    'Unicode URLEncode
                a = Mid(strIn, sl + kl + 1, 4)
                URLDecode = URLDecode & ChrW("&H" & a)
                sl = sl + 6
            Case "E"    'UTF-8 URLEncode
                hh = Mid(strIn, sl + kl, 2)
                a = Int("&H" & hh)    'ascii?
                If Abs(a) < 128 Then
                    sl = sl + 3
                    URLDecode = URLDecode & Chr(a)
                Else
                    hi = Mid(strIn, sl + 3 + kl, 2)
                    hl = Mid(strIn, sl + 6 + kl, 2)
                    a = ("&H" & hh And &HF) * 2 ^ 12 Or ("&H" & hi And &H3F) * 2 ^ 6 Or ("&H" & hl And &H3F)
                    If a < 0 Then a = a + 65536
                    URLDecode = URLDecode & ChrW(a)
                    sl = sl + 9
                End If
            Case Else    'Asc URLEncode
                hh = Mid(strIn, sl + kl, 2)    '??
                a = Int("&H" & hh)    'ascii?
                If Abs(a) < 128 Then
                    sl = sl + 3
                Else
                    hi = Mid(strIn, sl + 3 + kl, 2)    '??
                    'a = Int("&H" & hh & hi) '?ascii?
                    a = (Int("&H" & hh) - 194) * 64 + Int("&H" & hi)
                    sl = sl + 6
                End If
                URLDecode = URLDecode & ChrW(a)
        End Select
        tl = sl
        sl = InStr(sl, strIn, Key, 1)
    Loop
    URLDecode = URLDecode & Mid(strIn, tl)
End Function

Public Function URLEncode(ByVal txt As String) As String
    Dim buffer As String, i As Long, c As Long, n As Long
    buffer = String$(Len(txt) * 12, "%")
 
    For i = 1 To Len(txt)
        c = AscW(Mid$(txt, i, 1)) And 65535
        Select Case c
            Case 48 To 57, 65 To 90, 97 To 122, 45, 46, 95  ' Unescaped 0-9A-Za-z-._ '
                n = n + 1
                Mid$(buffer, n) = ChrW(c)
            Case Is <= 127            ' Escaped UTF-8 1 bytes U+0000 to U+007F '
                n = n + 3
                Mid$(buffer, n - 1) = Right$(Hex$(256 + c), 2)
            Case Is <= 2047           ' Escaped UTF-8 2 bytes U+0080 to U+07FF '
                n = n + 6
                Mid$(buffer, n - 4) = Hex$(192 + (c \ 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case 55296 To 57343       ' Escaped UTF-8 4 bytes U+010000 to U+10FFFF '
                i = i + 1
                c = 65536 + (c Mod 1024) * 1024 + (AscW(Mid$(txt, i, 1)) And 1023)
                n = n + 12
                Mid$(buffer, n - 10) = Hex$(240 + (c \ 262144))
                Mid$(buffer, n - 7) = Hex$(128 + ((c \ 4096) Mod 64))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
            Case Else                 ' Escaped UTF-8 3 bytes U+0800 to U+FFFF '
                n = n + 9
                Mid$(buffer, n - 7) = Hex$(224 + (c \ 4096))
                Mid$(buffer, n - 4) = Hex$(128 + ((c \ 64) Mod 64))
                Mid$(buffer, n - 1) = Hex$(128 + (c Mod 64))
        End Select
    Next
    URLEncode = Left$(buffer, n)
End Function

Public Function GetMimeType(ByVal FileName As String) As String
    Dim strMimeType As String
    Dim strExt As String
    Dim objReg As RegistryUtility
    Dim objFSO As Scripting.FileSystemObject
    strMimeType = "application/unknown"
    Set objFSO = New Scripting.FileSystemObject
    strExt = objFSO.GetExtensionName(FileName)
    On Error Resume Next
    Set objReg = New RegistryUtility
    strMimeType = CStr(objReg.ReadRegValue("HKEY_CLASSES_ROOT\." & strExt & "\Content Type"))
    GetMimeType = strMimeType
End Function

Public Function EncodeBase64(ByVal Text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(Text, vbFromUnicode)
  Dim objXML As Variant
  Dim objNode As Variant
  Set objXML = CreateObject("MSXML2.DOMDocument")
  Set objNode = objXML.createElement("b64")
  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = Replace(objNode.Text, vbLf, "")
End Function

Public Function DecodeBase64(ByVal strData As String) As Byte()
    Dim objXML2 As MSXML2.DOMDocument60
    Dim objNode As MSXML2.IXMLDOMElement
    Set objXML2 = New MSXML2.DOMDocument60
    Set objNode = objXML2.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue
End Function

Public Function ConstructQueryParameters(ByVal Url As String, ByVal Params As Scripting.Dictionary) As String
    Dim arrKeys() As Variant
    Dim strUrl As String
    strUrl = Url
    Dim i As Long
    arrKeys = Params.Keys
    If UBound(arrKeys) >= 0 Then
        strUrl = strUrl & "?"
        For i = LBound(arrKeys) To UBound(arrKeys)
            If i = UBound(arrKeys) Then
                strUrl = strUrl & arrKeys(i) & "=" & Params.Item(arrKeys(i))
            Else: strUrl = strUrl & arrKeys(i) & "=" & Params.Item(arrKeys(i)) & "&"
            End If
        Next
    End If
    ConstructQueryParameters = strUrl
End Function

