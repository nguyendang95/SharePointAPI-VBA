VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DriveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_MimeType As String
Private m_Hashes As Hashes

Public Property Get MimeType() As String
    MimeType = m_MimeType
End Property

Public Property Let MimeType(Value As String)
    m_MimeType = Value
End Property

Public Property Get Hashes() As Hashes
    Set Hashes = m_Hashes
End Property

Public Property Set Hashes(Value As Hashes)
    Set m_Hashes = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If MimeType <> vbNullString Then .Add "mimeType", MimeType
        If Not Hashes Is Nothing Then .Add "hashes", Hashes.ToJson
    End With
    Set ToJson = objJson
End Function


