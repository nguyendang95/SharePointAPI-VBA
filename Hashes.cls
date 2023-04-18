VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hashes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Crc32Hash As String
Private m_QuickXorHash As String
Private m_Sha1Hash As String
Private m_Sha256Hash As String

Public Property Get Crc32Hash() As String
    Crc32Hash = m_Crc32Hash
End Property

Public Property Let Crc32Hash(Value As String)
    m_Crc32Hash = Value
End Property

Public Property Get QuickXorHash() As String
    QuickXorHash = m_QuickXorHash
End Property

Public Property Let QuickXorHash(Value As String)
    m_QuickXorHash = Value
End Property

Public Property Get Sha1Hash() As String
    Sha1Hash = m_Sha1Hash
End Property

Public Property Let Sha1Hash(Value As String)
    m_Sha1Hash = Value
End Property

Public Property Get Sha256Hash() As String
    Sha256Hash = m_Sha256Hash
End Property

Public Property Let Sha256Hash(Value As String)
    m_Sha256Hash = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Crc32Hash <> vbNullString Then .Add "crc32Hash", Crc32Hash
        If QuickXorHash <> vbNullString Then .Add "quickXorHash", QuickXorHash
        If Sha1Hash <> vbNullString Then .Add "sha1Hash", Sha1Hash
        If Sha256Hash <> vbNullString Then .Add "sha256Hash", Sha256Hash
    End With
    Set ToJson = objJson
End Function



