VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CheckInParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CheckInAs As String
Private m_Comment As String

Public Property Get CheckInAs() As String
    CheckInAs = m_CheckInAs
End Property

Public Property Let CheckInAs(Value As String)
    m_CheckInAs = Value
End Property

Public Property Get Comment() As String
    Comment = m_Comment
End Property

Public Property Let Comment(Value As String)
    m_Comment = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If CheckInAs <> vbNullString Then .Add "checkInAs", CheckInAs
        If Comment <> vbNullString Then .Add "comment", Comment
    End With
    Set ToJson = objJson
End Function


