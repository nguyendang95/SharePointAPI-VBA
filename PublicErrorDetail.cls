VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PublicErrorDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Code As String
Private m_Message As String
Private m_Target As String

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Code(Value As String)
    m_Code = Value
End Property

Public Property Get Message() As String
    Message = m_Message
End Property

Public Property Let Message(Value As String)
    m_Message = Value
End Property

Public Property Get Target() As String
    Target = m_Target
End Property

Public Property Let Target(Value As String)
    m_Target = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Code <> vbNullString Then .Add "code", Code
        If Message <> vbNullString Then .Add "message", Message
        If Target <> vbNullString Then .Add "target", Target
    End With
    Set ToJson = objJson
End Function


