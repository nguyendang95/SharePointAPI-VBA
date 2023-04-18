VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InnerError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Code As String
Private m_ErrorType As String
Private m_Message As String
Private m_StackTrack As String
Private m_ThrowSite As String

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Code(Value As String)
    m_Code = Value
End Property

Public Property Get ErrorType() As String
    ErrorType = m_ErrorType
End Property

Public Property Let ErrorType(Value As String)
    m_ErrorType = Value
End Property

Public Property Get Message() As String
    Message = m_Message
End Property

Public Property Let Message(Value As String)
    m_Message = Value
End Property

Public Property Get StackTrack() As String
    StackTrack = m_StackTrack
End Property

Public Property Let StackTrack(Value As String)
    m_StackTrack = Value
End Property

Public Property Get ThrowSite() As String
    ThrowSite = m_ThrowSite
End Property

Public Property Let ThrowSite(Value As String)
    m_ThrowSite = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Code <> vbNullString Then .Add "code", Code
        If ErrorType <> vbNullString Then .Add "errorType", ErrorType
        If Message <> vbNullString Then .Add "message", Message
        If StackTrack <> vbNullString Then .Add "stackTrack", StackTrack
        If ThrowSite <> vbNullString Then .Add "throwSite", ThrowSite
    End With
    Set ToJson = objJson
End Function
