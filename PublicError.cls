VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PublicError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Code As String
Private m_Details As New Collection
Private m_InnerError As PublicInnerError
Private m_Message As String
Private m_Target As String

Public Property Get Code() As String
    Code = m_Code
End Property

Public Property Let Code(Value As String)
    m_Code = Value
End Property

Public Property Get Details() As Collection
    Set Details = m_Details
End Property

Public Property Set Details(Value As Collection)
    Set m_Details = Value
End Property

Public Property Get InnerError() As PublicInnerError
    Set InnerError = m_InnerError
End Property

Public Property Set InnerError(Value As PublicInnerError)
    Set m_InnerError = Value
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
        If Details.Count > 0 Then .Add "details", Details
        If Not InnerError Is Nothing Then .Add "innerError", InnerError.ToJson
        If Message <> vbNullString Then .Add "message", Message
        If Target <> vbNullString Then .Add "target", Target
    End With
    Set ToJson = objJson
End Function

