VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Code As String
Private m_Message As String
Private m_InnerError As InnerError

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

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

Public Property Get InnerError() As InnerError
    Set InnerError = m_InnerError
End Property

Public Property Set InnerError(Value As InnerError)
    Set m_InnerError = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If Code <> vbNullString Then .Add "code", Code
        If Message <> vbNullString Then .Add "message", Message
        If Not InnerError Is Nothing Then .Add "innerError", InnerError.ToJson
    End With
    Set ToJson = objJson
End Function


