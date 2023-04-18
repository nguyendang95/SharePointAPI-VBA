VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Request"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_DependsOn As String
Private m_Method As String
Private m_Url As String
Private m_Headers As String
Private m_Body As String

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get DependsOn() As String
    DependsOn = m_DependsOn
End Property

Public Property Let DependsOn(Value As String)
    m_DependsOn = Value
End Property

Public Property Get Method() As String
    Method = m_Method
End Property

Public Property Let Method(Value As String)
    m_Method = Value
End Property

Public Property Get Url() As String
    Url = m_Url
End Property

Public Property Let Url(Value As String)
    m_Url = Value
End Property

Public Property Get Headers() As String
    Headers = m_Headers
End Property

Public Property Let Headers(Value As String)
    m_Headers = Value
End Property

Public Property Get Body() As String
    Body = m_Body
End Property

Public Property Let Body(Value As String)
    m_Body = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Id <> vbNullString Then .Add "id", Id
        If DependsOn <> vbNullString Then .Add "dependsOn", DependsOn
        If Method <> vbNullString Then .Add "method", Method
        If Url <> vbNullString Then .Add "url", Url
        If Headers <> vbNullString Then .Add "headers", Headers
        If Body <> vbNullString Then .Add "body", Body
    End With
    Set ToJson = objJson
End Function
