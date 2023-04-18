VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharePointIdentity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DisplayName As String
Private m_Id As String
Private m_LoginName As String
Private m_Thumbnails As New Collection

Public Property Get DisplayName() As String
    DisplayName = m_DisplayName
End Property

Public Property Let DisplayName(Value As String)
    m_DisplayName = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get LoginName() As String
    LoginName = m_LoginName
End Property

Public Property Let LoginName(Value As String)
    m_LoginName = Value
End Property

Public Property Get Thumbnails() As Collection
    Set Thumbnails = m_Thumbnails
End Property

Public Property Set Thumbnails(Value As Collection)
    Set m_Thumbnails = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If DisplayName <> vbNullString Then .Add "displayName", DisplayName
        If Id <> vbNullString Then .Add "id", Id
        If LoginName <> vbNullString Then .Add "loginName", LoginName
    End With
    Set ToJson = objJson
End Function
