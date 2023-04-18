VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Identity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DisplayName As String
Private m_Id As String

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

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If DisplayName <> vbNullString Then .Add "displayName", DisplayName
        If Id <> vbNullString Then .Add "id", Id
    End With
    Set ToJson = objJson
End Function

