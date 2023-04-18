VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SortBy As String
Private m_SortOrder As String
Private m_ViewType As String

Public Property Get SortBy() As String
    SortBy = m_SortBy
End Property

Public Property Let SortBy(Value As String)
    m_SortBy = Value
End Property

Public Property Get SortOrder() As String
    SortOrder = m_SortOrder
End Property

Public Property Let SortOrder(Value As String)
    m_SortOrder = Value
End Property

Public Property Get ViewType() As String
    ViewType = m_ViewType
End Property

Public Property Let ViewType(Value As String)
    m_ViewType = Value
End Property


Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If SortBy <> vbNullString Then .Add "sortBy", SortBy
        If SortOrder <> vbNullString Then .Add "sortOrder", SortOrder
        If ViewType <> vbNullString Then .Add "viewType", ViewType
    End With
    Set ToJson = objJson
End Function
