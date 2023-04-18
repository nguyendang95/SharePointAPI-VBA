VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DefaultLanguage As String
Private m_Descriptions As New Collection
Private m_Formula As String

Public Property Get DefaultLanguage() As String
   DefaultLanguage = m_DefaultLanguage
End Property

Public Property Let DefaultLanguage(Value As String)
   m_DefaultLanguage = Value
End Property

Public Property Get Descriptions() As Collection
  Set Descriptions = m_Descriptions
End Property

Public Property Set Descriptions(Value As Collection)
  Set m_Descriptions = Value
End Property

Public Property Get Formula() As String
   Formula = m_Formula
End Property

Public Property Let Formula(Value As String)
   m_Formula = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If DefaultLanguage <> vbNullString Then .Add "defaultLanguage", DefaultLanguage
        If Descriptions.Count > 0 Then .Add "descriptions", Descriptions
        If Formula <> vbNullString Then .Add "formula", Formula
    End With
    Set ToJson = objJson
End Function

