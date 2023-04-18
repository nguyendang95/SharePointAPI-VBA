VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalculatedColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Format As String
Private m_Formula As String
Private m_OutputType As String

Public Property Get Format() As String
   Format = m_Format
End Property

Public Property Let Format(Value As String)
   m_Format = Value
End Property

Public Property Get Formula() As String
   Formula = m_Formula
End Property

Public Property Let Formula(Value As String)
   m_Formula = Value
End Property

Public Property Get OutputType() As String
   OutputType = m_OutputType
End Property

Public Property Let OutputType(Value As String)
   m_OutputType = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Format <> vbNullString Then .Add "format", Format
        If Formula <> vbNullString Then .Add "formula", Formula
        If OutputType <> vbNullString Then .Add "outputType", OutputType
    End With
    Set ToJson = objJson
End Function

