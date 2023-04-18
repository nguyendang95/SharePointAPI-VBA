VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_DisplayAs As String
Private m_Format As String

Public Property Get DisplayAs() As String
   DisplayAs = m_DisplayAs
End Property

Public Property Let DisplayAs(Value As String)
   m_DisplayAs = Value
End Property

Public Property Get Format() As String
   Format = m_Format
End Property

Public Property Let Format(Value As String)
   m_Format = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If DisplayAs <> vbNullString Then .Add "displayAs", DisplayAs
        If Format <> vbNullString Then .Add "format", Format
    End With
    Set ToJson = objJson
End Function

