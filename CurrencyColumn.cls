VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CurrencyColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Locale As String

Public Property Get Locale() As String
   Locale = m_Locale
End Property

Public Property Let Locale(Value As String)
   m_Locale = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Locale <> vbNullString Then .Add "locale", Locale
    End With
    Set ToJson = objJson
End Function

