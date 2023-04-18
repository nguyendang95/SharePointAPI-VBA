VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IncompleteData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_MissingDataBeforeDateTime As String
Private m_WasThrottled As Variant

Public Property Get MissingDataBeforeDateTime() As String
    MissingDataBeforeDateTime = m_MissingDataBeforeDateTime
End Property

Public Property Let MissingDataBeforeDateTime(Value As String)
    m_MissingDataBeforeDateTime = Value
End Property

Public Property Get WasThrottled() As Variant
    WasThrottled = m_WasThrottled
End Property

Public Property Let WasThrottled(Value As Variant)
    m_WasThrottled = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If MissingDataBeforeDateTime <> vbNullString Then .Add "missingDataBeforeDateTime", MissingDataBeforeDateTime
        If IsEmpty(WasThrottled) = False Then .Add "WasThrottled", WasThrottled
    End With
    Set ToJson = objJson
End Function


