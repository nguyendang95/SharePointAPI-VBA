VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemAnalytics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AllTime As ItemActivityStat
Private m_LastSevenDays As ItemActivityStat

Public Property Get AllTime() As ItemActivityStat
    Set AllTime = m_AllTime
End Property

Public Property Set AllTime(Value As ItemActivityStat)
    Set m_AllTime = Value
End Property

Public Property Get LastSevenDays() As ItemActivityStat
    Set LastSevenDays = m_LastSevenDays
End Property

Public Property Set LastSevenDays(Value As ItemActivityStat)
    Set m_LastSevenDays = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not AllTime Is Nothing Then .Add "allTime", AllTime.ToJson
        If Not LastSevenDays Is Nothing Then .Add "lastSevenDays", LastSevenDays.ToJson
    End With
    Set ToJson = objJson
End Function

