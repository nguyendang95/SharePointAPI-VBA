VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeStamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Date As String
Private m_Time As String
Private m_TimeZone As String

Public Property Get TimeStampDate() As String
    TimeStampDate = m_Date
End Property

Public Property Let TimeStampDate(Value As String)
    m_Date = Value
End Property

Public Property Get Time() As String
    Time = m_Time
End Property

Public Property Let Time(Value As String)
    m_Time = Value
End Property

Public Property Get TimeZone() As String
    TimeZone = m_TimeZone
End Property

Public Property Let TimeZone(Value As String)
    m_TimeZone = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Date <> vbNullString Then .Add "date", Date
        If Time <> vbNullString Then .Add "time", Time
        If TimeZone <> vbNullString Then .Add "timeZone", TimeZone
    End With
    Set ToJson = objJson
End Function

