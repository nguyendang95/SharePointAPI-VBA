VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SearchResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_OnClickTelemetryUrl As String

Public Property Get OnClickTelemetryUrl() As String
    OnClickTelemetryUrl = m_OnClickTelemetryUrl
End Property

Public Property Let OnClickTelemetryUrl(Value As String)
    m_OnClickTelemetryUrl = Value
End Property



