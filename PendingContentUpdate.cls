VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PendingContentUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_QueuedDateTime As String

Public Property Get QueuedDateTime() As String
    QueuedDateTime = m_QueuedDateTime
End Property

Public Property Let QueuedDateTime(Value As String)
    m_QueuedDateTime = Value
End Property



