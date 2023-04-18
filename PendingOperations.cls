VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PendingOperations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PendingContentUpdate As PendingContentUpdate

Public Property Get PendingContentUpdate() As PendingContentUpdate
    Set PendingContentUpdate = m_PendingContentUpdate
End Property

Public Property Set PendingContentUpdate(Value As PendingContentUpdate)
    Set m_PendingContentUpdate = Value
End Property



