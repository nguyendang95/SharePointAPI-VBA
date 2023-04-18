VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BatchRequestResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Responses As New Collection

Public Property Get Responses() As Collection
    Set Responses = m_Responses
End Property

Public Property Set Responses(Value As Collection)
    Set m_Responses = Value
End Property



