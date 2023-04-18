VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Deleted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_State As String

Public Property Get State() As String
    State = m_State
End Property

Public Property Let State(Value As String)
    m_State = Value
End Property





