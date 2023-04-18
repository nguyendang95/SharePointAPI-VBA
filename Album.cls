VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Album"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CoveredImageItemId As String

Public Property Get CoveredImageItemId() As String
    CoveredImageItemId = m_CoveredImageItemId
End Property

Public Property Let CoveredImageItemId(Value As String)
    m_CoveredImageItemId = Value
End Property




