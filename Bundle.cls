VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Bundle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Album As Album
Private m_ChildCount As Variant

Public Property Get Album() As Album
    Set Album = m_Album
End Property

Public Property Set Album(Value As Album)
    Set m_Album = m_Album
End Property

Public Property Get ChildCount() As Variant
    ChildCount = m_ChildCount
End Property

Public Property Let ChildCount(Value As Variant)
    m_ChildCount = m_ChildCount
End Property

