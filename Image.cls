VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Width As Variant
Private m_Height As Variant

Public Property Get Width() As Variant
    Width = m_Width
End Property

Public Property Let Width(Value As Variant)
    m_Width = Value
End Property

Public Property Get Height() As Variant
    Height = m_Height
End Property

Public Property Let Height(Value As Variant)
    m_Height = Value
End Property



