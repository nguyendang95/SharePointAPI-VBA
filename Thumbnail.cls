VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Thumbnail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Height As Variant
Private m_SourceItemId As String
Private m_Url As Variant
Private m_Width As Variant

Public Property Get Height() As Variant
    Height = m_Height
End Property

Public Property Let Height(Value As Variant)
    m_Height = Value
End Property

Public Property Get SourceItemId() As String
    SourceItemId = m_SourceItemId
End Property

Public Property Let SourceItemId(Value As String)
    m_SourceItemId = Value
End Property

Public Property Get Url() As Variant
    Url = m_Url
End Property

Public Property Let Url(Value As Variant)
    m_Url = Value
End Property

Public Property Get Width() As Variant
    Width = m_Width
End Property

Public Property Let Width(Value As Variant)
    m_Width = Value
End Property





