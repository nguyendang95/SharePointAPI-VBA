VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HyperlinkOrPictureColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_IsPicture As Variant

Public Property Get IsPicture() As Variant
   IsPicture = m_IsPicture
End Property

Public Property Let IsPicture(Value As Variant)
   m_IsPicture = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(IsPicture) = False Then .Add "IsPicture", IsPicture
    End With
    Set ToJson = objJson
End Function

