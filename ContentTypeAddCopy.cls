VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContentTypeAddCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ContentType As String

Public Property Get ContentType() As String
   ContentType = m_ContentType
End Property

Public Property Let ContentType(Value As String)
   m_ContentType = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If ContentType <> vbNullString Then .Add "contentType", ContentType
    End With
    Set ToJson = objJson
End Function
