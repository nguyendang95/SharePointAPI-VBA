VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AllowMultipleLines As Variant
Private m_AppendChangesToExistingText As Variant
Private m_LinesForEditing As Variant
Private m_MaxLength As Variant
Private m_TextType As String

Public Property Get AllowMultipleLines() As Variant
   AllowMultipleLines = m_AllowMultipleLines
End Property

Public Property Let AllowMultipleLines(Value As Variant)
   m_AllowMultipleLines = Value
End Property

Public Property Get AppendChangesToExistingText() As Variant
   AppendChangesToExistingText = m_AppendChangesToExistingText
End Property

Public Property Let AppendChangesToExistingText(Value As Variant)
   m_AppendChangesToExistingText = Value
End Property

Public Property Get LinesForEditing() As Variant
   LinesForEditing = m_LinesForEditing
End Property

Public Property Let LinesForEditing(Value As Variant)
   m_LinesForEditing = Value
End Property

Public Property Get MaxLength() As Variant
   MaxLength = m_MaxLength
End Property

Public Property Let MaxLength(Value As Variant)
   m_MaxLength = Value
End Property

Public Property Get TextType() As String
   TextType = m_TextType
End Property

Public Property Let TextType(Value As String)
   m_TextType = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(AllowMultipleLines) = False Then .Add "allowMultipleLines", AllowMultipleLines
        If IsEmpty(AppendChangesToExistingText) = False Then .Add "appendChangesToExistingText", AppendChangesToExistingText
        If IsEmpty(LinesForEditing) = False Then .Add "linesForEditing", LinesForEditing
        If IsEmpty(MaxLength) = False Then .Add "maxLength", MaxLength
        If TextType <> vbNullString Then .Add "textType", TextType
    End With
    Set ToJson = objJson
End Function

