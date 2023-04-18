VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DriveFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_ChildCount As Variant
Private m_View As FolderView

Public Property Get ChildCount() As Variant
   ChildCount = m_ChildCount
End Property

Public Property Let ChildCount(Value As Variant)
   m_ChildCount = Value
End Property

Public Property Get View() As FolderView
    Set View = m_View
End Property

Public Property Set View(Value As FolderView)
    Set m_View = Value
End Property


Public Function ToJson() As Variant
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(ChildCount) = False Then .Add "name", ChildCount
        If Not View Is Nothing Then .Add "view", View.ToJson
    End With
    Set ToJson = objJson
End Function
