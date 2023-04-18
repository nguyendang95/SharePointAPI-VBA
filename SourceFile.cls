VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SourceFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_SourceFile As ItemReference
Private m_DestinationFileName As String

Public Property Get SourceFile() As ItemReference
   Set SourceFile = m_SourceFile
End Property

Public Property Set SourceFile(Value As ItemReference)
   Set m_SourceFile = Value
End Property

Public Property Get DestinationFileName() As String
   DestinationFileName = m_DestinationFileName
End Property

Public Property Let DestinationFileName(Value As String)
   m_DestinationFileName = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not SourceFile Is Nothing Then .Add "sourceFile", SourceFile.ToJson
        If DestinationFileName <> vbNullString Then .Add "destinationFileName", DestinationFileName
    End With
    Set ToJson = objJson
End Function

