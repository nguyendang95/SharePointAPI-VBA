VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DriveItemUploadableProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Description As String
Private m_FileSize As Variant
Private m_FileSystemInfo As FileSystemInfo
Private m_Name As String
Private m_ConflictBehavior As String

Public Property Get ConflictBehavior() As String
    ConflictBehavior = m_ConflictBehavior
End Property

Public Property Let ConflictBehavior(Value As String)
    m_ConflictBehavior = Value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(Value As String)
    m_Description = Value
End Property

Public Property Get FileSize() As Variant
    Set FileSize = m_FileSize
End Property

Public Property Set FileSize(Value As Variant)
    Set m_FileSize = Value
End Property

Public Property Get FileSystemInfo() As FileSystemInfo
    Set FileSystemInfo = m_FileSystemInfo
End Property

Public Property Set FileSystemInfo(Value As FileSystemInfo)
    Set m_FileSystemInfo = Value
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(Value As String)
    m_Name = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Description <> vbNullString Then .Add "description", Description
        If IsEmpty(FileSize) = False Then .Add "fileSize", FileSize
        If Not FileSystemInfo Is Nothing Then .Add "fileSystemInfo", FileSystemInfo.ToJson
        If Name <> vbNullString Then .Add "name", Name
        If ConflictBehavior <> vbNullString Then .Add "@microsoft.graph.conflictBehavior", ConflictBehavior
    End With
    Set ToJson = objJson
End Function

