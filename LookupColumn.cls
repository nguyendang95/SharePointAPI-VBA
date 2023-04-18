VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LookupColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AllowMultipleValues As Variant
Private m_AllowUnlimitedLength As Variant
Private m_ColumnName As String
Private m_ListId As String
Private m_PrimaryLookupColumnId As String

Public Property Get AllowMultipleValues() As Variant
   AllowMultipleValues = m_AllowMultipleValues
End Property

Public Property Let AllowMultipleValues(Value As Variant)
   m_AllowMultipleValues = Value
End Property

Public Property Get AllowUnlimitedLength() As Variant
   AllowUnlimitedLength = m_AllowUnlimitedLength
End Property

Public Property Let AllowUnlimitedLength(Value As Variant)
   m_AllowUnlimitedLength = Value
End Property

Public Property Get ColumnName() As String
   ColumnName = m_ColumnName
End Property

Public Property Let ColumnName(Value As String)
   m_ColumnName = Value
End Property

Public Property Get ListId() As String
   ListId = m_ListId
End Property

Public Property Let ListId(Value As String)
   m_ListId = Value
End Property

Public Property Get PrimaryLookupColumnId() As String
   PrimaryLookupColumnId = m_PrimaryLookupColumnId
End Property

Public Property Let PrimaryLookupColumnId(Value As String)
   m_PrimaryLookupColumnId = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(AllowMultipleValues) = False Then .Add "AllowMultipleValues", AllowMultipleValues
        If IsEmpty(AllowUnlimitedLength) = False Then .Add "AllowUnlimitedLength", AllowUnlimitedLength
        If ColumnName <> vbNullString Then .Add "columnName", ColumnName
        If ListId <> vbNullString Then .Add "listId", ListId
        If PrimaryLookupColumnId <> vbNullString Then .Add "primaryLookupColumnId", PrimaryLookupColumnId
    End With
    Set ToJson = objJson
End Function

