VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SiteCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Hostname As String
Private m_DataLocationCode As String
Private m_Root As Root

Public Property Get Hostname() As String
   Hostname = m_Hostname
End Property

Public Property Let Hostname(Value As String)
   m_Hostname = Value
End Property

Public Property Get DataLocationCode() As String
   DataLocationCode = m_DataLocationCode
End Property

Public Property Let DataLocationCode(Value As String)
   m_DataLocationCode = Value
End Property

Public Property Get Root() As Root
  Set Root = m_Root
End Property

Public Property Set Root(Value As Root)
  Set m_Root = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Hostname <> vbNullString Then .Add "hostname", Hostname
        If DataLocationCode <> vbNullString Then .Add "dataLocationCode", DataLocationCode
        If Not Root Is Nothing Then .Add "root", Root.ToJson
    End With
    Set ToJson = objJson
End Function

