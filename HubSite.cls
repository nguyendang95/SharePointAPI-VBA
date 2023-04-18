VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HubSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_HubSiteUrls As New Collection
Private m_PropagateToExistingLists As Variant

Public Property Get HubSiteUrls() As Collection
   Set HubSiteUrls = m_HubSiteUrls
End Property

Public Property Set HubSiteUrls(Value As Collection)
   Set m_HubSiteUrls = Value
End Property

Public Property Get PropagateToExistingLists() As Variant
   PropagateToExistingLists = m_PropagateToExistingLists
End Property

Public Property Let PropagateToExistingLists(Value As Variant)
   m_PropagateToExistingLists = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If HubSiteUrls.Count > 0 Then .Add "hubSiteUrls", HubSiteUrls
        If IsEmpty(PropagateToExistingLists) = False Then .Add "propagateToExistingLists", PropagateToExistingLists
    End With
    Set ToJson = objJson
End Function
