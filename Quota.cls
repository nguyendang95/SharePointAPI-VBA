VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Quota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Deleted As Variant
Private m_Remaining As Variant
Private m_State As String
Private m_StoragePlanInformation As StoragePlanInformation
Private m_Total As Variant
Private m_Used As Variant

Public Property Get Deleted() As Variant
    Deleted = m_Deleted
End Property

Public Property Let Deleted(Value As Variant)
    m_Deleted = Value
End Property

Public Property Get Remaining() As Variant
    Remaining = m_Remaining
End Property

Public Property Let Remaining(Value As Variant)
    m_Remaining = Value
End Property

Public Property Get State() As String
    State = m_State
End Property

Public Property Let State(Value As String)
    m_State = Value
End Property

Public Property Get StoragePlanInformation() As StoragePlanInformation
    Set StoragePlanInformation = m_StoragePlanInformation
End Property

Public Property Set StoragePlanInformation(Value As StoragePlanInformation)
    Set m_StoragePlanInformation = Value
End Property

Public Property Get Total() As Variant
    Total = m_Total
End Property

Public Property Let Total(Value As Variant)
    m_Total = Value
End Property

Public Property Get Used() As Variant
    Used = m_Used
End Property

Public Property Let Used(Value As Variant)
    m_Used = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(Deleted) = False Then .Add "Deleted", Deleted
        If IsEmpty(Remaining) = False Then .Add "Remaining", Remaining
        If State <> vbNullString Then .Add "state", State
        If Not StoragePlanInformation Is Nothing Then .Add "storagePlanInformation", StoragePlanInformation.ToJson
        If IsEmpty(Total) = False Then .Add "Total", Total
        If IsEmpty(Used) = False Then .Add "Used", Used
    End With
    Set ToJson = objJson
End Function


