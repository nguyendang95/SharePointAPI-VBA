VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemActivityStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Activities As New Collection
Private m_IncompletedData As IncompleteData
Private m_IsTrending As Variant
Private m_StartDateTime As String
Private m_EndDateTime As String
Private m_Create As ItemActionStat
Private m_Delete As ItemActionStat
Private m_Edit As ItemActionStat
Private m_Move As ItemActionStat
Private m_Access As ItemActionStat


Public Property Get Activities() As Collection
    Set Activities = m_Activities
End Property

Public Property Set Activities(Value As Collection)
    Set m_Activities = Value
End Property

Public Property Get IncompletedData() As IncompleteData
    Set IncompletedData = m_IncompletedData
End Property

Public Property Set IncompletedData(Value As IncompleteData)
    Set m_IncompletedData = Value
End Property

Public Property Get IsTrending() As Variant
    IsTrending = m_IsTrending
End Property

Public Property Let IsTrending(Value As Variant)
    m_IsTrending = Value
End Property

Public Property Get StartDateTime() As String
    StartDateTime = m_StartDateTime
End Property

Public Property Let StartDateTime(Value As String)
    m_StartDateTime = Value
End Property

Public Property Get EndDateTime() As String
    EndDateTime = m_EndDateTime
End Property

Public Property Let EndDateTime(Value As String)
    m_EndDateTime = Value
End Property

Public Property Get Create() As ItemActionStat
    Set Create = m_Create
End Property

Public Property Set Create(Value As ItemActionStat)
    Set m_Create = Value
End Property

Public Property Get Delete() As ItemActionStat
    Set Delete = m_Delete
End Property

Public Property Set Delete(Value As ItemActionStat)
    Set m_Delete = Value
End Property

Public Property Get Edit() As ItemActionStat
    Set Edit = m_Edit
End Property

Public Property Set Edit(Value As ItemActionStat)
    Set m_Edit = Value
End Property

Public Property Get Move() As ItemActionStat
    Set Move = m_Move
End Property

Public Property Set Move(Value As ItemActionStat)
    Set m_Move = Value
End Property

Public Property Get Access() As ItemActionStat
    Set Access = m_Access
End Property

Public Property Set Access(Value As ItemActionStat)
    Set m_Access = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Activities.Count > 0 Then .Add "activities", Activities
        If Not IncompletedData Is Nothing Then .Add "incompletedData", IncompletedData.ToJson
        If IsEmpty(IsTrending) = False Then .Add "IsTrending", IsTrending
        If StartDateTime <> vbNullString Then .Add "startDateTime", StartDateTime
        If EndDateTime <> vbNullString Then .Add "endDateTime", EndDateTime
        If Not Create Is Nothing Then .Add "create", Create.ToJson
        If Not Delete Is Nothing Then .Add "delete", Delete.ToJson
        If Not Edit Is Nothing Then .Add "edit", Edit.ToJson
        If Not Move Is Nothing Then .Add "move", Move.ToJson
        If Not Access Is Nothing Then .Add "access", Access.ToJson
    End With
    Set ToJson = objJson
End Function

