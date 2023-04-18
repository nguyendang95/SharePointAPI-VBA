VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RichLongRunningOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CreatedDateTime As String
Private m_Error As PublicError
Private m_Id As String
Private m_LastActionDateTime As String
Private m_PercentageComplete As Variant
Private m_ResourceId As String
Private m_ResourceLocation As String
Private m_Status As String
Private m_StatusDetail As String
Private m_Type As String

Public Property Get CreatedDateTime() As String
    CreatedDateTime = m_CreatedDateTime
End Property

Public Property Let CreatedDateTime(Value As String)
    m_CreatedDateTime = Value
End Property

Public Property Get Error() As PublicError
    Set Error = m_Error
End Property

Public Property Set Error(Value As PublicError)
    Set m_Error = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get LastActionDateTime() As String
    LastActionDateTime = m_LastActionDateTime
End Property

Public Property Let LastActionDateTime(Value As String)
    m_LastActionDateTime = Value
End Property

Public Property Get PercentageComplete() As Variant
    PercentageComplete = m_PercentageComplete
End Property

Public Property Let PercentageComplete(Value As Variant)
    m_PercentageComplete = Value
End Property

Public Property Get ResourceId() As String
    ResourceId = m_ResourceId
End Property

Public Property Let ResourceId(Value As String)
    m_ResourceId = Value
End Property

Public Property Get ResourceLocation() As String
    ResourceLocation = m_ResourceLocation
End Property

Public Property Let ResourceLocation(Value As String)
    m_ResourceLocation = Value
End Property

Public Property Get Status() As String
    Status = m_Status
End Property

Public Property Let Status(Value As String)
    m_Status = Value
End Property

Public Property Get StatusDetail() As String
    StatusDetail = m_StatusDetail
End Property

Public Property Let StatusDetail(Value As String)
    m_StatusDetail = Value
End Property

Public Property Get RichLongRunningOperationType() As String
    RichLongRunningOperationType = m_Type
End Property

Public Property Let RichLongRunningOperationType(Value As String)
    m_Type = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If CreatedDateTime <> vbNullString Then .Add "createdDateTime", CreatedDateTime
        If Not Error Is Nothing Then .Add "error", Error.ToJson
        If Id <> vbNullString Then .Add "id", Id
        If LastActionDateTime <> vbNullString Then .Add "lastActionDateTime", LastActionDateTime
        If IsEmpty(PercentageComplete) = False Then .Add "PercentageComplete", PercentageComplete
        If ResourceId <> vbNullString Then .Add "resourceId", ResourceId
        If ResourceLocation <> vbNullString Then .Add "resourceLocation", ResourceLocation
        If Status <> vbNullString Then .Add "status", Status
        If StatusDetail <> vbNullString Then .Add "statusDetail", StatusDetail
        If RichLongRunningOperationType <> vbNullString Then .Add "type", RichLongRunningOperationType
    End With
    Set ToJson = objJson
End Function


