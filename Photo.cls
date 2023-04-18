VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Photo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_CameraMake As String
Private m_CameraModel As String
Private m_ExposureDenominator As Variant
Private m_ExposureNumerator As Variant
Private m_FNumber As Variant
Private m_FocalLength As Variant
Private m_Iso As Variant
Private m_Orientation As Variant
Private m_TakenDateTime As String

Public Property Get CameraMake() As String
    CameraMake = m_CameraMake
End Property

Public Property Let CameraMake(Value As String)
    m_CameraMake = Value
End Property

Public Property Get CameraModel() As String
    CameraModel = m_CameraModel
End Property

Public Property Let CameraModel(Value As String)
    m_CameraModel = Value
End Property

Public Property Get ExposureDenominator() As Variant
    ExposureDenominator = m_ExposureDenominator
End Property

Public Property Let ExposureDenominator(Value As Variant)
    m_ExposureDenominator = Value
End Property

Public Property Get ExposureNumerator() As Variant
    ExposureNumerator = m_ExposureNumerator
End Property

Public Property Let ExposureNumerator(Value As Variant)
    m_ExposureNumerator = Value
End Property

Public Property Get FNumber() As Variant
    FNumber = m_FNumber
End Property

Public Property Let FNumber(Value As Variant)
    m_FNumber = Value
End Property

Public Property Get FocalLength() As Variant
    FocalLength = m_FocalLength
End Property

Public Property Let FocalLength(Value As Variant)
    m_FocalLength = Value
End Property

Public Property Get Iso() As Variant
    Iso = m_Iso
End Property

Public Property Let Iso(Value As Variant)
    m_Iso = Value
End Property

Public Property Get Orientation() As Variant
    Orientation = m_Orientation
End Property

Public Property Let Orientation(Value As Variant)
    m_Orientation = Value
End Property

Public Property Get TakenDateTime() As String
    TakenDateTime = m_TakenDateTime
End Property

Public Property Let TakenDateTime(Value As String)
    m_TakenDateTime = Value
End Property



