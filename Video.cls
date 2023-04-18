VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Video"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_AudioBitsPerSample As Variant
Private m_AudioChannels As Variant
Private m_AudioFormat As String
Private m_AudioSamplesPerSecond As Variant
Private m_Bitrate As Variant
Private m_Duration As Variant
Private m_FourCC As String
Private m_FrameRate As Variant
Private m_Height As Variant
Private m_Width As Variant

Public Property Get AudioBitsPerSample() As Variant
    AudioBitsPerSample = m_AudioBitsPerSample
End Property

Public Property Let AudioBitsPerSample(Value As Variant)
    m_AudioBitsPerSample = Value
End Property

Public Property Get AudioChannels() As Variant
    AudioChannels = m_AudioChannels
End Property

Public Property Let AudioChannels(Value As Variant)
    m_AudioChannels = Value
End Property

Public Property Get AudioFormat() As String
    AudioFormat = m_AudioFormat
End Property

Public Property Let AudioFormat(Value As String)
    m_AudioFormat = Value
End Property

Public Property Get AudioSamplesPerSecond() As Variant
    AudioSamplesPerSecond = m_AudioSamplesPerSecond
End Property

Public Property Let AudioSamplesPerSecond(Value As Variant)
    m_AudioSamplesPerSecond = Value
End Property

Public Property Get Bitrate() As Variant
    Bitrate = m_Bitrate
End Property

Public Property Let Bitrate(Value As Variant)
    m_Bitrate = Value
End Property

Public Property Get Duration() As Variant
    Duration = m_Duration
End Property

Public Property Let Duration(Value As Variant)
    m_Duration = Value
End Property

Public Property Get FourCC() As String
    FourCC = m_FourCC
End Property

Public Property Let FourCC(Value As String)
    m_FourCC = Value
End Property

Public Property Get FrameRate() As Variant
    FrameRate = m_FrameRate
End Property

Public Property Let FrameRate(Value As Variant)
    m_FrameRate = Value
End Property

Public Property Get Height() As Variant
    Height = m_Height
End Property

Public Property Let Height(Value As Variant)
    m_Height = Value
End Property

Public Property Get Width() As Variant
    Width = m_Width
End Property

Public Property Let Width(Value As Variant)
    m_Width = Value
End Property



