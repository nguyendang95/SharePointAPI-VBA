VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Audio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Album As String
Private m_AlbumArtist As String
Private m_Artist As String
Private m_Bitrate As Variant
Private m_Composers As String
Private m_Copyright As String
Private m_Disc As Variant
Private m_DiscCount As Variant
Private m_Duration As Variant
Private m_Genre As String
Private m_HasDrm As Variant
Private m_IsVariableBitrate As Variant
Private m_Title As String
Private m_Track As Variant
Private m_TrackCount As Variant
Private m_Year As Variant

Public Property Get Album() As String
    Album = m_Album
End Property

Public Property Let Album(Value As String)
    m_Album = Value
End Property

Public Property Get AlbumArtist() As String
    AlbumArtist = m_AlbumArtist
End Property

Public Property Let AlbumArtist(Value As String)
    m_AlbumArtist = Value
End Property

Public Property Get Artist() As String
    Artist = m_Artist
End Property

Public Property Let Artist(Value As String)
    m_Artist = Value
End Property

Public Property Get Bitrate() As Double
    Bitrate = m_Bitrate
End Property

Public Property Let Bitrate(Value As Double)
    m_Bitrate = Value
End Property

Public Property Get Composers() As String
    Composers = m_Composers
End Property

Public Property Let Composers(Value As String)
    m_Composers = Value
End Property

Public Property Get Copyright() As String
    Copyright = m_Copyright
End Property

Public Property Let Copyright(Value As String)
    m_Copyright = Value
End Property

Public Property Get Disc() As Integer
    Disc = m_Disc
End Property

Public Property Let Disc(Value As Integer)
    m_Disc = Value
End Property

Public Property Get DiscCount() As Integer
    DiscCount = m_DiscCount
End Property

Public Property Let DiscCount(Value As Integer)
    m_DiscCount = Value
End Property

Public Property Get Duration() As Double
    Duration = m_Duration
End Property

Public Property Let Duration(Value As Double)
    m_Duration = Value
End Property

Public Property Get Genre() As String
    Genre = m_Genre
End Property

Public Property Let Genre(Value As String)
    m_Genre = Value
End Property

Public Property Get HasDrm() As Boolean
    HasDrm = m_HasDrm
End Property

Public Property Let HasDrm(Value As Boolean)
    m_HasDrm = Value
End Property

Public Property Get IsVariableBitrate() As Boolean
    IsVariableBitrate = m_IsVariableBitrate
End Property

Public Property Let IsVariableBitrate(Value As Boolean)
    m_IsVariableBitrate = Value
End Property

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(Value As String)
    m_Title = Value
End Property

Public Property Get Track() As Long
    Track = m_Track
End Property

Public Property Let Track(Value As Long)
    m_Track = Value
End Property

Public Property Get TrackCount() As Long
    TrackCount = m_TrackCount
End Property

Public Property Let TrackCount(Value As Long)
    m_TrackCount = Value
End Property

Public Property Get Year() As Long
    Year = m_Year
End Property

Public Property Let Year(Value As Long)
    m_Year = Value
End Property
