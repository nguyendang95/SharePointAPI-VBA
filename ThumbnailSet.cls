VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThumbnailSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Large As Thumbnail
Private m_Medium As Thumbnail
Private m_Small As Thumbnail
Private m_Source As Thumbnail

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get Large() As Thumbnail
    Set Large = m_Large
End Property

Public Property Set Large(Value As Thumbnail)
    Set m_Large = Value
End Property

Public Property Get Medium() As Thumbnail
    Set Medium = m_Medium
End Property

Public Property Set Medium(Value As Thumbnail)
    Set m_Medium = Value
End Property

Public Property Get Small() As Thumbnail
    Set Small = m_Small
End Property

Public Property Set Small(Value As Thumbnail)
    Set m_Small = Value
End Property

Public Property Get Source() As Thumbnail
    Set Source = m_Source
End Property

Public Property Set Source(Value As Thumbnail)
    Set m_Source = Value
End Property



