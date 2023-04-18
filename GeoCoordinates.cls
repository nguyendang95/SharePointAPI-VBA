VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeoCoordinates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Altitude As Variant
Private m_Latitude As Variant
Private m_Longitude As Variant

Public Property Get Altitude() As Variant
    Altitude = m_Altitude
End Property

Public Property Let Altitude(Value As Variant)
    m_Altitude = Value
End Property

Public Property Get Latitude() As Variant
    Latitude = m_Latitude
End Property

Public Property Let Latitude(Value As Variant)
    m_Latitude = Value
End Property

Public Property Get Longitude() As Variant
    Longitude = m_Longitude
End Property

Public Property Let Longitude(Value As Variant)
    m_Longitude = Value
End Property



