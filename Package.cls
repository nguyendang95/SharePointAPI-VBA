VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Package"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PackageType As String

Public Property Get PackageType() As String
    PackageType = m_PackageType
End Property

Public Property Let PackageType(Value As String)
    m_PackageType = Value
End Property



