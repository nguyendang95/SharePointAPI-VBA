VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IdentitySet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Application As Identity
Private m_Device As Identity
Private m_User As Identity

Public Property Get Application() As Identity
  Set Application = m_Application
End Property

Public Property Set Application(Value As Identity)
  Set m_Application = Value
End Property

Public Property Get Device() As Identity
  Set Device = m_Device
End Property

Public Property Set Device(Value As Identity)
  Set m_Device = Value
End Property

Public Property Get User() As Identity
  Set User = m_User
End Property

Public Property Set User(Value As Identity)
  Set m_User = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If Not Application Is Nothing Then .Add "application", Application.ToJson
        If Not Device Is Nothing Then .Add "device", Device.ToJson
        If Not User Is Nothing Then .Add "user", User.ToJson
    End With
    Set ToJson = objJson
End Function

