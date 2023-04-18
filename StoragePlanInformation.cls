VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StoragePlanInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_UpgradeAvailable As Variant

Public Property Get UpgradeAvailable() As Variant
    UpgradeAvailable = m_UpgradeAvailable
End Property

Public Property Let UpgradeAvailable(Value As Variant)
    m_UpgradeAvailable = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If IsEmpty(UpgradeAvailable) = False Then .Add "UpgradeAvailable", UpgradeAvailable
    End With
    Set ToJson = objJson
End Function

