VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldValueSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Fields As Scripting.Dictionary

Public Property Get Fields() As Scripting.Dictionary
    Set Fields = m_Fields
End Property

Public Property Let Fields(Value As Scripting.Dictionary)
    Set m_Fields = Value
End Property

