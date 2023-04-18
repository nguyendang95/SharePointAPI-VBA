VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Response"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Id As String
Private m_Status As Integer
Private m_Headers As Scripting.Dictionary
Private m_Body As Object

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Property Get Status() As Integer
    Status = m_Status
End Property

Public Property Let Status(Value As Integer)
    m_Status = Value
End Property

Public Property Get Headers() As Scripting.Dictionary
    Set Headers = m_Headers
End Property

Public Property Set Headers(Value As Scripting.Dictionary)
    Set m_Headers = Value
End Property

Public Property Get Body() As Object
    Set Body = m_Body
End Property

Public Property Set Body(Value As Object)
    Set m_Body = Value
End Property

