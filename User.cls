VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_BusinessPhones As New Collection
Private m_DisplayName As String
Private m_GivenName As String
Private m_JobTitle As String
Private m_Mail As String
Private m_MobilePhone As String
Private m_OfficeLocation As String
Private m_PreferredLanguage As String
Private m_Surname As String
Private m_UserPrincipalName As String
Private m_Id As String

Public Property Get BusinessPhones() As Collection
    Set BusinessPhones = m_BusinessPhones
End Property

Public Property Set BusinessPhones(Value As Collection)
    Set m_BusinessPhones = Value
End Property

Public Property Get DisplayName() As String
    DisplayName = m_DisplayName
End Property

Public Property Let DisplayName(Value As String)
    m_DisplayName = Value
End Property

Public Property Get GivenName() As String
    GivenName = m_GivenName
End Property

Public Property Let GivenName(Value As String)
    m_GivenName = Value
End Property

Public Property Get JobTitle() As String
    JobTitle = m_JobTitle
End Property

Public Property Let JobTitle(Value As String)
    m_JobTitle = Value
End Property

Public Property Get Mail() As String
    Mail = m_Mail
End Property

Public Property Let Mail(Value As String)
    m_Mail = Value
End Property

Public Property Get MobilePhone() As String
    MobilePhone = m_MobilePhone
End Property

Public Property Let MobilePhone(Value As String)
    m_MobilePhone = Value
End Property

Public Property Get OfficeLocation() As String
    OfficeLocation = m_OfficeLocation
End Property

Public Property Let OfficeLocation(Value As String)
    m_OfficeLocation = Value
End Property

Public Property Get PreferredLanguage() As String
    PreferredLanguage = m_PreferredLanguage
End Property

Public Property Let PreferredLanguage(Value As String)
    m_PreferredLanguage = Value
End Property

Public Property Get Surname() As String
    Surname = m_Surname
End Property

Public Property Let Surname(Value As String)
    m_Surname = Value
End Property

Public Property Get UserPrincipalName() As String
    UserPrincipalName = m_UserPrincipalName
End Property

Public Property Let UserPrincipalName(Value As String)
    m_UserPrincipalName = Value
End Property

Public Property Get Id() As String
    Id = m_Id
End Property

Public Property Let Id(Value As String)
    m_Id = Value
End Property

Public Function ToJson() As Scripting.Dictionary
    Dim objJson As Scripting.Dictionary
    Set objJson = New Scripting.Dictionary
    With objJson
        If BusinessPhones.Count > 0 Then .Add "businessPhone", BusinessPhones
        If DisplayName <> vbNullString Then .Add "displayName", DisplayName
        If GivenName <> vbNullString Then .Add "givenName", GivenName
        If JobTitle <> vbNullString Then .Add "jobTitle", JobTitle
        If Mail <> vbNullString Then .Add "mail", Mail
        If MobilePhone <> vbNullString Then .Add "mobilePhone", MobilePhone
        If OfficeLocation <> vbNullString Then .Add "officeLocation", OfficeLocation
        If PreferredLanguage <> vbNullString Then .Add "preferredLanguage", PreferredLanguage
        If Surname <> vbNullString Then .Add "surname", Surname
        If UserPrincipalName <> vbNullString Then .Add "userPrincipalName", UserPrincipalName
        If Id <> vbNullString Then .Add "id", Id
    End With
    Set ToJson = objJson
End Function

