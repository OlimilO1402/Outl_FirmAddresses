VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmExtCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Klasse repräsentiert eine allgemeine Firma, die ebenso "Teil" aber ein externer Teil der FIRMA ist, und allgemein zu Kunden oder Subunternehmen zählt
'Klasse repräsentiert die Firma eines Kunden oder Subunternehmens der FIRMA. Offices werden hier nicht gebraucht
Implements FirmCompany
Public PFN      As String
Public Name     As String
Public m_Abbrev As String
Public Contacts As List 'Of FirmContact

Friend Sub New_(aName As String) ', aAbbrev As String)
    Me.Name = aName
    'm_Abbrev = aAbbrev
    'If Len(m_Abbrev) = 0 Then
       m_Abbrev = GetAbbrev
    'End If
End Sub

Private Function GetAbbrev() As String
    Dim s As String
    If Len(m_Abbrev) = 0 Then
        If NameContains("Firma1") Then
            s = "Fi1"
        ElseIf NameContains("Firma2") Then
            s = "Fi2"
        ElseIf NameContains("Firma3") Then
            s = "Fi3"
        ElseIf NameContains("Firma4") Then
            s = "Fi4"
        ElseIf NameContains("Firma5") Then
            s = "Fi5"
        '
        '...
        '
        End If
    End If
    GetAbbrev = s
End Function
Public Property Get Abbrev() As String
    Abbrev = m_Abbrev
End Property
Private Function NameContains(s As String) As Boolean
    NameContains = InStr(1, Me.Name, s, vbTextCompare)
End Function

Public Property Get Key() As String
    Key = m_Abbrev
End Property

Private Property Get FirmCompany_Abbrev() As String
    FirmCompany_Abbrev = Me.Abbrev
End Property

Private Property Set FirmCompany_Contacts(ByVal Value As List) 'FirmContacts)
    Set Me.Contacts = Value
End Property
Private Property Get FirmCompany_Contacts() As FirmContacts
    Set FirmCompany_Contacts = Me.Contacts
End Property

Private Property Get FirmCompany_Key() As String
    FirmCompany_Key = Me.Key
End Property

Private Property Let FirmCompany_Name(ByVal RHS As String)
    Me.Name = RHS
End Property
Private Property Get FirmCompany_Name() As String
    FirmCompany_Name = Me.Name
End Property

Private Property Set FirmCompany_Offices(ByVal Value As List) 'RHS As FirmOffices)
    '
End Property
Private Property Get FirmCompany_Offices() As FirmOffices
    Set FirmCompany_Offices = Nothing
End Property
