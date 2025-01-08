VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Company As FirmCompany
Public Name  As String '"Nachname, Vorname"
Public Tel   As String 'Telefonnummer mit Durchwahl
Public Mob   As String 'Mobilnummer
Public Email As String 'Emailadresse

Friend Sub New_(aCompany As FirmCompany, aName As String, aTel As String, aMob As String, aEmail As String)
    Set Me.Company = aCompany
    Me.Name = aName
    Me.Tel = aTel
    Me.Mob = aMob
    Me.Email = aEmail
End Sub

Public Function Key() As String
    If Len(Me.Email) <> 0 Then
        Key = LCase(Me.Email)
    ElseIf Len(Me.Name) <> 0 Then
        Key = LCase(Me.Name)
    Else
        Key = Tel & " " & Mob
    End If
End Function

Public Property Get AbteilungAbbrev() As String
    'entweder von der Abteilung innerhalb der Firma
    'oder von externer Firma die Abkürzung
    'momentan noch keine Abteilung da
    Dim s As String: s = Company.Abbrev
    If Len(s) = 0 Then
        '? was noch?
    End If
    AbteilungAbbrev = s
End Property

Friend Function ToStr() As String
    ToStr = Me.Name & vbTab & Me.Tel & vbTab & Me.Mob & vbTab & Me.Email '& vbCrLf
End Function

