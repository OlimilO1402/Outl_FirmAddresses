VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmContacts_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_List As Collection '(Of FirmContact)

Private Sub Class_Initialize()
    Set m_List = New Collection
End Sub

Public Function Add(aNewContact As FirmContact) As FirmContact
    Set Add = aNewContact
    Dim Key As String: Key = Add.Key
    If ContainsKey(Key) Then
        Set Add = m_List.Item(Key)
    Else
        m_List.Add Add, Key
    End If
End Function

Public Function ContainsKey(Key As String) As Boolean
Try: On Error Resume Next
    If IsEmpty(m_List(Key)) Then:  'DoNothing
    ContainsKey = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Property Get Item(Key As String) As FirmContact
    Set Item = m_List.Item(Key)
End Property

Public Property Get Contains(aContact As FirmContact) As Boolean
    Contains = ContainsKey(aContact.Key)
End Property

Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get List() As Collection
    Set List = m_List
End Property

Public Function ToStr() As String
    Dim s As String
    Dim ctc As FirmContact
    For Each ctc In m_List
        s = s & ctc.ToStr & vbCrLf
    Next
    ToStr = s
End Function

