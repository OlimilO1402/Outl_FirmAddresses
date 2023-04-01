VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirmExtCompanys_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_List As Collection '(Of FirmExtCompany)

Private Sub Class_Initialize()
    Set m_List = New Collection
End Sub

Public Function Add(aNewCompany As FirmExtCompany) As FirmExtCompany
    Set Add = aNewCompany
    Dim Key As String: Key = Add.Key
    If ContainsKey(Key) Then
        Set Add = m_List.Item(Key)
    Else
        m_List.Add Add, Key
    End If
End Function

Private Function ContainsKey(Key As String) As Boolean
Try: On Error Resume Next
    If IsEmpty(m_List(Key)) Then:  'DoNothing
    ContainsKey = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Property Get Item(Key As String) As FirmExtCompany
    Set Item = m_List.Item(Key)
End Property

Public Property Get Contains(aCompany As FirmExtCompany) As Boolean
    Contains = ContainsKey(aCompany.Key)
End Property

Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get List() As Collection
    Set List = m_List
End Property

Public Function ToFiles(Prepfn As String)
    Dim pfn As String
    Dim cmp As FirmExtCompany
    For Each cmp In m_List
        pfn = Prepfn & cmp.Key & ".txt"
        cmp.ToFile pfn
    Next
End Function
