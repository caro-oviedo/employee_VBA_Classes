VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Person"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pFirstName As String
Private pLastname As String
Private pcAddress As New c_Address


Public Property Get FirstName() As String
    FirstName = pFirstName

End Property

Public Property Let FirstName(ByVal FirstNameArg As String)
    pFirstName = FirstNameArg

End Property

Public Property Get LastName() As String
    LastName = pLastname
End Property

Public Property Let LastName(ByVal LastNameArg As String)
    pLastname = LastNameArg
End Property



Public Property Get Address() As c_Address
    Set Address = pcAddress
End Property

Public Property Set Address(ByVal AddressArg As c_Address)
    Set pcAddress = AddressArg

End Property



Public Property Get FullName() As String
    FullName = pFirstName & " " & pLastname

End Property



