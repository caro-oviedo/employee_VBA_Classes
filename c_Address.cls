VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Address"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pStreetAddress As String
Private pStreetAddress_2 As String
Private pCity As String
Private pState As String
Private pZipCode As String



Public Property Get StreetAddress() As String
    StreetAddress = pStreetAddress

End Property

Public Property Let StreetAddress(ByVal StreetAddressArg As String)
    pStreetAddress = StreetAddressArg
    

End Property




Public Property Get StreetAddress_2() As String
    StreetAddress_2 = pStreetAddress_2

End Property

Public Property Let StreetAddress_2(ByVal StreetAddress_2Arg As String)
    pStreetAddress_2 = StreetAddress_2Arg

End Property




Public Property Get City() As String
    City = pCity

End Property

Public Property Let City(ByVal CityArg As String)
    pCity = CityArg

End Property



Public Property Get State() As String
    State = pState

End Property

Public Property Let State(ByVal StateArg As String)
    pState = StateArg

End Property




Public Property Get ZipCode() As String
    ZipCode = pZipCode

End Property

Public Property Let ZipCode(ByVal ZipCodeArg As String)
    pZipCode = ZipCodeArg

End Property



Public Sub WriteAddressToSpreadSheet()
    Sheet1.Range("G5").Value = pStreetAddress
    Sheet1.Range("H5").Value = pStreetAddress_2
    Sheet1.Range("I5").Value = pCity
    Sheet1.Range("J5").Value = pState
    Sheet1.Range("K5").Value = pZipCode
End Sub






