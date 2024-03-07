VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pcPerson As New c_Person
Private pEmploymentType As enumEmployeeType
Private pEmployeeNumber As String
Private pBranch As enumBranch
Private pDepartment As enumDepartment



Public Property Get Person() As c_Person
    Set Person = pcPerson
End Property

Public Property Set Person(ByVal PersonArg As c_Person)
    Set pcPerson = PersonArg

End Property

Public Property Get eEmploymentType() As enumEmployeeType
    eEmploymentType = pEmploymentType

End Property

Public Property Let eEmploymentType(ByVal eEmploymentTypeArg As enumEmployeeType)
    pEmploymentType = eEmploymentTypeArg

End Property

Public Property Get EmployeeNumber() As String
   EmployeeNumber = pEmployeeNumber
End Property

Public Property Let EmployeeNumber(ByVal EmployeeNumberArg As String)
    pEmployeeNumber = EmployeeNumberArg

End Property


Public Property Get Department() As enumDepartment
    Department = pDepartment

End Property

Public Property Let Department(ByVal DepartmentArg As enumDepartment)
    pDepartment = DepartmentArg

End Property


Public Property Get Branch() As enumBranch
    Branch = pBranch

End Property

Public Property Let Branch(ByVal BranchArg As enumBranch)
    pBranch = BranchArg

End Property


Public Sub WriteInfoToSpreadSheet()
    Sheet1.Range("A5").Value = pEmployeeNumber
    Sheet1.Range("B5").Value = pEmploymentType
    Sheet1.Range("C5").Value = pBranch
    Sheet1.Range("D5").Value = pDepartment
    Sheet1.Range("E5").Value = pcPerson.FirstName
    Sheet1.Range("F5").Value = pcPerson.LastName
    
End Sub




