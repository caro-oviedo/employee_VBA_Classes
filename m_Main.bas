Attribute VB_Name = "m_Main"
Option Explicit

Sub EmloyeeDemo()
    Dim cEmployee As New c_Employee
    
    With cEmployee
        .Branch = e_Huston
        .Department = e_Accounting
        .eEmploymentType = e_FullTime
        
        
        .EmployeeNumber = 548934
        
        
        With .Person
            .FirstName = "Tom"
            .LastName = "Jones"
                   
            With .Address
                .City = "Avellaneda"
                .State = "Pineyro"
                .StreetAddress = " 123 Del Pino"
                .StreetAddress_2 = "Apt 2"
                .ZipCode = "1870"
                .WriteAddressToSpreadSheet
            End With
        End With
        .WriteInfoToSpreadSheet
        
    End With
    
End Sub
