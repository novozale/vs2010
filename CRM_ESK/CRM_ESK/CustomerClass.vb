Partial Public Class CustomerClass
    Private CompanyID_value As String
    Private ScalaCustomerCode_value As String
    Private CompanyName_value As String
    Private CompanyAddress_value As String
    Private CompanyPhone_value As String
    Private CompanyEMail_value As String
    Private CustomerGroup_value As String
    Private EndMarket_value As String
    Private IsIKA_value As String
    Private Potencial_value As Double

    Public Property CompanyID() As String
        Get
            Return CompanyID_value
        End Get
        Set(ByVal value As String)
            CompanyID_value = value
        End Set
    End Property

    Public Property ScalaCustomerCode() As String
        Get
            Return ScalaCustomerCode_value
        End Get
        Set(ByVal value As String)
            ScalaCustomerCode_value = value
        End Set
    End Property

    Public Property CompanyName() As String
        Get
            Return CompanyName_value
        End Get
        Set(ByVal value As String)
            CompanyName_value = value
        End Set
    End Property

    Public Property CompanyAddress() As String
        Get
            Return CompanyAddress_value
        End Get
        Set(ByVal value As String)
            CompanyAddress_value = value
        End Set
    End Property

    Public Property CompanyPhone() As String
        Get
            Return CompanyPhone_value
        End Get
        Set(ByVal value As String)
            CompanyPhone_value = value
        End Set
    End Property

    Public Property CompanyEMail() As String
        Get
            Return CompanyEMail_value
        End Get
        Set(ByVal value As String)
            CompanyEMail_value = value
        End Set
    End Property

    Public Property CustomerGroup() As String
        Get
            Return CustomerGroup_value
        End Get
        Set(ByVal value As String)
            CustomerGroup_value = value
        End Set
    End Property

    Public Property EndMarket() As String
        Get
            Return EndMarket_value
        End Get
        Set(ByVal value As String)
            EndMarket_value = value
        End Set
    End Property

    Public Property IsIKA() As String
        Get
            Return IsIKA_value
        End Get
        Set(ByVal value As String)
            IsIKA_value = value
        End Set
    End Property

    Public Property Potencial() As Double
        Get
            Return Potencial_value
        End Get
        Set(ByVal value As Double)
            Potencial_value = value
        End Set
    End Property

    Public Function GetItem(ByVal i As Integer) As Object
        Select Case i
            Case 1
                Return CompanyID_value
            Case 2
                Return ScalaCustomerCode_value
            Case 3
                Return CompanyName_value
            Case 4
                Return CompanyAddress_value
            Case 5
                Return CompanyPhone_value
            Case 6
                Return CompanyEMail_value
            Case 7
                Return CustomerGroup_value
            Case 8
                Return EndMarket_value
            Case 9
                Return IsIKA_value
            Case 10
                Return Potencial_value
             Case Else
                Return DBNull.Value
        End Select
    End Function

    Public Function GetItem(ByVal MyName As String) As Object
        Select Case MyName
            Case "CompanyID"
                Return CompanyID_value
            Case "ScalaCustomerCode"
                Return ScalaCustomerCode_value
            Case "CompanyName"
                Return CompanyName_value
            Case "CompanyAddress"
                Return CompanyAddress_value
            Case "CompanyPhone"
                Return CompanyPhone_value
            Case "CompanyEMail"
                Return CompanyEMail_value
            Case "CustomerGroup"
                Return CustomerGroup_value
            Case "EndMarket"
                Return EndMarket_value
            Case "IsIKA"
                Return IsIKA_value
            Case "Potencial"
                Return Potencial_value
            Case Else
                Return DBNull.Value
        End Select
    End Function
End Class
