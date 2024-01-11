Partial Public Class ProjectClass
    Private ProjectID_value As String
    Private CompanyID_value As String
    Private ScalaCustomerCode_value As String
    Private CompanyName_value As String
    Private ProjectName_value As String
    Private ProjectSumm_value As Double
    Private ProjectComment_value As String
    Private FirstDate_value As Date
    Private LastDate_value As Date
    Private StartDate_value As Date
    Private CloseDate_value As String
    Private ProposalDate_value As Date
    Private ProjectAddr_value As String
    Private Investor_value As String
    Private Contractor_value As String
    Private ResponciblePerson_value As String
    Private ManufacturersList_value As String
    Private AlterManufacturers_value As Boolean
    Private Competitors_value As String
    Private AdditionalExpencesPerCent_value As Double
    Private IsApproved_value As Boolean
    Private ProjectStage_value As String
    Private ParentProjectID_value As String
    Private ParentProjectName_value As String
    Private InvestProject_value As Boolean

    Public Property ProjectID() As String
        Get
            Return ProjectID_value
        End Get
        Set(ByVal value As String)
            ProjectID_value = value
        End Set
    End Property

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

    Public Property ProjectName() As String
        Get
            Return ProjectName_value
        End Get
        Set(ByVal value As String)
            ProjectName_value = value
        End Set
    End Property

    Public Property ProjectSumm() As Double
        Get
            Return ProjectSumm_value
        End Get
        Set(ByVal value As Double)
            ProjectSumm_value = value
        End Set
    End Property

    Public Property ProjectComment() As String
        Get
            Return ProjectComment_value
        End Get
        Set(ByVal value As String)
            ProjectComment_value = value
        End Set
    End Property

    Public Property FirstDate() As Date
        Get
            Return FirstDate_value
        End Get
        Set(ByVal value As Date)
            FirstDate_value = value
        End Set
    End Property

    Public Property LastDate() As Date
        Get
            Return LastDate_value
        End Get
        Set(ByVal value As Date)
            LastDate_value = value
        End Set
    End Property

    Public Property StartDate() As Date
        Get
            Return StartDate_value
        End Get
        Set(ByVal value As Date)
            StartDate_value = value
        End Set
    End Property

    Public Property CloseDate() As String
        Get
            Return CloseDate_value
        End Get
        Set(ByVal value As String)
            CloseDate_value = value
        End Set
    End Property

    Public Property ProposalDate() As Date
        Get
            Return ProposalDate_value
        End Get
        Set(ByVal value As Date)
            ProposalDate_value = value
        End Set
    End Property

    Public Property ProjectAddr() As String
        Get
            Return ProjectAddr_value
        End Get
        Set(ByVal value As String)
            ProjectAddr_value = value
        End Set
    End Property

    Public Property Investor() As String
        Get
            Return Investor_value
        End Get
        Set(ByVal value As String)
            Investor_value = value
        End Set
    End Property

    Public Property Contractor() As String
        Get
            Return Contractor_value
        End Get
        Set(ByVal value As String)
            Contractor_value = value
        End Set
    End Property

    Public Property ResponciblePerson() As String
        Get
            Return ResponciblePerson_value
        End Get
        Set(ByVal value As String)
            ResponciblePerson_value = value
        End Set
    End Property

    Public Property ManufacturersList() As String
        Get
            Return ManufacturersList_value
        End Get
        Set(ByVal value As String)
            ManufacturersList_value = value
        End Set
    End Property

    Public Property AlterManufacturers() As Boolean
        Get
            Return AlterManufacturers_value
        End Get
        Set(ByVal value As Boolean)
            AlterManufacturers_value = value
        End Set
    End Property

    Public Property Competitors() As String
        Get
            Return Competitors_value
        End Get
        Set(ByVal value As String)
            Competitors_value = value
        End Set
    End Property

    Public Property AdditionalExpencesPerCent() As Double
        Get
            Return AdditionalExpencesPerCent_value
        End Get
        Set(ByVal value As Double)
            AdditionalExpencesPerCent_value = value
        End Set
    End Property

    Public Property IsApproved() As Boolean
        Get
            Return IsApproved_value
        End Get
        Set(ByVal value As Boolean)
            IsApproved_value = value
        End Set
    End Property

    Public Property ProjectStage() As String
        Get
            Return ProjectStage_value
        End Get
        Set(ByVal value As String)
            ProjectStage_value = value
        End Set
    End Property

    Public Property ParentProjectID() As String
        Get
            Return ParentProjectID_value
        End Get
        Set(ByVal value As String)
            ParentProjectID_value = value
        End Set
    End Property

    Public Property ParentProjectName() As String
        Get
            Return ParentProjectName_value
        End Get
        Set(ByVal value As String)
            ParentProjectName_value = value
        End Set
    End Property

    Public Property InvestProject() As Boolean
        Get
            Return InvestProject_value
        End Get
        Set(ByVal value As Boolean)
            InvestProject_value = value
        End Set
    End Property

    Public Function GetItem(ByVal i As Integer) As Object
        Select Case i
            Case 1
                Return ProjectID_value
            Case 2
                Return CompanyID_value
            Case 3
                Return ScalaCustomerCode_value
            Case 4
                Return CompanyName_value
            Case 5
                Return ProjectName_value
            Case 6
                Return ProjectSumm_value
            Case 7
                Return ProjectComment_value
            Case 8
                Return FirstDate_value
            Case 9
                Return LastDate_value
            Case 10
                Return StartDate_value
            Case 11
                Return CloseDate_value
            Case 12
                Return ProposalDate_value
            Case 13
                Return ProjectAddr_value
            Case 14
                Return Investor_value
            Case 15
                Return Contractor_value
            Case 16
                Return ResponciblePerson_value
            Case 17
                Return ManufacturersList_value
            Case 18
                Return AlterManufacturers_value
            Case 19
                Return Competitors_value
            Case 20
                Return AdditionalExpencesPerCent_value
            Case 21
                Return IsApproved_value
            Case 22
                Return ProjectStage_value
            Case 23
                Return ParentProjectID_value
            Case 24
                Return ParentProjectName_value
            Case 25
                Return InvestProject_value
            Case Else
                Return DBNull.Value
        End Select
    End Function

    Public Function GetItem(ByVal MyName As String) As Object
        Select Case MyName
            Case "ProjectID"
                Return ProjectID_value
            Case "CompanyID"
                Return CompanyID_value
            Case "ScalaCustomerCode"
                Return ScalaCustomerCode_value
            Case "CompanyName"
                Return CompanyName_value
            Case "ProjectName"
                Return ProjectName_value
            Case "ProjectSumm"
                Return ProjectSumm_value
            Case "ProjectComment"
                Return ProjectComment_value
            Case "FirstDate"
                Return FirstDate_value
            Case "LastDate"
                Return LastDate_value
            Case "StartDate"
                Return StartDate_value
            Case "CloseDate"
                Return CloseDate_value
            Case "ProposalDate"
                Return ProposalDate_value
            Case "ProjectAddr"
                Return ProjectAddr_value
            Case "Investor"
                Return Investor_value
            Case "Contractor"
                Return Contractor_value
            Case "ResponciblePerson"
                Return ResponciblePerson_value
            Case "ManufacturersList"
                Return ManufacturersList_value
            Case "AlterManufacturers"
                Return AlterManufacturers_value
            Case "Competitors"
                Return Competitors_value
            Case "AdditionalExpencesPerCent"
                Return AdditionalExpencesPerCent_value
            Case "IsApproved"
                Return IsApproved_value
            Case "ProjectStage"
                Return ProjectStage_value
            Case "ParentProjectID"
                Return ParentProjectID_value
            Case "ParentProjectName"
                Return ParentProjectName_value
            Case "InvestProject"
                Return InvestProject_value
            Case Else
                Return DBNull.Value
        End Select
    End Function
End Class
