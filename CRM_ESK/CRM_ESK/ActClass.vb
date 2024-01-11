Partial Public Class ActClass
    Private EventID_value As String                    '---ID
    Private ActionPlannedDate_value As Date
    Private DirectionName_value As String
    Private EventTypeName_value As String
    Private ScalaCustomerCode_value As String
    Private CompanyName_value As String
    Private ContactName_value As String
    Private ContactPhone_value As String
    Private ContactEMail_value As String
    Private ActionName_value As String
    Private ProjectInfo_value As String
    Private ActionSumm_value As Double
    Private ActionComments_value As String
    Private ActionResultName_value As String
    Private FullName_value As String
    Private CompanyAddress_value As String
    Private CompanyPhone_value As String
    Private CompanyEMail_value As String
    Private IsIKA_value As String
    Private IsApproved_value As Boolean

    Public Property EventID() As String
        Get
            Return EventID_value
        End Get
        Set(ByVal value As String)
            EventID_value = value
        End Set
    End Property

    Public Property ActionPlannedDate() As Date
        Get
            Return ActionPlannedDate_value
        End Get
        Set(ByVal value As Date)
            ActionPlannedDate_value = value
        End Set
    End Property

    Public Property DirectionName() As String
        Get
            Return DirectionName_value
        End Get
        Set(ByVal value As String)
            DirectionName_value = value
        End Set
    End Property

    Public Property EventTypeName() As String
        Get
            Return EventTypeName_value
        End Get
        Set(ByVal value As String)
            EventTypeName_value = value
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

    Public Property ContactName() As String
        Get
            Return ContactName_value
        End Get
        Set(ByVal value As String)
            ContactName_value = value
        End Set
    End Property

    Public Property ContactPhone() As String
        Get
            Return ContactPhone_value
        End Get
        Set(ByVal value As String)
            ContactPhone_value = value
        End Set
    End Property

    Public Property ContactEMail() As String
        Get
            Return ContactEMail_value
        End Get
        Set(ByVal value As String)
            ContactEMail_value = value
        End Set
    End Property

    Public Property ActionName() As String
        Get
            Return ActionName_value
        End Get
        Set(ByVal value As String)
            ActionName_value = value
        End Set
    End Property

    Public Property ProjectInfo() As String
        Get
            Return ProjectInfo_value
        End Get
        Set(ByVal value As String)
            ProjectInfo_value = value
        End Set
    End Property

    Public Property ActionSumm() As Double
        Get
            Return ActionSumm_value
        End Get
        Set(ByVal value As Double)
            ActionSumm_value = value
        End Set
    End Property

    Public Property ActionComments() As String
        Get
            Return ActionComments_value
        End Get
        Set(ByVal value As String)
            ActionComments_value = value
        End Set
    End Property

    Public Property ActionResultName() As String
        Get
            Return ActionResultName_value
        End Get
        Set(ByVal value As String)
            ActionResultName_value = value
        End Set
    End Property

    Public Property FullName() As String
        Get
            Return FullName_value
        End Get
        Set(ByVal value As String)
            FullName_value = value
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

    Public Property IsIKA() As String
        Get
            Return IsIKA_value
        End Get
        Set(ByVal value As String)
            IsIKA_value = value
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

    Public Function GetItem(ByVal i As Integer) As Object
        Select Case i
            Case 1
                Return EventID_value
            Case 2
                Return ActionPlannedDate_value
            Case 3
                Return DirectionName_value
            Case 4
                Return EventTypeName_value
            Case 5
                Return ScalaCustomerCode_value
            Case 6
                Return CompanyName_value
            Case 7
                Return ContactName_value
            Case 8
                Return ContactPhone_value
            Case 9
                Return ContactEMail_value
            Case 10
                Return ActionName_value
            Case 11
                Return ProjectInfo_value
            Case 12
                Return ActionSumm_value
            Case 13
                Return ActionComments_value
            Case 14
                Return ActionResultName_value
            Case 15
                Return FullName_value
            Case 16
                Return CompanyAddress_value
            Case 17
                Return CompanyPhone_value
            Case 18
                Return CompanyEMail_value
            Case 19
                Return IsIKA_value
            Case 20
                Return IsApproved_value
            Case Else
                Return DBNull.Value
        End Select
    End Function

    Public Function GetItem(ByVal MyName As String) As Object
        Select Case MyName
            Case "EventID"
                Return EventID_value
            Case "ActionPlannedDate"
                Return ActionPlannedDate_value
            Case "DirectionName"
                Return DirectionName_value
            Case "EventTypeName"
                Return EventTypeName_value
            Case "ScalaCustomerCode"
                Return ScalaCustomerCode_value
            Case "CompanyName"
                Return CompanyName_value
            Case "ContactName"
                Return ContactName_value
            Case "ContactPhone"
                Return ContactPhone_value
            Case "ContactEMail"
                Return ContactEMail_value
            Case "ActionName"
                Return ActionName_value
            Case "ProjectInfo"
                Return ProjectInfo_value
            Case "ActionSumm"
                Return ActionSumm_value
            Case "ActionComments"
                Return ActionComments_value
            Case "ActionResultName"
                Return ActionResultName_value
            Case "FullName"
                Return FullName_value
            Case "CompanyAddress"
                Return CompanyAddress_value
            Case "CompanyPhone"
                Return CompanyPhone_value
            Case "CompanyEMail"
                Return CompanyEMail_value
            Case "IsIKA"
                Return IsIKA_value
            Case "IsApproved"
                Return IsApproved_value
            Case Else
                Return DBNull.Value
        End Select

    End Function
End Class
