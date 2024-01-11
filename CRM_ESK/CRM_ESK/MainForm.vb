'Option Strict Off
Imports Syncfusion.WinForms.DataGridConverter
Imports Syncfusion.WinForms.DataGrid.Events
Imports Syncfusion.WinForms.DataGrid.Enums
Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.Controls
Imports Syncfusion.Windows.Forms
Imports System
Imports Syncfusion.Data
Imports System.Drawing
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports Syncfusion.WinForms.DataGrid.Interactivity
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Security.Permissions
Imports System.Collections.Generic
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Globalization
Imports Syncfusion.WinForms.GridCommon.ScrollAxis


Public Class MainForm
    Private propertyAccessProvider As IPropertyAccessProvider = Nothing

    Private Tab0Settings As Integer = 0
    Private Tab1Settings As Integer = 0
    Private Tab2Settings As Integer = 0
    Private Tab3Settings As Integer = 0

    Private autoFitOptions As New RowAutoFitOptions()
    Private autoHeight As Integer
    Private FilterText As String = ""
    Private ShowOnlyActive As Integer = 0           '---Показывать все записи или только план / факт в sfdatagrid2
    Private ShowSfDataGri2Plan As Integer = 1       '---Показывать планы в sfdatagrid2
    Private ShowSfDataGri2Fact As Integer = 0       '---Показывать факты в sfdatagrid2
    Private ShowSfDataGri2Details As Integer = 0    '---Показывать детали в форме планов
    Private ShowSfDataGri4Details As Integer = 0    '---Показывать детали в форме проектов

    Private LoadFlag = 0                            '---флаг - идет загрузка формы или нет
    Private SfDataGrid2FilterBlockFlag = 0          '---флаг блокировки фильтра в колонке план - факт

    Public Sub New()
        InitializeComponent()
        'Font = New Font(New FontFamily("Arial Unicode MS"), 8.0F, FontStyle.Bold)
        'Font = New Font(New FontFamily("Microsoft Sans Serif"), 8.0F, FontStyle.Bold)
    End Sub

    Private Sub MainForm_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '// 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub MainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске определяем параметры - Год, компания, пользователь и т.д.
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        AddHandler SfDataGrid1.QueryRowHeight, AddressOf sfDataGrid1_QueryRowHeight
        AddHandler SfDataGrid1.QueryRowStyle, AddressOf SfDataGrid1_QueryRowStyle
        AddHandler SfDataGrid1.SelectionChanged, AddressOf SfDataGrid1_SelectionChanged

        AddHandler SfDataGrid2.QueryRowHeight, AddressOf sfDataGrid2_QueryRowHeight
        AddHandler SfDataGrid2.QueryCellStyle, AddressOf SfDataGrid2_QueryCellStyle
        AddHandler SfDataGrid2.ToolTipOpening, AddressOf SfDataGrid2_ToolTipOpening
        AddHandler SfDataGrid2.SelectionChanged, AddressOf SfDataGrid2_SelectionChanged

        AddHandler SfDataGrid4.QueryRowHeight, AddressOf sfDataGrid4_QueryRowHeight
        AddHandler SfDataGrid4.QueryRowStyle, AddressOf SfDataGrid4_QueryRowStyle
        AddHandler SfDataGrid4.SelectionChanged, AddressOf SfDataGrid4_SelectionChanged

        AddHandler SfDataGrid7.QueryRowHeight, AddressOf sfDataGrid7_QueryRowHeight
        AddHandler SfDataGrid7.QueryRowStyle, AddressOf SfDataGrid7_QueryRowStyle
        AddHandler SfDataGrid7.SelectionChanged, AddressOf SfDataGrid7_SelectionChanged

        '---параметры запуска
        Try
            Dim Scala As New SfwIII.Application

            Declarations.CompanyID = Scala.ActiveProcess.CommonVars.CompanyCode
            Declarations.Year = Mid(Scala.ActiveProcess.CommonVars.FiscalYear, 3)
            Declarations.UserCode = Scala.ActiveProcess.CommonVars.UserCode
            'Declarations.UserCode = "novozhilov"

        Catch
            MsgBox("Программа должна запускаться только из меню Scala", MsgBoxStyle.Critical, "Внимание!")
            Application.Exit()
        End Try

        '---ID пользователя
        MySQLStr = "SELECT UserID, FullName "
        MySQLStr = MySQLStr & "FROM  ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Upper(UserName) = N'" & UCase(Trim(Declarations.UserCode)) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Не найден ID сотрудника, соответствующий логину на вход в Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            Application.Exit()
        Else
            Declarations.UserID = Declarations.MyRec.Fields("UserID").Value
            Declarations.FullName = Declarations.MyRec.Fields("FullName").Value
            trycloseMyRec()
        End If

        '---продавец
        MySQLStr = "SELECT ST010300.ST01001 AS SC, ST010300.ST01002 AS FullName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 "
        MySQLStr = MySQLStr & "WHERE (UPPER(ScalaSystemDB.dbo.ScaUsers.UserName) = UPPER(N'" & Declarations.UserCode & "')) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Не найден код продавца, соответствующий логину на вход в Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            Application.Exit()
        Else
            Declarations.SalesmanCode = Declarations.MyRec.Fields("SC").Value
            trycloseMyRec()
        End If

        '---кост центр пользователя
        MySQLStr = "SELECT SUBSTRING(ST01021, 7, 3) AS CC "
        MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ST01002 = N'" & Declarations.FullName & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            MsgBox("Не найден Кост центр сотрудника, соответствующий логину на вход в Scala. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            Application.Exit()
        Else
            Declarations.CC = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If

        '---Является ли пользователь членом группы CRMManaregs или CRMDirector или ProjectDirector
        CheckRights(Declarations.UserCode, "CRMManagers")
        CheckRights1(Declarations.UserCode, "CRMDirector")
        CheckRights2(Declarations.UserCode, "ProjectDirector")

        '---Разрешено ли пользователям из групп CRMManagers и CRMDirector менять 
        '---пользователя - владельца события при создании (редактировании) 1-можно 0-нельзя
        MySQLStr = "SELECT Value "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Config WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (Parameter = N'AllowChangeUser') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            Declarations.AllowChangeUser = "0"
        Else
            Declarations.AllowChangeUser = Declarations.MyRec.Fields("Value").Value
            trycloseMyRec()
        End If

        TabControl1.SelectedIndex = 0
        LoadDataToScheduleList()
    End Sub

    Private Sub TabControl1_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор закладки
        '// для выбранной закладки выводим данные
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        Select Case sender.selectedtab.text
            Case "План на месяц"
                LoadDataToScheduleList()
            Case "Список действий"
                LoadDataToTabActList(Nothing, Nothing, Nothing, Nothing, Nothing)
            Case "Проекты"
                LoadDataToProjects(Nothing, Nothing)
            Case "Клиенты"
                LoadDataToCustomers()
            Case Else
        End Select
    End Sub

    Private Sub LoadDataToCustomers()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных на закладку "Клиенты"
        '//
        '/////////////////////////////////////////////////////////////////////////////////////

        If Tab3Settings = 0 Then
            LoadFlag = 1

            LoadCustomersList()
            TextBox4.Text = ""
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()

            '-----параметры таблицы
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True

            Tab3Settings = 1
            LoadFlag = 0
        End If
    End Sub

    Private Sub LoadDataToProjects(ByVal MyStartData As DateTime, ByVal MyFinData As DateTime)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных на закладку "Проекты"
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCtrlDate As DateTime

        If Tab2Settings = 0 Then
            LoadFlag = 1

            '---Фильтр по датам
            Try
                If MyStartData = MyCtrlDate Then
                    DateTimePicker4.Value = DateAdd(DateInterval.Year, -1, CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())))
                Else
                    DateTimePicker4.Value = MyStartData
                End If
                If MyFinData = MyCtrlDate Then
                    DateTimePicker3.Value = DateAdd(DateInterval.Year, 1, CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())))
                Else
                    DateTimePicker3.Value = MyFinData
                End If
            Catch ex As Exception
            End Try

            SplitContainer2.Panel2Collapsed = True
            LoadProjectsList()
            TextBox3.Text = ""
            Button23.Text = "Показывать детали"
            ShowSfDataGri4Details = 0
            CheckProjectButtons()
            CheckProjectApprovButton()

            '-----параметры таблицы
            SetInitProjectTableParams(SfDataGrid4)
            SetProjectTableParams(SfDataGrid4)
            SfDataGrid4.Visible = True

            Tab2Settings = 1
            LoadFlag = 0
        End If
    End Sub

    Private Sub LoadDataToScheduleList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных на закладку "План на месяц"
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка продавцов
        Dim MyDs As New DataSet                       '

        If Tab0Settings = 0 Then
            LoadFlag = 1
            '-----год
            Dim comboSource1 As New Dictionary(Of Integer, String)
            For i As Integer = 2012 To (Now().Year + 1)
                comboSource1.Add(i, i.ToString)
            Next i
            ComboBox4.DataSource = New BindingSource(comboSource1, Nothing)
            ComboBox4.DisplayMember = "Value"
            ComboBox4.ValueMember = "Key"
            ComboBox4.SelectedValue = Now().Year
            'ComboBox4.SelectedValue = 2018

            '-----месяц
            Dim comboSource As New Dictionary(Of Integer, String)
            comboSource.Add(1, "Январь")
            comboSource.Add(2, "Февраль")
            comboSource.Add(3, "Март")
            comboSource.Add(4, "Апрель")
            comboSource.Add(5, "Май")
            comboSource.Add(6, "Июнь")
            comboSource.Add(7, "Июль")
            comboSource.Add(8, "Август")
            comboSource.Add(9, "Сентябрь")
            comboSource.Add(10, "Октябрь")
            comboSource.Add(11, "Ноябрь")
            comboSource.Add(12, "Декабрь")
            ComboBox5.DataSource = New BindingSource(comboSource, Nothing)
            ComboBox5.DisplayMember = "Value"
            ComboBox5.ValueMember = "Key"
            ComboBox5.SelectedValue = Now().Month

            '---Список продавцов
            If Declarations.MyPermission = True Then
                '---Доступны все продавцы
                MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
                MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
                MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
                MySQLStr = MySQLStr & "WHERE(ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "
                MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
            ElseIf Declarations.MyCCPermission = True Then
                '---Доступны продавцы определенного кост центра
                MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
                MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
                MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
                MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
                MySQLStr = MySQLStr & "(tbl_CRM_CCOwners.CCOwn = N'" & Declarations.CC & "') "
                MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
            Else
                '---только один продавец (вошедший в систему)
                MySQLStr = "SELECT ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
                MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
                MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
                MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
                MySQLStr = MySQLStr & "(ScalaSystemDB.dbo.ScaUsers.UserID = " & Declarations.UserID & ") "
                MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
            End If
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                ComboBox6.DisplayMember = "FullName" 'Это то что будет отображаться
                ComboBox6.ValueMember = "UserID"   'это то что будет храниться
                ComboBox6.DataSource = MyDs.Tables(0).DefaultView
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            ComboBox6.SelectedValue = Declarations.UserID
            'ComboBox6.SelectedValue = 394
            LoadFlag = 0
            SplitContainer1.Panel2Collapsed = True

            LoadScheduleList()

            '-----параметры таблицы
            SetInitScheduleTableParams(SfDataGrid2)
            SetScheduleTableParams(SfDataGrid2)
            CheckSfDataGrid2Activity()
            CheckPlanFactState()
            LoadScheduleListDetail()
            SfDataGrid2.Visible = True
            CheckPlanesButtonSf2()

            Tab0Settings = 1
        End If
    End Sub

    Private Sub LoadCustomersList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка клиентов sfdatagrid7
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        '---------------------Данные по клиентам
        Try
            MySQLStr = "SELECT tbl_CRM_Companies.CompanyID, ISNULL(tbl_CRM_Companies.ScalaCustomerCode, '') AS ScalaCustomerCode, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies.CompanyName, '') AS CompanyName, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies.CompanyAddress, '') AS CompanyAddress, ISNULL(tbl_CRM_Companies.CompanyPhone, '') AS CompanyPhone "
            MySQLStr = MySQLStr & ", ISNULL(tbl_CRM_Companies.CompanyEMail, '') AS CompanyEMail,"
            MySQLStr = MySQLStr & "ISNULL(tbl_RexelCustomerGroup.RussianName,'') AS CustomerGroup, ISNULL(tbl_RexelEndMarkets.RussianName, '') AS EndMarket, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies_Ext.IsIKA, N'') AS IsIKA, ISNULL(tbl_CRM_Companies_Ext.Potencial, 0) AS Potencial "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Companies WITH(NOLOCK) LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_RexelCustomerGroup ON tbl_CRM_Companies.RCGCode = tbl_RexelCustomerGroup.RCGCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_RexelEndMarkets ON tbl_CRM_Companies.EMCode = tbl_RexelEndMarkets.EMCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & " tbl_CRM_Companies_Ext ON tbl_CRM_Companies.CompanyID = tbl_CRM_Companies_Ext.CompanyID "
            MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName "
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs, "MasterData")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            '-----
            Dim MyStr As CustomerClass
            Dim MyList As List(Of CustomerClass)
            MyList = New List(Of CustomerClass)
            For i As Integer = 0 To MyDs.Tables(0).Rows.Count - 1
                MyStr = New CustomerClass
                MyStr.CompanyID = MyDs.Tables(0).Rows(i).Item("CompanyID").ToString()
                MyStr.ScalaCustomerCode = MyDs.Tables(0).Rows(i).Item("ScalaCustomerCode")
                MyStr.CompanyName = MyDs.Tables(0).Rows(i).Item("CompanyName")
                MyStr.CompanyAddress = MyDs.Tables(0).Rows(i).Item("CompanyAddress")
                MyStr.CompanyPhone = MyDs.Tables(0).Rows(i).Item("CompanyPhone")
                MyStr.CompanyEMail = MyDs.Tables(0).Rows(i).Item("CompanyEMail")
                MyStr.CustomerGroup = MyDs.Tables(0).Rows(i).Item("CustomerGroup")
                MyStr.EndMarket = MyDs.Tables(0).Rows(i).Item("EndMarket")
                MyStr.IsIKA = MyDs.Tables(0).Rows(i).Item("IsIKA")
                MyStr.Potencial = MyDs.Tables(0).Rows(i).Item("Potencial")

                MyList.Add(MyStr)
            Next
            '-----
            SfDataGrid7.Visible = False
            SfDataGrid7.DataSource = MyList
        Catch ex As Exception
        End Try
    End Sub

    Private Sub LoadProjectsList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка проектов sfdatagrid4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        '---------------------Данные по проектам
        Try
            MySQLStr = "exec spp_CRM_GetProjectInfo N'" _
                + Format(DateTimePicker4.Value, "dd/MM/yyyy") + "', N'" _
                + Format(DateTimePicker3.Value, "dd/MM/yyyy") + "' "
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs, "MasterData")
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            '-----
            Dim MyStr As ProjectClass
            Dim MyList As List(Of ProjectClass)
            MyList = New List(Of ProjectClass)
            For i As Integer = 0 To MyDs.Tables(0).Rows.Count - 1
                MyStr = New ProjectClass
                MyStr.ProjectID = MyDs.Tables(0).Rows(i).Item("ProjectID").ToString()
                MyStr.CompanyID = MyDs.Tables(0).Rows(i).Item("CompanyID").ToString()
                MyStr.ScalaCustomerCode = MyDs.Tables(0).Rows(i).Item("ScalaCustomerCode")
                MyStr.CompanyName = MyDs.Tables(0).Rows(i).Item("CompanyName")
                MyStr.ProjectName = MyDs.Tables(0).Rows(i).Item("ProjectName")
                MyStr.ProjectSumm = MyDs.Tables(0).Rows(i).Item("ProjectSumm")
                MyStr.ProjectComment = MyDs.Tables(0).Rows(i).Item("ProjectComment")
                MyStr.FirstDate = MyDs.Tables(0).Rows(i).Item("FirstDate")
                MyStr.LastDate = MyDs.Tables(0).Rows(i).Item("LastDate")
                MyStr.StartDate = MyDs.Tables(0).Rows(i).Item("StartDate")
                If Not IsDBNull(MyDs.Tables(0).Rows(i).Item("CloseDate")) Then
                    MyStr.CloseDate = Format(MyDs.Tables(0).Rows(i).Item("CloseDate"), "dd/MM/yyyy")
                Else
                    MyStr.CloseDate = ""
                End If
                MyStr.ProposalDate = MyDs.Tables(0).Rows(i).Item("ProposalDate")
                MyStr.ProjectAddr = MyDs.Tables(0).Rows(i).Item("ProjectAddr")
                MyStr.Investor = MyDs.Tables(0).Rows(i).Item("Investor")
                MyStr.Contractor = MyDs.Tables(0).Rows(i).Item("Contractor")
                MyStr.ResponciblePerson = MyDs.Tables(0).Rows(i).Item("ResponciblePerson")
                MyStr.ManufacturersList = MyDs.Tables(0).Rows(i).Item("ManufacturersList")
                MyStr.AlterManufacturers = MyDs.Tables(0).Rows(i).Item("AlterManufacturers")
                MyStr.Competitors = MyDs.Tables(0).Rows(i).Item("Competitors")
                MyStr.AdditionalExpencesPerCent = MyDs.Tables(0).Rows(i).Item("AdditionalExpencesPerCent")
                MyStr.IsApproved = MyDs.Tables(0).Rows(i).Item("IsApproved")
                MyStr.ProjectStage = MyDs.Tables(0).Rows(i).Item("ProjectStage")
                MyStr.ParentProjectID = MyDs.Tables(0).Rows(i).Item("ParentProjectID")
                MyStr.ParentProjectName = MyDs.Tables(0).Rows(i).Item("ParentProjectName")
                MyStr.InvestProject = MyDs.Tables(0).Rows(i).Item("InvestProject")

                MyList.Add(MyStr)
            Next
            '-----
            'SfDataGrid1.DataSource = MyDs.Tables(0)
            SfDataGrid4.Visible = False
            SfDataGrid4.DataSource = MyList
        Catch ex As Exception
        End Try

    End Sub

    Public Sub LoadScheduleList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка действий sfdatagrid2
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySqlStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        If LoadFlag = 0 Then
            '---------------------Данные по действиям
            Try
                MySqlStr = "exec spp_CRM_GetMonthSchedule " + ComboBox4.SelectedValue.ToString + ", " + ComboBox5.SelectedValue.ToString + ", " + ComboBox6.SelectedValue.ToString
                Try
                    MyAdapter = New SqlClient.SqlDataAdapter(MySqlStr, Declarations.MyNETConnStr)
                    MyAdapter.SelectCommand.CommandTimeout = 600
                    MyAdapter.Fill(MyDs, "MasterData")
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Catch ex As Exception
            End Try

            '======Динамические данные
            '-----формирование класса работы с данными
            Dim MyDict = New Dictionary(Of String, Type)
            For i As Integer = 0 To MyDs.Tables(0).Columns.Count - 1
                MyDict.Add(MyDs.Tables(0).Columns(i).ToString, MyDs.Tables(0).Columns(i).DataType)
            Next
            Dim MyShClassType As Type = CreateMonthScheduleClassType("MonthScheduleClass", MyDict)
            '-----окончание формирования класса работы с данными

            '-----заполнение списка
            Dim Mylist As IList = CType(Activator.CreateInstance(GetType(List(Of )).MakeGenericType(New Type() {MyShClassType})), IList)
            For i As Integer = 0 To MyDs.Tables(0).Rows.Count - 1
                Dim MyShClass As Object = Activator.CreateInstance(MyShClassType)
                For j As Integer = 0 To MyDs.Tables(0).Columns.Count - 1
                    If Not IsDBNull(MyDs.Tables(0).Rows(i).Item(j)) Then
                        MyShClass.SetItem(MyDs.Tables(0).Columns(j).ToString, MyDs.Tables(0).Rows(i).Item(j))
                    End If
                Next
                Mylist.Add(MyShClass)
            Next
            '-----окончание заполнения списка
            SfDataGrid2.Visible = False
            SfDataGrid2.DataSource = Mylist
            '=======конец динамических данных

        End If
    End Sub

    Private Sub LoadProjectListDetail()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка деталей проектов sfdatagrid4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySqlStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet
        Dim MyAdapter1 As SqlClient.SqlDataAdapter
        Dim MyDs1 As New DataSet
        Dim MyProjectID As String

        If SfDataGrid4.SelectedItem Is Nothing Then
            SfDataGrid5.DataSource = Nothing
            SfDataGrid6.DataSource = Nothing
        Else
            MyProjectID = SfDataGrid4.SelectedItem.GetItem("ProjectID").ToString
            '-------------------Данные по активностям
            Try
                MySqlStr = "exec spp_CRM_GetProjectActivityDetails N'" + MyProjectID + "'"
                Try
                    MyAdapter = New SqlClient.SqlDataAdapter(MySqlStr, Declarations.MyNETConnStr)
                    MyAdapter.SelectCommand.CommandTimeout = 600
                    MyAdapter.Fill(MyDs)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Catch ex As Exception
            End Try
            SfDataGrid5.Visible = False
            SfDataGrid5.DataSource = MyDs.Tables(0)
            SetSfDataGrid5InitTableParams(SfDataGrid5)
            SetSfDataGrid5TableParams(SfDataGrid5)
            SfDataGrid5.Visible = True
            '-------------------Данные по заказам
            Try
                MySqlStr = "exec spp_CRM_GetProjectOrdersDetails N'" + MyProjectID + "'"
                Try
                    MyAdapter1 = New SqlClient.SqlDataAdapter(MySqlStr, Declarations.MyNETConnStr)
                    MyAdapter1.SelectCommand.CommandTimeout = 600
                    MyAdapter1.Fill(MyDs1)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Catch ex As Exception
            End Try
            SfDataGrid6.Visible = False
            SfDataGrid6.DataSource = MyDs1.Tables(0)
            SetSfDataGrid6InitTableParams(SfDataGrid6)
            SetSfDataGrid6TableParams(SfDataGrid6)
            SfDataGrid6.Visible = True
        End If
    End Sub

    Private Sub LoadScheduleListDetail()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка деталей действий sfdatagrid3
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySqlStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet
        Dim MyCompanyID As String

        If SfDataGrid2.SelectedItem Is Nothing Then
            SfDataGrid3.DataSource = Nothing
        Else
            MyCompanyID = SfDataGrid2.SelectedItem.GetItem("CompanyID").ToString
            '-------------------Данные по деталям
            Try
                MySqlStr = "exec spp_CRM_GetMonthScheduleDetails " + ComboBox4.SelectedValue.ToString + ", " + ComboBox5.SelectedValue.ToString + ", N'" + MyCompanyID + "'"
                Try
                    MyAdapter = New SqlClient.SqlDataAdapter(MySqlStr, Declarations.MyNETConnStr)
                    MyAdapter.SelectCommand.CommandTimeout = 600
                    MyAdapter.Fill(MyDs)
                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            Catch ex As Exception
            End Try
            SfDataGrid3.Visible = False
            SfDataGrid3.DataSource = MyDs.Tables(0)
            SetSfDataGrid3InitTableParams(SfDataGrid3)
            SetSfDataGrid3TableParams(SfDataGrid3)
            SfDataGrid3.Visible = True
        End If

    End Sub

    Public Function ReqTable(Of T)(ByVal MySQLStr As String) As List(Of T)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование списка типа Т из таблицы
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim result As New List(Of T)()

        InitMyConn(False)
        Dim request = Declarations.MyConn.Execute(MySQLStr)

        While Not request.EOF
            result.Add( _
                CType(Activator.CreateInstance( _
                        GetType(T), New Object() {request.Fields}),  _
                    T))
            request.MoveNext()
        End While

        Return result
    End Function

    Public Shared Function CreateClass(ByVal className As String, ByVal properties As Dictionary(Of String, Type)) As Type
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Динамическое создание класса для работы с планами
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim myDomain As AppDomain = AppDomain.CurrentDomain
        Dim myAsmName As New AssemblyName("EskRuAssembly")
        Dim myAssembly As AssemblyBuilder = myDomain.DefineDynamicAssembly(myAsmName, AssemblyBuilderAccess.Run)
        Dim myModule As ModuleBuilder = myAssembly.DefineDynamicModule("EskRuModule")
        Dim myType As TypeBuilder = myModule.DefineType(className, TypeAttributes.Public And TypeAttributes.Class)
        'Dim getSetAttr As MethodAttributes = MethodAttributes.Public And MethodAttributes.SpecialName And MethodAttributes.HideBySig
        Dim getSetAttr As MethodAttributes = MethodAttributes.Public
        myType.DefineDefaultConstructor(MethodAttributes.Public)

        For Each o In properties
            Dim prop As PropertyBuilder = myType.DefineProperty(o.Key, PropertyAttributes.HasDefault, o.Value, Type.EmptyTypes)
            Dim field As FieldBuilder = myType.DefineField("_" + o.Key, o.Value, FieldAttributes.Private)

            Dim getter As MethodBuilder = myType.DefineMethod("get_" + o.Key, getSetAttr, o.Value, Type.EmptyTypes)
            Dim getterIL As ILGenerator = getter.GetILGenerator()
            getterIL.Emit(OpCodes.Ldarg_0)
            getterIL.Emit(OpCodes.Ldfld, field)
            getterIL.Emit(OpCodes.Ret)
            prop.SetGetMethod(getter)

            Dim setter As MethodBuilder = myType.DefineMethod("set_" + o.Key, getSetAttr, Nothing, New Type() {o.Value})
            Dim setterIL As ILGenerator = setter.GetILGenerator()
            setterIL.Emit(OpCodes.Ldarg_0)
            setterIL.Emit(OpCodes.Ldarg_1)
            setterIL.Emit(OpCodes.Stfld, field)
            setterIL.Emit(OpCodes.Ret)
            prop.SetSetMethod(setter)
        Next


        Return myType.CreateType()
    End Function

    Private Sub LoadDataToTabActList(ByVal MyStartData As DateTime, ByVal MyFinData As DateTime, ByVal MySalesmanID As Integer, ByVal MyActions As String, ByVal MyActivity As String)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных на закладку "Список действий"
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка продавцов
        Dim MyDs As New DataSet                       '
        Dim MyCtrlDate As DateTime

        If Tab1Settings = 0 Then
            LoadFlag = 1
            '---Фильтр по датам
            Try
                If MyStartData = MyCtrlDate Then
                    DateTimePicker1.Value = DateAdd(DateInterval.Year, -1, CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())))
                Else
                    DateTimePicker1.Value = MyStartData
                End If
                If MyFinData = MyCtrlDate Then
                    DateTimePicker2.Value = DateAdd(DateInterval.Year, 1, CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())))
                Else
                    DateTimePicker2.Value = MyFinData
                End If
            Catch ex As Exception
            End Try

            '---Список продавцов
            If Declarations.MyPermission = True Then
                '---Доступны все продавцы
                MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
                MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
                MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
                MySQLStr = MySQLStr & "WHERE(ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "
                MySQLStr = MySQLStr & "Union ALL "
                MySQLStr = MySQLStr & "SELECT 0 AS UserID, ' Все пользователи' as FullName "
                MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
            ElseIf Declarations.MyCCPermission = True Then
                '---Доступны продавцы определенного кост центра
                MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
                MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
                MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
                MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
                MySQLStr = MySQLStr & "(tbl_CRM_CCOwners.CCOwn = N'" & Declarations.CC & "') "
                MySQLStr = MySQLStr & "Union ALL "
                MySQLStr = MySQLStr & "SELECT 0 AS UserID, ' Все пользователи' as FullName "
                MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
            Else
                '---только один продавец (вошедший в систему)
                MySQLStr = "SELECT ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
                MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
                MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
                MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
                MySQLStr = MySQLStr & "(ScalaSystemDB.dbo.ScaUsers.UserID = " & Declarations.UserID & ") "
                MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
            End If
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                ComboBox1.DisplayMember = "FullName" 'Это то что будет отображаться
                ComboBox1.ValueMember = "UserID"   'это то что будет храниться
                ComboBox1.DataSource = MyDs.Tables(0).DefaultView
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            '---Выбранный продавец
            If IsNothing(MySalesmanID) Or MySalesmanID = 0 Then
                ComboBox1.SelectedValue = Declarations.UserID
            Else
                ComboBox1.SelectedValue = MySalesmanID
            End If

            '---Действия по продавцу или все действия по клиентам
            If IsNothing(MyActions) Then
                ComboBox2.Text = "Только выбранного продавца"
            Else
                ComboBox2.Text = MyActions
            End If

            '---Только активные действия или все
            If IsNothing(MyActivity) Then
                ComboBox3.Text = "Только активные действия"
            Else
                ComboBox3.Text = MyActivity
            End If
            LoadFlag = 0

            LoadActionList()
            TextBox1.Text = ""
            CheckPlanesButton()
            CheckButtons()

            '-----параметры таблицы
            SetInitTableParams(SfDataGrid1)
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True

            Tab1Settings = 1
        End If
    End Sub

    Private Sub SetInitCustomerTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление начальных параметров элемента SfDataGrid7
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MySfDataGrid.AutoGenerateColumnsMode = AutoGenerateColumnsMode.SmartReset
        MySfDataGrid.Style.HeaderStyle.BackColor = Color.LightGray
        MySfDataGrid.GroupPanel.Height = 50
        MySfDataGrid.HeaderRowHeight = 50
        MySfDataGrid.RowHeight = 50
        MySfDataGrid.AllowResizingColumns = True
        MySfDataGrid.AllowFiltering = True
        MySfDataGrid.ShowRowHeader = True
        MySfDataGrid.AllowTriStateSorting = True
        MySfDataGrid.ShowSortNumbers = True

        Try
            For i As Integer = 1 To 9
                MySfDataGrid.Columns(i).HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
            Next (i)
        Catch
        End Try
        MySfDataGrid.Columns("ScalaCustomerCode").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("ScalaCustomerCode").CellStyle.TextColor = Color.Navy
        MySfDataGrid.Columns("CompanyName").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("CompanyName").CellStyle.TextColor = Color.Navy
        MySfDataGrid.AutoGenerateColumnsMode = AutoGenerateColumnsMode.SmartReset
        MySfDataGrid.SearchController = New SearchControllerExt5(MySfDataGrid)

        MySfDataGrid.SearchController.AllowHighlightSearchText = True
        MySfDataGrid.TableControl.Invalidate()

        '--------------контекстное меню
        MySfDataGrid.RecordContextMenu = New ContextMenuStrip()
        MySfDataGrid.RecordContextMenu.Items.Add("Создать", Nothing, AddressOf OnSfDataGrid7CreateClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Редактировать", Nothing, AddressOf OnSfDataGrid7EditClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Доп. Информация", Nothing, AddressOf OnSfDataGrid7AddInfoClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Объединить", Nothing, AddressOf OnSfDataGrid7UnionClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Удалить", Nothing, AddressOf OnSfDataGrid7DeleteClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять все фильтры", Nothing, AddressOf OnSfDataGrid7RemoveAllFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять Фильтр в текущей колонке", Nothing, AddressOf OnSfDataGrid7RemoveFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Фильтр по выбранному", Nothing, AddressOf OnSfDataGrid7FilterBySelectClicked)
        AddHandler MySfDataGrid.ContextMenuOpening, AddressOf SfDataGrid7_ContextMenuOpening

        '------------тотал
        If MySfDataGrid.TableSummaryRows.Count = 0 Then
            Dim tableSummaryRow1 As New GridTableSummaryRow()
            tableSummaryRow1.Name = "TableSummary"
            tableSummaryRow1.ShowSummaryInRow = True
            tableSummaryRow1.Title = " Количество клиентов: {TotalCompanies}"
            tableSummaryRow1.Position = VerticalPosition.Top

            Dim summaryColumn1 As New GridSummaryColumn()
            summaryColumn1.Name = "TotalCompanies"
            summaryColumn1.SummaryType = SummaryType.CountAggregate
            summaryColumn1.Format = "{Count}"
            summaryColumn1.MappingName = "CompanyID"
            tableSummaryRow1.SummaryColumns.Add(summaryColumn1)
            MySfDataGrid.TableSummaryRows.Add(tableSummaryRow1)
        End If
    End Sub

    Private Sub SetInitProjectTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление начальных параметров элемента SfDataGrid4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MySfDataGrid.AutoGenerateColumnsMode = AutoGenerateColumnsMode.SmartReset
        MySfDataGrid.Style.HeaderStyle.BackColor = Color.LightGray
        MySfDataGrid.GroupPanel.Height = 50
        MySfDataGrid.HeaderRowHeight = 50
        MySfDataGrid.RowHeight = 50
        MySfDataGrid.AllowResizingColumns = True
        MySfDataGrid.AllowFiltering = True
        MySfDataGrid.ShowRowHeader = True
        MySfDataGrid.AllowTriStateSorting = True
        MySfDataGrid.ShowSortNumbers = True

        Try
            For i As Integer = 2 To 24
                MySfDataGrid.Columns(i).HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
            Next (i)
        Catch
        End Try
        MySfDataGrid.Columns("FirstDate").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("FirstDate").CellStyle.TextColor = Color.Navy
        MySfDataGrid.Columns("LastDate").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("LastDate").CellStyle.TextColor = Color.Navy
        MySfDataGrid.Columns("CompanyName").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("CompanyName").CellStyle.TextColor = Color.Black
        MySfDataGrid.Columns("ProjectName").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("ProjectName").CellStyle.TextColor = Color.DarkGreen
        MySfDataGrid.AutoGenerateColumnsMode = AutoGenerateColumnsMode.SmartReset
        MySfDataGrid.SearchController = New SearchControllerExt4(MySfDataGrid)

        MySfDataGrid.SearchController.AllowHighlightSearchText = True
        MySfDataGrid.TableControl.Invalidate()

        '--------------контекстное меню
        MySfDataGrid.RecordContextMenu = New ContextMenuStrip()
        MySfDataGrid.RecordContextMenu.Items.Add("Создать", Nothing, AddressOf OnSfDataGrid4CreateClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Редактировать", Nothing, AddressOf OnSfDataGrid4EditClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Утвердить", Nothing, AddressOf OnSfDataGrid4ApproveClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Удалить", Nothing, AddressOf OnSfDataGrid4DeleteClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять все фильтры", Nothing, AddressOf OnSfDataGrid4RemoveAllFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять Фильтр в текущей колонке", Nothing, AddressOf OnSfDataGrid4RemoveFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Фильтр по выбранному", Nothing, AddressOf OnSfDataGrid4FilterBySelectClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Найти родительский проект", Nothing, AddressOf OnSfDataGrid4FindParentClicked)

        AddHandler MySfDataGrid.ContextMenuOpening, AddressOf SfDataGrid4_ContextMenuOpening
        '------------тотал
        If MySfDataGrid.TableSummaryRows.Count = 0 Then
            Dim tableSummaryRow1 As New GridTableSummaryRow()
            tableSummaryRow1.Name = "TableSummary"
            tableSummaryRow1.ShowSummaryInRow = True
            tableSummaryRow1.Title = " Количество проектов: {TotalProjects}"
            tableSummaryRow1.Position = VerticalPosition.Top

            Dim summaryColumn1 As New GridSummaryColumn()
            summaryColumn1.Name = "TotalProjects"
            summaryColumn1.SummaryType = SummaryType.CountAggregate
            summaryColumn1.Format = "{Count}"
            summaryColumn1.MappingName = "ProjectID"
            tableSummaryRow1.SummaryColumns.Add(summaryColumn1)
            MySfDataGrid.TableSummaryRows.Add(tableSummaryRow1)
        End If
    End Sub

    Private Sub SetSfDataGrid6InitTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление начальных параметров элемента SfDataGrid6
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MySfDataGrid.Style.HeaderStyle.BackColor = Color.LightGray
        MySfDataGrid.AllowResizingColumns = True
        MySfDataGrid.AllowFiltering = True
        MySfDataGrid.ShowRowHeader = True
        MySfDataGrid.AllowSorting = True
        MySfDataGrid.ShowSortNumbers = True
        MySfDataGrid.AllowTriStateSorting = True

        For i As Integer = 0 To 3
            Try
                MySfDataGrid.Columns(i).HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
            Catch
            End Try
        Next i

        '--------------контекстное меню
        MySfDataGrid.RecordContextMenu = New ContextMenuStrip()
        MySfDataGrid.RecordContextMenu.Items.Add("Снять все фильтры", Nothing, AddressOf OnSfDataGrid6RemoveAllFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять Фильтр в текущей колонке", Nothing, AddressOf OnSfDataGrid6RemoveFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Фильтр по выбранному", Nothing, AddressOf OnSfDataGrid6FilterBySelectClicked)

        '-------------Total
        If MySfDataGrid.TableSummaryRows.Count = 0 Then
            Dim tableSummaryRow1 As New GridTableSummaryRow()
            tableSummaryRow1.Name = "TableSummary"
            tableSummaryRow1.ShowSummaryInRow = True
            tableSummaryRow1.Title = " Итого в документах: {TotalSumm}"
            tableSummaryRow1.Position = VerticalPosition.Top

            Dim summaryColumn1 As New GridSummaryColumn()
            summaryColumn1.Name = "TotalSumm"
            summaryColumn1.SummaryType = SummaryType.DoubleAggregate
            summaryColumn1.Format = "{Sum:c}"
            summaryColumn1.MappingName = "DocumentSumm"
            tableSummaryRow1.SummaryColumns.Add(summaryColumn1)
            MySfDataGrid.TableSummaryRows.Add(tableSummaryRow1)
        End If
    End Sub

    Private Sub SetSfDataGrid5InitTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление начальных параметров элемента SfDataGrid5
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MySfDataGrid.Style.HeaderStyle.BackColor = Color.LightGray
        MySfDataGrid.AllowResizingColumns = True
        MySfDataGrid.AllowFiltering = True
        MySfDataGrid.ShowRowHeader = True
        MySfDataGrid.AllowSorting = True
        MySfDataGrid.ShowSortNumbers = True
        MySfDataGrid.AllowTriStateSorting = True

        For i As Integer = 0 To 4
            Try
                MySfDataGrid.Columns(i).HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
            Catch
            End Try
        Next i

        '--------------контекстное меню
        MySfDataGrid.RecordContextMenu = New ContextMenuStrip()
        MySfDataGrid.RecordContextMenu.Items.Add("Снять все фильтры", Nothing, AddressOf OnSfDataGrid5RemoveAllFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять Фильтр в текущей колонке", Nothing, AddressOf OnSfDataGrid5RemoveFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Фильтр по выбранному", Nothing, AddressOf OnSfDataGrid5FilterBySelectClicked)
    End Sub

    Private Sub SetSfDataGrid3InitTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление начальных параметров элемента SfDataGrid3
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MySfDataGrid.Style.HeaderStyle.BackColor = Color.LightGray
        MySfDataGrid.AllowResizingColumns = True
        MySfDataGrid.AllowFiltering = True
        MySfDataGrid.ShowRowHeader = True
        MySfDataGrid.AllowSorting = True
        MySfDataGrid.ShowSortNumbers = True
        MySfDataGrid.AllowTriStateSorting = True

        For i As Integer = 1 To 6
            Try
                MySfDataGrid.Columns(i).HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
            Catch
            End Try
        Next i

        '--------------контекстное меню
        MySfDataGrid.RecordContextMenu = New ContextMenuStrip()
        MySfDataGrid.RecordContextMenu.Items.Add("Снять все фильтры", Nothing, AddressOf OnSfDataGrid3RemoveAllFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять Фильтр в текущей колонке", Nothing, AddressOf OnSfDataGrid3RemoveFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Фильтр по выбранному", Nothing, AddressOf OnSfDataGrid3FilterBySelectClicked)
    End Sub

    Private Sub SetInitScheduleTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление начальных параметров элемента SfDataGrid2
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MySfDataGrid.Style.HeaderStyle.BackColor = Color.LightGray
        MySfDataGrid.GroupPanel.Height = 50
        MySfDataGrid.AllowResizingColumns = True
        MySfDataGrid.AllowFiltering = False
        MySfDataGrid.ShowRowHeader = True
        MySfDataGrid.AllowSorting = False
        MySfDataGrid.ShowSortNumbers = True
        MySfDataGrid.AllowTriStateSorting = True
        MySfDataGrid.SearchController = New SearchControllerExt2(MySfDataGrid)

        '-----Total
        Dim tableSummaryRow1 As New GridTableSummaryRow()
        tableSummaryRow1.Name = "TableSummary"
        tableSummaryRow1.ShowSummaryInRow = True
        tableSummaryRow1.Title = " Количество действий  : {TotalEvents}"
        tableSummaryRow1.Position = VerticalPosition.Top

        Dim summaryColumn1 As New GridSummaryColumn()
        summaryColumn1.Name = "TotalEvents"
        summaryColumn1.SummaryType = SummaryType.Int32Aggregate
        summaryColumn1.Format = "{Sum}"
        summaryColumn1.MappingName = "RowTotal"

        tableSummaryRow1.SummaryColumns.Add(summaryColumn1)
        MySfDataGrid.TableSummaryRows.Add(tableSummaryRow1)

        '--------------контекстное меню
        MySfDataGrid.RecordContextMenu = New ContextMenuStrip()
        MySfDataGrid.RecordContextMenu.Items.Add("Создать", Nothing, AddressOf OnSfDataGrid2CreateClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Редактировать этот день этой компании", Nothing, AddressOf OnSfDataGrid2EditDayClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Редактировать этот месяц этой компании", Nothing, AddressOf OnSfDataGrid2EditMonthClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Все проекты в этом месяце", Nothing, AddressOf OnSfDataGrid2MonthProjectsClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Проекты этой компании в этом месяце", Nothing, AddressOf OnSfDataGrid2CompanyMonthProjectsClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять все фильтры", Nothing, AddressOf OnSfDataGrid2RemoveAllFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять Фильтр в текущей колонке", Nothing, AddressOf OnSfDataGrid2RemoveFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Фильтр по выбранному", Nothing, AddressOf OnSfDataGrid2FilterBySelectClicked)

        AddHandler MySfDataGrid.ContextMenuOpening, AddressOf SfDataGrid2_ContextMenuOpening
    End Sub

    Private Sub SetInitTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление начальных параметров элемента SfDataGrid1
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MySfDataGrid.AutoGenerateColumnsMode = AutoGenerateColumnsMode.SmartReset
        MySfDataGrid.Style.HeaderStyle.BackColor = Color.LightGray
        MySfDataGrid.GroupPanel.Height = 50
        MySfDataGrid.HeaderRowHeight = 50
        MySfDataGrid.RowHeight = 50
        MySfDataGrid.AllowResizingColumns = True
        MySfDataGrid.AllowFiltering = True
        MySfDataGrid.ShowRowHeader = True
        MySfDataGrid.AllowTriStateSorting = True
        MySfDataGrid.ShowSortNumbers = True
        'MySfDataGrid.ShowBusyIndicator = True
        Try
            For i As Integer = 1 To 19
                MySfDataGrid.Columns(i).HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
                MySfDataGrid.Columns(i).HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
            Next (i)
        Catch
        End Try
        MySfDataGrid.Columns("ActionPlannedDate").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("ActionPlannedDate").CellStyle.TextColor = Color.Navy
        MySfDataGrid.Columns("CompanyName").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("CompanyName").CellStyle.TextColor = Color.Black
        MySfDataGrid.Columns("ActionName").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("ActionName").CellStyle.TextColor = Color.Black
        MySfDataGrid.Columns("ProjectInfo").CellStyle.Font.Bold = True
        MySfDataGrid.Columns("ProjectInfo").CellStyle.TextColor = Color.DarkGreen
        MySfDataGrid.AutoGenerateColumnsMode = AutoGenerateColumnsMode.SmartReset
        MySfDataGrid.SearchController = New SearchControllerExt(MySfDataGrid)

        MySfDataGrid.SearchController.AllowHighlightSearchText = True
        MySfDataGrid.TableControl.Invalidate()

        '--------------контекстное меню
        MySfDataGrid.RecordContextMenu = New ContextMenuStrip()
        MySfDataGrid.RecordContextMenu.Items.Add("Создать", Nothing, AddressOf OnCreateClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Копировать", Nothing, AddressOf OnCopyClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Редактировать", Nothing, AddressOf OnEditClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Просмотр", Nothing, AddressOf OnViewClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Передать", Nothing, AddressOf OnTransferClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Удалить", Nothing, AddressOf OnDeleteClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять все фильтры", Nothing, AddressOf OnRemoveAllFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять Фильтр в текущей колонке", Nothing, AddressOf OnRemoveFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Фильтр по выбранному", Nothing, AddressOf OnFilterBySelectClicked)

        AddHandler MySfDataGrid.ContextMenuOpening, AddressOf SfDataGrid1_ContextMenuOpening

        '------------тотал
        'Dim tableSummaryRow1 As New GridTableSummaryRow()
        'tableSummaryRow1.Name = "TableSummary"
        'tableSummaryRow1.ShowSummaryInRow = False
        'tableSummaryRow1.Position = VerticalPosition.Top

        'Dim summaryColumn1 As New GridSummaryColumn()
        'summaryColumn1.Name = "TotalRecords"
        'summaryColumn1.SummaryType = SummaryType.CountAggregate
        'summaryColumn1.Format = "Всего записей: {Count}"
        'summaryColumn1.MappingName = "ActionPlannedDate"


        'tableSummaryRow1.SummaryColumns.Add(summaryColumn1)
        'MySfDataGrid.TableSummaryRows.Add(tableSummaryRow1)

        If MySfDataGrid.TableSummaryRows.Count = 0 Then
            Dim tableSummaryRow1 As New GridTableSummaryRow()
            tableSummaryRow1.Name = "TableSummary"
            tableSummaryRow1.ShowSummaryInRow = True
            tableSummaryRow1.Title = " Количество записей: {TotalEvents}"
            tableSummaryRow1.Position = VerticalPosition.Top

            Dim summaryColumn1 As New GridSummaryColumn()
            summaryColumn1.Name = "TotalEvents"
            summaryColumn1.SummaryType = SummaryType.CountAggregate
            summaryColumn1.Format = "{Count:d}"
            summaryColumn1.MappingName = "ActionPlannedDate"
            tableSummaryRow1.SummaryColumns.Add(summaryColumn1)
            MySfDataGrid.TableSummaryRows.Add(tableSummaryRow1)
        End If
    End Sub

    Private Sub SetCustomerTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление параметров элемента SfDataGrid7
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 1 To 8
            MySfDataGrid.Columns(i).AllowTextWrapping = True
            MySfDataGrid.Columns(i).AllowHeaderTextWrapping = True
        Next i

        '-----CompanyID
        MySfDataGrid.Columns("CompanyID").HeaderText = "CID"
        MySfDataGrid.Columns("CompanyID").Visible = False
        '-----ScalaCustomerCode
        MySfDataGrid.Columns("ScalaCustomerCode").HeaderText = "Код Скала"
        MySfDataGrid.Columns("ScalaCustomerCode").Width = 80
        '-----CompanyName
        MySfDataGrid.Columns("CompanyName").HeaderText = "Компания"
        MySfDataGrid.Columns("CompanyName").Width = 250
        '-----CompanyAddress
        MySfDataGrid.Columns("CompanyAddress").HeaderText = "Адрес"
        MySfDataGrid.Columns("CompanyAddress").Width = 450
        '-----CompanyPhone
        MySfDataGrid.Columns("CompanyPhone").HeaderText = "Телефон"
        MySfDataGrid.Columns("CompanyPhone").Width = 150
        '-----CompanyEMail
        MySfDataGrid.Columns("CompanyEMail").HeaderText = "E-Mail"
        MySfDataGrid.Columns("CompanyEMail").Width = 180
        '-----CustomerGroup
        MySfDataGrid.Columns("CustomerGroup").HeaderText = "Группа Rexel"
        MySfDataGrid.Columns("CustomerGroup").Width = 400
        '-----EndMarket
        MySfDataGrid.Columns("EndMarket").HeaderText = "Рынок Rexel"
        MySfDataGrid.Columns("EndMarket").Width = 150
        '-----IsIKA
        MySfDataGrid.Columns("IsIKA").HeaderText = "Вид КА"
        MySfDataGrid.Columns("IsIKA").Width = 200
        '-----Potencial
        MySfDataGrid.Columns("Potencial").HeaderText = "Потенциал (Руб)"
        MySfDataGrid.Columns("Potencial").Width = 150


    End Sub

    Private Sub SetProjectTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление параметров элемента SfDataGrid4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 2 To 24
            MySfDataGrid.Columns(i).AllowTextWrapping = True
            MySfDataGrid.Columns(i).AllowHeaderTextWrapping = True
        Next i

        '-----названия и параметры колонок
        '-----ProjectID
        MySfDataGrid.Columns("ProjectID").HeaderText = "PID"
        MySfDataGrid.Columns("ProjectID").Visible = False
        '-----CompanyID
        MySfDataGrid.Columns("CompanyID").HeaderText = "CID"
        MySfDataGrid.Columns("CompanyID").Visible = False
        '-----ScalaCustomerCode
        MySfDataGrid.Columns("ScalaCustomerCode").HeaderText = "Код Скала"
        MySfDataGrid.Columns("ScalaCustomerCode").Width = 80
        '-----CompanyName
        MySfDataGrid.Columns("CompanyName").HeaderText = "Компания"
        MySfDataGrid.Columns("CompanyName").Width = 150
        '-----ProjectName
        MySfDataGrid.Columns("ProjectName").HeaderText = "Проект"
        MySfDataGrid.Columns("ProjectName").Width = 150
        '-----ProjectSumm
        MySfDataGrid.Columns("ProjectSumm").HeaderText = "Сумма"
        MySfDataGrid.Columns("ProjectSumm").Width = 90
        '-----ProjectComment
        MySfDataGrid.Columns("ProjectComment").HeaderText = "Комментарий"
        MySfDataGrid.Columns("ProjectComment").Width = 150
        '-----FirstDate
        MySfDataGrid.Columns("FirstDate").HeaderText = "Начало"
        MySfDataGrid.Columns("FirstDate").Width = 80
        '-----LastDate
        MySfDataGrid.Columns("LastDate").HeaderText = "Окончание"
        MySfDataGrid.Columns("LastDate").Width = 80
        '-----StartDate
        MySfDataGrid.Columns("StartDate").HeaderText = "Занесен"
        MySfDataGrid.Columns("StartDate").Width = 80
        '-----CloseDate
        MySfDataGrid.Columns("CloseDate").HeaderText = "Закрыт"
        MySfDataGrid.Columns("CloseDate").Width = 80
        '-----ProposalDate
        MySfDataGrid.Columns("ProposalDate").HeaderText = "Подача предложения к:"
        MySfDataGrid.Columns("ProposalDate").Width = 120
        '-----ProjectAddr
        MySfDataGrid.Columns("ProjectAddr").HeaderText = "Адрес"
        MySfDataGrid.Columns("ProjectAddr").Width = 150
        '-----Investor
        MySfDataGrid.Columns("Investor").HeaderText = "Инвестор"
        MySfDataGrid.Columns("Investor").Width = 100
        '-----Contractor
        MySfDataGrid.Columns("Contractor").HeaderText = "Контрактор"
        MySfDataGrid.Columns("Contractor").Width = 100
        '-----ResponciblePerson
        MySfDataGrid.Columns("ResponciblePerson").HeaderText = "Ответственный"
        MySfDataGrid.Columns("ResponciblePerson").Width = 100
        '-----ManufacturersList
        MySfDataGrid.Columns("ManufacturersList").HeaderText = "Производители"
        MySfDataGrid.Columns("ManufacturersList").Width = 100
        '-----AlterManufacturers
        MySfDataGrid.Columns("AlterManufacturers").HeaderText = "Возможна альтернатива"
        MySfDataGrid.Columns("AlterManufacturers").Width = 100
        '-----Competitors
        MySfDataGrid.Columns("Competitors").HeaderText = "Конкуренты"
        MySfDataGrid.Columns("Competitors").Width = 100
        '-----AdditionalExpencesPerCent
        MySfDataGrid.Columns("AdditionalExpencesPerCent").HeaderText = "Доп расходы (%)"
        MySfDataGrid.Columns("AdditionalExpencesPerCent").Width = 100
        '-----IsApproved
        MySfDataGrid.Columns("IsApproved").HeaderText = "Утвержден"
        MySfDataGrid.Columns("IsApproved").Width = 100
        '-----ProjectStage
        MySfDataGrid.Columns("ProjectStage").HeaderText = "Стадия проекта"
        MySfDataGrid.Columns("ProjectStage").Width = 120
        '-----ParentProjectID
        MySfDataGrid.Columns("ParentProjectID").HeaderText = "PPID"
        MySfDataGrid.Columns("ParentProjectID").Visible = False
        '-----ParentProjectName
        MySfDataGrid.Columns("ParentProjectName").HeaderText = "Родительский проект"
        MySfDataGrid.Columns("ParentProjectName").Width = 120
        '-----InvestProject
        MySfDataGrid.Columns("InvestProject").HeaderText = "Проект с услугами"
        MySfDataGrid.Columns("InvestProject").Width = 120

    End Sub

    Private Sub SetSfDataGrid6TableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление параметров элемента SfDataGrid6
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        '-----названия и параметры колонок
        '-----DocumentNumber
        MySfDataGrid.Columns("DocumentNumber").HeaderText = "Документ"
        MySfDataGrid.Columns("DocumentNumber").Width = 200
        '-----DocumentDate
        MySfDataGrid.Columns("DocumentDate").HeaderText = "Дата"
        MySfDataGrid.Columns("DocumentDate").Width = 200
        '-----DocumentSumm
        MySfDataGrid.Columns("DocumentSumm").HeaderText = "Сумма (Руб)"
        MySfDataGrid.Columns("DocumentSumm").Width = 200
        '-----SalesmanName
        MySfDataGrid.Columns("SalesmanName").HeaderText = "Продавец"
        MySfDataGrid.Columns("SalesmanName").Width = 200
    End Sub

    Private Sub SetSfDataGrid5TableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление параметров элемента SfDataGrid5
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        '-----названия и параметры колонок
        '-----DirectionName
        MySfDataGrid.Columns("DirectionName").HeaderText = "Направление"
        MySfDataGrid.Columns("DirectionName").Width = 150
        '-----EventTypeName
        MySfDataGrid.Columns("EventTypeName").HeaderText = "Активность"
        MySfDataGrid.Columns("EventTypeName").Width = 150
        '-----CompanyName
        MySfDataGrid.Columns("CompanyName").HeaderText = "Компания"
        MySfDataGrid.Columns("CompanyName").Width = 260
        '-----ActionTime
        MySfDataGrid.Columns("ActionTime").HeaderText = "Дата"
        MySfDataGrid.Columns("ActionTime").Width = 150
        '-----FullName
        MySfDataGrid.Columns("FullName").HeaderText = "Продавец"
        MySfDataGrid.Columns("FullName").Width = 150
    End Sub

    Private Sub SetSfDataGrid3TableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление параметров элемента SfDataGrid3
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        '-----названия и параметры колонок
        '-----CompanyID
        MySfDataGrid.Columns("CompanyID").HeaderText = "ID"
        MySfDataGrid.Columns("CompanyID").Visible = False
        '-----DocumentType
        MySfDataGrid.Columns("DocumentType").HeaderText = "Документ"
        MySfDataGrid.Columns("DocumentType").Width = 200
        '-----DocumentNumber
        MySfDataGrid.Columns("DocumentNumber").HeaderText = "Номер документа"
        MySfDataGrid.Columns("DocumentNumber").Width = 200
        '-----DocumentDate
        MySfDataGrid.Columns("DocumentDate").HeaderText = "Дата"
        MySfDataGrid.Columns("DocumentDate").Width = 200
        '-----DocumentSumm
        MySfDataGrid.Columns("DocumentSumm").HeaderText = "Сумма (Руб)"
        MySfDataGrid.Columns("DocumentSumm").Width = 200
        '-----SalesmanName
        MySfDataGrid.Columns("SalesmanName").HeaderText = "Продавец"
        MySfDataGrid.Columns("SalesmanName").Width = 200
    End Sub

    Private Sub SetScheduleTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление параметров элемента SfDataGrid2
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        'MySfDataGrid.AutoGenerateColumnsMode = AutoGenerateColumnsMode.SmartReset

        MySfDataGrid.Columns("CompanyScalaCode").AllowFiltering = True
        Try
            MySfDataGrid.Columns("CompanyScalaCode").HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
            MySfDataGrid.Columns("CompanyScalaCode").HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
            MySfDataGrid.Columns("CompanyScalaCode").HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
        Catch
        End Try
        MySfDataGrid.Columns("CompanyName").AllowFiltering = True
        Try
            MySfDataGrid.Columns("CompanyName").HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
            MySfDataGrid.Columns("CompanyName").HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
            MySfDataGrid.Columns("CompanyName").HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
        Catch
        End Try
        MySfDataGrid.Columns("CustProject").AllowFiltering = True
        Try
            MySfDataGrid.Columns("CustProject").HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
            MySfDataGrid.Columns("CustProject").HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
            MySfDataGrid.Columns("CustProject").HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
        Catch
        End Try
        MySfDataGrid.Columns("OrdersQTY").AllowFiltering = True
        Try
            MySfDataGrid.Columns("OrdersQTY").HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
            MySfDataGrid.Columns("OrdersQTY").HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
            MySfDataGrid.Columns("OrdersQTY").HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
        Catch
        End Try
        MySfDataGrid.Columns("Status").AllowFiltering = True
        Try
            MySfDataGrid.Columns("Status").HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
            MySfDataGrid.Columns("Status").HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
            MySfDataGrid.Columns("Status").HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
        Catch
        End Try
        MySfDataGrid.Columns("RowTotal").AllowFiltering = True
        Try
            MySfDataGrid.Columns("RowTotal").HeaderStyle.FilterIcon = New Bitmap(Image.FromFile("FilterIcon.png"))
            MySfDataGrid.Columns("RowTotal").HeaderStyle.FilteredIcon = New Bitmap(Image.FromFile("FilteredIcon.png"))
            MySfDataGrid.Columns("RowTotal").HeaderStyle.SortIcon = New Bitmap(Image.FromFile("ArrowUp.ico"))
        Catch
        End Try
        MySfDataGrid.Columns("CompanyScalaCode").AllowSorting = True
        MySfDataGrid.Columns("CompanyName").AllowSorting = True
        MySfDataGrid.Columns("CustProject").AllowSorting = True
        MySfDataGrid.Columns("OrdersQTY").AllowSorting = True
        MySfDataGrid.Columns("Status").AllowSorting = True
        MySfDataGrid.Columns("RowTotal").AllowSorting = True

        For i As Integer = 1 To MySfDataGrid.ColumnCount - 2
            MySfDataGrid.Columns(i).AllowTextWrapping = True
            MySfDataGrid.Columns(i).AllowHeaderTextWrapping = True
        Next i

        '-----названия и параметры колонок
        For i As Integer = 1 To MySfDataGrid.ColumnCount - 2
            MySfDataGrid.Columns(i).Width = 40
            If MySfDataGrid.Columns(i).MappingName.Equals("CompanyID") Or _
                MySfDataGrid.Columns(i).MappingName.Equals("CompanyScalaCode") Or _
                MySfDataGrid.Columns(i).MappingName.Equals("CompanyName") Or _
                MySfDataGrid.Columns(i).MappingName.Equals("CustProject") Or _
                MySfDataGrid.Columns(i).MappingName.Equals("OrdersQTY") Or _
                MySfDataGrid.Columns(i).MappingName.Equals("Status") Or _
                MySfDataGrid.Columns(i).MappingName.Equals("RowTotal") Then
                MySfDataGrid.Columns(i).HeaderStyle.BackColor = Color.LightGray
            Else
                If MySfDataGrid.Columns(i).MappingName.Contains("Сб") Or _
                    MySfDataGrid.Columns(i).MappingName.Contains("Вс") Then
                    MySfDataGrid.Columns(i).HeaderStyle.BackColor = Color.FromArgb(200, 185, 255, 171)
                Else
                    MySfDataGrid.Columns(i).HeaderStyle.BackColor = Color.White
                End If
            End If
        Next i

        '-----CompanyID
        MySfDataGrid.Columns("CompanyID").HeaderText = "ID"
        MySfDataGrid.Columns("CompanyID").Visible = False
        '-----CompanyScalaCode
        MySfDataGrid.Columns("CompanyScalaCode").HeaderText = "Код Скала"
        MySfDataGrid.Columns("CompanyScalaCode").Width = 88
        '-----CompanyName
        MySfDataGrid.Columns("CompanyName").HeaderText = "Компания"
        MySfDataGrid.Columns("CompanyName").Width = 200
        '-----CustProject
        MySfDataGrid.Columns("CustProject").HeaderText = "Проекты"
        MySfDataGrid.Columns("CustProject").Width = 88
        '-----OrdersQTY
        MySfDataGrid.Columns("OrdersQTY").HeaderText = "Scala"
        MySfDataGrid.Columns("OrdersQTY").Width = 80
        MySfDataGrid.Columns("OrdersQTY").ShowHeaderToolTip = True
        MySfDataGrid.Style.ToolTipStyle.BorderThickness = 2
        MySfDataGrid.Style.ToolTipStyle.BorderColor = Color.DarkBlue
        MySfDataGrid.ToolTipOption.AutoPopDelay = 10000
        '-----Status
        MySfDataGrid.Columns("Status").HeaderText = "Статус"
        MySfDataGrid.Columns("Status").Width = 80
        '-----Дни
        For i As Integer = 6 To MySfDataGrid.ColumnCount - 2
            MySfDataGrid.Columns(i).HeaderText = Replace(Replace(MySfDataGrid.Columns(i).MappingName, "d", ""), "_", Chr(10))
        Next
        '-----RowTotal
        MySfDataGrid.Columns("RowTotal").HeaderText = "Итого"
        MySfDataGrid.Columns("RowTotal").Width = 80

        MySfDataGrid.FrozenColumnCount = 6
        MySfDataGrid.FooterColumnCount = 1
        MySfDataGrid.Style.FreezePaneLineStyle.Weight = 3

    End Sub

    Private Sub SetTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление параметров элемента SfDataGrid1
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 1 To 19
            MySfDataGrid.Columns(i).AllowTextWrapping = True
            MySfDataGrid.Columns(i).AllowHeaderTextWrapping = True
        Next i


        '-----названия и параметры колонок
        '-----EventID
        MySfDataGrid.Columns("EventID").HeaderText = "ID"
        MySfDataGrid.Columns("EventID").Visible = False
        '-----ActionPlannedDate
        MySfDataGrid.Columns("ActionPlannedDate").HeaderText = "Планируемая дата"
        MySfDataGrid.Columns("ActionPlannedDate").Width = 100

        '-----DirectionName
        MySfDataGrid.Columns("DirectionName").HeaderText = "Направление"
        MySfDataGrid.Columns("DirectionName").Width = 90
        '-----EventTypeName
        MySfDataGrid.Columns("EventTypeName").HeaderText = "Способ"
        MySfDataGrid.Columns("EventTypeName").Width = 90
        '-----ScalaCustomerCode
        MySfDataGrid.Columns("ScalaCustomerCode").HeaderText = "Код Скала"
        MySfDataGrid.Columns("ScalaCustomerCode").Width = 80
        '-----CompanyName
        MySfDataGrid.Columns("CompanyName").HeaderText = "Компания"
        MySfDataGrid.Columns("CompanyName").Width = 150
        '-----ContactName
        MySfDataGrid.Columns("ContactName").HeaderText = "Ф.И.О. контакт"
        MySfDataGrid.Columns("ContactName").Width = 100
        '-----ContactPhone
        MySfDataGrid.Columns("ContactPhone").HeaderText = "Телефон"
        MySfDataGrid.Columns("ContactPhone").Width = 100
        '-----ContactEMail
        MySfDataGrid.Columns("ContactEMail").HeaderText = "E mail"
        MySfDataGrid.Columns("ContactEMail").Width = 100
        '-----ActionName
        MySfDataGrid.Columns("ActionName").HeaderText = "Цель"
        MySfDataGrid.Columns("ActionName").Width = 100
        '-----ProjectInfo
        MySfDataGrid.Columns("ProjectInfo").HeaderText = "Проект"
        MySfDataGrid.Columns("ProjectInfo").Width = 100
        '-----ActionSumm
        MySfDataGrid.Columns("ActionSumm").HeaderText = "Сумма вопроса"
        MySfDataGrid.Columns("ActionSumm").Width = 100
        '-----ActionComments
        MySfDataGrid.Columns("ActionComments").HeaderText = "Комментарии"
        MySfDataGrid.Columns("ActionComments").Width = 200
        '-----ActionResultName
        MySfDataGrid.Columns("ActionResultName").HeaderText = "Результат"
        MySfDataGrid.Columns("ActionResultName").Width = 200
        '-----FullName
        MySfDataGrid.Columns("FullName").HeaderText = "Продавец"
        MySfDataGrid.Columns("FullName").Width = 100
        '-----CompanyAddress
        MySfDataGrid.Columns("CompanyAddress").HeaderText = "Адрес компании"
        MySfDataGrid.Columns("CompanyAddress").Width = 150
        '-----CompanyPhone
        MySfDataGrid.Columns("CompanyPhone").HeaderText = "Телефон компании"
        MySfDataGrid.Columns("CompanyPhone").Width = 100
        '-----CompanyEMail
        MySfDataGrid.Columns("CompanyEMail").HeaderText = "Email компании"
        MySfDataGrid.Columns("CompanyEMail").Width = 100
        '-----IsIKA
        MySfDataGrid.Columns("IsIKA").HeaderText = "Ключевой клиент"
        MySfDataGrid.Columns("IsIKA").Width = 150
        '-----IsApproved
        MySfDataGrid.Columns("IsApproved").HeaderText = "Утверждено"
        MySfDataGrid.Columns("IsApproved").Width = 90

    End Sub

    Private Sub sfDataGrid7_QueryRowHeight(ByVal sender As Object, ByVal e As QueryRowHeightEventArgs)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка высоты строк, чтобы содержимое помещалось полностью. sfDataGrid7
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid7.AutoSizeController.GetAutoRowHeight(e.RowIndex, autoFitOptions, autoHeight) Then
            If e.RowIndex <> 0 Then
                If autoHeight > 24 Then
                    e.Height = autoHeight
                    e.Handled = True
                End If
            Else
                e.Height = autoHeight + 20
                e.Handled = True
            End If
        End If

    End Sub

    Private Sub sfDataGrid4_QueryRowHeight(ByVal sender As Object, ByVal e As QueryRowHeightEventArgs)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка высоты строк, чтобы содержимое помещалось полностью. sfDataGrid4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid4.AutoSizeController.GetAutoRowHeight(e.RowIndex, autoFitOptions, autoHeight) Then
            If e.RowIndex <> 0 Then
                If autoHeight > 24 Then
                    e.Height = autoHeight
                    e.Handled = True
                End If
            Else
                e.Height = autoHeight + 20
                e.Handled = True
            End If
        End If

    End Sub

    Private Sub sfDataGrid2_QueryRowHeight(ByVal sender As Object, ByVal e As QueryRowHeightEventArgs)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка высоты строк, чтобы содержимое помещалось полностью. sfDataGrid2
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid2.AutoSizeController.GetAutoRowHeight(e.RowIndex, autoFitOptions, autoHeight) Then
            If e.RowIndex <> 0 Then
                If autoHeight > 24 Then
                    e.Height = autoHeight
                    e.Handled = True
                End If
            Else
                e.Height = autoHeight + 20
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub sfDataGrid1_QueryRowHeight(ByVal sender As Object, ByVal e As QueryRowHeightEventArgs)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка высоты строк, чтобы содержимое помещалось полностью. sfDataGrid1
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid1.AutoSizeController.GetAutoRowHeight(e.RowIndex, autoFitOptions, autoHeight) Then
            If e.RowIndex <> 0 Then
                If autoHeight > 24 Then
                    e.Height = autoHeight
                    e.Handled = True
                End If
            Else
                'e.Height = (((autoHeight + 10) / 5) ^ 2)
                e.Height = autoHeight + 20
                'e.Height = 48
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub SfDataGrid7_QueryRowStyle(ByVal sender As Object, ByVal e As QueryRowStyleEventArgs)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк SfDataGrid7
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.RowType = RowType.DefaultRow Then
            If e.RowData.ScalaCustomerCode = "" Then
                e.Style.BackColor = Color.LightYellow
            Else
                e.Style.BackColor = Color.White
            End If
        End If
    End Sub

    Private Sub SfDataGrid4_QueryRowStyle(ByVal sender As Object, ByVal e As QueryRowStyleEventArgs)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк SfDataGrid4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.RowType = RowType.DefaultRow Then
            If e.RowData.IsApproved Then
                e.Style.BackColor = Color.White
            Else
                e.Style.BackColor = Color.LightYellow
            End If
        End If
    End Sub

    Private Sub SfDataGrid2_QueryCellStyle(ByVal sender As Object, ByVal e As QueryCellStyleEventArgs)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк SfDataGrid2
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.Column.MappingName.Equals("CompanyID") Or _
            e.Column.MappingName.Equals("CompanyScalaCode") Or _
            e.Column.MappingName.Equals("CompanyName") Then
            e.Style.BackColor = Color.FromArgb(100, 232, 232, 232)
        ElseIf e.Column.MappingName.Equals("CustProject") Or _
            e.Column.MappingName.Equals("OrdersQTY") Then
            If e.DisplayText.Equals("0") = False Then
                e.Style.BackColor = Color.Blue
                e.Style.TextColor = Color.White
                e.Style.Font.Bold = True
                e.Style.HorizontalAlignment = HorizontalAlignment.Center
            Else
                e.Style.BackColor = Color.FromArgb(100, 232, 232, 232)
            End If
        ElseIf e.Column.MappingName.Equals("RowTotal") Then
            If e.DisplayText.Equals("0") = False Then
                Dim rowData = SfDataGrid2.GetRecordAtRowIndex(e.RowIndex)
                If rowData.GetItem("Status").Equals("План") Then
                    e.Style.BackColor = Color.FromArgb(200, 61, 8, 114)
                Else
                    e.Style.BackColor = Color.Green
                End If
                e.Style.TextColor = Color.White
                e.Style.Font.Bold = True
                e.Style.HorizontalAlignment = HorizontalAlignment.Center
            Else
                e.Style.BackColor = Color.FromArgb(100, 229, 224, 236)
                e.Style.HorizontalAlignment = HorizontalAlignment.Center
            End If
        ElseIf e.Column.MappingName.Equals("Status") Then
            If e.DisplayText.Equals("План") Then
                e.Style.BackColor = Color.FromArgb(100, 206, 219, 254)
            Else
                e.Style.BackColor = Color.White
            End If
        Else
            If e.DisplayText.Equals("") = False Then
                Dim rowData = SfDataGrid2.GetRecordAtRowIndex(e.RowIndex)
                If rowData.GetItem("Status").Equals("План") Then
                    e.Style.BackColor = Color.FromArgb(200, 61, 8, 114)
                Else
                    e.Style.BackColor = Color.Green
                End If
                e.Style.TextColor = Color.White
                e.Style.Font.Bold = True
                e.Style.HorizontalAlignment = HorizontalAlignment.Center
            Else
                If e.Column.MappingName.Contains("Сб") Or _
                    e.Column.MappingName.Contains("Вс") Then
                    e.Style.BackColor = Color.FromArgb(100, 227, 255, 221)
                Else
                    Dim rowData = SfDataGrid2.GetRecordAtRowIndex(e.RowIndex)
                    If rowData.GetItem("Status").Equals("План") Then
                        e.Style.BackColor = Color.FromArgb(100, 206, 219, 254)
                    Else
                        e.Style.BackColor = Color.White
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub SfDataGrid1_QueryRowStyle(ByVal sender As Object, ByVal e As QueryRowStyleEventArgs)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсветка строк SfDataGrid1
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.RowType = RowType.DefaultRow Then
            '----при работе с таблицей
            'Dim dataRowView = TryCast(e.RowData, DataRowView)
            'Dim dataRow = dataRowView.Row
            'Dim cellValue = dataRow("ActionResultName").ToString()
            'If cellValue.Equals("") Then
            '    e.Style.BackColor = Color.FromArgb(100, 227, 255, 221)
            'Else
            '    e.Style.BackColor = Color.FromArgb(100, 232, 232, 232)
            'End If
            '-----при работе с листом
            If e.RowData.ActionResultName.Equals("") = False Then
                '-----Закрытое действие
                e.Style.BackColor = Color.FromArgb(100, 232, 232, 232)
            Else
                '-----действие в работе
                If DateAdd(DateInterval.Day, 1, e.RowData.ActionPlannedDate) > Now() Then
                    '-----не просрочено
                    If e.RowData.IsApproved = True Then
                        '-----Утверждено
                        e.Style.BackColor = Color.FromArgb(100, 227, 255, 221)
                    Else
                        '-----Не утверждено
                        e.Style.BackColor = Color.White
                    End If
                ElseIf DateDiff(DateInterval.Day, e.RowData.ActionPlannedDate, Now()) < 8 Then
                    '-----Просрочено до 8 дней
                    If e.RowData.IsApproved = True Then
                        '-----Утверждено
                        e.Style.BackColor = Color.FromArgb(100, 255, 246, 217)
                    Else
                        '-----Не утверждено
                        e.Style.BackColor = Color.LightYellow
                    End If
                Else
                    '-----Просрочено более 8 дней
                    If e.RowData.IsApproved = True Then
                        '-----Утверждено
                        e.Style.BackColor = Color.FromArgb(100, 254, 217, 214)
                    Else
                        '-----Не утверждено
                        e.Style.BackColor = Color.LightYellow
                    End If

                End If

            End If
        End If
    End Sub

    Private Sub CheckProjectApprovButton()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление доступности кнопки утверждения проектов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If (Declarations.MyPDPermission = True) Then
            If SfDataGrid4.SelectedItem Is Nothing Then
                '-----Ничего не выбрано
                Button19.Enabled = False
                Button19.Visible = False
            Else
                If SfDataGrid4.SelectedItem.IsApproved = True Then
                    '-----Уже утвержден
                    Button19.Enabled = False
                    Button19.Visible = False
                Else
                    '-----Не утвержден
                    Button19.Enabled = True
                    Button19.Visible = True
                End If
            End If
        Else
            Button19.Enabled = False
            Button19.Visible = False
        End If
    End Sub


    Private Sub CheckPlanesButtonSf2()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление доступности кнопки утверждения планов на закладке план на месяц
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        If (Declarations.MyCCPermission = True Or Declarations.MyPermission = True) Then
            Button13.Enabled = True
            Button13.Visible = True
        Else
            Button13.Enabled = False
            Button13.Visible = False
        End If
    End Sub

    Private Sub CheckPlanesButton()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление доступности кнопки утверждения планов на закладке список действий
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        If (Declarations.MyCCPermission = True Or Declarations.MyPermission = True) Then
            Button4.Enabled = True
            Button4.Visible = True
        Else
            Button4.Enabled = False
            Button4.Visible = False
        End If
    End Sub

    Private Sub LoadActionList()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка действий
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        If LoadFlag = 0 Then
            MySQLStr = "SELECT tbl_CRM_Events.EventID, "
            MySQLStr = MySQLStr & "tbl_CRM_Events.ActionPlannedDate, "
            MySQLStr = MySQLStr & "tbl_CRM_Directions.DirectionName, "
            MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_EventTypes.EventTypeID = 999999 THEN tbl_CRM_Events.EventTypeDescription ELSE tbl_CRM_EventTypes.EventTypeName END AS EventTypeName, "
            MySQLStr = MySQLStr & "Ltrim(Rtrim(ISNULL(tbl_CRM_Companies.ScalaCustomerCode, ''))) as ScalaCustomerCode, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies.CompanyName, '') as CompanyName, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Contacts.ContactName, '') as ContactName, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Contacts.ContactPhone, '') as ContactPhone, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Contacts.ContactEMail, '') AS ContactEMail, "
            MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_Actions.ActionID = 999999 THEN tbl_CRM_Events.ActionDescription ELSE tbl_CRM_Actions.ActionName END AS ActionName, "
            MySQLStr = MySQLStr & "LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(tbl_CRM_Projects.ProjectName, ''))) + ' ' + LTRIM(RTRIM(ISNULL(tbl_CRM_Projects.ProjectComment, ''))))) AS ProjectInfo, "
            MySQLStr = MySQLStr & "tbl_CRM_Events.ActionSumm, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Events.ActionComments, '') AS ActionComments, "
            MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_ActionsResultTypes.ActionResultID = 999999 THEN ISNULL(tbl_CRM_Events.ActionResultDescription, '') ELSE ISNULL(tbl_CRM_ActionsResultTypes.ActionResultName, '') END AS ActionResultName, "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers.FullName, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies.CompanyAddress, '') as CompanyAddress, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies.CompanyPhone, '') AS CompanyPhone, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies.CompanyEMail, '') AS CompanyEMail, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies_Ext.IsIKA, N'') AS IsIKA, tbl_CRM_Events.IsApproved "
            MySQLStr = MySQLStr & "FROM tbl_CRM_ActionsResultTypes WITH (NOLOCK) RIGHT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Events INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_EventTypes ON tbl_CRM_Events.EventTypeID = tbl_CRM_EventTypes.EventTypeID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Directions ON tbl_CRM_Events.DirectionID = tbl_CRM_Directions.DirectionID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Contacts ON tbl_CRM_Events.ContactID = tbl_CRM_Contacts.ContactID AND tbl_CRM_Companies.CompanyID = tbl_CRM_Contacts.CompanyID AND "
            MySQLStr = MySQLStr & "tbl_CRM_Events.CompanyID = tbl_CRM_Contacts.CompanyID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Actions ON tbl_CRM_Events.ActionID = tbl_CRM_Actions.ActionID INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.UserID = ScalaSystemDB.dbo.ScaUsers.UserID INNER JOIN "
            MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 ON tbl_CRM_ActionsResultTypes.ActionResultID = tbl_CRM_Events.ActionResultID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Projects ON tbl_CRM_Events.ProjectID = tbl_CRM_Projects.ProjectID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies_Ext ON tbl_CRM_Companies.CompanyID = tbl_CRM_Companies_Ext.CompanyID "
            MySQLStr = MySQLStr & " WHERE (tbl_CRM_Events.ActionPlannedDate >= CONVERT(DATETIME, '" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103)) "
            MySQLStr = MySQLStr & " AND (tbl_CRM_Events.ActionPlannedDate <= CONVERT(DATETIME, '" & Format(DateTimePicker2.Value, "dd/MM/yyyy") & "', 103)) "
            If ComboBox1.SelectedValue = 0 Then                                 '-----Выбраны все продавцы
                If Declarations.MyPermission = True Then                        '-----Вообще все продавцы
                Else                                                            '-----Все продавцы принадлежащие к кост центру
                    If ComboBox2.Text = "Только выбранного продавца" Then       '-----Только по продавцам, принадлежащим к кост центру
                        MySQLStr = MySQLStr & "AND (SUBSTRING(ST010300.ST01021, 7, 3) = N'" & Declarations.CC & "') "
                    Else                                                        '-----По всем продавцам для клиентов кост центра
                        MySQLStr = MySQLStr & "AND (tbl_CRM_Events.CompanyID IN(SELECT tbl_CRM_Events.CompanyID "
                        MySQLStr = MySQLStr & "FROM tbl_CRM_Events INNER JOIN "
                        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.UserID = ScalaSystemDB.dbo.ScaUsers.UserID INNER JOIN "
                        MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 "
                        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Events.ActionPlannedDate >= CONVERT(DATETIME, '" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103)) "
                        MySQLStr = MySQLStr & "AND (tbl_CRM_Events.ActionPlannedDate <= CONVERT(DATETIME, '" & Format(DateTimePicker2.Value, "dd/MM/yyyy") & "', 103)) "
                        MySQLStr = MySQLStr & "AND (SUBSTRING(ST010300.ST01021, 7, 3) = N'" & Declarations.CC & "') "
                        MySQLStr = MySQLStr & "GROUP BY tbl_CRM_Events.CompanyID "
                        MySQLStr = MySQLStr & "UNION "
                        MySQLStr = MySQLStr & "Select tbl_CRM_Companies.CompanyID "
                        MySQLStr = MySQLStr & "FROM tbl_CRM_Companies INNER JOIN "
                        MySQLStr = MySQLStr & "SL010300 ON tbl_CRM_Companies.ScalaCustomerCode = SL010300.SL01001 INNER JOIN "
                        MySQLStr = MySQLStr & "ST010300 AS ST010300_1 ON SL010300.SL01035 = ST010300_1.ST01001 "
                        MySQLStr = MySQLStr & "WHERE (SUBSTRING(ST010300_1.ST01021, 7, 3) = N'" & Declarations.CC & "'))) "
                    End If
                End If
            Else                                                                '-----Выбран один продавец
                If ComboBox2.Text = "Только выбранного продавца" Then           '-----Только по выбранному продавцу
                    MySQLStr = MySQLStr & "AND (tbl_CRM_Events.UserID = " & ComboBox1.SelectedValue & ")"
                Else                                                            '-----По всем продавцам для клиентов выбранного продавца
                    MySQLStr = MySQLStr & "AND (tbl_CRM_Events.CompanyID IN(SELECT CompanyID "
                    MySQLStr = MySQLStr & "FROM tbl_CRM_Events "
                    MySQLStr = MySQLStr & "WHERE (ActionPlannedDate >= CONVERT(DATETIME, '" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103)) "
                    MySQLStr = MySQLStr & "AND (ActionPlannedDate <= CONVERT(DATETIME, '" & Format(DateTimePicker2.Value, "dd/MM/yyyy") & "', 103)) "
                    MySQLStr = MySQLStr & "AND (UserID = " & ComboBox1.SelectedValue & ") "
                    MySQLStr = MySQLStr & "GROUP BY CompanyID "
                    MySQLStr = MySQLStr & "UNION "
                    MySQLStr = MySQLStr & "Select tbl_CRM_Companies.CompanyID "
                    MySQLStr = MySQLStr & "FROM tbl_CRM_Companies INNER JOIN "
                    MySQLStr = MySQLStr & "SL010300 ON tbl_CRM_Companies.ScalaCustomerCode = SL010300.SL01001 "
                    MySQLStr = MySQLStr & "WHERE (SL010300.SL01035 = N'" & Declarations.SalesmanCode & "'))) "
                End If
            End If
            If ComboBox3.Text = "Только активные действия" Then
                MySQLStr = MySQLStr & " AND (tbl_CRM_Events.ActionResultID IS NULL) "
            End If
            MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Events.ActionPlannedDate "

            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                '-----
                Dim MyStr As ActClass
                Dim MyList As List(Of ActClass)
                MyList = New List(Of ActClass)
                For i As Integer = 0 To MyDs.Tables(0).Rows.Count - 1
                    MyStr = New ActClass
                    MyStr.EventID = MyDs.Tables(0).Rows(i).Item("EventID").ToString()
                    MyStr.ActionPlannedDate = MyDs.Tables(0).Rows(i).Item("ActionPlannedDate")
                    MyStr.DirectionName = MyDs.Tables(0).Rows(i).Item("DirectionName")
                    MyStr.EventTypeName = MyDs.Tables(0).Rows(i).Item("EventTypeName")
                    MyStr.ScalaCustomerCode = MyDs.Tables(0).Rows(i).Item("ScalaCustomerCode")
                    MyStr.CompanyName = MyDs.Tables(0).Rows(i).Item("CompanyName")
                    MyStr.ContactName = MyDs.Tables(0).Rows(i).Item("ContactName")
                    MyStr.ContactPhone = MyDs.Tables(0).Rows(i).Item("ContactPhone")
                    MyStr.ContactEMail = MyDs.Tables(0).Rows(i).Item("ContactEMail")
                    MyStr.ActionName = MyDs.Tables(0).Rows(i).Item("ActionName")
                    MyStr.ProjectInfo = MyDs.Tables(0).Rows(i).Item("ProjectInfo")
                    MyStr.ActionSumm = MyDs.Tables(0).Rows(i).Item("ActionSumm")
                    MyStr.ActionComments = MyDs.Tables(0).Rows(i).Item("ActionComments")
                    MyStr.ActionResultName = MyDs.Tables(0).Rows(i).Item("ActionResultName")
                    MyStr.FullName = MyDs.Tables(0).Rows(i).Item("FullName")
                    MyStr.CompanyAddress = MyDs.Tables(0).Rows(i).Item("CompanyAddress")
                    MyStr.CompanyPhone = MyDs.Tables(0).Rows(i).Item("CompanyPhone")
                    MyStr.CompanyEMail = MyDs.Tables(0).Rows(i).Item("CompanyEMail")
                    MyStr.IsIKA = MyDs.Tables(0).Rows(i).Item("IsIKA")
                    MyStr.IsApproved = MyDs.Tables(0).Rows(i).Item("IsApproved")

                    MyList.Add(MyStr)
                Next
                '-----
                'SfDataGrid1.DataSource = MyDs.Tables(0)
                SfDataGrid1.Visible = False
                SfDataGrid1.DataSource = MyList
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End If
    End Sub

    Private Sub DateTimePicker1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.Validated
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Изменение начала диапазона выбора действий sfdatagrid1
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            CheckPlanesButton()
            CheckButtons()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub DateTimePicker2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker2.Validated
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Изменение конца диапазона выбора действий sfdatagrid1
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            CheckPlanesButton()
            CheckButtons()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранным продавцом sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            CheckPlanesButton()
            CheckButtons()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранной опцией - отображать данные для всех продавцов данных клиентов или нет sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            CheckPlanesButton()
            CheckButtons()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранной опцией - отображать все события или только активные sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            CheckPlanesButton()
            CheckButtons()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка данных в Excel из SfDataGrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim options = New ExcelExportingOptions()

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Files(*.Xls)|*.Xls|Files(*.Xlsx)|*.Xlsx"
        saveFileDialog.AddExtension = True
        saveFileDialog.DefaultExt = ".Xls"
        saveFileDialog.FileName = "Book1"
        If saveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK AndAlso saveFileDialog.CheckPathExists Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim excelEngine = SfDataGrid1.ExportToExcel(SfDataGrid1.View, options)
            Dim workBook = excelEngine.Excel.Workbooks(0)
            workBook.SaveAs(saveFileDialog.FileName)
            System.Windows.Forms.Cursor.Current = Cursors.Default
            If MessageBox.Show("Хотите открыть файл сейчас?", "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                Dim proc As New System.Diagnostics.Process()
                proc.StartInfo.FileName = saveFileDialog.FileName
                proc.Start()
            End If
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранной опцией - отображать все события или только активные sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Button5_Click_Func()
    End Sub

    Public Sub Button5_Click_Func()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция Загрузки данных в соответствии с выбранной опцией - отображать все события или только активные sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            CheckPlanesButton()
            CheckButtons()
            System.Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Включение или выключение возможности группировки sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid1.ShowGroupDropArea = True Then
            SfDataGrid1.GroupColumnDescriptions.Clear()
            SfDataGrid1.ShowGroupDropArea = False
            Button1.Text = "Включить группировку"
        Else
            SfDataGrid1.ShowGroupDropArea = True
            Button1.Text = "Выключить группировку"
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Быстрый поиск - изменение текста поиска sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        FilterText = TextBox1.Text
        SfDataGrid1.View.Filter = AddressOf FilterRecords
        SfDataGrid1.View.RefreshFilter()
        SfDataGrid1.SearchController.Search(FilterText)

    End Sub

    Public Function FilterRecords(ByVal o As Object) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Фильтрация таблицы в соответствии с текстом поиска sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim item = TryCast(o, ActClass)
        If item IsNot Nothing Then
            If item.ActionPlannedDate.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.DirectionName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.EventTypeName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ScalaCustomerCode.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ContactName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ContactPhone.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ContactEMail.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ActionName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ProjectInfo.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ActionSumm.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ActionComments.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ActionResultName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.FullName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyAddress.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyPhone.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyEMail.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.IsIKA.ToUpper().Contains(FilterText.ToUpper()) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создать действие в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim ActDate As DateTime
        Dim CompanyID As String
        Dim UserID As Integer
        Dim i As Integer

        CompanyID = ""
        UserID = ComboBox1.SelectedValue
        ActDate = New DateTime(Now.Year, Now.Month, Now.Day)
        Declarations.MyResult = 0
        CreateAction(ActDate, CompanyID, UserID)
        If Declarations.MyResult = 1 Then
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            i = 2
            For Each record In SfDataGrid1.View.Records
                If record.Data.EventID = Declarations.MyEventID Then
                    SfDataGrid1.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckPlanesButton()
            CheckButtons()
        End If
    End Sub

    Private Sub OnSfDataGrid2CreateClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создать действие в CRM из контекстного меню sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim provider = New CultureInfo("en-US")
        Dim ActDate As DateTime
        Dim CompanyID As String
        Dim UserID As Integer
        Dim SelRowIndex As Integer
        Dim SelColumnIndex As Integer

        If SfDataGrid2.CurrentCell.Column.ToString().Equals("CompanyScalaCode") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("CompanyName") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("CustProject") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("OrdersQTY") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("Status") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("RowTotal") Then
            ActDate = DateTime.ParseExact(Microsoft.VisualBasic.Strings.Right("00" + Now().Day, 2) _
                + "/" + Microsoft.VisualBasic.Strings.Right("00" + ComboBox5.SelectedValue.ToString(), 2) + "/" + ComboBox4.SelectedValue.ToString(), "dd/MM/yyyy", provider)
        Else
            ActDate = DateTime.ParseExact(Microsoft.VisualBasic.Strings.Right("00" + GetDayFromHeader(SfDataGrid2.CurrentCell.Column.MappingName).ToString(), 2) _
                + "/" + Microsoft.VisualBasic.Strings.Right("00" + ComboBox5.SelectedValue.ToString(), 2) + "/" + ComboBox4.SelectedValue.ToString(), "dd/MM/yyyy", provider)
        End If

        CompanyID = SfDataGrid2.SelectedItem.GetItem("CompanyID").ToString
        UserID = ComboBox6.SelectedValue

        SelRowIndex = SfDataGrid2.CurrentCell.RowIndex()
        SelColumnIndex = SfDataGrid2.CurrentCell.ColumnIndex

        Declarations.MyResult = 0
        CreateAction(ActDate, CompanyID, UserID)
        If Declarations.MyResult = 1 Then
            LoadScheduleList()
            SetScheduleTableParams(SfDataGrid2)
            CheckSfDataGrid2Activity()
            CheckPlanFactState()
            SfDataGrid2.Visible = True
            LoadScheduleListDetail()
            SfDataGrid2.MoveToCurrentCell(New RowColumnIndex(SelRowIndex, SelColumnIndex))
            CheckPlanesButtonSf2()
        End If
    End Sub

    Private Sub OnSfDataGrid7CreateClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню создание клиента sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyResult = 0
        CreateClient()
        If Declarations.MyResult = 1 Then
            LoadCustomersList()
            TextBox4.Text = ""
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
            i = 2
            For Each record In SfDataGrid7.View.Records
                If record.Data.CompanyID = Declarations.MyClientID Then
                    SfDataGrid7.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()
        End If
    End Sub

    Private Sub OnSfDataGrid4CreateClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню создание проекта sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyResult = 0
        CreateProject()
        If Declarations.MyResult = 1 Then
            LoadProjectsList()
            TextBox3.Text = ""
            SetInitProjectTableParams(SfDataGrid4)
            SetProjectTableParams(SfDataGrid4)
            SfDataGrid4.Visible = True
            i = 2
            For Each record In SfDataGrid4.View.Records
                If record.Data.ProjectID = Declarations.MyProjectID Then
                    SfDataGrid4.MoveToCurrentCell(New RowColumnIndex(i, 3))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckProjectButtons()
            CheckProjectApprovButton()
        End If
    End Sub

    Private Sub OnCreateClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создать действие в CRM из контекстного меню sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim ActDate As DateTime
        Dim CompanyID As String
        Dim UserID As Integer
        Dim i As Integer

        CompanyID = ""
        UserID = ComboBox1.SelectedValue
        ActDate = New DateTime(Now.Year, Now.Month, Now.Day)
        Declarations.MyResult = 0
        CreateAction(ActDate, CompanyID, UserID)
        If Declarations.MyResult = 1 Then
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            i = 2
            For Each record In SfDataGrid1.View.Records
                If record.Data.EventID = Declarations.MyEventID Then
                    SfDataGrid1.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckPlanesButton()
            CheckButtons()
        End If
    End Sub

    Private Sub CreateAction(ByVal MyDate As DateTime, ByVal MyCompany As String, ByVal UserID As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание действия в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddEvent = New AddEvent
        MyAddEvent.StartParam = "Create"
        MyAddEvent.ActDate = MyDate
        MyAddEvent.CompanyID = MyCompany
        MyAddEvent.UserID = UserID
        Declarations.MyEventID = "00000000-0000-0000-0000-000000000000"
        MyAddEvent.ShowDialog()
    End Sub

    Private Sub OnSfDataGrid2EditDayClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Переход к редактированию выбранного дня выбранной компании
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDateStart As DateTime
        Dim MyDateFin As DateTime
        Dim MyScalaCode As String
        Dim MyCompanyName As String
        Dim MyDayNum As Integer
        'Dim MyCol As Integer

        Tab1Settings = 0
        Dim provider = New CultureInfo("en-US")
        MyDateStart = DateTime.ParseExact("01/" + Microsoft.VisualBasic.Strings.Right("00" + ComboBox5.SelectedValue.ToString(), 2) + "/" + ComboBox4.SelectedValue.ToString(), "dd/MM/yyyy", provider)
        MyDateFin = MyDateStart.AddMonths(1)
        MyDateFin = MyDateFin.AddDays(-1)
        'MyCol = SfDataGrid2.CurrentCell.ColumnIndex
        MyDayNum = GetDayFromHeader(SfDataGrid2.CurrentCell.Column.MappingName) - 1

        LoadDataToTabActList(MyDateStart, MyDateFin, ComboBox6.SelectedValue, "Всех продавцов клиентов", "Все действия")

        SfDataGrid1.ClearFilters()
        '-----Фильтр по компании
        MyScalaCode = SfDataGrid2.SelectedItem.GetItem("CompanyScalaCode")
        SfDataGrid1.Columns("ScalaCustomerCode").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyScalaCode})
        MyCompanyName = SfDataGrid2.SelectedItem.GetItem("CompanyName")
        SfDataGrid1.Columns("CompanyName").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyCompanyName})
        '-----Фильтр по дате
        SfDataGrid1.Columns("ActionPlannedDate").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyDateStart.AddDays(MyDayNum)})
        '-----
        CheckPlanesButton()
        CheckButtons()
        TabControl1.SelectedIndex = 1
    End Sub

    Private Function GetDayFromHeader(ByVal MyStr As String) As Integer
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Получение номера дня из заголовка
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim WrkStr As String
        Dim MyNum As Integer

        WrkStr = ""
        For i As Integer = 0 To MyStr.Length - 1
            If MyStr(i).ToString().Equals("_") Then
                Exit For
            ElseIf MyStr(i).ToString().Equals("d") = False Then
                WrkStr = WrkStr + MyStr(i)
            End If
        Next
        Try
            MyNum = Integer.Parse(WrkStr)
        Catch ex As Exception
            GetDayFromHeader = 0
            Exit Function
        End Try
        GetDayFromHeader = MyNum
    End Function

    Private Sub OnSfDataGrid2EditMonthClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Переход к редактированию выбранного месяца выбранной компании
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDateStart As DateTime
        Dim MyDateFin As DateTime
        Dim MyScalaCode As String
        Dim MyCompanyName As String

        Tab1Settings = 0
        Dim provider = New CultureInfo("en-US")
        MyDateStart = DateTime.ParseExact("01/" + Microsoft.VisualBasic.Strings.Right("00" + ComboBox5.SelectedValue.ToString(), 2) + "/" + ComboBox4.SelectedValue.ToString(), "dd/MM/yyyy", provider)
        MyDateFin = MyDateStart.AddMonths(1)
        MyDateFin = MyDateFin.AddDays(-1)
        LoadDataToTabActList(MyDateStart, MyDateFin, ComboBox6.SelectedValue, "Всех продавцов клиентов", "Все действия")
        '-----Фильтр по компании
        SfDataGrid1.ClearFilters()
        MyScalaCode = SfDataGrid2.SelectedItem.GetItem("CompanyScalaCode")
        SfDataGrid1.Columns("ScalaCustomerCode").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyScalaCode})
        MyCompanyName = SfDataGrid2.SelectedItem.GetItem("CompanyName")
        SfDataGrid1.Columns("CompanyName").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyCompanyName})
        '-----
        CheckPlanesButton()
        CheckButtons()
        TabControl1.SelectedIndex = 1
    End Sub

    Private Sub OnSfDataGrid2MonthProjectsClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Переход к списку проектов активных в этом месяце
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDateStart As DateTime
        Dim MyDateFin As DateTime

        Tab2Settings = 0
        Dim provider = New CultureInfo("en-US")
        MyDateStart = DateTime.ParseExact("01/" + Microsoft.VisualBasic.Strings.Right("00" + ComboBox5.SelectedValue.ToString(), 2) + "/" + ComboBox4.SelectedValue.ToString(), "dd/MM/yyyy", provider)
        MyDateFin = MyDateStart.AddMonths(1)
        MyDateFin = MyDateFin.AddDays(-1)
        LoadDataToProjects(MyDateStart, MyDateFin)
        SfDataGrid4.ClearFilters()
        '-----
        CheckProjectButtons()
        CheckProjectApprovButton()
        TabControl1.SelectedIndex = 2
    End Sub

    Private Sub OnSfDataGrid2CompanyMonthProjectsClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Переход к списку проектов выбранной компании, активных в этом месяце
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyDateStart As DateTime
        Dim MyDateFin As DateTime
        Dim MyScalaCode As String
        Dim MyCompanyName As String

        Tab2Settings = 0
        Dim provider = New CultureInfo("en-US")
        MyDateStart = DateTime.ParseExact("01/" + Microsoft.VisualBasic.Strings.Right("00" + ComboBox5.SelectedValue.ToString(), 2) + "/" + ComboBox4.SelectedValue.ToString(), "dd/MM/yyyy", provider)
        MyDateFin = MyDateStart.AddMonths(1)
        MyDateFin = MyDateFin.AddDays(-1)
        LoadDataToProjects(MyDateStart, MyDateFin)
        '-----Фильтр по компании
        SfDataGrid4.ClearFilters()
        MyScalaCode = SfDataGrid2.SelectedItem.GetItem("CompanyScalaCode")
        SfDataGrid4.Columns("ScalaCustomerCode").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyScalaCode})
        MyCompanyName = SfDataGrid2.SelectedItem.GetItem("CompanyName")
        SfDataGrid4.Columns("CompanyName").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyCompanyName})
        '-----
        CheckProjectButtons()
        CheckProjectApprovButton()
        TabControl1.SelectedIndex = 2
    End Sub


    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Копировать действие в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyOldEventID = SfDataGrid1.SelectedItem.GetItem("EventID").ToString()
        Declarations.MyEventID = "00000000-0000-0000-0000-000000000000"
        Declarations.MyResult = 0
        CopyAction()
        If Declarations.MyResult = 1 Then
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            i = 2
            For Each record In SfDataGrid1.View.Records
                If record.Data.EventID = Declarations.MyEventID Then
                    SfDataGrid1.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckPlanesButton()
            CheckButtons()
        End If
    End Sub

    Private Sub OnCopyClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Копировать действие в CRM из контекстного меню
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyOldEventID = SfDataGrid1.SelectedItem.GetItem("EventID").ToString()
        Declarations.MyEventID = "00000000-0000-0000-0000-000000000000"
        Declarations.MyResult = 0
        CopyAction()
        If Declarations.MyResult = 1 Then
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            i = 2
            For Each record In SfDataGrid1.View.Records
                If record.Data.EventID = Declarations.MyEventID Then
                    SfDataGrid1.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckPlanesButton()
            CheckButtons()
        End If
    End Sub

    Private Sub CopyAction()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Копирование действия в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddEvent = New AddEvent
        MyAddEvent.StartParam = "Copy"
        MyAddEvent.ShowDialog()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удалить действие в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyEventID = SfDataGrid1.SelectedItem.GetItem("EventID").ToString()
        DeleteAction()
        LoadActionList()
        TextBox1.Text = ""
        SetTableParams(SfDataGrid1)
        SfDataGrid1.Visible = True
        CheckPlanesButton()
        CheckButtons()
    End Sub

    Private Sub OnSfDataGrid7DeleteClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню удаления клиента sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DeleteClient(SfDataGrid7.SelectedItem.GetItem("CompanyID").ToString()) = True Then
            LoadCustomersList()
            TextBox4.Text = ""
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()

            '-----параметры таблицы
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
        End If
    End Sub

    Private Sub OnSfDataGrid4DeleteClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню удаления проекта sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DeleteProject(SfDataGrid4.SelectedItem.GetItem("ProjectID").ToString()) = True Then
            LoadProjectsList()
            TextBox3.Text = ""
            CheckProjectButtons()
            CheckProjectApprovButton()

            '-----параметры таблицы
            SetInitProjectTableParams(SfDataGrid4)
            SetProjectTableParams(SfDataGrid4)
            SfDataGrid4.Visible = True
        End If
    End Sub

    Private Sub OnDeleteClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удалить действие в CRM из контекстного меню
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyEventID = SfDataGrid1.SelectedItem.GetItem("EventID").ToString()
        DeleteAction()
        LoadActionList()
        TextBox1.Text = ""
        SetTableParams(SfDataGrid1)
        SfDataGrid1.Visible = True
        CheckPlanesButton()
        CheckButtons()
    End Sub

    Private Sub DeleteAction()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление действия в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyRez As MsgBoxResult
        Dim MyRez1 As String
        Dim MyCalID As String
        Dim MyEmail As String

        MyCalID = ""
        MyEmail = ""

        MyRez = MsgBox("Вы уверены, что хотите удалить данное действие?", MsgBoxStyle.YesNo, "Внимание!")
        If MyRez = MsgBoxResult.Yes Then
            '---Удаление записи в календаре
            MySQLStr = "SELECT CalEventID "
            MySQLStr = MySQLStr & "FROM tbl_CRM_EventsInCalendar "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Trim(Declarations.MyEventID) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Else
                MyCalID = Declarations.MyRec.Fields("CalEventID").Value
                trycloseMyRec()
                MySQLStr = "SELECT RM.dbo.RM660100.RM66003 AS Email "
                MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers INNER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_Events ON ScalaSystemDB.dbo.ScaUsers.UserID = tbl_CRM_Events.UserID INNER JOIN "
                MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 INNER JOIN "
                MySQLStr = MySQLStr & "RM.dbo.RM660100 ON ST010300.ST01001 = RM.dbo.RM660100.RM66001 "
                MySQLStr = MySQLStr & "WHERE (tbl_CRM_Events.EventID = '" & Trim(Declarations.MyEventID) & "') "
                MySQLStr = MySQLStr & "AND (RM.dbo.RM660100.RM66003 <> '') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Else
                    MyEmail = Declarations.MyRec.Fields("Email").Value
                    trycloseMyRec()

                    '---Это от офиса 365
                    'Dim MyObj As New spbadm4.Esk365ServiceClient
                    'Dim MyCalendarEvent As New spbadm4.DeleteCalendarEventType

                    'MyCalendarEvent.CalendarEventIDOld = MyCalID
                    'MyCalendarEvent.Email = MyEmail
                    'MyCalendarEvent.Login = "Esk365ServiceUser"
                    'Try
                    '    MyRez1 = MyObj.DeleteCalendarEvent(MyCalendarEvent)
                    'Catch ex As Exception
                    'End Try

                    '---Это от MS Exchange
                    'Dim MyObj As New spbadm4_EWS.EskEWSServiceClient
                    'Dim MyCalendarEvent As New spbadm4_EWS.DeleteCalendarEventType

                    'MyCalendarEvent.CalendarEventIDOld = MyCalID
                    'MyCalendarEvent.Email = MyEmail
                    'MyCalendarEvent.Login = "EskEWSServiceUser"
                    'Try
                    '    MyRez1 = MyObj.DeleteCalendarEvent(MyCalendarEvent)
                    'Catch ex As Exception
                    'End Try

                    '---Это от Zimbra
                    Dim MyObj As New CalendarZimbraService.CalendarZimbraServiceClient
                    Dim MyCalendarEvent As New CalendarZimbraService.DeleteCalendarEventType

                    MyCalendarEvent.CalendarEventIDOld = MyCalID
                    MyCalendarEvent.Email = MyEmail
                    MyCalendarEvent.Login = "CalZimbraServiceUser"
                    Try
                        MyRez1 = MyObj.DeleteCalendarEvent(MyCalendarEvent)
                    Catch ex As Exception
                    End Try

                    If MyCalID.Equals("") = False Then
                        MySQLStr = "DELETE FROM tbl_CRM_EventsInCalendar "
                        MySQLStr = MySQLStr & "WHERE (EventID = '" & Trim(Declarations.MyEventID) & "') "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                    End If
                End If
            End If

            '---Удаление аттачментов
            MySQLStr = "DELETE FROM tbl_CRM_Attachments "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Trim(Declarations.MyEventID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---Удаление записи о действии
            MySQLStr = "DELETE FROM tbl_CRM_Events "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Trim(Declarations.MyEventID) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактировать действие в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim ActionID As String
        Dim i As Integer

        ActionID = SfDataGrid1.SelectedItem.GetItem("EventID").ToString()
        Declarations.MyResult = 0
        EditAction(ActionID)
        If Declarations.MyResult = 1 Then
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            i = 2
            For Each record In SfDataGrid1.View.Records
                If record.Data.EventID = Declarations.MyEventID Then
                    SfDataGrid1.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckPlanesButton()
            CheckButtons()
        End If
    End Sub

    Private Sub OnSfDataGrid7UnionClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню объединения клиентов sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyClientID = SfDataGrid7.SelectedItem.GetItem("CompanyID").ToString()
        Declarations.MyResult = 0
        UnionClients()
        If Declarations.MyResult = 1 Then
            LoadCustomersList()
            TextBox4.Text = ""
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
            i = 2
            For Each record In SfDataGrid7.View.Records
                If record.Data.CompanyID = Declarations.MyNewClientID Then
                    SfDataGrid7.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()
        End If
    End Sub

    Private Sub OnSfDataGrid7AddInfoClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню редактирования доп информации клиента sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyClientID = SfDataGrid7.SelectedItem.GetItem("CompanyID").ToString()
        Declarations.MyResult = 0
        EditClientAddInfo()
        If Declarations.MyResult = 1 Then
            LoadCustomersList()
            TextBox4.Text = ""
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
            i = 2
            For Each record In SfDataGrid7.View.Records
                If record.Data.CompanyID = Declarations.MyClientID Then
                    SfDataGrid7.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()
        End If
    End Sub

    Private Sub OnSfDataGrid7EditClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню редактирование клиента sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyClientID = SfDataGrid7.SelectedItem.GetItem("CompanyID").ToString()
        Declarations.MyResult = 0
        EditClient()
        If Declarations.MyResult = 1 Then
            LoadCustomersList()
            TextBox4.Text = ""
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
            i = 2
            For Each record In SfDataGrid7.View.Records
                If record.Data.CompanyID = Declarations.MyClientID Then
                    SfDataGrid7.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()
        End If
    End Sub

    Private Sub OnSfDataGrid4EditClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню редактирование проекта sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyProjectID = SfDataGrid4.SelectedItem.GetItem("ProjectID").ToString()
        Declarations.MyResult = 0
        EditProject()
        If Declarations.MyResult = 1 Then
            LoadProjectsList()
            TextBox3.Text = ""
            SetInitProjectTableParams(SfDataGrid4)
            SetProjectTableParams(SfDataGrid4)
            SfDataGrid4.Visible = True
            i = 2
            For Each record In SfDataGrid4.View.Records
                If record.Data.ProjectID = Declarations.MyProjectID Then
                    SfDataGrid4.MoveToCurrentCell(New RowColumnIndex(i, 3))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckProjectButtons()
            CheckProjectApprovButton()
        End If
    End Sub

    Private Sub OnSfDataGrid4ApproveClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// контекстное меню утверждение проекта sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim i As Integer

        Declarations.MyProjectID = SfDataGrid4.SelectedItem.GetItem("ProjectID").ToString()
        ApproveProjects(Declarations.MyProjectID)
        LoadProjectsList()
        TextBox3.Text = ""
        SetInitProjectTableParams(SfDataGrid4)
        SetProjectTableParams(SfDataGrid4)
        SfDataGrid4.Visible = True
        i = 2
        For Each record In SfDataGrid4.View.Records
            If record.Data.ProjectID = Declarations.MyProjectID Then
                SfDataGrid4.MoveToCurrentCell(New RowColumnIndex(i, 3))
                Exit For
            End If
            i = i + 1
        Next record
        CheckProjectButtons()
        CheckProjectApprovButton()
    End Sub

    Private Sub OnEditClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактировать действие в CRM из контекстного меню
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim ActionID As String
        Dim i As Integer

        ActionID = SfDataGrid1.SelectedItem.GetItem("EventID").ToString()
        Declarations.MyResult = 0
        EditAction(ActionID)
        If Declarations.MyResult = 1 Then
            LoadActionList()
            TextBox1.Text = ""
            SetTableParams(SfDataGrid1)
            SfDataGrid1.Visible = True
            i = 2
            For Each record In SfDataGrid1.View.Records
                If record.Data.EventID = Declarations.MyEventID Then
                    SfDataGrid1.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckPlanesButton()
            CheckButtons()
        End If
    End Sub

    Private Sub EditAction(ByVal ActionID As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование действия в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddEvent = New AddEvent
        MyAddEvent.StartParam = "Edit"
        Declarations.MyEventID = ActionID
        MyAddEvent.ShowDialog()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Просмотреть действие в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyEventID = SfDataGrid1.SelectedItem.GetItem("EventID").ToString()
        ViewAction()
    End Sub

    Private Sub OnViewClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Просмотреть действие в CRM из контекстного меню
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyEventID = SfDataGrid1.SelectedItem.GetItem("EventID").ToString()
        ViewAction()
    End Sub

    Private Sub ViewAction()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Просмотр действия в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyViewEvent = New ViewEvent
        MyViewEvent.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Передать действие в CRM другому продавцу мз SfDataGrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyEventID As String

        If SfDataGrid1.SelectedItems.Count = 0 Then
            MyEventID = ""
        Else
            MyEventID = SfDataGrid1.SelectedItem.GetItem("EventID").ToString()
        End If
        TransferAction(ComboBox1.SelectedValue, MyEventID)
        Button5_Click_Func()
    End Sub

    Private Sub OnTransferClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Передать действие в CRM другому продавцу из контекстного меню SfDataGrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TransferAction(ComboBox1.SelectedValue, SfDataGrid1.SelectedItem.GetItem("EventID").ToString())
        Button5_Click_Func()
    End Sub

    Private Sub TransferAction(ByVal UserID As Integer, ByVal EventID As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Передача действия в CRM другому продавцу
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MySendReturnAction = New SendReturnAction
        MySendReturnAction.MyUserID = UserID
        MySendReturnAction.MyEventID = EventID
        MySendReturnAction.ShowDialog()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Утвердить планы в CRM SfDataGrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ApproveActions()
        Button5_Click_Func()
    End Sub

    Private Sub ApproveActions()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Утверждение планов в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyPlansApprovement = New PlansApprovement
        MyPlansApprovement.ShowDialog()
    End Sub

    Private Sub OnSfDataGrid6RemoveAllFiltersClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка всех фильтров sfdatagrid6
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SfDataGrid6.ClearFilters()
    End Sub

    Private Sub OnSfDataGrid5RemoveAllFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка всех фильтров sfdatagrid5
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SfDataGrid5.ClearFilters()
    End Sub

    Private Sub OnSfDataGrid3RemoveAllFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка всех фильтров sfdatagrid3
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SfDataGrid3.ClearFilters()
    End Sub

    Private Sub OnSfDataGrid2RemoveAllFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка всех фильтров sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----план - факт
        SfDataGrid2FilterBlockFlag = 1
        CheckBoxPlan.Checked = True
        CheckBoxFact.Checked = True
        SfDataGrid2FilterBlockFlag = 0
        SfDataGrid2.ClearFilters()
    End Sub

    Private Sub OnSfDataGrid7RemoveAllFiltersClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка всех фильтров sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SfDataGrid7.ClearFilters()
    End Sub

    Private Sub OnSfDataGrid4RemoveAllFiltersClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка всех фильтров sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SfDataGrid4.ClearFilters()
    End Sub

    Private Sub OnRemoveAllFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка всех фильтров sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SfDataGrid1.ClearFilters()
    End Sub

    Private Sub OnSfDataGrid6RemoveFiltersClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка фильтра в выбранной колонке sfdatagrid6
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCol As Integer

        MyCol = SfDataGrid6.CurrentCell.ColumnIndex
        SfDataGrid6.ClearFilter(SfDataGrid6.Columns(MyCol - 1))
    End Sub

    Private Sub OnSfDataGrid5RemoveFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка фильтра в выбранной колонке sfdatagrid5
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCol As Integer

        MyCol = SfDataGrid5.CurrentCell.ColumnIndex
        SfDataGrid5.ClearFilter(SfDataGrid5.Columns(MyCol - 1))
    End Sub

    Private Sub OnSfDataGrid3RemoveFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка фильтра в выбранной колонке sfdatagrid3
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCol As Integer

        MyCol = SfDataGrid3.CurrentCell.ColumnIndex
        SfDataGrid3.ClearFilter(SfDataGrid3.Columns(MyCol - 1))
    End Sub

    Private Sub OnSfDataGrid2RemoveFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка фильтра в выбранной колонке sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyColStr As String
        Dim MyCol As Integer

        MyCol = SfDataGrid2.CurrentCell.ColumnIndex
        MyColStr = SfDataGrid2.CurrentCell.Column.MappingName
        If MyColStr.Equals("Status") Then
            '-----план - факт
            SfDataGrid2FilterBlockFlag = 1
            CheckBoxPlan.Checked = True
            CheckBoxFact.Checked = True
            SfDataGrid2FilterBlockFlag = 0
            SfDataGrid2.ClearFilter(SfDataGrid2.Columns(MyCol - 1))
        Else
            SfDataGrid2.ClearFilter(SfDataGrid2.Columns(MyCol - 1))
        End If
    End Sub

    Private Sub OnSfDataGrid7RemoveFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка фильтра в выбранной колонке sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCol As String

        MyCol = SfDataGrid7.CurrentCell.Column.MappingName
        SfDataGrid7.ClearFilter(SfDataGrid7.Columns(MyCol))
    End Sub

    Private Sub OnSfDataGrid4RemoveFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка фильтра в выбранной колонке sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCol As String

        MyCol = SfDataGrid4.CurrentCell.Column.MappingName
        SfDataGrid4.ClearFilter(SfDataGrid4.Columns(MyCol))
    End Sub

    Private Sub OnRemoveFiltersClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка фильтра в выбранной колонке sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCol As String

        MyCol = SfDataGrid1.CurrentCell.Column.MappingName
        SfDataGrid1.ClearFilter(SfDataGrid1.Columns(MyCol))
    End Sub

    Private Sub OnSfDataGrid6FilterBySelectClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка фильтра по выбранному элементу sfdatagrid6
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyColStr As String
        Dim MyVal As Object

        MyColStr = SfDataGrid6.CurrentCell.Column.MappingName
        MyVal = SfDataGrid6.SelectedItem(MyColStr).ToString

        SfDataGrid6.ClearFilter(SfDataGrid6.CurrentCell.Column)
        SfDataGrid6.Columns(MyColStr).FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyVal})
    End Sub

    Private Sub OnSfDataGrid5FilterBySelectClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка фильтра по выбранному элементу sfdatagrid5
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyColStr As String
        Dim MyVal As Object

        MyColStr = SfDataGrid5.CurrentCell.Column.MappingName
        MyVal = SfDataGrid5.SelectedItem(MyColStr).ToString

        SfDataGrid5.ClearFilter(SfDataGrid5.CurrentCell.Column)
        SfDataGrid5.Columns(MyColStr).FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyVal})
    End Sub

    Private Sub OnSfDataGrid3FilterBySelectClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка фильтра по выбранному элементу sfdatagrid3
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyColStr As String
        Dim MyVal As Object

        MyColStr = SfDataGrid3.CurrentCell.Column.MappingName
        MyVal = SfDataGrid3.SelectedItem(MyColStr).ToString

        SfDataGrid3.ClearFilter(SfDataGrid3.CurrentCell.Column)
        SfDataGrid3.Columns(MyColStr).FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyVal})
    End Sub

    Private Sub OnSfDataGrid2FilterBySelectClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка фильтра по выбранному элементу sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyColStr As String
        Dim MyVal As String

        MyColStr = SfDataGrid2.CurrentCell.Column.MappingName
        If MyColStr.Equals("Status") Then
            '-----план или факт
            MyVal = SfDataGrid2.SelectedItem.GetItem(MyColStr).ToString
            SfDataGrid2FilterBlockFlag = 1
            If MyVal.Equals("План") Then
                CheckBoxPlan.Checked = True
                CheckBoxFact.Checked = False
            Else
                CheckBoxPlan.Checked = False
                CheckBoxFact.Checked = True
            End If
            SfDataGrid2FilterBlockFlag = 0
            SfDataGrid2.ClearFilter(SfDataGrid2.CurrentCell.Column)
            SfDataGrid2.Columns(MyColStr).FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyVal})
        Else
            MyVal = SfDataGrid2.SelectedItem.GetItem(MyColStr).ToString

            SfDataGrid2.ClearFilter(SfDataGrid2.CurrentCell.Column)
            SfDataGrid2.Columns(MyColStr).FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyVal})
        End If
    End Sub

    Private Sub OnSfDataGrid7FilterBySelectClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка фильтра по выбранному элементу sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCol As String
        Dim MyVal As Object

        MyCol = SfDataGrid7.CurrentCell.Column.MappingName
        MyVal = SfDataGrid7.SelectedItem.GetItem(MyCol)

        SfDataGrid7.ClearFilter(SfDataGrid7.CurrentCell.Column)
        SfDataGrid7.Columns(MyCol).FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyVal})
    End Sub

    Private Sub OnSfDataGrid4FilterBySelectClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка фильтра по выбранному элементу sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCol As String
        Dim MyVal As Object

        MyCol = SfDataGrid4.CurrentCell.Column.MappingName
        MyVal = SfDataGrid4.SelectedItem.GetItem(MyCol)

        SfDataGrid4.ClearFilter(SfDataGrid4.CurrentCell.Column)
        SfDataGrid4.Columns(MyCol).FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyVal})
    End Sub

    Private Sub OnSfDataGrid4FindParentClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// поиск родительского проекта в sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyVal As Object
        Dim i As Integer

        MyVal = SfDataGrid4.SelectedItem.GetItem("ParentProjectID")
        i = 2
        For Each record In SfDataGrid4.View.Records
            If UCase(record.Data.ProjectID) = UCase(MyVal) Then
                SfDataGrid4.MoveToCurrentCell(New RowColumnIndex(i, 3))
                Exit For
            End If
            i = i + 1
        Next record
        CheckProjectButtons()
        LoadProjectListDetail()
    End Sub

    Private Sub OnFilterBySelectClicked(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка фильтра по выбранному элементу sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyCol As String
        Dim MyVal As Object

        MyCol = SfDataGrid1.CurrentCell.Column.MappingName
        MyVal = SfDataGrid1.SelectedItem.GetItem(MyCol)

        SfDataGrid1.ClearFilter(SfDataGrid1.CurrentCell.Column)
        SfDataGrid1.Columns(MyCol).FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.Equals, .FilterValue = MyVal})
    End Sub

    Private Sub SfDataGrid2_ContextMenuOpening(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка доступности элементов контекстного меню sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Разделители
        SfDataGrid2.RecordContextMenu.Items(3).Enabled = False
        SfDataGrid2.RecordContextMenu.Items(6).Enabled = False

        If SfDataGrid2.CurrentCell.Column.ToString().Equals("CompanyScalaCode") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("CompanyName") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("CustProject") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("OrdersQTY") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("Status") Or _
            SfDataGrid2.CurrentCell.Column.ToString().Equals("RowTotal") Then
            SfDataGrid2.RecordContextMenu.Items(1).Visible = False
            SfDataGrid2.RecordContextMenu.Items(8).Visible = True
            SfDataGrid2.RecordContextMenu.Items(9).Visible = True
        Else
            SfDataGrid2.RecordContextMenu.Items(1).Visible = True
            SfDataGrid2.RecordContextMenu.Items(8).Visible = False
            SfDataGrid2.RecordContextMenu.Items(9).Visible = False
        End If
    End Sub

    Private Sub SfDataGrid7_ContextMenuOpening(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка доступности элементов контекстного меню sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid7.SelectedItem.ScalaCustomerCode = "" Then
            SfDataGrid7.RecordContextMenu.Items(1).Visible = True
            SfDataGrid7.RecordContextMenu.Items(2).Visible = True
            SfDataGrid7.RecordContextMenu.Items(3).Visible = True
            SfDataGrid7.RecordContextMenu.Items(4).Visible = True
            SfDataGrid7.RecordContextMenu.Items(5).Visible = True
        Else
            SfDataGrid7.RecordContextMenu.Items(1).Visible = False
            SfDataGrid7.RecordContextMenu.Items(2).Visible = False
            SfDataGrid7.RecordContextMenu.Items(3).Visible = False
            SfDataGrid7.RecordContextMenu.Items(4).Visible = False
            SfDataGrid7.RecordContextMenu.Items(5).Visible = False
        End If
    End Sub

    Private Sub SfDataGrid4_ContextMenuOpening(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка доступности элементов контекстного меню sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Разделители
        SfDataGrid4.RecordContextMenu.Items(3).Enabled = False
        SfDataGrid4.RecordContextMenu.Items(5).Enabled = False
        SfDataGrid4.RecordContextMenu.Items(9).Enabled = False

        If (Declarations.MyPDPermission = True) Then
            If SfDataGrid4.SelectedItem.IsApproved Then
                '-----проект утвержден
                SfDataGrid4.RecordContextMenu.Items(2).Visible = False
            Else
                '-----проект не утвержден
                SfDataGrid4.RecordContextMenu.Items(2).Visible = True
            End If
        Else
            SfDataGrid4.RecordContextMenu.Items(2).Visible = False
        End If

        If SfDataGrid4.SelectedItem.IsApproved Then
            '-----проект утвержден
            SfDataGrid4.RecordContextMenu.Items(3).Visible = False
            SfDataGrid4.RecordContextMenu.Items(4).Visible = False
        Else
            '-----проект не утвержден
            SfDataGrid4.RecordContextMenu.Items(3).Visible = True
            SfDataGrid4.RecordContextMenu.Items(4).Visible = True
        End If
        

        If SfDataGrid4.SelectedItem.ParentProjectID = "00000000-0000-0000-0000-000000000000" Then
            '-----нет родительского проекта
            SfDataGrid4.RecordContextMenu.Items(9).Visible = False
            SfDataGrid4.RecordContextMenu.Items(10).Visible = False
        Else
            '-----есть родительский проект
            If SfDataGrid4.GroupColumnDescriptions.Count = 0 Then
                '---группировок нет
                SfDataGrid4.RecordContextMenu.Items(9).Visible = True
                SfDataGrid4.RecordContextMenu.Items(10).Visible = True
            Else
                '---группировки есть
                SfDataGrid4.RecordContextMenu.Items(9).Visible = False
                SfDataGrid4.RecordContextMenu.Items(10).Visible = False
            End If

        End If
    End Sub

    Private Sub SfDataGrid1_ContextMenuOpening(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка доступности элементов контекстного меню sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '-----Разделители
        SfDataGrid1.RecordContextMenu.Items(5).Enabled = False
        SfDataGrid1.RecordContextMenu.Items(7).Enabled = False

        If SfDataGrid1.SelectedItem.ActionResultName.Equals("") Then
            '-----действие не закрыто
            If SfDataGrid1.SelectedItem.FullName.Equals(Declarations.FullName) Then
                '-----Действие принадлежит сотруднику
                SfDataGrid1.RecordContextMenu.Items(2).Visible = True
                If SfDataGrid1.SelectedItem.IsApproved Then
                    '-----Действие утверждено
                    SfDataGrid1.RecordContextMenu.Items(5).Visible = False
                    SfDataGrid1.RecordContextMenu.Items(6).Visible = False
                Else
                    '-----Действие не утверждено
                    SfDataGrid1.RecordContextMenu.Items(5).Visible = True
                    SfDataGrid1.RecordContextMenu.Items(6).Visible = True
                End If
            Else
                '-----Действие не принадлежит сотруднику
                If Declarations.AllowChangeUser = 1 And (Declarations.MyPermission = True Or Declarations.MyCCPermission = True) Then
                    '-----Сотрудник может менять действия другого сотрудника
                    SfDataGrid1.RecordContextMenu.Items(2).Visible = True
                    If SfDataGrid1.SelectedItem.IsApproved Then
                        '-----Действие утверждено
                        SfDataGrid1.RecordContextMenu.Items(5).Visible = False
                        SfDataGrid1.RecordContextMenu.Items(6).Visible = False
                    Else
                        '-----Действие не утверждено
                        SfDataGrid1.RecordContextMenu.Items(5).Visible = True
                        SfDataGrid1.RecordContextMenu.Items(6).Visible = True
                    End If
                Else
                    '-----Сотрудник не может менять действия другого сотрудника
                    SfDataGrid1.RecordContextMenu.Items(2).Visible = False
                    SfDataGrid1.RecordContextMenu.Items(5).Visible = False
                    SfDataGrid1.RecordContextMenu.Items(6).Visible = False
                End If
            End If
        Else
            '-----действие закрыто
            SfDataGrid1.RecordContextMenu.Items(2).Visible = False
            SfDataGrid1.RecordContextMenu.Items(4).Visible = False
            SfDataGrid1.RecordContextMenu.Items(5).Visible = False
            SfDataGrid1.RecordContextMenu.Items(6).Visible = False
        End If
    End Sub

    Private Sub SfDataGrid7_SelectionChanged()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Событие смены выбора элемента sfDataGrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckCustomerButtons()
    End Sub

    Private Sub SfDataGrid4_SelectionChanged()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Событие смены выбора элемента sfDataGrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckProjectButtons()
        LoadProjectListDetail()
    End Sub

    Private Sub SfDataGrid2_SelectionChanged()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Событие смены выбора элемента sfDataGrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadScheduleListDetail()
    End Sub

    Private Sub SfDataGrid1_SelectionChanged()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Событие смены выбора элемента sfDataGrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckButtons()
    End Sub

    Private Sub CheckCustomerDownloadButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния кнопки выгрузки клиентов sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Declarations.MyCCPermission _
            Or Declarations.MyPermission _
            Or Declarations.MyPDPermission Then
            Button31.Visible = True
        Else
            Button31.Visible = False
        End If
    End Sub

    Private Sub CheckCustomerButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния кнопок при изменении выбора sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid7.SelectedItem Is Nothing Then
            '-----Ничего не выбрано
            Button25.Enabled = False
            Button26.Enabled = False
            Button28.Enabled = False
            Button29.Enabled = False
        Else
            Button28.Enabled = True
            If SfDataGrid7.SelectedItem.ScalaCustomerCode = "" Then
                Button25.Enabled = True
                Button26.Enabled = True
                Button29.Enabled = True
            Else
                Button25.Enabled = False
                Button26.Enabled = False
                Button29.Enabled = False
            End If
        End If
    End Sub

    Private Sub CheckProjectButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния кнопок при изменении выбора sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid4.SelectedItem Is Nothing Then
            '-----Ничего не выбрано
            Button22.Enabled = True
            Button20.Enabled = False
            Button21.Enabled = False
        Else
            '-----Строка выбрана
            Button22.Enabled = True
            Button20.Enabled = True
            If SfDataGrid4.SelectedItem.IsApproved Then
                '-----проект утвержден
                Button21.Enabled = False
            Else
                '-----проект не утвержден
                Button21.Enabled = True
            End If
        End If
    End Sub

    Private Sub CheckButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния кнопок при изменении выбора sfdatagrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid1.SelectedItem Is Nothing Then
            '-----Ничего не выбрано
            Button12.Enabled = False
            Button6.Enabled = False
            Button2.Enabled = False
            Button10.Enabled = False
        Else
            '-----Строка выбрана
            Button12.Enabled = True
            Button10.Enabled = True
            If SfDataGrid1.SelectedItem.ActionResultName.Equals("") Then
                '-----действие не закрыто
                If SfDataGrid1.SelectedItem.FullName.Equals(Declarations.FullName) Then
                    '-----Действие принадлежит сотруднику
                    Button2.Enabled = True
                    If SfDataGrid1.SelectedItem.IsApproved Then
                        '-----Действие утверждено
                        Button6.Enabled = False
                    Else
                        '-----Действие не утверждено
                        Button6.Enabled = True
                    End If
                Else
                    '-----Действие не принадлежит сотруднику
                    If Declarations.AllowChangeUser = 1 And (Declarations.MyPermission = True Or Declarations.MyCCPermission = True) Then
                        '-----Сотрудник может менять действия другого сотрудника
                        Button2.Enabled = True
                        If SfDataGrid1.SelectedItem.IsApproved Then
                            '-----Действие утверждено
                            Button6.Enabled = False
                        Else
                            '-----Действие не утверждено
                            Button6.Enabled = True
                        End If
                    Else
                        '-----Сотрудник не может менять действия другого сотрудника
                        Button6.Enabled = False
                        Button2.Enabled = False
                    End If
                End If
            Else
                '-----действие закрыто
                Button12.Enabled = True
                Button6.Enabled = False
                Button2.Enabled = False
                Button10.Enabled = True
            End If
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных в закладке план на месяц sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Button9_Click_Func()
    End Sub

    Public Sub Button9_Click_Func()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция Обновления данных в закладке план на месяц sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadScheduleList()
        SetScheduleTableParams(SfDataGrid2)
        CheckSfDataGrid2Activity()
        CheckPlanFactState()
        SfDataGrid2.Visible = True
        LoadScheduleListDetail()
        CheckPlanesButtonSf2()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных в закладке план на месяц по изменению года sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            LoadScheduleList()
            SetScheduleTableParams(SfDataGrid2)
            CheckSfDataGrid2Activity()
            CheckPlanFactState()
            SfDataGrid2.Visible = True
            LoadScheduleListDetail()
            CheckPlanesButtonSf2()
        End If
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных в закладке план на месяц по изменению месяца sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            LoadScheduleList()
            SetScheduleTableParams(SfDataGrid2)
            CheckSfDataGrid2Activity()
            CheckPlanFactState()
            SfDataGrid2.Visible = True
            LoadScheduleListDetail()
            CheckPlanesButtonSf2()
        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных в закладке план на месяц по изменению продавца sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If LoadFlag = 0 Then
            LoadScheduleList()
            SetScheduleTableParams(SfDataGrid2)
            CheckSfDataGrid2Activity()
            CheckPlanFactState()
            SfDataGrid2.Visible = True
            LoadScheduleListDetail()
            CheckPlanesButtonSf2()
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка данных в Excel из SfDataGrid1
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim options = New ExcelExportingOptions()

        options.ExportBorders = True
        options.ExportColumnWidth = True
        options.ExportFreezePanes = True
        options.ExportGroupSummary = True
        options.ExportPageOptions = True
        options.ExportRowHeight = True
        options.ExportStyle = True
        options.AllowOutlining = True

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Files(*.Xls)|*.Xls|Files(*.Xlsx)|*.Xlsx"
        saveFileDialog.AddExtension = True
        saveFileDialog.DefaultExt = ".Xls"
        saveFileDialog.FileName = "Book1"
        If saveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK AndAlso saveFileDialog.CheckPathExists Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim excelEngine = SfDataGrid2.ExportToExcel(SfDataGrid2.View, options)
            Dim workBook = excelEngine.Excel.Workbooks(0)
            workBook.SaveAs(saveFileDialog.FileName)
            System.Windows.Forms.Cursor.Current = Cursors.Default
            If MessageBox.Show("Хотите открыть файл сейчас?", "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                Dim proc As New System.Diagnostics.Process()
                proc.StartInfo.FileName = saveFileDialog.FileName
                proc.Start()
            End If
        End If

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Быстрый поиск - изменение текста поиска в sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        FilterText = TextBox2.Text
        SfDataGrid2.View.Filter = AddressOf FilterRecords2
        SfDataGrid2.View.RefreshFilter()
        SfDataGrid2.SearchController.Search(FilterText)
    End Sub

    Public Function FilterRecords2(ByVal o As Object) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Фильтрация таблицы в соответствии с текстом поиска 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If IsNothing(o) = False Then
            If o.CompanyName.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                o.CompanyScalaCode.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                o.CustProject.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                o.OrdersQTY.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                o.Status.ToString().ToUpper().Contains(FilterText.ToUpper()) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub SfDataGrid2_ToolTipOpening(ByVal sender As Object, ByVal e As ToolTipOpeningEventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Вывод кастомного tooltip для sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.DisplayText = "Scala" Then
            e.ToolTipInfo.Items(0).Text = "Для данного месяца " + Chr(13) + Chr(10) +
                "Выводится количество: " + Chr(13) + Chr(10) +
                "- Заказов всех типов, " + Chr(13) + Chr(10) +
                "- Коммерческих предожений, " + Chr(13) + Chr(10) +
                "- СФ, по которым выходит срок оплаты "
            e.ToolTipInfo.Items(0).Style.TextAlignment = ContentAlignment.MiddleLeft
        End If
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Показывать все записи или только те, по которым есть планы / факты
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ShowOnlyActive = 0 Then
            SfDataGrid2.ClearFilter(SfDataGrid2.Columns("RowTotal"))
            SfDataGrid2.Columns("RowTotal").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.NotEquals, .FilterValue = 0})
            Button15.Text = "Показать все записи"
            ShowOnlyActive = 1
        Else
            SfDataGrid2.ClearFilter(SfDataGrid2.Columns("RowTotal"))
            Button15.Text = "Показать только активности"
            ShowOnlyActive = 0
        End If
    End Sub

    Private Sub CheckSfDataGrid2Activity()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка флага Показывать все записи или только те, по которым есть планы / факты и выставление фильтра
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ShowOnlyActive = 0 Then
            SfDataGrid2.ClearFilter(SfDataGrid2.Columns("RowTotal"))

        Else
            SfDataGrid2.ClearFilter(SfDataGrid2.Columns("RowTotal"))
            SfDataGrid2.Columns("RowTotal").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.NotEquals, .FilterValue = 0})
        End If
    End Sub

    Private Sub CheckBoxFact_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBoxFact.CheckStateChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Изменение флага выводить факт или нет в sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckBoxFact.Checked = True Then
            ShowSfDataGri2Fact = 1
        Else
            ShowSfDataGri2Fact = 0
        End If
        CheckPlanFactState()
    End Sub

    Private Sub CheckBoxPlan_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBoxPlan.CheckStateChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Изменение флага выводить план или нет в sfdatagrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckBoxPlan.Checked = True Then
            ShowSfDataGri2Plan = 1
        Else
            ShowSfDataGri2Plan = 0
        End If
        CheckPlanFactState()
    End Sub

    Private Sub CheckPlanFactState()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка флагов вывода плана и факта и выставление фильтров
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid2FilterBlockFlag = 0 Then
            Try
                SfDataGrid2.ClearFilter(SfDataGrid2.Columns("Status"))
            Catch ex As Exception
            End Try

            If ShowSfDataGri2Plan = 1 Then
            Else
                Try
                    SfDataGrid2.Columns("Status").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.NotEquals, .FilterValue = "План"})
                Catch ex As Exception
                End Try
            End If

            If ShowSfDataGri2Fact = 1 Then
            Else
                Try
                    SfDataGrid2.Columns("Status").FilterPredicates.Add(New FilterPredicate() With {.FilterType = FilterType.NotEquals, .FilterValue = "Факт"})
                Catch ex As Exception
                End Try
            End If
        End If
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Показывать дополнительную информацию для планирования
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ShowSfDataGri2Details = 0 Then
            Button16.Text = "Скрыть детали"
            ShowSfDataGri2Details = 1
            SplitContainer1.Panel2Collapsed = False
        Else
            Button16.Text = "Показывать детали"
            ShowSfDataGri2Details = 0
            SplitContainer1.Panel2Collapsed = True
        End If
        LoadScheduleListDetail()

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Быстрый поиск - изменение текста поиска sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        FilterText = TextBox3.Text
        SfDataGrid4.View.Filter = AddressOf FilterRecordsSfDataGrid4
        SfDataGrid4.View.RefreshFilter()
        SfDataGrid4.SearchController.Search(FilterText)
    End Sub

    Public Function FilterRecordsSfDataGrid4(ByVal o As Object) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Фильтрация таблицы в соответствии с текстом поиска sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim item = TryCast(o, ProjectClass)
        If item IsNot Nothing Then
            If item.ScalaCustomerCode.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ProjectName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ProjectSumm.ToString.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ProjectComment.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.FirstDate.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.LastDate.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.StartDate.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CloseDate.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ProposalDate.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ProjectAddr.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.Investor.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.Contractor.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ResponciblePerson.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.ManufacturersList.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.Competitors.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.AdditionalExpencesPerCent.ToString().ToUpper().Contains(FilterText.ToUpper()) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Включение или выключение возможности группировки sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid4.ShowGroupDropArea = True Then
            SfDataGrid4.GroupColumnDescriptions.Clear()
            SfDataGrid4.ShowGroupDropArea = False
            Button24.Text = "Включить группировку"
        Else
            SfDataGrid4.ShowGroupDropArea = True
            Button24.Text = "Выключить группировку"
        End If
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// нажатие кнопки создание проекта sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyResult = 0
        CreateProject()
        If Declarations.MyResult = 1 Then
            LoadProjectsList()
            TextBox3.Text = ""
            SetInitProjectTableParams(SfDataGrid4)
            SetProjectTableParams(SfDataGrid4)
            SfDataGrid4.Visible = True
            i = 2
            For Each record In SfDataGrid4.View.Records
                If record.Data.ProjectID = Declarations.MyProjectID Then
                    SfDataGrid4.MoveToCurrentCell(New RowColumnIndex(i, 3))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckProjectButtons()
            CheckProjectApprovButton()
        End If
    End Sub

    Private Sub CreateProject()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// процедура создания проекта sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddProject = New AddProject
        MyAddProject.StartParam = "Create"
        MyAddProject.SourceForm = "MainForm"
        MyAddProject.ShowDialog()
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// нажатие кнопки редактирование проекта sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyProjectID = SfDataGrid4.SelectedItem.GetItem("ProjectID").ToString()
        Declarations.MyResult = 0
        EditProject()
        If Declarations.MyResult = 1 Then
            LoadProjectsList()
            TextBox3.Text = ""
            SetInitProjectTableParams(SfDataGrid4)
            SetProjectTableParams(SfDataGrid4)
            SfDataGrid4.Visible = True
            i = 2
            For Each record In SfDataGrid4.View.Records
                If record.Data.ProjectID = Declarations.MyProjectID Then
                    SfDataGrid4.MoveToCurrentCell(New RowColumnIndex(i, 3))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckProjectButtons()
            CheckProjectApprovButton()
        End If
    End Sub

    Private Sub EditProject()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// процедура редактирования проекта sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddProject = New AddProject
        MyAddProject.StartParam = "Edit"
        MyAddProject.SourceForm = "MainForm"
        MyAddProject.ShowDialog()
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// нажатие кнопки удаления проекта sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DeleteProject(SfDataGrid4.SelectedItem.GetItem("ProjectID").ToString()) = True Then
            LoadProjectsList()
            TextBox3.Text = ""
            CheckProjectButtons()
            CheckProjectApprovButton()

            '-----параметры таблицы
            SetInitProjectTableParams(SfDataGrid4)
            SetProjectTableParams(SfDataGrid4)
            SfDataGrid4.Visible = True
        End If
    End Sub

    Private Function DeleteProject(ByVal MyProjectID As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление проекта SfDataGrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---Проверка - можно ли удалять, может быть есть ссылки на него в CRM
        MySQLStr = "SELECT COUNT(ProjectID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            '---можно удалять
            trycloseMyRec()
            '---------Дополнительная проверка - можно ли удалять, может быть, есть ссылки на него в заказах на продажу.
            MySQLStr = "SELECT COUNT(OrderID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_SalesHdr_ProjectAddInfo "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & MyProjectID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                '---можно удалять
                trycloseMyRec()
                '---Удаление дополнительной информации по проекту
                MySQLStr = "DELETE FROM tbl_CRM_Project_Details "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & MyProjectID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---удаление групп продуктов в проекте
                MySQLStr = "DELETE FROM tbl_CRM_Project_ProdGroups "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & MyProjectID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---Удаление расширенной информации по проекту
                '---tbl_CRM_Projects_Ext
                MySQLStr = "DELETE FROM tbl_CRM_Projects_Ext "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & MyProjectID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---tbl_CRM_Projects_StagesHistory
                MySQLStr = "DELETE FROM tbl_CRM_Projects_StagesHistory "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & MyProjectID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---tbl_CRM_Projects_Stages
                MySQLStr = "DELETE FROM tbl_CRM_Projects_Stages "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & MyProjectID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---Удаление проекта
                MySQLStr = "DELETE FROM tbl_CRM_Projects "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & MyProjectID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
                DeleteProject = True
            Else
                trycloseMyRec()
                MsgBox("Данный проект нельзя удалять, так как на него есть ссылки в заказах на продажу.", MsgBoxStyle.Critical, "Внимание!")
                DeleteProject = False
            End If
        Else
            trycloseMyRec()
            MsgBox("Данный проект нельзя удалять, так как на него есть ссылки в таблице действий. Удалить такой проект можно только удалив сначала все действия с этим проектом.", MsgBoxStyle.Critical, "Внимание!")
            DeleteProject = False
        End If
    End Function

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление списка проектов (кнопка) sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProjectsList()
        TextBox3.Text = ""
        CheckProjectButtons()
        CheckProjectApprovButton()

        '-----параметры таблицы
        SetInitProjectTableParams(SfDataGrid4)
        SetProjectTableParams(SfDataGrid4)
        SfDataGrid4.Visible = True
    End Sub

    Private Sub DateTimePicker4_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker4.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор начальной даты диапазона проектов sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProjectsList()
        TextBox3.Text = ""
        CheckProjectButtons()
        CheckProjectApprovButton()

        '-----параметры таблицы
        SetInitProjectTableParams(SfDataGrid4)
        SetProjectTableParams(SfDataGrid4)
        SfDataGrid4.Visible = True
    End Sub

    Private Sub DateTimePicker3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker3.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор конечной даты диапазона проектов sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProjectsList()
        TextBox3.Text = ""
        CheckProjectButtons()
        CheckProjectApprovButton()

        '-----параметры таблицы
        SetInitProjectTableParams(SfDataGrid4)
        SetProjectTableParams(SfDataGrid4)
        SfDataGrid4.Visible = True
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка данных в Excel из SfDataGrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim options = New ExcelExportingOptions()

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Files(*.Xls)|*.Xls|Files(*.Xlsx)|*.Xlsx"
        saveFileDialog.AddExtension = True
        saveFileDialog.DefaultExt = ".Xls"
        saveFileDialog.FileName = "Book1"
        If saveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK AndAlso saveFileDialog.CheckPathExists Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim excelEngine = SfDataGrid4.ExportToExcel(SfDataGrid4.View, options)
            Dim workBook = excelEngine.Excel.Workbooks(0)
            workBook.SaveAs(saveFileDialog.FileName)
            System.Windows.Forms.Cursor.Current = Cursors.Default
            If MessageBox.Show("Хотите открыть файл сейчас?", "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                Dim proc As New System.Diagnostics.Process()
                proc.StartInfo.FileName = saveFileDialog.FileName
                proc.Start()
            End If
        End If
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// нажатие кнопки утверждения проектов sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyProjectID = SfDataGrid4.SelectedItem.GetItem("ProjectID").ToString()
        ApproveProjects(Declarations.MyProjectID)
        LoadProjectsList()
        TextBox3.Text = ""
        SetInitProjectTableParams(SfDataGrid4)
        SetProjectTableParams(SfDataGrid4)
        SfDataGrid4.Visible = True
        i = 2
        For Each record In SfDataGrid4.View.Records
            If record.Data.ProjectID = Declarations.MyProjectID Then
                SfDataGrid4.MoveToCurrentCell(New RowColumnIndex(i, 3))
                Exit For
            End If
            i = i + 1
        Next record
        CheckProjectButtons()
        CheckProjectApprovButton()
    End Sub

    Private Sub ApproveProjects(ByVal ProjectID As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// процедура утверждения проектов sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "UPDATE tbl_CRM_Projects "
        MySQLStr = MySQLStr & "SET IsApproved = 1 "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & ProjectID & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Показывать дополнительную информацию для планирования
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ShowSfDataGri4Details = 0 Then
            Button23.Text = "Скрыть детали"
            ShowSfDataGri4Details = 1
            SplitContainer2.Panel2Collapsed = False
        Else
            Button23.Text = "Показывать детали"
            ShowSfDataGri4Details = 0
            SplitContainer2.Panel2Collapsed = True
        End If
        LoadProjectListDetail()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Быстрый поиск - изменение текста поиска sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        FilterText = TextBox4.Text
        SfDataGrid7.View.Filter = AddressOf FilterRecordsSfDataGrid7
        SfDataGrid7.View.RefreshFilter()
        SfDataGrid7.SearchController.Search(FilterText)
    End Sub

    Public Function FilterRecordsSfDataGrid7(ByVal o As Object) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Фильтрация таблицы в соответствии с текстом поиска sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Dim item = TryCast(o, CustomerClass)
        If item IsNot Nothing Then
            If item.ScalaCustomerCode.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyName.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyAddress.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyAddress.ToString.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyPhone.ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CompanyEMail.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.CustomerGroup.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.EndMarket.ToString().ToUpper().Contains(FilterText.ToUpper()) Or _
                item.IsIKA.ToString().ToUpper().Contains(FilterText.ToUpper()) Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных по клиентам в sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadCustomersList()
        TextBox4.Text = ""
        CheckCustomerButtons()
        CheckCustomerDownloadButtons()

        '-----параметры таблицы
        SetInitCustomerTableParams(SfDataGrid7)
        SetCustomerTableParams(SfDataGrid7)
        SfDataGrid7.Visible = True
    End Sub

    Private Sub Button34_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button34.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Включение или выключение возможности группировки sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid7.ShowGroupDropArea = True Then
            SfDataGrid7.GroupColumnDescriptions.Clear()
            SfDataGrid7.ShowGroupDropArea = False
            Button34.Text = "Включить группировку"
        Else
            SfDataGrid7.ShowGroupDropArea = True
            Button34.Text = "Выключить группировку"
        End If
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// создание клиента sfdatagrid7 по кнопке
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyResult = 0
        CreateClient()
        If Declarations.MyResult = 1 Then
            LoadCustomersList()
            TextBox4.Text = ""
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
            i = 2
            For Each record In SfDataGrid7.View.Records
                If record.Data.CompanyID = Declarations.MyClientID Then
                    SfDataGrid7.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()
        End If
    End Sub

    Private Sub CreateClient()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция создания клиента sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddClient = New AddClient
        MyAddClient.StartParam = "Create"
        MyAddClient.SourceForm = "MainForm"
        MyAddClient.ShowDialog()
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование клиента sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyClientID = SfDataGrid7.SelectedItem.GetItem("CompanyID").ToString()
        Declarations.MyResult = 0
        EditClient()
        If Declarations.MyResult = 1 Then
            LoadCustomersList()
            TextBox4.Text = ""
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
            i = 2
            For Each record In SfDataGrid7.View.Records
                If record.Data.CompanyID = Declarations.MyClientID Then
                    SfDataGrid7.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()
        End If
    End Sub

    Private Sub EditClient()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция редактирования клиента sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddClient = New AddClient
        MyAddClient.StartParam = "Edit"
        MyAddClient.SourceForm = "MainForm"
        MyAddClient.ShowDialog()
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление клиента sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DeleteClient(SfDataGrid7.SelectedItem.GetItem("CompanyID").ToString()) = True Then
            LoadCustomersList()
            TextBox4.Text = ""
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()

            '-----параметры таблицы
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
        End If
    End Sub

    Private Function DeleteClient(ByVal MyCompanyID As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция удаления клиента sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---Проверка - можно ли удалять, может быть есть ссылки на него
        MySQLStr = "SELECT COUNT(CompanyID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (CompanyID = '" & MyCompanyID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            '---можно удалять
            trycloseMyRec()
            '---Удаление контактов
            MySQLStr = "DELETE FROM tbl_CRM_Contacts "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & MyCompanyID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '---Удаление дополнительной информации
            MySQLStr = "DELETE FROM tbl_CRM_Companies_Ext "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & MyCompanyID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '---Удаление самой записи
            MySQLStr = "DELETE FROM tbl_CRM_Companies "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & MyCompanyID & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            DeleteClient = True
        Else
            trycloseMyRec()
            MsgBox("Данную компанию нельзя удалять, так как на нее есть ссылки в таблице действий. Удалить такую компанию можно или удалив сначала все действия по этой компаниии, или объединив эту компанию с другой.", MsgBoxStyle.Critical, "Внимание!")
            DeleteClient = False
        End If

    End Function

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование доп информации по клиенту sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyClientID = SfDataGrid7.SelectedItem.GetItem("CompanyID").ToString()
        Declarations.MyResult = 0
        EditClientAddInfo()
        If Declarations.MyResult = 1 Then
            LoadCustomersList()
            TextBox4.Text = ""
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
            i = 2
            For Each record In SfDataGrid7.View.Records
                If record.Data.CompanyID = Declarations.MyClientID Then
                    SfDataGrid7.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()
        End If
    End Sub

    Private Sub EditClientAddInfo()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция редактирования доп информации по клиенту sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCustomerExtInfo = New CustomerExtInfo
        MyCustomerExtInfo.ShowDialog()
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Объединение клиентов sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim i As Integer

        Declarations.MyClientID = SfDataGrid7.SelectedItem.GetItem("CompanyID").ToString()
        Declarations.MyResult = 0
        UnionClients()
        If Declarations.MyResult = 1 Then
            LoadCustomersList()
            TextBox4.Text = ""
            SetInitCustomerTableParams(SfDataGrid7)
            SetCustomerTableParams(SfDataGrid7)
            SfDataGrid7.Visible = True
            i = 2
            For Each record In SfDataGrid7.View.Records
                If record.Data.CompanyID = Declarations.MyNewClientID Then
                    SfDataGrid7.MoveToCurrentCell(New RowColumnIndex(i, 2))
                    Exit For
                End If
                i = i + 1
            Next record
            CheckCustomerButtons()
            CheckCustomerDownloadButtons()
        End If
    End Sub

    Private Sub UnionClients()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Функция объединения клиентов sfdatagrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCustomerMerge = New CustomerMerge
        MyCustomerMerge.SrcForm = "MainForm"
        MyCustomerMerge.ShowDialog()
    End Sub

    Private Sub Button31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button31.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка данных в Excel из SfDataGrid7
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim options = New ExcelExportingOptions()

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "Files(*.Xls)|*.Xls|Files(*.Xlsx)|*.Xlsx"
        saveFileDialog.AddExtension = True
        saveFileDialog.DefaultExt = ".Xls"
        saveFileDialog.FileName = "Book1"
        If saveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK AndAlso saveFileDialog.CheckPathExists Then
            System.Windows.Forms.Cursor.Current = Cursors.WaitCursor
            Dim excelEngine = SfDataGrid7.ExportToExcel(SfDataGrid7.View, options)
            Dim workBook = excelEngine.Excel.Workbooks(0)
            workBook.SaveAs(saveFileDialog.FileName)
            System.Windows.Forms.Cursor.Current = Cursors.Default
            If MessageBox.Show("Хотите открыть файл сейчас?", "Export to Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                Dim proc As New System.Diagnostics.Process()
                proc.StartInfo.FileName = saveFileDialog.FileName
                proc.Start()
            End If
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Утверждение данных из SfDataGrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ApproveActions()
        Button9_Click_Func()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Передача активности другому продавцу из SfDataGrid2
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TransferAction(ComboBox6.SelectedValue, "")
        Button9_Click_Func()
    End Sub

    Private Sub SfDataGrid4_SelectionChanged1(ByVal sender As Object, ByVal e As Syncfusion.WinForms.DataGrid.Events.SelectionChangedEventArgs) Handles SfDataGrid4.SelectionChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена выбора проекта SfDataGrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckProjectApprovButton()
    End Sub
End Class
