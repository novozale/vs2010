Imports Syncfusion.Data
Imports Syncfusion.WinForms.DataGrid
Imports Syncfusion.WinForms.DataGrid.Events
Imports Syncfusion.WinForms.DataGrid.Enums


Public Class ProjectList
    Private propertyAccessProvider As IPropertyAccessProvider = Nothing
    Private autoFitOptions As New RowAutoFitOptions()
    Private autoHeight As Integer
    Private FilterText As String = ""
    Private LoadFlag = 0                            '---флаг - идет загрузка формы или нет

    Public Sub New()
        InitializeComponent()
        'Font = New Font(New FontFamily("Arial Unicode MS"), 8.0F, FontStyle.Bold)
        'Font = New Font(New FontFamily("Microsoft Sans Serif"), 8.0F, FontStyle.Bold)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ProjectList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape и ALT+F4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub ProjectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка окна и данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        AddHandler SfDataGrid4.QueryRowHeight, AddressOf sfDataGrid4_QueryRowHeight
        AddHandler SfDataGrid4.QueryRowStyle, AddressOf SfDataGrid4_QueryRowStyle
        AddHandler SfDataGrid4.SelectionChanged, AddressOf SfDataGrid4_SelectionChanged

        LoadDataToProjects(Nothing, Nothing)
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

    Private Sub SfDataGrid4_SelectionChanged()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Событие смены выбора элемента sfDataGrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckProjectButtons()
    End Sub

    Private Sub CheckProjectButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния кнопок при изменении выбора sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SfDataGrid4.SelectedItem Is Nothing Then
            '-----Ничего не выбрано
            Button4.Enabled = False
        Else
            '-----Строка выбрана
            Button4.Enabled = True
        End If
    End Sub

    Private Sub LoadDataToProjects(ByVal MyStartData As DateTime, ByVal MyFinData As DateTime)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных на закладку "Проекты"
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyCtrlDate As DateTime

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

        LoadProjectsList()
        TextBox3.Text = ""
        CheckProjectButtons()

        '-----параметры таблицы
        SetInitProjectTableParams(SfDataGrid4)
        SetProjectTableParams(SfDataGrid4)
        SfDataGrid4.Visible = True

        LoadFlag = 0
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
            For i As Integer = 2 To 23
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
        MySfDataGrid.RecordContextMenu.Items.Add("Снять все фильтры", Nothing, AddressOf OnSfDataGrid4RemoveAllFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Снять Фильтр в текущей колонке", Nothing, AddressOf OnSfDataGrid4RemoveFiltersClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("Фильтр по выбранному", Nothing, AddressOf OnSfDataGrid4FilterBySelectClicked)
        MySfDataGrid.RecordContextMenu.Items.Add("---------", Nothing, Nothing)
        MySfDataGrid.RecordContextMenu.Items.Add("Выбрать проект", Nothing, AddressOf OnSfDataGrid4SelectPProjectClicked)
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

    Private Sub OnSfDataGrid4SelectPProjectClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// выбор родительского проекта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyParentProjectID = SfDataGrid4.SelectedItem.GetItem("ProjectID").ToString()
        MyAddProject.TextBox12.Text = SfDataGrid4.SelectedItem.GetItem("ProjectName").ToString()

        Me.Close()
    End Sub

    Private Sub OnSfDataGrid4RemoveAllFiltersClicked()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка всех фильтров sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SfDataGrid4.ClearFilters()
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

    Private Sub SfDataGrid4_ContextMenuOpening(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Установка доступности элементов контекстного меню sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SfDataGrid4.RecordContextMenu.Items(0).Enabled = True
        SfDataGrid4.RecordContextMenu.Items(1).Enabled = True
        SfDataGrid4.RecordContextMenu.Items(2).Enabled = True
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

                MyList.Add(MyStr)
            Next
            '-----
            'SfDataGrid1.DataSource = MyDs.Tables(0)
            SfDataGrid4.Visible = False
            SfDataGrid4.DataSource = MyList
        Catch ex As Exception
        End Try

    End Sub

    Private Sub SetProjectTableParams(ByRef MySfDataGrid As SfDataGrid)
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление параметров элемента SfDataGrid4
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 2 To 23
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

    End Sub

    Private Sub DateTimePicker4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
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

        '-----параметры таблицы
        SetInitProjectTableParams(SfDataGrid4)
        SetProjectTableParams(SfDataGrid4)
        SfDataGrid4.Visible = True
    End Sub

    Private Sub DateTimePicker3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
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

        '-----параметры таблицы
        SetInitProjectTableParams(SfDataGrid4)
        SetProjectTableParams(SfDataGrid4)
        SfDataGrid4.Visible = True
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление списка проектов (кнопка) sfdatagrid4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        LoadProjectsList()
        TextBox3.Text = ""
        CheckProjectButtons()

        '-----параметры таблицы
        SetInitProjectTableParams(SfDataGrid4)
        SetProjectTableParams(SfDataGrid4)
        SfDataGrid4.Visible = True
    End Sub

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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором родительского проекта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyParentProjectID = SfDataGrid4.SelectedItem.GetItem("ProjectID").ToString()
        MyAddProject.TextBox12.Text = SfDataGrid4.SelectedItem.GetItem("ProjectName").ToString()

        Me.Close()
    End Sub
End Class