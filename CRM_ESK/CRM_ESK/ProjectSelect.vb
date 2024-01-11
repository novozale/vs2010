Public Class ProjectSelect
    Public StartParam As String

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ProjectSelect_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub ProjectSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список проектов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ComboBox1.SelectedItem = "Только незакрытые"

        If StartParam = "Edit" Then
            Label2.Text = MyAddEvent.TextBox6.Text
            Button5.Text = "Отмена"
            Button4.Enabled = True
            Button4.Visible = True
        Else
            'Label2.Text = Trim(MyCustomersWithProjects.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString())
            Button5.Text = "Выход"
            Button4.Enabled = False
            Button4.Visible = False
        End If
        LoadData()
        CheckButtons()
    End Sub

    Private Sub LoadData()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка проектов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT ProjectID, CompanyID, ProjectName, ProjectComment, StartDate, CloseDate, FirstDate, LastDate, ProjectAddr, Investor, Contractor, "
        MySQLStr = MySQLStr & "ResponciblePerson, ProposalDate, ManufacturersList, AlterManufacturers, Competitors, AdditionalExpencesPerCent, ISNULL(IsApproved,0) as IsApproved, CASE WHEN IsIPG = 0 THEN '' ELSE '+' END AS IsIPG "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Projects WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Declarations.MyClientID & "') "
        If ComboBox1.SelectedItem = "Только незакрытые" Then
            MySQLStr = MySQLStr & "AND (CloseDate IS NULL) "
        End If
        MySQLStr = MySQLStr & "ORDER BY StartDate "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "ID"
        DataGridView1.Columns(0).Width = 40
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "CID"
        DataGridView1.Columns(1).Width = 40
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).HeaderText = "Проект"
        DataGridView1.Columns(2).Width = 250
        DataGridView1.Columns(3).HeaderText = "Комментарий"
        DataGridView1.Columns(3).Width = 300
        DataGridView1.Columns(4).HeaderText = "Дата занесения в CRM"
        DataGridView1.Columns(4).Width = 120
        DataGridView1.Columns(5).HeaderText = "Дата закрытия проекта"
        DataGridView1.Columns(5).Width = 120
        DataGridView1.Columns(6).HeaderText = "Дата начала проекта"
        DataGridView1.Columns(6).Width = 120
        DataGridView1.Columns(7).HeaderText = "Дата окончания проекта"
        DataGridView1.Columns(7).Width = 120
        DataGridView1.Columns(8).HeaderText = "Адрес проекта"
        DataGridView1.Columns(8).Width = 120
        DataGridView1.Columns(9).HeaderText = "Инвестор"
        DataGridView1.Columns(9).Width = 120
        DataGridView1.Columns(10).HeaderText = "Контрактор"
        DataGridView1.Columns(10).Width = 120
        DataGridView1.Columns(11).HeaderText = "Ответственное лицо (инженер)"
        DataGridView1.Columns(11).Width = 120
        DataGridView1.Columns(12).HeaderText = "Дата подачи предложения"
        DataGridView1.Columns(12).Width = 80
        DataGridView1.Columns(13).HeaderText = "Список производителей"
        DataGridView1.Columns(13).Width = 120
        DataGridView1.Columns(14).HeaderText = "Альтернативные производители"
        DataGridView1.Columns(14).Width = 80
        DataGridView1.Columns(15).HeaderText = "Конкуренты"
        DataGridView1.Columns(15).Width = 120
        DataGridView1.Columns(16).HeaderText = "% доп расходов"
        DataGridView1.Columns(16).Width = 120
        DataGridView1.Columns(17).HeaderText = "Утвержден"
        DataGridView1.Columns(17).Width = 120
        DataGridView1.Columns(18).HeaderText = "Проект IPG"
        DataGridView1.Columns(18).Width = 40

    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление состояния кнопок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button4.Enabled = False
            Button8.Enabled = False
            Button9.Enabled = False
        Else
            If Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(5).Value.ToString()) = "" Then '---проект не закрыт
                Button8.Enabled = True
                If Me.DataGridView1.SelectedRows.Item(0).Cells(17).Value = True Then '---проект утвержден
                    If Declarations.MyPDPermission = True Then '--Директор по проектам
                        Button9.Enabled = True
                    Else
                        Button9.Enabled = False
                    End If
                Else
                    Button9.Enabled = True
                End If
                Button4.Enabled = True
            Else                                                                               '---проект закрыт
                Button8.Enabled = False
                Button9.Enabled = False
                Button4.Enabled = False
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        LoadData()
        CheckButtons()
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных в соответствии с выбранным значением
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        LoadData()
        CheckButtons()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание проекта
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyAddProject = New AddProject
        MyAddProject.StartParam = "Create"
        MyAddProject.SourceForm = "ProjectSelect"
        MyAddProject.ShowDialog()
        LoadData()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyProjectID Then
                DataGridView1.CurrentCell = DataGridView1.Item(2, i)
            End If
        Next
        '---проверка состояния кнопок
        CheckButtons()
    End Sub



    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором контакта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Edit" Then
            ProjectSelect()
        Else
            ProjectEdit()
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена выбора строки
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CheckButtons()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление проекта
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---Проверка - можно ли удалять, может быть есть ссылки на него в CRM
        MySQLStr = "SELECT COUNT(ProjectID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            '---можно удалять
            trycloseMyRec()
            '---------Дополнительная проверка - можно ли удалять, может быть, есть ссылки на него в заказах на продажу.
            MySQLStr = "SELECT COUNT(OrderID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_SalesHdr_ProjectAddInfo "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                '---можно удалять
                trycloseMyRec()
                '---Удаление дополнительной информации по проекту
                MySQLStr = "DELETE FROM tbl_CRM_Project_Details "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---удаление групп продуктов в проекте
                MySQLStr = "DELETE FROM tbl_CRM_Project_ProdGroups "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---Удаление расширенной информации по проекту
                '---tbl_CRM_Projects_Ext
                MySQLStr = "DELETE FROM tbl_CRM_Projects_Ext "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---tbl_CRM_Projects_StagesHistory
                MySQLStr = "DELETE FROM tbl_CRM_Projects_StagesHistory "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---tbl_CRM_Projects_Stages
                MySQLStr = "DELETE FROM tbl_CRM_Projects_Stages "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                '---Удаление проекта
                MySQLStr = "DELETE FROM tbl_CRM_Projects "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)

                LoadData()
                CheckButtons()
            Else
                trycloseMyRec()
                MsgBox("Данный проект нельзя удалять, так как на него есть ссылки в заказах на продажу.", MsgBoxStyle.Critical, "Внимание!")
            End If
        Else
            trycloseMyRec()
            MsgBox("Данный проект нельзя удалять, так как на него есть ссылки в таблице действий. Удалить такой проект можно только удалив сначала все действия с этим проектом.", MsgBoxStyle.Critical, "Внимание!")
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором контакта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ProjectSelect()
    End Sub

    Private Sub ProjectSelect()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором контакта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(5).Value.ToString()) = "" Then
            MyAddEvent.TextBox8.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString()) + " " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString())
            Declarations.MyProjectID = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())

            Me.Close()
        End If
    End Sub

    Private Sub ProjectEdit()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие проекта на редактирование
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyProjectID = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MyAddProject = New AddProject
        MyAddProject.StartParam = "Edit"
        MyAddProject.SourceForm = "ProjectSelect"
        MyAddProject.ShowDialog()
        LoadData()
        '---текущей строкой сделать отредактированную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyProjectID Then
                DataGridView1.CurrentCell = DataGridView1.Item(2, i)
            End If
        Next
        '---проверка состояния кнопок
        CheckButtons()
    End Sub

    Private Sub Button8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование проекта
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        ProjectEdit()
    End Sub
End Class