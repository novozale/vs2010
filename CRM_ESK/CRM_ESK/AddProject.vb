Imports ADODB
Imports System.IO
Imports System.Xml

Public Class AddProject
    Public StartParam As String
    Public SourceForm As String

    Private Sub AddProject_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Declarations.MyResult = 0
            Me.Close()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения информации
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Declarations.MyResult = 0
        Me.Close()
    End Sub

    Private Sub AddProject_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в форму
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyGUID As Guid
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка направлений
        Dim MyDs As New DataSet

        '-----список стадий проекта
        MySQLStr = "SELECT ID, Name "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Projects_StagesCFG WITH(NOLOCK) "
        MySQLStr = MySQLStr & "ORDER BY ID "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox2.DisplayMember = "Name" 'Это то что будет отображаться
            ComboBox2.ValueMember = "ID"   'это то что будет храниться
            ComboBox2.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----В зависимости от того - это новый проект или редактирование - читаем дополнительные данные и выставляем значения
        If StartParam = "Create" Then
            '---создаем запись о проекте
            MyGUID = Guid.NewGuid
            Declarations.MyProjectID = MyGUID.ToString
            Declarations.MyParentProjectID = "00000000-0000-0000-0000-000000000000"
            CheckBox1.Checked = False
            If SourceForm = "ProjectSelect" Then
                '---находим и выставляем компанию
                MySQLStr = "Select CompanyName "
                MySQLStr = MySQLStr & "FROM tbl_CRM_Companies "
                MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(Declarations.MyClientID) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Else
                    TextBox13.Text = Declarations.MyRec.Fields("CompanyName").Value
                    Button5.Enabled = False
                End If
            Else
                '---MainForm - не делаем ничего
                Button5.Enabled = True
            End If
            TextBox7.Text = Declarations.FullName
            '---Загружена детальная информация или нет
            CheckBox5.Checked = False
            Button11.Enabled = False
        Else    '-----редактирование
            MySQLStr = "SELECT tbl_CRM_Projects.ProjectID, tbl_CRM_Projects.CompanyID, tbl_CRM_Projects.ProjectName, tbl_CRM_Projects.ProjectSumm, tbl_CRM_Projects.ProjectComment, tbl_CRM_Projects.StartDate, "
            MySQLStr = MySQLStr & "tbl_CRM_Projects.CloseDate, tbl_CRM_Projects.FirstDate, tbl_CRM_Projects.LastDate, ISNULL(tbl_CRM_Projects.ProjectAddr, '') AS ProjectAddr, ISNULL(tbl_CRM_Projects.Investor, '') AS Investor, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects.Contractor, '') AS Contractor, ISNULL(tbl_CRM_Projects.ResponciblePerson, '') AS ResponciblePerson, ISNULL(tbl_CRM_Projects.ProposalDate, CONVERT(datetime, "
            MySQLStr = MySQLStr & "'01/01/1900', 103)) AS ProposalDate, ISNULL(tbl_CRM_Projects.ManufacturersList, '') AS ManufacturersList, ISNULL(tbl_CRM_Projects.AlterManufacturers, 0) AS AlterManufacturers, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects.Competitors, '') AS Competitors, ISNULL(tbl_CRM_Projects.AdditionalExpencesPerCent, 0) AS AdditionalExpencesPerCent, ISNULL(tbl_CRM_Projects.IsApproved, 0) "
            MySQLStr = MySQLStr & "AS IsApproved, ISNULL(tbl_CRM_Projects.IsIPG, 0) AS IsIPG, ISNULL(tbl_CRM_Projects_Ext.ParentProjectID, '00000000-0000-0000-0000-000000000000') AS ParentProjectID, ISNULL(tbl_CRM_Projects_1.ProjectName, "
            MySQLStr = MySQLStr & "'') AS ParentProjectName, ISNULL(tbl_CRM_Projects_Stages.ProjectStageID, CASE WHEN tbl_CRM_Projects.CloseDate IS NULL THEN 1 ELSE 100 END) AS ProjectStageID, "
            MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyName "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Companies INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Projects WITH (NOLOCK) ON tbl_CRM_Companies.CompanyID = tbl_CRM_Projects.CompanyID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Projects_Stages WITH (NOLOCK) ON tbl_CRM_Projects.ProjectID = tbl_CRM_Projects_Stages.ProjectID LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Projects AS tbl_CRM_Projects_1 WITH (NOLOCK) RIGHT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Projects_Ext WITH (NOLOCK) ON tbl_CRM_Projects_1.ProjectID = tbl_CRM_Projects_Ext.ParentProjectID ON tbl_CRM_Projects.ProjectID = tbl_CRM_Projects_Ext.ProjectID "
            MySQLStr = MySQLStr & "WHERE (tbl_CRM_Projects.ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Данный проект не существует, возможно, удален другим пользователем. Обновите данные в окне со списком проектов.", MsgBoxStyle.Critical, "Внимание!")
                Me.Close()
            Else
                Declarations.MyRec.MoveFirst()
                TextBox13.Text = Declarations.MyRec.Fields("CompanyName").Value
                Declarations.CompanyID = Declarations.MyRec.Fields("CompanyID").Value
                Declarations.MyClientID = Declarations.MyRec.Fields("CompanyID").Value
                TextBox6.Text = Declarations.MyRec.Fields("ProjectName").Value
                TextBox2.Text = Format(Declarations.MyRec.Fields("ProjectSumm").Value, ".00")
                TextBox1.Text = Declarations.MyRec.Fields("ProjectComment").Value
                DateTimePicker1.Value = Declarations.MyRec.Fields("FirstDate").Value
                DateTimePicker2.Value = Declarations.MyRec.Fields("LastDate").Value
                TextBox3.Text = Declarations.MyRec.Fields("ProjectAddr").Value
                TextBox4.Text = Declarations.MyRec.Fields("Investor").Value
                TextBox5.Text = Declarations.MyRec.Fields("Contractor").Value
                TextBox7.Text = Declarations.MyRec.Fields("ResponciblePerson").Value
                DateTimePicker3.Value = Declarations.MyRec.Fields("ProposalDate").Value
                TextBox8.Text = Declarations.MyRec.Fields("ManufacturersList").Value
                If Declarations.MyRec.Fields("AlterManufacturers").Value = -1 Then
                    CheckBox2.Checked = True
                Else
                    CheckBox2.Checked = False
                End If
                TextBox9.Text = Declarations.MyRec.Fields("Competitors").Value
                If Declarations.MyRec.Fields("AdditionalExpencesPerCent").Value = 0 Then
                    TextBox10.Text = ""
                Else
                    TextBox10.Text = Format(Declarations.MyRec.Fields("AdditionalExpencesPerCent").Value, ".00")
                End If

                If Declarations.MyRec.Fields("IsApproved").Value = -1 Then
                    CheckBox3.Checked = True
                Else
                    CheckBox3.Checked = False
                End If

                If Declarations.MyRec.Fields("IsIPG").Value = -1 Then
                    CheckBox4.Checked = True
                Else
                    CheckBox4.Checked = False
                End If
                If Declarations.MyRec.Fields("CloseDate").Value.ToString = "" Then
                    CheckBox1.Checked = False
                Else
                    CheckBox1.Checked = True
                End If
                Declarations.MyParentProjectID = Declarations.MyRec.Fields("ParentProjectID").Value
                TextBox12.Text = Declarations.MyRec.Fields("ParentProjectName").Value
                ComboBox2.SelectedValue = Declarations.MyRec.Fields("ProjectStageID").Value
                trycloseMyRec()
                '---составная строка - группы товаров в проекте
                MySQLStr = "SELECT distinct (SELECT tbl_CRM_ProdGroupsList.ItemGroupName + ';' AS 'data()' "
                MySQLStr = MySQLStr & "FROM tbl_CRM_Project_ProdGroups AS t2 INNER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_ProdGroupsList ON t2.ProdGroupID = tbl_CRM_ProdGroupsList.ID "
                MySQLStr = MySQLStr & "WHERE (t2.ProjectID = t1.ProjectID) ORDER BY tbl_CRM_ProdGroupsList.ID For xml path('')) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_CRM_Project_ProdGroups AS t1 INNER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_ProdGroupsList AS tbl_CRM_ProdGroupsList_1 ON t1.ProdGroupID = tbl_CRM_ProdGroupsList_1.ID "
                MySQLStr = MySQLStr & "WHERE (t1.ProjectID = '" & Declarations.MyProjectID & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                    TextBox11.Text = ""
                Else
                    Declarations.MyRec.MoveFirst()
                    TextBox11.Text = Declarations.MyRec.Fields("CC").Value
                    trycloseMyRec()
                End If
            End If

        End If
        '---Аттачменты
        LoadAttachments()
        CheckAttachmentsButtons()
        TextBox6.Select()

        '---Загружена детальная информация или нет
        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Details "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            CheckBox5.Checked = False
        Else
            If Declarations.MyRec.Fields("CC").Value > 0 Then
                trycloseMyRec()
                CheckBox5.Checked = True
            Else
                trycloseMyRec()
                CheckBox5.Checked = False
            End If
        End If
        If CheckBox5.Checked = True Then
            Button11.Enabled = True
        Else
            Button11.Enabled = False
        End If

        '-----Проверяем права пользователя, и в зависимости от них блокируем отдельные поля.
        CheckRights()

    End Sub

    Private Function LoadAttachments()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// загрузка данных об аттачментах в форму
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyAdapter4 As SqlClient.SqlDataAdapter    'для списка аттачментов
        Dim MyDs4 As New DataSet
        Dim MySQLStr As String

        MySQLStr = "SELECT AttachmentID, ProjectID, AttachmentName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Attachments WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
        MySQLStr = MySQLStr & "ORDER BY AttachmentName "
        Try
            MyAdapter4 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter4.SelectCommand.CommandTimeout = 600
            MyAdapter4.Fill(MyDs4)
            DataGridView1.DataSource = MyDs4.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "ID"
        DataGridView1.Columns(0).Width = 0
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "ProjID"
        DataGridView1.Columns(1).Width = 0
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).HeaderText = "Имя файла"
        DataGridView1.Columns(2).Width = 600
    End Function

    Public Function CheckAttachmentsButtons()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление состояния кнопок, ответственных за аттачменты
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button7.Enabled = False
            Button8.Enabled = False
        Else
            Button7.Enabled = True
            Button8.Enabled = True
        End If
    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна выбора клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////
        MyCustomerSelect = New CustomerSelect
        MyCustomerSelect.SourceForm = "AddProject"
        MyCustomerSelect.ShowDialog()
        If TextBox13.Text <> "" Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub CheckRights()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка и выставление состояния кнопок и полей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Declarations.MyPDPermission = True Then    '---директор по проектам - все поля
            If SourceForm = "ProjectSelect" Then
                Button5.Enabled = False
            Else
                If StartParam = "Edit" Then
                    '---Проверка - есть ссылки на проект в CRM
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
                            Button5.Enabled = True
                        Else
                            Button5.Enabled = False
                        End If
                    Else
                        Button5.Enabled = False
                    End If
                Else
                    Button5.Enabled = True
                End If
            End If
            TextBox6.Enabled = True
            TextBox2.Enabled = True
            TextBox1.Enabled = True
            DateTimePicker1.Enabled = True
            DateTimePicker2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            DateTimePicker3.Enabled = True
            TextBox8.Enabled = True
            CheckBox2.Enabled = True
            TextBox9.Enabled = True
            TextBox10.Enabled = True
            CheckBox3.Enabled = True
            CheckBox4.Enabled = True
            If StartParam = "Create" Then
                CheckBox1.Enabled = False
                Button3.Enabled = False
            Else
                CheckBox1.Enabled = True
                Button3.Enabled = True
            End If
            Button4.Enabled = True
            Button9.Enabled = True
            Button10.Enabled = True
        Else
            If CheckBox3.Checked = True Then  '---Утвержден - ничего нельзя менять
                Button5.Enabled = False
                TextBox6.Enabled = False
                TextBox2.Enabled = False
                TextBox1.Enabled = False
                DateTimePicker1.Enabled = False
                DateTimePicker2.Enabled = False
                TextBox3.Enabled = False
                TextBox4.Enabled = False
                TextBox5.Enabled = False
                DateTimePicker3.Enabled = False
                TextBox8.Enabled = False
                CheckBox2.Enabled = False
                TextBox9.Enabled = False
                TextBox10.Enabled = False
                CheckBox3.Enabled = False
                CheckBox4.Enabled = False
                CheckBox1.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
                Button9.Enabled = False
                Button10.Enabled = False
                If CheckBox1.Checked = True Then
                    ComboBox2.Enabled = False
                Else
                    ComboBox2.Enabled = True
                End If
            Else   '---не утвержден - закрываются только поля директора по проектам.
                If SourceForm = "ProjectSelect" Then
                    Button5.Enabled = False
                Else
                    If StartParam = "Edit" Then
                        '---Проверка - есть ссылки на проект в CRM
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
                                Button5.Enabled = True
                            Else
                                Button5.Enabled = False
                            End If
                        Else
                            Button5.Enabled = False
                        End If
                    Else
                        Button5.Enabled = True
                    End If
                End If
                TextBox6.Enabled = True
                TextBox2.Enabled = True
                TextBox1.Enabled = True
                DateTimePicker1.Enabled = True
                DateTimePicker2.Enabled = True
                TextBox3.Enabled = True
                TextBox4.Enabled = True
                TextBox5.Enabled = True
                DateTimePicker3.Enabled = True
                TextBox8.Enabled = True
                CheckBox2.Enabled = True
                TextBox9.Enabled = True
                TextBox10.Enabled = False
                CheckBox3.Enabled = False
                CheckBox4.Enabled = True
                If StartParam = "Create" Then
                    CheckBox1.Enabled = False
                Else
                    CheckBox1.Enabled = True
                End If
                Button3.Enabled = False
                Button4.Enabled = True
                Button9.Enabled = True
                Button10.Enabled = True
                If CheckBox1.Checked = True Then
                    ComboBox2.Enabled = False
                Else
                    ComboBox2.Enabled = True
                End If
            End If
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна выбора проекта 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyProjectList = New ProjectList
        MyProjectList.ShowDialog()
        Me.SelectNextControl(sender, True, True, True, False)
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка поля проекта 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TextBox12.Text = ""
        Declarations.MyParentProjectID = "00000000-0000-0000-0000-000000000000"
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна со списком групп товаров для выбора
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Create" Then
            If CheckDataFiling(1) = True Then
                SaveNewData()
                StartParam = "Edit"
                CheckRights()
                GetNewGroups()
            End If
        Else
            GetNewGroups()
        End If
        Me.SelectNextControl(sender, True, True, True, False)
    End Sub

    Private Sub TextBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox6.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub CheckBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CheckBox1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Function CheckDataFiling(ByVal MyParam As Integer) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '// MyParam = 0 проверяем все MyParam = 1 проверяем все кроме групп товаров в проекте
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox13.Text) = "" Then
            MsgBox("Поле ""Компания"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox1.Select()
            Exit Function
        End If

        If Trim(TextBox6.Text) = "" Then
            MsgBox("Поле ""Название проекта"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox1.Select()
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""Сумма (РУБ)"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox2.Select()
            Exit Function
        End If

        If Format(DateTimePicker1.Value, "dd/MM/yyyy") = "01/01/1900" Then
            MsgBox("дата начала проекта должна быть заполнена. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            DateTimePicker1.Select()
            Exit Function
        End If

        If Format(DateTimePicker2.Value, "dd/MM/yyyy") = "01/01/1900" Then
            MsgBox("дата окончания проекта должна быть заполнена. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            DateTimePicker2.Select()
            Exit Function
        End If

        If DateTimePicker2.Value <= DateTimePicker1.Value Then
            MsgBox("дата окончания проекта должна быть больше даты начала проекта. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            DateTimePicker1.Select()
            Exit Function
        End If

        If Format(DateTimePicker3.Value, "dd/MM/yyyy") = "01/01/1900" Then
            MsgBox("дата подачи предложения должна быть заполнена. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            DateTimePicker3.Select()
            Exit Function
        End If

        If DateTimePicker3.Value >= DateTimePicker2.Value Then
            MsgBox("дата подачи предложения должна быть меньше даты окончания проекта. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            DateTimePicker3.Select()
            Exit Function
        End If

        '---Проверка - есть ли для данного клиента проект с таким названием
        MySQLStr = "SELECT COUNT(ProjectName) AS CC "
        MySQLStr = MySQLStr & "FROM  tbl_CRM_Projects WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Declarations.MyClientID & "') AND "
        MySQLStr = MySQLStr & "(UPPER(ProjectName) = N'" & UCase(Trim(TextBox6.Text)) & "') AND "
        MySQLStr = MySQLStr & "(ProjectID <> N'" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            '---такого названия нет, можно созздавать
            trycloseMyRec()
        Else
            trycloseMyRec()
            MsgBox("Проект с таким названием уже существует для данного клиента. Введите другое название.", MsgBoxStyle.Critical, "Внимание!")
            CheckDataFiling = False
            TextBox1.Select()
            Exit Function
        End If

        '---Проверка - есть ли Скальский код у клиента. Если нет - такой проект утверждать нельзя.
        'If CheckBox3.CheckState = CheckState.Checked Then
        MySQLStr = "SELECT COUNT(CompanyID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Companies "
        MySQLStr = MySQLStr & "WHERE (ScalaCustomerCode IS NOT NULL) "
        MySQLStr = MySQLStr & "AND (CompanyID = '" & Declarations.MyClientID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            '---Скальский код отсутствует
            trycloseMyRec()
            MsgBox("Клиент, для которого вы ввели информацию по данному проекту, отсутствует в Scala и у него не существует Скальского кода клиента. Вам необходимо сначала завести клиента в Scala, дождаться выгрузки данного клиента в CRM, после чего назначить этому клиенту проект. При необходимости (если есть 2 одинаковых клиента, один со Скальским кодом, другой без), можно восползоваться функционалом ""Объединить"" для двух и более одинаковых клиентов. ", MsgBoxStyle.Critical, "Внимание!")
            CheckDataFiling = False
            TextBox6.Select()
            Exit Function
        Else

            trycloseMyRec()
        End If
        'End If

        '---Проверка - если выбрано закрыть проект - можно ли (если есть незакрытые действия по этому проекту, то нельзя)
        If CheckBox1.CheckState = CheckState.Checked Then
            MySQLStr = "SELECT COUNT(EventID) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ProjectID = N'" & Declarations.MyProjectID & "') AND "
            MySQLStr = MySQLStr & "(ActionClosed IS NULL) "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                '---незакрытых действий нет, можно закрывать
                trycloseMyRec()
            Else
                trycloseMyRec()
                MsgBox("Вы выбрали опцию 'Закрыть проект'. Однако по этому проекту есть незакрытые действия и закрывать проект, пока эти действия не будут закрыты, нельзя.", MsgBoxStyle.Critical, "Внимание!")
                CheckDataFiling = False
                TextBox1.Select()
                Exit Function
            End If
        End If

        '---Проверка - если выбрано "Утвержден" - должно быть заполнено поле % доп расходов по проекту 
        'If CheckBox3.CheckState = CheckState.Checked Then
        '    If Trim(TextBox10.Text) = "" Then
        '        MsgBox("Вы выбрали опцию 'Утвержден'. Однако при этом не заполнили поле '% доп расходов по проекту...'. Это поле обязательно должно быть заполнено.", MsgBoxStyle.Critical, "Внимание!")
        '        CheckDataFiling = False
        '        TextBox10.Select()
        '        Exit Function
        '    End If
        'End If

        '---Проверка - если выбрано "Утвержден" - должна быть загружена дополнительная информация по проекту
        'If CheckBox3.CheckState = CheckState.Checked Then
        '    MySQLStr = "SELECT COUNT(*) AS CC "
        '    MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Details WITH(NOLOCK) "
        '    MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
        '    InitMyConn(False)
        '    InitMyRec(False, MySQLStr)
        '    If Declarations.MyRec.Fields("CC").Value = 0 Then
        '        '---Дополнительные данные не загружены
        '        trycloseMyRec()
        '        MsgBox("Вы выбрали опцию ""Утвержден"". Однако при этом не загрузили детальную информацию по проекту из Excel файла ""Обоснование проекта"". Детальная информация должна быть загружена до утверждения.", MsgBoxStyle.Critical, "Внимание!")
        '        CheckDataFiling = False
        '        Button3.Select()
        '        Exit Function
        '    Else
        '        trycloseMyRec()
        '    End If
        'End If

        '---Проверка выбора групп товаров, поставляющихся в рамках данного проекта
        If MyParam = 0 Then     '---проверяем только если параметр равен 0
            If Trim(TextBox11.Text) = "" Then
                MsgBox("необходимо выбрать группы товаров, которые будут поставляться в рамках проекта", MsgBoxStyle.Critical, "Внимание!")
                CheckDataFiling = False
                Button4.Select()
                Exit Function
            End If
        End If

        CheckDataFiling = True

    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение информации о проекте
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling(0) = True Then
            If StartParam = "Create" Then
                SaveNewData()
            Else
                UpdateData()
            End If
            Declarations.MyResult = 1
            Me.Close()
        End If
    End Sub

    Private Sub SaveNewData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// сохранение данных в случае создания новой записи
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---создание записи в таблице tbl_CRM_Projects
        MySQLStr = "INSERT INTO tbl_CRM_Projects "
        MySQLStr = MySQLStr & "(ProjectID, CompanyID, ProjectName, ProjectSumm, ProjectComment, StartDate, CloseDate, FirstDate, LastDate, "
        MySQLStr = MySQLStr & "ProjectAddr, Investor, Contractor, ResponciblePerson, ProposalDate, ManufacturersList, AlterManufacturers, "
        MySQLStr = MySQLStr & "Competitors, AdditionalExpencesPerCent, IsApproved, IsIPG) "
        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyProjectID & "', "
        MySQLStr = MySQLStr & "'" & Declarations.MyClientID & "', "
        MySQLStr = MySQLStr & "N'" & TextBox6.Text & "', "
        MySQLStr = MySQLStr & Replace(CStr(CDbl(TextBox2.Text)), ",", ".") & ", "
        MySQLStr = MySQLStr & "N'" & TextBox1.Text & "', "
        MySQLStr = MySQLStr & "GetDate(), "
        MySQLStr = MySQLStr & "NULL, "
        MySQLStr = MySQLStr & "Convert(datetime,'" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103), "
        MySQLStr = MySQLStr & "Convert(datetime,'" & Format(DateTimePicker2.Value, "dd/MM/yyyy") & "', 103), "
        MySQLStr = MySQLStr & "N'" & TextBox3.Text & "', "
        MySQLStr = MySQLStr & "N'" & TextBox4.Text & "', "
        MySQLStr = MySQLStr & "N'" & TextBox5.Text & "', "
        MySQLStr = MySQLStr & "N'" & TextBox7.Text & "', "
        MySQLStr = MySQLStr & "Convert(datetime,'" & Format(DateTimePicker3.Value, "dd/MM/yyyy") & "', 103), "
        MySQLStr = MySQLStr & "N'" & TextBox8.Text & "', "
        MySQLStr = MySQLStr & CheckBox2.CheckState & ", "
        MySQLStr = MySQLStr & "N'" & TextBox9.Text & "', "
        If Trim(TextBox10.Text) = "" Then
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & Replace(CStr(CDbl(TextBox10.Text)), ",", ".") & ", "
        End If
        MySQLStr = MySQLStr & CheckBox3.CheckState & ", "
        If CheckBox4.CheckState = CheckState.Checked Then
            MySQLStr = MySQLStr & "1) "
        Else
            MySQLStr = MySQLStr & "0) "
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---создание записи в таблице tbl_CRM_Projects_Ext
        If Declarations.MyParentProjectID <> "00000000-0000-0000-0000-000000000000" Then
            MySQLStr = "DELETE FROM tbl_CRM_Projects_Ext "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            MySQLStr = "INSERT INTO tbl_CRM_Projects_Ext "
            MySQLStr = MySQLStr & "(ProjectID, ParentProjectID) "
            MySQLStr = MySQLStr & "VALUES     ('" & Declarations.MyProjectID & "', "
            MySQLStr = MySQLStr & "'" & Declarations.MyParentProjectID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

        '---создание записи в таблице tbl_CRM_Projects_Stages
        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Projects_Stages "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            MySQLStr = "INSERT INTO tbl_CRM_Projects_Stages "
            MySQLStr = MySQLStr & "(ProjectID, ProjectStageID, CHUser) "
            MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyProjectID & "', "
            MySQLStr = MySQLStr & CStr(ComboBox2.SelectedValue) & ", "
            MySQLStr = MySQLStr & "N'" & Declarations.FullName & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        Else
            MySQLStr = "UPDATE tbl_CRM_Projects_Stages "
            MySQLStr = MySQLStr & "SET ProjectStageID = " & CStr(ComboBox2.SelectedValue) & ", "
            MySQLStr = MySQLStr & "CHUser = N '" & Declarations.FullName & "' "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

        '---создание записи в таблице tbl_CRM_Projects_History
        MySQLStr = "INSERT INTO tbl_CRM_Projects_History "
        MySQLStr = MySQLStr & "(ProjectID, CompanyID, ProjectName, ProjectSumm, ProjectComment, StartDate, CloseDate, FirstDate, "
        MySQLStr = MySQLStr & "LastDate, ProjectAddr, Investor, Contractor, ResponciblePerson, ProposalDate, "
        MySQLStr = MySQLStr & "ManufacturersList, AlterManufacturers, Competitors, AdditionalExpencesPerCent, IsApproved, IsIPG, Editor, ActDate) "
        MySQLStr = MySQLStr & "SELECT ProjectID, CompanyID, ProjectName, ProjectSumm, ProjectComment, StartDate, CloseDate, FirstDate, LastDate, "
        MySQLStr = MySQLStr & "ProjectAddr, Investor, Contractor, ResponciblePerson, ProposalDate, "
        MySQLStr = MySQLStr & "ManufacturersList, AlterManufacturers, Competitors, AdditionalExpencesPerCent, IsApproved, IsIPG, "
        MySQLStr = MySQLStr & "N'" & Declarations.FullName & "' AS Editor, Getdate() AS ActDate "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Projects "
        MySQLStr = MySQLStr & "WHERE (ProjectID = N'" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub

    Private Sub UpdateData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// сохранение данных в случае редактирования записи
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyRez As MsgBoxResult

        If CheckBox1.CheckState = CheckState.Checked Then
            MyRez = MsgBox("В окне редактирования выбрано закрытие проекта. Это значит, что после сохранения этот проект будет недоступен для редактирования и использования (на него нелльзя будет ссылаться). Вы уверены, что хотите сохранить проект и тем закрыть его? ", MsgBoxStyle.YesNo, "Внимание!")
        Else
            MyRez = MsgBoxResult.Yes
        End If
        If MyRez = MsgBoxResult.Yes Then
            '---Редактирование существующей записи в таблице tbl_CRM_Projects
            MySQLStr = "UPDATE tbl_CRM_Projects "
            MySQLStr = MySQLStr & "SET CompanyID = N'" & Declarations.MyClientID & "', "
            MySQLStr = MySQLStr & "ProjectName = N'" & TextBox6.Text & "', "
            MySQLStr = MySQLStr & "ProjectComment = N'" & TextBox1.Text & "', "
            If CheckBox1.CheckState = CheckState.Checked Then
                MySQLStr = MySQLStr & "CloseDate = GetDate(), "
            Else
                MySQLStr = MySQLStr & "CloseDate = NULL, "
            End If
            MySQLStr = MySQLStr & "ProjectSumm = " & Replace(CStr(CDbl(TextBox2.Text)), ",", ".") & ", "
            MySQLStr = MySQLStr & "FirstDate = Convert(datetime,'" & Format(DateTimePicker1.Value, "dd/MM/yyyy") & "', 103), "
            MySQLStr = MySQLStr & "LastDate = Convert(datetime,'" & Format(DateTimePicker2.Value, "dd/MM/yyyy") & "', 103), "
            MySQLStr = MySQLStr & "ProjectAddr = N'" & TextBox3.Text & "', "
            MySQLStr = MySQLStr & "Investor = N'" & TextBox4.Text & "', "
            MySQLStr = MySQLStr & "Contractor = N'" & TextBox5.Text & "', "
            MySQLStr = MySQLStr & "ResponciblePerson = N'" & TextBox7.Text & "', "
            MySQLStr = MySQLStr & "ProposalDate = Convert(datetime,'" & Format(DateTimePicker3.Value, "dd/MM/yyyy") & "', 103), "
            MySQLStr = MySQLStr & "ManufacturersList = N'" & TextBox8.Text & "', "
            MySQLStr = MySQLStr & "AlterManufacturers = " & CheckBox2.CheckState & ", "
            MySQLStr = MySQLStr & "Competitors = N'" & TextBox9.Text & "', "
            If Trim(TextBox10.Text) = "" Then
                MySQLStr = MySQLStr & "AdditionalExpencesPerCent = NULL, "
            Else
                MySQLStr = MySQLStr & "AdditionalExpencesPerCent = " & Replace(CStr(CDbl(TextBox10.Text)), ",", ".") & ", "
            End If
            MySQLStr = MySQLStr & "IsApproved = " & CheckBox3.CheckState & ", "
            If CheckBox4.CheckState = CheckState.Checked Then
                MySQLStr = MySQLStr & "IsIPG = 1 "
            Else
                MySQLStr = MySQLStr & "IsIPG = 0 "
            End If
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            '---редактирование записи в таблице tbl_CRM_Projects_Ext
            MySQLStr = "DELETE FROM tbl_CRM_Projects_Ext "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            If Declarations.MyParentProjectID <> "00000000-0000-0000-0000-000000000000" Then
                MySQLStr = "INSERT INTO tbl_CRM_Projects_Ext "
                MySQLStr = MySQLStr & "(ProjectID, ParentProjectID) "
                MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyProjectID & "', "
                MySQLStr = MySQLStr & "'" & Declarations.MyParentProjectID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If

            '---редактирование записи в таблице tbl_CRM_Projects_Stages
            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Projects_Stages "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.Fields("CC").Value = 0 Then
                MySQLStr = "INSERT INTO tbl_CRM_Projects_Stages "
                MySQLStr = MySQLStr & "(ProjectID, ProjectStageID, CHUser) "
                MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyProjectID & "', "
                MySQLStr = MySQLStr & CStr(ComboBox2.SelectedValue) & ", "
                MySQLStr = MySQLStr & "N'" & Declarations.FullName & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            Else
                MySQLStr = "UPDATE tbl_CRM_Projects_Stages "
                MySQLStr = MySQLStr & "SET ProjectStageID = " & CStr(ComboBox2.SelectedValue) & ", "
                MySQLStr = MySQLStr & "CHUser = N'" & Declarations.FullName & "' "
                MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If

            '---создание записи в таблице tbl_CRM_Projects_History
            MySQLStr = "INSERT INTO tbl_CRM_Projects_History "
            MySQLStr = MySQLStr & "(ProjectID, CompanyID, ProjectName, ProjectSumm, ProjectComment, StartDate, CloseDate, FirstDate, "
            MySQLStr = MySQLStr & "LastDate, ProjectAddr, Investor, Contractor, ResponciblePerson, ProposalDate, "
            MySQLStr = MySQLStr & "ManufacturersList, AlterManufacturers, Competitors, AdditionalExpencesPerCent, IsApproved, IsIPG, Editor, ActDate) "
            MySQLStr = MySQLStr & "SELECT ProjectID, CompanyID, ProjectName, ProjectSumm, ProjectComment, StartDate, CloseDate, FirstDate, LastDate, "
            MySQLStr = MySQLStr & "ProjectAddr, Investor, Contractor, ResponciblePerson, ProposalDate, "
            MySQLStr = MySQLStr & "ManufacturersList, AlterManufacturers, Competitors, AdditionalExpencesPerCent, IsApproved, IsIPG, "
            MySQLStr = MySQLStr & "N'" & Declarations.FullName & "' AS Editor, Getdate() AS ActDate "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Projects "
            MySQLStr = MySQLStr & "WHERE (ProjectID = N'" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

    End Sub

    Private Sub TextBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox2.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox2_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox2.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения поля
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox2.Text) <> "" Then
            If InStr(TextBox2.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Сумма (РУБ)"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox2.Text
                Catch ex As Exception
                    MsgBox("В поле ""Сумма (РУБ)"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
            e.Cancel = False
        Else
            MsgBox("В поле ""Сумма (РУБ)"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
            e.Cancel = True
            Exit Sub
        End If
    End Sub

    Private Sub TextBox10_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox10.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox10_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox10.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения поля
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox10.Text) <> "" Then
            If InStr(TextBox10.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""% доп расходов по  проекту..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox10.Text
                Catch ex As Exception
                    MsgBox("В поле ""% доп расходов по  проекту..."" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
            If MyRez < 0 Then
                MsgBox("В поле ""% доп расходов по  проекту..."" должно быть введено ПОЛОЖИТЕЛЬНОЕ число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            End If
            e.Cancel = False
        End If
    End Sub

    Private Sub TextBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox5.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox7_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox7.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox8_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox8.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub TextBox9_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBox9.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub DateTimePicker1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub DateTimePicker2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DateTimePicker2.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
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

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(DateTimePicker1, True, True, True, False)
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(DateTimePicker1, True, True, True, False)
    End Sub

    Private Sub DateTimePicker3_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker3.ValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(DateTimePicker1, True, True, True, False)
    End Sub

    Private Sub CheckBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CheckBox2.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub CheckBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CheckBox3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна импорта детальной информации  по проекту
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MyProjectDetailsImport = New ProjectDetailsImport
        MyProjectDetailsImport.ShowDialog()
        '---Загружена детальная информация или нет
        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Details "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            CheckBox5.Checked = False
        Else
            If Declarations.MyRec.Fields("CC").Value > 0 Then
                trycloseMyRec()
                CheckBox5.Checked = True
            Else
                trycloseMyRec()
                CheckBox5.Checked = False
            End If
        End If
        If CheckBox5.Checked = True Then
            Button11.Enabled = True
        Else
            Button11.Enabled = False
        End If
        Me.SelectNextControl(sender, True, True, True, False)
    End Sub

    Private Sub GetNewGroups()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна со списком групп товаров для выбора и выбор
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MyItemGroupsInProject = New ItemGroupsInProject
        MyItemGroupsInProject.ShowDialog()
        '---составная строка - группы товаров в проекте
        MySQLStr = "SELECT distinct (SELECT tbl_CRM_ProdGroupsList.ItemGroupName + ';' AS 'data()' "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_ProdGroups AS t2 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_ProdGroupsList ON t2.ProdGroupID = tbl_CRM_ProdGroupsList.ID "
        MySQLStr = MySQLStr & "WHERE (t2.ProjectID = t1.ProjectID) ORDER BY tbl_CRM_ProdGroupsList.ID For xml path('')) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_ProdGroups AS t1 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_ProdGroupsList AS tbl_CRM_ProdGroupsList_1 ON t1.ProdGroupID = tbl_CRM_ProdGroupsList_1.ID "
        MySQLStr = MySQLStr & "WHERE (t1.ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            TextBox11.Text = ""
        Else
            Declarations.MyRec.MoveFirst()
            TextBox11.Text = Declarations.MyRec.Fields("CC").Value
            trycloseMyRec()
        End If
    End Sub

    Private Sub CheckBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles CheckBox4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена стадии проекта
        '//
        '////////////////////////////////////////////////////////////////////////////////
        If ComboBox2.SelectedValue = 100 Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If

        Me.SelectNextControl(sender, True, True, True, False)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Извлечение аттачмента из БД
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim mstream As ADODB.Stream

        Try
            MySQLStr = "SELECT AttachmentID, ProjectID, AttachmentName, AttachmentBody "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Attachments WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (AttachmentID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            MySQLStr = MySQLStr & "ORDER BY AttachmentName "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            SaveFileDialog1.FileName = Declarations.MyRec.Fields("AttachmentName").Value
            If SaveFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If (SaveFileDialog1.FileName <> "") Then
                    mstream = New ADODB.Stream
                    mstream.Type = StreamTypeEnum.adTypeBinary
                    mstream.Open()
                    mstream.Write(Declarations.MyRec.Fields("AttachmentBody").Value)
                    mstream.SaveToFile(SaveFileDialog1.FileName, SaveOptionsEnum.adSaveCreateOverWrite)
                End If
            End If
            trycloseMyRec()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление аттачмента из БД
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_CRM_Project_Attachments "
        MySQLStr = MySQLStr & "WHERE (AttachmentID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        LoadAttachments()
        CheckAttachmentsButtons()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление аттачмента к действию
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Create" Or StartParam = "Copy" Then
            '---Если запись только создается - ее надо сначала правильно заполнить и сохранить
            If CheckDataFiling(False) = True Then
                SaveNewData()
                StartParam = "Edit"
                CreateNewAttachment()
            End If
        Else
            CreateNewAttachment()
        End If
        LoadAttachments()
        CheckAttachmentsButtons()
    End Sub

    Private Sub CreateNewAttachment()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// занесение нового аттачмента в БД
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyGUID As Guid
        Dim FName As String
        Dim i As Integer
        Dim mstream As ADODB.Stream
        Dim FInfo As FileInfo

        If OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If (OpenFileDialog1.FileName <> "") Then
                MyGUID = Guid.NewGuid
                Declarations.MyAttachmentID = MyGUID.ToString
                '--имя файла без пути
                i = InStrRev(OpenFileDialog1.FileName, "\")
                FName = Microsoft.VisualBasic.Right(OpenFileDialog1.FileName, Len(OpenFileDialog1.FileName) - i)
                FInfo = New FileInfo(OpenFileDialog1.FileName)
                If FInfo.Length <> 0 Then
                    Try
                        MySQLStr = "INSERT INTO tbl_CRM_Project_Attachments "
                        MySQLStr = MySQLStr & "(AttachmentID, ProjectID, AttachmentName, AttachmentBody) "
                        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyAttachmentID & "', "
                        MySQLStr = MySQLStr & "'" & Declarations.MyProjectID & "', "
                        MySQLStr = MySQLStr & "N'" & FName & "', "
                        MySQLStr = MySQLStr & "NULL) "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)

                        MySQLStr = "SELECT AttachmentID, ProjectID, AttachmentName, AttachmentBody "
                        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Attachments WITH(NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (AttachmentID = '" & Declarations.MyAttachmentID & "') "
                        MySQLStr = MySQLStr & "ORDER BY AttachmentName "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        mstream = New ADODB.Stream
                        mstream.Type = StreamTypeEnum.adTypeBinary
                        mstream.Open()
                        mstream.LoadFromFile(OpenFileDialog1.FileName)
                        Declarations.MyRec.Fields("AttachmentBody").Value = mstream.Read
                        Declarations.MyRec.Update()
                        trycloseMyRec()

                    Catch ex As Exception
                        MySQLStr = "DELETE FROM tbl_CRM_Project_Attachments "
                        MySQLStr = MySQLStr & "WHERE (AttachmentID = '" & Declarations.MyAttachmentID & "') "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)
                        MsgBox(ex.ToString)
                    End Try
                Else
                    MsgBox("Файл " & OpenFileDialog1.FileName & " имеет нулевой размер и не может быть импортирован. ", MsgBoxStyle.Critical, "Внимние!")
                End If
            End If
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка детальной информации по проекту
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Cursor = Cursors.WaitCursor
        If My.Settings.UseOffice = "LibreOffice" Then
            UploadProductsToLO()
        Else
            UploadProductsToExcel()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub UploadProductsToLO()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка детальной информации по проекту в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFrame As Object
        Dim i As Integer                              'счетчик строк

        LOSetNotation(0)
        oServiceManager = CreateObject("com.sun.star.ServiceManager")
        oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
        oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
        Dim arg(1)
        arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
        arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
        oWorkBook = oDesktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg)
        oSheet = oWorkBook.getSheets().getByIndex(0)
        oFrame = oWorkBook.getCurrentController.getFrame

        i = 1
        ExportProductsHeaderToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)
        ExportProductsBodyToLO(oServiceManager, oDispatcher, oDesktop, oWorkBook, _
            oSheet, oFrame, i)

        Dim args() As Object
        ReDim args(0)
        args(0) = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
        args(0).Name = "ToPoint"
        args(0).Value = "$A$1"
        oDispatcher.executeDispatch(oFrame, ".uno:GoToCell", "", 0, args)

        oWorkBook.CurrentController.Frame.ContainerWindow.Visible = True
        oWorkBook.CurrentController.Frame.ContainerWindow.toFront()
    End Sub

    Private Sub UploadProductsToExcel()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка детальной информации по проекту в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As Object
        Dim MyWRKBook As Object
        Dim i As Integer                              'счетчик строк

        MyObj = CreateObject("Excel.Application")
        MyObj.SheetsInNewWorkbook = 1
        MyWRKBook = MyObj.Workbooks.Add

        i = 1
        ExportProductsHeaderToExcel(MyWRKBook, i)
        ExportProductsBodyToExcel(MyWRKBook, i)

        MyWRKBook.ActiveSheet.Range("A1").Select()
        MyObj.Application.Visible = True
        MyObj = Nothing

    End Sub

    Public Function ExportProductsHeaderToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '-----Ширина колонок
        oSheet.getColumns().getByName("A").Width = 3800
        oSheet.getColumns().getByName("B").Width = 9500
        oSheet.getColumns().getByName("C").Width = 3800
        oSheet.getColumns().getByName("D").Width = 7600
        oSheet.getColumns().getByName("E").Width = 3800
        oSheet.getColumns().getByName("F").Width = 3800
        oSheet.getColumns().getByName("G").Width = 3800
        oSheet.getColumns().getByName("H").Width = 3800
        oSheet.getColumns().getByName("I").Width = 1900
        oSheet.getColumns().getByName("J").Width = 3800

        '-----заголовок листа
        MySQLStr = "SELECT     tbl_CRM_Companies.ScalaCustomerCode + '  ' + tbl_CRM_Companies.CompanyName AS Company, "
        MySQLStr = MySQLStr & "'""' + tbl_CRM_Projects.ProjectName + '""' + ' Запланирован С ' + CONVERT(nvarchar(30), tbl_CRM_Projects.FirstDate,103) "
        MySQLStr = MySQLStr & " + ' По ' + CONVERT(nvarchar(30), tbl_CRM_Projects.LastDate, 103) AS Project, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects.ResponciblePerson, '') AS ResponciblePerson, tbl_CRM_Projects.AdditionalExpencesPerCent "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Projects INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Projects.CompanyID = tbl_CRM_Companies.CompanyID "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Projects.ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
        Else
            Declarations.MyRec.MoveFirst()
            oSheet.getCellRangeByName("B" & CStr(i)).String = "Детальная информация по проекту " & Declarations.MyRec.Fields("Project").Value
            oSheet.getCellRangeByName("B" & CStr(i + 1)).String = "Клиент " & Declarations.MyRec.Fields("Company").Value
            oSheet.getCellRangeByName("B" & CStr(i + 2)).String = "Ответственный " & Declarations.MyRec.Fields("ResponciblePerson").Value
            oSheet.getCellRangeByName("B" & CStr(i + 3)).String = "Процент доп. расходов " & Format(Declarations.MyRec.Fields("AdditionalExpencesPerCent").Value, "# ##0.00")
            trycloseMyRec()
        End If
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 3), "Arial")
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i + 3))
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i) & ":B" & CStr(i), 11)
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "B" & CStr(i + 1) & ":B" & CStr(i + 3), 9)

        '-----заголовок таблицы
        i = 6
        oSheet.getCellRangeByName("A" & CStr(i)).String = "Код запаса"
        oSheet.getCellRangeByName("B" & CStr(i)).String = "Имя запаса"
        oSheet.getCellRangeByName("C" & CStr(i)).String = "Код поставщика"
        oSheet.getCellRangeByName("D" & CStr(i)).String = "Поставщик"
        oSheet.getCellRangeByName("E" & CStr(i)).String = "Код товара поставщика"
        oSheet.getCellRangeByName("F" & CStr(i)).String = "Количество"
        oSheet.getCellRangeByName("G" & CStr(i)).String = "Себестоимость"
        oSheet.getCellRangeByName("H" & CStr(i)).String = "Цена"
        oSheet.getCellRangeByName("I" & CStr(i)).String = "Код валюты"
        oSheet.getCellRangeByName("J" & CStr(i)).String = "Валюта"
        LOFontSetFamilyName(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i), "Arial")
        LOFontSetSize(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i), 8)
        LOFontSetBold(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i))
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).CellBackColor = RGB(174, 249, 255)    '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        LOWrapText(oServiceManager, oDispatcher, oFrame, "A" & CStr(i) & ":J" & CStr(i))
        LOSetBorders(oServiceManager, oSheet, "A" & CStr(i) & ":J" & CStr(i), 70, 70, RGB(0, 0, 0)) '---тут уроды поменяли местами - на самом деле это не RGB а BGR!!!! первый цвет синий
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).VertJustify = 2
        oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i)).HoriJustify = 2

        i = 7
    End Function

    Public Function ExportProductsHeaderToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка заголовка продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '-----заголовок листа
        MySQLStr = "SELECT     tbl_CRM_Companies.ScalaCustomerCode + '  ' + tbl_CRM_Companies.CompanyName AS Company, "
        MySQLStr = MySQLStr & "'""' + tbl_CRM_Projects.ProjectName + '""' + ' Запланирован С ' + CONVERT(nvarchar(30), tbl_CRM_Projects.FirstDate,103) "
        MySQLStr = MySQLStr & " + ' По ' + CONVERT(nvarchar(30), tbl_CRM_Projects.LastDate, 103) AS Project, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects.ResponciblePerson, '') AS ResponciblePerson, tbl_CRM_Projects.AdditionalExpencesPerCent "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Projects INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Projects.CompanyID = tbl_CRM_Companies.CompanyID "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Projects.ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
        Else
            Declarations.MyRec.MoveFirst()
            MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "Детальная информация по проекту " & Declarations.MyRec.Fields("Project").Value
            MyWRKBook.ActiveSheet.Range("B" & CStr(i + 1)) = "Клиент " & Declarations.MyRec.Fields("Company").Value
            MyWRKBook.ActiveSheet.Range("B" & CStr(i + 2)) = "Ответственный " & Declarations.MyRec.Fields("ResponciblePerson").Value
            MyWRKBook.ActiveSheet.Range("B" & CStr(i + 3)) = "Процент доп. расходов " & Format(Declarations.MyRec.Fields("AdditionalExpencesPerCent").Value, "# ##0.00")
            trycloseMyRec()
        End If
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i)).Font.Size = 11
        MyWRKBook.ActiveSheet.Range("B" & CStr(i + 1) & ":B" & CStr(i + 3)).Font.Size = 9
        MyWRKBook.ActiveSheet.Range("B" & CStr(i) & ":B" & CStr(i + 3)).Font.Bold = True

        '-----заголовок таблицы
        i = 6
        MyWRKBook.ActiveSheet.Range("A" & CStr(i)) = "Код запаса"
        MyWRKBook.ActiveSheet.Range("B" & CStr(i)) = "Имя запаса"
        MyWRKBook.ActiveSheet.Range("C" & CStr(i)) = "Код поставщика"
        MyWRKBook.ActiveSheet.Range("D" & CStr(i)) = "Поставщик"
        MyWRKBook.ActiveSheet.Range("E" & CStr(i)) = "Код товара поставщика"
        MyWRKBook.ActiveSheet.Range("F" & CStr(i)) = "Количество"
        MyWRKBook.ActiveSheet.Range("G" & CStr(i)) = "Себестоимость"
        MyWRKBook.ActiveSheet.Range("H" & CStr(i)) = "Цена"
        MyWRKBook.ActiveSheet.Range("I" & CStr(i)) = "Код валюты"
        MyWRKBook.ActiveSheet.Range("J" & CStr(i)) = "Валюта"

        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Font.Size = 8
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Font.Bold = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).WrapText = True
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).VerticalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).HorizontalAlignment = -4108
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Select()
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Borders(5).LineStyle = -4142
        MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Borders(6).LineStyle = -4142

        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Borders(7)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Borders(8)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Borders(9)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Borders(10)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Borders(11)
            .LineStyle = 1
            .Weight = 4
            .ColorIndex = -4105
        End With
        With MyWRKBook.ActiveSheet.Range("A" & CStr(i) & ":J" & CStr(i)).Interior
            .ColorIndex = 36
            .Pattern = 1
            .PatternColorIndex = -4105
        End With

        MyWRKBook.ActiveSheet.Columns("A:A").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("B:B").ColumnWidth = 50
        MyWRKBook.ActiveSheet.Columns("C:C").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("D:D").ColumnWidth = 40
        MyWRKBook.ActiveSheet.Columns("E:E").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("F:F").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("G:G").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("H:H").ColumnWidth = 20
        MyWRKBook.ActiveSheet.Columns("I:I").ColumnWidth = 10
        MyWRKBook.ActiveSheet.Columns("J:J").ColumnWidth = 20

        i = 7
    End Function

    Public Function ExportProductsBodyToLO(ByRef oServiceManager As Object, ByRef oDispatcher As Object, _
        ByRef oDesktop As Object, ByRef oWorkBook As Object, ByRef oSheet As Object, _
        ByRef oFrame As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка списка продуктов в LibreOffice
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyArrStr() As Object
        Dim MyArr() As Object
        Dim MyL As Double
        Dim j As Integer
        Dim MyRange As Object

        MySQLStr = "SELECT tbl_CRM_Project_Details.ScalaItemCode, SC010300.SC01002 + ' ' + SC010300.SC01003 AS Name, "
        MySQLStr = MySQLStr & "SC010300.SC01058, PL010300.PL01002, SC010300.SC01060, tbl_CRM_Project_Details.QTY, "
        MySQLStr = MySQLStr & "tbl_CRM_Project_Details.ProjectPriCost, tbl_CRM_Project_Details.ProjectPrice, "
        MySQLStr = MySQLStr & "tbl_CRM_Project_Details.CurrCode, SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Details INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_CRM_Project_Details.ScalaItemCode = SC010300.SC01001 INNER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON tbl_CRM_Project_Details.CurrCode = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Project_Details.ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            Declarations.MyRec.MoveLast()
            MyL = Declarations.MyRec.RecordCount - 1
            ReDim MyArrStr(MyL)
            Declarations.MyRec.MoveFirst()
            j = 0
            While Not Declarations.MyRec.EOF
                ReDim MyArr(9)
                MyArr(0) = Declarations.MyRec.Fields(0).Value.ToString
                MyArr(1) = Declarations.MyRec.Fields(1).Value.ToString
                MyArr(2) = Declarations.MyRec.Fields(2).Value.ToString
                MyArr(3) = Declarations.MyRec.Fields(3).Value.ToString
                MyArr(4) = Declarations.MyRec.Fields(4).Value.ToString
                MyArr(5) = CDbl(Declarations.MyRec.Fields(5).Value)
                MyArr(6) = CDbl(Declarations.MyRec.Fields(6).Value)
                MyArr(7) = CDbl(Declarations.MyRec.Fields(7).Value)
                MyArr(8) = CInt(Declarations.MyRec.Fields(8).Value)
                MyArr(9) = Declarations.MyRec.Fields(9).Value.ToString
                MyArrStr(j) = MyArr

                j = j + 1
                Declarations.MyRec.MoveNext()
            End While
            MyRange = oSheet.getCellRangeByName("A" & CStr(i) & ":J" & CStr(i + MyL))
            MyRange.setDataArray(MyArrStr)
            LOFormatCells(oServiceManager, oDispatcher, oFrame, "F" & CStr(i) & ":H" & CStr(i + MyL), 4)
        End If
    End Function

    Public Function ExportProductsBodyToExcel(ByRef MyWRKBook As Object, ByRef i As Integer)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выгрузка списка продуктов в Excel
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT tbl_CRM_Project_Details.ScalaItemCode, SC010300.SC01002 + ' ' + SC010300.SC01003 AS Name, "
        MySQLStr = MySQLStr & "SC010300.SC01058, PL010300.PL01002, SC010300.SC01060, tbl_CRM_Project_Details.QTY, "
        MySQLStr = MySQLStr & "tbl_CRM_Project_Details.ProjectPriCost, tbl_CRM_Project_Details.ProjectPrice, "
        MySQLStr = MySQLStr & "tbl_CRM_Project_Details.CurrCode, SYCD0100.SYCD009 "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Details INNER JOIN "
        MySQLStr = MySQLStr & "SC010300 ON tbl_CRM_Project_Details.ScalaItemCode = SC010300.SC01001 INNER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 INNER JOIN "
        MySQLStr = MySQLStr & "SYCD0100 ON tbl_CRM_Project_Details.CurrCode = SYCD0100.SYCD001 "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Project_Details.ProjectID = '" & Declarations.MyProjectID & "') "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
            trycloseMyRec()
        Else
            MyWRKBook.ActiveSheet.Range("A" & CStr(i)).CopyFromRecordset(Declarations.MyRec)
            trycloseMyRec()
        End If
    End Function

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование запроса на портал на утверждение проекта
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim docQTY As Integer
        Dim MyRez As MsgBoxResult

        If CheckDataFiling(False) = True Then
            If StartParam = "Create" Then
                SaveNewData()
            Else
                UpdateData()
            End If
            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Attachments "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                MsgBox("Невозможно проверить количество документов, присоединенных к данному проекту. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                Exit Sub
            Else
                Declarations.MyRec.MoveFirst()
                docQTY = Declarations.MyRec.Fields("CC").Value
                trycloseMyRec()
            End If
            If docQTY = 0 Then
                MyRez = MsgBox("У вас нет ни одного документа, присоединенного к проекту. Вы уверены, что хотите создать заявку на утверждение проекта на портале?", MsgBoxStyle.YesNo, "Внимание!")
                If MyRez = MsgBoxResult.No Then
                    Exit Sub
                End If
            End If

            '-----создание заявки на утверждение проекта
            If CreateProjectConfirmrequest() = True Then
                MsgBox("Создана заявка на портале на утверждение проекта.", MsgBoxStyle.OkOnly, "Внимание!")
            End If
        End If
        
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование запроса на портал на утверждение деталей проекта и загрузку спецификации
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim docQTY As Integer
        Dim MyRez As MsgBoxResult
        Dim MyStr As String

        If CheckDataFiling(False) = True Then
            If StartParam = "Create" Then
                SaveNewData()
            Else
                UpdateData()
            End If

            MySQLStr = "SELECT COUNT(*) AS CC "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Attachments "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                trycloseMyRec()
                MsgBox("Невозможно проверить количество документов, присоединенных к данному проекту. Обратитесь к администратору.", MsgBoxStyle.Critical, "Внимание!")
                Exit Sub
            Else
                Declarations.MyRec.MoveFirst()
                docQTY = Declarations.MyRec.Fields("CC").Value
                trycloseMyRec()
            End If
            If docQTY = 0 Then
                MsgBox("У вас нет ни одного документа, присоединенного к проекту. Для утверждения деталей проекта и загрузки спецификации необходимо спецификацию проекта присоединить к проекту (добавить файл).", MsgBoxStyle.Critical, "Внимание!")
                Exit Sub
            End If
            MyStr = "Для утверждения деталей проекта и загрузки спецификации необходимо " & Chr(13) & Chr(10)
            MyStr = MyStr & "присоединить к проекту (добавить файл) следующие файлы: " & Chr(13) & Chr(10)
            MyStr = MyStr & "1) Файл расчета проекта " & Chr(13) & Chr(10)
            MyStr = MyStr & "2) Предложение клиенту " & Chr(13) & Chr(10)
            MyStr = MyStr & "3) Предложение от поставщика " & Chr(13) & Chr(10)
            MyStr = MyStr & "4) Письмо от менеджера по перевозкам со стоимостью доставки до клиента " & Chr(13) & Chr(10)
            MyStr = MyStr & "5) Письмо от менеджера по перевозкам со стоимостью доставки от поставщика " & Chr(13) & Chr(10) & Chr(13) & Chr(10)
            MyStr = MyStr & "Вы присоединили к проекту все необходимые документы и готовы разместить заявку на портале? "
            MyRez = MsgBox(MyStr, MsgBoxStyle.YesNo, "Внимание!")
            If MyRez = MsgBoxResult.No Then
                Exit Sub
            End If

            '-----создание заявки на утверждение деталей проекта и загрузку спецификации
            If CreateProjectDetailConfirmrequest() = True Then
                MsgBox("Создана заявка на портале на утверждение проекта.", MsgBoxStyle.OkOnly, "Внимание!")
            End If

            MsgBox("Создана заявка на портале на утверждение деталей проекта и загрузку спецификации.", MsgBoxStyle.OkOnly, "Внимание!")
        End If
    End Sub

    Private Function CreateProjectConfirmrequest() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// создание заявки на портал на утверждение проекта
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim ProjectName As String
        Dim ProjectID As String
        Dim IsDetail As Integer
        Dim CustName As String
        Dim CustCode As String
        Dim ProjectAddr As String
        Dim ProjectDescr As String
        Dim Investor As String
        Dim Contractor As String
        Dim Salesman As String
        Dim ProjectSumm As String
        Dim ProjectStartDate As String
        Dim ProjectFinDate As String
        Dim ProjectPropDate As String
        Dim SuppList As String
        Dim AlterManufacturers As Integer
        Dim Competitors As String
        Dim ProdGroups As String
        Dim ProjectStage As String
        Dim IsApproved As Integer
        Dim AdditionalExpencesPerCent As String
        Dim DetIsLoad As Integer

        Dim MyAttachment As Byte()
        Dim FileName As String = ""

        MySQLStr = "SELECT tbl_CRM_Projects.ProjectName, tbl_CRM_Projects.ProjectID, 0 AS IsDetail, tbl_CRM_Companies.CompanyName, "
        MySQLStr = MySQLStr & "tbl_CRM_Companies.ScalaCustomerCode, ISNULL(tbl_CRM_Projects.ProjectAddr, "
        MySQLStr = MySQLStr & "N'') AS ProjectAddr, ISNULL(tbl_CRM_Projects.ProjectComment, N'') AS ProjectComment, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects.Investor, N'') AS Investor, ISNULL(tbl_CRM_Projects.Contractor, N'') "
        MySQLStr = MySQLStr & "AS Contractor, CASE WHEN ScalaSystemDB.dbo.ScaUsers.UserName IS NULL THEN '' ELSE 'eskru\' + "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers.UserName END AS salesman, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects.ProjectSumm, 0) AS ProjectSumm, CASE WHEN tbl_CRM_Projects.StartDate IS NULL THEN '' ELSE Convert(nvarchar(30), tbl_CRM_Projects.StartDate, 103) END as StartDate, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_Projects.LastDate IS NULL THEN '' ELSE Convert(nvarchar(30), tbl_CRM_Projects.LastDate, 103) END AS LastDate, CASE WHEN tbl_CRM_Projects.ProposalDate IS NULL THEN '' ELSE Convert(nvarchar(30), "
        MySQLStr = MySQLStr & "tbl_CRM_Projects.ProposalDate, 103) END AS ProposalDate , tbl_CRM_Projects.ManufacturersList, "
        MySQLStr = MySQLStr & "tbl_CRM_Projects.AlterManufacturers, tbl_CRM_Projects.Competitors, ISNULL(View_16.CC, N'') AS ProdGroups, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects_StagesCFG.Name, N'') AS ProjectStage, "
        MySQLStr = MySQLStr & "tbl_CRM_Projects.IsApproved, ISNULL(tbl_CRM_Projects.AdditionalExpencesPerCent, 0) AS AdditionalExpencesPerCent, "
        MySQLStr = MySQLStr & "CASE WHEN View_17.Qty IS NULL THEN 0 ELSE 1 END AS DetIsLoad "
        MySQLStr = MySQLStr & "FROM (SELECT ProjectID, COUNT(ProjectID) AS Qty "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Details "
        MySQLStr = MySQLStr & "GROUP BY ProjectID) AS View_17 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Projects.CompanyID = tbl_CRM_Companies.CompanyID ON "
        MySQLStr = MySQLStr & "View_17.ProjectID = tbl_CRM_Projects.ProjectID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects_StagesCFG INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects_Stages ON tbl_CRM_Projects_StagesCFG.ID = tbl_CRM_Projects_Stages.ProjectStageID "
        MySQLStr = MySQLStr & "ON tbl_CRM_Projects.ProjectID = tbl_CRM_Projects_Stages.ProjectID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT distinct '" & Declarations.MyProjectID & "' AS ID,  "
        MySQLStr = MySQLStr & "(SELECT tbl_CRM_ProdGroupsList.ItemGroupName + ';' AS 'data()' "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_ProdGroups AS t2 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_ProdGroupsList ON t2.ProdGroupID = tbl_CRM_ProdGroupsList.ID "
        MySQLStr = MySQLStr & "WHERE (t2.ProjectID = t1.ProjectID) ORDER BY tbl_CRM_ProdGroupsList.ID For xml path('')) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_ProdGroups AS t1 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_ProdGroupsList AS tbl_CRM_ProdGroupsList_1 ON t1.ProdGroupID = tbl_CRM_ProdGroupsList_1.ID "
        MySQLStr = MySQLStr & "WHERE (t1.ProjectID = '" & Declarations.MyProjectID & "') ) AS View_16 ON "
        MySQLStr = MySQLStr & "tbl_CRM_Projects.ProjectID = View_16.ID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Projects.ResponciblePerson = ScalaSystemDB.dbo.ScaUsers.FullName "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Projects.ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            CreateProjectConfirmrequest = False
            Exit Function
        Else
            Declarations.MyRec.MoveFirst()
            ProjectName = Declarations.MyRec.Fields("ProjectName").Value
            Dim MyGuid As Guid = Guid.Parse(Declarations.MyRec.Fields("ProjectID").Value)
            ProjectID = MyGuid.ToString("D")
            'ProjectID = Declarations.MyRec.Fields("ProjectID").Value.ToString()
            IsDetail = Declarations.MyRec.Fields("IsDetail").Value
            CustName = Declarations.MyRec.Fields("CompanyName").Value
            CustCode = Trim(Declarations.MyRec.Fields("ScalaCustomerCode").Value)
            ProjectAddr = Declarations.MyRec.Fields("ProjectAddr").Value
            ProjectDescr = Declarations.MyRec.Fields("ProjectComment").Value
            Investor = Declarations.MyRec.Fields("Investor").Value
            Contractor = Declarations.MyRec.Fields("Contractor").Value
            Salesman = Declarations.MyRec.Fields("salesman").Value
            ProjectSumm = Format(Declarations.MyRec.Fields("ProjectSumm").Value, "0")
            ProjectStartDate = Declarations.MyRec.Fields("StartDate").Value
            ProjectFinDate = Declarations.MyRec.Fields("LastDate").Value
            ProjectPropDate = Declarations.MyRec.Fields("ProposalDate").Value
            SuppList = Declarations.MyRec.Fields("ManufacturersList").Value
            AlterManufacturers = Declarations.MyRec.Fields("AlterManufacturers").Value
            Competitors = Declarations.MyRec.Fields("Competitors").Value
            ProdGroups = Declarations.MyRec.Fields("ProdGroups").Value
            ProjectStage = Declarations.MyRec.Fields("ProjectStage").Value
            IsApproved = Declarations.MyRec.Fields("IsApproved").Value
            AdditionalExpencesPerCent = Replace(Declarations.MyRec.Fields("AdditionalExpencesPerCent").Value.ToString(), ",", ".")
            DetIsLoad = Declarations.MyRec.Fields("DetIsLoad").Value

            trycloseMyRec()
        End If

        Dim listWebService As spbprd4.Lists = New spbprd4.Lists()
        listWebService.Credentials = New System.Net.NetworkCredential("developer", "!Devpass", "ESKRU")
        Dim listName = "{371adff7-b7b1-4d42-97d7-ec49e1b880fe}"
        Dim listView = ""
        Dim listItemId As String = ""


        Dim strBatch As String = "<Method ID='1' Cmd='New'>"
        strBatch = strBatch + "<Field Name='ID'>New</Field>"
        strBatch = strBatch + "<Field Name='Title'>" & ProjectName & "</Field>"                                         '---название проекта
        strBatch = strBatch + "<Field Name='ProjectID'>" & ProjectID & "</Field>"                                       '---ID проекта
        strBatch = strBatch + "<Field Name='_x0421__x043e__x0433__x043b__x04'>" & IsDetail & "</Field>"                 '---на утверждение проекта (0) или на согласование деталей (1)
        strBatch = strBatch + "<Field Name='_x041a__x043e__x043c__x043f__x04'>" & CustName & "</Field>"                 '---имя клиента
        strBatch = strBatch + "<Field Name='_x041a__x043e__x0434__x0020__x04'>" & CustCode & "</Field>"                 '---код клиента в Scala
        strBatch = strBatch + "<Field Name='_x0410__x0434__x0440__x0435__x04'>" & ProjectAddr & "</Field>"              '---адрес проекта
        strBatch = strBatch + "<Field Name='_x041e__x043f__x0438__x0441__x04'>" & ProjectDescr & "</Field>"             '---описание проекта
        strBatch = strBatch + "<Field Name='_x0418__x043d__x0432__x0435__x04'>" & Investor & "</Field>"                 '---инвестор проекта
        strBatch = strBatch + "<Field Name='_x041f__x043e__x0434__x0440__x04'>" & Contractor & "</Field>"               '---подрядчик проекта
        strBatch = strBatch + "<Field Name='_x041f__x0440__x043e__x0434__x04'>" & Salesman & "</Field>"                 '---продавец
        strBatch = strBatch + "<Field Name='_x0421__x0443__x043c__x043c__x04'>" & ProjectSumm & "</Field>"              '---сумма проекта
        strBatch = strBatch + "<Field Name='_x0434__x0430__x0442__x0430__x00'>" & ProjectStartDate & "</Field>"         '---дата начала проекта
        strBatch = strBatch + "<Field Name='_x043f__x043b__x0430__x043d__x00'>" & ProjectFinDate & "</Field>"           '---дата окончания проекта
        strBatch = strBatch + "<Field Name='_x0414__x0430__x0442__x0430__x00'>" & ProjectPropDate & "</Field>"          '---дата подачи предложения по проекту
        strBatch = strBatch + "<Field Name='_x041f__x0435__x0440__x0435__x04'>" & SuppList & "</Field>"                 '---список производителей
        strBatch = strBatch + "<Field Name='_x0412__x043e__x0437__x043c__x04'>" & AlterManufacturers & "</Field>"       '---возможность альтернативных производителей
        strBatch = strBatch + "<Field Name='_x041a__x043e__x043d__x043a__x04'>" & Competitors & "</Field>"              '---конкуренты
        strBatch = strBatch + "<Field Name='_x0413__x0440__x0443__x043f__x04'>" & ProdGroups & "</Field>"               '---группы продуктов
        strBatch = strBatch + "<Field Name='_x042d__x0442__x0430__x043f__x00'>" & ProjectStage & "</Field>"             '---стадия проекта
        strBatch = strBatch + "<Field Name='_x0423__x0442__x0432__x0435__x04'>" & IsApproved & "</Field>"               '---утвержден или нет
        strBatch = strBatch + "<Field Name='_x041f__x0440__x043e__x0446__x04'>" & AdditionalExpencesPerCent & "</Field>" '---дополнительные расходы по проекту
        strBatch = strBatch + "<Field Name='_x0414__x0435__x0442__x0430__x04'>" & DetIsLoad & "</Field>"                '---детальная информация по проекту загружена
        strBatch = strBatch + "</Method>"

        Dim xmlDoc As XmlDocument = New System.Xml.XmlDocument()
        Dim elBatch As System.Xml.XmlElement = xmlDoc.CreateElement("Batch")
        elBatch.SetAttribute("OnError", "Continue")
        elBatch.SetAttribute("ListVersion", "1")
        elBatch.SetAttribute("ViewName", listView)
        elBatch.InnerXml = strBatch

        Try
            Dim ndReturn As XmlNode = listWebService.UpdateListItems(listName, elBatch)
            '-----аттачменты
            MySQLStr = "SELECT AttachmentName, AttachmentBody "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Attachments "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
            Else
                Declarations.MyRec.MoveFirst()
                While Declarations.MyRec.EOF <> True
                    Dim NewDoc As XmlDocument = New XmlDocument
                    NewDoc.LoadXml(ndReturn.OuterXml)
                    Dim NewNdList As XmlNodeList = NewDoc.GetElementsByTagName("z:row")
                    listItemId = NewNdList(0).Attributes("ows_ID").Value.ToString
                    '---имя файла
                    FileName = Declarations.MyRec.Fields("AttachmentName").Value
                    '---аттачмент
                    MyAttachment = Declarations.MyRec.Fields("AttachmentBody").Value
                    listWebService.AddAttachment(listName, listItemId, FileName, MyAttachment)

                    Declarations.MyRec.MoveNext()
                End While
                trycloseMyRec()
            End If

        Catch ex As Exception
            MsgBox("Ошибка создания заявки на портале " + ex.Message, MsgBoxStyle.Critical, "Внимание!")
            CreateProjectConfirmrequest = False
            Exit Function
        End Try


        CreateProjectConfirmrequest = True
    End Function

    Private Function CreateProjectDetailConfirmrequest() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// создание заявки на портал на утверждение деталей проекта и загрузку спецификации
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim ProjectName As String
        Dim ProjectID As String
        Dim IsDetail As Integer
        Dim CustName As String
        Dim CustCode As String
        Dim ProjectAddr As String
        Dim ProjectDescr As String
        Dim Investor As String
        Dim Contractor As String
        Dim Salesman As String
        Dim ProjectSumm As String
        Dim ProjectStartDate As String
        Dim ProjectFinDate As String
        Dim ProjectPropDate As String
        Dim SuppList As String
        Dim AlterManufacturers As Integer
        Dim Competitors As String
        Dim ProdGroups As String
        Dim ProjectStage As String
        Dim IsApproved As Integer
        Dim AdditionalExpencesPerCent As String
        Dim DetIsLoad As Integer

        Dim MyAttachment As Byte()
        Dim FileName As String = ""

        MySQLStr = "SELECT tbl_CRM_Projects.ProjectName, tbl_CRM_Projects.ProjectID, -1 AS IsDetail, tbl_CRM_Companies.CompanyName, "
        MySQLStr = MySQLStr & "tbl_CRM_Companies.ScalaCustomerCode, ISNULL(tbl_CRM_Projects.ProjectAddr, "
        MySQLStr = MySQLStr & "N'') AS ProjectAddr, ISNULL(tbl_CRM_Projects.ProjectComment, N'') AS ProjectComment, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects.Investor, N'') AS Investor, ISNULL(tbl_CRM_Projects.Contractor, N'') "
        MySQLStr = MySQLStr & "AS Contractor, CASE WHEN ScalaSystemDB.dbo.ScaUsers.UserName IS NULL THEN '' ELSE 'eskru\' + "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers.UserName END AS salesman, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects.ProjectSumm, 0) AS ProjectSumm, CASE WHEN tbl_CRM_Projects.StartDate IS NULL THEN '' ELSE Convert(nvarchar(30), tbl_CRM_Projects.StartDate, 103) END as StartDate, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_Projects.LastDate IS NULL THEN '' ELSE Convert(nvarchar(30), tbl_CRM_Projects.LastDate, 103) END AS LastDate, CASE WHEN tbl_CRM_Projects.ProposalDate IS NULL THEN '' ELSE Convert(nvarchar(30), "
        MySQLStr = MySQLStr & "tbl_CRM_Projects.ProposalDate, 103) END AS ProposalDate , tbl_CRM_Projects.ManufacturersList, "
        MySQLStr = MySQLStr & "tbl_CRM_Projects.AlterManufacturers, tbl_CRM_Projects.Competitors, ISNULL(View_16.CC, N'') AS ProdGroups, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Projects_StagesCFG.Name, N'') AS ProjectStage, "
        MySQLStr = MySQLStr & "tbl_CRM_Projects.IsApproved, ISNULL(tbl_CRM_Projects.AdditionalExpencesPerCent, 0) AS AdditionalExpencesPerCent, "
        MySQLStr = MySQLStr & "CASE WHEN View_17.Qty IS NULL THEN 0 ELSE 1 END AS DetIsLoad "
        MySQLStr = MySQLStr & "FROM (SELECT ProjectID, COUNT(ProjectID) AS Qty "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Details "
        MySQLStr = MySQLStr & "GROUP BY ProjectID) AS View_17 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Projects.CompanyID = tbl_CRM_Companies.CompanyID ON "
        MySQLStr = MySQLStr & "View_17.ProjectID = tbl_CRM_Projects.ProjectID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects_StagesCFG INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects_Stages ON tbl_CRM_Projects_StagesCFG.ID = tbl_CRM_Projects_Stages.ProjectStageID "
        MySQLStr = MySQLStr & "ON tbl_CRM_Projects.ProjectID = tbl_CRM_Projects_Stages.ProjectID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT distinct '" & Declarations.MyProjectID & "' AS ID,  "
        MySQLStr = MySQLStr & "(SELECT tbl_CRM_ProdGroupsList.ItemGroupName + ';' AS 'data()' "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_ProdGroups AS t2 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_ProdGroupsList ON t2.ProdGroupID = tbl_CRM_ProdGroupsList.ID "
        MySQLStr = MySQLStr & "WHERE (t2.ProjectID = t1.ProjectID) ORDER BY tbl_CRM_ProdGroupsList.ID For xml path('')) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_ProdGroups AS t1 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_ProdGroupsList AS tbl_CRM_ProdGroupsList_1 ON t1.ProdGroupID = tbl_CRM_ProdGroupsList_1.ID "
        MySQLStr = MySQLStr & "WHERE (t1.ProjectID = '" & Declarations.MyProjectID & "') ) AS View_16 ON "
        MySQLStr = MySQLStr & "tbl_CRM_Projects.ProjectID = View_16.ID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Projects.ResponciblePerson = ScalaSystemDB.dbo.ScaUsers.FullName "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Projects.ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            CreateProjectDetailConfirmrequest = False
            Exit Function
        Else
            Declarations.MyRec.MoveFirst()
            ProjectName = Declarations.MyRec.Fields("ProjectName").Value
            Dim MyGuid As Guid = Guid.Parse(Declarations.MyRec.Fields("ProjectID").Value)
            ProjectID = MyGuid.ToString("D")
            'ProjectID = Declarations.MyRec.Fields("ProjectID").Value.ToString()
            IsDetail = Declarations.MyRec.Fields("IsDetail").Value
            CustName = Declarations.MyRec.Fields("CompanyName").Value
            CustCode = Trim(Declarations.MyRec.Fields("ScalaCustomerCode").Value)
            ProjectAddr = Declarations.MyRec.Fields("ProjectAddr").Value
            ProjectDescr = Declarations.MyRec.Fields("ProjectComment").Value
            Investor = Declarations.MyRec.Fields("Investor").Value
            Contractor = Declarations.MyRec.Fields("Contractor").Value
            Salesman = Declarations.MyRec.Fields("salesman").Value
            ProjectSumm = Declarations.MyRec.Fields("ProjectSumm").Value.ToString()
            ProjectStartDate = Declarations.MyRec.Fields("StartDate").Value
            ProjectFinDate = Declarations.MyRec.Fields("LastDate").Value
            ProjectPropDate = Declarations.MyRec.Fields("ProposalDate").Value
            SuppList = Declarations.MyRec.Fields("ManufacturersList").Value
            AlterManufacturers = Declarations.MyRec.Fields("AlterManufacturers").Value
            Competitors = Declarations.MyRec.Fields("Competitors").Value
            ProdGroups = Declarations.MyRec.Fields("ProdGroups").Value
            ProjectStage = Declarations.MyRec.Fields("ProjectStage").Value
            IsApproved = Declarations.MyRec.Fields("IsApproved").Value
            AdditionalExpencesPerCent = Declarations.MyRec.Fields("AdditionalExpencesPerCent").Value.ToString()
            DetIsLoad = Declarations.MyRec.Fields("DetIsLoad").Value

            trycloseMyRec()
        End If

        Dim listWebService As spbprd4.Lists = New spbprd4.Lists()
        listWebService.Credentials = New System.Net.NetworkCredential("developer", "!Devpass", "ESKRU")
        Dim listName = "{371adff7-b7b1-4d42-97d7-ec49e1b880fe}"
        Dim listView = ""
        Dim listItemId As String = ""

        Dim strBatch As String = "<Method ID='1' Cmd='New'>"
        strBatch = strBatch + "<Field Name='ID'>New</Field>"
        strBatch = strBatch + "<Field Name='Title'>" & ProjectName & "</Field>"                                         '---название проекта
        strBatch = strBatch + "<Field Name='ProjectID'>" & ProjectID & "</Field>"                                       '---ID проекта
        strBatch = strBatch + "<Field Name='_x0421__x043e__x0433__x043b__x04'>" & IsDetail & "</Field>"                 '---на утверждение проекта (0) или на согласование деталей (1)
        strBatch = strBatch + "<Field Name='_x041a__x043e__x043c__x043f__x04'>" & CustName & "</Field>"                 '---имя клиента
        strBatch = strBatch + "<Field Name='_x041a__x043e__x0434__x0020__x04'>" & CustCode & "</Field>"                 '---код клиента в Scala
        strBatch = strBatch + "<Field Name='_x0410__x0434__x0440__x0435__x04'>" & ProjectAddr & "</Field>"              '---адрес проекта
        strBatch = strBatch + "<Field Name='_x041e__x043f__x0438__x0441__x04'>" & ProjectDescr & "</Field>"             '---описание проекта
        strBatch = strBatch + "<Field Name='_x0418__x043d__x0432__x0435__x04'>" & Investor & "</Field>"                 '---инвестор проекта
        strBatch = strBatch + "<Field Name='_x041f__x043e__x0434__x0440__x04'>" & Contractor & "</Field>"               '---подрядчик проекта
        strBatch = strBatch + "<Field Name='_x041f__x0440__x043e__x0434__x04'>" & Salesman & "</Field>"                 '---продавец
        strBatch = strBatch + "<Field Name='_x0421__x0443__x043c__x043c__x04'>" & ProjectSumm & "</Field>"              '---сумма проекта
        strBatch = strBatch + "<Field Name='_x0434__x0430__x0442__x0430__x00'>" & ProjectStartDate & "</Field>"         '---дата начала проекта
        strBatch = strBatch + "<Field Name='_x043f__x043b__x0430__x043d__x00'>" & ProjectFinDate & "</Field>"           '---дата окончания проекта
        strBatch = strBatch + "<Field Name='_x0414__x0430__x0442__x0430__x00'>" & ProjectPropDate & "</Field>"          '---дата подачи предложения по проекту
        strBatch = strBatch + "<Field Name='_x041f__x0435__x0440__x0435__x04'>" & SuppList & "</Field>"                 '---список производителей
        strBatch = strBatch + "<Field Name='_x0412__x043e__x0437__x043c__x04'>" & AlterManufacturers & "</Field>"       '---возможность альтернативных производителей
        strBatch = strBatch + "<Field Name='_x041a__x043e__x043d__x043a__x04'>" & Competitors & "</Field>"              '---конкуренты
        strBatch = strBatch + "<Field Name='_x0413__x0440__x0443__x043f__x04'>" & ProdGroups & "</Field>"               '---группы продуктов
        strBatch = strBatch + "<Field Name='_x042d__x0442__x0430__x043f__x00'>" & ProjectStage & "</Field>"             '---стадия проекта
        strBatch = strBatch + "<Field Name='_x0423__x0442__x0432__x0435__x04'>" & IsApproved & "</Field>"               '---утвержден или нет
        strBatch = strBatch + "<Field Name='_x041f__x0440__x043e__x0446__x04'>" & AdditionalExpencesPerCent & "</Field>" '---дополнительные расходы по проекту
        strBatch = strBatch + "<Field Name='_x0414__x0435__x0442__x0430__x04'>" & DetIsLoad & "</Field>"                '---детальная информация по проекту загружена
        strBatch = strBatch + "</Method>"

        Dim xmlDoc As XmlDocument = New System.Xml.XmlDocument()
        Dim elBatch As System.Xml.XmlElement = xmlDoc.CreateElement("Batch")
        elBatch.SetAttribute("OnError", "Continue")
        elBatch.SetAttribute("ListVersion", "1")
        elBatch.SetAttribute("ViewName", listView)
        elBatch.InnerXml = strBatch

        Try
            Dim ndReturn As XmlNode = listWebService.UpdateListItems(listName, elBatch)
            '-----аттачменты
            MySQLStr = "SELECT AttachmentName, AttachmentBody "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Project_Attachments "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
            Else
                Declarations.MyRec.MoveFirst()
                While Declarations.MyRec.EOF <> True
                    Dim NewDoc As XmlDocument = New XmlDocument
                    NewDoc.LoadXml(ndReturn.OuterXml)
                    Dim NewNdList As XmlNodeList = NewDoc.GetElementsByTagName("z:row")
                    listItemId = NewNdList(0).Attributes("ows_ID").Value.ToString
                    '---имя файла
                    FileName = Declarations.MyRec.Fields("AttachmentName").Value
                    '---аттачмент
                    MyAttachment = Declarations.MyRec.Fields("AttachmentBody").Value
                    listWebService.AddAttachment(listName, listItemId, FileName, MyAttachment)

                    Declarations.MyRec.MoveNext()
                End While
                trycloseMyRec()
            End If

        Catch ex As Exception
            MsgBox("Ошибка создания заявки на портале " + ex.Message, MsgBoxStyle.Critical, "Внимание!")
            CreateProjectDetailConfirmrequest = False
            Exit Function
        End Try

        CreateProjectDetailConfirmrequest = True
    End Function
End Class