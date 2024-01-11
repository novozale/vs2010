Public Class ContactSelect

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ContactSelect_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub ContactSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список контактов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Label2.Text = MyAddEvent.TextBox6.Text
        LoadData()
        CheckButtons()
    End Sub

    Private Sub LoadData()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка контактов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT ContactID, CompanyID, ContactName, ContactPhone, ContactEMail, ISNULL(Comments,'') AS Comments, CASE WHEN FromScala = 0 THEN '' ELSE 'X' END AS FromScala "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Contacts WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Declarations.MyClientID & "') "
        MySQLStr = MySQLStr & "ORDER BY ContactName "

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
        DataGridView1.Columns(2).HeaderText = "Контактное лицо"
        DataGridView1.Columns(2).Width = 237
        DataGridView1.Columns(3).HeaderText = "Телефон"
        DataGridView1.Columns(3).Width = 150
        DataGridView1.Columns(4).HeaderText = "E-Mail"
        DataGridView1.Columns(4).Width = 150
        DataGridView1.Columns(5).HeaderText = "Комментарий"
        DataGridView1.Columns(5).Width = 150
        DataGridView1.Columns(6).HeaderText = "Из Scala"
        DataGridView1.Columns(6).Width = 40

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
            Button4.Enabled = True
            If Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(6).Value.ToString()) = "" Then
                Button8.Enabled = True
                Button9.Enabled = True
            Else
                Button8.Enabled = False
                Button9.Enabled = False
            End If
        End If
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором контакта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count <> 0 Then
            ContactSelect()
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

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором контакта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ContactSelect()
    End Sub

    Private Sub ContactSelect()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором контакта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddEvent.TextBox7.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString()) + " " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString()) + " " + Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(4).Value.ToString())
        Declarations.MyContactID = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())

        Me.Close()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание нового контакта
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        MyAddContact = New AddContact
        MyAddContact.StartParam = "Create"
        MyAddContact.ShowDialog()
        LoadData()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyContactID Then
                DataGridView1.CurrentCell = DataGridView1.Item(2, i)
            End If
        Next
        '---проверка состояния кнопок
        CheckButtons()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование контакта
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Declarations.MyContactID = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MyAddContact = New AddContact
        MyAddContact.StartParam = "Edit"
        MyAddContact.ShowDialog()
        LoadData()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyContactID Then
                DataGridView1.CurrentCell = DataGridView1.Item(2, i)
            End If
        Next
        '---проверка состояния кнопок
        CheckButtons()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление контакта
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---Проверка - можно ли удалять, может быть есть ссылки на него
        MySQLStr = "SELECT COUNT(ContactID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (ContactID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            '---можно удалять
            trycloseMyRec()
            '---Удаление контакта
            MySQLStr = "DELETE FROM tbl_CRM_Contacts "
            MySQLStr = MySQLStr & "WHERE (ContactID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            LoadData()
            CheckButtons()
        Else
            trycloseMyRec()
            MsgBox("Данный контакт нельзя удалять, так как на него есть ссылки в таблице действий. Удалить такой контакт можно только удалив сначала все действия с этим контактом.", MsgBoxStyle.Critical, "Внимание!")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление списка контактов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Label2.Text = MyAddEvent.TextBox6.Text
        LoadData()
        CheckButtons()
    End Sub
End Class