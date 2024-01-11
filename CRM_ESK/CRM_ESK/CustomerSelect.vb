Public Class CustomerSelect
    Public SourceForm As String

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна без выбора покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub CustomerSelect_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub CustomerSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список покупателей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        DataLoading()
        CheckButtons()
    End Sub

    Public Sub DataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в форму
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        MySQLStr = "SELECT tbl_CRM_Companies.CompanyID, tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName, "
        MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyAddress, tbl_CRM_Companies.CompanyPhone, tbl_CRM_Companies.CompanyEMail,"
        MySQLStr = MySQLStr & "tbl_RexelCustomerGroup.RussianName AS CustomerGroup, tbl_RexelEndMarkets.RussianName AS EndMarket, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies_Ext.IsIKA, N'') AS IsIKA "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Companies WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_RexelCustomerGroup ON tbl_CRM_Companies.RCGCode = tbl_RexelCustomerGroup.RCGCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_RexelEndMarkets ON tbl_CRM_Companies.EMCode = tbl_RexelEndMarkets.EMCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & " tbl_CRM_Companies_Ext ON tbl_CRM_Companies.CompanyID = tbl_CRM_Companies_Ext.CompanyID "
        MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName "

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
        DataGridView1.Columns(1).HeaderText = "Код в Scala"
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).HeaderText = "Название клиента"
        DataGridView1.Columns(2).Width = 170
        DataGridView1.Columns(3).HeaderText = "Адрес клиента"
        DataGridView1.Columns(3).Width = 300
        DataGridView1.Columns(4).HeaderText = "Телефон"
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).HeaderText = "E-Mail"
        DataGridView1.Columns(5).Width = 100
        DataGridView1.Columns(6).HeaderText = "Группа Rexel"
        DataGridView1.Columns(6).Width = 250
        DataGridView1.Columns(7).HeaderText = "Рынок Rexel"
        DataGridView1.Columns(7).Width = 180
        DataGridView1.Columns(8).HeaderText = "Вид КА"
        DataGridView1.Columns(8).Width = 120

    End Sub

    Public Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление состояния кнопок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button4.Enabled = False
            Button8.Enabled = False
            Button9.Enabled = False
            Button11.Enabled = False
            Button12.Enabled = False
        Else
            Button4.Enabled = True

            If Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(1).Value.ToString()) = "" Then
                Button8.Enabled = True
                Button9.Enabled = True
                Button11.Enabled = True
                Button12.Enabled = True
            Else
                Button8.Enabled = False
                Button9.Enabled = False
                Button11.Enabled = False
                Button12.Enabled = False
            End If
        End If
    End Sub

    Private Sub DataGridView1_ColumnHeaderMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.ColumnHeaderMouseClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Щелчок по заголовку таблицы 
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Button6.Text = "Подсветить все"
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CustomerSelect()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка состояния кнопок при изменении выделения  
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CheckButtons()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CustomerSelect()
    End Sub

    Private Sub CustomerSelect()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If SourceForm = "AddEvent" Then
            MyAddEvent.TextBox6.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString())
            Declarations.MyClientID = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        Else    '---AddProject
            MyAddProject.TextBox13.Text = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString())
            Declarations.MyClientID = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        End If
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск первого подходящего по критерию покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(5, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(5, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                    Exit Sub
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск следующего подходящего по критерию покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    MyRez = MsgBox("Поиск дошел до конца списка. Начать сначала?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(5, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(5, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                        Exit Sub
                    End If
                End If
            Next i
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Подсвечивание всех подходящих по критерию покупателей
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Windows.Forms.Cursor.Current = Cursors.WaitCursor
        If Button6.Text = "Подсветить все" Then
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(2, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(4, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(5, i).Value.ToString)), UCase(Trim(TextBox1.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(5, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                Else
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
                End If
            Next
            Button6.Text = "Снять выделение"
        Else
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
            Next
            Button6.Text = "Подсветить все"
        End If
        Windows.Forms.Cursor.Current = Cursors.Default
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор всех подходящих по критерию покупателей в отдельное окно
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox1.Select()
        Else
            MyCustomerSelectList = New CustomerSelectList
            MyCustomerSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание нового клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        AddNewClient()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Редактирование клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddClient = New AddClient
        MyAddClient.StartParam = "Edit"
        MyAddClient.SourceForm = "CustomerSelect"
        Declarations.MyClientID = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MyAddClient.ShowDialog()
        DataLoading()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyClientID Then
                DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                Exit For
            End If
        Next
        CheckButtons()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---Проверка - можно ли удалять, может быть есть ссылки на него
        MySQLStr = "SELECT COUNT(CompanyID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            '---можно удалять
            trycloseMyRec()
            '---Удаление контактов
            MySQLStr = "DELETE FROM tbl_CRM_Contacts "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '---Удаление дополнительной информации
            MySQLStr = "DELETE FROM tbl_CRM_Companies_Ext "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
            '---Удаление самой записи
            MySQLStr = "DELETE FROM tbl_CRM_Companies "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "')"
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            DataLoading()
            CheckButtons()
        Else
            trycloseMyRec()
            MsgBox("Данную компанию нельзя удалять, так как на нее есть ссылки в таблице действий. Удалить такую компанию можно или удалив сначала все действия по этой компаниии, или объединив эту компанию с другой.", MsgBoxStyle.Critical, "Внимание!")
        End If
    End Sub

    Private Sub AddNewClient()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание нового клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddClient = New AddClient
        MyAddClient.StartParam = "Create"
        MyAddClient.SourceForm = "CustomerSelect"
        MyAddClient.ShowDialog()
        DataLoading()
        '---текущей строкой сделать созданную
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyClientID Then
                DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                Exit For
            End If
        Next
        CheckButtons()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных по клиентам
        '//
        '////////////////////////////////////////////////////////////////////////////////

        DataLoading()
        CheckButtons()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Объединение клиентов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCustomerMerge = New CustomerMerge
        MyCustomerMerge.SrcForm = "CustomerSelect"
        MyCustomerMerge.ShowDialog()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Внесение дополнительной информации по клиенту
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCustomerExtInfo = New CustomerExtInfo
        Declarations.MyClientID = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MyCustomerExtInfo.ShowDialog()
        DataLoading()
        '---текущей строкой сделать редактируемую
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If Trim(DataGridView1.Item(0, i).Value.ToString) = Declarations.MyClientID Then
                DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                Exit For
            End If
        Next
        CheckButtons()
    End Sub
End Class