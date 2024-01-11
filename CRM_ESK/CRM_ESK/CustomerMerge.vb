Public Class CustomerMerge
    Public SrcForm As String

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна без объединения клиентов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyResult = 0
        Me.Close()
    End Sub

    Private Sub CustomerMerge_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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

    Private Sub CustomerMerge_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список покупателей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        DataLoading()
        CheckButtons()
    End Sub

    Private Sub DataLoading()
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
        MySQLStr = MySQLStr & "tbl_RexelCustomerGroup.RussianName AS CustomerGroup, tbl_RexelEndMarkets.RussianName AS EndMarket "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Companies WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_RexelCustomerGroup ON tbl_CRM_Companies.RCGCode = tbl_RexelCustomerGroup.RCGCode LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_RexelEndMarkets ON tbl_CRM_Companies.EMCode = tbl_RexelEndMarkets.EMCode "
        If SrcForm.Equals("MainForm") Then
            MySQLStr = MySQLStr & "WHERE (tbl_CRM_Companies.CompanyID <> N'" & Declarations.MyClientID & "') "
        Else
            MySQLStr = MySQLStr & "WHERE (tbl_CRM_Companies.CompanyID <> N'" & Trim(MyCustomerSelect.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        End If

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

        If SrcForm.Equals("MainForm") Then
            Label2.Text = MainForm.SfDataGrid7.SelectedItem.GetItem("CompanyName").ToString()
        Else
            Label2.Text = Trim(MyCustomerSelect.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString())
        End If


    End Sub

    Private Sub CheckButtons()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выставление состояния кнопок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button4.Enabled = False
        Else
            Button4.Enabled = True
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна с выбором покупателя
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Declarations.MyNewClientID = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString())
        MergeCustomers()
        If SrcForm.Equals("CustomerSelect") Then
            MyCustomerSelect.DataLoading()
            MyCustomerSelect.CheckButtons()
        End If
        Declarations.MyResult = 1
        Me.Close()
    End Sub

    Private Sub MergeCustomers()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Собственно объединение
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        '----Сначала перебрасываем контакты
        MySQLStr = "UPDATE tbl_CRM_Contacts "
        If SrcForm.Equals("MainForm") Then
            MySQLStr = MySQLStr & "SET CompanyID = '" & Declarations.MyNewClientID & "' "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Declarations.MyClientID & "') "
        Else
            MySQLStr = MySQLStr & "SET CompanyID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "' "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(MyCustomerSelect.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '----После контактов перебрасываем действия
        MySQLStr = "UPDATE tbl_CRM_Events "
        If SrcForm.Equals("MainForm") Then
            MySQLStr = MySQLStr & "SET CompanyID = '" & Declarations.MyNewClientID & "' "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Declarations.MyClientID & "') "
        Else
            MySQLStr = MySQLStr & "SET CompanyID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "' "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(MyCustomerSelect.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '----Также перебрасываем проекты
        MySQLStr = "UPDATE tbl_CRM_Projects "
        If SrcForm.Equals("MainForm") Then
            MySQLStr = MySQLStr & "SET CompanyID = '" & Declarations.MyNewClientID & "' "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Declarations.MyClientID & "') "
        Else
            MySQLStr = MySQLStr & "SET CompanyID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "' "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(MyCustomerSelect.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '----Ну и удаляем клиента
        MySQLStr = "DELETE FROM tbl_CRM_Companies "
        If SrcForm.Equals("MainForm") Then
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Declarations.MyClientID & "')"
        Else
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(MyCustomerSelect.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "')"
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '----Обновляем окно действий
        If SrcForm.Equals("CustomerSelect") Then
            Select Case MainForm.TabControl1.SelectedTab.Text
                Case "План на месяц"
                    MainForm.Button9_Click_Func()
                Case "Список действий"
                    MainForm.Button5_Click_Func()

                Case Else
            End Select
        End If
    End Sub
End Class