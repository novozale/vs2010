Public Class SupplierSelectList

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без выбора поставщика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub SupplierSelectList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором поставщика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        SupplierSelect()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором поставщика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        SupplierSelect()
    End Sub

    Private Sub SupplierSelect()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором поставщика
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 0 To MySupplierSelect.DataGridView1.Rows.Count - 1
            If Trim(MySupplierSelect.DataGridView1.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MySupplierSelect.DataGridView1.CurrentCell = MySupplierSelect.DataGridView1.Item(0, i)
                Me.Close()
                Exit Sub
            End If
        Next
        Me.Close()
    End Sub

    Private Sub SupplierSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список поставщиков
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If Trim(MySupplierSelect.TextBox1.Text) = "" Then
            '----В первое окно условие не введено - считаем, что во второе введено
            MySQLStr = "SELECT PL01001, PL01002, PL01003 + PL01004 + PL01005 AS PL01003, PL01025 "
            MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (UPPER(PL01001) LIKE N'%" & UCase(MySupplierSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(UPPER(PL01002) LIKE N'%" & UCase(MySupplierSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(UPPER(PL01003 + PL01004 + PL01005) LIKE N'%" & UCase(MySupplierSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(UPPER(PL01025) LIKE N'%" & UCase(MySupplierSelect.TextBox2.Text) & "%') "
            MySQLStr = MySQLStr & "ORDER BY PL01002"
        Else
            '----В первое окно условие введено
            If Trim(MySupplierSelect.TextBox2.Text) = "" Then
                '----Во второе окно условие введено
                MySQLStr = "SELECT PL01001, PL01002, PL01003 + PL01004 + PL01005 AS PL01003, PL01025 "
                MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (UPPER(PL01001) LIKE N'%" & UCase(MySupplierSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(UPPER(PL01002) LIKE N'%" & UCase(MySupplierSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(UPPER(PL01003 + PL01004 + PL01005) LIKE N'%" & UCase(MySupplierSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(UPPER(PL01025) LIKE N'%" & UCase(MySupplierSelect.TextBox1.Text) & "%') "
                MySQLStr = MySQLStr & "ORDER BY PL01002"
            Else
                '----Условия введены в оба окна
                MySQLStr = "SELECT PL01001, PL01002, PL01003 + PL01004 + PL01005 AS PL01003, PL01025 "
                MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE ((UPPER(PL01001) LIKE N'%" & UCase(MySupplierSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(PL01001) LIKE N'%" & UCase(MySupplierSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((UPPER(PL01002) LIKE N'%" & UCase(MySupplierSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(PL01002) LIKE N'%" & UCase(MySupplierSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((UPPER(PL01003 + PL01004 + PL01005) LIKE N'%" & UCase(MySupplierSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(PL01003 + PL01004 + PL01005) LIKE N'%" & UCase(MySupplierSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((UPPER(PL01025) LIKE N'%" & UCase(MySupplierSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(UPPER(PL01025) LIKE N'%" & UCase(MySupplierSelect.TextBox2.Text) & "%')) "
                MySQLStr = MySQLStr & "ORDER BY PL01002 "
            End If

        End If

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        DataGridView1.Columns(0).HeaderText = "Код поставщика"
        DataGridView1.Columns(0).Width = 90
        DataGridView1.Columns(1).HeaderText = "Имя поставщика"
        DataGridView1.Columns(1).Width = 140
        DataGridView1.Columns(2).HeaderText = "Адрес поставщика"
        DataGridView1.Columns(3).HeaderText = "ИНН покупателя"
        DataGridView1.Columns(3).Width = 130

        If DataGridView1.Rows.Count > 0 Then
            Button4.Enabled = True
        Else
            Button4.Enabled = False
        End If
    End Sub
End Class