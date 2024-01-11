Public Class CustomerSelectList

    Private Sub CustomerSelectList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// При запуске загружаем список покупателей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If Trim(MyCustomerSelect.TextBox1.Text) = "" Then
            '----В первое окно условие не введено - считаем, что во второе введено
            MySQLStr = "SELECT tbl_CRM_Companies.CompanyID, tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName, "
            MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyAddress, tbl_CRM_Companies.CompanyPhone, tbl_CRM_Companies.CompanyEMail, "
            MySQLStr = MySQLStr & "tbl_RexelCustomerGroup.RussianName AS CustomerGroup, tbl_RexelEndMarkets.RussianName AS EndMarket, "
            MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies_Ext.IsIKA, N'') AS IsIKA "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Companies WITH(NOLOCK) LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_RexelCustomerGroup ON tbl_CRM_Companies.RCGCode = tbl_RexelCustomerGroup.RCGCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_RexelEndMarkets ON tbl_CRM_Companies.EMCode = tbl_RexelEndMarkets.EMCode LEFT OUTER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies_Ext ON tbl_CRM_Companies.CompanyID = tbl_CRM_Companies_Ext.CompanyID "
            MySQLStr = MySQLStr & "WHERE (Upper(tbl_CRM_Companies.ScalaCustomerCode) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyName) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyAddress) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyPhone) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%') OR "
            MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyEMail) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%') "
            MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName "
        Else
            '----В первое окно условие введено
            If Trim(MyCustomerSelect.TextBox2.Text) = "" Then
                '----Во второе окно условие не введено
                MySQLStr = "SELECT tbl_CRM_Companies.CompanyID, tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName, "
                MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyAddress, tbl_CRM_Companies.CompanyPhone, tbl_CRM_Companies.CompanyEMail, "
                MySQLStr = MySQLStr & "tbl_RexelCustomerGroup.RussianName AS CustomerGroup, tbl_RexelEndMarkets.RussianName AS EndMarket, "
                MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies_Ext.IsIKA, N'') AS IsIKA "
                MySQLStr = MySQLStr & "FROM tbl_CRM_Companies WITH(NOLOCK) LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_RexelCustomerGroup ON tbl_CRM_Companies.RCGCode = tbl_RexelCustomerGroup.RCGCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_RexelEndMarkets ON tbl_CRM_Companies.EMCode = tbl_RexelEndMarkets.EMCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_Companies_Ext ON tbl_CRM_Companies.CompanyID = tbl_CRM_Companies_Ext.CompanyID "
                MySQLStr = MySQLStr & "WHERE (Upper(tbl_CRM_Companies.ScalaCustomerCode) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyName) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyAddress) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyPhone) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') OR "
                MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyEMail) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') "
                MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName "
            Else
                '----Условия введены в оба окна
                MySQLStr = "SELECT tbl_CRM_Companies.CompanyID, tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName, "
                MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyAddress, tbl_CRM_Companies.CompanyPhone, tbl_CRM_Companies.CompanyEMail, "
                MySQLStr = MySQLStr & "tbl_RexelCustomerGroup.RussianName AS CustomerGroup, tbl_RexelEndMarkets.RussianName AS EndMarket, "
                MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Companies_Ext.IsIKA, N'') AS IsIKA "
                MySQLStr = MySQLStr & "FROM tbl_CRM_Companies WITH(NOLOCK) LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_RexelCustomerGroup ON tbl_CRM_Companies.RCGCode = tbl_RexelCustomerGroup.RCGCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_RexelEndMarkets ON tbl_CRM_Companies.EMCode = tbl_RexelEndMarkets.EMCode LEFT OUTER JOIN "
                MySQLStr = MySQLStr & "tbl_CRM_Companies_Ext ON tbl_CRM_Companies.CompanyID = tbl_CRM_Companies_Ext.CompanyID "
                MySQLStr = MySQLStr & "WHERE ((Upper(tbl_CRM_Companies.ScalaCustomerCode) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.ScalaCustomerCode) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((Upper(tbl_CRM_Companies.CompanyName) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyName) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((Upper(tbl_CRM_Companies.CompanyAddress) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyAddress) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((Upper(tbl_CRM_Companies.CompanyPhone) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyPhone) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%')) OR "
                MySQLStr = MySQLStr & "((Upper(tbl_CRM_Companies.CompanyEMail) LIKE N'%" & UCase(MyCustomerSelect.TextBox1.Text) & "%') AND "
                MySQLStr = MySQLStr & "(Upper(tbl_CRM_Companies.CompanyEMail) LIKE N'%" & UCase(MyCustomerSelect.TextBox2.Text) & "%')) "
                MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName "
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

        If DataGridView1.Rows.Count > 0 Then
            Button4.Enabled = True
        Else
            Button4.Enabled = False
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без выбора покупателя
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором покупателя
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CustomerSelect()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором покупателя
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        CustomerSelect()
    End Sub

    Private Sub CustomerSelect()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с выбором покупателя
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        For i As Integer = 0 To MyCustomerSelect.DataGridView1.Rows.Count - 1
            If Trim(MyCustomerSelect.DataGridView1.Item(0, i).Value.ToString) = Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) Then
                MyCustomerSelect.DataGridView1.CurrentCell = MyCustomerSelect.DataGridView1.Item(1, i)
                Me.Close()
                Exit Sub
            End If
        Next
        Me.Close()
    End Sub
End Class