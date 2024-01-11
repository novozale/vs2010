Public Class SalesOrderList

    Private Sub SalesOrderList_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub SalesOrderList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна - загружаем исходную информацию
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        DataPreparation()
        ReCalculationSumm()
        ChangeButtonsStatus()
    End Sub

    Private Sub DataPreparation()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование списка заказов
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                   'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        MySQLStr = "SELECT tbl_CRM_Orders.ID, tbl_CRM_Orders.EventID, tbl_CRM_Orders.OrderNum, View_1.CustomerN, View_1.Customer, "
        MySQLStr = MySQLStr & "View_1.OrderDate, View_1.SalesmanN, View_1.Salesman, View_1.Summ "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Orders WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT OR010300.OR01001 AS OrderN, OR010300.OR01003 AS CustomerN, ISNULL(SL010300.SL01002, N'') AS Customer, "
        MySQLStr = MySQLStr & "OR010300.OR01015 AS OrderDate, OR010300.OR01019 AS SalesmanN, ISNULL(ST010300.ST01002, N'') AS Salesman, "
        MySQLStr = MySQLStr & "OR010300.OR01024 * SYCH0100.SYCH006 AS Summ "
        MySQLStr = MySQLStr & "FROM OR010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCH0100 ON OR010300.OR01028 = SYCH0100.SYCH001 AND OR010300.OR01015 >= SYCH0100.SYCH004 AND "
        MySQLStr = MySQLStr & "OR010300.OR01015 < SYCH0100.SYCH005 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "ST010300 ON OR010300.OR01019 = ST010300.ST01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 ON OR010300.OR01003 = SL010300.SL01001 "
        MySQLStr = MySQLStr & "UNION ALL "
        MySQLStr = MySQLStr & "SELECT OR200300.OR20001 AS OrderN, OR200300.OR20003 AS CustomerN, ISNULL(SL010300_1.SL01002, N'') AS Customer, "
        MySQLStr = MySQLStr & "OR200300.OR20015 AS OrderDate, OR200300.OR20019 AS SalesmanN, ISNULL(ST010300_1.ST01002, N'') AS Salesman, "
        MySQLStr = MySQLStr & "OR200300.OR20024 * SYCH0100_1.SYCH006 AS Summ "
        MySQLStr = MySQLStr & "FROM OR200300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "SYCH0100 AS SYCH0100_1 ON OR200300.OR20028 = SYCH0100_1.SYCH001 AND OR200300.OR20015 >= SYCH0100_1.SYCH004 AND "
        MySQLStr = MySQLStr & "OR200300.OR20015 < SYCH0100_1.SYCH005 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "ST010300 AS ST010300_1 ON OR200300.OR20019 = ST010300_1.ST01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "SL010300 AS SL010300_1 ON OR200300.OR20003 = SL010300_1.SL01001) AS View_1 ON "
        MySQLStr = MySQLStr & "tbl_CRM_Orders.OrderNum = View_1.OrderN "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Orders.EventID = '" & Declarations.MyEventID & "') "
        MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Orders.OrderNum "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        DataGridView1.Columns(0).HeaderText = "ID"
        DataGridView1.Columns(0).Width = 0
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "ActionID"
        DataGridView1.Columns(1).Width = 0
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).HeaderText = "N заказа на продажу"
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).HeaderText = "Код клиента"
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).HeaderText = "Название клиента"
        DataGridView1.Columns(4).Width = 200
        DataGridView1.Columns(5).HeaderText = "Дата заказа"
        DataGridView1.Columns(5).Width = 120
        DataGridView1.Columns(6).HeaderText = "Код продавца"
        DataGridView1.Columns(6).Width = 80
        DataGridView1.Columns(7).HeaderText = "Продавец"
        DataGridView1.Columns(7).Width = 200
        DataGridView1.Columns(8).HeaderText = "Сумма заказа"
        DataGridView1.Columns(8).Width = 120
    End Sub

    Private Sub ChangeButtonsStatus()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния кнопок
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If DataGridView1.SelectedRows.Count = 0 Then
            Button2.Enabled = False
        Else
            Button2.Enabled = True
        End If

    End Sub

    Private Sub ReCalculationSumm()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// пересчет общей суммы заказов, относящихся к данному действию
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySumm As Double

        MySumm = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            MySumm = MySumm + IIf(DataGridView1.Item(8, i).Value.ToString = "", 0, DataGridView1.Item(8, i).Value)
        Next

        Label2.Text = MySumm
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из редактирования списка заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление заказа из списка заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_CRM_Orders "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        DataPreparation()
        ReCalculationSumm()
        ChangeButtonsStatus()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Добавление заказа в список заказов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyAddSalesOrder = New AddSalesOrder
        MyAddSalesOrder.ShowDialog()

        DataPreparation()
        ReCalculationSumm()
        ChangeButtonsStatus()
    End Sub
End Class