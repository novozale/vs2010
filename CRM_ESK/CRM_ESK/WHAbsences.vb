Public Class WHAbsences

    Private Sub WHAbsences_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна - загружаем исходную информацию
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                   'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        '---список складов и склад по умолчанию
        MySQLStr = "SELECT SC23001, SC23001 + ' ' + SC23002 AS SC23002 "
        MySQLStr = MySQLStr & "FROM SC230300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC23006 = N'1') "
        MySQLStr = MySQLStr & "ORDER BY SC23001 "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "SC23002" 'Это то что будет отображаться
            ComboBox1.ValueMember = "SC23001"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        MySQLStr = "SELECT DISTINCT WHCode "
        MySQLStr = MySQLStr & "FROM tbl_CRM_CCOwners WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (CCSubord = N'" & Declarations.CC & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
        Else
            ComboBox1.SelectedValue = Declarations.MyRec.Fields("WHCode").Value
            trycloseMyRec()
        End If

        DataPreparation()
        ListPreparation()
        ChangeButtonsStatus()
    End Sub

    Public Function DataPreparation()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование списка продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                   'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter
        Dim MyDs As New DataSet

        MySQLStr = "SELECT SC010300.SC01001 AS ID, "
        MySQLStr = MySQLStr & "SC010300.SC01002 + ' ' + SC010300.SC01003 AS Name, "
        MySQLStr = MySQLStr & "ROUND(ISNULL(t2.SC39005, 0), 2) AS Price, "
        MySQLStr = MySQLStr & "SC010300.SC01060 AS SuppID, "
        MySQLStr = MySQLStr & "View_1.txt AS UnitName, "
        MySQLStr = MySQLStr & "SC010300.SC01042 - SC010300.SC01043 - SC010300.SC01044 AS AvailableQty, "
        MySQLStr = MySQLStr & "SC010300.SC01058 AS SuppCode, "
        MySQLStr = MySQLStr & "ISNULL(PL010300.PL01002, N'') + ' ' + ISNULL(PL010300.PL01003, N'') AS SuppName "
        MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT 0 AS Expr1, SC09002 AS txt "
        MySQLStr = MySQLStr & "FROM SC090300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 1 AS Expr1, SC09003 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_40 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 2 AS Expr1, SC09004 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_39 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 3 AS Expr1, SC09005 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_38 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 4 AS Expr1, SC09006 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_37 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 5 AS Expr1, SC09007 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_36 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 6 AS Expr1, SC09008 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_35 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 7 AS Expr1, SC09009 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_34 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 8 AS Expr1, SC09010 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_33 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 9 AS Expr1, SC09011 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_32 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 10 AS Expr1, SC09012 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_31 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 11 AS Expr1, SC09013 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_30 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 12 AS Expr1, SC09014 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_29 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 13 AS Expr1, SC09015 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_28 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 14 AS Expr1, SC09016 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_27 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 15 AS Expr1, SC09017 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_26 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 16 AS Expr1, SC09018 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_25 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 17 AS Expr1, SC09019 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_24 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 18 AS Expr1, SC09020 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_23 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 19 AS Expr1, SC09021 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_22 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 20 AS Expr1, SC09022 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_21 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 21 AS Expr1, SC09023 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_20 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 22 AS Expr1, SC09024 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_19 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 23 AS Expr1, SC09025 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_18 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 24 AS Expr1, SC09026 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_17 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 25 AS Expr1, SC09027 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_16 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 26 AS Expr1, SC09028 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_15 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 27 AS Expr1, SC09029 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_14 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 28 AS Expr1, SC09030 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_13 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 29 AS Expr1, SC09031 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_12 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 30 AS Expr1, SC09032 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_11 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 31 AS Expr1, SC09033 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_10 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 32 AS Expr1, SC09034 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_9 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 33 AS Expr1, SC09035 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_8 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 34 AS Expr1, SC09036 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_7 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 35 AS Expr1, SC09037 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_6 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 36 AS Expr1, SC09038 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_5 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 37 AS Expr1, SC09039 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_4 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 38 AS Expr1, SC09040 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_3 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 39 AS Expr1, SC09041 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_2 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC09001 = 'RUS') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 40 AS Expr1, SC09042 "
        MySQLStr = MySQLStr & "FROM SC090300 AS SC090300_1 WITH(NOLOCK)) AS View_1 ON "
        MySQLStr = MySQLStr & "SC010300.SC01135 = View_1.Expr1 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "PL010300 ON SC010300.SC01058 = PL010300.PL01001 LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT SC39001, SC39005 "
        MySQLStr = MySQLStr & "FROM SC390300 WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (SC39002 = N'00')) AS t2 ON "
        MySQLStr = MySQLStr & "SC010300.SC01001 = t2.SC39001 "
        MySQLStr = MySQLStr & "WHERE (SC010300.SC01001 <> N'00000000') AND "
        MySQLStr = MySQLStr & "(LTRIM(RTRIM(SC010300.SC01066)) <> N'8') "
        If Trim(TextBox1.Text) = "" Then
        Else
            MySQLStr = MySQLStr & "AND (SC010300.SC01058 = N'" & Trim(TextBox1.Text) & "') "
        End If
        MySQLStr = MySQLStr & "ORDER BY dbo.SC010300.SC01001  "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        DataGridView1.Columns(0).HeaderText = "Код Scala"
        DataGridView1.Columns(0).Width = 110
        DataGridView1.Columns(1).HeaderText = "Имя продукта"
        DataGridView1.Columns(1).Width = 300
        DataGridView1.Columns(2).HeaderText = "Прайс"
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).HeaderText = "Код постав"
        DataGridView1.Columns(3).Width = 110
        DataGridView1.Columns(4).HeaderText = "Ед изм"
        DataGridView1.Columns(4).Width = 40
        DataGridView1.Columns(5).HeaderText = "Доступно для заказа на всех складах"
        DataGridView1.Columns(5).Width = 115
        DataGridView1.Columns(6).HeaderText = "Поставщик ID"
        DataGridView1.Columns(6).Width = 70
        DataGridView1.Columns(7).HeaderText = "Поставщик"
        DataGridView1.Columns(7).Width = 300

    End Function

    Public Function ListPreparation()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Формирование списка товаров с завышенной ценой
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLstr As String
        Dim MyAdapter1 As SqlClient.SqlDataAdapter
        Dim MyDs1 As New DataSet

        MySQLstr = "SELECT CONVERT(nvarchar(50), tbl_CRM_WHAbsences.ID) AS ID, "
        MySQLstr = MySQLstr & "ISNULL(SC230300.SC23002, N'') AS WH, "
        MySQLstr = MySQLstr & "tbl_CRM_WHAbsences.ScalaCode, "
        MySQLstr = MySQLstr & "ISNULL(tbl_CRM_WHAbsences.SupplierCode, tbl_CRM_WHAbsences.OtherSupplierCode) AS SupplierCode, "
        MySQLstr = MySQLstr & "CASE WHEN SC010300.SC01001 IS NULL THEN tbl_CRM_WHAbsences.OtherItem ELSE Ltrim(Rtrim(Ltrim(Rtrim(ISNULL(SC010300.SC01002, ''))) + ' ' + Ltrim(Rtrim(ISNULL(SC010300.SC01002, ''))))) END AS ItemName, "
        MySQLstr = MySQLstr & "tbl_CRM_WHAbsences.ScalaSupplierCode, "
        MySQLstr = MySQLstr & "CASE WHEN PL010300.PL01001 IS NULL THEN tbl_CRM_WHAbsences.OtherSupplier ELSE Ltrim(Rtrim(ISNULL(PL010300.PL01002, ''))) END AS ScalaSupplierName, "
        MySQLstr = MySQLstr & "tbl_CRM_WHAbsences.RequestedQTY, "
        MySQLstr = MySQLstr & "tbl_CRM_WHAbsences.AvailableQTY "
        MySQLstr = MySQLstr & "FROM tbl_CRM_WHAbsences WITH(NOLOCK) LEFT OUTER JOIN "
        MySQLstr = MySQLstr & "SC230300 ON tbl_CRM_WHAbsences.WH = SC230300.SC23001 LEFT OUTER JOIN "
        MySQLstr = MySQLstr & "PL010300 ON tbl_CRM_WHAbsences.ScalaSupplierCode = PL010300.PL01001 LEFT OUTER JOIN "
        MySQLstr = MySQLstr & "SC010300 ON tbl_CRM_WHAbsences.ScalaCode = SC010300.SC01001 "
        MySQLstr = MySQLstr & "WHERE (tbl_CRM_WHAbsences.EventID = '" & Declarations.MyEventID & "') "
        MySQLstr = MySQLstr & "ORDER BY tbl_CRM_WHAbsences.ScalaCode, SupplierCode "

        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLstr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            DataGridView2.DataSource = MyDs1.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        DataGridView2.Columns(0).HeaderText = "ID строки"
        DataGridView2.Columns(0).Width = 0
        DataGridView2.Columns(0).Visible = False
        DataGridView2.Columns(1).HeaderText = "Склад"
        DataGridView2.Columns(1).Width = 130
        DataGridView2.Columns(2).HeaderText = "Код Scala"
        DataGridView2.Columns(2).Width = 120
        DataGridView2.Columns(3).HeaderText = "Код товара поставщика"
        DataGridView2.Columns(3).Width = 120
        DataGridView2.Columns(4).HeaderText = "Имя продукта"
        DataGridView2.Columns(4).Width = 300
        DataGridView2.Columns(5).HeaderText = "Код поставщика"
        DataGridView2.Columns(5).Width = 130
        DataGridView2.Columns(6).HeaderText = "Поставщик"
        DataGridView2.Columns(6).Width = 200
        DataGridView2.Columns(7).HeaderText = "Количество, запрошенное клиентом"
        DataGridView2.Columns(7).Width = 130
        DataGridView2.Columns(8).HeaderText = "Доступно для заказа на всех складах"
        DataGridView2.Columns(8).Width = 130


    End Function

    Public Sub ChangeButtonsStatus()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// смена состояния кнопок 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If DataGridView2.SelectedRows.Count = 0 Then
            Button12.Enabled = False
        Else
            Button12.Enabled = True
        End If

        If DataGridView1.SelectedRows.Count = 0 Then
            Button11.Enabled = False
        Else
            Button11.Enabled = True
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление списка запасов 
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        DataPreparation()
        ChangeButtonsStatus()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна со списком поставщиков
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        MySupplierSelect = New SupplierSelect
        MySupplierSelect.MySrcWin = "WHAbsences"
        MySupplierSelect.ShowDialog()
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

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Введен код поставщика - Обновляем список продуктов
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        If TextBox1.Modified = True Then
            TextBox2.Text = ""
            TextBox3.Text = ""
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            DataPreparation()
            ChangeButtonsStatus()
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox1.Validating
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Введен код поставщика - находим и подписываем его название
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If TextBox1.Modified = True Then
            If Trim(TextBox1.Text) = "" Then
                Label3.Text = ""
            Else
                MySQLStr = "SELECT PL01002, PL01003 + ' ' + PL01004 + ' ' + PL01005 AS PL01003 "
                MySQLStr = MySQLStr & "FROM PL010300 WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (PL01001 = N'" & Trim(TextBox1.Text) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    trycloseMyRec()
                    MsgBox("Вы ввели неверный код поставщика. Введите корректный или воспользуйтесь поиском.", vbCritical, "Внимание!")
                    Label3.Text = ""
                    e.Cancel = True
                    Exit Sub
                Else
                    Label3.Text = Declarations.MyRec.Fields("PL01002").Value & " " & Declarations.MyRec.Fields("PL01003").Value
                    trycloseMyRec()
                End If
            End If
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск первого подходящего по критерию запаса
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск следующего подходящего по критерию запаса
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Object

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = DataGridView1.CurrentCellAddress.Y + 1 To DataGridView1.Rows.Count
                If i = DataGridView1.Rows.Count Then
                    MyRez = MsgBox("Поиск дошел до конца списка. Начать сначала?", MsgBoxStyle.YesNo, "Внимание!")
                    If MyRez = 6 Then
                        i = 0
                    Else
                        Windows.Forms.Cursor.Current = Cursors.Default
                        Exit Sub
                    End If
                End If
                If DataGridView1.Rows.Count = 0 Then
                Else
                    If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                        Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                        DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                        Windows.Forms.Cursor.Current = Cursors.Default
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
        '// Подсвечивание всех подходящих по критерию запасов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Button6.Text = "Подсветить все" Then
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(1, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 _
                    Or InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox2.Text))) <> 0 And InStr(UCase(Trim(DataGridView1.Item(3, i).Value.ToString)), UCase(Trim(TextBox3.Text))) <> 0 Then
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Yellow
                Else
                    DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
                End If
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "Снять выдел."
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.Empty
            Next
            Windows.Forms.Cursor.Current = Cursors.Default
            Button6.Text = "Подсветить все"
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

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Поиск первого запаса по коду с начала строки
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" Then
        Else
            Windows.Forms.Cursor.Current = Cursors.WaitCursor
            For i As Integer = 0 To DataGridView1.Rows.Count - 1
                If StrComp(UCase(Trim(TextBox2.Text)), Microsoft.VisualBasic.Left(UCase(Trim(DataGridView1.Item(0, i).Value.ToString)), Len(UCase(Trim(TextBox2.Text)))), 1) = 0 Then
                    DataGridView1.CurrentCell = DataGridView1.Item(0, i)
                    Windows.Forms.Cursor.Current = Cursors.Default
                    Exit Sub
                End If
            Next i
            Exit Sub
            Windows.Forms.Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из редактирования строк запасов с завышенным прайсом
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выбор всех подходящих по критерию запасов в отдельное окно
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox2.Text) = "" And Trim(TextBox3.Text) = "" Then
            MsgBox("Необходимо ввести критерий поиска", MsgBoxStyle.OkOnly, "Внимание!")
            TextBox2.Select()
        Else
            MyItemSelectList = New ItemSelectList
            MyItemSelectList.MySrcWin = "WHAbsences"
            MyItemSelectList.ShowDialog()
        End If
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// открытие окна для добавления записи о товаре с высоким прайсом
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyItemWHAbsences = New ItemWHAbsences
        MyItemWHAbsences.ShowDialog()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление записи о товаре с высоким прайсом
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_CRM_WHAbsences "
        MySQLStr = MySQLStr & "WHERE (ID = '" & Trim(Me.DataGridView2.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        ListPreparation()
        ChangeButtonsStatus()
    End Sub
End Class