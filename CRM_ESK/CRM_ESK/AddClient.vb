Public Class AddClient
    Public StartParam As String
    Public SourceForm As String

    Private Sub AddClient_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub AddClient_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы и значений в форму
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для групп Rexel
        Dim MyDs As New DataSet
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    'для рынков Rexel
        Dim MyDs1 As New DataSet
        Dim MyGUID As Guid

        '---Группы клиентов Rexel
        MySQLStr = "SELECT RCGCode, RussianName "
        MySQLStr = MySQLStr & "FROM tbl_RexelCustomerGroup WITH(NOLOCK) "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT '0' AS RCGCode, '' AS RussianName "
        MySQLStr = MySQLStr & "ORDER BY RCGCode "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "RussianName" 'Это то что будет отображаться
            ComboBox1.ValueMember = "RCGCode"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---Рынки Rexel
        MySQLStr = "SELECT EMCode, RussianName "
        MySQLStr = MySQLStr & "FROM tbl_RexelEndMarkets WITH(NOLOCK) "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT '0' AS EMCode, '' AS RussianName "
        MySQLStr = MySQLStr & "ORDER BY EMCode"
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox2.DisplayMember = "RussianName" 'Это то что будет отображаться
            ComboBox2.ValueMember = "EMCode"   'это то что будет храниться
            ComboBox2.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----В зависимости от того - это новое действие или редактирование - читаем дополнительные данные и выставляем значения
        If StartParam = "Create" Then
            '---создаем запись о действии
            MyGUID = Guid.NewGuid
            Declarations.MyClientID = MyGUID.ToString
        Else
            MySQLStr = "SELECT CompanyID, ScalaCustomerCode, CompanyName, RCGCode, EMCode, CompanyAddress, CompanyPhone, CompanyEMail "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Companies WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Declarations.MyClientID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Данный клиент не существует, возможно, удален другим пользователем. Обновите данные в окне со списком клиентов.", MsgBoxStyle.Critical, "Внимание!")
                Me.Close()
            Else
                Declarations.MyRec.MoveFirst()
                TextBox6.Text = Declarations.MyRec.Fields("CompanyName").Value
                TextBox1.Text = Declarations.MyRec.Fields("CompanyAddress").Value
                TextBox2.Text = Declarations.MyRec.Fields("CompanyPhone").Value
                TextBox3.Text = Declarations.MyRec.Fields("CompanyEMail").Value
                If Declarations.MyRec.Fields("RCGCode").Value.ToString = "" Then
                    ComboBox1.SelectedValue = "0"
                Else
                    ComboBox1.SelectedValue = Declarations.MyRec.Fields("RCGCode").Value
                End If
                If Declarations.MyRec.Fields("EMCode").Value.ToString = "0" Then
                    ComboBox2.SelectedValue = "0"
                Else
                    ComboBox2.SelectedValue = Declarations.MyRec.Fields("EMCode").Value
                End If
                trycloseMyRec()
            End If
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

    Private Sub ComboBox1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox2.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(sender, True, True, True, False)
    End Sub

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(sender, True, True, True, False)
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If Trim(TextBox6.Text) = "" Then
            MsgBox("Поле ""Название"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox6.Select()
            Exit Function
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Поле ""Адрес"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox1.Select()
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""Телефон"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox2.Select()
            Exit Function
        End If

        If Trim(TextBox3.Text) = "" Then
            MsgBox("Поле ""E-mail"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox3.Select()
            Exit Function
        End If

        If ComboBox1.SelectedValue = "0" Then
            MsgBox("Поле ""Группа Rexel"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            ComboBox1.Select()
            Exit Function
        End If

        If ComboBox2.SelectedValue = "0" Then
            MsgBox("Поле ""Рынок Rexel"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            ComboBox2.Select()
            Exit Function
        End If

        '---Проверяем, что такого названия нет в БД
        MySQLStr = "SELECT COUNT(*) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Companies "
        MySQLStr = MySQLStr & "WHERE (UPPER(Ltrim(Rtrim(CompanyName))) = UPPER('" & Trim(TextBox6.Text) & "')) "
        MySQLStr = MySQLStr & "AND (CompanyID <> '" & Trim(Declarations.MyClientID) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            CheckDataFiling = True
        Else
            If Declarations.MyRec.Fields("CC").Value < 1 Then
                trycloseMyRec()
                CheckDataFiling = True
            Else
                trycloseMyRec()
                MsgBox("Компания с таким названием уже есть в базе данных (возможно, набранная в другом - верхнем или нижнем регистре). ", MsgBoxStyle.Critical, "Внимание")
                CheckDataFiling = False
                Exit Function
            End If

        End If

        CheckDataFiling = True

    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение информации о клиенте
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then
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

        '---создание записи в таблице tbl_CRM_Events
        MySQLStr = "INSERT INTO tbl_CRM_Companies "
        MySQLStr = MySQLStr & "(CompanyID, ScalaCustomerCode, CompanyName, RCGCode, EMCode, CompanyAddress, CompanyPhone, CompanyEMail, CreationDate) "
        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyClientID & "', "
        MySQLStr = MySQLStr & "NULL, "
        MySQLStr = MySQLStr & "N'" & TextBox6.Text & "', "
        MySQLStr = MySQLStr & "N'" & ComboBox1.SelectedValue & "', "
        MySQLStr = MySQLStr & "N'" & ComboBox2.SelectedValue & "', "
        MySQLStr = MySQLStr & "N'" & TextBox1.Text & "', "
        MySQLStr = MySQLStr & "N'" & TextBox2.Text & "', "
        MySQLStr = MySQLStr & "N'" & TextBox3.Text & "', "
        MySQLStr = MySQLStr & "GetDate()) "
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

        '---Редактирование существующей записи в таблице tbl_CRM_Events
        MySQLStr = "UPDATE tbl_CRM_Companies "
        MySQLStr = MySQLStr & "SET CompanyName = N'" & TextBox6.Text & "', "
        MySQLStr = MySQLStr & "RCGCode = N'" & ComboBox1.SelectedValue & "', "
        MySQLStr = MySQLStr & "EMCode = N'" & ComboBox2.SelectedValue & "', "
        MySQLStr = MySQLStr & "CompanyAddress = N'" & TextBox1.Text & "', "
        MySQLStr = MySQLStr & "CompanyPhone = N'" & TextBox2.Text & "', "
        MySQLStr = MySQLStr & "CompanyEMail = N'" & TextBox3.Text & "' "
        MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Declarations.MyClientID & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class