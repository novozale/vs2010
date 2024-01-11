Public Class AddContact
    Public StartParam As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения информации
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub AddContact_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub AddContact_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы и значений в форму
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyGUID As Guid
        Dim MySQLStr As String

        '----В зависимости от того - это новое действие или редактирование - читаем дополнительные данные и выставляем значения
        If StartParam = "Create" Then
            '---создаем запись о действии
            MyGUID = Guid.NewGuid
            Declarations.MyContactID = MyGUID.ToString
        Else
            MySQLStr = "SELECT ContactID, CompanyID, ContactName, ContactPhone, ContactEMail, FromScala, ISNULL(Comments,'') AS Comments "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Contacts WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (ContactID = '" & Declarations.MyContactID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Данный контакт не существует, возможно, удален другим пользователем. Обновите данные в окне со списком контактов.", MsgBoxStyle.Critical, "Внимание!")
                Me.Close()
            Else
                Declarations.MyRec.MoveFirst()
                TextBox6.Text = Declarations.MyRec.Fields("ContactName").Value
                TextBox1.Text = Declarations.MyRec.Fields("ContactPhone").Value
                TextBox2.Text = Declarations.MyRec.Fields("ContactEMail").Value
                TextBox3.Text = Declarations.MyRec.Fields("Comments").Value
                trycloseMyRec()
            End If
        End If
        Label2.Text = MyAddEvent.TextBox6.Text
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

    Private Sub Button2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button2.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с сохранением данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            SaveData()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с сохранением данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SaveData()
    End Sub

    Private Sub SaveData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с сохранением данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then
            If StartParam = "Create" Then
                SaveNewData()
            Else
                UpdateData()
            End If
            Me.Close()
        End If
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox6.Text) = "" Then
            MsgBox("Поле ""Ф.И.О"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox6.Select()
            Exit Function
        End If

        If Trim(TextBox1.Text) = "" Then
            MsgBox("Поле ""Телефон"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox1.Select()
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" Then
            MsgBox("Поле ""E-Mail"" должно быть заполнено. ", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox2.Select()
            Exit Function
        End If

        CheckDataFiling = True
    End Function

    Private Sub SaveNewData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// сохранение данных в случае создания новой записи
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---создание записи в таблице tbl_CRM_Contacts
        MySQLStr = "INSERT INTO tbl_CRM_Contacts "
        MySQLStr = MySQLStr & "(ContactID, CompanyID, ContactName, ContactPhone, ContactEMail, FromScala, CreationDate, Comments) "
        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyContactID & "', "
        MySQLStr = MySQLStr & "'" & Declarations.MyClientID & "', "
        MySQLStr = MySQLStr & "N'" & TextBox6.Text & "', "
        MySQLStr = MySQLStr & "N'" & TextBox1.Text & "', "
        MySQLStr = MySQLStr & "N'" & TextBox2.Text & "', "
        MySQLStr = MySQLStr & "0, "
        MySQLStr = MySQLStr & "GetDate(), "
        MySQLStr = MySQLStr & "N'" & TextBox3.Text & "') "
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
        MySQLStr = "Update tbl_CRM_Contacts "
        MySQLStr = MySQLStr & "SET ContactName = N'" & TextBox6.Text & "', "
        MySQLStr = MySQLStr & "ContactPhone = N'" & TextBox1.Text & "', "
        MySQLStr = MySQLStr & "ContactEMail = N'" & TextBox2.Text & "', "
        MySQLStr = MySQLStr & "Comments = N'" & TextBox3.Text & "' "
        MySQLStr = MySQLStr & "WHERE (ContactID = '" & Declarations.MyContactID & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class