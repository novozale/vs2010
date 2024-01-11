Imports System
Imports System.IO
Imports ADODB

Public Class AddEvent
    Public StartParam As String
    Public ActDate As DateTime
    Public CompanyID As String
    Public UserID As Integer

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения изменений
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckEventClose() = True Then
            CheckAddInfo_CancelAction()
            Declarations.MyResult = 0
            Me.Close()
        End If
    End Sub

    Private Sub AddEvent_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If

        If e.KeyData = Keys.Escape Then
            If CheckEventClose() = True Then
                CheckAddInfo_CancelAction()
                Declarations.MyResult = 0
                Me.Close()
            End If
        End If
    End Sub

    Private Function CheckEventClose() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка - если действие уже сохранено как закрытое, а мы выходим по кнопке
        '// отмена - возможно, мы что - то не записали, а дальнейшее редактирование будет недоступно
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "SELECT COUNT(EventID) AS CC "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') AND "
        MySQLStr = MySQLStr & "(ActionResultID IS NOT NULL) "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.Fields("CC").Value = 0 Then
            '---закрытой записи нет
            trycloseMyRec()
            CheckEventClose = True
        Else
            '---есть закрытая запись
            trycloseMyRec()
            MsgBox("При предыдущем сохранении информации о данном действии был сохранен результат действия. Это значит, что после выхода из данного окна данное действие нельзя будет отредактировать. Для корректного выхода из данного окна в этой ситуации (проверка данных, которые необходимо сохранить), воспользуйтесь кнопкой сохранения данных.", MsgBoxStyle.Critical, "Внимание!")
            CheckEventClose = False
        End If
    End Function

    Private Sub CheckAddInfo_CancelAction()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление лишней дополнительной информации при отмене сохранения действия
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyActionResult As Integer

        MySQLStr = "SELECT ISNULL(ActionResultID, 0) AS ActionResultID "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & " ') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            MyActionResult = 0
        Else
            MyActionResult = Declarations.MyRec.Fields("ActionResultID").Value
            trycloseMyRec()
        End If

        If MyActionResult <> 1 Then '---продажа
            MySQLStr = "DELETE FROM tbl_CRM_Orders "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

        If MyActionResult <> 2 Then '---Отказ - высокие цены
            MySQLStr = "DELETE FROM tbl_CRM_HighPrice "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

        If MyActionResult <> 3 Then '---Отказ - Нет товара на складе
            MySQLStr = "DELETE FROM tbl_CRM_WHAbsences "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

    End Sub

    Private Sub AddEvent_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы и значений в форму
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка направлений
        Dim MyDs As New DataSet
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    'для списка способов контакта
        Dim MyDs1 As New DataSet
        Dim MyAdapter2 As SqlClient.SqlDataAdapter    'для списка действий
        Dim MyDs2 As New DataSet
        Dim MyAdapter3 As SqlClient.SqlDataAdapter    'для списка результатов
        Dim MyDs3 As New DataSet
        Dim MyAdapter4 As SqlClient.SqlDataAdapter    'для списка пользователей
        Dim MyDs4 As New DataSet
        Dim MyAdapter5 As SqlClient.SqlDataAdapter    'для списка транспорта
        Dim MyDs5 As New DataSet
        Dim MyGUID As Guid

        '---направление
        MySQLStr = "SELECT DirectionID, DirectionName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Directions WITH(NOLOCK) "
        MySQLStr = MySQLStr & "ORDER BY DirectionID "
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "DirectionName" 'Это то что будет отображаться
            ComboBox1.ValueMember = "DirectionID"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        ComboBox1.SelectedValue = 2

        '---Способ контакта
        MySQLStr = "SELECT EventTypeID, EventTypeName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_EventTypes WITH(NOLOCK) "
        MySQLStr = MySQLStr & "ORDER BY EventTypeID "
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox2.DisplayMember = "EventTypeName" 'Это то что будет отображаться
            ComboBox2.ValueMember = "EventTypeID"   'это то что будет храниться
            ComboBox2.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        ComboBox2.SelectedValue = 4

        '---Действие
        MySQLStr = "SELECT ActionID, ActionName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Actions WITH(NOLOCK) "
        MySQLStr = MySQLStr & "ORDER BY ActionID "
        Try
            MyAdapter2 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter2.SelectCommand.CommandTimeout = 600
            MyAdapter2.Fill(MyDs2)
            ComboBox4.DisplayMember = "ActionName" 'Это то что будет отображаться
            ComboBox4.ValueMember = "ActionID"   'это то что будет храниться
            ComboBox4.DataSource = MyDs2.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        ComboBox4.SelectedValue = 999999

        '---Результат
        MySQLStr = "SELECT ActionResultID, ActionResultName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_ActionsResultTypes WITH(NOLOCK) "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 0 AS ActionResultID, '' AS ActionResultName "
        MySQLStr = MySQLStr & "ORDER BY ActionResultID "
        Try
            MyAdapter3 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter3.SelectCommand.CommandTimeout = 600
            MyAdapter3.Fill(MyDs3)
            ComboBox5.DisplayMember = "ActionResultName" 'Это то что будет отображаться
            ComboBox5.ValueMember = "ActionResultID"   'это то что будет храниться
            ComboBox5.DataSource = MyDs3.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---в зависимости от значения параметра - Разрешено ли пользователям из групп CRMManagers и CRMDirector менять 
        '---пользователя - владельца события при создании (редактировании) 1-можно 0-нельзя
        '---выставляем видимость управляющих элементов
        If Declarations.AllowChangeUser = "0" Then '---не разрешено
            Label14.Visible = False
            ComboBox3.Visible = False
        Else
            If Declarations.MyCCPermission = True Or Declarations.MyPermission = True Then
                Label14.Visible = True
                ComboBox3.Visible = True
            Else
                Label14.Visible = False
                ComboBox3.Visible = False
            End If
        End If

        '---Список пользователей (не блокированных)
        If Declarations.MyPermission = True Then
            '---Доступны все пользователи
            MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
            MySQLStr = MySQLStr & "WHERE(ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "
            MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        ElseIf Declarations.MyCCPermission = True Then
            '---Доступны пользователи определенного кост центра
            MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
            MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
            MySQLStr = MySQLStr & "(tbl_CRM_CCOwners.CCOwn = N'" & Declarations.CC & "') "
            MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        Else
            '---только один пользователь (вошедший в систему)
            MySQLStr = "SELECT ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
            MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
            MySQLStr = MySQLStr & "(ScalaSystemDB.dbo.ScaUsers.UserID = " & Declarations.UserID & ") "
            MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        End If
        Try
            MyAdapter4 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter4.SelectCommand.CommandTimeout = 600
            MyAdapter4.Fill(MyDs4)
            ComboBox3.DisplayMember = "FullName" 'Это то что будет отображаться
            ComboBox3.ValueMember = "UserID"   'это то что будет храниться
            ComboBox3.DataSource = MyDs4.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---Список транспорта
        MySQLStr = "SELECT TransportID, TransportName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Transport WITH(NOLOCK) "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT 0 AS TransportID, '' AS TransportName "
        MySQLStr = MySQLStr & "ORDER BY TransportID "
        Try
            MyAdapter5 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter5.SelectCommand.CommandTimeout = 600
            MyAdapter5.Fill(MyDs5)
            ComboBox6.DisplayMember = "TransportName" 'Это то что будет отображаться
            ComboBox6.ValueMember = "TransportID"   'это то что будет храниться
            ComboBox6.DataSource = MyDs5.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '----В зависимости от того - это новое действие или редактирование - читаем дополнительные данные и выставляем значения
        If StartParam = "Create" Then
            '---создаем запись о действии
            MyGUID = Guid.NewGuid
            Declarations.MyEventID = MyGUID.ToString
            TextBox1.Enabled = False
            TextBox2.Enabled = True
            TextBox5.Enabled = False
            If Declarations.MyCCPermission = True Or Declarations.MyPermission = True Then
                ComboBox3.SelectedValue = UserID
            Else
                ComboBox3.SelectedValue = Declarations.UserID
            End If
            ComboBox6.Enabled = False
            TextBox9.Enabled = False
            '---утверждаемые
            ComboBox1.Enabled = True
            ComboBox2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
            ComboBox4.Enabled = True
            DateTimePicker1.Enabled = True
            '---Ввод передаваемых параметрами данных
            DateTimePicker1.Value = ActDate
            If CompanyID.Equals("") = False Then
                Declarations.MyClientID = Trim(CompanyID)
                MySQLStr = "SELECT CompanyName "
                MySQLStr = MySQLStr & "FROM tbl_CRM_Companies "
                MySQLStr = MySQLStr & "WHERE (CompanyID = '" & Trim(CompanyID) & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                    Declarations.MyClientID = ""
                Else
                    TextBox6.Text = Declarations.MyRec.Fields("CompanyName").Value
                End If
                trycloseMyRec()
            End If

            ComboBox1.Select()
        ElseIf StartParam = "Edit" Then
            '---редактируем существующую запись
            MySQLStr = "SELECT tbl_CRM_Events.*, tbl_CRM_Companies.CompanyName, LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(tbl_CRM_Contacts.ContactName,''))) "
            MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ISNULL(tbl_CRM_Contacts.ContactPhone,''))) + ' ' + LTRIM(RTRIM(ISNULL(tbl_CRM_Contacts.ContactEMail,''))))) AS ContactFullName, "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers.FullName, Ltrim(Rtrim(Ltrim(Rtrim(ISNULL(tbl_CRM_Projects.ProjectName, ''))) + ' ' + Ltrim(Rtrim(ISNULL(tbl_CRM_Projects.ProjectComment, ''))))) AS ProjectInfo "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Contacts ON tbl_CRM_Events.ContactID = tbl_CRM_Contacts.ContactID AND "
            MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyID = tbl_CRM_Contacts.CompanyID INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.UserID = ScalaSystemDB.dbo.ScaUsers.UserID "
            MySQLStr = MySQLStr & "LEFT OUTER JOIN tbl_CRM_Projects ON tbl_CRM_Events.ProjectID = tbl_CRM_Projects.ProjectID "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Данное действие не существует, возможно, удалено другим пользователем. Обновите данные в окне со списком действий.", MsgBoxStyle.Critical, "Внимание!")
                Me.Close()
            Else
                Declarations.MyRec.MoveFirst()
                ComboBox1.SelectedValue = Declarations.MyRec.Fields("DirectionID").Value
                ComboBox2.SelectedValue = Declarations.MyRec.Fields("EventTypeID").Value
                If ComboBox2.SelectedValue = 999999 Then
                    TextBox1.Enabled = True
                    TextBox1.Text = Declarations.MyRec.Fields("EventTypeDescription").Value
                Else
                    TextBox1.Enabled = False
                End If
                Declarations.MyClientID = Declarations.MyRec.Fields("CompanyID").Value
                TextBox6.Text = Declarations.MyRec.Fields("CompanyName").Value
                Declarations.MyContactID = Declarations.MyRec.Fields("ContactID").Value
                TextBox7.Text = Declarations.MyRec.Fields("ContactFullName").Value.ToString
                ComboBox4.SelectedValue = Declarations.MyRec.Fields("ActionID").Value
                If ComboBox4.SelectedValue = 999999 Then
                    TextBox2.Enabled = True
                    TextBox2.Text = Declarations.MyRec.Fields("ActionDescription").Value
                Else
                    TextBox2.Enabled = False
                End If
                DateTimePicker1.Value = Declarations.MyRec.Fields("ActionPlannedDate").Value
                TextBox3.Text = Declarations.MyRec.Fields("ActionSumm").Value.ToString
                TextBox4.Text = Declarations.MyRec.Fields("ActionComments").Value
                ComboBox5.SelectedValue = IIf(Declarations.MyRec.Fields("ActionResultID").Value.ToString = "", 0, Declarations.MyRec.Fields("ActionResultID").Value)
                If ComboBox5.SelectedValue = 999999 Then
                    TextBox5.Enabled = True
                    TextBox5.Text = Declarations.MyRec.Fields("ActionResultDescription").Value
                Else
                    TextBox5.Enabled = False
                End If
                Declarations.OwnerID = Declarations.MyRec.Fields("OwnerID").Value
                Declarations.MyProjectID = Declarations.MyRec.Fields("ProjectID").Value.ToString
                TextBox8.Text = Declarations.MyRec.Fields("ProjectInfo").Value
                ComboBox3.SelectedValue = Declarations.MyRec.Fields("UserID").Value
                ComboBox6.SelectedValue = IIf(Declarations.MyRec.Fields("TransportID").Value.ToString = "", 0, Declarations.MyRec.Fields("TransportID").Value)
                If ComboBox2.SelectedValue = 4 Then
                    ComboBox6.Enabled = True
                Else
                    ComboBox6.SelectedValue = 0
                    ComboBox6.Enabled = False
                End If
                If ComboBox6.SelectedValue = 0 Then
                    TextBox9.Text = ""
                    TextBox9.Enabled = False
                Else
                    TextBox9.Text = Declarations.MyRec.Fields("TransportDistance").Value
                    TextBox9.Enabled = True
                End If

                If Declarations.MyRec.Fields("IsApproved").Value = True Then '---утвержденная запись
                    ComboBox1.Enabled = False
                    ComboBox2.Enabled = False
                    Button3.Enabled = False
                    Button4.Enabled = False
                    ComboBox4.Enabled = False
                    DateTimePicker1.Enabled = False
                    TextBox1.Enabled = False
                    TextBox2.Enabled = False
                Else                                                         '---запись не утверждена
                    ComboBox1.Enabled = True
                    ComboBox2.Enabled = True
                    Button3.Enabled = True
                    Button4.Enabled = True
                    ComboBox4.Enabled = True
                    DateTimePicker1.Enabled = True
                End If

                trycloseMyRec()
                ComboBox1.Select()

            End If
        ElseIf StartParam = "Copy" Then
            '---Создаем новую запись на основе старой
            MyGUID = Guid.NewGuid
            Declarations.MyEventID = MyGUID.ToString

            MySQLStr = "SELECT tbl_CRM_Events.*, tbl_CRM_Companies.CompanyName, LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(tbl_CRM_Contacts.ContactName,''))) "
            MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ISNULL(tbl_CRM_Contacts.ContactPhone,''))) + ' ' + LTRIM(RTRIM(ISNULL(tbl_CRM_Contacts.ContactEMail,''))))) AS ContactFullName, "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers.FullName, Ltrim(Rtrim(Ltrim(Rtrim(ISNULL(tbl_CRM_Projects.ProjectName, ''))) + ' ' + Ltrim(Rtrim(ISNULL(tbl_CRM_Projects.ProjectComment, ''))))) AS ProjectInfo "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Events WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Contacts ON tbl_CRM_Events.ContactID = tbl_CRM_Contacts.ContactID AND "
            MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyID = tbl_CRM_Contacts.CompanyID INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.UserID = ScalaSystemDB.dbo.ScaUsers.UserID "
            MySQLStr = MySQLStr & "LEFT OUTER JOIN tbl_CRM_Projects ON tbl_CRM_Events.ProjectID = tbl_CRM_Projects.ProjectID "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Trim(Declarations.MyOldEventID) & "') "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                trycloseMyRec()
                MsgBox("Действие, на основе которого вы делаете копию, не существует, возможно, удалено другим пользователем. Обновите данные в окне со списком действий.", MsgBoxStyle.Critical, "Внимание!")
                Me.Close()
            Else
                Declarations.MyRec.MoveFirst()
                ComboBox1.SelectedValue = Declarations.MyRec.Fields("DirectionID").Value
                ComboBox2.SelectedValue = Declarations.MyRec.Fields("EventTypeID").Value
                If ComboBox2.SelectedValue = 999999 Then
                    TextBox1.Enabled = True
                    TextBox1.Text = Declarations.MyRec.Fields("EventTypeDescription").Value
                Else
                    TextBox1.Enabled = False
                End If
                Declarations.MyClientID = Declarations.MyRec.Fields("CompanyID").Value
                TextBox6.Text = Declarations.MyRec.Fields("CompanyName").Value
                Declarations.MyContactID = Declarations.MyRec.Fields("ContactID").Value
                TextBox7.Text = Declarations.MyRec.Fields("ContactFullName").Value.ToString
                ComboBox4.SelectedValue = Declarations.MyRec.Fields("ActionID").Value
                If ComboBox4.SelectedValue = 999999 Then
                    TextBox2.Enabled = True
                    TextBox2.Text = Declarations.MyRec.Fields("ActionDescription").Value
                Else
                    TextBox2.Enabled = False
                End If
                Declarations.OwnerID = Declarations.UserID
                'Declarations.MyProjectID = Declarations.MyRec.Fields("ProjectID").Value.ToString
                'TextBox8.Text = Declarations.MyRec.Fields("ProjectInfo").Value
                ComboBox3.SelectedValue = Declarations.UserID
                If ComboBox2.SelectedValue = 4 Then
                    ComboBox6.Enabled = True
                Else
                    ComboBox6.SelectedValue = 0
                    ComboBox6.Enabled = False
                End If

                '---утверждаемые
                ComboBox1.Enabled = True
                ComboBox2.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True
                ComboBox4.Enabled = True
                DateTimePicker1.Enabled = True

                trycloseMyRec()
                ComboBox1.Select()

            End If
        End If
        '---Аттачменты
        LoadAttachments()
        CheckAttachmentsButtons()
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

        MySQLStr = "SELECT AttachmentID, EventID, AttachmentName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Attachments WITH(NOLOCK) "
        MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
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
        DataGridView1.Columns(1).HeaderText = "EvID"
        DataGridView1.Columns(1).Width = 0
        DataGridView1.Columns(1).Visible = False
        DataGridView1.Columns(2).HeaderText = "Имя файла"
        DataGridView1.Columns(2).Width = 550
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

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox1, True, True, True, False)
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

    Private Sub ComboBox2_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ComboBox2.SelectedValue = 999999 Then
            TextBox1.Enabled = True
        Else
            TextBox1.Text = ""
            TextBox1.Enabled = False
        End If

        If ComboBox2.SelectedValue = 4 Then
            ComboBox6.Enabled = True
        Else
            ComboBox6.SelectedValue = 0
            ComboBox6.Enabled = False
        End If

        Me.SelectNextControl(ComboBox2, True, True, True, False)
    End Sub

    Private Sub ComboBox4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox4_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ComboBox4.SelectedValue = 999999 Then
            TextBox2.Enabled = True
        Else
            TextBox2.Text = ""
            TextBox2.Enabled = False
        End If

        Me.SelectNextControl(ComboBox4, True, True, True, False)
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

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(DateTimePicker1, True, True, True, False)
    End Sub

    Private Sub ComboBox5_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox5.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            If StartParam = "Create" Then
                '---Если запись только создается - ее надо сначала правильно заполнить и сохранить
                If CheckDataFiling(False) = True Then
                    SaveNewData()
                    StartParam = "Edit"
                End If
            End If

            If ComboBox5.SelectedValue = 999999 Then
                TextBox5.Enabled = True
            Else
                TextBox5.Text = ""
                TextBox5.Enabled = False
            End If

            Me.SelectNextControl(sender, True, True, True, False)
            If ComboBox5.SelectedValue <> 999999 And ComboBox5.SelectedValue <> 0 Then
                '---при необходимости вводим дополнительную информацию
                GetAdditionalInfo()
                Me.SelectNextControl(sender, True, True, True, False)
            End If
        End If
    End Sub

    Private Sub ComboBox5_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If StartParam = "Create" And ComboBox5.SelectedValue <> Nothing Then
            '---Если запись только создается - ее надо сначала правильно заполнить и сохранить
            If CheckDataFiling(False) = True Then
                SaveNewData()
                StartParam = "Edit"
            End If
        End If

        If ComboBox5.SelectedValue = 999999 Then
            TextBox5.Enabled = True
        Else
            TextBox5.Text = ""
            TextBox5.Enabled = False
        End If

        CheckAddInfo_ResultChanging()

        Me.SelectNextControl(sender, True, True, True, False)
        If ComboBox5.SelectedValue <> 999999 And ComboBox5.SelectedValue <> 0 Then
            '---при необходимости вводим дополнительную информацию
            GetAdditionalInfo()
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

    Private Sub TextBox3_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox3.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения поля
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox3.Text) <> "" Then
            If InStr(TextBox3.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Сумма действия (РУБ)"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox3.Text
                Catch ex As Exception
                    MsgBox("В поле ""Сумма действия (РУБ)"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
            e.Cancel = False
        Else
            'MsgBox("В поле ""Сумма действия (РУБ)"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
            'e.Cancel = True
            'Exit Sub
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

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна выбора клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        MyCustomerSelect = New CustomerSelect
        MyCustomerSelect.SourceForm = "AddEvent"
        MyCustomerSelect.ShowDialog()
        TextBox7.Text = ""
        If TextBox6.Text <> "" Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна выбора контакта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If TextBox6.Text = "" Then
            MsgBox("Вы можете выбрать контакт только после того, как выбран клиент.", MsgBoxStyle.Critical, "Внимание!")
            Button3.Select()
        Else
            MyContactSelect = New ContactSelect
            MyContactSelect.ShowDialog()
            If TextBox7.Text <> "" Then
                Me.SelectNextControl(sender, True, True, True, False)
            End If
        End If
    End Sub

    Private Sub Button3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна выбора клиента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            MyCustomerSelect = New CustomerSelect
            MyCustomerSelect.ShowDialog()
            TextBox7.Text = ""
            If TextBox6.Text <> "" Then
                Me.SelectNextControl(sender, True, True, True, False)
            End If
        End If
    End Sub

    Private Sub Button4_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button4.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна выбора контакта
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            If TextBox6.Text = "" Then
                MsgBox("Вы можете выбрать контакт только после того, как выбран клиент.", MsgBoxStyle.Critical, "Внимание!")
                Button3.Select()
            Else
                MyContactSelect = New ContactSelect
                MyContactSelect.ShowDialog()
                If TextBox7.Text <> "" Then
                    Me.SelectNextControl(sender, True, True, True, False)
                End If
            End If
        End If
    End Sub

    Private Function CheckDataFiling(ByVal AddCheck As Boolean) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyCount As Integer
        Dim MyDbl As Double

        If Trim(TextBox1.Text) = "" And ComboBox2.SelectedValue = 999999 Then
            MsgBox("Поле ""Другой способ контакта"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox1.Select()
            Exit Function
        End If

        If Trim(TextBox6.Text) = "" Then
            MsgBox("Необходимо выбрать клиента, для которого создается действие.", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Button3.Select()
            Exit Function
        End If

        If Trim(TextBox7.Text) = "" Then
            MsgBox("Необходимо выбрать контакт, для которого создается действие.", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            Button4.Select()
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" And ComboBox4.SelectedValue = 999999 Then
            MsgBox("Поле ""Другое действие"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox2.Select()
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" And ComboBox5.SelectedValue = 999999 Then
            MsgBox("Поле ""Другой результат"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox5.Select()
            Exit Function
        End If

        If Trim(TextBox3.Text) = "" Then
            'MsgBox("Поле ""Сумма действия (РУБ)"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            'CheckDataFiling = False
            'TextBox3.Select()
            'Exit Function
        End If

        If Trim(TextBox3.Text) <> "" Then
            Try
                MyDbl = CDbl(TextBox3.Text)
            Catch ex As Exception
                MsgBox("Поле ""Сумма действия (РУБ)"" должно быть заполнено числом", MsgBoxStyle.Critical, "Внимание")
                CheckDataFiling = False
                TextBox3.Select()
                Exit Function
            End Try
        End If

        If DateTimePicker1.Value < DateAdd(DateInterval.Day, -8, CDate(DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()))) Then
            MsgBox("Плановая дата выполнения действия не должна быть меньше сегодняшней больше чем на 8 дней.", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            DateTimePicker1.Select()
            Exit Function
        End If

        If TextBox9.Enabled = True Then
            If Trim(TextBox9.Text) = "" Then
                MsgBox("Поле ""Расстояние (км)"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
                CheckDataFiling = False
                TextBox9.Select()
                Exit Function
            End If

            Try
                MyDbl = CDbl(TextBox9.Text)
            Catch ex As Exception
                MsgBox("Поле ""Расстояние (км)"" должно быть заполнено числом", MsgBoxStyle.Critical, "Внимание")
                CheckDataFiling = False
                TextBox9.Select()
                Exit Function
            End Try
        End If

        If AddCheck = True Then
            '---Проверка занесения дополнительной (в т.ч. маркетинговой информации) в случае закрытия действия
            If ComboBox5.SelectedValue = 1 Then '---продажа
                MySQLStr = "SELECT COUNT(ID) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_CRM_Orders WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                MyCount = Declarations.MyRec.Fields("CC").Value
                trycloseMyRec()
                If MyCount = 0 Then
                    MsgBox("Вами выбран результат действия - 'продажа'. При этом вы должны указать хотя бы один заказ ненулевого типа, созданный в результате этого действия. ", MsgBoxStyle.Critical, "Внимание")
                    CheckDataFiling = False
                    Button5.Select()
                    Exit Function
                End If
            End If

            If ComboBox5.SelectedValue = 2 Then '---Отказ - высокие цены
                MySQLStr = "SELECT COUNT(ID) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_CRM_HighPrice WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                MyCount = Declarations.MyRec.Fields("CC").Value
                trycloseMyRec()
                If MyCount = 0 Then
                    MsgBox("Вами выбран результат действия - 'Отказ - высокие цены'. При этом вы должны указать хотя бы для одного запаса, по которому был отказ, наши цены и ожидания клиента. ", MsgBoxStyle.Critical, "Внимание")
                    CheckDataFiling = False
                    Button5.Select()
                    Exit Function
                End If
            End If

            If ComboBox5.SelectedValue = 3 Then '---Отказ - Нет товара на складе
                MySQLStr = "SELECT COUNT(ID) AS CC "
                MySQLStr = MySQLStr & "FROM tbl_CRM_WHAbsences WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                MyCount = Declarations.MyRec.Fields("CC").Value
                trycloseMyRec()
                If MyCount = 0 Then
                    MsgBox("Вами выбран результат действия - 'Отказ - Нет товара на складе'. При этом вы должны указать хотя бы для одного запаса, по которому был отказ, доступное количество на нашем складе и запрошенное клиентом. ", MsgBoxStyle.Critical, "Внимание")
                    CheckDataFiling = False
                    Button5.Select()
                    Exit Function
                End If
            End If
        End If

        CheckDataFiling = True

    End Function

    Private Function SaveData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As MsgBoxResult
        Dim MySQLStr As String

        If CheckDataFiling(True) = True Then
            If ComboBox5.SelectedValue <> 0 Then
                MyRez = MsgBox("В окне редактирования выбран результат действия. Это значит, что после сохранения это действие будет недоступно для редактирования. Вы уверены, что хотите сохранить действие и тем закрыть его? ", MsgBoxStyle.YesNo, "Внимание!")
            Else
                MyRez = MsgBoxResult.Yes
            End If
            If MyRez = MsgBoxResult.Yes Then
                If StartParam = "Create" Or StartParam = "Copy" Then
                    SaveNewData()
                Else
                    UpdateData()
                End If
                '---Напоминалка в календарь
                'If CreateCalendarEvent(Declarations.MyEventID) = False Then
                'If CreateCalendarEventEWS(Declarations.MyEventID) = False Then
                If CreateCalendarEventZimbra(Declarations.MyEventID) = False Then
                    MySQLStr = "Exec spp_CRM_SendAppointment N'" & Declarations.MyEventID & "' "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                End If
                Declarations.MyResult = 1
                Me.Close()
            End If
        End If
    End Function

    Public Function CreateCalendarEvent(ByVal MyEventID As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание события в календаре в офисе 365
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As New spbadm4.Esk365ServiceClient
        Dim MyCalendarEvent As New spbadm4.CreateCalendarEventType
        Dim MyRez As String
        Dim MySQLStr As String
        Dim MyDate As DateTime

        MySQLStr = "SELECT tbl_CRM_Directions.DirectionName, tbl_CRM_EventTypes.EventTypeName, ISNULL(tbl_CRM_Companies.CompanyName, '') AS CompanyName, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Contacts.ContactName, '') AS ContactName, ISNULL(tbl_CRM_Contacts.ContactPhone, '') AS ContactPhone, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Contacts.ContactEMail, '') AS ContactEMail, tbl_CRM_Events.ActionPlannedDate, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Events.ActionComments, '') AS ActionComments, RM.dbo.RM660100.RM66003 AS UserEmail, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_EventsInCalendar.CalEventID, '') AS CalEventID, "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(tbl_CRM_Events.ActionDescription, '') = '' THEN tbl_CRM_Actions.ActionName ELSE ISNULL(tbl_CRM_Events.ActionDescription, '') END AS ActionName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Directions ON tbl_CRM_Events.DirectionID = tbl_CRM_Directions.DirectionID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_EventTypes ON tbl_CRM_Events.EventTypeID = tbl_CRM_EventTypes.EventTypeID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Contacts ON tbl_CRM_Events.ContactID = tbl_CRM_Contacts.ContactID AND tbl_CRM_Companies.CompanyID = tbl_CRM_Contacts.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.UserID = ScalaSystemDB.dbo.ScaUsers.UserID INNER JOIN "
        MySQLStr = MySQLStr & "RM.dbo.RM660100 ON ScalaSystemDB.dbo.ScaUsers.FullName = RM.dbo.RM660100.RM66002 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Actions ON tbl_CRM_Events.ActionID = tbl_CRM_Actions.ActionID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_EventsInCalendar ON tbl_CRM_Events.EventID = tbl_CRM_EventsInCalendar.EventID "
        MySQLStr = MySQLStr & "WHERE (RM.dbo.RM660100.RM66003 <> '') AND (tbl_CRM_Events.EventID = '" & MyEventID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Return False
        Else
            MyCalendarEvent.CalendarEventIDOld = Declarations.MyRec.Fields("CalEventID").Value

            MyCalendarEvent.Subject = Declarations.MyRec.Fields("DirectionName").Value & " "
            MyCalendarEvent.Subject = MyCalendarEvent.Subject & Declarations.MyRec.Fields("EventTypeName").Value
            MyCalendarEvent.Subject = MyCalendarEvent.Subject & " Компания " & Declarations.MyRec.Fields("CompanyName").Value

            MyCalendarEvent.Body = Declarations.MyRec.Fields("DirectionName").Value & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & Declarations.MyRec.Fields("EventTypeName").Value
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Компания " & Declarations.MyRec.Fields("CompanyName").Value & Chr(13) & Chr(10)
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Контакт: " & Declarations.MyRec.Fields("ContactName").Value &
            MyCalendarEvent.Body = MyCalendarEvent.Body & " Телефон: " & Declarations.MyRec.Fields("ContactPhone").Value
            MyCalendarEvent.Body = MyCalendarEvent.Body & " Email: " & Declarations.MyRec.Fields("ContactEMail").Value & Chr(13) & Chr(10)
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Действие: " & Declarations.MyRec.Fields("ActionName").Value & Chr(13) & Chr(10)
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Комментарий: " & Declarations.MyRec.Fields("ActionComments").Value & Chr(13) & Chr(10)

            MyDate = Declarations.MyRec.Fields("ActionPlannedDate").Value
            MyCalendarEvent.Start = New DateTime(MyDate.Year, MyDate.Month, MyDate.Day, 9, 0, 0)
            MyCalendarEvent.Finish = New DateTime(MyDate.Year, MyDate.Month, MyDate.Day, 17, 30, 0)
            MyCalendarEvent.Timezone = "Russian Standard Time"
            MyCalendarEvent.Email = Declarations.MyRec.Fields("UserEmail").Value
            MyCalendarEvent.Login = "Esk365ServiceUser"

            Try
                MyRez = MyObj.CreateCalendarEvent(MyCalendarEvent)
                If MyRez.Equals("") Then
                    Return False
                Else
                    MySQLStr = "DELETE FROM tbl_CRM_EventsInCalendar "
                    MySQLStr = MySQLStr & "WHERE (EventID = '" & MyEventID & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    MySQLStr = "INSERT INTO tbl_CRM_EventsInCalendar "
                    MySQLStr = MySQLStr & "(EventID, CalEventID) "
                    MySQLStr = MySQLStr & "VALUES ('" & MyEventID & "', "
                    MySQLStr = MySQLStr & "N'" & MyRez & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    Return True
                End If
            Catch ex As Exception
                Return False
            End Try
        End If
    End Function

    Public Function CreateCalendarEventEWS(ByVal MyEventID As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание события в календаре при помощи EWS
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As New spbadm4_EWS.EskEWSServiceClient
        Dim MyCalendarEvent As New spbadm4_EWS.CreateCalendarEventType
        Dim MySQLStr As String
        Dim MyDate As DateTime
        Dim MyRez As String

        MySQLStr = "SELECT tbl_CRM_Directions.DirectionName, tbl_CRM_EventTypes.EventTypeName, ISNULL(tbl_CRM_Companies.CompanyName, '') AS CompanyName, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Contacts.ContactName, '') AS ContactName, ISNULL(tbl_CRM_Contacts.ContactPhone, '') AS ContactPhone, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Contacts.ContactEMail, '') AS ContactEMail, tbl_CRM_Events.ActionPlannedDate, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Events.ActionComments, '') AS ActionComments, RM.dbo.RM660100.RM66003 AS UserEmail, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_EventsInCalendar.CalEventID, '') AS CalEventID, "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(tbl_CRM_Events.ActionDescription, '') = '' THEN tbl_CRM_Actions.ActionName ELSE ISNULL(tbl_CRM_Events.ActionDescription, '') END AS ActionName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Directions ON tbl_CRM_Events.DirectionID = tbl_CRM_Directions.DirectionID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_EventTypes ON tbl_CRM_Events.EventTypeID = tbl_CRM_EventTypes.EventTypeID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Contacts ON tbl_CRM_Events.ContactID = tbl_CRM_Contacts.ContactID AND tbl_CRM_Companies.CompanyID = tbl_CRM_Contacts.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.UserID = ScalaSystemDB.dbo.ScaUsers.UserID INNER JOIN "
        MySQLStr = MySQLStr & "RM.dbo.RM660100 ON ScalaSystemDB.dbo.ScaUsers.FullName = RM.dbo.RM660100.RM66002 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Actions ON tbl_CRM_Events.ActionID = tbl_CRM_Actions.ActionID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_EventsInCalendar ON tbl_CRM_Events.EventID = tbl_CRM_EventsInCalendar.EventID "
        MySQLStr = MySQLStr & "WHERE (RM.dbo.RM660100.RM66003 <> '') AND (tbl_CRM_Events.EventID = '" & MyEventID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Return False
        Else
            MyCalendarEvent.CalendarEventIDOld = Declarations.MyRec.Fields("CalEventID").Value

            MyCalendarEvent.Subject = Declarations.MyRec.Fields("DirectionName").Value & " "
            MyCalendarEvent.Subject = MyCalendarEvent.Subject & Declarations.MyRec.Fields("EventTypeName").Value
            MyCalendarEvent.Subject = MyCalendarEvent.Subject & " Компания " & Declarations.MyRec.Fields("CompanyName").Value

            MyCalendarEvent.Body = "<p>" & Declarations.MyRec.Fields("DirectionName").Value & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & Declarations.MyRec.Fields("EventTypeName").Value & "</p>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & "<p> Компания " & Replace(Declarations.MyRec.Fields("CompanyName").Value, """", "'") & "<br>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Контакт: " & Declarations.MyRec.Fields("ContactName").Value & "<br>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & "   Телефон: " & Declarations.MyRec.Fields("ContactPhone").Value & "<br>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & "   Email: " & Declarations.MyRec.Fields("ContactEMail").Value & "<br>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Действие: " & Declarations.MyRec.Fields("ActionName").Value & "<br>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Комментарий: " & Declarations.MyRec.Fields("ActionComments").Value & "</p>"

            MyDate = Declarations.MyRec.Fields("ActionPlannedDate").Value
            MyCalendarEvent.Start = New DateTime(MyDate.Year, MyDate.Month, MyDate.Day, 9, 0, 0)
            MyCalendarEvent.Finish = New DateTime(MyDate.Year, MyDate.Month, MyDate.Day, 17, 30, 0)
            MyCalendarEvent.Timezone = "Russian Standard Time"
            MyCalendarEvent.Email = Declarations.MyRec.Fields("UserEmail").Value
            MyCalendarEvent.Login = "EskEWSServiceUser"

            Try
                MyRez = MyObj.CreateCalendarEvent(MyCalendarEvent)
                If MyRez.Equals("") Then
                    Return False
                Else
                    MySQLStr = "DELETE FROM tbl_CRM_EventsInCalendar "
                    MySQLStr = MySQLStr & "WHERE (EventID = '" & MyEventID & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    MySQLStr = "INSERT INTO tbl_CRM_EventsInCalendar "
                    MySQLStr = MySQLStr & "(EventID, CalEventID) "
                    MySQLStr = MySQLStr & "VALUES ('" & MyEventID & "', "
                    MySQLStr = MySQLStr & "N'" & MyRez & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    Return True
                End If
            Catch ex As Exception
                Return False
            End Try
        End If
    End Function

    Public Function CreateCalendarEventZimbra(ByVal MyEventID As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание события в календаре при помощи Zimbra SOAP
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As New CalendarZimbraService.CalendarZimbraServiceClient
        Dim MyCalendarEvent As New CalendarZimbraService.CreateCalendarEventType
        Dim MySQLStr As String
        Dim MyDate As DateTime
        Dim MyRez As String

        MySQLStr = "SELECT tbl_CRM_Directions.DirectionName, tbl_CRM_EventTypes.EventTypeName, ISNULL(tbl_CRM_Companies.CompanyName, '') AS CompanyName, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Contacts.ContactName, '') AS ContactName, ISNULL(tbl_CRM_Contacts.ContactPhone, '') AS ContactPhone, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Contacts.ContactEMail, '') AS ContactEMail, tbl_CRM_Events.ActionPlannedDate, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Events.ActionComments, '') AS ActionComments, RM.dbo.RM660100.RM66003 AS UserEmail, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_EventsInCalendar.CalEventID, '') AS CalEventID, "
        MySQLStr = MySQLStr & "CASE WHEN ISNULL(tbl_CRM_Events.ActionDescription, '') = '' THEN tbl_CRM_Actions.ActionName ELSE ISNULL(tbl_CRM_Events.ActionDescription, '') END AS ActionName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Directions ON tbl_CRM_Events.DirectionID = tbl_CRM_Directions.DirectionID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_EventTypes ON tbl_CRM_Events.EventTypeID = tbl_CRM_EventTypes.EventTypeID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Contacts ON tbl_CRM_Events.ContactID = tbl_CRM_Contacts.ContactID AND tbl_CRM_Companies.CompanyID = tbl_CRM_Contacts.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.UserID = ScalaSystemDB.dbo.ScaUsers.UserID INNER JOIN "
        MySQLStr = MySQLStr & "RM.dbo.RM660100 ON ScalaSystemDB.dbo.ScaUsers.FullName = RM.dbo.RM660100.RM66002 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Actions ON tbl_CRM_Events.ActionID = tbl_CRM_Actions.ActionID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_EventsInCalendar ON tbl_CRM_Events.EventID = tbl_CRM_EventsInCalendar.EventID "
        MySQLStr = MySQLStr & "WHERE (RM.dbo.RM660100.RM66003 <> '') AND (tbl_CRM_Events.EventID = '" & MyEventID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Return False
        Else
            MyCalendarEvent.CalendarEventIDOld = Declarations.MyRec.Fields("CalEventID").Value

            'MyCalendarEvent.Subject = Declarations.MyRec.Fields("DirectionName").Value & " "
            'MyCalendarEvent.Subject = MyCalendarEvent.Subject & Declarations.MyRec.Fields("EventTypeName").Value
            'MyCalendarEvent.Subject = MyCalendarEvent.Subject & " Компания " & Declarations.MyRec.Fields("CompanyName").Value
            MyCalendarEvent.Subject = "CRM APPOINTMENT: "
            MyCalendarEvent.Subject = MyCalendarEvent.Subject & Declarations.MyRec.Fields("DirectionName").Value & " "
            MyCalendarEvent.Subject = MyCalendarEvent.Subject & Declarations.MyRec.Fields("EventTypeName").Value

            MyCalendarEvent.Body = Declarations.MyRec.Fields("DirectionName").Value & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & Declarations.MyRec.Fields("EventTypeName").Value & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Компания " & Replace(Declarations.MyRec.Fields("CompanyName").Value, """", "'") & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Контакт: " & Declarations.MyRec.Fields("ContactName").Value & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & "   Телефон: " & Declarations.MyRec.Fields("ContactPhone").Value & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & "   Email: " & Declarations.MyRec.Fields("ContactEMail").Value & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Действие: " & Declarations.MyRec.Fields("ActionName").Value & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Комментарий: " & Declarations.MyRec.Fields("ActionComments").Value & " "

            MyDate = Declarations.MyRec.Fields("ActionPlannedDate").Value
            MyCalendarEvent.Start = New DateTime(MyDate.Year, MyDate.Month, MyDate.Day, 9, 0, 0)
            MyCalendarEvent.Finish = New DateTime(MyDate.Year, MyDate.Month, MyDate.Day, 17, 30, 0)
            MyCalendarEvent.Timezone = "Russian Standard Time"
            MyCalendarEvent.Email = Declarations.MyRec.Fields("UserEmail").Value
            MyCalendarEvent.Login = "CalZimbraServiceUser"

            Try
                MyRez = MyObj.CreateCalendarEvent(MyCalendarEvent)
                If MyRez.Equals("") Then
                    Return False
                Else
                    MySQLStr = "DELETE FROM tbl_CRM_EventsInCalendar "
                    MySQLStr = MySQLStr & "WHERE (EventID = '" & MyEventID & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    MySQLStr = "INSERT INTO tbl_CRM_EventsInCalendar "
                    MySQLStr = MySQLStr & "(EventID, CalEventID) "
                    MySQLStr = MySQLStr & "VALUES ('" & MyEventID & "', "
                    MySQLStr = MySQLStr & "N'" & MyRez & "') "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)
                    Return True
                End If
            Catch ex As Exception
                Return False
            End Try
        End If
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки Сохранение данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        SaveData()
    End Sub

    Private Sub Button1_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button1.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Нажатие кнопки Сохранение данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            SaveData()
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
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

    Private Sub SaveNewData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// сохранение данных в случае создания новой записи
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---создание записи в таблице tbl_CRM_Events
        Declarations.OwnerID = Declarations.UserID
        MySQLStr = "INSERT INTO tbl_CRM_Events "
        MySQLStr = MySQLStr & "(EventID, DirectionID, EventTypeID, EventTypeDescription, CompanyID, "
        MySQLStr = MySQLStr & "ContactID, ActionTime, ActionID, ActionDescription, ActionPlannedDate, "
        MySQLStr = MySQLStr & "ActionSumm, ActionComments, ActionResultID, ActionResultDescription, UserID, "
        MySQLStr = MySQLStr & "OwnerID, ActionClosed, ProjectID, TransportID, TransportDistance, IsApproved) "
        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyEventID & "', "             '--ID действия
        MySQLStr = MySQLStr & ComboBox1.SelectedValue & ", "                           '--Направление действия
        MySQLStr = MySQLStr & ComboBox2.SelectedValue & ", "                           '--Способ контакта
        If ComboBox2.SelectedValue = 999999 Then                                       '--Дополнительный способ контакта
            MySQLStr = MySQLStr & "'" & TextBox1.Text & "', "
        Else
            MySQLStr = MySQLStr & "NULL, "
        End If
        MySQLStr = MySQLStr & "'" & Declarations.MyClientID & "', "                    '--клиент
        MySQLStr = MySQLStr & "'" & Declarations.MyContactID & "', "                   '--контакт
        If DateTimePicker1.Value < Now() Then
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "', 103), " '--дата создания записи равна планируемой дате выполненния
        Else
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()) & "', 103), " '--дата создания записи фактическая
        End If
        MySQLStr = MySQLStr & ComboBox4.SelectedValue & ", "                           '--действие
        If ComboBox4.SelectedValue = 999999 Then                                       '--другое действие
            MySQLStr = MySQLStr & "'" & TextBox2.Text & "', "
        Else
            MySQLStr = MySQLStr & "NULL, "
        End If
        MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "', 103), " '--плановая дата выполнения действия
        If Trim(TextBox3.Text) = "" Then
            MySQLStr = MySQLStr & "0, "                  '--сумма действия
        Else
            MySQLStr = MySQLStr & Replace(TextBox3.Text, ",", ".") & ", "                  '--сумма действия
        End If

        MySQLStr = MySQLStr & "'" & TextBox4.Text & "', "                              '--комментарий
        If ComboBox5.SelectedValue = 0 Then                                            '--результат действия
            MySQLStr = MySQLStr & "NULL, "
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & ComboBox5.SelectedValue & ", "
            If ComboBox5.SelectedValue = 999999 Then
                MySQLStr = MySQLStr & "'" & TextBox5.Text & "', "
            Else
                MySQLStr = MySQLStr & "NULL, "
            End If
        End If
        MySQLStr = MySQLStr & ComboBox3.SelectedValue & ", "                           '--ID пользователя
        'MySQLStr = MySQLStr & Declarations.UserID & ", "                              '--ID владельца
        MySQLStr = MySQLStr & ComboBox3.SelectedValue & ", "                           '--ID владельца
        If ComboBox5.SelectedValue = 0 Then                                            '--Дата закрытия действия 
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()) & "', 103), "
        End If
        If Trim(TextBox8.Text) = "" Then
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & "'" & Declarations.MyProjectID & "', "
        End If
        If ComboBox6.SelectedValue = 0 Then                                            '--Дата закрытия действия 
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & ComboBox6.SelectedValue & ", "
        End If
        If Trim(TextBox9.Text) = "" Then
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & Replace(TextBox9.Text, ",", ".") & ", "
        End If
        MySQLStr = MySQLStr & "0) "


        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '-------если дата создания действия меньше даты создания проекта, на который ссылается -корректируем дату создания проекта
        If Trim(TextBox8.Text) <> "" Then
            MySQLStr = "UPDATE tbl_CRM_Projects "
            MySQLStr = MySQLStr & "SET StartDate = CASE WHEN CONVERT(DATETIME, '"
            If DateTimePicker1.Value < Now() Then
                MySQLStr = MySQLStr & DatePart(DateInterval.Day, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Year, DateTimePicker1.Value)
            Else
                MySQLStr = MySQLStr & DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())
            End If
            MySQLStr = MySQLStr & "', 103) < StartDate THEN "
            MySQLStr = MySQLStr & "CONVERT(DATETIME, '"
            If DateTimePicker1.Value < Now() Then
                MySQLStr = MySQLStr & DatePart(DateInterval.Day, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Year, DateTimePicker1.Value)
            Else
                MySQLStr = MySQLStr & DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now())
            End If
            MySQLStr = MySQLStr & "', 103) ELSE StartDate END "
            MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "

            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

    End Sub

    Private Sub UpdateData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// сохранение данных в случае редактирования записи
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---Редактирование существующей записи в таблице tbl_CRM_Events
        MySQLStr = "UPDATE tbl_CRM_Events "
        MySQLStr = MySQLStr & "SET DirectionID = " & ComboBox1.SelectedValue & ", "    '--Направление действия
        MySQLStr = MySQLStr & "EventTypeID = " & ComboBox2.SelectedValue & ", "        '--Способ контакта
        If ComboBox2.SelectedValue = 999999 Then                                       '--Дополнительный способ контакта
            MySQLStr = MySQLStr & "EventTypeDescription = N'" & TextBox1.Text & "', "
        Else
            MySQLStr = MySQLStr & "EventTypeDescription = NULL, "
        End If
        MySQLStr = MySQLStr & "CompanyID = '" & Declarations.MyClientID & "', "        '--клиент
        MySQLStr = MySQLStr & "ContactID = '" & Declarations.MyContactID & "', "       '--контакт
        MySQLStr = MySQLStr & "ActionID = " & ComboBox4.SelectedValue & ", "           '--действие
        If ComboBox4.SelectedValue = 999999 Then                                       '--другое действие
            MySQLStr = MySQLStr & "ActionDescription ='" & TextBox2.Text & "', "
        Else
            MySQLStr = MySQLStr & "ActionDescription = NULL, "
        End If
        MySQLStr = MySQLStr & "ActionPlannedDate = CONVERT(DATETIME, '" & DatePart(DateInterval.Day, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Month, DateTimePicker1.Value) & "/" & DatePart(DateInterval.Year, DateTimePicker1.Value) & "', 103), "
        If Trim(TextBox3.Text) = "" Then
            MySQLStr = MySQLStr & "ActionSumm = 0, " '--сумма действия
        Else
            MySQLStr = MySQLStr & "ActionSumm = " & Replace(TextBox3.Text, ",", ".") & ", " '--сумма действия
        End If
        MySQLStr = MySQLStr & "ActionComments = N'" & TextBox4.Text & "', "            '--комментарий
        If ComboBox5.SelectedValue = 0 Then                                            '--результат действия
            MySQLStr = MySQLStr & "ActionResultID = NULL, "
            MySQLStr = MySQLStr & "ActionResultDescription = NULL, "
        Else
            MySQLStr = MySQLStr & "ActionResultID = " & ComboBox5.SelectedValue & ", "
            If ComboBox5.SelectedValue = 999999 Then
                MySQLStr = MySQLStr & "ActionResultDescription = '" & TextBox5.Text & "', "
            Else
                MySQLStr = MySQLStr & "ActionResultDescription = NULL, "
            End If
        End If
        MySQLStr = MySQLStr & "UserID = " & ComboBox3.SelectedValue & ", "                '--ID пользователя
        'MySQLStr = MySQLStr & "OwnerID = " & Declarations.OwnerID & ", "                 '--ID владельца
        MySQLStr = MySQLStr & "OwnerID = " & ComboBox3.SelectedValue & ", "               '--ID владельца
        If ComboBox5.SelectedValue = 0 Then                                               '--Дата закрытия действия 
            MySQLStr = MySQLStr & "ActionClosed = NULL, "
        Else
            MySQLStr = MySQLStr & "ActionClosed = CONVERT(DATETIME, '" & DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()) & "', 103), "
        End If
        If Trim(TextBox8.Text) = "" Then
            MySQLStr = MySQLStr & "ProjectID = NULL, "
        Else
            MySQLStr = MySQLStr & "ProjectID = '" & Declarations.MyProjectID & "', "
        End If
        If ComboBox6.SelectedValue = 0 Then                                            '--Дата закрытия действия 
            MySQLStr = MySQLStr & "TransportID = NULL, "
        Else
            MySQLStr = MySQLStr & "TransportID = " & ComboBox6.SelectedValue & ", "
        End If
        If Trim(TextBox9.Text) = "" Then
            MySQLStr = MySQLStr & "TransportDistance = NULL, "
        Else
            MySQLStr = MySQLStr & "TransportDistance = " & Replace(TextBox9.Text, ",", ".") & ", "
        End If
        MySQLStr = MySQLStr & "IsApproved = IsApproved "
        MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "

        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
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
                        MySQLStr = "INSERT INTO tbl_CRM_Attachments "
                        MySQLStr = MySQLStr & "(AttachmentID, EventID, AttachmentName, AttachmentBody) "
                        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyAttachmentID & "', "
                        MySQLStr = MySQLStr & "'" & Declarations.MyEventID & "', "
                        MySQLStr = MySQLStr & "N'" & FName & "', "
                        MySQLStr = MySQLStr & "NULL) "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)

                        MySQLStr = "SELECT AttachmentID, EventID, AttachmentName, AttachmentBody "
                        MySQLStr = MySQLStr & "FROM tbl_CRM_Attachments WITH(NOLOCK) "
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
                        MySQLStr = "DELETE FROM tbl_CRM_Attachments "
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

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление аттачмента из БД
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_CRM_Attachments "
        MySQLStr = MySQLStr & "WHERE (AttachmentID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        LoadAttachments()
        CheckAttachmentsButtons()
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
            MySQLStr = "SELECT AttachmentID, EventID, AttachmentName, AttachmentBody "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Attachments WITH(NOLOCK) "
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

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окон ввода дополнительной информации в случае закрытия действия
        '//
        '////////////////////////////////////////////////////////////////////////////////

        GetAdditionalInfo()
        Me.SelectNextControl(sender, True, True, True, False)
    End Sub

    Private Sub GetAdditionalInfo()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окон ввода дополнительной информации 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ComboBox5.SelectedValue = 1 Then       '---продажа
            MySalesOrderList = New SalesOrderList
            MySalesOrderList.ShowDialog()
        ElseIf ComboBox5.SelectedValue = 2 Then   '---Отказ - высокие цены
            MyHighPrice = New HighPrice
            MyHighPrice.ShowDialog()
        ElseIf ComboBox5.SelectedValue = 3 Then   '---Отказ - Нет товара на складе
            MyWHAbsences = New WHAbsences
            MyWHAbsences.ShowDialog()
        End If

    End Sub

    Private Sub CheckAddInfo_ResultChanging()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление лишней дополнительной информации при смене выбора результата действия
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        If ComboBox5.SelectedValue <> 1 Then '---продажа
            MySQLStr = "DELETE FROM tbl_CRM_Orders "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

        If ComboBox5.SelectedValue <> 2 Then '---Отказ - высокие цены
            MySQLStr = "DELETE FROM tbl_CRM_HighPrice "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If

        If ComboBox5.SelectedValue <> 3 Then '---Отказ - Нет товара на складе
            MySQLStr = "DELETE FROM tbl_CRM_WHAbsences "
            MySQLStr = MySQLStr & "WHERE (EventID = '" & Declarations.MyEventID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)
        End If
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна выбора проекта 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If TextBox6.Text = "" Then
            MsgBox("Вы можете выбрать проект только после того, как выбран клиент.", MsgBoxStyle.Critical, "Внимание!")
            Button3.Select()
        Else
            MyProjectSelect = New ProjectSelect
            MyProjectSelect.StartParam = "Edit"
            MyProjectSelect.ShowDialog()
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub Button10_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button10.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Очистка поля проекта 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        TextBox8.Text = ""
    End Sub

    Private Sub Button9_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button9.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Открытие окна выбора проекта 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If TextBox6.Text = "" Then
            MsgBox("Вы можете выбрать проект только после того, как выбран клиент.", MsgBoxStyle.Critical, "Внимание!")
            Button3.Select()
        Else
            MyProjectSelect = New ProjectSelect
            MyProjectSelect.ShowDialog()
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox3_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox3.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox3_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.SelectNextControl(ComboBox2, True, True, True, False)
    End Sub

    Private Sub ComboBox6_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox6.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по Enter
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Enter Then
            If ComboBox6.SelectedValue = 0 Then
                TextBox9.Text = ""
                TextBox9.Enabled = False
            Else
                TextBox9.Enabled = True
            End If

            Me.SelectNextControl(sender, True, True, True, False)
        End If
    End Sub

    Private Sub ComboBox6_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// переход на следующее поле по выбору значения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If ComboBox6.SelectedValue = 0 Then
            TextBox9.Text = ""
            TextBox9.Enabled = False
        Else
            TextBox9.Enabled = True
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

    Private Sub TextBox9_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox9.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения поля
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If TextBox9.Enabled = True Then
            If Trim(TextBox9.Text) <> "" Then
                If InStr(TextBox9.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                    MsgBox("В поле ""Расстояние (км)"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                Else
                    Try
                        MyRez = TextBox9.Text
                    Catch ex As Exception
                        MsgBox("В поле ""Расстояние (км)"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                        e.Cancel = True
                        Exit Sub
                    End Try
                End If
                e.Cancel = False
            ElseIf ComboBox6.SelectedValue <> 0 Then
                MsgBox("В поле ""Расстояние (км)"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            End If
        End If
    End Sub
End Class