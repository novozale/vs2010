Public Class PlansApprovement

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна утверждения планов
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub PlansApprovement_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Запрет закрытия окна по ALT+F4
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Alt + Keys.F4 Then
            e.Handled = True
        End If
    End Sub

    Private Sub PlansApprovement_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в окно
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка продавцов
        Dim MyDs As New DataSet

        '---Список продавцов
        If Declarations.MyPermission = True Then
            '---Доступны все продавцы
            MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
            MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "
            MySQLStr = MySQLStr & "AND (ScalaSystemDB.dbo.ScaUsers.UserID <> " & Declarations.UserID & ") "
            MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        ElseIf Declarations.MyCCPermission = True Then
            '---Доступны продавцы определенного кост центра
            MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
            MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
            MySQLStr = MySQLStr & "(tbl_CRM_CCOwners.CCOwn = N'" & Declarations.CC & "') "
            MySQLStr = MySQLStr & "AND (ScalaSystemDB.dbo.ScaUsers.UserID <> " & Declarations.UserID & ") "
            MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        Else
            '---только один продавец (вошедший в систему)
            MySQLStr = "SELECT ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
            MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
            MySQLStr = MySQLStr & "(ScalaSystemDB.dbo.ScaUsers.UserID = " & Declarations.UserID & ") "
            MySQLStr = MySQLStr & "AND (ScalaSystemDB.dbo.ScaUsers.UserID <> " & Declarations.UserID & ") "
            MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        End If
        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            ComboBox1.DisplayMember = "FullName" 'Это то что будет отображаться
            ComboBox1.ValueMember = "UserID"   'это то что будет храниться
            ComboBox1.DataSource = MyDs.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        'ComboBox1.SelectedValue = Declarations.UserID

        '---Загрузка данных
        DataLoading()
    End Sub

    Public Function DataLoading()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка списка действий (в соответствии с параметрами)
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка активностей
        Dim MyDs As New DataSet

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        '---активности только выбранного продавца
        MySQLStr = "Select "
        MySQLStr = MySQLStr & "tbl_CRM_Events.EventID, "
        MySQLStr = MySQLStr & "tbl_CRM_Events.ActionPlannedDate, "
        MySQLStr = MySQLStr & "tbl_CRM_Directions.DirectionName, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_EventTypes.EventTypeID = 999999 THEN tbl_CRM_Events.EventTypeDescription ELSE tbl_CRM_EventTypes.EventTypeName END AS EventTypeName, "
        MySQLStr = MySQLStr & "tbl_CRM_Companies.ScalaCustomerCode, "
        MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyName, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_Actions.ActionID = 999999 THEN tbl_CRM_Events.ActionDescription ELSE tbl_CRM_Actions.ActionName END AS ActionName, "
        MySQLStr = MySQLStr & "Ltrim(Rtrim(Ltrim(Rtrim(ISNULL(tbl_CRM_Projects.ProjectName, ''))) + ' ' + Ltrim(Rtrim(ISNULL(tbl_CRM_Projects.ProjectComment, ''))))) AS ProjectInfo, "
        MySQLStr = MySQLStr & "tbl_CRM_Events.IsApproved "
        MySQLStr = MySQLStr & "FROM tbl_CRM_ActionsResultTypes WITH(NOLOCK) RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Events INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_EventTypes ON tbl_CRM_Events.EventTypeID = tbl_CRM_EventTypes.EventTypeID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Directions ON tbl_CRM_Events.DirectionID = tbl_CRM_Directions.DirectionID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Contacts ON tbl_CRM_Events.ContactID = tbl_CRM_Contacts.ContactID AND "
        MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyID = tbl_CRM_Contacts.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Actions ON tbl_CRM_Events.ActionID = tbl_CRM_Actions.ActionID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.UserID = ScalaSystemDB.dbo.ScaUsers.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 ON "
        MySQLStr = MySQLStr & "tbl_CRM_ActionsResultTypes.ActionResultID = tbl_CRM_Events.ActionResultID "
        MySQLStr = MySQLStr & "LEFT OUTER JOIN tbl_CRM_Projects ON tbl_CRM_Events.ProjectID = tbl_CRM_Projects.ProjectID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies_Ext ON tbl_CRM_Companies.CompanyID = tbl_CRM_Companies_Ext.CompanyID "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Events.UserID = " & ComboBox1.SelectedValue & ") "
        MySQLStr = MySQLStr & " AND (tbl_CRM_Events.ActionResultID IS NULL) "
        MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Events.ActionPlannedDate "

        Try
            MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter.SelectCommand.CommandTimeout = 600
            MyAdapter.Fill(MyDs)
            DataGridView1.DataSource = MyDs.Tables(0)
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---заголовки
        DataGridView1.Columns(0).HeaderText = "ID"
        DataGridView1.Columns(0).Width = 40
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "Плани руемая дата"
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).HeaderText = "Направ ление"
        DataGridView1.Columns(2).Width = 70
        DataGridView1.Columns(3).HeaderText = "Способ контакта"
        DataGridView1.Columns(3).Width = 120
        DataGridView1.Columns(4).HeaderText = "Код клиента в Scala"
        DataGridView1.Columns(4).Width = 60
        DataGridView1.Columns(5).HeaderText = "Клиент"
        DataGridView1.Columns(5).Width = 150
        DataGridView1.Columns(6).HeaderText = "Действие"
        DataGridView1.Columns(6).Width = 120
        DataGridView1.Columns(7).HeaderText = "Проект"
        DataGridView1.Columns(7).Width = 120
        DataGridView1.Columns(8).HeaderText = "Утвержден"
        DataGridView1.Columns(8).Width = 100
    End Function

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в соответствии с выбранным продавцом
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---загрузка данных
        DataLoading()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Обновление данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        '---загрузка данных
        DataLoading()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// подсветка активностей
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim row As DataGridViewRow

        row = DataGridView1.Rows(e.RowIndex)
        If row.Cells(8).Value = 0 Then
            row.DefaultCellStyle.BackColor = Color.White
        Else
            row.DefaultCellStyle.BackColor = Color.LightGray
        End If
    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния утверждения плана
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyID As String

        If e.Button = Windows.Forms.MouseButtons.Left Then
            If e.ColumnIndex = 8 Then
                MyID = DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString
                ChangeApprovedState(DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString, DataGridView1.SelectedRows.Item(0).Cells(8).Value)
                '---загрузка данных
                DataLoading()
                '---текущей строкой сделать редактируемую
                For i As Integer = 0 To DataGridView1.Rows.Count - 1
                    If Trim(DataGridView1.Item(0, i).Value.ToString) = MyID Then
                        DataGridView1.CurrentCell = DataGridView1.Item(1, i)
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub ChangeApprovedState(ByVal MyGUID As String, ByVal CurrState As Boolean)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена состояния утверждения плана в БД
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "Update tbl_CRM_Events  "
        If CurrState = False Then
            MySQLStr = MySQLStr & "SET IsApproved = -1 "
        Else
            MySQLStr = MySQLStr & "SET IsApproved = 0 "
        End If
        MySQLStr = MySQLStr & "WHERE (EventID = '" & MyGUID & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class