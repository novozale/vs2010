Public Class SendReturnAction
    Public MyUserID As Integer = -1
    Public MyEventID As String = ""

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход из формы
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub SendReturnAction_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка окна, загрузка информации
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка продавцов
        Dim MyDs As New DataSet
        Dim MyAdapter1 As SqlClient.SqlDataAdapter    'для списка продавцов
        Dim MyDs1 As New DataSet
        Dim MyAdapter2 As SqlClient.SqlDataAdapter    'для списка продавцов
        Dim MyDs2 As New DataSet

        '---Список продавцов
        If Declarations.MyPermission = True Then
            '---Доступны все продавцы
            MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
            MySQLStr = MySQLStr & "WHERE(ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "
            MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        ElseIf Declarations.MyCCPermission = True Then
            '---Доступны продавцы определенного кост центра
            MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
            MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
            MySQLStr = MySQLStr & "(tbl_CRM_CCOwners.CCOwn = N'" & Declarations.CC & "') "
            MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        Else
            '---только один продавец (вошедший в систему)
            MySQLStr = "SELECT ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
            MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
            MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
            MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND "
            MySQLStr = MySQLStr & "(ScalaSystemDB.dbo.ScaUsers.UserID = " & Declarations.UserID & ") "
            MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        End If

        '---От кого передается
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

        '---кому возвращается
        Try
            MyAdapter2 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter2.SelectCommand.CommandTimeout = 600
            MyAdapter2.Fill(MyDs2)
            ComboBox5.DisplayMember = "FullName" 'Это то что будет отображаться
            ComboBox5.ValueMember = "UserID"   'это то что будет храниться
            ComboBox5.DataSource = MyDs2.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---Кому передается
        '---Доступны все продавцы
        MySQLStr = "SELECT Distinct ScalaSystemDB.dbo.ScaUsers.UserID, ScalaSystemDB.dbo.ScaUsers.FullName "
        MySQLStr = MySQLStr & "FROM ST010300 WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_CCOwners ON SUBSTRING(ST010300.ST01021, 7, 3) = tbl_CRM_CCOwners.CCSubord "
        MySQLStr = MySQLStr & "WHERE(ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) "
        MySQLStr = MySQLStr & "ORDER BY ScalaSystemDB.dbo.ScaUsers.FullName "
        Try
            MyAdapter1 = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
            MyAdapter1.SelectCommand.CommandTimeout = 600
            MyAdapter1.Fill(MyDs1)
            ComboBox2.DisplayMember = "FullName" 'Это то что будет отображаться
            ComboBox2.ValueMember = "UserID"   'это то что будет храниться
            ComboBox2.DataSource = MyDs1.Tables(0).DefaultView
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        '---Значения по умолчанию
        If MyUserID <> -1 Then
            ComboBox1.SelectedValue = MyUserID
            ComboBox5.SelectedValue = MyUserID
        End If

        GetTransferredActivities()
        GeBackTransferredActivities()

        If MyEventID.Equals("") = False Then
            ComboBox3.SelectedValue = MyEventID
        End If

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена выбора пользователя, от которого передавать активности
        '//
        '////////////////////////////////////////////////////////////////////////////////

        GetTransferredActivities()
        ComboBox2.Select()
    End Sub

    Private Sub GetTransferredActivities()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение списка передаваемых от пользователя активностей
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If ComboBox1.SelectedValue = Nothing Then
        Else
            MySQLStr = "SELECT ' Все' AS EventID, 'Все активности' AS Event "
            MySQLStr = MySQLStr & "Union "
            MySQLStr = MySQLStr & "SELECT Convert(nvarchar(50), tbl_CRM_Events.EventID) as EventID, tbl_CRM_Directions.DirectionName + ' ' + "
            MySQLStr = MySQLStr & "tbl_CRM_EventTypes.EventTypeName + ' ' + LTRIM(RTRIM(ISNULL(tbl_CRM_Companies.ScalaCustomerCode, ''))) "
            MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(tbl_CRM_Companies.CompanyName)) + ' ' + CONVERT(nvarchar(30), tbl_CRM_Events.ActionPlannedDate, 103) AS Event "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Events INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Directions ON tbl_CRM_Events.DirectionID = tbl_CRM_Directions.DirectionID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_EventTypes ON tbl_CRM_Events.EventTypeID = tbl_CRM_EventTypes.EventTypeID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID "
            MySQLStr = MySQLStr & "WHERE (tbl_CRM_Events.UserID = " & ComboBox1.SelectedValue & ") "
            MySQLStr = MySQLStr & "AND (tbl_CRM_Events.ActionResultID IS NULL) "
            MySQLStr = MySQLStr & "Order By Event "
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                ComboBox3.DisplayMember = "Event" 'Это то что будет отображаться
                ComboBox3.ValueMember = "EventID"   'это то что будет храниться
                ComboBox3.DataSource = MyDs.Tables(0).DefaultView
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            ComboBox3.SelectedValue = " Все"
        End If
    End Sub

    Private Sub ComboBox5_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.SelectedValueChanged
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Смена выбора пользователя, кому возвращать активности
        '//
        '////////////////////////////////////////////////////////////////////////////////

        GeBackTransferredActivities()
        ComboBox4.Select()
    End Sub

    Private Sub GeBackTransferredActivities()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// получение списка возвращаемых пользователю активностей
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyAdapter As SqlClient.SqlDataAdapter     '
        Dim MyDs As New DataSet                       '

        If ComboBox5.SelectedValue = Nothing Then
        Else
            MySQLStr = "SELECT ' Все' AS EventID, 'Все активности' AS Event "
            MySQLStr = MySQLStr & "Union "
            MySQLStr = MySQLStr & "SELECT Convert(nvarchar(50), tbl_CRM_Events.EventID) as EventID, tbl_CRM_Directions.DirectionName + ' ' + "
            MySQLStr = MySQLStr & "tbl_CRM_EventTypes.EventTypeName + ' ' + LTRIM(RTRIM(ISNULL(tbl_CRM_Companies.ScalaCustomerCode, ''))) "
            MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(tbl_CRM_Companies.CompanyName)) + ' ' + CONVERT(nvarchar(30), tbl_CRM_Events.ActionPlannedDate, 103) AS Event "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Events INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Directions ON tbl_CRM_Events.DirectionID = tbl_CRM_Directions.DirectionID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_EventTypes ON tbl_CRM_Events.EventTypeID = tbl_CRM_EventTypes.EventTypeID INNER JOIN "
            MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID "
            MySQLStr = MySQLStr & "WHERE (tbl_CRM_Events.OwnerID = " & ComboBox5.SelectedValue & ") "
            MySQLStr = MySQLStr & "AND (tbl_CRM_Events.UserID <> " & ComboBox5.SelectedValue & ") "
            MySQLStr = MySQLStr & "AND (tbl_CRM_Events.ActionResultID IS NULL) "
            MySQLStr = MySQLStr & "Order By Event "
            Try
                MyAdapter = New SqlClient.SqlDataAdapter(MySQLStr, Declarations.MyNETConnStr)
                MyAdapter.SelectCommand.CommandTimeout = 600
                MyAdapter.Fill(MyDs)
                ComboBox4.DisplayMember = "Event" 'Это то что будет отображаться
                ComboBox4.ValueMember = "EventID"   'это то что будет храниться
                ComboBox4.DataSource = MyDs.Tables(0).DefaultView
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
            ComboBox4.SelectedValue = " Все"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Передача владения действиями
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim IDList As New List(Of String)
        Dim i As Integer

        '-----Список для изменения в календарь
        MySQLStr = "SELECT EventID "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events "
        MySQLStr = MySQLStr & "WHERE (UserID = " & ComboBox1.SelectedValue & ") "
        MySQLStr = MySQLStr & "AND (tbl_CRM_Events.ActionResultID IS NULL) "
        If ComboBox3.SelectedValue.Equals(" Все") = False Then
            MySQLStr = MySQLStr & "And (EventID  = '" & ComboBox3.SelectedValue & "') "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                IDList.Add(Declarations.MyRec.Fields("EventID").Value)
                Declarations.MyRec.MoveNext()
            End While
        End If

        '-----изменения в события
        MySQLStr = "UPDATE tbl_CRM_Events "
        MySQLStr = MySQLStr & "SET UserID = " & ComboBox2.SelectedValue & " "
        MySQLStr = MySQLStr & "WHERE (UserID = " & ComboBox1.SelectedValue & ") "
        MySQLStr = MySQLStr & "AND (tbl_CRM_Events.ActionResultID IS NULL) "
        If ComboBox3.SelectedValue.Equals(" Все") = False Then
            MySQLStr = MySQLStr & "And (EventID  = '" & ComboBox3.SelectedValue & "') "
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '-----изменения в календарь
        For i = 0 To IDList.Count - 1
            AddEvent.CreateCalendarEvent(IDList(i))
            i = i + 1
        Next

        MsgBox("Передача владения действиями осуществлена.", MsgBoxStyle.OkOnly, "Внимание!")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Возвращение владения действиями
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim IDList As New List(Of String)
        Dim i As Integer

        '-----Список для изменения в календарь
        MySQLStr = "SELECT EventID "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events "
        MySQLStr = MySQLStr & "WHERE (OwnerID = " & ComboBox5.SelectedValue & ") "
        MySQLStr = MySQLStr & "AND (UserID <> OwnerID) "
        MySQLStr = MySQLStr & "AND (tbl_CRM_Events.ActionResultID IS NULL) "
        If ComboBox3.SelectedValue.Equals(" Все") = False Then
            MySQLStr = MySQLStr & "And (EventID  = '" & ComboBox3.SelectedValue & "') "
        End If
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
        Else
            Declarations.MyRec.MoveFirst()
            While Declarations.MyRec.EOF = False
                IDList.Add(Declarations.MyRec.Fields("EventID").Value)
                Declarations.MyRec.MoveNext()
            End While
        End If

        '-----изменения в события
        MySQLStr = "UPDATE tbl_CRM_Events "
        MySQLStr = MySQLStr & "SET UserID = OwnerID "
        MySQLStr = MySQLStr & "WHERE (OwnerID = " & ComboBox5.SelectedValue & ") "
        MySQLStr = MySQLStr & "AND (UserID <> OwnerID) "
        MySQLStr = MySQLStr & "AND (tbl_CRM_Events.ActionResultID IS NULL) "
        If ComboBox4.SelectedValue.Equals(" Все") = False Then
            MySQLStr = MySQLStr & "And (EventID  = '" & ComboBox4.SelectedValue & "') "
        End If
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '-----изменения в календарь
        For i = 0 To IDList.Count - 1
            AddEvent.CreateCalendarEvent(IDList(i))
            i = i + 1
        Next

        MsgBox("Возврат владения действиями осуществлен.", MsgBoxStyle.OkOnly, "Внимание!")
    End Sub
End Class