Public Class AddSalesOrder

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход из окна добавления заказа
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub AddSalesOrder_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход с сохранением данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            SaveData()
            Me.Close()
        End If
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка введенных данных - можно ли сохранять
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim OrderOvnerID As String                    'ID владельца заказа
        Dim OrrderOvnerFIO As String                  'ФИО владельца заказа
        Dim TeamFlag As Integer                       'флаг - 

        MySQLStr = "SELECT OR010300.OR01001, ISNULL(ScalaSystemDB.dbo.ScaUsers.UserID, 0) AS UserID, ISNULL(ScalaSystemDB.dbo.ScaUsers.FullName, '') AS FullName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers WITH(NOLOCK) RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "OR010300 ON ST010300.ST01001 = OR010300.OR01019 "
        MySQLStr = MySQLStr & "WHERE (OR010300.OR01002 <> 0) AND (OR010300.OR01001 = N'" & Microsoft.VisualBasic.Right("0000000000" & TextBox1.Text, 10) & "') "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT OR200300.OR20001, ISNULL(ScaUsers_1.UserID, 0) AS UserID, ISNULL(ScaUsers_1.FullName, '') AS FullName "
        MySQLStr = MySQLStr & "FROM ScalaSystemDB.dbo.ScaUsers AS ScaUsers_1 WITH(NOLOCK) RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "ST010300 AS ST010300_1 ON ScaUsers_1.FullName = ST010300_1.ST01002 RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "OR200300 ON ST010300_1.ST01001 = OR200300.OR20019 "
        MySQLStr = MySQLStr & "WHERE (OR200300.OR20001 = N'" & Microsoft.VisualBasic.Right("0000000000" & TextBox1.Text, 10) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '---Наличие заказа в Scala
            trycloseMyRec()
            MsgBox("Заказа на продажу с номером " & Microsoft.VisualBasic.Right("0000000000" & TextBox1.Text, 10) & " нет в базе данных или он находится в 0 типе.", MsgBoxStyle.Critical, "Внимание!")
            CheckData = False
            Exit Function
        Else
            If Declarations.MyRec.Fields("UserID").Value <> Declarations.UserID Then
                '---Принадлежность заказа продавцу
                OrderOvnerID = Declarations.MyRec.Fields("UserID").Value.ToString
                OrrderOvnerFIO = Declarations.MyRec.Fields("FullName").Value.ToString
                trycloseMyRec()
                '---тут проверяем тех, кто работает в команде.
                TeamFlag = 0
                MySQLStr = "SELECT ScaUsers_1.UserID as CC "
                MySQLStr = MySQLStr & "FROM tbl_SalesCommission_Groups INNER JOIN "
                MySQLStr = MySQLStr & "tbl_SalesCommission_Groups AS tbl_SalesCommission_Groups_1 ON "
                MySQLStr = MySQLStr & "tbl_SalesCommission_Groups.GroupName = tbl_SalesCommission_Groups_1.GroupName INNER JOIN "
                MySQLStr = MySQLStr & "ST010300 ON tbl_SalesCommission_Groups_1.SalesmanCode = ST010300.ST01001 INNER JOIN "
                MySQLStr = MySQLStr & "ST010300 AS ST010300_1 ON tbl_SalesCommission_Groups.SalesmanCode = ST010300_1.ST01001 INNER JOIN "
                MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON ST010300_1.ST01002 = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
                MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers AS ScaUsers_1 ON ST010300.ST01002 = ScaUsers_1.FullName "
                MySQLStr = MySQLStr & "WHERE (ScalaSystemDB.dbo.ScaUsers.UserID = " & Declarations.UserID & ") "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
                Else
                    Declarations.MyRec.MoveFirst()
                    While Declarations.MyRec.EOF = False
                        If Declarations.MyRec.Fields("CC").Value.ToString = OrderOvnerID Then
                            TeamFlag = 1
                        End If
                        Declarations.MyRec.MoveNext()
                    End While
                End If
                trycloseMyRec()

                If TeamFlag = 1 Then 'человек из команды
                Else
                    MsgBox("Заказ на продажу с номером " & Microsoft.VisualBasic.Right("0000000000" & TextBox1.Text, 10) & " принадлежит другому продавцу с ID = " & OrderOvnerID & " и Ф.И.О = " & OrrderOvnerFIO & ".", MsgBoxStyle.Critical, "Внимание!")
                    CheckData = False
                    Exit Function
                End If
            Else
                trycloseMyRec()
            End If
        End If
        MySQLStr = "SELECT tbl_CRM_Orders.OrderNum, tbl_CRM_Events.ActionPlannedDate "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Orders WITH(NOLOCK) INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Events ON tbl_CRM_Orders.EventID = tbl_CRM_Events.EventID "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Orders.OrderNum = N'" & Microsoft.VisualBasic.Right("0000000000" & TextBox1.Text, 10) & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            '---Не указан в качестве результата ни в одном другом действии
            trycloseMyRec()
        Else
            MsgBox("Заказ на продажу с номером " & Microsoft.VisualBasic.Right("0000000000" & TextBox1.Text, 10) & " уже указан как результат действия (продажа) от " & Declarations.MyRec.Fields("ActionPlannedDate").Value.ToString & ".", MsgBoxStyle.Critical, "Внимание!")
            trycloseMyRec()
            CheckData = False
            Exit Function
        End If
        CheckData = True
    End Function

    Private Sub SaveData()
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyGUID As Guid

        MyGUID = Guid.NewGuid
        Declarations.MyOrderID = MyGUID.ToString
        MySQLStr = "INSERT INTO tbl_CRM_Orders "
        MySQLStr = MySQLStr & "(ID, EventID, OrderNum) "
        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyOrderID & "', "
        MySQLStr = MySQLStr & "'" & Declarations.MyEventID & "', "
        MySQLStr = MySQLStr & "N'" & Microsoft.VisualBasic.Right("0000000000" & TextBox1.Text, 10) & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class