Imports ADODB
Imports System.Net.Mail

Module Functions
    Public Sub InitMyConn()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Инициализация соединения с БД, чтение глобальных переменных
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        On Error GoTo MyCatch
        If MyConn Is Nothing Then
            MyConn = New ADODB.Connection
            MyConn.CursorLocation = 3
            MyConn.CommandTimeout = 600
            MyConn.ConnectionTimeout = 300
            If Declarations.MyConnStr = "" Then
                'Declarations.MyConnStr = "Provider=SQLOLEDB;User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=SPBDVL2"
                Declarations.MyConnStr = "Provider=SQLOLEDB;User ID=sa;Password=sqladmin;DATABASE=ScaDataDB;SERVER=SQLCLS"
                Declarations.MyNETConnStr = Replace(Declarations.MyConnStr, "Provider=SQLOLEDB;", "")
                Declarations.MyNETConnStr = Declarations.MyNETConnStr & ";Timeout=0;"
            End If
            MyConn.Open(Declarations.MyConnStr)
        End If
        Exit Sub
MyCatch:
        EventLog.WriteEntry("CRM_ESK_DailyJob", "Ошибка Functions 1")
    End Sub

    Public Sub trycloseMyRec()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '//Попытка закрытия рекордсета
        '//
        '////////////////////////////////////////////////////////////////////////////////////////

        On Error Resume Next
        MyRec.Close()
    End Sub

    Public Sub InitMyRec(ByVal sql As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//Открытие рекордсета
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyErr

        On Error GoTo MyCatch
        InitMyConn()
        If MyRec Is Nothing Then
            MyRec = New ADODB.Recordset
        End If
        trycloseMyRec()
        MyRec.LockType = LockTypeEnum.adLockOptimistic
        MyRec.Open(sql, MyConn)
        If MyConn.Errors.Count > 0 Then
            For Each MyErr In MyConn.Errors
                Err.Raise(MyErr.Number, MyErr.Source, MyErr.Description)
            Next MyErr
        End If
        Exit Sub
MyCatch:
        EventLog.WriteEntry("CRM_ESK_DailyJob", "Ошибка Functions 2")
    End Sub

    Public Sub SendMyReminder(ByVal Subject As String, ByVal MyWrkString As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// отправка сообщения о ошибке по почте в ИТ
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Try
            Dim smtp As SmtpClient = New SmtpClient(My.Settings.SMTPService)
            Dim msg As New MailMessage

            msg.To.Add(My.Settings.MessageTo)
            If Trim(My.Settings.MessageCC) <> "" Then
                msg.CC.Add(My.Settings.MessageCC)
            End If
            msg.From = New MailAddress(My.Settings.MessageFrom)
            msg.Subject = Subject
            msg.Body = MyWrkString
            smtp.Send(msg)
        Catch ex As Exception
            EventLog.WriteEntry("CRM_ESK_DailyJob", ex.Message)
        End Try
    End Sub

    Public Sub CheckProjects()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка состояния проектов
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyGuid As Guid
        Dim MyProjectID As String
        Dim MyCompanyID As String
        Dim MyProjectName As String
        Dim MyProjectComment As String
        Dim MyProjectAddr As String
        Dim MyResponciblePerson As String
        Dim MyManufacturersList As String
        Dim MyEmail As String
        Dim MyCompanyName As String
        Dim MyCompanyAddress As String
        Dim MyUserID As Integer
        Dim MyContactID As String
        Dim MyContactName As String
        Dim MySchedDate As DateTime
        Dim MyOldUser As Integer
        Dim MyOldDate As DateTime
        Dim MyDelayInDays As Integer

        MyDelayInDays = My.Settings.MyDelayInDays
        '--------------Получение списка незакрытых проектов------------------------------
        '--------------с последним событием старше MyDelayInDays дней--------------------
        MySQLStr = "SELECT tbl_CRM_Projects_2.ProjectID, tbl_CRM_Projects_2.CompanyID, "
        MySQLStr = MySQLStr & "tbl_CRM_Projects_2.ProjectName, tbl_CRM_Projects_2.ProjectComment, "
        MySQLStr = MySQLStr & "tbl_CRM_Projects_2.ProjectAddr, "
        MySQLStr = MySQLStr & "tbl_CRM_Projects_2.ResponciblePerson, tbl_CRM_Projects_2.ManufacturersList, "
        MySQLStr = MySQLStr & "tbl_CRM_Projects_2.IsApproved, tbl_CRM_Projects_2.IsIPG, RM.dbo.RM660100.RM66003 AS Email, "
        MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyName, tbl_CRM_Companies.CompanyAddress, "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers.UserID, ISNULL(View_5.ContactID, "
        MySQLStr = MySQLStr & "View_6.ContactID) AS ContactID, "
        MySQLStr = MySQLStr & "ISNULL(View_5.ContactName, View_6.ContactName) AS ContactName "
        MySQLStr = MySQLStr & "FROM (SELECT ProjectID, MAX(ActDate) AS ActDate "
        MySQLStr = MySQLStr & "FROM (SELECT tbl_CRM_Projects_History.ProjectID, tbl_CRM_Projects_History.ActDate "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Projects_History INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects ON tbl_CRM_Projects_History.ProjectID = tbl_CRM_Projects.ProjectID "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Projects.CloseDate IS NULL) AND (tbl_CRM_Projects.IsApproved = 1) "
        MySQLStr = MySQLStr & "UNION "
        MySQLStr = MySQLStr & "SELECT tbl_CRM_Events.ProjectID, tbl_CRM_Events.ActionPlannedDate AS ActDate "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects AS tbl_CRM_Projects_1 ON tbl_CRM_Events.ProjectID = tbl_CRM_Projects_1.ProjectID "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Projects_1.CloseDate IS NULL) AND (tbl_CRM_Projects_1.IsApproved = 1)) AS View_2 "
        MySQLStr = MySQLStr & "GROUP BY ProjectID) AS View_3 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects AS tbl_CRM_Projects_2 ON View_3.ProjectID = tbl_CRM_Projects_2.ProjectID INNER JOIN "
        MySQLStr = MySQLStr & "RM.dbo.RM660100 ON tbl_CRM_Projects_2.ResponciblePerson = RM.dbo.RM660100.RM66002 INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Projects_2.ResponciblePerson = ScalaSystemDB.dbo.ScaUsers.FullName INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Projects_2.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT tbl_CRM_Contacts.CompanyID, tbl_CRM_Contacts.ContactID, tbl_CRM_Contacts.ContactName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Contacts INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT CompanyID, MAX(CreationDate) AS CreationDate "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Contacts AS tbl_CRM_Contacts_1 "
        MySQLStr = MySQLStr & "GROUP BY CompanyID) AS View_8 ON tbl_CRM_Contacts.CreationDate = View_8.CreationDate "
        MySQLStr = MySQLStr & "AND tbl_CRM_Contacts.CompanyID = View_8.CompanyID "
        MySQLStr = MySQLStr & "GROUP BY tbl_CRM_Contacts.ContactID, tbl_CRM_Contacts.ContactName, tbl_CRM_Contacts.CompanyID) "
        MySQLStr = MySQLStr & "AS View_6 ON tbl_CRM_Companies.CompanyID = View_6.CompanyID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT TOP (100) PERCENT View_7.ProjectID, tbl_CRM_Contacts_1.ContactID, tbl_CRM_Contacts_1.ContactName "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events AS tbl_CRM_Events_2 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Contacts AS tbl_CRM_Contacts_1 ON tbl_CRM_Events_2.ContactID = tbl_CRM_Contacts_1.ContactID INNER JOIN "
        MySQLStr = MySQLStr & "(SELECT TOP (100) PERCENT tbl_CRM_Projects_3.ProjectID, MAX(tbl_CRM_Events_1.ActionPlannedDate) AS ActionPlannedDate "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Events AS tbl_CRM_Events_1 INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects AS tbl_CRM_Projects_3 ON tbl_CRM_Events_1.ProjectID = tbl_CRM_Projects_3.ProjectID "
        MySQLStr = MySQLStr & "GROUP BY tbl_CRM_Projects_3.ProjectID) AS View_7 ON tbl_CRM_Events_2.ProjectID = View_7.ProjectID AND "
        MySQLStr = MySQLStr & "tbl_CRM_Events_2.ActionPlannedDate = View_7.ActionPlannedDate "
        MySQLStr = MySQLStr & "GROUP BY tbl_CRM_Contacts_1.ContactID, tbl_CRM_Contacts_1.ContactName, View_7.ProjectID "
        MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Contacts_1.ContactID) AS View_5 ON tbl_CRM_Projects_2.ProjectID = View_5.ProjectID "
        MySQLStr = MySQLStr & "WHERE (RM.dbo.RM660100.RM66003 <> '') AND (View_3.ActDate < DATEADD(dd, -" & CStr(MyDelayInDays) & "1, GETDATE())) "
        MySQLStr = MySQLStr & "AND (ScalaSystemDB.dbo.ScaUsers.IsBlocked = 0) AND (tbl_CRM_Projects_2.IsApproved = 1) "
        MySQLStr = MySQLStr & "ORDER BY tbl_CRM_Projects_2.ResponciblePerson"
        InitMyConn()
        InitMyRec(MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
        Else
            Declarations.MyRec.MoveFirst()
            MyOldUser = 0
            While Declarations.MyRec.EOF = False
                '------------------------получение данных из запроса---------------------
                MyGuid = Guid.NewGuid()
                MyProjectID = Declarations.MyRec.Fields("ProjectID").Value
                MyCompanyID = Declarations.MyRec.Fields("CompanyID").Value
                MyProjectName = Declarations.MyRec.Fields("ProjectName").Value
                MyProjectComment = Declarations.MyRec.Fields("ProjectComment").Value
                MyProjectAddr = Declarations.MyRec.Fields("ProjectAddr").Value
                MyResponciblePerson = Declarations.MyRec.Fields("ResponciblePerson").Value
                MyManufacturersList = Declarations.MyRec.Fields("ManufacturersList").Value
                MyEmail = Declarations.MyRec.Fields("Email").Value
                MyCompanyName = Declarations.MyRec.Fields("CompanyName").Value
                MyCompanyAddress = Declarations.MyRec.Fields("CompanyAddress").Value
                MyUserID = Declarations.MyRec.Fields("UserID").Value
                MyContactID = Declarations.MyRec.Fields("ContactID").Value
                MyContactName = Declarations.MyRec.Fields("ContactName").Value
                If MyUserID = MyOldUser Then
                    '---несколько проектов в один день
                    MySchedDate = DateAdd(DateInterval.Day, 1, MyOldDate)
                    Do While Weekday(MySchedDate, FirstDayOfWeek.Monday) = 6 Or Weekday(MySchedDate, FirstDayOfWeek.Monday) = 7
                        MySchedDate = DateAdd(DateInterval.Day, 1, MySchedDate)
                        If (MySchedDate - Today).TotalDays > 30 Then
                            MySchedDate = DateAdd(DateInterval.Day, 5, Today)
                        End If
                    Loop
                Else
                    '---только 1 проект
                    MySchedDate = DateAdd(DateInterval.Day, 5, Today)
                    Do While Weekday(MySchedDate, FirstDayOfWeek.Monday) = 6 Or Weekday(MySchedDate, FirstDayOfWeek.Monday) = 7
                        MySchedDate = DateAdd(DateInterval.Day, 1, MySchedDate)
                    Loop
                End If
                MyOldUser = MyUserID
                MyOldDate = MySchedDate

                '--------------------создание действия в CRM
                MySQLStr = "INSERT INTO tbl_CRM_Events "
                MySQLStr = MySQLStr & "(EventID, DirectionID, EventTypeID, EventTypeDescription, CompanyID, "
                MySQLStr = MySQLStr & "ContactID, ActionTime, ActionID, ActionDescription, ActionPlannedDate, "
                MySQLStr = MySQLStr & "ActionSumm, ActionComments, ActionResultID, ActionResultDescription, UserID, "
                MySQLStr = MySQLStr & "OwnerID, ActionClosed, ProjectID, TransportID, TransportDistance, IsApproved) "
                MySQLStr = MySQLStr & "VALUES ('" & MyGuid.ToString & "', "                     '--ID действия
                MySQLStr = MySQLStr & "2, "                                                     '--Направление действия
                MySQLStr = MySQLStr & "1, "                                                     '--Способ контакта
                MySQLStr = MySQLStr & "NULL, "                                                  '--Дополнительный способ контакта
                MySQLStr = MySQLStr & "'" & MyCompanyID & "', "                                 '--клиент
                MySQLStr = MySQLStr & "'" & MyContactID & "', "                                 '--контакт
                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & DatePart(DateInterval.Day, Now()) & "/" & DatePart(DateInterval.Month, Now()) & "/" & DatePart(DateInterval.Year, Now()) & "', 103), " '--дата создания записи фактическая
                MySQLStr = MySQLStr & "999999, "                                                '--действие
                MySQLStr = MySQLStr & "'Обновление информации о проекте', "                     '--другое действие
                MySQLStr = MySQLStr & "CONVERT(DATETIME, '" & DatePart(DateInterval.Day, MySchedDate) & "/" & DatePart(DateInterval.Month, MySchedDate) & "/" & DatePart(DateInterval.Year, MySchedDate) & "', 103), " '--плановая дата выполнения действия
                MySQLStr = MySQLStr & "0, "                                                     '--сумма действия
                MySQLStr = MySQLStr & "'', "                                                    '--комментарий
                MySQLStr = MySQLStr & "NULL, "                                                  '--результат действия
                MySQLStr = MySQLStr & "NULL, "
                MySQLStr = MySQLStr & MyUserID.ToString & ", "                                  '--ID пользователя
                MySQLStr = MySQLStr & MyUserID.ToString & ", "                                  '--ID владельца
                MySQLStr = MySQLStr & "NULL, "                                                  '--Дата закрытия действия 
                MySQLStr = MySQLStr & "'" & MyProjectID & "', "                                 '--проект
                MySQLStr = MySQLStr & "NULL, "                                                  '--транспорт 
                MySQLStr = MySQLStr & "NULL, "
                MySQLStr = MySQLStr & "1) "
                InitMyConn()
                Declarations.MyConn.Execute(MySQLStr)

                '------------------Занесение информации в календарь----------------------
                If CreateCalendarEvent(MyGuid.ToString) = False Then
                    MySQLStr = "Exec spp_CRM_SendAppointment N'" & MyGuid.ToString & "' "
                    InitMyConn()
                    Declarations.MyConn.Execute(MySQLStr)
                End If

                '------------------Отправка письма с уведомлением о создании задачи------
                SendInfoByEmail(MyEmail, MyResponciblePerson, MySchedDate, MyCompanyName, MyCompanyAddress,
                    MyProjectName, MyProjectAddr, MyProjectComment, MyManufacturersList,
                    MyContactName)

                'Console.WriteLine(MyGuid.ToString & " _ " & MyCompanyName & " _ " & MyProjectName & " _ " & MyEmail & " _ " & MySchedDate.ToString)
                'Console.WriteLine(MySchedDate.ToString)
                Declarations.MyRec.MoveNext()
            End While
        End If

    End Sub

    Public Function CreateCalendarEvent(ByVal MyEventID As String) As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Создание события в календаре в офисе 365
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyObj As New spbadm4_EWS.EskEWSServiceClient
        Dim MyCalendarEvent As New spbadm4_EWS.CreateCalendarEventType
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
        InitMyConn()
        InitMyRec(MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            Return False
        Else
            MyCalendarEvent.CalendarEventIDOld = Declarations.MyRec.Fields("CalEventID").Value

            MyCalendarEvent.Subject = Declarations.MyRec.Fields("DirectionName").Value & " "
            MyCalendarEvent.Subject = MyCalendarEvent.Subject & Declarations.MyRec.Fields("EventTypeName").Value
            MyCalendarEvent.Subject = MyCalendarEvent.Subject & " Компания " & Declarations.MyRec.Fields("CompanyName").Value

            MyCalendarEvent.Body = "<p>" & Declarations.MyRec.Fields("DirectionName").Value & " "
            MyCalendarEvent.Body = MyCalendarEvent.Body & Declarations.MyRec.Fields("EventTypeName").Value & "</p>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & "<p> Компания " & Declarations.MyRec.Fields("CompanyName").Value & "<br>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & "Контакт: " & Declarations.MyRec.Fields("ContactName").Value & "<br>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & " Телефон: " & Declarations.MyRec.Fields("ContactPhone").Value & "<br>"
            MyCalendarEvent.Body = MyCalendarEvent.Body & " Email: " & Declarations.MyRec.Fields("ContactEMail").Value & "<br>"
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
                    InitMyConn()
                    Declarations.MyConn.Execute(MySQLStr)
                    MySQLStr = "INSERT INTO tbl_CRM_EventsInCalendar "
                    MySQLStr = MySQLStr & "(EventID, CalEventID) "
                    MySQLStr = MySQLStr & "VALUES ('" & MyEventID & "', "
                    MySQLStr = MySQLStr & "N'" & MyRez & "') "
                    InitMyConn()
                    Declarations.MyConn.Execute(MySQLStr)
                    Return True
                End If
            Catch ex As Exception
                Return False
            End Try
        End If
    End Function

    Private Sub SendInfoByEmail(ByVal ToEmail As String, ByVal MyResponciblePerson As String, ByVal MySchedDate As DateTime,
                                ByVal MyCompanyName As String, ByVal MyCompanyAddress As String, ByVal MyProjectName As String,
                                ByVal MyProjectAddr As String, ByVal MyProjectComment As String, ByVal MyManufacturersList As String,
                                ByVal MyContactName As String)
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Отправка по почте уведомления о создании действия в CRM
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim smtp As Net.Mail.SmtpClient
        Dim msg As Net.Mail.MailMessage
        Dim MyMsgStr As String

        smtp = New Net.Mail.SmtpClient(My.Settings.SMTPService)
        msg = New Net.Mail.MailMessage
        msg.To.Add(ToEmail)
        msg.From = New Net.Mail.MailAddress("reportserver@elektroskandia.ru")
        msg.Subject = "Уведомление о назначении задачи в CRM"

        MyMsgStr = "Уважаемый " & MyResponciblePerson & "!" & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "Вам на " & Format(MySchedDate, "dd/MM/yyyy") & "в CRM сформирована задача " & Chr(13)
        MyMsgStr = MyMsgStr & "Уточнить и обновить статус проекта: " & Chr(13) & Chr(13)
        MyMsgStr = MyMsgStr & "Клиент: " & MyCompanyName & "." & Chr(13)
        MyMsgStr = MyMsgStr & "Адрес клиента: " & MyCompanyAddress & "." & Chr(13)
        MyMsgStr = MyMsgStr & "Проект: " & MyProjectName & "." & Chr(13)
        MyMsgStr = MyMsgStr & "Адрес проекта: " & MyProjectAddr & "." & Chr(13)
        MyMsgStr = MyMsgStr & "Комментарий к проекту: " & MyProjectComment & "." & Chr(13)
        MyMsgStr = MyMsgStr & "Производители в проекте: " & MyManufacturersList & "." & Chr(13)
        MyMsgStr = MyMsgStr & "Контактное лицо: " & MyContactName & Chr(13) & "." & Chr(13)
        MyMsgStr = MyMsgStr + "_______________________________" & Chr(13)
        MyMsgStr = MyMsgStr + "С уважением," & Chr(13)
        MyMsgStr = MyMsgStr + "ООО ""Электроскандия Рус"". " & Chr(13) & Chr(13)
        msg.Body = MyMsgStr
        smtp.Send(msg)
    End Sub
End Module
