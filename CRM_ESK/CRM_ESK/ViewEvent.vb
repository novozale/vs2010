Imports ADODB
Imports System.IO

Public Class ViewEvent

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения изменений
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Button2.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения изменений
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ViewEvent_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub ViewEvent_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы и значений в форму
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyDs As New DataSet

        MySQLStr = "SELECT tbl_CRM_Events.EventID, tbl_CRM_Directions.DirectionName, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_EventTypes.EventTypeID = 999999 THEN tbl_CRM_Events.EventTypeDescription ELSE tbl_CRM_EventTypes.EventTypeName END "
        MySQLStr = MySQLStr & "AS EventTypeName, tbl_CRM_Companies.ScalaCustomerCode, tbl_CRM_Companies.CompanyName, tbl_CRM_Companies.CompanyAddress, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_Actions.ActionID = 999999 THEN tbl_CRM_Events.ActionDescription ELSE tbl_CRM_Actions.ActionName END AS ActionName, "
        MySQLStr = MySQLStr & "tbl_CRM_Events.ActionPlannedDate, tbl_CRM_Events.ActionSumm, tbl_CRM_Events.ActionComments, tbl_CRM_Companies.CompanyPhone, "
        MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyEMail, tbl_CRM_Contacts.ContactName, tbl_CRM_Contacts.ContactPhone, tbl_CRM_Contacts.ContactEMail, "
        MySQLStr = MySQLStr & "CASE WHEN tbl_CRM_ActionsResultTypes.ActionResultID = 999999 THEN tbl_CRM_Events.ActionResultDescription ELSE tbl_CRM_ActionsResultTypes.ActionResultName "
        MySQLStr = MySQLStr & "END AS ActionResultName, ScalaSystemDB.dbo.ScaUsers.FullName, LTRIM(RTRIM(LTRIM(RTRIM(ISNULL(tbl_CRM_Projects.ProjectName, ''))) "
        MySQLStr = MySQLStr & "+ ' ' + LTRIM(RTRIM(ISNULL(tbl_CRM_Projects.ProjectComment, ''))))) AS ProjectInfo, ISNULL(tbl_CRM_Transport.TransportName, '') AS TransportName, "
        MySQLStr = MySQLStr & "ISNULL(tbl_CRM_Events.TransportDistance, 0) AS TransportDistance "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Transport WITH(NOLOCK) RIGHT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Events INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_EventTypes ON tbl_CRM_Events.EventTypeID = tbl_CRM_EventTypes.EventTypeID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Directions ON tbl_CRM_Events.DirectionID = tbl_CRM_Directions.DirectionID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Companies ON tbl_CRM_Events.CompanyID = tbl_CRM_Companies.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Contacts ON tbl_CRM_Events.ContactID = tbl_CRM_Contacts.ContactID AND "
        MySQLStr = MySQLStr & "tbl_CRM_Companies.CompanyID = tbl_CRM_Contacts.CompanyID INNER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Actions ON tbl_CRM_Events.ActionID = tbl_CRM_Actions.ActionID INNER JOIN "
        MySQLStr = MySQLStr & "ScalaSystemDB.dbo.ScaUsers ON tbl_CRM_Events.UserID = ScalaSystemDB.dbo.ScaUsers.UserID INNER JOIN "
        MySQLStr = MySQLStr & "ST010300 ON ScalaSystemDB.dbo.ScaUsers.FullName = ST010300.ST01002 ON "
        MySQLStr = MySQLStr & "tbl_CRM_Transport.TransportID = tbl_CRM_Events.TransportID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_ActionsResultTypes ON tbl_CRM_Events.ActionResultID = tbl_CRM_ActionsResultTypes.ActionResultID LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "tbl_CRM_Projects ON tbl_CRM_Events.ProjectID = tbl_CRM_Projects.ProjectID "
        MySQLStr = MySQLStr & "WHERE (tbl_CRM_Events.EventID = N'" & Declarations.MyEventID & "') "

        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            MsgBox("Данное действие не существует, возможно, удалено другим пользователем. Обновите данные в окне со списком действий.", MsgBoxStyle.Critical, "Внимание!")
            Me.Close()
        Else
            Declarations.MyRec.MoveFirst()
            TextBox6.Text = Declarations.MyRec.Fields("DirectionName").Value
            TextBox7.Text = Declarations.MyRec.Fields("EventTypeName").Value
            TextBox8.Text = Declarations.MyRec.Fields("CompanyName").Value
            TextBox9.Text = Trim(Trim(Declarations.MyRec.Fields("ContactName").Value.ToString) + " " + Trim(Declarations.MyRec.Fields("ContactPhone").Value.ToString) + " " + Trim(Declarations.MyRec.Fields("ContactEMail").Value.ToString))
            TextBox10.Text = Declarations.MyRec.Fields("ActionName").Value
            DateTimePicker1.Value = Declarations.MyRec.Fields("ActionPlannedDate").Value
            TextBox3.Text = Declarations.MyRec.Fields("ActionSumm").Value
            TextBox4.Text = Declarations.MyRec.Fields("ActionComments").Value
            TextBox11.Text = Declarations.MyRec.Fields("ActionResultName").Value.ToString
            TextBox1.Text = Declarations.MyRec.Fields("ProjectInfo").Value.ToString
            TextBox5.Text = Declarations.MyRec.Fields("TransportName").Value.ToString
            If Declarations.MyRec.Fields("TransportDistance").Value = 0 Then
                TextBox2.Text = ""
            Else
                TextBox2.Text = Declarations.MyRec.Fields("TransportDistance").Value.ToString
            End If
            trycloseMyRec()
            Button2.Select()
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
            Button6.Enabled = False
            Button8.Enabled = False
        Else
            Button6.Enabled = True
            Button8.Enabled = True
        End If
    End Function

    Private Sub Button6_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button6.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// просмотр аттачмента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ViewAttachment()
    End Sub

    Private Sub DataGridView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.DoubleClick
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// просмотр аттачмента
        '//
        '////////////////////////////////////////////////////////////////////////////////

        ViewAttachment()
    End Sub

    Private Sub Button8_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button8.Click
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
            trycloseMyRec()
            MsgBox(ex.ToString)
        End Try
    End Sub

    Private Sub ViewAttachment()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// просмотр аттачмента
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim mstream As ADODB.Stream
        Dim SavePath As String
        Dim DirectoryName As String

        Try
            MySQLStr = "SELECT AttachmentID, EventID, AttachmentName, AttachmentBody "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Attachments WITH(NOLOCK) "
            MySQLStr = MySQLStr & "WHERE (AttachmentID = '" & Trim(Me.DataGridView1.SelectedRows.Item(0).Cells(0).Value.ToString()) & "') "
            MySQLStr = MySQLStr & "ORDER BY AttachmentName "
            InitMyConn(False)
            InitMyRec(False, MySQLStr)
            DirectoryName = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())
            Directory.CreateDirectory(DirectoryName)
            SavePath = DirectoryName & "\" & Declarations.MyRec.Fields("AttachmentName").Value
            mstream = New ADODB.Stream
            mstream.Type = StreamTypeEnum.adTypeBinary
            mstream.Open()
            mstream.Write(Declarations.MyRec.Fields("AttachmentBody").Value)
            mstream.SaveToFile(SavePath, SaveOptionsEnum.adSaveCreateOverWrite)
            trycloseMyRec()
            Process.Start(SavePath)
        Catch ex As Exception
            trycloseMyRec()
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class