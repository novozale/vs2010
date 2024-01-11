Public Class CustomerExtInfo

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения изменений
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Declarations.MyResult = 0
        Me.Close()
    End Sub

    Private Sub CustomerExtInfo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка формы и данных в нее
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        MySQLStr = "SELECT * "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Companies_Ext "
        MySQLStr = MySQLStr & "WHERE (CompanyID = N'" & Declarations.MyClientID & "') "
        InitMyConn(False)
        InitMyRec(False, MySQLStr)
        If Declarations.MyRec.EOF = True And Declarations.MyRec.BOF = True Then
            trycloseMyRec()
            ComboBox1.Text = ""
            TextBox1.Text = ""
        Else
            ComboBox1.Text = Trim(Declarations.MyRec.Fields("IsIKA").Value)
            TextBox1.Text = Declarations.MyRec.Fields("Potencial").Value
            trycloseMyRec()
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение введенных данных и выход из формы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка

        If CheckData() = True Then

            '---таблица tbl_CRM_Companies_Ext
            MySQLStr = "DELETE FROM tbl_CRM_Companies_Ext "
            MySQLStr = MySQLStr & "WHERE (CompanyID = N'" & Declarations.MyClientID & "') "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            If Not Trim(ComboBox1.Text).Equals("") Or Not Trim(TextBox1.Text).Equals("") Then
                MySQLStr = "INSERT INTO tbl_CRM_Companies_Ext "
                MySQLStr = MySQLStr & "(CompanyID, IsIKA, Potencial) "
                MySQLStr = MySQLStr & "VALUES (N'" & Declarations.MyClientID & "', "
                MySQLStr = MySQLStr & "N'" & Trim(ComboBox1.Text) & "', "
                MySQLStr = MySQLStr & Replace(TextBox1.Text, ",", ".") & ") "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If

            '---таблица tbl_CustomerCard0300
            MySQLStr = "UPDATE tbl_CustomerCard0300 "
            MySQLStr = MySQLStr & "SET IKA = tbl_RexelIKATypes.ID "
            MySQLStr = MySQLStr & "FROM tbl_CustomerCard0300 INNER JOIN "
            MySQLStr = MySQLStr & "(SELECT ScalaCustomerCode, N'" & Trim(ComboBox1.Text) & "' AS IKAName "
            MySQLStr = MySQLStr & "FROM tbl_CRM_Companies "
            MySQLStr = MySQLStr & "WHERE (CompanyID = N'" & Declarations.MyClientID & "')) "
            MySQLStr = MySQLStr & "AS View_12 ON tbl_CustomerCard0300.SL01001 = View_12.ScalaCustomerCode INNER JOIN "
            MySQLStr = MySQLStr & "tbl_RexelIKATypes ON View_12.IKAName = tbl_RexelIKATypes.Name "
            InitMyConn(False)
            Declarations.MyConn.Execute(MySQLStr)

            Declarations.MyResult = 1
            Me.Close()

        End If
    End Sub

    Private Function CheckData() As Boolean
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка правильности заполнения полей
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MyDbl As Double

        If Not Trim(TextBox1.Text).Equals("") Then
            Try
                MyDbl = TextBox1.Text
            Catch ex As Exception
                MsgBox("Потенциал (Руб) должнен быть заполнен числом.", MsgBoxStyle.OkOnly, "Внимание!")
                CheckData = False
                TextBox1.Select()
                Exit Function
            End Try
        End If

        CheckData = True
    End Function
End Class