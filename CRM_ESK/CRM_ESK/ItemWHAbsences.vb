Public Class ItemWHAbsences

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ItemWHAbsences_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub ItemWHAbsences_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в окно по открытию формы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Label15.Text = MyWHAbsences.ComboBox1.SelectedValue
        Label9.Text = Trim(MyWHAbsences.DataGridView1.SelectedRows.Item(0).Cells(0).Value)
        Label10.Text = Trim(MyWHAbsences.DataGridView1.SelectedRows.Item(0).Cells(3).Value)
        Label11.Text = Trim(MyWHAbsences.DataGridView1.SelectedRows.Item(0).Cells(6).Value)
        TextBox5.Text = Trim(MyWHAbsences.DataGridView1.SelectedRows.Item(0).Cells(5).Value)
        TextBox5.Enabled = False
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

    Private Sub TextBox4_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox4.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox4.Text) <> "" Then
            If InStr(TextBox4.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Запрошенное покупателем количество"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox4.Text
                Catch ex As Exception
                    MsgBox("В поле ""Запрошенное покупателем количество"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub TextBox5_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox5.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox5.Text) <> "" Then
            If InStr(TextBox5.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Доступно к заказу на всех складах"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("В поле ""Доступно к заказу на всех складах"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// сохранение данных
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckDataFiling() = True Then
            SaveData()
            Me.Close()
        End If
    End Sub

    Private Function CheckDataFiling() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Проверка заполнения полей в окне
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If Trim(TextBox4.Text) = "" Then
            MsgBox("Поле ""Запрошенное покупателем количество"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox4.Select()
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" Then
            MsgBox("Поле ""Доступно к заказу на всех складах"" должно быть заполнено.", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox5.Select()
            Exit Function
        End If

        If Trim(TextBox1.Text) = "" And Trim(TextBox2.Text) <> "" Then
            MsgBox("Поле ""Другой код запаса поставщика"" должно быть заполнено, если заполнено поле ""Название запаса (нет в Scala)"".", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox1.Select()
            Exit Function
        End If

        If Trim(TextBox2.Text) = "" And Trim(TextBox1.Text) <> "" Then
            MsgBox("Поле ""Название запаса (нет в Scala)"" должно быть заполнено, если заполнено поле ""Другой код запаса поставщика"".", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox2.Select()
            Exit Function
        End If

        CheckDataFiling = True

    End Function

    Private Sub SaveData()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Сохранение данных
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String
        Dim MyGUID As Guid

        MyGUID = Guid.NewGuid
        Declarations.MyWHAbsencesID = MyGUID.ToString
        MySQLStr = "INSERT INTO tbl_CRM_WHAbsences "
        MySQLStr = MySQLStr & "(ID, EventID, WH, ScalaCode, SupplierCode, OtherSupplierCode, OtherItem, ScalaSupplierCode, OtherSupplier, RequestedQTY, AvailableQTY ) "
        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyWHAbsencesID & "', "
        MySQLStr = MySQLStr & "'" & Declarations.MyEventID & "', "
        MySQLStr = MySQLStr & "'" & Label15.Text & "', "
        If TextBox2.Text = "" Then
            MySQLStr = MySQLStr & "N'" & Replace(Label9.Text, "'", "''") & "', "
        Else
            MySQLStr = MySQLStr & "NULL, "
        End If
        If TextBox1.Text = "" Then
            MySQLStr = MySQLStr & "N'" & Replace(Label10.Text, "'", "''") & "', "
        Else
            MySQLStr = MySQLStr & "NULL, "
        End If
        If TextBox1.Text = "" Then
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & "N'" & Replace(TextBox1.Text, "'", "''") & "', "
        End If
        If TextBox2.Text = "" Then
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & "N'" & Replace(TextBox2.Text, "'", "''") & "', "
        End If
        If TextBox3.Text = "" Then
            MySQLStr = MySQLStr & "N'" & Replace(Label11.Text, "'", "''") & "', "
            MySQLStr = MySQLStr & "NULL, "
        Else
            MySQLStr = MySQLStr & "NULL, "
            MySQLStr = MySQLStr & "N'" & Replace(TextBox3.Text, "'", "''") & "', "
        End If
        MySQLStr = MySQLStr & Replace(TextBox4.Text, ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(TextBox5.Text, ",", ".") & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MyWHAbsences.ListPreparation()
        MyWHAbsences.ChangeButtonsStatus()
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Доступность / недоступность поля "Доступное количество"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckAvlField()
    End Sub

    Private Sub TextBox2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Доступность / недоступность поля "Доступное количество"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckAvlField()
    End Sub

    Private Sub TextBox3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Доступность / недоступность поля "Доступное количество"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckAvlField()
    End Sub

    Private Sub CheckAvlField()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Доступность / недоступность поля "Доступное количество"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" Then '---запас из Scala
            TextBox5.Enabled = False
            TextBox5.Text = Trim(MyWHAbsences.DataGridView1.SelectedRows.Item(0).Cells(5).Value)
        Else
            TextBox5.Enabled = True
        End If

    End Sub
End Class