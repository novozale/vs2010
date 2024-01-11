Public Class ItemHighPrice

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub ItemHighPrice_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// выход без сохранения по Escape
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        If e.KeyData = Keys.Escape Then
            Me.Close()
        End If
    End Sub

    Private Sub ItemHighPrice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в окно по открытию формы
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////

        Label9.Text = Trim(MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(0).Value)
        Label10.Text = Trim(MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(4).Value)
        Label11.Text = Trim(MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(7).Value)
        TextBox4.Text = Trim(IIf(MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(2).Value = 0, "", MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString()))
        TextBox6.Text = Trim(IIf(MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(3).Value = 0, "", MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString()))
        If MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(2).Value <> 0 Then
            TextBox4.Enabled = False
        Else
            TextBox4.Enabled = True
        End If
        If MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(3).Value <> 0 Then
            TextBox6.Enabled = False
        Else
            TextBox6.Enabled = True
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
                MsgBox("В поле ""Наш прайс РУБ"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox4.Text
                Catch ex As Exception
                    MsgBox("В поле ""Наш прайс РУБ"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
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
                MsgBox("В поле ""Прайс ожидаемый покупателем РУБ"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox5.Text
                Catch ex As Exception
                    MsgBox("В поле ""Прайс ожидаемый покупателем РУБ"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
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

        If Trim(TextBox6.Text) = "" Then
            MsgBox("Поле ""Наша себестоимость РУБ"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox6.Select()
            Exit Function
        End If

        If Trim(TextBox4.Text) = "" Then
            MsgBox("Поле ""Наш прайс РУБ"" должно быть заполнено", MsgBoxStyle.Critical, "Внимание")
            CheckDataFiling = False
            TextBox4.Select()
            Exit Function
        End If

        If Trim(TextBox5.Text) = "" Then
            MsgBox("Поле ""Прайс ожидаемый покупателем РУБ"" должно быть заполнено.", MsgBoxStyle.Critical, "Внимание")
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
        Declarations.MyHighPriceID = MyGUID.ToString
        MySQLStr = "INSERT INTO tbl_CRM_HighPrice "
        MySQLStr = MySQLStr & "(ID, EventID, ScalaCode, SupplierCode, OtherSupplierCode, OtherItem, ScalaSupplierCode, OtherSupplier, OurPrice, EstimatedPrice, OurPriCost) "
        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyHighPriceID & "', "
        MySQLStr = MySQLStr & "'" & Declarations.MyEventID & "', "
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
        MySQLStr = MySQLStr & Replace(TextBox5.Text, ",", ".") & ", "
        MySQLStr = MySQLStr & Replace(TextBox6.Text, ",", ".") & ") "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
        MyHighPrice.ListPreparation()
        MyHighPrice.ChangeButtonsStatus()
    End Sub

    Private Sub TextBox1_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Доступность / недоступность поля "Наш прайс РУБ"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckAvlField()
    End Sub

    Private Sub TextBox2_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox2.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Доступность / недоступность поля "Наш прайс РУБ"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckAvlField()
    End Sub

    Private Sub TextBox3_Validated(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox3.Validated
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Доступность / недоступность поля "Наш прайс РУБ"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        CheckAvlField()
    End Sub

    Private Sub CheckAvlField()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Доступность / недоступность поля "Наш прайс РУБ"
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(2).Value <> 0 Then '---запас из Scala
            TextBox4.Enabled = False
            TextBox4.Text = Trim(MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(2).Value.ToString())
        Else
            TextBox4.Enabled = True
        End If

        If TextBox1.Text = "" And TextBox2.Text = "" And TextBox3.Text = "" And MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(3).Value <> 0 Then '---запас из Scala
            TextBox6.Enabled = False
            TextBox6.Text = Trim(MyHighPrice.DataGridView1.SelectedRows.Item(0).Cells(3).Value.ToString())
        Else
            TextBox6.Enabled = True
        End If

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

    Private Sub TextBox6_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles TextBox6.Validating
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка - числовое ли значение введено
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MyRez As Double
        Dim aa As New System.Globalization.NumberFormatInfo

        If Trim(TextBox6.Text) <> "" Then
            If InStr(TextBox6.Text, aa.CurrentInfo.NumberGroupSeparator) Then
                MsgBox("В поле ""Наша себестоимость РУБ"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                e.Cancel = True
                Exit Sub
            Else
                Try
                    MyRez = TextBox6.Text
                Catch ex As Exception
                    MsgBox("В поле ""Наша себестоимость РУБ"" должно быть введено число", MsgBoxStyle.Critical, "Внимание!")
                    e.Cancel = True
                    Exit Sub
                End Try
            End If
        End If
        e.Cancel = False
    End Sub
End Class