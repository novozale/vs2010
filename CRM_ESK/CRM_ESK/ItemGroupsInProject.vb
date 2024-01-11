Public Class ItemGroupsInProject

    Private Sub ItemGroupsInProject_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка данных в форму
        '// 
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String                        'рабочая строка
        Dim MyAdapter As SqlClient.SqlDataAdapter     'для списка активностей
        Dim MyDs As New DataSet

        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect

        MySQLStr = "SELECT tbl_CRM_ProdGroupsList.ID, CONVERT(bit, ISNULL(View_7.Selected, 0)) AS IsSelected, tbl_CRM_ProdGroupsList.ItemGroupName, "
        MySQLStr = MySQLStr & "ISNULL(View_7.ProdGroupComment, '') AS ProdGroupComment "
        MySQLStr = MySQLStr & "FROM tbl_CRM_ProdGroupsList LEFT OUTER JOIN "
        MySQLStr = MySQLStr & "(SELECT ProdGroupID, ProdGroupComment, - 1 AS Selected "
        MySQLStr = MySQLStr & "FROM tbl_CRM_Project_ProdGroups "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "')) AS View_7 ON tbl_CRM_ProdGroupsList.ID = View_7.ProdGroupID "
        MySQLStr = MySQLStr & "Order By ID "

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
        DataGridView1.Columns(1).HeaderText = "В про ек те"
        DataGridView1.Columns(1).Width = 40
        DataGridView1.Columns(2).HeaderText = "Группа продуктов"
        DataGridView1.Columns(2).Width = 300
        DataGridView1.Columns(2).ReadOnly = True
        DataGridView1.Columns(3).HeaderText = "Детали (комментарий)"
        DataGridView1.Columns(3).Width = 450
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход без сохранения
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход с сохранением информации
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If CheckData() = True Then
            SaveData()
            Me.Close()
        End If
    End Sub

    Private Function CheckData() As Boolean
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// проверка корректности введенных данных
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim SelectedQTY As Integer

        SelectedQTY = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(1).Value = True Then
                SelectedQTY = SelectedQTY + 1
            End If
        Next
        If SelectedQTY = 0 Then
            MsgBox("Для проекта необходимо выбрать хотя бы одну группу товаров, которая будет поставляться в рамках проекта", MsgBoxStyle.Critical, "Внимание!")
            CheckData = False
            Exit Function
        End If

        CheckData = True
    End Function

    Private Sub Savedata()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// сохранение введенных данных
        '//
        '////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        '---очистка старого
        MySQLStr = "DELETE FROM tbl_CRM_Project_ProdGroups "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & Declarations.MyProjectID & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)

        '---занесение нового
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If DataGridView1.Rows(i).Cells(1).Value = True Then
                MySQLStr = "INSERT INTO tbl_CRM_Project_ProdGroups "
                MySQLStr = MySQLStr & "(ID, ProjectID, ProdGroupID, ProdGroupComment) "
                MySQLStr = MySQLStr & "VALUES (NEWID(), "
                MySQLStr = MySQLStr & "'" & Declarations.MyProjectID & "', "
                MySQLStr = MySQLStr & DataGridView1.Rows(i).Cells(0).Value.ToString & ", "
                MySQLStr = MySQLStr & "N'" & DataGridView1.Rows(i).Cells(3).Value.ToString & "') "
                InitMyConn(False)
                Declarations.MyConn.Execute(MySQLStr)
            End If
        Next
    End Sub
End Class