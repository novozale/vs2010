Public Class ProjectDetailsImport

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// Выход 
        '//
        '////////////////////////////////////////////////////////////////////////////////

        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Процедура загрузки из Excel детальной информации по проекту  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Label3.Text = ""
        Me.Refresh()
        System.Windows.Forms.Application.DoEvents()
        Button1.Enabled = False
        Button2.Enabled = False
        If My.Settings.UseOffice = "LibreOffice" Then
            ImportDataFromLO()
        Else
            ImportDataFromExcel()
        End If
        Button1.Enabled = True
        Button2.Enabled = True
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub ImportDataFromExcel()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из Excel  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyVersion As String                     'Версия документа
        Dim appXLSRC As Object
        Dim MySQLStr As String
        Dim MyCurrency As Integer                   'валюта детальной информации
        Dim i As Double                             'счетчик строк
        Dim SupplierItemCode As String              'код товара поставщика
        Dim ScalaItemCode As String                 'код товара в Scala
        Dim MyPriCost As Double                     'себестоимость товара
        Dim MyPrice As Double                       'цена товара
        Dim MyQTY As Double                         'количество товара

        If OpenFileDialog1.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog1.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()

                appXLSRC = CreateObject("Excel.Application")
                appXLSRC.Workbooks.Open(OpenFileDialog1.FileName)

                '---Удаляем старую информацию
                DeleteAddInfo(Declarations.MyProjectID)

                '---Проверяем версию Excel файла с деталььными данными
                MyVersion = Trim(appXLSRC.Worksheets(1).Range("A1").Value)
                MySQLStr = "SELECT Version "
                MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH(NOLOCK) "
                MySQLStr = MySQLStr & "WHERE (Name = N'Импорт детальной информации по проекту') "
                InitMyConn(False)
                InitMyRec(False, MySQLStr)
                If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                    MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    trycloseMyRec()
                    DeleteAddInfo(Declarations.MyProjectID)
                    Exit Sub
                Else
                    If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                        trycloseMyRec()
                    Else
                        MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & Trim(Declarations.MyRec.Fields("Version").Value) & ".", vbCritical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        trycloseMyRec()
                        DeleteAddInfo(Declarations.MyProjectID)
                        Exit Sub
                    End If
                End If

                '---Проверяем, что проставлен корректный код валюты
                Try
                    MyCurrency = appXLSRC.Worksheets(1).Range("E9").Value
                Catch ex As Exception
                    MsgBox("Ошибка проставления валюты в Excel файле ячейка E9: " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    DeleteAddInfo(Declarations.MyProjectID)
                    Exit Sub
                End Try

                If MyCurrency <> 0 And MyCurrency <> 1 And MyCurrency <> 12 Then
                    MsgBox("В Excel файле ячейка E9 должна быть проставлена валюта: 0 - рубли или 1 - доллары или 12 - евро.", MsgBoxStyle.Critical, "Внимание!")
                    appXLSRC.DisplayAlerts = 0
                    appXLSRC.Workbooks.Close()
                    appXLSRC.DisplayAlerts = 1
                    appXLSRC.Quit()
                    appXLSRC = Nothing
                    DeleteAddInfo(Declarations.MyProjectID)
                    Exit Sub
                End If

                '---импорт детальной информации 
                i = 27
                While appXLSRC.Worksheets(1).Range("B" & i).Value <> Nothing Or appXLSRC.Worksheets(1).Range("C" & i).Value <> Nothing
                    '---коды товара
                    If appXLSRC.Worksheets(1).Range("B" & i).Value = Nothing Then
                        SupplierItemCode = ""
                    Else
                        SupplierItemCode = Trim(appXLSRC.Worksheets(1).Range("B" & i).Value.ToString)
                    End If

                    If appXLSRC.Worksheets(1).Range("C" & i).Value = Nothing Then
                        ScalaItemCode = ""
                    Else
                        ScalaItemCode = Trim(appXLSRC.Worksheets(1).Range("C" & i).Value.ToString)
                    End If

                    If SupplierItemCode = "" And ScalaItemCode = "" Then
                        MsgBox("В Excel файле строка " & CStr(i) & " Должны быть заполнены или код товара поставщика, или код скала непустыми значениями.", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        DeleteAddInfo(Declarations.MyProjectID)
                        Exit Sub
                    End If

                    If ScalaItemCode <> "" Then
                        MySQLStr = "SELECT COUNT(*) AS CC "
                        MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaItemCode & "') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If (Declarations.MyRec.Fields("CC").Value = 0) Then
                            trycloseMyRec()
                            MsgBox("В Excel файле в ячейку C" & CStr(i) & " внесено значение кода товара в Scala: " & ScalaItemCode & " , которое отсутствует в базе данных. Проверьте код и ведите корректный.", MsgBoxStyle.Critical, "Внимание!")
                            appXLSRC.DisplayAlerts = 0
                            appXLSRC.Workbooks.Close()
                            appXLSRC.DisplayAlerts = 1
                            appXLSRC.Quit()
                            appXLSRC = Nothing
                            DeleteAddInfo(Declarations.MyProjectID)
                            Exit Sub
                        Else
                            trycloseMyRec()
                        End If
                    End If

                    '---Количество товара
                    Try
                        MyQTY = appXLSRC.Worksheets(1).Range("E" & i).Value
                    Catch ex As Exception
                        MsgBox("В Excel файле в ячейку E" & CStr(i) & " (Количество) должно быть внесено число.", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        DeleteAddInfo(Declarations.MyProjectID)
                        Exit Sub
                    End Try

                    If MyQTY <= 0 Then
                        MsgBox("В Excel файле в ячейку E" & CStr(i) & " (Количество) должно быть внесено число больше 0. Количество  не может быть 0 и меньше 0.", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        DeleteAddInfo(Declarations.MyProjectID)
                        Exit Sub
                    End If

                    '---Себестоимость товара
                    Try
                        MyPriCost = appXLSRC.Worksheets(1).Range("I" & i).Value
                    Catch ex As Exception
                        MsgBox("В Excel файле в ячейку I" & CStr(i) & " (Себестоимость) должно быть внесено число.", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        DeleteAddInfo(Declarations.MyProjectID)
                        Exit Sub
                    End Try

                    If MyPriCost <= 0 Then
                        MsgBox("В Excel файле в ячейку I" & CStr(i) & " (Себестоимость) должно быть внесено число больше 0. Себестоимость не может быть 0 и меньше 0.", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        DeleteAddInfo(Declarations.MyProjectID)
                        Exit Sub
                    End If

                    '---Цена товара
                    Try
                        MyPrice = appXLSRC.Worksheets(1).Range("J" & i).Value
                    Catch ex As Exception
                        MsgBox("В Excel файле в ячейку J" & CStr(i) & " (Цена без НДС) должно быть внесено число.", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        DeleteAddInfo(Declarations.MyProjectID)
                        Exit Sub
                    End Try

                    If MyPrice <= 0 Then
                        MsgBox("В Excel файле в ячейку J" & CStr(i) & " (Цена без НДС) должно быть внесено число больше 0. Цена не может быть 0.", MsgBoxStyle.Critical, "Внимание!")
                        appXLSRC.DisplayAlerts = 0
                        appXLSRC.Workbooks.Close()
                        appXLSRC.DisplayAlerts = 1
                        appXLSRC.Quit()
                        appXLSRC = Nothing
                        DeleteAddInfo(Declarations.MyProjectID)
                        Exit Sub
                    End If

                    '---Занесение информации в БД
                    MySQLStr = "INSERT INTO tbl_CRM_Project_Details "
                    MySQLStr = MySQLStr & "(ProjectID, SupplierItemCode, ScalaItemCode, QTY, ProjectPriCost, ProjectPrice, CurrCode) "
                    MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyProjectID & "', "
                    If SupplierItemCode = "" Then
                        MySQLStr = MySQLStr & "NULL, "
                    Else
                        MySQLStr = MySQLStr & "N'" & SupplierItemCode & "', "
                    End If
                    If ScalaItemCode = "" Then
                        MySQLStr = MySQLStr & "NULL, "
                    Else
                        MySQLStr = MySQLStr & "N'" & ScalaItemCode & "', "
                    End If
                    MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ", "
                    MySQLStr = MySQLStr & Replace(CStr(MyPriCost), ",", ".") & ", "
                    MySQLStr = MySQLStr & Replace(CStr(MyPrice), ",", ".") & ", "
                    MySQLStr = MySQLStr & CStr(MyCurrency) & ") "
                    InitMyConn(False)
                    Declarations.MyConn.Execute(MySQLStr)


                    Label3.Text = CStr(i - 26)
                    Me.Refresh()
                    System.Windows.Forms.Application.DoEvents()
                    i = i + 1
                End While


                appXLSRC.DisplayAlerts = 0
                appXLSRC.Workbooks.Close()
                appXLSRC.DisplayAlerts = 1
                appXLSRC.Quit()
                appXLSRC = Nothing
                MsgBox("Импорт детальной информации по проекту произведен", MsgBoxStyle.OkOnly, "Внимание!")
            End If
        End If
    End Sub

    Private Sub ImportDataFromLO()
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Загрузка из Libre Office  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MyVersion As String                     'Версия документа
        Dim MySQLStr As String
        Dim MyCurrency As Integer                   'валюта детальной информации
        Dim i As Double                             'счетчик строк
        Dim SupplierItemCode As String              'код товара поставщика
        Dim ScalaItemCode As String                 'код товара в Scala
        Dim MyPriCost As Double                     'себестоимость товара
        Dim MyPrice As Double                       'цена товара
        Dim MyQTY As Double                         'количество товара
        Dim oServiceManager As Object
        Dim oDispatcher As Object
        Dim oDesktop As Object
        Dim oWorkBook As Object
        Dim oSheet As Object
        Dim oFileName As String

        If OpenFileDialog2.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            If (OpenFileDialog2.FileName = "") Then
            Else
                Me.Cursor = Cursors.WaitCursor
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()
                Try
                    oServiceManager = CreateObject("com.sun.star.ServiceManager")
                    oDesktop = oServiceManager.createInstance("com.sun.star.frame.Desktop")
                    oDispatcher = oServiceManager.createInstance("com.sun.star.frame.DispatchHelper")
                    oFileName = Replace(OpenFileDialog2.FileName, "\", "/")
                    oFileName = "file:///" + oFileName
                    Dim arg(1)
                    arg(0) = mAkePropertyValue("Hidden", True, oServiceManager)
                    arg(1) = mAkePropertyValue("MacroExecutionMode", 4, oServiceManager)
                    oWorkBook = oDesktop.loadComponentFromURL(oFileName, "_blank", 0, arg)
                    oSheet = oWorkBook.getSheets().getByIndex(0)

                    '---Удаляем старую информацию
                    DeleteAddInfo(Declarations.MyProjectID)

                    '---Проверяем версию листа Excel
                    MyVersion = oSheet.getCellRangeByName("A1").String
                    If MyVersion = "" Then
                        MsgBox("В импортируемом листе Excel в ячейке 'A1' не проставлена версия листа Excel ", MsgBoxStyle.Critical, "Внимание!")
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    Else
                        MySQLStr = "SELECT Version "
                        MySQLStr = MySQLStr & "FROM tbl_VersionImportItemsFromExcel WITH(NOLOCK) "
                        MySQLStr = MySQLStr & "WHERE (Name = N'Импорт детальной информации по проекту') "
                        InitMyConn(False)
                        InitMyRec(False, MySQLStr)
                        If Declarations.MyRec.BOF = True And Declarations.MyRec.EOF = True Then
                            MsgBox("В Scala не проставлена текущая версия листа Excel. Обратитесь к администратору", vbCritical, "Внимание!")
                            trycloseMyRec()
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            Exit Sub
                        Else
                            If Trim(Declarations.MyRec.Fields("Version").Value) = MyVersion Then
                                trycloseMyRec()
                            Else
                                MsgBox("Вы пытаетесь работать с некорректной версией листа Excel. Надо работать с версией " & Trim(Declarations.MyRec.Fields("Version").Value) & ".", vbCritical, "Внимание!")
                                trycloseMyRec()
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                Exit Sub
                            End If
                        End If
                    End If

                    '---Проверяем, что проставлен корректный код валюты
                    Try
                        MyCurrency = oSheet.getCellRangeByName("E9").Value
                    Catch ex As Exception
                        MsgBox("Ошибка проставления валюты в Excel файле ячейка E9: " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                        trycloseMyRec()
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End Try

                    If MyCurrency <> 0 And MyCurrency <> 1 And MyCurrency <> 12 Then
                        MsgBox("В Excel файле ячейка E9 должна быть проставлена валюта: 0 - рубли или 1 - доллары или 12 - евро.", MsgBoxStyle.Critical, "Внимание!")
                        trycloseMyRec()
                        Me.Cursor = Cursors.Default
                        oWorkBook.Close(True)
                        Exit Sub
                    End If

                    '---импорт детальной информации
                    i = 27
                    While oSheet.getCellRangeByName("B" & i).String.Equals("") = False Or oSheet.getCellRangeByName("C" & i).String.Equals("") = False
                        '---коды товара
                        SupplierItemCode = oSheet.getCellRangeByName("B" & i).String
                        ScalaItemCode = oSheet.getCellRangeByName("C" & i).String

                        If SupplierItemCode = "" And ScalaItemCode = "" Then
                            MsgBox("В Excel файле строка " & CStr(i) & " Должны быть заполнены или код товара поставщика, или код скала непустыми значениями.", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            DeleteAddInfo(Declarations.MyProjectID)
                            Exit Sub
                        End If

                        If ScalaItemCode <> "" Then
                            MySQLStr = "SELECT COUNT(*) AS CC "
                            MySQLStr = MySQLStr & "FROM SC010300 WITH(NOLOCK) "
                            MySQLStr = MySQLStr & "WHERE (SC01001 = N'" & ScalaItemCode & "') "
                            InitMyConn(False)
                            InitMyRec(False, MySQLStr)
                            If (Declarations.MyRec.Fields("CC").Value = 0) Then
                                trycloseMyRec()
                                MsgBox("В Excel файле в ячейку C" & CStr(i) & " внесено значение кода товара в Scala: " & ScalaItemCode & " , которое отсутствует в базе данных. Проверьте код и ведите корректный.", MsgBoxStyle.Critical, "Внимание!")
                                Me.Cursor = Cursors.Default
                                oWorkBook.Close(True)
                                DeleteAddInfo(Declarations.MyProjectID)
                                Exit Sub
                            Else
                                trycloseMyRec()
                            End If
                        End If

                        '---Количество товара
                        Try
                            MyQTY = oSheet.getCellRangeByName("E" & i).Value
                        Catch ex As Exception
                            MsgBox("В Excel файле в ячейку E" & CStr(i) & " (Количество) должно быть внесено число.", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            DeleteAddInfo(Declarations.MyProjectID)
                            Exit Sub
                        End Try

                        If MyQTY <= 0 Then
                            MsgBox("В Excel файле в ячейку E" & CStr(i) & " (Количество) должно быть внесено число больше 0. Количество  не может быть 0 и меньше 0.", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            DeleteAddInfo(Declarations.MyProjectID)
                            Exit Sub
                        End If

                        '---Себестоимость товара
                        Try
                            MyPriCost = oSheet.getCellRangeByName("I" & i).Value
                        Catch ex As Exception
                            MsgBox("В Excel файле в ячейку I" & CStr(i) & " (Себестоимость) должно быть внесено число.", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            DeleteAddInfo(Declarations.MyProjectID)
                            Exit Sub
                        End Try

                        If MyPriCost <= 0 Then
                            MsgBox("В Excel файле в ячейку I" & CStr(i) & " (Себестоимость) должно быть внесено число больше 0. Себестоимость не может быть 0 и меньше 0.", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            DeleteAddInfo(Declarations.MyProjectID)
                            Exit Sub
                        End If

                        '---Цена товара
                        Try
                            MyPrice = oSheet.getCellRangeByName("J" & i).Value
                        Catch ex As Exception
                            MsgBox("В Excel файле в ячейку J" & CStr(i) & " (Цена без НДС) должно быть внесено число.", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            DeleteAddInfo(Declarations.MyProjectID)
                            Exit Sub
                        End Try

                        If MyPrice <= 0 Then
                            MsgBox("В Excel файле в ячейку J" & CStr(i) & " (Цена без НДС) должно быть внесено число больше 0. Цена не может быть 0.", MsgBoxStyle.Critical, "Внимание!")
                            Me.Cursor = Cursors.Default
                            oWorkBook.Close(True)
                            DeleteAddInfo(Declarations.MyProjectID)
                            Exit Sub
                        End If

                        '---Занесение информации в БД
                        MySQLStr = "INSERT INTO tbl_CRM_Project_Details "
                        MySQLStr = MySQLStr & "(ProjectID, SupplierItemCode, ScalaItemCode, QTY, ProjectPriCost, ProjectPrice, CurrCode) "
                        MySQLStr = MySQLStr & "VALUES ('" & Declarations.MyProjectID & "', "
                        If SupplierItemCode = "" Then
                            MySQLStr = MySQLStr & "NULL, "
                        Else
                            MySQLStr = MySQLStr & "N'" & SupplierItemCode & "', "
                        End If
                        If ScalaItemCode = "" Then
                            MySQLStr = MySQLStr & "NULL, "
                        Else
                            MySQLStr = MySQLStr & "N'" & ScalaItemCode & "', "
                        End If
                        MySQLStr = MySQLStr & Replace(CStr(MyQTY), ",", ".") & ", "
                        MySQLStr = MySQLStr & Replace(CStr(MyPriCost), ",", ".") & ", "
                        MySQLStr = MySQLStr & Replace(CStr(MyPrice), ",", ".") & ", "
                        MySQLStr = MySQLStr & CStr(MyCurrency) & ") "
                        InitMyConn(False)
                        Declarations.MyConn.Execute(MySQLStr)

                        Label3.Text = CStr(i - 26)
                        Me.Refresh()
                        System.Windows.Forms.Application.DoEvents()
                        i = i + 1
                    End While
                Catch ex As Exception
                    MsgBox("ошибка : " & ex.Message, MsgBoxStyle.Critical, "Внимание!")
                End Try
                Me.Cursor = Cursors.Default
                oWorkBook.Close(True)
                MsgBox("Импорт данных произведен.", MsgBoxStyle.OkOnly, "Внимание!")

                Me.Cursor = Cursors.Default
                Me.Refresh()
                System.Windows.Forms.Application.DoEvents()
            End If
        End If
    End Sub

    Private Sub DeleteAddInfo(ByVal MyProjectID As String)
        '////////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Удаление дополнительной информации по проекту  
        '//
        '////////////////////////////////////////////////////////////////////////////////////////
        Dim MySQLStr As String

        MySQLStr = "DELETE FROM tbl_CRM_Project_Details "
        MySQLStr = MySQLStr & "WHERE (ProjectID = '" & MyProjectID & "') "
        InitMyConn(False)
        Declarations.MyConn.Execute(MySQLStr)
    End Sub
End Class