Imports System.Reflection
Imports System.CodeDom.Compiler


Module DynamicClasses
    Public Function CreateMonthScheduleClassType(ByVal className As String, ByVal properties As Dictionary(Of String, Type)) As Type
        '/////////////////////////////////////////////////////////////////////////////////////
        '//
        '// Динамическое создание класса для работы с планами
        '//
        '/////////////////////////////////////////////////////////////////////////////////////
        Dim myCode As CodeDomProvider = CodeDomProvider.CreateProvider("VB")
        Dim myPar As New CompilerParameters()
        Dim myCodeBody As New System.Text.StringBuilder()
        Dim i As Integer

        myCodeBody.AppendLine("Public Class " + className + " ")
        '-----properties
        For Each o In properties
            myCodeBody.AppendLine("Private " + o.Key + "_value As " + o.Value.ToString + " ")

            myCodeBody.AppendLine("Public Property " + o.Key + "() As " + o.Value.ToString + " ")
            myCodeBody.AppendLine("Get ")
            myCodeBody.AppendLine("Return " + o.Key + "_value ")
            myCodeBody.AppendLine("End Get ")
            myCodeBody.AppendLine("Set(ByVal value As " + o.Value.ToString + ") ")
            myCodeBody.AppendLine(o.Key + "_value = value ")
            myCodeBody.AppendLine("End Set ")
            myCodeBody.AppendLine("End Property ")
        Next

        '-----Функция GetItem(String)
        myCodeBody.AppendLine("Public Function GetItem(ByVal ItemName As String) As Object ")
        myCodeBody.AppendLine("Select Case ItemName ")
        For Each o In properties
            myCodeBody.AppendLine("Case """ + o.Key + """ ")
            myCodeBody.AppendLine("Return " + o.Key + "_value ")
        Next
        myCodeBody.AppendLine("Case Else ")
        myCodeBody.AppendLine("Return Nothing ")
        myCodeBody.AppendLine("End Select ")
        myCodeBody.AppendLine("End Function ")

        '-----Функция SetItem(String)
        myCodeBody.AppendLine("Public Function SetItem(ByVal ItemName As String, ByVal ItemValue as object) ")
        myCodeBody.AppendLine("Select Case ItemName ")
        For Each o In properties
            myCodeBody.AppendLine("Case """ + o.Key + """ ")
            myCodeBody.AppendLine(o.Key + "_value = ItemValue ")
        Next
        myCodeBody.AppendLine("Case Else ")
        myCodeBody.AppendLine("End Select ")
        myCodeBody.AppendLine("End Function ")

        '-----Функция GetItem(Integer)
        myCodeBody.AppendLine("Public Function GetItem(ByVal ItemID As Integer) As Object ")
        myCodeBody.AppendLine("Select Case ItemID ")
        i = 0
        For Each o In properties
            myCodeBody.AppendLine("Case " + i.ToString() + " ")
            myCodeBody.AppendLine("Return " + o.Key + "_value ")
            i = i + 1
        Next
        myCodeBody.AppendLine("Case Else ")
        myCodeBody.AppendLine("Return Nothing ")
        myCodeBody.AppendLine("End Select ")
        myCodeBody.AppendLine("End Function ")

        '-----Функция SetItem(Integer)
        myCodeBody.AppendLine("Public Function SetItem(ByVal ItemID As Integer, ByVal ItemValue as object) ")
        myCodeBody.AppendLine("Select Case ItemID ")
        i = 0
        For Each o In properties
            myCodeBody.AppendLine("Case " + i.ToString() + " ")
            myCodeBody.AppendLine(o.Key + "_value = ItemValue ")
            i = i + 1
        Next
        myCodeBody.AppendLine("Case Else ")
        myCodeBody.AppendLine("End Select ")
        myCodeBody.AppendLine("End Function ")

        myCodeBody.AppendLine("End Class ")

        Dim myResult As CompilerResults = myCode.CompileAssemblyFromSource(myPar, myCodeBody.ToString())
        If myResult.Errors.HasErrors Then
            CreateMonthScheduleClassType = Nothing
        End If

        Dim myAsm As Assembly = myResult.CompiledAssembly()
        Dim myType As Type = myAsm.GetType(className)
        'Dim myCls As Object = myAsm.CreateInstance(className, True)

        'CreateMonthScheduleClass = myCls
        CreateMonthScheduleClassType = myType
    End Function
End Module
