Module Main

    Sub Main()
        '////////////////////////////////////////////////////////////////////////////////
        '//
        '// она и в Африке main
        '//
        '////////////////////////////////////////////////////////////////////////////////

        If EventLog.SourceExists("CRM_ESK_DailyJob") = False Then '--первый раз запустить от имени администратора для создания лога
            EventLog.CreateEventSource("CRM_ESK_DailyJob", "Application")
        End If

        If My.Settings.MyDebug = "YES" Then
            EventLog.WriteEntry("CRM_ESK_DailyJob", "Старт процедуры проверки проектов в CRM")
        End If
        CheckProjects()
        If My.Settings.MyDebug = "YES" Then
            EventLog.WriteEntry("CRM_ESK_DailyJob", "Окончание процедуры проверки проектов в CRM")
        End If
    End Sub

End Module
