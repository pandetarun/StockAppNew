Imports System.IO
Imports System.Threading
Imports System.Configuration
Imports StockApplication

Public Class StockAppHourlyService


    Protected Overrides Sub OnStart(ByVal args() As String)
        Me.WriteToFile("StockApp Service started at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Me.ScheduleService()
    End Sub

    Protected Overrides Sub OnStop()
        Me.WriteToFile("StockApp Service stopped at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Me.Schedular.Dispose()
    End Sub

    Private Schedular As Timer

    Public Sub ScheduleService()
        Try
            Schedular = New Timer(New TimerCallback(AddressOf SchedularCallback))

            'Set the Default Time.
            Dim scheduledTime As DateTime = DateTime.MinValue
            Dim dailyTimeStart As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 9, 0, 0, 0) '#9:00:00 AM#
            Dim dailyTimeEnd As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 16, 0, 0, 0) '#4:00:00 PM#
            Dim weekendStartTimeToGetNSEData As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 8, 50, 0, 0) '#8:50:00 AM#
            Dim weekendEndTimeToGetNSEData As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 9, 10, 0, 0) ' #9:10:00 AM#

            Dim intervalMinutes As Integer = Convert.ToInt32(ConfigurationManager.AppSettings("IntervalMinutes"))

            'Set the Scheduled Time by adding the Interval to Current Time.
            scheduledTime = DateTime.Now.AddMinutes(intervalMinutes)
            If DateTime.Now > scheduledTime Then
                'If Scheduled Time is passed set Schedule for the next Interval.
                scheduledTime = scheduledTime.AddMinutes(intervalMinutes)
            End If
            'End If
            Try
                If Weekday(Today) > 1 And Weekday(Today) < 7 Then
                    'Stock Data collection will only happen on Weekdays
                    If DateTime.Now.TimeOfDay > dailyTimeStart.TimeOfDay And DateTime.Now.TimeOfDay < dailyTimeEnd.TimeOfDay Then
                        'Indices details fetch and store
                        Me.WriteToFile("NSEDetails entry started" & DateTime.Now.TimeOfDay.ToString)
                        Dim tmpNSEIndicesDetails As NSEIndicesDetails
                        tmpNSEIndicesDetails = New NSEIndicesDetails()
                        tmpNSEIndicesDetails.getIndicesDetailsAndStore()
                        Me.WriteToFile("NSEDetails entry End" & DateTime.Now.TimeOfDay.ToString)
                        'Hourly Data fetch and entry
                        Me.WriteToFile("hourlyStockdata entry started" & DateTime.Now.TimeOfDay.ToString)
                        Dim tmpHourlyStockQuote As HourlyStockQuote
                        tmpHourlyStockQuote = New HourlyStockQuote()
                        tmpHourlyStockQuote.GetAndStoreHourlyData()
                        Me.WriteToFile("hourlyStockdata entry End" & DateTime.Now.TimeOfDay.ToString)
                    End If
                ElseIf Weekday(Today) = 1 And (DateTime.Now.TimeOfDay > weekendStartTimeToGetNSEData.TimeOfDay And DateTime.Now.TimeOfDay < weekendEndTimeToGetNSEData.TimeOfDay) Then
                    'NSE List will get updated every Sunday
                    Me.WriteToFile("NSEList entry started" & DateTime.Now.TimeOfDay.ToString)
                    Dim tmpNSEindices As NSEindices
                    tmpNSEindices = New NSEindices()
                    tmpNSEindices.getIndicesListAndStore()
                    Me.WriteToFile("NSEList entry End" & DateTime.Now.TimeOfDay.ToString)
                End If
            Catch ex As Exception
                WriteToFile("StockApp Service Error in getting stock data" + ex.Message + ex.StackTrace)
            End Try
            'Get the difference in Minutes between the Scheduled and Current Time.
            Dim timeSpan As TimeSpan
            Dim dueTime As Integer
            If Weekday(Today) > 1 And Weekday(Today) < 7 And DateTime.Now.TimeOfDay < dailyTimeStart.TimeOfDay Then
                Me.WriteToFile("daily time = " & dailyTimeStart)
                timeSpan = dailyTimeStart.Subtract(DateTime.Now)
                Me.WriteToFile("withintime of weekday condition Next scheduled time set as = " & timeSpan.ToString("%d") & "days and " & timeSpan.ToString("hh\:mm\:ss"))
            ElseIf DateTime.Now.TimeOfDay > dailyTimeEnd.TimeOfDay Or Weekday(Today) = 1 Or Weekday(Today) = 7 Then
                Dim now As DateTime = DateTime.Now
                Dim myDate = New DateTime(now.Year, now.Month, now.Day + 1, 9, 0, 0, 0)
                timeSpan = myDate.Subtract(DateTime.Now)
                Me.WriteToFile(" outsidetime condition - Next scheduled time set as = " & timeSpan.ToString("%d") & "days and " & timeSpan.ToString("hh\:mm\:ss"))
            ElseIf Weekday(Today) > 1 And Weekday(Today) < 7 Then
                timeSpan = scheduledTime.Subtract(DateTime.Now)
                Me.WriteToFile("weekday condition Next scheduled time set as = " & timeSpan.ToString("%d") & "days and " & timeSpan.ToString("hh\:mm\:ss"))
            End If
            dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds)
            'Change the Timer's Due Time.
            Schedular.Change(dueTime, Timeout.Infinite)
        Catch ex As Exception
            WriteToFile("StockApp Service Error on: {0} " + ex.Message + ex.StackTrace)
            'Stop the Windows Service.
            Using serviceController As New System.ServiceProcess.ServiceController("StockService")
                serviceController.[Stop]()
            End Using
        End Try
    End Sub

    Private Sub SchedularCallback(e As Object)
        Me.WriteToFile("StockApp Service callback Log: " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Me.ScheduleService()
    End Sub

    Private Sub WriteToFile(text As String)
        Dim path As String = "C:\ServiceLog.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub
End Class