Imports System.IO
Imports System.Threading
Imports System.Configuration
Imports StockApplication
Imports TechnicalIndicatorsCalculationServiceProj
Imports System.ServiceModel
Imports System.ServiceModel.Web
Imports System.ServiceModel.Description

Public Class StockAppHourlyService


    Protected Overrides Sub OnStart(ByVal args() As String)
        Me.WriteToFile("StockAppDataDownload Service started at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Me.ScheduleService()
    End Sub

    Protected Overrides Sub OnStop()
        Me.WriteToFile("StockAppDataDownload Service stopped at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Me.Schedular.Dispose()
    End Sub

    Private Schedular As Timer

    Public Sub ScheduleService()
        Try
            Schedular = New Timer(New TimerCallback(AddressOf SchedularCallback))

            'Set the Default Time.
            Dim scheduledTime As DateTime = DateTime.MinValue
            Dim dailyTimeStart As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 9, 5, 0, 0) '#9:00:00 AM#
            Dim dailyTimeEnd As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 16, 5, 0, 0) '#4:00:00 PM#
            Dim weekendStartTimeToGetNSEData As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 8, 50, 0, 0) '#8:50:00 AM#
            Dim weekendEndTimeToGetNSEData As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 9, 10, 0, 0) ' #9:10:00 AM#
            Dim weekdayTimeToGetDailyStockDataStart As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 17, 0, 0, 0)
            Dim weekdayTimeToGetDailyStockDataEnd As DateTime = New DateTime(Now.Year, Now.Month, Now.Day, 17, 5, 0, 0)
            Dim timeSpanForWeekDaywithinLimit As TimeSpan

            Dim intervalMinutes As Integer = Convert.ToInt32(ConfigurationManager.AppSettings("IntervalMinutes"))
            'Set the Scheduled Time by adding the Interval to Current Time.
            scheduledTime = DateTime.Now.AddMinutes(intervalMinutes)
            Try
                If Weekday(Today) > 1 And Weekday(Today) < 7 Then
                    'Stock Data collection will only happen on Weekdays betwen 9 to 4PM
                    If DateTime.Now.TimeOfDay >= dailyTimeStart.TimeOfDay And DateTime.Now.TimeOfDay < dailyTimeEnd.TimeOfDay Then
                        'Indices details fetch and store
                        Me.WriteToFile("NSEDetails entry started " & DateTime.Now.TimeOfDay.ToString)
                        Dim tmpNSEIndicesDetails As NSEIndicesDetails = New NSEIndicesDetails()
                        tmpNSEIndicesDetails.getIndicesDetailsAndStore()
                        Me.WriteToFile("NSEDetails entry End " & DateTime.Now.TimeOfDay.ToString)
                        'Hourly Data fetch and entry
                        Me.WriteToFile("hourlyStockdata entry started" & DateTime.Now.TimeOfDay.ToString)
                        Dim tmpHourlyStockQuote As HourlyStockQuote = New HourlyStockQuote()
                        tmpHourlyStockQuote.GetAndStoreHourlyData()
                        Me.WriteToFile("hourlyStockdata entry End " & DateTime.Now.TimeOfDay.ToString)
                        'IntraDay techinal indicator service call starts
                        Using cf As New ChannelFactory(Of IConnectorService)(New WebHttpBinding(), "http://localhost:6060")
                            cf.Endpoint.Behaviors.Add(New WebHttpBehavior())
                            Dim channel As IConnectorService = cf.CreateChannel()
                            Me.WriteToFile("StockAppDataDownload Service calling connector Service Up " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
                            channel.ProcessIntraDayTechnicalIndicators()
                            Me.WriteToFile("StockAppDataDownload Service called connector Service Up  " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
                        End Using
                    End If
                    'Daily stock collection data for daily table will happen at 5PM every day 
                    If DateTime.Now.TimeOfDay >= weekdayTimeToGetDailyStockDataStart.TimeOfDay And DateTime.Now.TimeOfDay < weekdayTimeToGetDailyStockDataEnd.TimeOfDay Then
                        'Daily stock entry 
                        Me.WriteToFile("DailyStockDetails entry started " & DateTime.Now.TimeOfDay.ToString)
                        Dim tmpDailyStockQuote As DailyStockQuote = New DailyStockQuote()
                        tmpDailyStockQuote.getDailyStockDetailsAndStore()
                        Me.WriteToFile("DailyStockDetails entry End " & DateTime.Now.TimeOfDay.ToString)
                        'Daily techinal indicator service call starts
                        Using cf As New ChannelFactory(Of IConnectorService)(New WebHttpBinding(), "http://localhost:6060")
                            cf.Endpoint.Behaviors.Add(New WebHttpBehavior())
                            Dim channel As IConnectorService = cf.CreateChannel()
                            Me.WriteToFile("StockAppDataDownload Service calling connector Service Up " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
                            channel.ProcessDailyTechnicalIndicators()
                            Me.WriteToFile("StockAppDataDownload Service called connector Service Up  " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
                        End Using
                    End If
                ElseIf Weekday(Today) = 1 And (DateTime.Now.TimeOfDay > weekendStartTimeToGetNSEData.TimeOfDay And DateTime.Now.TimeOfDay < weekendEndTimeToGetNSEData.TimeOfDay) Then
                    'NSE List will get updated every Sunday
                    Me.WriteToFile("NSEList entry started " & DateTime.Now.TimeOfDay.ToString)
                    Dim tmpNSEindices As NSEindices = New NSEindices()
                        tmpNSEindices.getIndicesListAndStore()
                    Me.WriteToFile("NSEList entry End " & DateTime.Now.TimeOfDay.ToString)
                End If
            Catch ex As Exception
                WriteToFile("StockAppDataDownload Service Error in getting stock data " + ex.Message + ex.StackTrace)
            End Try
            'Get the difference in Minutes between the Scheduled and Current Time.
            Dim timeSpan As TimeSpan
            Dim dueTime As Integer
            If DateTime.Now > scheduledTime Then
                'If Scheduled Time is passed set Schedule for the next Interval.
                scheduledTime = scheduledTime.AddMinutes(intervalMinutes)
            End If
            timeSpanForWeekDaywithinLimit = scheduledTime.Subtract(DateTime.Now)
            If Weekday(Today) > 1 And Weekday(Today) < 7 And DateTime.Now.TimeOfDay < dailyTimeStart.TimeOfDay Then
                'weekday and time is earlier than 9AM
                Me.WriteToFile("daily time = " & dailyTimeStart)
                timeSpan = dailyTimeStart.Subtract(DateTime.Now)
                Me.WriteToFile("withintime of weekday condition Next scheduled time set as = " & timeSpan.ToString("%d") & "days and " & timeSpan.ToString("hh\:mm\:ss"))
            ElseIf DateTime.Now.TimeOfDay > weekdayTimeToGetDailyStockDataStart.TimeOfDay Or Weekday(Today) = 1 Or Weekday(Today) = 7 Then
                'setting next execution time to next day morning 9 AM in case current time is more than 5PM or current day is weekend
                Dim now As DateTime = DateTime.Now
                Dim myDate = New DateTime(now.Year, now.Month, now.Day + 1, 9, 0, 0, 0)
                timeSpan = myDate.Subtract(DateTime.Now)
                Me.WriteToFile(" outsidetime condition - Next scheduled time set as = " & timeSpan.ToString("%d") & "days and " & timeSpan.ToString("hh\:mm\:ss"))
            ElseIf DateTime.Now.TimeOfDay > dailyTimeEnd.TimeOfDay And DateTime.Now.TimeOfDay < weekdayTimeToGetDailyStockDataStart.TimeOfDay Then
                'setting next execution time to 5 PM in case current time is more than 4PM and less than 5PM
                timeSpan = weekdayTimeToGetDailyStockDataStart.Subtract(DateTime.Now)
                Me.WriteToFile(" outside reguar time but less than daily collection time - Next scheduled time set as = " & timeSpan.ToString("%d") & "days and " & timeSpan.ToString("hh\:mm\:ss"))
            ElseIf Weekday(Today) > 1 And Weekday(Today) < 7 Then
                'Weekday within data collection limit of 9 to 4PM
                timeSpan = timeSpanForWeekDaywithinLimit
                Me.WriteToFile("weekday condition Next scheduled time set as = " & timeSpan.ToString("%d") & "days and " & timeSpan.ToString("hh\:mm\:ss"))
            End If
            Try
                dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds)
            Catch ex As Exception
                WriteToFile("StockAppDataDownload Service Error on calculating due time : {0} " + ex.Message + ex.StackTrace)
                dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds + 5000)
            End Try
            'Change the Timer's Due Time.
            Schedular.Change(dueTime, Timeout.Infinite)
            Catch ex As Exception
                WriteToFile("StockAppDataDownload Service Error on: {0} " + ex.Message + ex.StackTrace)
            'Stop the Windows Service.
            Using serviceController As New System.ServiceProcess.ServiceController("StockAppDataDownload")
                serviceController.[Stop]()
            End Using
        End Try
    End Sub

    Private Sub SchedularCallback(e As Object)
        Me.WriteToFile("StockAppDataDownload Service callback Log: " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Me.ScheduleService()
    End Sub

    Private Sub WriteToFile(text As String)
        Dim path As String = "D:\Tarun\StockApp\Log\StockAppDataDownloadServiceLog.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub
End Class