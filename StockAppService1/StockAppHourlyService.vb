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
            Dim TimeStart As DateTime = #5:00:00 AM#
            Dim TimeEnd As DateTime = #7:30:00 PM#

            Dim intervalMinutes As Integer = Convert.ToInt32(ConfigurationManager.AppSettings("IntervalMinutes"))

            'Set the Scheduled Time by adding the Interval to Current Time.
            scheduledTime = DateTime.Now.AddMinutes(intervalMinutes)
            If DateTime.Now > scheduledTime Then
                'If Scheduled Time is passed set Schedule for the next Interval.
                scheduledTime = scheduledTime.AddMinutes(intervalMinutes)
            End If
            'End If

            If DateTime.Now.TimeOfDay > TimeStart.TimeOfDay And DateTime.Now.TimeOfDay < TimeEnd.TimeOfDay Then
                'Indices details fetch and store
                Me.WriteToFile("NSEDetails entry started" & DateTime.Now.TimeOfDay.ToString)
                Dim tmpNSEIndicesDetails As NSEIndicesDetails
                tmpNSEIndicesDetails = New NSEIndicesDetails()
                tmpNSEIndicesDetails.getIndicesDetailsAndStore()
                Me.WriteToFile("NSEDetails entry End" & DateTime.Now.TimeOfDay.ToString)
                'Hourly Data fetch and entry
                Me.WriteToFile("hourlStockdata entry started" & DateTime.Now.TimeOfDay.ToString)
                Dim tmpHourlyStockQuote As HourlyStockQuote
                tmpHourlyStockQuote = New HourlyStockQuote()
                tmpHourlyStockQuote.GetAndStoreHourlyData()
                Me.WriteToFile("hourlStockdata entry End" & DateTime.Now.TimeOfDay.ToString)

            End If
            Dim timeSpan As TimeSpan = scheduledTime.Subtract(DateTime.Now)
            'Dim schedule As String = String.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds)

            'Me.WriteToFile((Convert.ToString("Simple Service scheduled to run after: ") & schedule) + " {0}")

            'Get the difference in Minutes between the Scheduled and Current Time.
            Dim dueTime As Integer = Convert.ToInt32(timeSpan.TotalMilliseconds)

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