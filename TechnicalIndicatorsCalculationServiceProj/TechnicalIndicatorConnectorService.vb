Imports System.IO
Imports System.Threading
Imports System.Configuration
Imports StockApplication

' NOTE: You can use the "Rename" command on the context menu to change the class name "ConnectorService" in both code and config file together.
Public Class TechnicalIndicatorConnectorService
    Implements IConnectorService

    Public Async Sub ProcessIntraDayTechnicalIndicators() Implements IConnectorService.ProcessIntraDayTechnicalIndicators
        WriteToFile("TechnicalIndicatorConnectorService Call received" & DateTime.Now.TimeOfDay.ToString)
        Await Task.Delay(50)
        CalculteIntraDayTechnicalIndicators()
        WriteToFile("TechnicalIndicatorConnectorService call finished" & DateTime.Now.TimeOfDay.ToString)
    End Sub

    Public Async Sub ProcessDailyTechnicalIndicators() Implements IConnectorService.ProcessDailyTechnicalIndicators
        WriteToFile("TechnicalIndicatorConnectorService Call received" & DateTime.Now.TimeOfDay.ToString)
        Await Task.Delay(50)
        CalculteDailyTechnicalIndicators()
        WriteToFile("TechnicalIndicatorConnectorService call finished" & DateTime.Now.TimeOfDay.ToString)
    End Sub

    Private Sub WriteToFile(text As String)
        Dim path As String = "D:\Tarun\StockApp\Log\TechnicalIndicatorsCalculationServiceLog.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub

    Private Sub CalculteIntraDayTechnicalIndicators()

        Try
            Me.WriteToFile("IntraDay calculation started" & DateTime.Now.TimeOfDay.ToString)
            Dim tmpCalculateIntradayIndicators As CalculateTechnicalIndicators = New CalculateTechnicalIndicators()
            tmpCalculateIntradayIndicators.CalculateIntradayIndicators()
            Me.WriteToFile("IntraDay calculation End" & DateTime.Now.TimeOfDay.ToString)
        Catch ex As Exception
            WriteToFile("TechnicalIndicatorConnectorService Error in calculating intraday data " + ex.Message + ex.StackTrace)
            ''Stop the Windows Service.
            'Using serviceController As New System.ServiceProcess.ServiceController("StockAppDataDownload")
            '    serviceController.[Stop]()
            'End Using
        End Try
    End Sub

    Public Sub CalculteDailyTechnicalIndicators()

        Try
            Me.WriteToFile("Daily calculation started" & DateTime.Now.TimeOfDay.ToString)
            Dim tmpCalculateIntradayIndicators As CalculateTechnicalIndicators = New CalculateTechnicalIndicators()
            tmpCalculateIntradayIndicators.CalculateIntradayIndicators()
            Me.WriteToFile("Daily calculation End" & DateTime.Now.TimeOfDay.ToString)
        Catch ex As Exception
            WriteToFile("TechnicalIndicatorConnectorService Error in calculating daily data " + ex.Message + ex.StackTrace)
            ''Stop the Windows Service.
            'Using serviceController As New System.ServiceProcess.ServiceController("StockAppDataDownload")
            '    serviceController.[Stop]()
            'End Using
        End Try
    End Sub

End Class
