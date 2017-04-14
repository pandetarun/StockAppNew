Imports System.IO
Imports System.Threading
Imports System.Configuration
Imports StockApplication

' NOTE: You can use the "Rename" command on the context menu to change the class name "ConnectorService" in both code and config file together.
Public Class TechnicalIndicatorConnectorService
    Implements IConnectorService

    Public Sub DoWork() Implements IConnectorService.DoWork
        WriteToFile("Connector Called")

    End Sub

    Private Sub WriteToFile(text As String)
        Dim path As String = "D:\Tarun\StockApp\Log\TechnicalIndicatorsCalculationServiceLog.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub

    Public Sub CalculteIntraDayTechnicalIndicators()

        Try
            'IntraDay moving average starts
            Me.WriteToFile("IntraDayMovingAverage entry started" & DateTime.Now.TimeOfDay.ToString)
            Dim tmpMovingAverageCalculation As MovingAverage = New MovingAverage()
            tmpMovingAverageCalculation.CalculateAndStoreIntraDayMA()
            Me.WriteToFile("IntraDayMovingAverage entry End" & DateTime.Now.TimeOfDay.ToString)
            'IntraDay bollinger band starts
            Me.WriteToFile("IntraDay Bollinger Band entry started" & DateTime.Now.TimeOfDay.ToString)
            Dim tmpBBCalculation As BollingerBands = New BollingerBands()
            tmpBBCalculation.CalculateAndStoreIntradayBollingerBands()
            Me.WriteToFile("IntraDay Bollinger Band entry End" & DateTime.Now.TimeOfDay.ToString)
            'IntraDay MACD starts
            Me.WriteToFile("IntraDay MACD entry started" & DateTime.Now.TimeOfDay.ToString)
            Dim tmpMACDCalculation As MACD = New MACD()
            tmpMACDCalculation.CalculateAndStoreIntraDayMACD()
            Me.WriteToFile("IntraDay MACD entry End" & DateTime.Now.TimeOfDay.ToString)
        Catch ex As Exception
            WriteToFile("StockAppDataDownload Service Error in getting stock data " + ex.Message + ex.StackTrace)
            'Stop the Windows Service.
            Using serviceController As New System.ServiceProcess.ServiceController("StockAppDataDownload")
                serviceController.[Stop]()
            End Using
        End Try
    End Sub

    Public Sub CalculteDailyTechnicalIndicators()

        Try
            'Daily moving average starts
            Me.WriteToFile("DailyMovingAverage entry started" & DateTime.Now.TimeOfDay.ToString)
            Dim tmpMovingAverageCalculation As MovingAverage = New MovingAverage()
            tmpMovingAverageCalculation.CalculateAndStoreDayMA()
            Me.WriteToFile("DailyMovingAverage entry End" & DateTime.Now.TimeOfDay.ToString)
        Catch ex As Exception
            WriteToFile("StockAppDataDownload Service Error in getting stock data " + ex.Message + ex.StackTrace)
            'Stop the Windows Service.
            Using serviceController As New System.ServiceProcess.ServiceController("StockAppDataDownload")
                serviceController.[Stop]()
            End Using
        End Try
    End Sub

End Class
