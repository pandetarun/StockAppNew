' NOTE: You can use the "Rename" command on the context menu to change the class name "GenerateTechnicalIndicator" in both code and config file together.
Imports System.IO
Imports System.Threading
Imports System.Configuration
Imports StockApplication

Public Class GenerateTechnicalIndicator
    Implements IGenerateTechnicalIndicator

    Public Async Sub GenerateIndication() Implements IGenerateTechnicalIndicator.GenerateIndication
        WriteToFile("TechnicalIndicatorsGeneration Call received" & DateTime.Now.TimeOfDay.ToString)
        Await Task.Delay(50)
        GenerateMovingAverageIndiction()
        WriteToFile("TechnicalIndicatorsGeneration call finished" & DateTime.Now.TimeOfDay.ToString)
    End Sub


    Private Sub GenerateMovingAverageIndiction()

    End Sub

    Private Sub GenerateBollingerBandIndiction()

    End Sub

    Private Sub GenerateMACDIndiction()

    End Sub

    Private Sub WriteToFile(text As String)
        Dim path As String = "D:\Tarun\StockApp\Log\TechnicalIndicatorsGenerationServiceLog.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub
End Class
