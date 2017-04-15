Imports System.IO
Imports System.Threading
Imports System.Configuration
Imports StockApplication
Imports System.ServiceModel
Imports System.ServiceModel.Web
Imports System.ServiceModel.Description

Public Class TechnicalIndicatorsCalculationService

    Protected Overrides Sub OnStart(ByVal args() As String)
        Me.WriteToFile("TechnicalIndicatorsCalculationService Service about to start at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Dim host As WebServiceHost = New WebServiceHost(GetType(TechnicalIndicatorConnectorService), New Uri("http://localhost:6060/"))
        Dim ep As ServiceEndpoint = host.AddServiceEndpoint(GetType(IConnectorService), New WebHttpBinding(), "")
        host.Open()
        Me.WriteToFile("TechnicalIndicatorsCalculationService Service started at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
    End Sub

    Protected Overrides Sub OnStop()
        Me.WriteToFile("TechnicalIndicatorsCalculationService Service stopped at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
    End Sub

    Private Sub WriteToFile(text As String)
        Dim path As String = "D:\Tarun\StockApp\Log\TechnicalIndicatorsCalculationServiceLog.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub
End Class
