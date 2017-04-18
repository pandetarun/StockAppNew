Imports System.IO
Imports System.Threading
Imports System.Configuration
Imports StockApplication
Imports System.ServiceModel
Imports System.ServiceModel.Web
Imports System.ServiceModel.Description

Public Class GenerateIndicationSrv

    Protected Overrides Sub OnStart(ByVal args() As String)
        Me.WriteToFile("TechnicalIndicatorsGenerationService Service about to start at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
        Dim host As WebServiceHost = New WebServiceHost(GetType(GenerateTechnicalIndicator), New Uri("http://localhost:6080/"))
        Dim ep As ServiceEndpoint = host.AddServiceEndpoint(GetType(IGenerateTechnicalIndicator), New WebHttpBinding(), "")
        host.Open()
        Me.WriteToFile("TechnicalIndicatorsGenerationService Service started at " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"))
    End Sub

    Protected Overrides Sub OnStop()
        ' Add code here to perform any tear-down necessary to stop your service.
    End Sub

    Private Sub WriteToFile(text As String)
        Dim path As String = "D:\Tarun\StockApp\Log\TechnicalIndicatorsGenerationServiceLog.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub

End Class
