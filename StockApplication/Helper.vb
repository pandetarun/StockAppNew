Imports System.Net
Public Class Helper

    Public Shared Function GetDataFromUrl(ByVal urlToGet As String) As String
        Dim serverUrl As String
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse
        Dim strResponse As String
        Dim counter As Integer
        counter = 0
        StockAppLogger.Log("GetDataFromUrl Start", "Helper")
Retry:
        Try
            serverUrl = urlToGet
            request = WebRequest.Create(serverUrl)
            request = CreateRequest(request)
            request.Timeout = 10000
            response = request.GetResponse()
            Dim sReader As New IO.StreamReader(response.GetResponseStream)
            strResponse = sReader.ReadToEnd()
            response.Close()
        Catch ex As Exception
            StockAppLogger.LogError("Error Occurred in getting data from URL = " & urlToGet, ex, "Helper")
            If counter = 0 Then
                counter = counter + 1
                StockAppLogger.LogInfo("Retrying fetch for URL = " & urlToGet, "Helper")
                GoTo Retry
            End If
            StockAppLogger.LogInfo("Error Retry failed in getting data from URL = " & urlToGet, "Helper")
            Return Nothing
        End Try
        StockAppLogger.Log("GetDataFromUrl End", "Helper")
        Return strResponse
    End Function

    Private Shared Function CreateRequest(ByRef request As HttpWebRequest) As HttpWebRequest
        StockAppLogger.Log("CreateRequest Start", "Helper")
        request.UseDefaultCredentials = True
        request.Method = WebRequestMethods.Http.Post
        request.ContentType = "application/json"
        With request
            .UserAgent = "User-Agent: Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.0.10) Gecko/20070216 Firefox/1.5.0.10"
            .KeepAlive = False
            .Method = "GET"
            .AllowAutoRedirect = True
            .Headers.Add("Pragma", "no-cache")
            .Headers.Add("Cache-Control", "no-cache")
            .Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
            .ServicePoint.Expect100Continue = False
        End With
        StockAppLogger.Log("CreateRequest End", "Helper")
        Return request
    End Function
End Class
