Imports System
Imports System.IO

Public Class StockAppLogger
    Dim tmpWriter As StreamWriter
    Shared tmpStockAppLogger As StockAppLogger
    Dim logFile As String

    Public Function InitializeLogger() As StockAppLogger
        If tmpStockAppLogger Is Nothing Then
            tmpStockAppLogger = New StockAppLogger()
            tmpStockAppLogger.logFile = "D:\Tarun\StockApp\StockApplication\log.txt"
        End If
        Return tmpStockAppLogger
    End Function

    Public Sub Log(logMessage As String, exec As Exception)
        tmpWriter.Write(vbCrLf + "Log Entry : ")
        tmpWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
        DateTime.Now.ToLongDateString())
        tmpWriter.WriteLine("  :")
        tmpWriter.WriteLine("  :{0}", logMessage)
        tmpWriter.WriteLine(exec)
        tmpWriter.WriteLine("-------------------------------")
    End Sub

    Public Sub Log(logMessage As String)
        tmpWriter.Write(vbCrLf + "Log Entry : ")
        tmpWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
        DateTime.Now.ToLongDateString())
        tmpWriter.WriteLine("  :")
        tmpWriter.WriteLine("  :{0}", logMessage)
        tmpWriter.WriteLine("-------------------------------")
    End Sub

    Public Shared Sub DumpLog(r As StreamReader)
        Dim line As String
        line = r.ReadLine()
        While Not (line Is Nothing)
            Console.WriteLine(line)
            line = r.ReadLine()
        End While
    End Sub
End Class
