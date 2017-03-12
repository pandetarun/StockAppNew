Imports System
Imports System.IO

Public Class StockAppLogger
    Dim tmpWriter As StreamWriter
    Shared tmpStockAppLogger As StockAppLogger
    Dim logFile As String
    Dim className As String

    Public Shared Function InitializeLogger(ByVal className As String) As StockAppLogger
        If tmpStockAppLogger Is Nothing Then
            tmpStockAppLogger = New StockAppLogger()
            tmpStockAppLogger.logFile = My.Settings.ApplicationFileLocation & "\log.txt"
        End If
        tmpStockAppLogger.className = className
        Return tmpStockAppLogger
    End Function

    Public Function LogError(logMessage As String, exec As Exception) As Boolean
        tmpWriter = File.AppendText(tmpStockAppLogger.logFile)
        'tmpWriter.Write(vbCrLf + "Log Entry : ")
        tmpWriter.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
        tmpWriter.Write(" " & tmpStockAppLogger.className & "  : Error")
        tmpWriter.Write("  :{0}", logMessage)
        tmpWriter.WriteLine(exec)
        tmpWriter.WriteLine("-------------------------------")
        tmpWriter.Flush()
        tmpWriter.Close()
        Return True
    End Function

    Public Function Log(logMessage As String) As Boolean
        If My.Settings.logLevel = "Debug" Then
            tmpWriter = File.AppendText(tmpStockAppLogger.logFile)
            'tmpWriter.Write(vbCrLf + "Log Entry : ")
            tmpWriter.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
            tmpWriter.Write(" " & tmpStockAppLogger.className & "  :")
            tmpWriter.WriteLine("  :{0}", logMessage)
            tmpWriter.WriteLine("-------------------------------")
            tmpWriter.Flush()
            tmpWriter.Close()
        End If
        Return True
    End Function

    Public Shared Sub DumpLog(r As StreamReader)
        Dim line As String
        line = r.ReadLine()
        While Not (line Is Nothing)
            Console.WriteLine(line)
            line = r.ReadLine()
        End While
    End Sub
End Class
