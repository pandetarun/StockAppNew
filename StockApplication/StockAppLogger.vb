Imports System
Imports System.IO

Public Class StockAppLogger


    Shared logFile As String = My.Settings.ApplicationFileLocation & "\log.txt"
    Shared errorLogFile As String = My.Settings.ApplicationFileLocation & "\Errorlog.txt"
    Shared className As String

    Public Shared Function LogError(logMessage As String, exec As Exception) As Boolean

        Dim tmpWriter As StreamWriter
        tmpWriter = File.AppendText(errorLogFile)
        'tmpWriter.Write(vbCrLf + "Log Entry : ")
        tmpWriter.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
        tmpWriter.Write(" " & className & "  : Error")
        tmpWriter.Write("  :{0}", logMessage)
        tmpWriter.WriteLine(exec)
        tmpWriter.WriteLine("-------------------------------")
        tmpWriter.Flush()
        tmpWriter.Close()
        Return True
    End Function

    Public Shared Function Log(logMessage As String) As Boolean
        If My.Settings.logLevel = "Debug" Then
            Dim tmpWriter As StreamWriter
            tmpWriter = File.AppendText(logFile)
            'tmpWriter.Write(vbCrLf + "Log Entry : ")
            tmpWriter.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
            tmpWriter.Write(" " & className & "  :")
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
