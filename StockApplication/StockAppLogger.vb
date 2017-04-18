Imports System
Imports System.IO

Public Class StockAppLogger


    Shared logFile As String = My.Settings.ApplicationFileLocation & "\Log\log.txt"
    Shared errorLogFile As String = My.Settings.ApplicationFileLocation & "\Log\Errorlog.txt"
    Shared className As String

    Public Shared Function LogInfo(logMessage As String, Optional className As String = "") As Boolean
        Try
            If My.Settings.logLevel = "Info" Or My.Settings.logLevel = "Debug" Then
                Dim tmpWriter As StreamWriter
                tmpWriter = File.AppendText(logFile)
                'tmpWriter.Write(vbCrLf + "Log Entry : ")
                tmpWriter.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
                tmpWriter.Write(" Info " & className & "  : ")
                tmpWriter.Write("  :{0}", logMessage)
                tmpWriter.WriteLine("-------------------------------")
                tmpWriter.Flush()
                tmpWriter.Close()
            End If
        Catch ex As Exception
            WriteToFile("LogInfo throws error" & DateTime.Now.TimeOfDay.ToString)
            Return False
        End Try
        Return True
    End Function

    Public Shared Function LogError(logMessage As String, exec As Exception, Optional className As String = "") As Boolean

        Dim tmpWriter As StreamWriter
        Try
            tmpWriter = File.AppendText(errorLogFile)
            'tmpWriter.Write(vbCrLf + "Log Entry : ")
            tmpWriter.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
            tmpWriter.Write(" " & className & "  : Error")
            tmpWriter.Write("  :{0}", logMessage)
            tmpWriter.WriteLine(exec)
            tmpWriter.WriteLine("-------------------------------")
            tmpWriter.Flush()
            tmpWriter.Close()
        Catch ex As Exception
            WriteToFile("LogError throws error" & DateTime.Now.TimeOfDay.ToString)
            Return False
        End Try
        Return True
    End Function

    Public Shared Function Log(logMessage As String, Optional className As String = "") As Boolean
        Try
            If My.Settings.logLevel = "Debug" Then
                Dim tmpWriter As StreamWriter
                tmpWriter = File.AppendText(logFile)
                'tmpWriter.Write(vbCrLf + "Log Entry : ")
                tmpWriter.Write("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString())
                tmpWriter.Write(" Debug " & className & "  :")
                tmpWriter.WriteLine("  :{0}", logMessage)
                tmpWriter.WriteLine("-------------------------------")
                tmpWriter.Flush()
                tmpWriter.Close()
            End If
        Catch ex As Exception
            WriteToFile("Log throws error" & DateTime.Now.TimeOfDay.ToString)
            Return False
        End Try
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

    Private Shared Sub WriteToFile(text As String)
        Dim path As String = "D:\Tarun\StockApp\Log\UnhandledErrorLog.txt"
        Using writer As New StreamWriter(path, True)
            writer.WriteLine(String.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")))
            writer.Close()
        End Using
    End Sub

End Class
