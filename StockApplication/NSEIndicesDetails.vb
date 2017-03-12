Imports System.Collections.Generic
Imports System
Imports System.IO

Public Class NSEIndicesDetails
    Dim pathOfFileofURLs As String = "D:\Tarun\StockApp\URLs.txt"

    Public Function getIndicesDetailsAndStore() As Boolean
        Dim NSEindicesUrlsList As List(Of String) = New List(Of String)
        NSEindicesUrlsList = GetNSEUrlForAllIndices()

        Return True
    End Function

    Public Function GetNSEUrlForAllIndices() As List(Of String)
        Dim NSEindicesUrlsList As List(Of String) = New List(Of String)
        Dim reader = File.OpenText(pathOfFileofURLs)
        Dim line As String = Nothing

        While (reader.Peek() <> -1)
            line = reader.ReadLine()
            NSEindicesUrlsList.Add(line)
        End While
        Return NSEindicesUrlsList
    End Function

    Public Function getAllIndicesDetails() As List(Of String)
        Dim NSEindicesDetailsList As List(Of String) = New List(Of String)


        Return NSEindicesDetailsList
    End Function
End Class
