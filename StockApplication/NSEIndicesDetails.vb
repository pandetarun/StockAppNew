Imports System.Collections.Generic
Imports System
Imports System.IO

Public Class NSEIndicesDetails
    Dim pathOfFileofURLs As String = My.Settings.ApplicationFileLocation & "\URLs.txt"

    Public Function getIndicesDetailsAndStore() As Boolean
        Dim NSEindicesUrlsList As List(Of String) = New List(Of String)
        NSEindicesUrlsList = GetNSEUrlForAllIndices()
        getAllIndicesDetails(NSEindicesUrlsList)


        Return True
    End Function

    Private Function GetNSEUrlForAllIndices() As List(Of String)
        Dim NSEindicesUrlsList As List(Of String) = New List(Of String)
        Dim reader = File.OpenText(pathOfFileofURLs)
        Dim line As String = Nothing

        While (reader.Peek() <> -1)
            line = reader.ReadLine()
            NSEindicesUrlsList.Add(line)
        End While
        Return NSEindicesUrlsList
    End Function

    Private Function getAllIndicesDetails(ByVal NSEindicesUrlsList As List(Of String)) As List(Of String)
        Dim NSEindicesDetailsList As List(Of String) = New List(Of String)

        For Each NSEIndicesURL In NSEindicesUrlsList
            getNSEIndicesDetails(NSEIndicesURL)
        Next

        Return NSEindicesDetailsList
    End Function

    Private Function getNSEIndicesDetails(ByVal NSEIndicesURL As String)
        Dim rawIndicesData As String

        rawIndicesData = Helper.GetDataFromUrl(NSEIndicesURL)
    End Function

End Class
