Imports System.Collections.Generic
Public Class NSEindices
    Dim indexName As String
    Dim lastPrice As Double
    Dim priceChange As Double
    Dim percentageChange As Double

    Public Function getIndicesListAndStore() As Boolean
        Dim rawIndicesData As String

        rawIndicesData = Helper.GetDataFromUrl("https://www.nseindia.com/homepage/Indices1.json")

        Return True
    End Function

    Private Function parseAndPopulateObjects(ByVal rawIndicesData As String) As List(Of NSEindices)
        Dim NSEindicesList As List(Of NSEindices) = New List(Of NSEindices)

        Dim NSEIndicesData As NSEindices

        Return NSEindicesList
    End Function

    Private Function storeIndicesDatainDB() As Boolean

        Return True
    End Function

End Class
