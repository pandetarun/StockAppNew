Public Class NSEindices
    Public Function getIndicesListAndStore() As Boolean
        Dim rawIndicesData As String

        rawIndicesData = Helper.GetDataFromUrl("https://www.nseindia.com/homepage/Indices1.json")

        Return True
    End Function

    Private Function parseAndPopulateObjects(ByVal rawIndicesData As String) As List(Of NSEindices)
        Dim NSEindicesList As NSEindices() = New NSEindices()


        Dim NSEIndicesData As NSEindices

        Return NSEindicesList
    End Function

    Private Function storeIndicesDatainDB() As Boolean

        Return True
    End Function

End Class
