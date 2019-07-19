Public Class clsDatasetInfo

    Public ReadOnly Property DatasetName As String
    Public ReadOnly Property DatasetID As Integer

    Public Sub New(name As String, id As Integer)
        DatasetName = name
        DatasetID = id
    End Sub
End Class
