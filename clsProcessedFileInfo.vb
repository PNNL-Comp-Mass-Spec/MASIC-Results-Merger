Public Class clsProcessedFileInfo

    Public Const COLLISION_MODE_NOT_DEFINED As String = "Collision_Mode_Not_Defined"

    Public Property BaseName As String

    Protected ReadOnly mOutputFiles As Dictionary(Of String, String)

    ''' <summary>
    ''' The Key is the collision mode and the value is the output file path
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property OutputFiles As Dictionary(Of String, String)
        Get
            Return mOutputFiles
        End Get
    End Property

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="baseDatasetName"></param>
    Public Sub New(baseDatasetName As String)
        BaseName = baseDatasetName
        mOutputFiles = New Dictionary(Of String, String)
    End Sub

    Public Sub AddOutputFile(collisionMode As String, outputFilePath As String)
        If String.IsNullOrEmpty(collisionMode) Then
            mOutputFiles.Add(COLLISION_MODE_NOT_DEFINED, outputFilePath)
        Else
            mOutputFiles.Add(collisionMode, outputFilePath)
        End If

    End Sub
End Class
