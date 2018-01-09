Public Class clsScanStatsData : Implements IComparable(Of clsScanStatsData)

    Public ReadOnly ScanNumber As Integer
    Public Property ElutionTime As String
    Public Property ScanType As String
    Public Property TotalIonIntensity As String
    Public Property BasePeakIntensity As String
    Public Property BasePeakMZ As String
    Public Property CollisionMode As String          ' Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)
    Public Property ReporterIonData As String        ' Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="scanNum"></param>
    Public Sub New(scanNum As Integer)
        ScanNumber = scanNum
    End Sub

    Public Function CompareTo(other As clsScanStatsData) As Integer Implements IComparable(Of clsScanStatsData).CompareTo
        If Me.ScanNumber < other.ScanNumber Then
            Return -1
        ElseIf Me.ScanNumber > other.ScanNumber Then
            Return 1
        Else
            Return 0
        End If
    End Function

End Class
