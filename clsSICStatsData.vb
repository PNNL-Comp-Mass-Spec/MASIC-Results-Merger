Public Class clsSICStatsData : Implements IComparable(Of clsSICStatsData)

    Public ReadOnly Property FragScanNumber As Integer
    Public Property OptimalScanNumber As String
    Public Property PeakMaxIntensity As String
    Public Property PeakSignalToNoiseRatio As String
    Public Property FWHMInScans As String
    Public Property PeakArea As String
    Public Property ParentIonIntensity As String
    Public Property ParentIonMZ As String
    Public Property StatMomentsArea As String

    ''' <summary>
    ''' Constructor
    ''' </summary>
    ''' <param name="fragScanNum"></param>
    Public Sub New(fragScanNum As Integer)
        FragScanNumber = fragScanNum
    End Sub

    Public Function CompareTo(other As clsSICStatsData) As Integer Implements IComparable(Of clsSICStatsData).CompareTo
        If Me.FragScanNumber < other.FragScanNumber Then
            Return -1
        ElseIf Me.FragScanNumber > other.FragScanNumber Then
            Return 1
        Else
            Return 0
        End If
    End Function

End Class
