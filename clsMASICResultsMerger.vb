Option Strict On

' This class merges the contents of a tab-delimited peptide hit results file
' (e.g. from Sequest, XTandem, or Inspect) with the corresponding MASIC results files, 
' appending the relevant MASIC stats for each peptide hit result
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started November 26, 2008
'
' E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnl.gov/ or http://omics.pnl.gov
' -------------------------------------------------------------------------------
' 
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at 
' http://www.apache.org/licenses/LICENSE-2.0
'
' Notice: This computer software was prepared by Battelle Memorial Institute, 
' hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
' Department of Energy (DOE).  All rights in the computer software are reserved 
' by DOE on behalf of the United States Government and the Contractor as 
' provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
' WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
' SOFTWARE.  This notice including this sentence must appear on any copies of 
' this computer software.

Public Class clsMASICResultsMerger
    Inherits clsProcessFilesBaseClass

    Public Sub New()
        MyBase.mFileDate = "December 1, 2008"
        InitializeLocalVariables()
    End Sub

#Region "Constants and Enums"

    Public Const SIC_STATS_FILE_EXTENSION As String = "_SICStats.txt"
    Public Const SCAN_STATS_FILE_EXTENSION As String = "_ScanStats.txt"
    Public Const REPORTER_IONS_FILE_EXTENSION As String = "_ReporterIons.txt"

    Public Const DEFAULT_SCAN_NUMBER_COLUMN As Integer = 2

    ' Error codes specialized for this class
    Public Enum eResultsProcessorErrorCodes As Integer
        NoError = 0
        MissingMASICFiles = 1
        UnspecifiedError = -1
    End Enum

    Protected Enum eScanStatsColumns
        Dataset = 0
        ScanNumber = 1
        ScanTime = 2
        ScanType = 3
        TotalIonIntensity = 4
        BasePeakIntensity = 5
        BasePeakMZ = 6
        BasePeakSignalToNoiseRatio = 7
        IonCount = 8
        IonCountRaw = 9
    End Enum

    Protected Enum eSICStatsColumns
        Dataset = 0
        ParentIonIndex = 1
        MZ = 2
        SurveyScanNumber = 3
        FragScanNumber = 4
        OptimalPeakApexScanNumber = 5
        PeakApexOverrideParentIonIndex = 6
        CustomSICPeak = 7
        PeakScanStart = 8
        PeakScanEnd = 9
        PeakScanMaxIntensity = 10
        PeakMaxIntensity = 11
        PeakSignalToNoiseRatio = 12
        FWHMInScans = 13
        PeakArea = 14
        ParentIonIntensity = 15
        PeakBaselineNoiseLevel = 16
        PeakBaselineNoiseStDev = 17
        PeakBaselinePointsUsed = 18
        StatMomentsArea = 19
        CenterOfMassScan = 20
        PeakStDev = 21
        PeakSkew = 22
        PeakKSStat = 23
        StatMomentsDataCountUsed = 24
    End Enum

    Protected Enum eReporterIonStatsColumns
        Dataset = 0
        ScanNumber = 1
        CollisionMode = 2
        ParentIonMZ = 3
        BasePeakIntensity = 4
        BasePeakMZ = 5
        ReporterIonIntensityMax = 6
    End Enum
#End Region

#Region "Structures"

    Protected Structure udtMASICFileNamesType
        Public DatasetName As String
        Public ScanStatsFileName As String
        Public SICStatsFileName As String
        Public ReporterIonsFileName As String
        Public Sub Initialize()
            DatasetName = String.Empty
            ScanStatsFileName = String.Empty
            SICStatsFileName = String.Empty
            ReporterIonsFileName = String.Empty
        End Sub
    End Structure

    Protected Structure udtScanStatsType
        Public ScanNumber As Integer
        Public ElutionTime As String
        Public ScanType As String
        Public TotalIonIntensity As String
        Public BasePeakIntensity As String
        Public BasePeakMZ As String
        Public CollisionMode As String          ' Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)
        Public ReporterIonData As String        ' Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)
    End Structure

    Protected Structure udtSICStatsType
        Public FragScanNumber As Integer
        Public OptimalScanNumber As String
        Public PeakMaxIntensity As String
        Public PeakSignalToNoiseRatio As String
        Public FWHMInScans As String
        Public PeakArea As String
        Public ParentIonIntensity As String
        Public ParentIonMZ As String
        Public StatMomentsArea As String     
    End Structure

#End Region

#Region "Classwide Variables"
    Protected mWarnMissingParameterFileSection As Boolean
    Protected mLocalErrorCode As eResultsProcessorErrorCodes

    Protected mMASICResultsFolderPath As String = String.Empty
    Protected mScanNumberColumn As Integer          ' For the input file, defines which column tracks scan number; the first column is column 1 (not zero)
    Protected mSeparateByCollisionMode As Boolean   ' When true, then a separate output file will be created for each collision mode type; this is only possible if a _ReporterIons.txt file exists

#End Region

#Region "Properties"
    Public ReadOnly Property LocalErrorCode() As eResultsProcessorErrorCodes
        Get
            Return mLocalErrorCode
        End Get
    End Property

    Public Property MASICResultsFolderPath() As String
        Get
            If mMASICResultsFolderPath Is Nothing Then mMASICResultsFolderPath = String.Empty
            Return mMASICResultsFolderPath
        End Get
        Set(ByVal value As String)
            If value Is Nothing Then value = String.Empty
            mMASICResultsFolderPath = value
        End Set
    End Property
    Public Property SeparateByCollisionMode() As Boolean
        Get
            Return mSeparateByCollisionMode
        End Get
        Set(ByVal value As Boolean)
            mSeparateByCollisionMode = value
        End Set
    End Property

    Public Property ScanNumberColumn() As Integer
        Get
            Return mScanNumberColumn
        End Get
        Set(ByVal value As Integer)
            mScanNumberColumn = value
        End Set
    End Property
    Public Property WarnMissingParameterFileSection() As Boolean
        Get
            Return mWarnMissingParameterFileSection
        End Get
        Set(ByVal Value As Boolean)
            mWarnMissingParameterFileSection = Value
        End Set
    End Property
#End Region

    Protected Function FindMASICFiles(ByVal strMASICResultsFolder As String, ByVal strInputFilePath As String, ByRef udtMASICFileNames As udtMASICFileNamesType) As Boolean

        Dim blnSuccess As Boolean = False
        Dim strDatasetName As String
        Dim strCandidateFilePath As String

        Dim intCharIndex As Integer

        Try
            MyBase.mProgressStepDescription = "Looking for MASIC data files that correspond to " & System.IO.Path.GetFileName(strInputFilePath)
            ShowMessage(MyBase.mProgressStepDescription)

            ' Parse out the dataset name and parent folder from strInputFilePath
            strDatasetName = System.IO.Path.GetFileNameWithoutExtension(strInputFilePath)

            ' Use a Do loop to try various possible dataset names
            Do
                strCandidateFilePath = System.IO.Path.Combine(strMASICResultsFolder, strDatasetName & SIC_STATS_FILE_EXTENSION)

                If System.IO.File.Exists(strCandidateFilePath) Then
                    ' SICStats file was found
                    ' Update udtMASICFileNames, then look for the other files

                    udtMASICFileNames.DatasetName = strDatasetName
                    udtMASICFileNames.SICStatsFileName = System.IO.Path.GetFileName(strCandidateFilePath)

                    strCandidateFilePath = System.IO.Path.Combine(strMASICResultsFolder, strDatasetName & SCAN_STATS_FILE_EXTENSION)
                    If System.IO.File.Exists(strCandidateFilePath) Then
                        udtMASICFileNames.ScanStatsFileName = System.IO.Path.GetFileName(strCandidateFilePath)
                    End If

                    strCandidateFilePath = System.IO.Path.Combine(strMASICResultsFolder, strDatasetName & REPORTER_IONS_FILE_EXTENSION)
                    If System.IO.File.Exists(strCandidateFilePath) Then
                        udtMASICFileNames.ReporterIonsFileName = System.IO.Path.GetFileName(strCandidateFilePath)
                    End If

                    blnSuccess = True

                Else
                    ' Find the last underscore in strDatasetName, then remove it and any text after it
                    intCharIndex = strDatasetName.LastIndexOf("_"c)

                    If intCharIndex > 0 Then
                        strDatasetName = strDatasetName.Substring(0, intCharIndex)
                    Else
                        ' No more underscores; we're unable to determine the dataset name
                        Exit Do
                    End If
                End If

            Loop While Not blnSuccess

        Catch ex As Exception
            HandleException("Error in FindMASICFiles", ex)
            blnSuccess = False
        End Try

        Return blnSuccess
    End Function

    Protected Function FlattenArray(ByVal strSplitLine() As String, ByVal intIndexStart As Integer) As String
        Dim strText As String = String.Empty
        Dim strColumn As String

        Dim intIndex As Integer

        If Not strSplitLine Is Nothing AndAlso strSplitLine.Length > 0 Then
            For intIndex = intIndexStart To strSplitLine.Length - 1

                If strSplitLine(intIndex) Is Nothing Then
                    strColumn = String.Empty
                Else
                    strColumn = String.Copy(strSplitLine(intIndex))
                End If

                If intIndex > intIndexStart Then
                    strText &= ControlChars.Tab & strColumn
                Else
                    strText = String.Copy(strColumn)
                End If
            Next
        End If

        Return strText

    End Function

    Public Overrides Function GetErrorMessage() As String
        ' Returns "" if no error

        Dim strErrorMessage As String

        If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError Or _
           MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError Then
            Select Case mLocalErrorCode
                Case eResultsProcessorErrorCodes.NoError
                    strErrorMessage = ""
                Case eResultsProcessorErrorCodes.UnspecifiedError
                    strErrorMessage = "Unspecified localized error"
                Case eResultsProcessorErrorCodes.MissingMASICFiles
                    strErrorMessage = "Missing MASIC Files"
                Case Else
                    ' This shouldn't happen
                    strErrorMessage = "Unknown error state"
            End Select
        Else
            strErrorMessage = MyBase.GetBaseClassErrorMessage()
        End If

        Return strErrorMessage
    End Function

    Private Sub InitializeLocalVariables()
        MyBase.ShowMessages = False

        mMASICResultsFolderPath = String.Empty
        mScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN
        mSeparateByCollisionMode = False

        mWarnMissingParameterFileSection = True

        mLocalErrorCode = eResultsProcessorErrorCodes.NoError

    End Sub

    Private Function LoadParameterFileSettings(ByVal strParameterFilePath As String) As Boolean

        Const OPTIONS_SECTION As String = "MASICResultsMerger"

        Dim objSettingsFile As New XmlSettingsFileAccessor

        Dim intValue As Integer

        Try

            If strParameterFilePath Is Nothing OrElse strParameterFilePath.Length = 0 Then
                ' No parameter file specified; nothing to load
                Return True
            End If

            If Not System.IO.File.Exists(strParameterFilePath) Then
                ' See if strParameterFilePath points to a file in the same directory as the application
                strParameterFilePath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), System.IO.Path.GetFileName(strParameterFilePath))
                If Not System.IO.File.Exists(strParameterFilePath) Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.ParameterFileNotFound)
                    Return False
                End If
            End If

            If objSettingsFile.LoadSettings(strParameterFilePath) Then
                If Not objSettingsFile.SectionPresent(OPTIONS_SECTION) Then
                    ShowErrorMessage("The node '<section name=""" & OPTIONS_SECTION & """> was not found in the parameter file: " & strParameterFilePath)
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidParameterFile)
                    Return False
                Else
                    mScanNumberColumn = objSettingsFile.GetParam(OPTIONS_SECTION, "ScanNumberColumn", DEFAULT_SCAN_NUMBER_COLUMN)
                    mSeparateByCollisionMode = objSettingsFile.GetParam(OPTIONS_SECTION, "SeparateByCollisionMode", False)
                End If
            End If

        Catch ex As Exception
            HandleException("Error in LoadParameterFileSettings", ex)
            Return False
        End Try

        Return True

    End Function

    Protected Function MergePeptideHitAndMASICFiles(ByVal strInputFilePath As String, _
                                                    ByVal strOutputFolderPath As String, _
                                                    ByVal intScanStatsCount As Integer, _
                                                    ByRef udtScanStats() As udtScanStatsType, _
                                                    ByVal intSICStatsCount As Integer, _
                                                    ByRef udtSICStats() As udtSICStatsType, _
                                                    ByVal strReporterIonHeaders As String) As Boolean


        Dim srInFile As System.IO.StreamReader
        Dim swOutfile() As System.IO.StreamWriter
        Dim intLinesWritten() As Integer

        Dim intOutputFileCount As Integer
        Dim strOutputFilePaths() As String

        Dim strLineIn As String

        Dim strSplitLine() As String
        Dim strAdditionalHeaders() As String

        Dim strHeaderLine As String
        Dim strAddonColumns As String
        Dim strCollisionModeCurrentScan As String = String.Empty

        Dim strBlankAdditionalColumns As String = String.Empty
        Dim strBlankAdditionalSICColumns As String = String.Empty
        Dim strBlankAdditionalReporterIonColumns As String = String.Empty

        Dim intIndex As Integer
        Dim intIndexMatch As Integer
        Dim intFragIndexMatch As Integer
        Dim intOutFileIndex As Integer
        Dim intEmptyOutFileCount As Integer

        Dim intLinesRead As Integer
        Dim intScanNumber As Integer

        Dim intScanNumbers() As Integer
        Dim intScanNumberPointerArray() As Integer

        Dim intFragScanNumbers() As Integer
        Dim intFragScanNumberPointerArray() As Integer

        Dim htCollisionModes As System.Collections.Hashtable
        Dim objHashTableEntry As Object
        Dim strCollisionModes() As String

        Dim blnWriteReporterIonStats As Boolean = False
        Dim blnSuccess As Boolean = False

        Try
            ' Open the input file
            srInFile = New System.IO.StreamReader(New System.IO.FileStream(strInputFilePath, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read))

        Catch ex As Exception
            HandleException("Error opening the input file: " & strInputFilePath, ex)
            Return False
        End Try

        Try
            ' Define the output file path
            If strOutputFolderPath Is Nothing Then
                strOutputFolderPath = String.Empty
            End If

            htCollisionModes = New System.Collections.Hashtable

            If mSeparateByCollisionMode Then
                ' Construct a list of the different collision modes in udtScanStats()

                intOutputFileCount = 0
                ReDim strCollisionModes(9)

                For intIndex = 0 To intScanStatsCount - 1
                    If Not htCollisionModes.ContainsKey(udtScanStats(intIndex).CollisionMode.ToLower) Then
                        ' Store this collision mode in htCollisionModes; the value stored will be the index in strCollisionModes()
                        htCollisionModes.Add(udtScanStats(intIndex).CollisionMode.ToLower, intOutputFileCount)

                        If intOutputFileCount >= strCollisionModes.Length Then
                            ReDim Preserve strCollisionModes(strCollisionModes.Length * 2 - 1)
                        End If

                        strCollisionModes(intOutputFileCount) = String.Copy(udtScanStats(intIndex).CollisionMode)
                        intOutputFileCount += 1
                    End If
                Next intIndex

                ReDim strOutputFilePaths(intOutputFileCount - 1)

                For intIndex = 0 To intOutputFileCount - 1
                    If strCollisionModes(intIndex) Is Nothing OrElse strCollisionModes(intIndex).Length = 0 Then
                        strCollisionModes(intIndex) = "na"
                    End If
                    strOutputFilePaths(intIndex) = System.IO.Path.Combine(strOutputFolderPath, System.IO.Path.GetFileNameWithoutExtension(strInputFilePath) & "_" & strCollisionModes(intIndex) & "_PlusSICStats.txt")
                Next
            Else
                intOutputFileCount = 1
                ReDim strOutputFilePaths(0)
                strOutputFilePaths(0) = System.IO.Path.Combine(strOutputFolderPath, System.IO.Path.GetFileNameWithoutExtension(strInputFilePath) & "_PlusSICStats.txt")
            End If

            ReDim swOutfile(intOutputFileCount - 1)
            ReDim intLinesWritten(intOutputFileCount - 1)

            ' Open the output file(s)
            For intIndex = 0 To intOutputFileCount - 1
                swOutfile(intIndex) = New System.IO.StreamWriter(New System.IO.FileStream(strOutputFilePaths(intIndex), IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.Read))
            Next

        Catch ex As Exception
            HandleException("Error creating the merged output file", ex)
            Return False
        End Try


        Try
            MyBase.mProgressStepDescription = "Parsing " & System.IO.Path.GetFileName(strInputFilePath) & " and writing " & System.IO.Path.GetFileName(strOutputFilePaths(0))
            ShowMessage(MyBase.mProgressStepDescription)

            If mScanNumberColumn < 1 Then
                ' Assume the scan number is in the second column
                mScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN
            End If

            ' Create a lookup array of the scan numbers in udtScanStats()
            ReDim intScanNumbers(intScanStatsCount - 1)
            ReDim intScanNumberPointerArray(intScanStatsCount - 1)

            For intIndex = 0 To intScanStatsCount - 1
                intScanNumbers(intIndex) = udtScanStats(intIndex).ScanNumber
                intScanNumberPointerArray(intIndex) = intIndex
            Next

            Array.Sort(intScanNumbers, intScanNumberPointerArray)

            ' Create a lookup array of the frag scan numbers in udtSICStats()
            ReDim intFragScanNumbers(intSICStatsCount - 1)
            ReDim intFragScanNumberPointerArray(intSICStatsCount - 1)

            For intIndex = 0 To intSICStatsCount - 1
                intFragScanNumbers(intIndex) = udtSICStats(intIndex).FragScanNumber
                intFragScanNumberPointerArray(intIndex) = intIndex
            Next

            Array.Sort(intFragScanNumbers, intFragScanNumberPointerArray)


            ' Read from srInFile and write out to the file(s) in swOutFile

            intLinesRead = 0
            Do While srInFile.Peek >= 0
                strLineIn = srInFile.ReadLine

                If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                    intLinesRead += 1
                    strSplitLine = strLineIn.Split(ControlChars.Tab)

                    If intLinesRead = 1 Then
                        ' Write out an updated header line
                        If strSplitLine.Length >= mScanNumberColumn AndAlso Integer.TryParse(strSplitLine(mScanNumberColumn - 1), intScanNumber) Then
                            ' The input file doesn't have a header line; we will add one, using generic column names for the data in the input file

                            strHeaderLine = String.Empty
                            For intIndex = 0 To strSplitLine.Length - 1
                                If intIndex = 0 Then
                                    strHeaderLine = "Column" & intIndex.ToString("00")
                                Else
                                    strHeaderLine &= ControlChars.Tab & "Column" & intIndex.ToString("00")
                                End If
                            Next
                        Else
                            ' The input file does have a text-based 
                            strHeaderLine = String.Copy(strLineIn)

                            ' Clear strSplitLine so that this line gets skipped
                            ReDim strSplitLine(-1)
                        End If

                        ' Append the ScanStats columns
                        strAddonColumns = "ElutionTime" & ControlChars.Tab & _
                                          "ScanType" & ControlChars.Tab & _
                                          "TotalIonIntensity" & ControlChars.Tab & _
                                          "BasePeakIntensity" & ControlChars.Tab & _
                                          "BasePeakMZ" & ControlChars.Tab

                        ' Append the SICStats columns
                        strAddonColumns &= "Optimal_Scan_Number" & ControlChars.Tab & _
                                           "PeakMaxIntensity" & ControlChars.Tab & _
                                           "PeakSignalToNoiseRatio" & ControlChars.Tab & _
                                           "FWHMInScans" & ControlChars.Tab & _
                                           "PeakArea" & ControlChars.Tab & _
                                           "ParentIonIntensity" & ControlChars.Tab & _
                                           "ParentIonMZ" & ControlChars.Tab & _
                                           "StatMomentsArea"

                        strBlankAdditionalSICColumns = String.Empty & ControlChars.Tab & _
                                                       String.Empty & ControlChars.Tab & _
                                                       String.Empty & ControlChars.Tab & _
                                                       String.Empty & ControlChars.Tab & _
                                                       String.Empty & ControlChars.Tab & _
                                                       String.Empty & ControlChars.Tab & _
                                                       String.Empty & ControlChars.Tab & _
                                                       String.Empty

                        If Not strReporterIonHeaders Is Nothing AndAlso strReporterIonHeaders.Length > 0 Then
                            ' Append the reporter ion stats columns
                            strAddonColumns &= ControlChars.Tab & strReporterIonHeaders
                            blnWriteReporterIonStats = True
                        End If

                        For intIndex = 0 To intOutputFileCount - 1
                            swOutfile(intIndex).WriteLine(strHeaderLine & ControlChars.Tab & strAddonColumns)
                        Next

                        ' Count the number of tabs in strAddonColumns and populate strBlankAdditionalColumns
                        strAdditionalHeaders = strAddonColumns.Split(ControlChars.Tab)

                        strBlankAdditionalColumns = String.Empty
                        For intIndex = 0 To strAdditionalHeaders.Length - 2
                            strBlankAdditionalColumns &= ControlChars.Tab
                        Next


                        ' Count the number of tabs in strReporterIonHeaders and populate strBlankAdditionalReporterIonColumns
                        strAdditionalHeaders = strReporterIonHeaders.Split(ControlChars.Tab)

                        strBlankAdditionalReporterIonColumns = String.Empty
                        For intIndex = 0 To strAdditionalHeaders.Length - 2
                            strBlankAdditionalReporterIonColumns &= ControlChars.Tab
                        Next

                    End If

                    If strSplitLine.Length >= mScanNumberColumn AndAlso Integer.TryParse(strSplitLine(mScanNumberColumn - 1), intScanNumber) Then
                        ' Look for intScanNumber in intScanNumbers
                        intIndexMatch = Array.BinarySearch(intScanNumbers, intScanNumber)

                        If intIndexMatch < 0 Then
                            ' Match not found; use the blank columns in strBlankAdditionalColumns
                            strAddonColumns = String.Copy(strBlankAdditionalColumns)
                        Else
                            With udtScanStats(intScanNumberPointerArray(intIndexMatch))
                                strAddonColumns = .ElutionTime & ControlChars.Tab & _
                                                  .ScanType & ControlChars.Tab & _
                                                  .TotalIonIntensity & ControlChars.Tab & _
                                                  .BasePeakIntensity & ControlChars.Tab & _
                                                  .BasePeakMZ
                            End With

                            intFragIndexMatch = Array.BinarySearch(intFragScanNumbers, intScanNumber)

                            If intFragIndexMatch < 0 Then
                                ' Match not found; use the blank columns in strBlankAdditionalSICColumns
                                strAddonColumns &= ControlChars.Tab & String.Copy(strBlankAdditionalSICColumns)
                            Else
                                With udtSICStats(intFragScanNumberPointerArray(intFragIndexMatch))
                                    strAddonColumns &= ControlChars.Tab & _
                                                       .OptimalScanNumber & ControlChars.Tab & _
                                                       .PeakMaxIntensity & ControlChars.Tab & _
                                                       .PeakSignalToNoiseRatio & ControlChars.Tab & _
                                                       .FWHMInScans & ControlChars.Tab & _
                                                       .PeakArea & ControlChars.Tab & _
                                                       .ParentIonIntensity & ControlChars.Tab & _
                                                       .ParentIonMZ & ControlChars.Tab & _
                                                       .StatMomentsArea
                                End With
                            End If

                            If blnWriteReporterIonStats Then
                                strCollisionModeCurrentScan = String.Empty

                                If intIndexMatch < 0 Then
                                    strAddonColumns &= ControlChars.Tab & _
                                                        String.Empty & ControlChars.Tab & _
                                                        strBlankAdditionalReporterIonColumns
                                Else
                                    With udtScanStats(intScanNumberPointerArray(intIndexMatch))
                                        If .CollisionMode Is Nothing OrElse .CollisionMode.Length = 0 Then
                                            ' Collision mode is not defined; append blank columns
                                            strAddonColumns &= ControlChars.Tab & _
                                                               String.Empty & ControlChars.Tab & _
                                                               strBlankAdditionalReporterIonColumns
                                        Else
                                            ' Collision mode is defined
                                            strAddonColumns &= ControlChars.Tab & _
                                                               .CollisionMode & ControlChars.Tab & _
                                                               .ReporterIonData

                                            strCollisionModeCurrentScan = String.Copy(.CollisionMode)
                                        End If
                                    End With
                                End If

                            End If
                        End If

                        intOutFileIndex = 0
                        If mSeparateByCollisionMode AndAlso intOutputFileCount > 1 Then
                            If Not strCollisionModeCurrentScan Is Nothing Then
                                ' Determine the correct output file
                                objHashTableEntry = htCollisionModes(strCollisionModeCurrentScan)

                                If Not objHashTableEntry Is Nothing Then
                                    intOutFileIndex = CInt(objHashTableEntry)
                                End If

                            End If
                        End If

                        swOutfile(intOutFileIndex).WriteLine(strLineIn & ControlChars.Tab & strAddonColumns)
                        intLinesWritten(intOutFileIndex) += 1

                    End If
                End If
            Loop

            ' Close the output files
            If Not swOutfile Is Nothing Then
                For intIndex = 0 To intOutputFileCount - 1
                    If Not swOutfile(intIndex) Is Nothing Then
                        swOutfile(intIndex).Close()
                    End If
                Next
            End If

            ' See if any of the files had no data written to them
            ' If there are,then delete the empty output file
            ' However, retain at least one output file
            intEmptyOutFileCount = 0
            For intIndex = 0 To intOutputFileCount - 1
                If intLinesWritten(intIndex) = 0 Then
                    intEmptyOutFileCount += 1
                End If
            Next

            If intEmptyOutFileCount > 0 Then
                If intEmptyOutFileCount = intOutputFileCount Then
                    ' All output files are empty
                    ' Pretend the first output file actually contains data
                    intLinesWritten(0) = 1
                End If

                For intIndex = 0 To intOutputFileCount - 1
                    ' Wait 250 msec before continuing
                    System.Threading.Thread.Sleep(250)

                    If intLinesWritten(intIndex) = 0 Then
                        Try
                            System.IO.File.Delete(strOutputFilePaths(intIndex))
                        Catch ex As Exception
                            ' Ignore errors here
                        End Try
                    End If
                Next
            End If

            blnSuccess = True

        Catch ex As Exception
            HandleException("Error in MergePeptideHitAndMASICFiles", ex)
            blnSuccess = False
        Finally
            If Not srInFile Is Nothing Then
                srInFile.Close()
            End If
        End Try

        Return blnSuccess

    End Function

    ' Main processing function
    Public Overloads Overrides Function ProcessFile(ByVal strInputFilePath As String, ByVal strOutputFolderPath As String, ByVal strParameterFilePath As String, ByVal blnResetErrorCode As Boolean) As Boolean
        ' Returns True if success, False if failure

        Dim udtMASICFileNames As udtMASICFileNamesType
        Dim strMASICResultsFolder As String

        Dim intScanStatsCount As Integer
        Dim udtScanStats() As udtScanStatsType

        Dim intSICStatsCount As Integer
        Dim udtSICStats() As udtSICStatsType

        Dim strReporterIonHeaders As String = String.Empty

        Dim blnSuccess As Boolean

        If blnResetErrorCode Then
            SetLocalErrorCode(eResultsProcessorErrorCodes.NoError)
        End If

        If Not LoadParameterFileSettings(strParameterFilePath) Then
            ShowErrorMessage("Parameter file load error: " & strParameterFilePath)

            If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError Then
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidParameterFile)
            End If
            Return False
        End If

        Try
            If strInputFilePath Is Nothing OrElse strInputFilePath.Length = 0 Then
                ShowMessage("Input file name is empty")
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.InvalidInputFilePath)
            Else
                ' Note that CleanupFilePaths() will update mOutputFolderPath, which is used by LogMessage()
                If Not CleanupFilePaths(strInputFilePath, strOutputFolderPath) Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.FilePathError)
                Else

                    MyBase.ResetProgress()

                    If mMASICResultsFolderPath Is Nothing OrElse mMASICResultsFolderPath.Length = 0 Then
                        strMASICResultsFolder = System.IO.Path.GetDirectoryName(strInputFilePath)
                    Else
                        strMASICResultsFolder = String.Copy(mMASICResultsFolderPath)
                    End If

                    ' Look for the corresponding MASIC files in the input folder
                    udtMASICFileNames.Initialize()
                    blnSuccess = FindMASICFiles(strMASICResultsFolder, strInputFilePath, udtMASICFileNames)

                    If Not blnSuccess Then
                        ShowErrorMessage("Error: Unable to find the MASIC data files in " & strMASICResultsFolder)
                        SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMASICFiles)
                        Return False
                    Else
                        If udtMASICFileNames.SICStatsFileName.Length = 0 Then
                            ShowErrorMessage("Error: the SIC stats file was not found in " & strMASICResultsFolder & "; unable to continue")
                            SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMASICFiles)
                            Return False
                        ElseIf udtMASICFileNames.ScanStatsFileName.Length = 0 Then
                            ShowErrorMessage("Error: the Scan stats file was not found in " & strMASICResultsFolder & "; unable to continue")
                            SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMASICFiles)
                            Return False
                        End If
                    End If

                    ' Read and cache the MASIC data 
                    ReDim udtScanStats(-1)
                    ReDim udtSICStats(-1)

                    blnSuccess = ReadMASICData(strMASICResultsFolder, udtMASICFileNames, _
                                               intScanStatsCount, udtScanStats, _
                                               intSICStatsCount, udtSICStats, _
                                               strReporterIonHeaders)

                    If blnSuccess Then
                        ' Merge the MASIC data with the input file
                        blnSuccess = MergePeptideHitAndMASICFiles(strInputFilePath, strOutputFolderPath, _
                                                                  intScanStatsCount, udtScanStats, _
                                                                  intSICStatsCount, udtSICStats, _
                                                                   strReporterIonHeaders)
                    End If

                    If blnSuccess Then
                        ShowMessage(String.Empty, False)
                    Else
                        SetLocalErrorCode(eResultsProcessorErrorCodes.UnspecifiedError)
                        ShowErrorMessage("Error")
                    End If
                End If
            End If
        Catch ex As Exception
            HandleException("Error in ProcessFile", ex)
        End Try

        Return blnSuccess

    End Function

    Protected Function ReadMASICData(ByVal strSourceFolder As String, _
                                     ByVal udtMASICFileNames As udtMASICFileNamesType, _
                                     ByRef intScanStatsCount As Integer, _
                                     ByRef udtScanStats() As udtScanStatsType, _
                                     ByRef intSICStatsCount As Integer, _
                                     ByRef udtSICStats() As udtSICStatsType, _
                                     ByRef strReporterIonHeaders As String) As Boolean

        Dim blnSuccess As Boolean = False

        Try

            blnSuccess = ReadScanStatsFile(strSourceFolder, udtMASICFileNames.ScanStatsFileName, intScanStatsCount, udtScanStats)

            If blnSuccess Then
                blnSuccess = ReadSICStatsFile(strSourceFolder, udtMASICFileNames.SICStatsFileName, intSICStatsCount, udtSICStats)
            End If

            If blnSuccess AndAlso udtMASICFileNames.ReporterIonsFileName.Length > 0 Then
                blnSuccess = ReadReporterIonStatsFile(strSourceFolder, udtMASICFileNames.ReporterIonsFileName, _
                                                      intScanStatsCount, udtScanStats, _
                                                      strReporterIonHeaders)
            End If

        Catch ex As Exception
            HandleException("Error in ReadMASICData", ex)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Protected Function ReadScanStatsFile(ByVal strSourceFolder As String, _
                                         ByVal strScanStatsFileName As String, _
                                         ByRef intScanStatsCount As Integer, _
                                         ByRef udtScanStats() As udtScanStatsType) As Boolean

        Dim srInFile As System.IO.StreamReader

        Dim strLineIn As String
        Dim strSplitLine() As String

        Dim intLinesRead As Integer
        Dim intScanNumber As Integer

        Dim blnSuccess As Boolean = False

        Try
            ' Initialize the storage array
            intScanStatsCount = 0
            ReDim udtScanStats(999)

            MyBase.mProgressStepDescription = "Reading the MASIC Scan Stats file: " & strScanStatsFileName
            ShowMessage(MyBase.mProgressStepDescription)

            srInFile = New System.IO.StreamReader(New System.IO.FileStream(System.IO.Path.Combine(strSourceFolder, strScanStatsFileName), IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read))

            intLinesRead = 0
            Do While srInFile.Peek >= 0
                strLineIn = srInFile.ReadLine

                If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                    intLinesRead += 1
                    strSplitLine = strLineIn.Split(ControlChars.Tab)

                    If strSplitLine.Length >= eScanStatsColumns.BasePeakMZ + 1 AndAlso Integer.TryParse(strSplitLine(eScanStatsColumns.ScanNumber), intScanNumber) Then

                        If intScanStatsCount >= udtScanStats.Length Then
                            ' Reserve more room in udtScanStats
                            ReDim Preserve udtScanStats(udtScanStats.Length * 2 - 1)
                        End If

                        With udtScanStats(intScanStatsCount)
                            .ScanNumber = intScanNumber

                            ' Note: the remaining values are stored as strings to prevent the number format from changing
                            .ElutionTime = String.Copy(strSplitLine(eScanStatsColumns.ScanTime))
                            .ScanType = String.Copy(strSplitLine(eScanStatsColumns.ScanType))
                            .TotalIonIntensity = String.Copy(strSplitLine(eScanStatsColumns.TotalIonIntensity))
                            .BasePeakIntensity = String.Copy(strSplitLine(eScanStatsColumns.BasePeakIntensity))
                            .BasePeakMZ = String.Copy(strSplitLine(eScanStatsColumns.BasePeakMZ))

                            .CollisionMode = String.Empty
                            .ReporterIonData = String.Empty
                        End With

                        intScanStatsCount += 1
                    End If
                End If
            Loop

            ' Shrink udtScanStats
            ReDim Preserve udtScanStats(intScanStatsCount - 1)

            blnSuccess = True

        Catch ex As Exception
            HandleException("Error in ReadScanStatsFile", ex)
            blnSuccess = False
        Finally
            If Not srInFile Is Nothing Then
                srInFile.Close()
            End If
        End Try

        Return blnSuccess

    End Function

    Protected Function ReadSICStatsFile(ByVal strSourceFolder As String, _
                                        ByVal strSICStatsFileName As String, _
                                        ByRef intSICStatsCount As Integer, _
                                        ByRef udtSICStats() As udtSICStatsType) As Boolean

        Dim srInFile As System.IO.StreamReader

        Dim strLineIn As String
        Dim strSplitLine() As String

        Dim intLinesRead As Integer
        Dim intFragScanNumber As Integer

        Dim blnSuccess As Boolean = False

        Try
            ' Initialize the storage array
            ReDim udtSICStats(999)
            intSICStatsCount = 0

            MyBase.mProgressStepDescription = "Reading the MASIC SIC Stats file: " & strSICStatsFileName
            ShowMessage(MyBase.mProgressStepDescription)

            srInFile = New System.IO.StreamReader(New System.IO.FileStream(System.IO.Path.Combine(strSourceFolder, strSICStatsFileName), IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read))

            intLinesRead = 0
            Do While srInFile.Peek >= 0
                strLineIn = srInFile.ReadLine

                If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                    intLinesRead += 1
                    strSplitLine = strLineIn.Split(ControlChars.Tab)

                    If strSplitLine.Length >= eSICStatsColumns.StatMomentsArea + 1 AndAlso Integer.TryParse(strSplitLine(eSICStatsColumns.FragScanNumber), intFragScanNumber) Then

                        If intSICStatsCount >= udtSICStats.Length Then
                            ' Reserve more room in udtSICStats
                            ReDim Preserve udtSICStats(udtSICStats.Length * 2 - 1)
                        End If

                        With udtSICStats(intSICStatsCount)
                            .FragScanNumber = intFragScanNumber

                            ' Note: the remaining values are stored as strings to prevent the number format from changing
                            .OptimalScanNumber = String.Copy(strSplitLine(eSICStatsColumns.OptimalPeakApexScanNumber))
                            .PeakMaxIntensity = String.Copy(strSplitLine(eSICStatsColumns.PeakMaxIntensity))
                            .PeakSignalToNoiseRatio = String.Copy(strSplitLine(eSICStatsColumns.PeakSignalToNoiseRatio))
                            .FWHMInScans = String.Copy(strSplitLine(eSICStatsColumns.FWHMInScans))
                            .PeakArea = String.Copy(strSplitLine(eSICStatsColumns.PeakArea))
                            .ParentIonIntensity = String.Copy(strSplitLine(eSICStatsColumns.ParentIonIntensity))
                            .ParentIonMZ = String.Copy(strSplitLine(eSICStatsColumns.MZ))
                            .StatMomentsArea = String.Copy(strSplitLine(eSICStatsColumns.StatMomentsArea))
                        End With

                        intSICStatsCount += 1
                    End If
                End If
            Loop

            ' Shrink udtSICStats
            ReDim Preserve udtSICStats(intSICStatsCount - 1)

            blnSuccess = True

        Catch ex As Exception
            HandleException("Error in ReadSICStatsFile", ex)
            blnSuccess = False
        Finally
            If Not srInFile Is Nothing Then
                srInFile.Close()
            End If
        End Try

        Return blnSuccess

    End Function

    Protected Function ReadReporterIonStatsFile(ByVal strSourceFolder As String, _
                                                ByVal strReporterIonStatsFileName As String, _
                                                ByRef intScanStatsCount As Integer, _
                                                ByRef udtScanStats() As udtScanStatsType, _
                                                ByRef strReporterIonHeaders As String) As Boolean

        Dim srInFile As System.IO.StreamReader

        Dim strLineIn As String
        Dim strSplitLine() As String

        Dim intScanNumbers() As Integer
        Dim intScanNumberPointerArray() As Integer

        Dim intLinesRead As Integer
        Dim intScanNumber As Integer

        Dim intIndex As Integer
        Dim intIndexMatch As Integer
        Dim intWarningCount As Integer = 0

        Dim blnSuccess As Boolean = False

        Try
            ' Create a lookup array of the scan numbers in udtScanStats()
            ReDim intScanNumbers(intScanStatsCount - 1)
            ReDim intScanNumberPointerArray(intScanStatsCount - 1)

            For intIndex = 0 To intScanStatsCount - 1
                intScanNumbers(intIndex) = udtScanStats(intIndex).ScanNumber
                intScanNumberPointerArray(intIndex) = intIndex
            Next

            Array.Sort(intScanNumbers, intScanNumberPointerArray)

            strReporterIonHeaders = String.Empty


            MyBase.mProgressStepDescription = "Reading the MASIC Reporter Ion Stats file: " & strReporterIonStatsFileName
            ShowMessage(MyBase.mProgressStepDescription)

            srInFile = New System.IO.StreamReader(New System.IO.FileStream(System.IO.Path.Combine(strSourceFolder, strReporterIonStatsFileName), IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read))

            intLinesRead = 0
            Do While srInFile.Peek >= 0
                strLineIn = srInFile.ReadLine

                If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                    intLinesRead += 1
                    strSplitLine = strLineIn.Split(ControlChars.Tab)

                    If intLinesRead = 1 Then
                        ' This is the header line; we need to cache it

                        If strSplitLine.Length >= eReporterIonStatsColumns.ReporterIonIntensityMax + 1 Then
                            strReporterIonHeaders = strSplitLine(eReporterIonStatsColumns.CollisionMode)
                            strReporterIonHeaders &= ControlChars.Tab & FlattenArray(strSplitLine, eReporterIonStatsColumns.ReporterIonIntensityMax)
                        Else
                            ' There aren't enough columns in the header line; this is unexpected
                            strReporterIonHeaders = "Collision Mode" & ControlChars.Tab & "AdditionalReporterIonColumns"
                        End If

                    End If

                    If strSplitLine.Length >= eReporterIonStatsColumns.ReporterIonIntensityMax + 1 AndAlso Integer.TryParse(strSplitLine(eReporterIonStatsColumns.ScanNumber), intScanNumber) Then

                        ' Look for intScanNumber in intScanNumbers
                        intIndexMatch = Array.BinarySearch(intScanNumbers, intScanNumber)

                        If intIndexMatch < 0 Then
                            If intWarningCount < 10 Then
                                ShowMessage("Warning: " & REPORTER_IONS_FILE_EXTENSION & " file refers to scan " & intScanNumber.ToString & ", but that scan was not in the _ScanStats.txt file")
                            ElseIf intWarningCount = 10 Then
                                ShowMessage("Warning: " & REPORTER_IONS_FILE_EXTENSION & " file has 10 or more scan numbers that are not defined in the _ScanStats.txt file")
                            End If
                            intWarningCount += 1
                        Else
                            ' Use intScanNumberPointerArray() to determine the correct storage location in udtScanStats
                            With udtScanStats(intScanNumberPointerArray(intIndexMatch))
                                If .ScanNumber <> intScanNumber Then
                                    ' Scan number mismatch; this shouldn't happen
                                    ShowMessage("Error: Scan number mismatch in ReadReporterIonStatsFile: " & .ScanNumber.ToString & " vs. " & intScanNumber.ToString)
                                Else
                                    .CollisionMode = String.Copy(strSplitLine(eReporterIonStatsColumns.CollisionMode))
                                    .ReporterIonData = FlattenArray(strSplitLine, eReporterIonStatsColumns.ReporterIonIntensityMax)
                                End If
                            End With

                        End If

                    End If
                End If
            Loop

            blnSuccess = True

        Catch ex As Exception
            HandleException("Error in ReadSICStatsFile", ex)
            blnSuccess = False
        Finally
            If Not srInFile Is Nothing Then
                srInFile.Close()
            End If
        End Try

        Return blnSuccess

    End Function

    Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eResultsProcessorErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(ByVal eNewErrorCode As eResultsProcessorErrorCodes, ByVal blnLeaveExistingErrorCodeUnchanged As Boolean)

        If blnLeaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> eResultsProcessorErrorCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            mLocalErrorCode = eNewErrorCode

            If eNewErrorCode = eResultsProcessorErrorCodes.NoError Then
                If MyBase.ErrorCode = clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(clsProcessFilesBaseClass.eProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

End Class
