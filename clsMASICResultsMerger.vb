Option Strict On
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Threading
Imports PHRPReader
Imports PRISM

' This class merges the contents of a tab-delimited peptide hit results file
' (e.g. from Sequest, XTandem, or MSGF+) with the corresponding MASIC results files,
' appending the relevant MASIC stats for each peptide hit result
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started November 26, 2008
'
' E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com
' Website: http://ncrr.pnnl.gov/ or http://omics.pnnl.gov
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
        MyBase.mFileDate = "November 21, 2013"
        InitializeLocalVariables()
    End Sub

#Region "Constants and Enums"

    Public Const SIC_STATS_FILE_EXTENSION As String = "_SICStats.txt"
    Public Const SCAN_STATS_FILE_EXTENSION As String = "_ScanStats.txt"
    Public Const REPORTER_IONS_FILE_EXTENSION As String = "_ReporterIons.txt"

    Public Const RESULTS_SUFFIX As String = "_PlusSICStats.txt"
    Public Const DEFAULT_SCAN_NUMBER_COLUMN As Integer = 2

    Protected Const SIC_STAT_COLUMN_COUNT_TO_ADD As Integer = 7

    ' Error codes specialized for this class
    Public Enum eResultsProcessorErrorCodes As Integer
        NoError = 0
        MissingMASICFiles = 1
        MissingMageFiles = 2
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

    Protected Structure udtDatasetInfoType
        Public DatasetName As String
        Public DatasetID As Integer
    End Structure

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

    Protected Structure udtScanStatsType : Implements IComparable(Of udtScanStatsType)

        Public ScanNumber As Integer
        Public ElutionTime As String
        Public ScanType As String
        Public TotalIonIntensity As String
        Public BasePeakIntensity As String
        Public BasePeakMZ As String
        Public CollisionMode As String          ' Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)
        Public ReporterIonData As String        ' Comes from _ReporterIons.txt file (Nothing if the file doesn't exist)

        Public Function CompareTo(other As udtScanStatsType) As Integer Implements IComparable(Of udtScanStatsType).CompareTo
            If Me.ScanNumber < other.ScanNumber Then
                Return -1
            ElseIf Me.ScanNumber > other.ScanNumber Then
                Return 1
            Else
                Return 0
            End If
        End Function
    End Structure

    Protected Structure udtSICStatsType : Implements IComparable(Of udtSICStatsType)
        Public FragScanNumber As Integer
        Public OptimalScanNumber As String
        Public PeakMaxIntensity As String
        Public PeakSignalToNoiseRatio As String
        Public FWHMInScans As String
        Public PeakArea As String
        Public ParentIonIntensity As String
        Public ParentIonMZ As String
        Public StatMomentsArea As String

        Public Function CompareTo(other As udtSICStatsType) As Integer Implements IComparable(Of udtSICStatsType).CompareTo
            If Me.FragScanNumber < other.FragScanNumber Then
                Return -1
            ElseIf Me.FragScanNumber > other.FragScanNumber Then
                Return 1
            Else
                Return 0
            End If
        End Function
    End Structure

#End Region

#Region "Classwide Variables"

    Protected mLocalErrorCode As eResultsProcessorErrorCodes

    Protected mMASICResultsFolderPath As String = String.Empty
    ' For the input file, defines which column tracks scan number; the first column is column 1 (not zero)
    ' When true, then a separate output file will be created for each collision mode type; this is only possible if a _ReporterIons.txt file exists

    Protected WithEvents mPHRPReader As clsPHRPReader

    ' For each KeyValuePair, the key is the base file name and the values are the output file paths for the base file
    ' There will be one output file for each base file if mSeparateByCollisionMode=false; multiple files if it is true
    Protected mProcessedDatasets As List(Of clsProcessedFileInfo)

#End Region

#Region "Properties"
    Public ReadOnly Property LocalErrorCode As eResultsProcessorErrorCodes
        Get
            Return mLocalErrorCode
        End Get
    End Property

    Public Property MageResults As Boolean

    Public Property MASICResultsFolderPath As String
        Get
            If mMASICResultsFolderPath Is Nothing Then mMASICResultsFolderPath = String.Empty
            Return mMASICResultsFolderPath
        End Get
        Set
            If Value Is Nothing Then Value = String.Empty
            mMASICResultsFolderPath = Value
        End Set
    End Property

    Public ReadOnly Property ProcessedDatasets As List(Of clsProcessedFileInfo)
        Get
            Return mProcessedDatasets
        End Get
    End Property

    Public Property SeparateByCollisionMode As Boolean

    Public Property ScanNumberColumn As Integer

    Public Property WarnMissingParameterFileSection As Boolean

#End Region

    Protected Function FindMASICFiles(strMASICResultsFolder As String, udtDatasetInfo As udtDatasetInfoType, ByRef udtMASICFileNames As udtMASICFileNamesType) As Boolean

        Dim blnSuccess = False
        Dim strDatasetName As String
        Dim blnTriedDatasetID = False

        Dim strCandidateFilePath As String

        Dim intCharIndex As Integer

        Try
            Console.WriteLine()
            MyBase.mProgressStepDescription = "Looking for MASIC data files that correspond to " & udtDatasetInfo.DatasetName
            ShowMessage(MyBase.mProgressStepDescription)

            strDatasetName = String.Copy(udtDatasetInfo.DatasetName)

            ' Use a Do loop to try various possible dataset names
            Do
                strCandidateFilePath = Path.Combine(strMASICResultsFolder, strDatasetName & SIC_STATS_FILE_EXTENSION)

                If File.Exists(strCandidateFilePath) Then
                    ' SICStats file was found
                    ' Update udtMASICFileNames, then look for the other files

                    udtMASICFileNames.DatasetName = strDatasetName
                    udtMASICFileNames.SICStatsFileName = Path.GetFileName(strCandidateFilePath)

                    strCandidateFilePath = Path.Combine(strMASICResultsFolder, strDatasetName & SCAN_STATS_FILE_EXTENSION)
                    If File.Exists(strCandidateFilePath) Then
                        udtMASICFileNames.ScanStatsFileName = Path.GetFileName(strCandidateFilePath)
                    End If

                    strCandidateFilePath = Path.Combine(strMASICResultsFolder, strDatasetName & REPORTER_IONS_FILE_EXTENSION)
                    If File.Exists(strCandidateFilePath) Then
                        udtMASICFileNames.ReporterIonsFileName = Path.GetFileName(strCandidateFilePath)
                    End If

                    blnSuccess = True

                Else
                    ' Find the last underscore in strDatasetName, then remove it and any text after it
                    intCharIndex = strDatasetName.LastIndexOf("_"c)

                    If intCharIndex > 0 Then
                        strDatasetName = strDatasetName.Substring(0, intCharIndex)
                    Else
                        If Not blnTriedDatasetID AndAlso udtDatasetInfo.DatasetID > 0 Then
                            strDatasetName = udtDatasetInfo.DatasetID & "_" & udtDatasetInfo.DatasetName
                            blnTriedDatasetID = True
                        Else
                            ' No more underscores; we're unable to determine the dataset name
                            Exit Do
                        End If

                    End If
                End If

            Loop While Not blnSuccess

        Catch ex As Exception
            HandleException("Error in FindMASICFiles", ex)
            blnSuccess = False
        End Try

        Return blnSuccess
    End Function

    Protected Function GetAdditionalMASICHeaders() As List(Of String)

        Dim lstAddonColumns = New List(Of String)

        ' Append the ScanStats columns
        lstAddonColumns.Add("ElutionTime")
        lstAddonColumns.Add("ScanType")
        lstAddonColumns.Add("TotalIonIntensity")
        lstAddonColumns.Add("BasePeakIntensity")
        lstAddonColumns.Add("BasePeakMZ")

        ' Append the SICStats columns
        lstAddonColumns.Add("Optimal_Scan_Number")
        lstAddonColumns.Add("PeakMaxIntensity")
        lstAddonColumns.Add("PeakSignalToNoiseRatio")
        lstAddonColumns.Add("FWHMInScans")
        lstAddonColumns.Add("PeakArea")
        lstAddonColumns.Add("ParentIonIntensity")
        lstAddonColumns.Add("ParentIonMZ")
        lstAddonColumns.Add("StatMomentsArea")

        Return lstAddonColumns
    End Function

    Protected Function FlattenList(lstData As List(Of String)) As String

        Dim sbFlattened = New StringBuilder

        For intIndex = 0 To lstData.Count - 1
            If intIndex > 0 Then
                sbFlattened.Append(ControlChars.Tab)
            End If
            sbFlattened.Append(lstData(intIndex))
        Next

        Return sbFlattened.ToString()

    End Function

    Protected Function FlattenArray(strSplitLine() As String, intIndexStart As Integer) As String
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

        If MyBase.ErrorCode = eProcessFilesErrorCodes.LocalizedError Or
           MyBase.ErrorCode = eProcessFilesErrorCodes.NoError Then
            Select Case mLocalErrorCode
                Case eResultsProcessorErrorCodes.NoError
                    strErrorMessage = ""
                Case eResultsProcessorErrorCodes.UnspecifiedError
                    strErrorMessage = "Unspecified localized error"
                Case eResultsProcessorErrorCodes.MissingMASICFiles
                    strErrorMessage = "Missing MASIC Files"
                Case eResultsProcessorErrorCodes.MissingMageFiles
                    strErrorMessage = "Missing Mage Extractor Files"
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
        ScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN
        SeparateByCollisionMode = False

        WarnMissingParameterFileSection = True

        mLocalErrorCode = eResultsProcessorErrorCodes.NoError

        mProcessedDatasets = New List(Of clsProcessedFileInfo)

    End Sub

    Private Function LoadParameterFileSettings(strParameterFilePath As String) As Boolean

        Const OPTIONS_SECTION = "MASICResultsMerger"

        Dim objSettingsFile As New XmlSettingsFileAccessor

        Try

            If strParameterFilePath Is Nothing OrElse strParameterFilePath.Length = 0 Then
                ' No parameter file specified; nothing to load
                Return True
            End If

            If Not File.Exists(strParameterFilePath) Then
                ' See if strParameterFilePath points to a file in the same directory as the application
                strParameterFilePath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), Path.GetFileName(strParameterFilePath))
                If Not File.Exists(strParameterFilePath) Then
                    MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.ParameterFileNotFound)
                    Return False
                End If
            End If

            If objSettingsFile.LoadSettings(strParameterFilePath) Then
                If Not objSettingsFile.SectionPresent(OPTIONS_SECTION) Then
                    ShowErrorMessage("The node '<section name=""" & OPTIONS_SECTION & """> was not found in the parameter file: " & strParameterFilePath)
                    MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.InvalidParameterFile)
                    Return False
                Else
                    ScanNumberColumn = objSettingsFile.GetParam(OPTIONS_SECTION, "ScanNumberColumn", DEFAULT_SCAN_NUMBER_COLUMN)
                    SeparateByCollisionMode = objSettingsFile.GetParam(OPTIONS_SECTION, "SeparateByCollisionMode", False)
                End If
            End If

        Catch ex As Exception
            HandleException("Error in LoadParameterFileSettings", ex)
            Return False
        End Try

        Return True

    End Function

    Protected Function MergePeptideHitAndMASICFiles(strInputFilePath As String,
      strOutputFolderPath As String,
      dctScanStats As Dictionary(Of Integer, udtScanStatsType),
      dctSICStats As Dictionary(Of Integer, udtSICStatsType),
      strReporterIonHeaders As String) As Boolean


        Dim swOutfile() As StreamWriter
        Dim intLinesWritten() As Integer

        Dim intOutputFileCount As Integer

        ' The Key is the collision mode and the value is the path
        Dim strOutputFilePaths() As KeyValuePair(Of String, String)

        Dim baseFileName = String.Empty

        Dim strLineIn As String

        Dim strSplitLine() As String

        Dim strHeaderLine As String
        Dim strAddonColumns As String
        Dim strCollisionModeCurrentScan As String

        Dim strBlankAdditionalColumns As String = String.Empty
        Dim strBlankAdditionalSICColumns As String = String.Empty
        Dim strBlankAdditionalReporterIonColumns As String = String.Empty

        Dim intOutFileIndex As Integer
        Dim intEmptyOutFileCount As Integer

        Dim intLinesRead As Integer
        Dim intScanNumber As Integer

        Dim dctCollisionModeFileMap As Dictionary(Of String, Integer)

        Dim blnWriteReporterIonStats = False
        Dim blnSuccess As Boolean

        Try

            baseFileName = Path.GetFileNameWithoutExtension(strInputFilePath)

            If String.IsNullOrWhiteSpace(strReporterIonHeaders) Then strReporterIonHeaders = String.Empty

            ' Define the output file path
            If strOutputFolderPath Is Nothing Then
                strOutputFolderPath = String.Empty
            End If

            dctCollisionModeFileMap = New Dictionary(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)

            If SeparateByCollisionMode Then
                ' Construct a list of the different collision modes in dctScanStats

                intOutputFileCount = 0

                For Each udtItem In dctScanStats.Values
                    If Not dctCollisionModeFileMap.ContainsKey(udtItem.CollisionMode) Then
                        ' Store this collision mode in htCollisionModes; the value stored will be the index in strCollisionModes()
                        dctCollisionModeFileMap.Add(udtItem.CollisionMode, intOutputFileCount)
                        intOutputFileCount += 1
                    End If
                Next

                ReDim strOutputFilePaths(intOutputFileCount - 1)

                If dctCollisionModeFileMap.Count = 0 Then
                    strOutputFilePaths(0) = New KeyValuePair(Of String, String)("na", Path.Combine(strOutputFolderPath, baseFileName & "_na" & RESULTS_SUFFIX))
                Else
                    For Each oItem In dctCollisionModeFileMap
                        Dim strCollisionMode As String = oItem.Key
                        If String.IsNullOrWhiteSpace(strCollisionMode) Then strCollisionMode = "na"
                        strOutputFilePaths(oItem.Value) = New KeyValuePair(Of String, String)(strCollisionMode, Path.Combine(strOutputFolderPath, baseFileName & "_" & strCollisionMode & RESULTS_SUFFIX))
                    Next
                End If
            Else
                intOutputFileCount = 1
                ReDim strOutputFilePaths(0)
                strOutputFilePaths(0) = New KeyValuePair(Of String, String)("", Path.Combine(strOutputFolderPath, baseFileName & RESULTS_SUFFIX))
            End If

            ReDim swOutfile(intOutputFileCount - 1)
            ReDim intLinesWritten(intOutputFileCount - 1)

            ' Open the output file(s)
            For intIndex = 0 To intOutputFileCount - 1
                swOutfile(intIndex) = New StreamWriter(New FileStream(strOutputFilePaths(intIndex).Value, FileMode.Create, FileAccess.Write, FileShare.Read))
            Next

        Catch ex As Exception
            HandleException("Error creating the merged output file", ex)
            Return False
        End Try


        Try
            MyBase.mProgressStepDescription = "Parsing " & Path.GetFileName(strInputFilePath) & " and writing " & Path.GetFileName(strOutputFilePaths(0).Value)
            ShowMessage(MyBase.mProgressStepDescription)

            If ScanNumberColumn < 1 Then
                ' Assume the scan number is in the second column
                ScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN
            End If


            ' Read from srInFile and write out to the file(s) in swOutFile
            Using srInFile = New StreamReader(New FileStream(strInputFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))

                intLinesRead = 0
                Do While srInFile.Peek > -1
                    strLineIn = srInFile.ReadLine
                    strCollisionModeCurrentScan = String.Empty

                    If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                        intLinesRead += 1
                        strSplitLine = strLineIn.Split(ControlChars.Tab)
                        strAddonColumns = String.Empty

                        If intLinesRead = 1 Then
                            ' Write out an updated header line
                            If strSplitLine.Length >= ScanNumberColumn AndAlso Integer.TryParse(strSplitLine(ScanNumberColumn - 1), intScanNumber) Then
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

                            Dim lstAdditionalHeaders As List(Of String)
                            lstAdditionalHeaders = GetAdditionalMASICHeaders()

                            ' Populate strBlankAdditionalColumns with tab characters based on the number of items in lstAdditionalHeaders
                            strBlankAdditionalColumns = New String(ControlChars.Tab, lstAdditionalHeaders.Count - 1)

                            strBlankAdditionalSICColumns = New String(ControlChars.Tab, SIC_STAT_COLUMN_COUNT_TO_ADD)

                            ' Initialize strBlankAdditionalReporterIonColumns
                            If strReporterIonHeaders.Length > 0 Then
                                strBlankAdditionalReporterIonColumns = New String(ControlChars.Tab, strReporterIonHeaders.Split(ControlChars.Tab).ToList().Count - 1)
                            End If

                            ' Initialize the AddOn Columns
                            strAddonColumns = FlattenList(lstAdditionalHeaders)

                            If strReporterIonHeaders.Length > 0 Then
                                ' Append the reporter ion stats columns
                                strAddonColumns &= ControlChars.Tab & strReporterIonHeaders
                                blnWriteReporterIonStats = True
                            End If

                            ' Write out the headers
                            For intIndex = 0 To intOutputFileCount - 1
                                swOutfile(intIndex).WriteLine(strHeaderLine & ControlChars.Tab & strAddonColumns)
                            Next

                        End If

                        If strSplitLine.Length >= ScanNumberColumn AndAlso Integer.TryParse(strSplitLine(ScanNumberColumn - 1), intScanNumber) Then
                            ' Look for intScanNumber in dctScanStats
                            Dim udtScanStats = New udtScanStatsType
                            If Not dctScanStats.TryGetValue(intScanNumber, udtScanStats) Then
                                ' Match not found; use the blank columns in strBlankAdditionalColumns
                                strAddonColumns = String.Copy(strBlankAdditionalColumns)
                            Else
                                With udtScanStats
                                    strAddonColumns = .ElutionTime & ControlChars.Tab &
                                       .ScanType & ControlChars.Tab &
                                       .TotalIonIntensity & ControlChars.Tab &
                                       .BasePeakIntensity & ControlChars.Tab &
                                       .BasePeakMZ
                                End With

                                Dim udtSICStats = New udtSICStatsType
                                If Not dctSICStats.TryGetValue(intScanNumber, udtSICStats) Then
                                    ' Match not found; use the blank columns in strBlankAdditionalSICColumns
                                    strAddonColumns &= ControlChars.Tab & strBlankAdditionalSICColumns

                                    If blnWriteReporterIonStats Then
                                        strAddonColumns &= ControlChars.Tab &
                                          String.Empty & ControlChars.Tab &
                                          strBlankAdditionalReporterIonColumns
                                    End If
                                Else
                                    With udtSICStats
                                        strAddonColumns &= ControlChars.Tab &
                                         .OptimalScanNumber & ControlChars.Tab &
                                         .PeakMaxIntensity & ControlChars.Tab &
                                         .PeakSignalToNoiseRatio & ControlChars.Tab &
                                         .FWHMInScans & ControlChars.Tab &
                                         .PeakArea & ControlChars.Tab &
                                         .ParentIonIntensity & ControlChars.Tab &
                                         .ParentIonMZ & ControlChars.Tab &
                                         .StatMomentsArea
                                    End With
                                End If

                                If blnWriteReporterIonStats Then
                                    With udtScanStats
                                        If String.IsNullOrWhiteSpace(.CollisionMode) Then
                                            ' Collision mode is not defined; append blank columns
                                            strAddonColumns &= ControlChars.Tab &
                                             String.Empty & ControlChars.Tab &
                                             strBlankAdditionalReporterIonColumns
                                        Else
                                            ' Collision mode is defined
                                            strAddonColumns &= ControlChars.Tab &
                                             .CollisionMode & ControlChars.Tab &
                                             .ReporterIonData

                                            strCollisionModeCurrentScan = String.Copy(.CollisionMode)
                                        End If
                                    End With


                                End If
                            End If

                            intOutFileIndex = 0
                            If SeparateByCollisionMode AndAlso intOutputFileCount > 1 Then
                                If Not strCollisionModeCurrentScan Is Nothing Then
                                    ' Determine the correct output file
                                    If Not dctCollisionModeFileMap.TryGetValue(strCollisionModeCurrentScan, intOutFileIndex) Then
                                        intOutFileIndex = 0
                                    End If
                                End If
                            End If

                            swOutfile(intOutFileIndex).WriteLine(strLineIn & ControlChars.Tab & strAddonColumns)
                            intLinesWritten(intOutFileIndex) += 1

                        End If

                    End If
                Loop

            End Using

            ' Close the output files
            If Not swOutfile Is Nothing Then
                For intIndex = 0 To intOutputFileCount - 1
                    If Not swOutfile(intIndex) Is Nothing Then
                        swOutfile(intIndex).Close()
                    End If
                Next
            End If

            ' See if any of the files had no data written to them
            ' If there are, then delete the empty output file
            ' However, retain at least one output file
            intEmptyOutFileCount = 0
            For intIndex = 0 To intOutputFileCount - 1
                If intLinesWritten(intIndex) = 0 Then
                    intEmptyOutFileCount += 1
                End If
            Next

            Dim outputPathEntry = New clsProcessedFileInfo(baseFileName)

            If intEmptyOutFileCount = 0 Then
                For Each item In strOutputFilePaths.ToList()
                    outputPathEntry.AddOutputFile(item.Key, item.Value)
                Next
            Else
                If intEmptyOutFileCount = intOutputFileCount Then
                    ' All output files are empty
                    ' Pretend the first output file actually contains data
                    intLinesWritten(0) = 1
                End If

                For intIndex = 0 To intOutputFileCount - 1
                    ' Wait 250 msec before continuing
                    Thread.Sleep(250)

                    If intLinesWritten(intIndex) = 0 Then
                        Try
                            ShowMessage("Deleting empty output file: " & ControlChars.NewLine & " --> " & Path.GetFileName(strOutputFilePaths(intIndex).Value))
                            File.Delete(strOutputFilePaths(intIndex).Value)
                        Catch ex As Exception
                            ' Ignore errors here
                        End Try
                    Else
                        outputPathEntry.AddOutputFile(strOutputFilePaths(intIndex).Key, strOutputFilePaths(intIndex).Value)
                    End If
                Next
            End If

            If outputPathEntry.OutputFiles.Count > 0 Then
                mProcessedDatasets.Add(outputPathEntry)
            End If

            blnSuccess = True

        Catch ex As Exception
            HandleException("Error in MergePeptideHitAndMASICFiles", ex)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Public Function MergeProcessedDatasets() As Boolean

        Try
            If mProcessedDatasets.Count = 1 Then
                ' Nothing to merge
                ShowMessage("Only one dataset has been processed by the MASICResultsMerger; nothing to merge")
                Return True
            End If

            ' Determine the base filename and collision modes used
            Dim baseFileName = String.Empty

            Dim collisionModes = New SortedSet(Of String)
            Dim datasetNameIdMap = New Dictionary(Of String, Integer)

            For Each processedDataset In mProcessedDatasets

                For Each processedFile In processedDataset.OutputFiles
                    If Not collisionModes.Contains(processedFile.Key) Then
                        collisionModes.Add(processedFile.Key)
                    End If
                Next

                If Not datasetNameIdMap.ContainsKey(processedDataset.BaseName) Then
                    datasetNameIdMap.Add(processedDataset.BaseName, datasetNameIdMap.Count + 1)
                End If

                ' Find the characters common to all of the processed datasets
                Dim candidateName = processedDataset.BaseName

                If String.IsNullOrEmpty(baseFileName) Then
                    baseFileName = String.Copy(candidateName)
                Else
                    Dim charsInCommon = 0

                    For intIndex = 0 To baseFileName.Length - 1
                        If intIndex >= candidateName.Length Then
                            Exit For
                        End If

                        If candidateName.ToLower().Chars(intIndex) <> baseFileName.ToLower().Chars(intIndex) Then
                            Exit For
                        End If

                        charsInCommon += 1
                    Next

                    If charsInCommon > 1 Then
                        baseFileName = baseFileName.Substring(0, charsInCommon)

                        ' Possibly backtrack to the previous underscore
                        Dim lastUnderscore = baseFileName.LastIndexOf("_", StringComparison.Ordinal)
                        If lastUnderscore >= 4 Then
                            baseFileName = baseFileName.Substring(0, lastUnderscore)
                        End If
                    End If

                End If
            Next

            If collisionModes.Count = 0 Then
                ShowErrorMessage("None of the processed datasets had any output files")
            End If

            baseFileName = "MergedData_" & baseFileName

            ' Open the output files
            Dim outputFileHandles = New Dictionary(Of String, StreamWriter)
            Dim outputFileHeaderWritten = New Dictionary(Of String, Boolean)

            For Each collisionMode In collisionModes
                Dim outputFileName As String

                If collisionMode = clsProcessedFileInfo.COLLISION_MODE_NOT_DEFINED Then
                    outputFileName = baseFileName & RESULTS_SUFFIX
                Else
                    outputFileName = baseFileName & "_" & collisionMode & RESULTS_SUFFIX
                End If

                outputFileHandles.Add(collisionMode, New StreamWriter(New FileStream(Path.Combine(mOutputFolderPath, outputFileName), FileMode.Create, FileAccess.Write)))
                outputFileHeaderWritten.Add(collisionMode, False)
            Next

            ' Create the DatasetMap file
            Using swOutFile = New StreamWriter(New FileStream(Path.Combine(mOutputFolderPath, baseFileName & "_DatasetMap.txt"), FileMode.Create, FileAccess.Write))

                swOutFile.WriteLine("DatasetID" & ControlChars.Tab & "DatasetName")

                For Each datasetMapping In datasetNameIdMap
                    swOutFile.WriteLine(datasetMapping.Value & ControlChars.Tab & datasetMapping.Key)
                Next

            End Using

            ' Merge the files
            For Each processedDataset In mProcessedDatasets
                Dim datasetId As Integer = datasetNameIdMap(processedDataset.BaseName)

                For Each sourcefile In processedDataset.OutputFiles

                    Dim collisionMode = sourcefile.Key

                    Dim swOutFile As StreamWriter = Nothing
                    If Not outputFileHandles.TryGetValue(collisionMode, swOutFile) Then
                        Console.WriteLine("Warning: unrecognized collison mode; skipping " & sourcefile.Value)
                        Continue For
                    End If

                    If Not File.Exists(sourcefile.Value) Then
                        Console.WriteLine("Warning: input file not found; skipping " & sourcefile.Value)
                        Continue For
                    End If

                    Using srSourceFile = New StreamReader(New FileStream(sourcefile.Value, FileMode.Open, FileAccess.Read, FileShare.Read))
                        Dim linesRead = 0
                        While srSourceFile.Peek > -1

                            Dim lineIn = srSourceFile.ReadLine()
                            linesRead += 1

                            If linesRead = 1 Then
                                If outputFileHeaderWritten(collisionMode) Then
                                    ' skip this line
                                    Continue While
                                End If

                                swOutFile.WriteLine("DatasetID" & ControlChars.Tab & lineIn)
                                outputFileHeaderWritten(collisionMode) = True

                            Else
                                swOutFile.WriteLine(datasetId & ControlChars.Tab & lineIn)
                            End If

                        End While
                    End Using

                Next
            Next

            For Each outputFile In outputFileHandles
                outputFile.Value.Close()
            Next

        Catch ex As Exception
            HandleException("Error in MergeProcessedDatasets", ex)
            Return False
        End Try

        Return True

    End Function

    ' Main processing function
    Public Overloads Overrides Function ProcessFile(strInputFilePath As String, strOutputFolderPath As String, strParameterFilePath As String, blnResetErrorCode As Boolean) As Boolean
        ' Returns True if success, False if failure

        Dim blnSuccess As Boolean
        Dim strMASICResultsFolder As String

        If blnResetErrorCode Then
            SetLocalErrorCode(eResultsProcessorErrorCodes.NoError)
        End If

        If Not LoadParameterFileSettings(strParameterFilePath) Then
            ShowErrorMessage("Parameter file load error: " & strParameterFilePath)

            If MyBase.ErrorCode = eProcessFilesErrorCodes.NoError Then
                MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.InvalidParameterFile)
            End If
            Return False
        End If

        If strInputFilePath Is Nothing OrElse strInputFilePath.Length = 0 Then
            ShowMessage("Input file name is empty")
            MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.InvalidInputFilePath)
            Return False
        End If

        ' Note that CleanupFilePaths() will update mOutputFolderPath, which is used by LogMessage()
        If Not CleanupFilePaths(strInputFilePath, strOutputFolderPath) Then
            MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.FilePathError)
            Return False
        End If

        Dim fiInputFile As FileInfo
        fiInputFile = New FileInfo(strInputFilePath)

        If String.IsNullOrWhiteSpace(mMASICResultsFolderPath) Then
            strMASICResultsFolder = fiInputFile.DirectoryName
        Else
            strMASICResultsFolder = String.Copy(mMASICResultsFolderPath)
        End If

        If MageResults Then
            blnSuccess = ProcessMageExtractorFile(fiInputFile, strMASICResultsFolder)
        Else
            blnSuccess = ProcessSingleJobFile(fiInputFile, strMASICResultsFolder)
        End If
        Return blnSuccess

    End Function

    Protected Function ProcessMageExtractorFile(fiInputFile As FileInfo, strMASICResultsFolder As String) As Boolean

        Dim udtMASICFileNames = New udtMASICFileNamesType

        Dim dctScanStats = New Dictionary(Of Integer, udtScanStatsType)
        Dim dctSICStats = New Dictionary(Of Integer, udtSICStatsType)

        Dim fiMetadataFile As FileInfo
        Dim strMetadataFile As String

        Dim lstColumns As List(Of String)
        Dim intJobColumnIndex As Integer

        Dim strHeaderLine As String
        Dim blnHeaderLineWritten = False
        Dim blnWriteReporterIonStats = False

        Dim strAddonColumns As String

        Dim strBlankAdditionalColumns As String
        Dim strBlankAdditionalSICColumns As String
        Dim strBlankAdditionalReporterIonColumns As String = String.Empty

        Dim strReporterIonHeaders As String = String.Empty

        Dim intJobsSuccessfullyMerged = 0

        Try

            ' Read the Mage Metadata file
            strMetadataFile = Path.Combine(fiInputFile.DirectoryName, Path.GetFileNameWithoutExtension(fiInputFile.Name) & "_metadata.txt")
            fiMetadataFile = New FileInfo(strMetadataFile)
            If Not fiMetadataFile.Exists Then
                ShowErrorMessage("Error: Mage Metadata File not found: " & fiMetadataFile.FullName)
                SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMageFiles)
                Return False
            End If

            ' Keys in this dictionary are the job, values are the DatasetID and DatasetName
            Dim dctJobToDatasetMap As Dictionary(Of Integer, udtDatasetInfoType)

            dctJobToDatasetMap = ReadMageMetadataFile(fiMetadataFile.FullName)
            If dctJobToDatasetMap Is Nothing OrElse dctJobToDatasetMap.Count = 0 Then
                ShowErrorMessage("Error: ReadMageMetadataFile returned an empty job mapping")
                Return False
            End If

            ' Open the Mage Extractor data file so that we can validate and cache the header row
            Using srInFile = New StreamReader(New FileStream(fiInputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.Read))
                strHeaderLine = srInFile.ReadLine()
                lstColumns = strHeaderLine.Split(ControlChars.Tab).ToList()
                intJobColumnIndex = lstColumns.IndexOf("Job")
                If intJobColumnIndex < 0 Then
                    ShowErrorMessage("Input file is not a valid Mage Extractor results file; it must contain a ""Job"" column: " & fiInputFile.FullName)
                    Return False
                End If
            End Using


            Dim lstAdditionalHeaders As List(Of String)
            lstAdditionalHeaders = GetAdditionalMASICHeaders()

            ' Populate strBlankAdditionalColumns with tab characters based on the number of items in lstAdditionalHeaders
            strBlankAdditionalColumns = New String(ControlChars.Tab, lstAdditionalHeaders.Count - 1)

            strBlankAdditionalSICColumns = New String(ControlChars.Tab, SIC_STAT_COLUMN_COUNT_TO_ADD)

            Dim strOutputFilePath As String
            strOutputFilePath = Path.GetFileNameWithoutExtension(fiInputFile.Name) & RESULTS_SUFFIX
            strOutputFilePath = Path.Combine(mOutputFolderPath, strOutputFilePath)

            ' Initialize the output file
            Using swOutFile = New StreamWriter(New FileStream(strOutputFilePath, FileMode.Create, FileAccess.Write, FileShare.Read))


                ' Open the Mage Extractor data file and read the data for each job
                mPHRPReader = New clsPHRPReader(fiInputFile.FullName, clsPHRPReader.ePeptideHitResultType.Unknown, False, False, False)
                mPHRPReader.EchoMessagesToConsole = False
                mPHRPReader.SkipDuplicatePSMs = False

                If Not mPHRPReader.CanRead Then
                    ShowErrorMessage("Aborting since PHRPReader is not ready: " & mPHRPReader.ErrorMessage)
                    Return False
                End If

                Dim intLastJob As Integer = -1
                Dim intJob As Integer = -1
                Dim blnMASICDataLoaded = False

                Do While mPHRPReader.MoveNext()

                    Dim oPSM As clsPSM = mPHRPReader.CurrentPSM
                    ' Parse out the job from the current line
                    lstColumns = oPSM.DataLineText.Split(ControlChars.Tab).ToList()

                    If Not Integer.TryParse(lstColumns(intJobColumnIndex), intJob) Then
                        ShowMessage("Warning: Job column does not contain a job number; skipping this entry: " & oPSM.DataLineText)
                        Continue Do
                    End If

                    If intJob <> intLastJob Then

                        ' New job; read and cache the MASIC data
                        blnMASICDataLoaded = False

                        Dim udtDatasetInfo = New udtDatasetInfoType

                        If Not dctJobToDatasetMap.TryGetValue(intJob, udtDatasetInfo) Then
                            ShowErrorMessage("Error: Job " & intJob & " was not defined in the Metadata file; unable to determine the dataset")
                        Else

                            ' Look for the corresponding MASIC files in the input folder
                            Dim blnSuccess As Boolean

                            udtMASICFileNames.Initialize()
                            blnSuccess = FindMASICFiles(strMASICResultsFolder, udtDatasetInfo, udtMASICFileNames)

                            If Not blnSuccess Then
                                ShowMessage("  Error: Unable to find the MASIC data files for dataset " & udtDatasetInfo.DatasetName & " in " & strMASICResultsFolder)
                                ShowMessage("         Job " & intJob & " will not have MASIC results")
                            Else
                                If udtMASICFileNames.SICStatsFileName.Length = 0 Then
                                    ShowMessage("  Error: the SIC stats file was not found for dataset " & udtDatasetInfo.DatasetName & " in " & strMASICResultsFolder)
                                    ShowMessage("         Job " & intJob & " will not have MASIC results")
                                    blnSuccess = False
                                ElseIf udtMASICFileNames.ScanStatsFileName.Length = 0 Then
                                    ShowMessage("  Error: the Scan stats file was not found for dataset " & udtDatasetInfo.DatasetName & " in " & strMASICResultsFolder)
                                    ShowMessage("         Job " & intJob & " will not have MASIC results")
                                    blnSuccess = False
                                End If
                            End If

                            If blnSuccess Then

                                ' Read and cache the MASIC data
                                dctScanStats = New Dictionary(Of Integer, udtScanStatsType)
                                dctSICStats = New Dictionary(Of Integer, udtSICStatsType)
                                strReporterIonHeaders = String.Empty

                                blnMASICDataLoaded = ReadMASICData(strMASICResultsFolder, udtMASICFileNames, dctScanStats, dctSICStats, strReporterIonHeaders)

                                If blnMASICDataLoaded Then
                                    intJobsSuccessfullyMerged += 1

                                    If intJobsSuccessfullyMerged = 1 Then

                                        ' Initialize strBlankAdditionalReporterIonColumns
                                        If strReporterIonHeaders.Length > 0 Then
                                            strBlankAdditionalReporterIonColumns = New String(ControlChars.Tab, strReporterIonHeaders.Split(ControlChars.Tab).ToList().Count - 1)
                                        End If

                                    End If
                                End If
                            End If

                        End If
                    End If

                    If blnMASICDataLoaded Then

                        If Not blnHeaderLineWritten Then

                            strAddonColumns = FlattenList(lstAdditionalHeaders)
                            If strReporterIonHeaders.Length > 0 Then
                                ' Append the reporter ion stats columns
                                strAddonColumns &= ControlChars.Tab & strReporterIonHeaders
                                blnWriteReporterIonStats = True
                            End If

                            swOutFile.WriteLine(strHeaderLine & ControlChars.Tab & strAddonColumns)

                            blnHeaderLineWritten = True
                        End If


                        ' Look for intScanNumber in dctScanStats
                        Dim udtScanStats = New udtScanStatsType
                        If Not dctScanStats.TryGetValue(oPSM.ScanNumber, udtScanStats) Then
                            ' Match not found; use the blank columns in strBlankAdditionalColumns
                            strAddonColumns = String.Copy(strBlankAdditionalColumns)
                        Else
                            With udtScanStats
                                strAddonColumns = .ElutionTime & ControlChars.Tab &
                                   .ScanType & ControlChars.Tab &
                                   .TotalIonIntensity & ControlChars.Tab &
                                   .BasePeakIntensity & ControlChars.Tab &
                                   .BasePeakMZ
                            End With

                            Dim udtSICStats = New udtSICStatsType
                            If Not dctSICStats.TryGetValue(oPSM.ScanNumber, udtSICStats) Then
                                ' Match not found; use the blank columns in strBlankAdditionalSICColumns
                                strAddonColumns &= ControlChars.Tab & strBlankAdditionalSICColumns

                                If blnWriteReporterIonStats Then
                                    strAddonColumns &= ControlChars.Tab &
                                      String.Empty & ControlChars.Tab &
                                      strBlankAdditionalReporterIonColumns
                                End If
                            Else
                                With udtSICStats
                                    strAddonColumns &= ControlChars.Tab &
                                     .OptimalScanNumber & ControlChars.Tab &
                                     .PeakMaxIntensity & ControlChars.Tab &
                                     .PeakSignalToNoiseRatio & ControlChars.Tab &
                                     .FWHMInScans & ControlChars.Tab &
                                     .PeakArea & ControlChars.Tab &
                                     .ParentIonIntensity & ControlChars.Tab &
                                     .ParentIonMZ & ControlChars.Tab &
                                     .StatMomentsArea
                                End With
                            End If

                            If blnWriteReporterIonStats Then

                                With udtScanStats
                                    If String.IsNullOrWhiteSpace(.CollisionMode) Then
                                        ' Collision mode is not defined; append blank columns
                                        strAddonColumns &= ControlChars.Tab &
                                         String.Empty & ControlChars.Tab &
                                         strBlankAdditionalReporterIonColumns
                                    Else
                                        ' Collision mode is defined
                                        strAddonColumns &= ControlChars.Tab &
                                         .CollisionMode & ControlChars.Tab &
                                         .ReporterIonData

                                    End If
                                End With


                            End If
                        End If

                        swOutFile.WriteLine(oPSM.DataLineText & ControlChars.Tab & strAddonColumns)
                    Else
                        swOutFile.WriteLine(oPSM.DataLineText & ControlChars.Tab & strBlankAdditionalColumns)
                    End If

                    UpdateProgress(mPHRPReader.PercentComplete)
                    intLastJob = intJob
                Loop

            End Using

            If intJobsSuccessfullyMerged > 0 Then
                Console.WriteLine()
                ShowMessage("Merged MASIC results for " & intJobsSuccessfullyMerged & " jobs")
            End If

        Catch ex As Exception
            HandleException("Error in ProcessMageExtractorFile", ex)
            Return False
        End Try

        If intJobsSuccessfullyMerged > 0 Then
            Return True
        Else
            Return False
        End If


    End Function

    Protected Function ProcessSingleJobFile(fiInputFile As FileInfo, strMASICResultsFolder As String) As Boolean
        Dim udtMASICFileNames = New udtMASICFileNamesType

        Dim dctScanStats As Dictionary(Of Integer, udtScanStatsType)
        Dim dctSICStats As Dictionary(Of Integer, udtSICStatsType)


        Dim strReporterIonHeaders As String = String.Empty

        Dim blnSuccess As Boolean

        Try
            Dim udtDatasetInfo = New udtDatasetInfoType

            ' Note that FindMASICFiles will first try the full filename, and if it doesn't find a match,
            ' it will start removing text from the end of the filename by looking for underscores
            udtDatasetInfo.DatasetName = Path.GetFileNameWithoutExtension(fiInputFile.FullName)
            udtDatasetInfo.DatasetID = 0

            ' Look for the corresponding MASIC files in the input folder
            udtMASICFileNames.Initialize()
            blnSuccess = FindMASICFiles(strMASICResultsFolder, udtDatasetInfo, udtMASICFileNames)

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
            dctScanStats = New Dictionary(Of Integer, udtScanStatsType)
            dctSICStats = New Dictionary(Of Integer, udtSICStatsType)

            blnSuccess = ReadMASICData(strMASICResultsFolder, udtMASICFileNames, dctScanStats, dctSICStats, strReporterIonHeaders)

            If blnSuccess Then
                ' Merge the MASIC data with the input file
                blnSuccess = MergePeptideHitAndMASICFiles(fiInputFile.FullName, mOutputFolderPath,
                 dctScanStats,
                 dctSICStats,
                 strReporterIonHeaders)
            End If

            If blnSuccess Then
                ShowMessage(String.Empty, False)
            Else
                SetLocalErrorCode(eResultsProcessorErrorCodes.UnspecifiedError)
                ShowErrorMessage("Error")
            End If

        Catch ex As Exception
            HandleException("Error in ProcessSingleJobFile", ex)
        End Try

        Return blnSuccess
    End Function

    Protected Function ReadMASICData(strSourceFolder As String,
      udtMASICFileNames As udtMASICFileNamesType,
      ByRef dctScanStats As Dictionary(Of Integer, udtScanStatsType),
      ByRef dctSICStats As Dictionary(Of Integer, udtSICStatsType),
      ByRef strReporterIonHeaders As String) As Boolean

        Dim blnSuccess As Boolean

        Try

            blnSuccess = ReadScanStatsFile(strSourceFolder, udtMASICFileNames.ScanStatsFileName, dctScanStats)

            If blnSuccess Then
                blnSuccess = ReadSICStatsFile(strSourceFolder, udtMASICFileNames.SICStatsFileName, dctSICStats)
            End If

            If blnSuccess AndAlso udtMASICFileNames.ReporterIonsFileName.Length > 0 Then
                blnSuccess = ReadReporterIonStatsFile(strSourceFolder, udtMASICFileNames.ReporterIonsFileName, dctScanStats, strReporterIonHeaders)
            End If

        Catch ex As Exception
            HandleException("Error in ReadMASICData", ex)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Protected Function ReadScanStatsFile(strSourceFolder As String,
     strScanStatsFileName As String,
     ByRef dctScanStats As Dictionary(Of Integer, udtScanStatsType)) As Boolean

        Dim strLineIn As String
        Dim strSplitLine() As String

        Dim intLinesRead As Integer
        Dim intScanNumber As Integer

        Dim blnSuccess = False

        Try
            ' Initialize dctScanStats
            If dctScanStats Is Nothing Then
                dctScanStats = New Dictionary(Of Integer, udtScanStatsType)
            Else
                dctScanStats.Clear()
            End If

            MyBase.mProgressStepDescription = "  Reading: " & strScanStatsFileName
            ShowMessage(MyBase.mProgressStepDescription)

            Using srInFile = New StreamReader(New FileStream(Path.Combine(strSourceFolder, strScanStatsFileName), FileMode.Open, FileAccess.Read, FileShare.Read))

                intLinesRead = 0
                Do While srInFile.Peek > -1
                    strLineIn = srInFile.ReadLine

                    If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                        intLinesRead += 1
                        strSplitLine = strLineIn.Split(ControlChars.Tab)

                        If strSplitLine.Length >= eScanStatsColumns.BasePeakMZ + 1 AndAlso Integer.TryParse(strSplitLine(eScanStatsColumns.ScanNumber), intScanNumber) Then

                            Dim udtScanStats As udtScanStatsType
                            With udtScanStats
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

                            dctScanStats.Add(intScanNumber, udtScanStats)
                        End If
                    End If
                Loop

            End Using

            blnSuccess = True

        Catch ex As Exception
            HandleException("Error in ReadScanStatsFile", ex)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Protected Function ReadMageMetadataFile(strMetadataFilePath As String) As Dictionary(Of Integer, udtDatasetInfoType)

        Dim dctJobToDatasetMap = New Dictionary(Of Integer, udtDatasetInfoType)
        Dim strLineIn As String
        Dim lstData As List(Of String)
        Dim blnHeadersParsed As Boolean

        Dim intJobIndex As Integer = -1
        Dim intDatasetIndex As Integer = -1
        Dim intDatasetIDIndex As Integer = -1

        Try
            Using srInFile = New StreamReader(New FileStream(strMetadataFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))

                Do While srInFile.Peek > -1
                    strLineIn = srInFile.ReadLine()

                    If Not String.IsNullOrWhiteSpace(strLineIn) Then
                        lstData = strLineIn.Split(ControlChars.Tab).ToList()

                        If Not blnHeadersParsed Then
                            ' Look for the Job and Dataset columns
                            intJobIndex = lstData.IndexOf("Job")
                            intDatasetIndex = lstData.IndexOf("Dataset")
                            intDatasetIDIndex = lstData.IndexOf("Dataset_ID")

                            If intJobIndex < 0 Then
                                ShowErrorMessage("Job column not found in the metadata file: " & strMetadataFilePath)
                                Return Nothing
                            ElseIf intDatasetIndex < 0 Then
                                ShowErrorMessage("Dataset column not found in the metadata file: " & strMetadataFilePath)
                                Return Nothing
                            ElseIf intDatasetIDIndex < 0 Then
                                ShowErrorMessage("Dataset_ID column not found in the metadata file: " & strMetadataFilePath)
                                Return Nothing
                            End If

                            blnHeadersParsed = True
                            Continue Do
                        End If

                        If lstData.Count > intDatasetIndex Then
                            Dim intJobNumber As Integer
                            Dim intDatasetID As Integer

                            If Integer.TryParse(lstData(intJobIndex), intJobNumber) Then
                                If Integer.TryParse(lstData(intDatasetIDIndex), intDatasetID) Then
                                    Dim udtDatasetInfo = New udtDatasetInfoType
                                    udtDatasetInfo.DatasetID = intDatasetID
                                    udtDatasetInfo.DatasetName = lstData(intDatasetIndex)

                                    dctJobToDatasetMap.Add(intJobNumber, udtDatasetInfo)
                                Else
                                    ShowMessage("Warning: Dataest_ID number not numeric in metadata file, line " & strLineIn)
                                End If
                            Else
                                ShowMessage("Warning: Job number not numeric in metadata file, line " & strLineIn)
                            End If
                        End If
                    End If
                Loop
            End Using

        Catch ex As Exception
            HandleException("Error in ReadMageMetadataFile", ex)
            Return Nothing
        End Try

        Return dctJobToDatasetMap

    End Function

    Protected Function ReadSICStatsFile(strSourceFolder As String,
       strSICStatsFileName As String,
       ByRef dctSICStats As Dictionary(Of Integer, udtSICStatsType)) As Boolean

        Dim strLineIn As String
        Dim strSplitLine() As String

        Dim intLinesRead As Integer
        Dim intFragScanNumber As Integer

        Dim blnSuccess As Boolean

        Try
            ' Initialize dctSICStats
            If dctSICStats Is Nothing Then
                dctSICStats = New Dictionary(Of Integer, udtSICStatsType)
            Else
                dctSICStats.Clear()
            End If

            MyBase.mProgressStepDescription = "  Reading: " & strSICStatsFileName
            ShowMessage(MyBase.mProgressStepDescription)

            Using srInFile = New StreamReader(New FileStream(Path.Combine(strSourceFolder, strSICStatsFileName), FileMode.Open, FileAccess.Read, FileShare.Read))

                intLinesRead = 0
                Do While srInFile.Peek > -1
                    strLineIn = srInFile.ReadLine

                    If Not strLineIn Is Nothing AndAlso strLineIn.Length > 0 Then
                        intLinesRead += 1
                        strSplitLine = strLineIn.Split(ControlChars.Tab)

                        If strSplitLine.Length >= eSICStatsColumns.StatMomentsArea + 1 AndAlso Integer.TryParse(strSplitLine(eSICStatsColumns.FragScanNumber), intFragScanNumber) Then

                            Dim udtSICStats As udtSICStatsType
                            With udtSICStats
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

                            dctSICStats.Add(intFragScanNumber, udtSICStats)
                        End If
                    End If
                Loop

            End Using

            blnSuccess = True

        Catch ex As Exception
            HandleException("Error in ReadSICStatsFile", ex)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Protected Function ReadReporterIonStatsFile(strSourceFolder As String,
      strReporterIonStatsFileName As String,
      dctScanStats As Dictionary(Of Integer, udtScanStatsType),
      ByRef strReporterIonHeaders As String) As Boolean

        Dim strLineIn As String
        Dim strSplitLine() As String

        Dim intLinesRead As Integer
        Dim intScanNumber As Integer

        Dim intWarningCount = 0

        Dim blnSuccess As Boolean

        Try
            strReporterIonHeaders = String.Empty

            MyBase.mProgressStepDescription = "  Reading: " & strReporterIonStatsFileName
            ShowMessage(MyBase.mProgressStepDescription)

            Using srInFile = New StreamReader(New FileStream(Path.Combine(strSourceFolder, strReporterIonStatsFileName), FileMode.Open, FileAccess.Read, FileShare.Read))

                intLinesRead = 0
                Do While srInFile.Peek > -1
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
                            Dim udtScanStats = New udtScanStatsType

                            If Not dctScanStats.TryGetValue(intScanNumber, udtScanStats) Then
                                If intWarningCount < 10 Then
                                    ShowMessage("Warning: " & REPORTER_IONS_FILE_EXTENSION & " file refers to scan " & intScanNumber.ToString & ", but that scan was not in the _ScanStats.txt file")
                                ElseIf intWarningCount = 10 Then
                                    ShowMessage("Warning: " & REPORTER_IONS_FILE_EXTENSION & " file has 10 or more scan numbers that are not defined in the _ScanStats.txt file")
                                End If
                                intWarningCount += 1
                            Else

                                If udtScanStats.ScanNumber <> intScanNumber Then
                                    ' Scan number mismatch; this shouldn't happen
                                    ShowMessage("Error: Scan number mismatch in ReadReporterIonStatsFile: " & udtScanStats.ScanNumber.ToString & " vs. " & intScanNumber.ToString)
                                Else
                                    udtScanStats.CollisionMode = String.Copy(strSplitLine(eReporterIonStatsColumns.CollisionMode))
                                    udtScanStats.ReporterIonData = FlattenArray(strSplitLine, eReporterIonStatsColumns.ReporterIonIntensityMax)

                                    dctScanStats(intScanNumber) = udtScanStats
                                End If

                            End If

                        End If
                    End If
                Loop

            End Using

            blnSuccess = True

        Catch ex As Exception
            HandleException("Error in ReadSICStatsFile", ex)
            blnSuccess = False
        End Try

        Return blnSuccess

    End Function

    Private Sub SetLocalErrorCode(eNewErrorCode As eResultsProcessorErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(eNewErrorCode As eResultsProcessorErrorCodes, blnLeaveExistingErrorCodeUnchanged As Boolean)

        If blnLeaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> eResultsProcessorErrorCodes.NoError Then
            ' An error code is already defined; do not change it
        Else
            mLocalErrorCode = eNewErrorCode

            If eNewErrorCode = eResultsProcessorErrorCodes.NoError Then
                If MyBase.ErrorCode = eProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

    Private Sub mPHRPReader_ErrorEvent(strErrorMessage As String) Handles mPHRPReader.ErrorEvent
        ShowErrorMessage(strErrorMessage)
    End Sub

    Private Sub mPHRPReader_MessageEvent(strMessage As String) Handles mPHRPReader.MessageEvent
        ShowMessage(strMessage)
    End Sub

    Private Sub mPHRPReader_WarningEvent(strWarningMessage As String) Handles mPHRPReader.WarningEvent
        ShowMessage("Warning: " & strWarningMessage)
    End Sub

End Class
