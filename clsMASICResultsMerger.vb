Option Strict On
Imports System.IO
Imports System.Runtime.InteropServices
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
' Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/
' -------------------------------------------------------------------------------
'
' Licensed under the Apache License, Version 2.0; you may not use this file except
' in compliance with the License.  You may obtain a copy of the License at
' http://www.apache.org/licenses/LICENSE-2.0
'

Public Class clsMASICResultsMerger
    Inherits FileProcessor.ProcessFilesBase

    Public Sub New()
        MyBase.mFileDate = "January 8, 2018"
        InitializeLocalVariables()
    End Sub

#Region "Constants and Enums"

    Public Const SIC_STATS_FILE_EXTENSION As String = "_SICStats.txt"
    Public Const SCAN_STATS_FILE_EXTENSION As String = "_ScanStats.txt"
    Public Const REPORTER_IONS_FILE_EXTENSION As String = "_ReporterIons.txt"

    Public Const RESULTS_SUFFIX As String = "_PlusSICStats.txt"
    Public Const DEFAULT_SCAN_NUMBER_COLUMN As Integer = 2

    Private Const SIC_STAT_COLUMN_COUNT_TO_ADD As Integer = 7

    ''' <summary>
    ''' Error codes specialized for this class
    ''' </summary>
    Public Enum eResultsProcessorErrorCodes As Integer
        NoError = 0
        MissingMASICFiles = 1
        MissingMageFiles = 2
        UnspecifiedError = -1
    End Enum

    Private Enum eScanStatsColumns
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

    Private Enum eSICStatsColumns
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

    Private Enum eReporterIonStatsColumns
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

    Private Structure udtDatasetInfoType
        Public DatasetName As String
        Public DatasetID As Integer
    End Structure

    Private Structure udtMASICFileNamesType
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

#End Region

#Region "Classwide Variables"

    Private mLocalErrorCode As eResultsProcessorErrorCodes

    Private mMASICResultsFolderPath As String = String.Empty

    Private mPHRPReader As clsPHRPReader

    Private mProcessedDatasets As List(Of clsProcessedFileInfo)

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

    ''' <summary>
    ''' When true, a separate output file will be created for each collision mode type; this is only possible if a _ReporterIons.txt file exists
    ''' </summary>
    ''' <returns></returns>
    Public Property SeparateByCollisionMode As Boolean

    ''' <summary>
    ''' For the input file, defines which column tracks scan number; the first column is column 1 (not zero)
    ''' </summary>
    ''' <returns></returns>
    Public Property ScanNumberColumn As Integer

#End Region

    Private Function FindMASICFiles(masicResultsFolder As String, udtDatasetInfo As udtDatasetInfoType, ByRef udtMASICFileNames As udtMASICFileNamesType) As Boolean

        Dim success = False
        Dim datasetName As String
        Dim triedDatasetID = False

        Dim candidateFilePath As String

        Dim charIndex As Integer

        Try
            Console.WriteLine()
            ShowMessage("Looking for MASIC data files that correspond to " & udtDatasetInfo.DatasetName)

            datasetName = String.Copy(udtDatasetInfo.DatasetName)

            ' Use a Do loop to try various possible dataset names
            Do
                candidateFilePath = Path.Combine(masicResultsFolder, datasetName & SIC_STATS_FILE_EXTENSION)

                If File.Exists(candidateFilePath) Then
                    ' SICStats file was found
                    ' Update udtMASICFileNames, then look for the other files

                    udtMASICFileNames.DatasetName = datasetName
                    udtMASICFileNames.SICStatsFileName = Path.GetFileName(candidateFilePath)

                    candidateFilePath = Path.Combine(masicResultsFolder, datasetName & SCAN_STATS_FILE_EXTENSION)
                    If File.Exists(candidateFilePath) Then
                        udtMASICFileNames.ScanStatsFileName = Path.GetFileName(candidateFilePath)
                    End If

                    candidateFilePath = Path.Combine(masicResultsFolder, datasetName & REPORTER_IONS_FILE_EXTENSION)
                    If File.Exists(candidateFilePath) Then
                        udtMASICFileNames.ReporterIonsFileName = Path.GetFileName(candidateFilePath)
                    End If

                    success = True

                Else
                    ' Find the last underscore in datasetName, then remove it and any text after it
                    charIndex = datasetName.LastIndexOf("_"c)

                    If charIndex > 0 Then
                        datasetName = datasetName.Substring(0, charIndex)
                    Else
                        If Not triedDatasetID AndAlso udtDatasetInfo.DatasetID > 0 Then
                            datasetName = udtDatasetInfo.DatasetID & "_" & udtDatasetInfo.DatasetName
                            triedDatasetID = True
                        Else
                            ' No more underscores; we're unable to determine the dataset name
                            Exit Do
                        End If

                    End If
                End If

            Loop While Not success

            Return success

        Catch ex As Exception
            HandleException("Error in FindMASICFiles", ex)
            Return False
        End Try

    End Function

    Private Sub FindScanNumColumn(inputFile As FileSystemInfo, splitLine As IList(Of String))

        ' Check for a column named "ScanNum" or "ScanNumber" or "Scan Num" or "Scan Number"
        ' If found, override ScanNumberColumn
        For colIndex = 0 To splitLine.Count - 1
            Select Case splitLine(colIndex).ToLower()
                Case "scan", "scannum", "scan num", "scannumber", "scan number", "scan#", "scan #"
                    If ScanNumberColumn <> colIndex + 1 Then
                        ScanNumberColumn = colIndex + 1
                        ShowMessage(
                            String.Format("Note: Reading scan numbers from column {0} ({1}) in file {2}",
                                          ScanNumberColumn, splitLine(colIndex), inputFile.Name))
                    End If
            End Select
        Next
    End Sub

    Private Function GetAdditionalMASICHeaders() As List(Of String)

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

    Private Function FlattenList(lstData As IReadOnlyList(Of String)) As String

        Dim sbFlattened = New StringBuilder

        For index = 0 To lstData.Count - 1
            If index > 0 Then
                sbFlattened.Append(ControlChars.Tab)
            End If
            sbFlattened.Append(lstData(index))
        Next

        Return sbFlattened.ToString()

    End Function

    Private Function FlattenArray(splitLine As IList(Of String), indexStart As Integer) As String
        Dim text As String = String.Empty
        Dim column As String

        Dim index As Integer

        If Not splitLine Is Nothing AndAlso splitLine.Count > 0 Then
            For index = indexStart To splitLine.Count - 1

                If splitLine(index) Is Nothing Then
                    column = String.Empty
                Else
                    column = String.Copy(splitLine(index))
                End If

                If index > indexStart Then
                    text &= ControlChars.Tab & column
                Else
                    text = String.Copy(column)
                End If
            Next
        End If

        Return text

    End Function

    ''' <summary>
    ''' Get the current error state, if any
    ''' </summary>
    ''' <returns>Returns an empty string if no error</returns>
    Public Overrides Function GetErrorMessage() As String

        Dim errorMessage As String

        If MyBase.ErrorCode = eProcessFilesErrorCodes.LocalizedError Or
           MyBase.ErrorCode = eProcessFilesErrorCodes.NoError Then
            Select Case mLocalErrorCode
                Case eResultsProcessorErrorCodes.NoError
                    errorMessage = ""
                Case eResultsProcessorErrorCodes.UnspecifiedError
                    errorMessage = "Unspecified localized error"
                Case eResultsProcessorErrorCodes.MissingMASICFiles
                    errorMessage = "Missing MASIC Files"
                Case eResultsProcessorErrorCodes.MissingMageFiles
                    errorMessage = "Missing Mage Extractor Files"
                Case Else
                    ' This shouldn't happen
                    errorMessage = "Unknown error state"
            End Select
        Else
            errorMessage = MyBase.GetBaseClassErrorMessage()
        End If

        Return errorMessage
    End Function

    Private Sub InitializeLocalVariables()

        mMASICResultsFolderPath = String.Empty
        ScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN
        SeparateByCollisionMode = False

        mLocalErrorCode = eResultsProcessorErrorCodes.NoError

        mProcessedDatasets = New List(Of clsProcessedFileInfo)

    End Sub

    Private Function MergePeptideHitAndMASICFiles(
      inputFile As FileSystemInfo,
      outputFolderPath As String,
      dctScanStats As Dictionary(Of Integer, clsScanStatsData),
      dctSICStats As IReadOnlyDictionary(Of Integer, clsSICStatsData),
      reporterIonHeaders As String) As Boolean

        Dim swOutfile() As StreamWriter
        Dim linesWritten() As Integer

        Dim outputFileCount As Integer

        ' The Key is the collision mode and the value is the path
        Dim outputFilePaths() As KeyValuePair(Of String, String)

        Dim baseFileName As String

        Dim blankAdditionalColumns As String = String.Empty
        Dim blankAdditionalSICColumns As String = String.Empty
        Dim blankAdditionalReporterIonColumns As String = String.Empty

        Dim dctCollisionModeFileMap As Dictionary(Of String, Integer) = Nothing

        Try
            If Not inputFile.Exists Then
                ShowErrorMessage("File not found: " & inputFile.FullName)
                Return False
            End If

            baseFileName = Path.GetFileNameWithoutExtension(inputFile.Name)

            If String.IsNullOrWhiteSpace(reporterIonHeaders) Then reporterIonHeaders = String.Empty

            ' Define the output file path
            If outputFolderPath Is Nothing Then
                outputFolderPath = String.Empty
            End If

            If SeparateByCollisionMode Then
                outputFilePaths = SummarizeCollisionModes(
                    inputFile,
                    baseFileName,
                    outputFolderPath,
                    dctScanStats,
                    dctCollisionModeFileMap)

                outputFileCount = outputFilePaths.Length

                If outputFileCount < 1 Then
                    Return False
                End If
            Else
                dctCollisionModeFileMap = New Dictionary(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)

                outputFileCount = 1
                ReDim outputFilePaths(0)
                outputFilePaths(0) = New KeyValuePair(Of String, String)("", Path.Combine(outputFolderPath, baseFileName & RESULTS_SUFFIX))
            End If

            ReDim swOutfile(outputFileCount - 1)
            ReDim linesWritten(outputFileCount - 1)

            ' Open the output file(s)
            For index = 0 To outputFileCount - 1
                swOutfile(index) = New StreamWriter(New FileStream(outputFilePaths(index).Value, FileMode.Create, FileAccess.Write, FileShare.Read))
            Next

        Catch ex As Exception
            HandleException("Error creating the merged output file", ex)
            Return False
        End Try

        Try
            Console.WriteLine()
            ShowMessage("Parsing " & inputFile.Name & " and writing " & Path.GetFileName(outputFilePaths(0).Value))

            If ScanNumberColumn < 1 Then
                ' Assume the scan number is in the second column
                ScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN
            End If

            ' Read from srInFile and write out to the file(s) in swOutFile
            Using srInFile = New StreamReader(New FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.Read))

                Dim linesRead = 0
                Dim writeReporterIonStats = False

                While Not srInFile.EndOfStream
                    Dim lineIn = srInFile.ReadLine()
                    Dim collisionModeCurrentScan = String.Empty

                    If Not lineIn Is Nothing AndAlso lineIn.Length > 0 Then
                        linesRead += 1
                        Dim splitLine = lineIn.Split(ControlChars.Tab)

                        If linesRead = 1 Then

                            Dim headerLine As String

                            Dim integerValue As Integer

                            ' Write out an updated header line
                            If splitLine.Length >= ScanNumberColumn AndAlso Integer.TryParse(splitLine(ScanNumberColumn - 1), integerValue) Then
                                ' The input file doesn't have a header line; we will add one, using generic column names for the data in the input file

                                headerLine = String.Empty
                                For index = 0 To splitLine.Length - 1
                                    If index = 0 Then
                                        headerLine = "Column" & index.ToString("00")
                                    Else
                                        headerLine &= ControlChars.Tab & "Column" & index.ToString("00")
                                    End If
                                Next
                            Else
                                ' The input file does have a text-based header
                                headerLine = String.Copy(lineIn)

                                FindScanNumColumn(inputFile, splitLine)

                                ' Clear splitLine so that this line gets skipped
                                ReDim splitLine(-1)
                            End If

                            Dim lstAdditionalHeaders = GetAdditionalMASICHeaders()

                            ' Populate blankAdditionalColumns with tab characters based on the number of items in lstAdditionalHeaders
                            blankAdditionalColumns = New String(ControlChars.Tab, lstAdditionalHeaders.Count - 1)

                            blankAdditionalSICColumns = New String(ControlChars.Tab, SIC_STAT_COLUMN_COUNT_TO_ADD)

                            ' Initialize blankAdditionalReporterIonColumns
                            If reporterIonHeaders.Length > 0 Then
                                blankAdditionalReporterIonColumns = New String(ControlChars.Tab, reporterIonHeaders.Split(ControlChars.Tab).ToList().Count - 1)
                            End If

                            ' Initialize the AddOn header columns
                            Dim addonHeaders = FlattenList(lstAdditionalHeaders)

                            If reporterIonHeaders.Length > 0 Then
                                ' Append the reporter ion stats columns
                                addonHeaders &= ControlChars.Tab & reporterIonHeaders
                                writeReporterIonStats = True
                            End If

                            ' Write out the headers
                            For index = 0 To outputFileCount - 1
                                swOutfile(index).WriteLine(headerLine & ControlChars.Tab & addonHeaders)
                            Next

                        End If

                        Dim scanNumber As Integer
                        If splitLine.Length < ScanNumberColumn OrElse Not Integer.TryParse(splitLine(ScanNumberColumn - 1), scanNumber) Then
                            Continue While
                        End If

                        ' Look for scanNumber in dctScanStats
                        Dim scanStatsEntry As clsScanStatsData = Nothing
                        Dim addonColumns As String

                        If Not dctScanStats.TryGetValue(scanNumber, scanStatsEntry) Then
                            ' Match not found; use the blank columns in blankAdditionalColumns
                            addonColumns = String.Copy(blankAdditionalColumns)
                        Else
                            With scanStatsEntry
                                addonColumns = .ElutionTime & ControlChars.Tab &
                                                  .ScanType & ControlChars.Tab &
                                                  .TotalIonIntensity & ControlChars.Tab &
                                                  .BasePeakIntensity & ControlChars.Tab &
                                                  .BasePeakMZ
                            End With

                            Dim sicStatsEntry As clsSICStatsData = Nothing
                            If Not dctSICStats.TryGetValue(scanNumber, sicStatsEntry) Then
                                ' Match not found; use the blank columns in blankAdditionalSICColumns
                                addonColumns &= ControlChars.Tab & blankAdditionalSICColumns

                                If writeReporterIonStats Then
                                    addonColumns &= ControlChars.Tab &
                                                       String.Empty & ControlChars.Tab &
                                                       blankAdditionalReporterIonColumns
                                End If
                            Else
                                With sicStatsEntry
                                    addonColumns &= ControlChars.Tab &
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

                            If writeReporterIonStats Then
                                With scanStatsEntry
                                    If String.IsNullOrWhiteSpace(.CollisionMode) Then
                                        ' Collision mode is not defined; append blank columns
                                        addonColumns &= ControlChars.Tab &
                                                           String.Empty & ControlChars.Tab &
                                                           blankAdditionalReporterIonColumns
                                    Else
                                        ' Collision mode is defined
                                        addonColumns &= ControlChars.Tab &
                                                           .CollisionMode & ControlChars.Tab &
                                                           .ReporterIonData

                                        collisionModeCurrentScan = String.Copy(.CollisionMode)
                                    End If
                                End With

                            ElseIf SeparateByCollisionMode Then
                                collisionModeCurrentScan = String.Copy(scanStatsEntry.CollisionMode)
                            End If
                        End If

                        Dim outFileIndex = 0
                        If SeparateByCollisionMode AndAlso outputFileCount > 1 Then
                            If Not collisionModeCurrentScan Is Nothing Then
                                ' Determine the correct output file
                                If Not dctCollisionModeFileMap.TryGetValue(collisionModeCurrentScan, outFileIndex) Then
                                    outFileIndex = 0
                                End If
                            End If
                        End If

                        swOutfile(outFileIndex).WriteLine(lineIn & ControlChars.Tab & addonColumns)
                        linesWritten(outFileIndex) += 1

                    End If
                End While

            End Using

            ' Close the output files
            If Not swOutfile Is Nothing Then
                For index = 0 To outputFileCount - 1
                    If Not swOutfile(index) Is Nothing Then
                        swOutfile(index).Close()
                    End If
                Next
            End If

            ' See if any of the files had no data written to them
            ' If there are, then delete the empty output file
            ' However, retain at least one output file
            Dim emptyOutFileCount = 0
            For index = 0 To outputFileCount - 1
                If linesWritten(index) = 0 Then
                    emptyOutFileCount += 1
                End If
            Next

            Dim outputPathEntry = New clsProcessedFileInfo(baseFileName)

            If emptyOutFileCount = 0 Then
                For Each item In outputFilePaths.ToList()
                    outputPathEntry.AddOutputFile(item.Key, item.Value)
                Next
            Else
                If emptyOutFileCount = outputFileCount Then
                    ' All output files are empty
                    ' Pretend the first output file actually contains data
                    linesWritten(0) = 1
                End If

                For index = 0 To outputFileCount - 1
                    ' Wait 250 msec before continuing
                    Thread.Sleep(250)

                    If linesWritten(index) = 0 Then
                        Try
                            ShowMessage("Deleting empty output file: " & ControlChars.NewLine & " --> " & Path.GetFileName(outputFilePaths(index).Value))
                            File.Delete(outputFilePaths(index).Value)
                        Catch ex As Exception
                            ' Ignore errors here
                        End Try
                    Else
                        outputPathEntry.AddOutputFile(outputFilePaths(index).Key, outputFilePaths(index).Value)
                    End If
                Next
            End If

            If outputPathEntry.OutputFiles.Count > 0 Then
                mProcessedDatasets.Add(outputPathEntry)
            End If

            Return True

        Catch ex As Exception
            HandleException("Error in MergePeptideHitAndMASICFiles", ex)
            Return False
        End Try

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

                    For index = 0 To baseFileName.Length - 1
                        If index >= candidateName.Length Then
                            Exit For
                        End If

                        If candidateName.ToLower().Chars(index) <> baseFileName.ToLower().Chars(index) Then
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
                        While Not srSourceFile.EndOfStream

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

    ''' <summary>
    ''' Main processing function
    ''' </summary>
    ''' <param name="inputFilePath">Input file path</param>
    ''' <param name="outputFolderPath">Output folder path</param>
    ''' <param name="parameterFilePath">Parameter file path (Ignored)</param>
    ''' <param name="resetErrorCode">If true, reset the error code</param>
    ''' <returns>True if success, False if failure</returns>
    Public Overloads Overrides Function ProcessFile(inputFilePath As String, outputFolderPath As String, parameterFilePath As String, resetErrorCode As Boolean) As Boolean

        Dim success As Boolean
        Dim masicResultsFolder As String

        If resetErrorCode Then
            SetLocalErrorCode(eResultsProcessorErrorCodes.NoError)
        End If

        If inputFilePath Is Nothing OrElse inputFilePath.Length = 0 Then
            ShowMessage("Input file name is empty")
            MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.InvalidInputFilePath)
            Return False
        End If

        ' Note that CleanupFilePaths() will update mOutputFolderPath, which is used by LogMessage()
        If Not CleanupFilePaths(inputFilePath, outputFolderPath) Then
            MyBase.SetBaseClassErrorCode(eProcessFilesErrorCodes.FilePathError)
            Return False
        End If

        Dim fiInputFile As FileInfo
        fiInputFile = New FileInfo(inputFilePath)

        If String.IsNullOrWhiteSpace(mMASICResultsFolderPath) Then
            masicResultsFolder = fiInputFile.DirectoryName
        Else
            masicResultsFolder = String.Copy(mMASICResultsFolderPath)
        End If

        If MageResults Then
            success = ProcessMageExtractorFile(fiInputFile, masicResultsFolder)
        Else
            success = ProcessSingleJobFile(fiInputFile, masicResultsFolder)
        End If
        Return success

    End Function

    Private Function ProcessMageExtractorFile(fiInputFile As FileInfo, masicResultsFolder As String) As Boolean

        Dim udtMASICFileNames = New udtMASICFileNamesType

        Dim dctScanStats = New Dictionary(Of Integer, clsScanStatsData)
        Dim dctSICStats = New Dictionary(Of Integer, clsSICStatsData)

        Try

            ' Read the Mage Metadata file
            Dim metadataFile = Path.Combine(fiInputFile.DirectoryName, Path.GetFileNameWithoutExtension(fiInputFile.Name) & "_metadata.txt")
            Dim fiMetadataFile = New FileInfo(metadataFile)
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

            Dim headerLine As String
            Dim jobColumnIndex As Integer

            ' Open the Mage Extractor data file so that we can validate and cache the header row
            Using srInFile = New StreamReader(New FileStream(fiInputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.Read))
                headerLine = srInFile.ReadLine()
                Dim lstColumns = headerLine.Split(ControlChars.Tab).ToList()
                jobColumnIndex = lstColumns.IndexOf("Job")
                If jobColumnIndex < 0 Then
                    ShowErrorMessage("Input file is not a valid Mage Extractor results file; it must contain a ""Job"" column: " & fiInputFile.FullName)
                    Return False
                End If
            End Using


            Dim lstAdditionalHeaders = GetAdditionalMASICHeaders()

            ' Populate blankAdditionalColumns with tab characters based on the number of items in lstAdditionalHeaders
            Dim blankAdditionalColumns = New String(ControlChars.Tab, lstAdditionalHeaders.Count - 1)

            Dim blankAdditionalSICColumns = New String(ControlChars.Tab, SIC_STAT_COLUMN_COUNT_TO_ADD)

            Dim outputFileName = Path.GetFileNameWithoutExtension(fiInputFile.Name) & RESULTS_SUFFIX
            Dim outputFilePath = Path.Combine(mOutputFolderPath, outputFileName)

            Dim jobsSuccessfullyMerged = 0


            ' Initialize the output file
            Using swOutFile = New StreamWriter(New FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.Read))

                ' Open the Mage Extractor data file and read the data for each job
                mPHRPReader = New clsPHRPReader(fiInputFile.FullName, clsPHRPReader.ePeptideHitResultType.Unknown, False, False, False) With {
                    .EchoMessagesToConsole = False,
                    .SkipDuplicatePSMs = False
                }

                RegisterEvents(mPHRPReader)

                If Not mPHRPReader.CanRead Then
                    ShowErrorMessage("Aborting since PHRPReader is not ready: " & mPHRPReader.ErrorMessage)
                    Return False
                End If

                Dim lastJob As Integer = -1
                Dim job As Integer = -1

                Dim masicDataLoaded = False
                Dim headerLineWritten = False
                Dim writeReporterIonStats As Boolean
                Dim reporterIonHeaders = String.Empty
                Dim blankAdditionalReporterIonColumns = String.Empty

                Do While mPHRPReader.MoveNext()

                    Dim oPSM As clsPSM = mPHRPReader.CurrentPSM

                    ' Parse out the job from the current line
                    Dim lstColumns = oPSM.DataLineText.Split(ControlChars.Tab).ToList()

                    If Not Integer.TryParse(lstColumns(jobColumnIndex), job) Then
                        ShowMessage("Warning: Job column does not contain a job number; skipping this entry: " & oPSM.DataLineText)
                        Continue Do
                    End If

                    If job <> lastJob Then

                        ' New job; read and cache the MASIC data
                        masicDataLoaded = False

                        Dim udtDatasetInfo = New udtDatasetInfoType

                        If Not dctJobToDatasetMap.TryGetValue(job, udtDatasetInfo) Then
                            ShowErrorMessage("Error: Job " & job & " was not defined in the Metadata file; unable to determine the dataset")
                        Else

                            ' Look for the corresponding MASIC files in the input folder
                            Dim success As Boolean

                            udtMASICFileNames.Initialize()
                            success = FindMASICFiles(masicResultsFolder, udtDatasetInfo, udtMASICFileNames)

                            If Not success Then
                                ShowMessage("  Error: Unable to find the MASIC data files for dataset " & udtDatasetInfo.DatasetName & " in " & masicResultsFolder)
                                ShowMessage("         Job " & job & " will not have MASIC results")
                            Else
                                If udtMASICFileNames.SICStatsFileName.Length = 0 Then
                                    ShowMessage("  Error: the SIC stats file was not found for dataset " & udtDatasetInfo.DatasetName & " in " & masicResultsFolder)
                                    ShowMessage("         Job " & job & " will not have MASIC results")
                                    success = False
                                ElseIf udtMASICFileNames.ScanStatsFileName.Length = 0 Then
                                    ShowMessage("  Error: the Scan stats file was not found for dataset " & udtDatasetInfo.DatasetName & " in " & masicResultsFolder)
                                    ShowMessage("         Job " & job & " will not have MASIC results")
                                    success = False
                                End If
                            End If

                            If success Then

                                ' Read and cache the MASIC data
                                dctScanStats = New Dictionary(Of Integer, clsScanStatsData)
                                dctSICStats = New Dictionary(Of Integer, clsSICStatsData)
                                reporterIonHeaders = String.Empty

                                masicDataLoaded = ReadMASICData(masicResultsFolder, udtMASICFileNames, dctScanStats, dctSICStats, reporterIonHeaders)

                                If masicDataLoaded Then
                                    jobsSuccessfullyMerged += 1

                                    If jobsSuccessfullyMerged = 1 Then

                                        ' Initialize blankAdditionalReporterIonColumns
                                        If reporterIonHeaders.Length > 0 Then
                                            blankAdditionalReporterIonColumns = New String(ControlChars.Tab, reporterIonHeaders.Split(ControlChars.Tab).ToList().Count - 1)
                                        End If

                                    End If
                                End If
                            End If

                        End If
                    End If

                    If masicDataLoaded Then

                        If Not headerLineWritten Then

                            Dim addonHeaderColumns = FlattenList(lstAdditionalHeaders)
                            If reporterIonHeaders.Length > 0 Then
                                ' Append the reporter ion stats columns
                                addonHeaderColumns &= ControlChars.Tab & reporterIonHeaders
                                writeReporterIonStats = True
                            End If

                            swOutFile.WriteLine(headerLine & ControlChars.Tab & addonHeaderColumns)

                            headerLineWritten = True
                        End If


                        ' Look for scanNumber in dctScanStats
                        Dim scanStatsEntry As clsScanStatsData = Nothing
                        Dim addonColumns As String

                        If Not dctScanStats.TryGetValue(oPSM.ScanNumber, scanStatsEntry) Then
                            ' Match not found; use the blank columns in blankAdditionalColumns
                            addonColumns = String.Copy(blankAdditionalColumns)
                        Else
                            With scanStatsEntry
                                addonColumns = .ElutionTime & ControlChars.Tab &
                                   .ScanType & ControlChars.Tab &
                                   .TotalIonIntensity & ControlChars.Tab &
                                   .BasePeakIntensity & ControlChars.Tab &
                                   .BasePeakMZ
                            End With

                            Dim sicStatsEntry As clsSICStatsData = Nothing
                            If Not dctSICStats.TryGetValue(oPSM.ScanNumber, sicStatsEntry) Then
                                ' Match not found; use the blank columns in blankAdditionalSICColumns
                                addonColumns &= ControlChars.Tab & blankAdditionalSICColumns

                                If writeReporterIonStats Then
                                    addonColumns &= ControlChars.Tab &
                                      String.Empty & ControlChars.Tab &
                                      blankAdditionalReporterIonColumns
                                End If
                            Else
                                With sicStatsEntry
                                    addonColumns &= ControlChars.Tab &
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

                            If writeReporterIonStats Then

                                With scanStatsEntry
                                    If String.IsNullOrWhiteSpace(.CollisionMode) Then
                                        ' Collision mode is not defined; append blank columns
                                        addonColumns &= ControlChars.Tab &
                                         String.Empty & ControlChars.Tab &
                                         blankAdditionalReporterIonColumns
                                    Else
                                        ' Collision mode is defined
                                        addonColumns &= ControlChars.Tab &
                                         .CollisionMode & ControlChars.Tab &
                                         .ReporterIonData

                                    End If
                                End With


                            End If
                        End If

                        swOutFile.WriteLine(oPSM.DataLineText & ControlChars.Tab & addonColumns)
                    Else
                        swOutFile.WriteLine(oPSM.DataLineText & ControlChars.Tab & blankAdditionalColumns)
                    End If

                    UpdateProgress("Loading data from " & fiInputFile.Name, mPHRPReader.PercentComplete)
                    lastJob = job
                Loop

            End Using

            If jobsSuccessfullyMerged > 0 Then
                Console.WriteLine()
                ShowMessage("Merged MASIC results for " & jobsSuccessfullyMerged & " jobs")
            End If

            If jobsSuccessfullyMerged > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            HandleException("Error in ProcessMageExtractorFile", ex)
            Return False
        End Try

    End Function

    Private Function ProcessSingleJobFile(fiInputFile As FileSystemInfo, masicResultsFolder As String) As Boolean
        Dim udtMASICFileNames = New udtMASICFileNamesType

        Dim dctScanStats As Dictionary(Of Integer, clsScanStatsData)
        Dim dctSICStats As Dictionary(Of Integer, clsSICStatsData)

        Dim reporterIonHeaders As String = String.Empty

        Dim success As Boolean

        Try
            Dim udtDatasetInfo = New udtDatasetInfoType

            ' Note that FindMASICFiles will first try the full filename, and if it doesn't find a match,
            ' it will start removing text from the end of the filename by looking for underscores
            udtDatasetInfo.DatasetName = Path.GetFileNameWithoutExtension(fiInputFile.FullName)
            udtDatasetInfo.DatasetID = 0

            ' Look for the corresponding MASIC files in the input folder
            udtMASICFileNames.Initialize()
            success = FindMASICFiles(masicResultsFolder, udtDatasetInfo, udtMASICFileNames)

            If Not success Then
                ShowErrorMessage("Error: Unable to find the MASIC data files in " & masicResultsFolder)
                SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMASICFiles)
                Return False
            Else
                If udtMASICFileNames.SICStatsFileName.Length = 0 Then
                    ShowErrorMessage("Error: the SIC stats file was not found in " & masicResultsFolder & "; unable to continue")
                    SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMASICFiles)
                    Return False
                ElseIf udtMASICFileNames.ScanStatsFileName.Length = 0 Then
                    ShowErrorMessage("Error: the Scan stats file was not found in " & masicResultsFolder & "; unable to continue")
                    SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMASICFiles)
                    Return False
                End If
            End If

            ' Read and cache the MASIC data
            dctScanStats = New Dictionary(Of Integer, clsScanStatsData)
            dctSICStats = New Dictionary(Of Integer, clsSICStatsData)

            success = ReadMASICData(masicResultsFolder, udtMASICFileNames, dctScanStats, dctSICStats, reporterIonHeaders)

            If success Then
                ' Merge the MASIC data with the input file
                success = MergePeptideHitAndMASICFiles(fiInputFile, mOutputFolderPath,
                 dctScanStats,
                 dctSICStats,
                 reporterIonHeaders)
            End If

            If success Then
                ShowMessage(String.Empty, False)
            Else
                SetLocalErrorCode(eResultsProcessorErrorCodes.UnspecifiedError)
                ShowErrorMessage("Error")
            End If

        Catch ex As Exception
            HandleException("Error in ProcessSingleJobFile", ex)
        End Try

        Return success

    End Function

    Private Function ReadMASICData(sourceFolder As String,
      udtMASICFileNames As udtMASICFileNamesType,
      dctScanStats As IDictionary(Of Integer, clsScanStatsData),
      dctSICStats As IDictionary(Of Integer, clsSICStatsData),
      <Out> ByRef reporterIonHeaders As String) As Boolean

        Dim success As Boolean

        Try

            success = ReadScanStatsFile(sourceFolder, udtMASICFileNames.ScanStatsFileName, dctScanStats)

            If success Then
                success = ReadSICStatsFile(sourceFolder, udtMASICFileNames.SICStatsFileName, dctSICStats)
            End If

            If success AndAlso udtMASICFileNames.ReporterIonsFileName.Length > 0 Then
                success = ReadReporterIonStatsFile(sourceFolder, udtMASICFileNames.ReporterIonsFileName, dctScanStats, reporterIonHeaders)
            Else
                reporterIonHeaders = String.Empty
            End If

            Return success

        Catch ex As Exception
            HandleException("Error in ReadMASICData", ex)
            reporterIonHeaders = String.Empty
            Return False
        End Try

    End Function

    Private Function ReadScanStatsFile(sourceFolder As String,
      scanStatsFileName As String,
      dctScanStats As IDictionary(Of Integer, clsScanStatsData)) As Boolean

        Dim lineIn As String
        Dim splitLine() As String

        Dim linesRead As Integer
        Dim scanNumber As Integer

        Try
            ' Initialize dctScanStats
            dctScanStats.Clear()

            ShowMessage("  Reading: " & scanStatsFileName)

            Using srInFile = New StreamReader(New FileStream(Path.Combine(sourceFolder, scanStatsFileName), FileMode.Open, FileAccess.Read, FileShare.Read))

                linesRead = 0
                While Not srInFile.EndOfStream
                    lineIn = srInFile.ReadLine()

                    If Not lineIn Is Nothing AndAlso lineIn.Length > 0 Then
                        linesRead += 1
                        splitLine = lineIn.Split(ControlChars.Tab)

                        If splitLine.Length >= eScanStatsColumns.BasePeakMZ + 1 AndAlso Integer.TryParse(splitLine(eScanStatsColumns.ScanNumber), scanNumber) Then

                            Dim scanStatsEntry = New clsScanStatsData(scanNumber)
                            With scanStatsEntry
                                ' Note: the remaining values are stored as strings to prevent the number format from changing
                                .ElutionTime = String.Copy(splitLine(eScanStatsColumns.ScanTime))
                                .ScanType = String.Copy(splitLine(eScanStatsColumns.ScanType))
                                .TotalIonIntensity = String.Copy(splitLine(eScanStatsColumns.TotalIonIntensity))
                                .BasePeakIntensity = String.Copy(splitLine(eScanStatsColumns.BasePeakIntensity))
                                .BasePeakMZ = String.Copy(splitLine(eScanStatsColumns.BasePeakMZ))

                                .CollisionMode = String.Empty
                                .ReporterIonData = String.Empty
                            End With

                            dctScanStats.Add(scanNumber, scanStatsEntry)
                        End If
                    End If
                End While

            End Using

            Return True

        Catch ex As Exception
            HandleException("Error in ReadScanStatsFile", ex)
            Return False
        End Try

    End Function

    Private Function ReadMageMetadataFile(metadataFilePath As String) As Dictionary(Of Integer, udtDatasetInfoType)

        Dim dctJobToDatasetMap = New Dictionary(Of Integer, udtDatasetInfoType)
        Dim lineIn As String
        Dim lstData As List(Of String)
        Dim headersParsed As Boolean

        Dim jobIndex As Integer = -1
        Dim datasetIndex As Integer = -1
        Dim datasetIDIndex As Integer = -1

        Try
            Using srInFile = New StreamReader(New FileStream(metadataFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))

                While Not srInFile.EndOfStream
                    lineIn = srInFile.ReadLine()

                    If Not String.IsNullOrWhiteSpace(lineIn) Then
                        lstData = lineIn.Split(ControlChars.Tab).ToList()

                        If Not headersParsed Then
                            ' Look for the Job and Dataset columns
                            jobIndex = lstData.IndexOf("Job")
                            datasetIndex = lstData.IndexOf("Dataset")
                            datasetIDIndex = lstData.IndexOf("Dataset_ID")

                            If jobIndex < 0 Then
                                ShowErrorMessage("Job column not found in the metadata file: " & metadataFilePath)
                                Return Nothing
                            ElseIf datasetIndex < 0 Then
                                ShowErrorMessage("Dataset column not found in the metadata file: " & metadataFilePath)
                                Return Nothing
                            ElseIf datasetIDIndex < 0 Then
                                ShowErrorMessage("Dataset_ID column not found in the metadata file: " & metadataFilePath)
                                Return Nothing
                            End If

                            headersParsed = True
                            Continue While
                        End If

                        If lstData.Count > datasetIndex Then
                            Dim jobNumber As Integer
                            Dim datasetID As Integer

                            If Integer.TryParse(lstData(jobIndex), jobNumber) Then
                                If Integer.TryParse(lstData(datasetIDIndex), datasetID) Then
                                    Dim udtDatasetInfo = New udtDatasetInfoType
                                    udtDatasetInfo.DatasetID = datasetID
                                    udtDatasetInfo.DatasetName = lstData(datasetIndex)

                                    dctJobToDatasetMap.Add(jobNumber, udtDatasetInfo)
                                Else
                                    ShowMessage("Warning: Dataest_ID number not numeric in metadata file, line " & lineIn)
                                End If
                            Else
                                ShowMessage("Warning: Job number not numeric in metadata file, line " & lineIn)
                            End If
                        End If
                    End If
                End While
            End Using

        Catch ex As Exception
            HandleException("Error in ReadMageMetadataFile", ex)
            Return Nothing
        End Try

        Return dctJobToDatasetMap

    End Function

    Private Function ReadSICStatsFile(sourceFolder As String,
      sicStatsFileName As String,
      dctSICStats As IDictionary(Of Integer, clsSICStatsData)) As Boolean

        Dim lineIn As String
        Dim splitLine() As String

        Dim linesRead As Integer
        Dim fragScanNumber As Integer

        Try
            ' Initialize dctSICStats
            dctSICStats.Clear()

            ShowMessage("  Reading: " & sicStatsFileName)

            Using srInFile = New StreamReader(New FileStream(Path.Combine(sourceFolder, sicStatsFileName), FileMode.Open, FileAccess.Read, FileShare.Read))

                linesRead = 0
                While Not srInFile.EndOfStream
                    lineIn = srInFile.ReadLine()

                    If Not lineIn Is Nothing AndAlso lineIn.Length > 0 Then
                        linesRead += 1
                        splitLine = lineIn.Split(ControlChars.Tab)

                        If splitLine.Length >= eSICStatsColumns.StatMomentsArea + 1 AndAlso Integer.TryParse(splitLine(eSICStatsColumns.FragScanNumber), fragScanNumber) Then

                            Dim sicStatsEntry = New clsSICStatsData(fragScanNumber)
                            With sicStatsEntry
                                ' Note: the remaining values are stored as strings to prevent the number format from changing
                                .OptimalScanNumber = String.Copy(splitLine(eSICStatsColumns.OptimalPeakApexScanNumber))
                                .PeakMaxIntensity = String.Copy(splitLine(eSICStatsColumns.PeakMaxIntensity))
                                .PeakSignalToNoiseRatio = String.Copy(splitLine(eSICStatsColumns.PeakSignalToNoiseRatio))
                                .FWHMInScans = String.Copy(splitLine(eSICStatsColumns.FWHMInScans))
                                .PeakArea = String.Copy(splitLine(eSICStatsColumns.PeakArea))
                                .ParentIonIntensity = String.Copy(splitLine(eSICStatsColumns.ParentIonIntensity))
                                .ParentIonMZ = String.Copy(splitLine(eSICStatsColumns.MZ))
                                .StatMomentsArea = String.Copy(splitLine(eSICStatsColumns.StatMomentsArea))
                            End With

                            dctSICStats.Add(fragScanNumber, sicStatsEntry)
                        End If
                    End If
                End While

            End Using

            Return True

        Catch ex As Exception
            HandleException("Error in ReadSICStatsFile", ex)
            Return False
        End Try

    End Function

    Private Function ReadReporterIonStatsFile(sourceFolder As String,
      reporterIonStatsFileName As String,
      dctScanStats As IDictionary(Of Integer, clsScanStatsData),
      <Out> ByRef reporterIonHeaders As String) As Boolean

        Dim lineIn As String
        Dim splitLine() As String

        Dim linesRead As Integer
        Dim scanNumber As Integer

        Dim warningCount = 0

        reporterIonHeaders = String.Empty

        Try

            ShowMessage("  Reading: " & reporterIonStatsFileName)

            Using srInFile = New StreamReader(New FileStream(Path.Combine(sourceFolder, reporterIonStatsFileName), FileMode.Open, FileAccess.Read, FileShare.Read))

                linesRead = 0
                While Not srInFile.EndOfStream
                    lineIn = srInFile.ReadLine()

                    If Not lineIn Is Nothing AndAlso lineIn.Length > 0 Then
                        linesRead += 1
                        splitLine = lineIn.Split(ControlChars.Tab)

                        If linesRead = 1 Then
                            ' This is the header line; we need to cache it

                            If splitLine.Length >= eReporterIonStatsColumns.ReporterIonIntensityMax + 1 Then
                                reporterIonHeaders = splitLine(eReporterIonStatsColumns.CollisionMode)
                                reporterIonHeaders &= ControlChars.Tab & FlattenArray(splitLine, eReporterIonStatsColumns.ReporterIonIntensityMax)
                            Else
                                ' There aren't enough columns in the header line; this is unexpected
                                reporterIonHeaders = "Collision Mode" & ControlChars.Tab & "AdditionalReporterIonColumns"
                            End If

                        End If

                        If splitLine.Length >= eReporterIonStatsColumns.ReporterIonIntensityMax + 1 AndAlso Integer.TryParse(splitLine(eReporterIonStatsColumns.ScanNumber), scanNumber) Then

                            ' Look for scanNumber in scanNumbers
                            Dim scanStatsEntry As clsScanStatsData = Nothing

                            If Not dctScanStats.TryGetValue(scanNumber, scanStatsEntry) Then
                                If warningCount < 10 Then
                                    ShowMessage("Warning: " & REPORTER_IONS_FILE_EXTENSION & " file refers to scan " & scanNumber.ToString & ", but that scan was not in the _ScanStats.txt file")
                                ElseIf warningCount = 10 Then
                                    ShowMessage("Warning: " & REPORTER_IONS_FILE_EXTENSION & " file has 10 or more scan numbers that are not defined in the _ScanStats.txt file")
                                End If
                                warningCount += 1
                            Else

                                If scanStatsEntry.ScanNumber <> scanNumber Then
                                    ' Scan number mismatch; this shouldn't happen
                                    ShowMessage("Error: Scan number mismatch in ReadReporterIonStatsFile: " & scanStatsEntry.ScanNumber.ToString & " vs. " & scanNumber.ToString)
                                Else
                                    scanStatsEntry.CollisionMode = String.Copy(splitLine(eReporterIonStatsColumns.CollisionMode))
                                    scanStatsEntry.ReporterIonData = FlattenArray(splitLine, eReporterIonStatsColumns.ReporterIonIntensityMax)
                                End If

                            End If

                        End If
                    End If
                End While

            End Using

            Return True

        Catch ex As Exception
            HandleException("Error in ReadSICStatsFile", ex)
            Return False
        End Try

    End Function

    Private Sub SetLocalErrorCode(eNewErrorCode As eResultsProcessorErrorCodes)
        SetLocalErrorCode(eNewErrorCode, False)
    End Sub

    Private Sub SetLocalErrorCode(eNewErrorCode As eResultsProcessorErrorCodes, leaveExistingErrorCodeUnchanged As Boolean)

        If leaveExistingErrorCodeUnchanged AndAlso mLocalErrorCode <> eResultsProcessorErrorCodes.NoError Then
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

    Private Function SummarizeCollisionModes(
      inputFile As FileSystemInfo,
      baseFileName As String,
      outputFolderPath As String,
      dctScanStats As Dictionary(Of Integer, clsScanStatsData),
      <Out> ByRef dctCollisionModeFileMap As Dictionary(Of String, Integer)) As KeyValuePair(Of String, String)()

        ' Construct a list of the different collision modes in dctScanStats

        dctCollisionModeFileMap = New Dictionary(Of String, Integer)(StringComparer.CurrentCultureIgnoreCase)

        Dim collisionModeTypeCount = 0

        For Each scanStatsItem In dctScanStats.Values
            If Not dctCollisionModeFileMap.ContainsKey(scanStatsItem.CollisionMode) Then
                ' Store this collision mode in htCollisionModes; the value stored will be the index in collisionModes()
                dctCollisionModeFileMap.Add(scanStatsItem.CollisionMode, collisionModeTypeCount)
                collisionModeTypeCount += 1
            End If
        Next

        If (dctCollisionModeFileMap.Count = 0) OrElse
           (dctCollisionModeFileMap.Count = 1 AndAlso String.IsNullOrWhiteSpace(dctCollisionModeFileMap.First.Key)) Then

            ' Try to load the collision mode info from the intput file
            ' MSGF+ results report this in the FragMethod column

            dctCollisionModeFileMap.Clear()
            collisionModeTypeCount = 0

            Try

                Using srInFile = New StreamReader(New FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.Read))

                    Dim linesRead = 0
                    Dim fragMethodColNumber = 0

                    While Not srInFile.EndOfStream
                        Dim lineIn = srInFile.ReadLine()

                        If Not lineIn Is Nothing AndAlso lineIn.Length > 0 Then
                            linesRead += 1
                            Dim splitLine = lineIn.Split(ControlChars.Tab)
                            If linesRead = 1 Then
                                ' Header line; look for the FragMethod column
                                For colIndex = 0 To splitLine.Length - 1
                                    If String.Equals(splitLine(colIndex), "FragMethod", StringComparison.OrdinalIgnoreCase) Then
                                        fragMethodColNumber = colIndex + 1
                                        Exit For
                                    End If
                                Next

                                If fragMethodColNumber = 0 Then
                                    ' Fragmentation method column not found
                                    ShowWarning("Unable to determine the collision mode for results being merged. " &
                                                "This is typically obtained from a MASIC _ReporterIons.txt file " &
                                                "or from the FragMethod column in the MSGF+ results file")
                                    Exit While
                                End If

                                ' Also look for the scan number column, auto-updating ScanNumberColumn if necessary
                                FindScanNumColumn(inputFile, splitLine)

                                Continue While
                            End If

                            Dim scanNumber As Integer
                            If splitLine.Length < ScanNumberColumn OrElse Not Integer.TryParse(splitLine(ScanNumberColumn - 1), scanNumber) Then
                                Continue While
                            End If

                            If splitLine.Length < fragMethodColNumber Then
                                Continue While
                            End If

                            Dim collisionMode = splitLine(fragMethodColNumber - 1)

                            If Not dctCollisionModeFileMap.ContainsKey(collisionMode) Then
                                ' Store this collision mode in htCollisionModes; the value stored will be the index in collisionModes()
                                dctCollisionModeFileMap.Add(collisionMode, collisionModeTypeCount)
                                collisionModeTypeCount += 1
                            End If

                            Dim scanStatsEntry As clsScanStatsData = Nothing
                            If Not dctScanStats.TryGetValue(scanNumber, scanStatsEntry) Then
                                scanStatsEntry = New clsScanStatsData(scanNumber)
                                scanStatsEntry.CollisionMode = collisionMode
                                dctScanStats.Add(scanNumber, scanStatsEntry)
                            Else
                                scanStatsEntry.CollisionMode = collisionMode
                            End If
                        End If
                    End While
                End Using

            Catch ex As Exception
                HandleException("Error extraction collision mode information from the intput file", ex)
                Return New KeyValuePair(Of String, String)(collisionModeTypeCount - 1) {}
            End Try

        End If

        If collisionModeTypeCount = 0 Then collisionModeTypeCount = 1

        Dim outputFilePaths = New KeyValuePair(Of String, String)(collisionModeTypeCount - 1) {}

        If dctCollisionModeFileMap.Count = 0 Then
            outputFilePaths(0) = New KeyValuePair(Of String, String)("na", Path.Combine(outputFolderPath, baseFileName & "_na" & RESULTS_SUFFIX))
        Else
            For Each oItem In dctCollisionModeFileMap
                Dim collisionMode As String = oItem.Key
                If String.IsNullOrWhiteSpace(collisionMode) Then collisionMode = "na"
                outputFilePaths(oItem.Value) = New KeyValuePair(Of String, String)(
                    collisionMode, Path.Combine(outputFolderPath, baseFileName & "_" & collisionMode & RESULTS_SUFFIX))
            Next
        End If

        Return outputFilePaths

    End Function

End Class
