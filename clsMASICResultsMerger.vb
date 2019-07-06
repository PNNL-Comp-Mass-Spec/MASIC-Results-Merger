Option Strict On
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports PHRPReader
Imports PRISM

' This class merges the contents of a tab-delimited peptide hit results file
' (e.g. from SEQUEST, X!Tandem, or MS-GF+) with the corresponding MASIC results files,
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
        MyBase.mFileDate = "July 5, 2019"
        InitializeLocalVariables()
    End Sub

#Region "Constants and Enums"

    Public Const SIC_STATS_FILE_EXTENSION As String = "_SICStats.txt"
    Public Const SCAN_STATS_FILE_EXTENSION As String = "_ScanStats.txt"
    Public Const REPORTER_IONS_FILE_EXTENSION As String = "_ReporterIons.txt"

    Public Const RESULTS_SUFFIX As String = "_PlusSICStats.txt"
    Public Const DEFAULT_SCAN_NUMBER_COLUMN As Integer = 2

    ''' <summary>
    ''' Error codes specialized for this class
    ''' </summary>
    Public Enum eResultsProcessorErrorCodes As Integer
        NoError = 0
        MissingMASICFiles = 1
        MissingMageFiles = 2
        UnspecifiedError = -1
    End Enum

    ' ReSharper disable UnusedMember.Local
    ' ReSharper disable UnusedMember.Global
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

    ' ReSharper restore UnusedMember.Local
    ' ReSharper restore UnusedMember.Global

#End Region

#Region "Classwide Variables"

    Private mLocalErrorCode As eResultsProcessorErrorCodes

    Private mMASICResultsDirectoryPath As String = String.Empty

    Private mPHRPReader As clsPHRPReader

    Private mProcessedDatasets As List(Of clsProcessedFileInfo)

#End Region

#Region "Properties"

    Public Property MageResults As Boolean

    Public Property MASICResultsDirectoryPath As String
        Get
            If mMASICResultsDirectoryPath Is Nothing Then mMASICResultsDirectoryPath = String.Empty
            Return mMASICResultsDirectoryPath
        End Get
        Set
            If Value Is Nothing Then Value = String.Empty
            mMASICResultsDirectoryPath = Value
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

    Private Function FindMASICFiles(
      masicResultsDirectory As String,
      datasetInfo As clsDatasetInfo,
      masicFiles As clsMASICFileInfo,
      masicFileSearchInfo As String,
      job As Integer) As Boolean

        Dim triedDatasetID = False
        Dim success = False

        Try
            Console.WriteLine()
            ShowMessage("Looking for MASIC data files that correspond to " & datasetInfo.DatasetName)

            Dim datasetName = String.Copy(datasetInfo.DatasetName)

            ' Use a loop to try various possible dataset names
            While True
                Dim scanStatsFile = New FileInfo(Path.Combine(masicResultsDirectory, datasetName & SCAN_STATS_FILE_EXTENSION))
                Dim sicStatsFile = New FileInfo(Path.Combine(masicResultsDirectory, datasetName & SIC_STATS_FILE_EXTENSION))
                Dim reportIonsFile = New FileInfo(Path.Combine(masicResultsDirectory, datasetName & REPORTER_IONS_FILE_EXTENSION))

                If scanStatsFile.Exists OrElse sicStatsFile.Exists Then

                    If scanStatsFile.Exists Then
                        masicFiles.ScanStatsFileName = scanStatsFile.Name
                    End If

                    If sicStatsFile.Exists Then
                        masicFiles.SICStatsFileName = sicStatsFile.Name
                    End If

                    If reportIonsFile.Exists Then
                        masicFiles.ReporterIonsFileName = reportIonsFile.Name
                    End If

                    success = True
                    Exit While
                End If

                ' Find the last underscore in datasetName, then remove it and any text after it
                Dim charIndex = datasetName.LastIndexOf("_"c)

                If charIndex > 0 Then
                    datasetName = datasetName.Substring(0, charIndex)
                Else
                    If Not triedDatasetID AndAlso datasetInfo.DatasetID > 0 Then
                        datasetName = datasetInfo.DatasetID & "_" & datasetInfo.DatasetName
                        triedDatasetID = True
                    Else
                        ' No more underscores; we're unable to determine the dataset name
                        Exit While
                    End If
                End If

            End While

        Catch ex As Exception
            HandleException("Error in FindMASICFiles", ex)
        End Try

        If Not success Then
            ShowMessage("  Error: Unable to find the MASIC data files " & masicFileSearchInfo)
            If job <> 0 Then
                ShowMessage("         Job " & job & " will not have MASIC results")
            End If
            Return False
        End If

        If String.IsNullOrWhiteSpace(masicFiles.ScanStatsFileName) AndAlso
           String.IsNullOrWhiteSpace(masicFiles.SICStatsFileName) Then
            ShowMessage("  Error: the SIC stats and/or scan stats files were not found " & masicFileSearchInfo)
            If job <> 0 Then
                ShowMessage("         Job " & job & " will not have MASIC results")
            End If
            Return False
        End If

        If String.IsNullOrWhiteSpace(masicFiles.ScanStatsFileName) Then
            ShowMessage("  Note: The MASIC SIC stats file was found, but the ScanStats file dose not exist " & masicFileSearchInfo)
        ElseIf String.IsNullOrWhiteSpace(masicFiles.SICStatsFileName) Then
            ShowMessage("  Note: The MASIC ScanStats file was found, but the SIC stats file dose not exist " & masicFileSearchInfo)
        End If

        Return True
    End Function

    Private Sub FindScanNumColumn(inputFile As FileSystemInfo, lineParts As IList(Of String))

        ' Check for a column named "ScanNum" or "ScanNumber" or "Scan Num" or "Scan Number"
        ' If found, override ScanNumberColumn
        For colIndex = 0 To lineParts.Count - 1
            If lineParts(colIndex).Equals("Scan", StringComparison.OrdinalIgnoreCase) OrElse
               lineParts(colIndex).Equals("ScanNum", StringComparison.OrdinalIgnoreCase) OrElse
               lineParts(colIndex).Equals("Scan Num", StringComparison.OrdinalIgnoreCase) OrElse
               lineParts(colIndex).Equals("ScanNumber", StringComparison.OrdinalIgnoreCase) OrElse
               lineParts(colIndex).Equals("Scan Number", StringComparison.OrdinalIgnoreCase) OrElse
               lineParts(colIndex).Equals("Scan#", StringComparison.OrdinalIgnoreCase) OrElse
               lineParts(colIndex).Equals("Scan #", StringComparison.OrdinalIgnoreCase) Then

                If ScanNumberColumn <> colIndex + 1 Then
                    ScanNumberColumn = colIndex + 1
                    ShowMessage(
                        String.Format("Note: Reading scan numbers from column {0} ({1}) in file {2}",
                                      ScanNumberColumn, lineParts(colIndex), inputFile.Name))
                End If
            End If
        Next
    End Sub

    Private Function GetScanStatsHeaders() As List(Of String)

        Dim scanStatsColumns = New List(Of String) From {
            "ElutionTime",
            "ScanType",
            "TotalIonIntensity",
            "BasePeakIntensity",
            "BasePeakMZ"
        }

        Return scanStatsColumns

    End Function

    Private Function GetSICStatsHeaders() As List(Of String)

        Dim sicStatsColumns = New List(Of String) From {
            "Optimal_Scan_Number",
            "PeakMaxIntensity",
            "PeakSignalToNoiseRatio",
            "FWHMInScans",
            "PeakArea",
            "ParentIonIntensity",
            "ParentIonMZ",
            "StatMomentsArea"
        }

        Return sicStatsColumns

    End Function

    Private Function FlattenList(lstData As IReadOnlyList(Of String)) As String
        Return String.Join(ControlChars.Tab, lstData)
    End Function

    Private Function FlattenArray(lineParts As IList(Of String), indexStart As Integer) As String

        Dim text As String = String.Empty
        Dim column As String

        Dim index As Integer

        If (lineParts Is Nothing) OrElse lineParts.Count <= 0 Then Return text

        For index = indexStart To lineParts.Count - 1

            If lineParts(index) Is Nothing Then
                column = String.Empty
            Else
                column = String.Copy(lineParts(index))
            End If

            If index > indexStart Then
                text &= ControlChars.Tab & column
            Else
                text = String.Copy(column)
            End If
        Next


        Return text

    End Function

    ''' <summary>
    ''' Get the current error state, if any
    ''' </summary>
    ''' <returns>Returns an empty string if no error</returns>
    Public Overrides Function GetErrorMessage() As String

        Dim errorMessage As String

        If MyBase.ErrorCode = ProcessFilesErrorCodes.LocalizedError Or
           MyBase.ErrorCode = ProcessFilesErrorCodes.NoError Then
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

        mMASICResultsDirectoryPath = String.Empty
        ScanNumberColumn = DEFAULT_SCAN_NUMBER_COLUMN
        SeparateByCollisionMode = False

        mLocalErrorCode = eResultsProcessorErrorCodes.NoError

        mProcessedDatasets = New List(Of clsProcessedFileInfo)

    End Sub

    Private Function MergePeptideHitAndMASICFiles(
      inputFile As FileSystemInfo,
      outputDirectoryPath As String,
      dctScanStats As Dictionary(Of Integer, clsScanStatsData),
      dctSICStats As IReadOnlyDictionary(Of Integer, clsSICStatsData),
      reporterIonHeaders As String) As Boolean

        Dim writers() As StreamWriter
        Dim linesWritten() As Integer

        Dim outputFileCount As Integer

        ' The Key is the collision mode and the value is the path
        Dim outputFilePaths() As KeyValuePair(Of String, String)

        Dim baseFileName As String

        Dim blankAdditionalScanStatsColumns As String = String.Empty
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
            If outputDirectoryPath Is Nothing Then
                outputDirectoryPath = String.Empty
            End If

            If SeparateByCollisionMode Then
                outputFilePaths = SummarizeCollisionModes(
                    inputFile,
                    baseFileName,
                    outputDirectoryPath,
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
                outputFilePaths(0) = New KeyValuePair(Of String, String)("", Path.Combine(outputDirectoryPath, baseFileName & RESULTS_SUFFIX))
            End If

            ReDim writers(outputFileCount - 1)
            ReDim linesWritten(outputFileCount - 1)

            ' Open the output file(s)
            For index = 0 To outputFileCount - 1
                writers(index) = New StreamWriter(New FileStream(outputFilePaths(index).Value, FileMode.Create, FileAccess.Write, FileShare.Read))
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

            ' Read from reader and write out to the file(s) in swOutFile
            Using reader = New StreamReader(New FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

                Dim linesRead = 0
                Dim writeReporterIonStats = False
                Dim writeSICStats = dctSICStats.Count > 0

                While Not reader.EndOfStream
                    Dim dataLine = reader.ReadLine()
                    Dim collisionModeCurrentScan = String.Empty

                    If String.IsNullOrWhiteSpace(dataLine) Then Continue While

                    linesRead += 1
                    Dim lineParts = dataLine.Split(ControlChars.Tab).ToList()

                    If linesRead = 1 Then

                        Dim headerLine As String

                        Dim integerValue As Integer

                        ' Write out an updated header line
                        If lineParts.Count >= ScanNumberColumn AndAlso Integer.TryParse(lineParts(ScanNumberColumn - 1), integerValue) Then
                            ' The input file doesn't have a header line; we will add one, using generic column names for the data in the input file

                            Dim genericHeaders = New List(Of String)
                            For index = 0 To lineParts.Count - 1
                                genericHeaders.Add("Column" & index.ToString("00"))
                            Next
                            headerLine = FlattenList(genericHeaders)
                        Else
                            ' The input file does have a text-based header
                            headerLine = String.Copy(dataLine)

                            FindScanNumColumn(inputFile, lineParts)

                            ' Clear splitLine so that this line gets skipped
                            lineParts.Clear()
                        End If

                        Dim scanStatsHeaders = GetScanStatsHeaders()
                        Dim sicStatsHeaders = GetSICStatsHeaders()

                        If Not writeSICStats Then
                            sicStatsHeaders.Clear()
                        End If

                        ' Populate blankAdditionalScanStatsColumns with tab characters based on the number of items in scanStatsHeaders
                        blankAdditionalScanStatsColumns = New String(ControlChars.Tab, scanStatsHeaders.Count - 1)

                        If writeSICStats Then
                            blankAdditionalSICColumns = New String(ControlChars.Tab, sicStatsHeaders.Count)
                        End If

                        ' Initialize blankAdditionalReporterIonColumns
                        If reporterIonHeaders.Length > 0 Then
                            blankAdditionalReporterIonColumns = New String(ControlChars.Tab, reporterIonHeaders.Split(ControlChars.Tab).ToList().Count - 1)
                        End If

                        ' Initialize the AddOn header columns
                        Dim addonHeaders = FlattenList(scanStatsHeaders)
                        If writeSICStats Then
                            addonHeaders &= ControlChars.Tab & FlattenList(sicStatsHeaders)
                        End If

                        If reporterIonHeaders.Length > 0 Then
                            ' Append the reporter ion stats columns
                            addonHeaders &= ControlChars.Tab & reporterIonHeaders
                            writeReporterIonStats = True
                        End If

                        ' Write out the headers
                        For index = 0 To outputFileCount - 1
                            writers(index).WriteLine(headerLine & ControlChars.Tab & addonHeaders)
                        Next

                    End If

                    Dim scanNumber As Integer
                    If lineParts.Count < ScanNumberColumn OrElse Not Integer.TryParse(lineParts(ScanNumberColumn - 1), scanNumber) Then
                        Continue While
                    End If

                    ' Look for scanNumber in dctScanStats
                    Dim scanStatsEntry As clsScanStatsData = Nothing
                    Dim addonColumns = New List(Of String)

                    If Not dctScanStats.TryGetValue(scanNumber, scanStatsEntry) Then
                        ' Match not found; use the blank columns in blankAdditionalScanStatsColumns
                        addonColumns.Add(blankAdditionalScanStatsColumns)
                    Else
                        addonColumns.Add(scanStatsEntry.ElutionTime)
                        addonColumns.Add(scanStatsEntry.ScanType)
                        addonColumns.Add(scanStatsEntry.TotalIonIntensity)
                        addonColumns.Add(scanStatsEntry.BasePeakIntensity)
                        addonColumns.Add(scanStatsEntry.BasePeakMZ)
                    End If

                    If writeSICStats Then
                        Dim sicStatsEntry As clsSICStatsData = Nothing
                        If Not dctSICStats.TryGetValue(scanNumber, sicStatsEntry) Then
                            ' Match not found; use the blank columns in blankAdditionalSICColumns
                            addonColumns.Add(blankAdditionalSICColumns)
                        Else
                            addonColumns.Add(sicStatsEntry.OptimalScanNumber)
                            addonColumns.Add(sicStatsEntry.PeakMaxIntensity)
                            addonColumns.Add(sicStatsEntry.PeakSignalToNoiseRatio)
                            addonColumns.Add(sicStatsEntry.FWHMInScans)
                            addonColumns.Add(sicStatsEntry.PeakArea)
                            addonColumns.Add(sicStatsEntry.ParentIonIntensity)
                            addonColumns.Add(sicStatsEntry.ParentIonMZ)
                            addonColumns.Add(sicStatsEntry.StatMomentsArea)
                        End If
                    End If

                    If writeReporterIonStats Then

                        If scanStatsEntry Is Nothing OrElse String.IsNullOrWhiteSpace(scanStatsEntry.CollisionMode) Then
                            ' Collision mode is not defined; append blank columns
                            addonColumns.Add(String.Empty)
                            addonColumns.Add(blankAdditionalReporterIonColumns)
                        Else
                            ' Collision mode is defined
                            addonColumns.Add(scanStatsEntry.CollisionMode)
                            addonColumns.Add(scanStatsEntry.ReporterIonData)

                            collisionModeCurrentScan = String.Copy(scanStatsEntry.CollisionMode)
                        End If

                    ElseIf SeparateByCollisionMode Then
                        If scanStatsEntry Is Nothing Then
                            collisionModeCurrentScan = String.Empty
                        Else
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

                    writers(outFileIndex).WriteLine(dataLine & ControlChars.Tab & FlattenList(addonColumns))
                    linesWritten(outFileIndex) += 1

                End While

            End Using

            ' Close the output files
            If Not writers Is Nothing Then
                For index = 0 To outputFileCount - 1
                    If Not writers(index) Is Nothing Then
                        writers(index).Close()
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

                outputFileHandles.Add(collisionMode, New StreamWriter(New FileStream(Path.Combine(mOutputDirectoryPath, outputFileName), FileMode.Create, FileAccess.Write)))
                outputFileHeaderWritten.Add(collisionMode, False)
            Next

            ' Create the DatasetMap file
            Using writer = New StreamWriter(New FileStream(Path.Combine(mOutputDirectoryPath, baseFileName & "_DatasetMap.txt"), FileMode.Create, FileAccess.Write))

                writer.WriteLine("DatasetID" & ControlChars.Tab & "DatasetName")

                For Each datasetMapping In datasetNameIdMap
                    writer.WriteLine(datasetMapping.Value & ControlChars.Tab & datasetMapping.Key)
                Next

            End Using

            ' Merge the files
            For Each processedDataset In mProcessedDatasets
                Dim datasetId As Integer = datasetNameIdMap(processedDataset.BaseName)

                For Each sourcefile In processedDataset.OutputFiles

                    Dim collisionMode = sourcefile.Key

                    Dim writer As StreamWriter = Nothing
                    If Not outputFileHandles.TryGetValue(collisionMode, writer) Then
                        Console.WriteLine("Warning: unrecognized collision mode; skipping " & sourcefile.Value)
                        Continue For
                    End If

                    If Not File.Exists(sourcefile.Value) Then
                        Console.WriteLine("Warning: input file not found; skipping " & sourcefile.Value)
                        Continue For
                    End If

                    Using reader = New StreamReader(New FileStream(sourcefile.Value, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        Dim linesRead = 0
                        While Not reader.EndOfStream

                            Dim dataLine = reader.ReadLine()
                            linesRead += 1

                            If linesRead = 1 Then
                                If outputFileHeaderWritten(collisionMode) Then
                                    ' skip this line
                                    Continue While
                                End If

                                writer.WriteLine("DatasetID" & ControlChars.Tab & dataLine)
                                outputFileHeaderWritten(collisionMode) = True

                            Else
                                writer.WriteLine(datasetId & ControlChars.Tab & dataLine)
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
    ''' <param name="outputDirectoryPath">Output directory path</param>
    ''' <param name="parameterFilePath">Parameter file path (Ignored)</param>
    ''' <param name="resetErrorCode">If true, reset the error code</param>
    ''' <returns>True if success, False if failure</returns>
    Public Overloads Overrides Function ProcessFile(inputFilePath As String, outputDirectoryPath As String, parameterFilePath As String, resetErrorCode As Boolean) As Boolean

        Dim success As Boolean
        Dim masicResultsDirectory As String

        If resetErrorCode Then
            SetLocalErrorCode(eResultsProcessorErrorCodes.NoError)
        End If

        If inputFilePath Is Nothing OrElse inputFilePath.Length = 0 Then
            ShowMessage("Input file name is empty")
            MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.InvalidInputFilePath)
            Return False
        End If

        ' Note that CleanupFilePaths() will update mOutputDirectoryPath, which is used by LogMessage()
        If Not CleanupFilePaths(inputFilePath, outputDirectoryPath) Then
            MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.FilePathError)
            Return False
        End If

        Dim fiInputFile As FileInfo
        fiInputFile = New FileInfo(inputFilePath)

        If String.IsNullOrWhiteSpace(mMASICResultsDirectoryPath) Then
            masicResultsDirectory = fiInputFile.DirectoryName
        Else
            masicResultsDirectory = String.Copy(mMASICResultsDirectoryPath)
        End If

        If MageResults Then
            success = ProcessMageExtractorFile(fiInputFile, masicResultsDirectory)
        Else
            success = ProcessSingleJobFile(fiInputFile, masicResultsDirectory)
        End If
        Return success

    End Function

    Private Function ProcessMageExtractorFile(fiInputFile As FileInfo, masicResultsDirectory As String) As Boolean

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
            Dim dctJobToDatasetMap As Dictionary(Of Integer, clsDatasetInfo)

            dctJobToDatasetMap = ReadMageMetadataFile(fiMetadataFile.FullName)
            If dctJobToDatasetMap Is Nothing OrElse dctJobToDatasetMap.Count = 0 Then
                ShowErrorMessage("Error: ReadMageMetadataFile returned an empty job mapping")
                Return False
            End If

            Dim headerLine As String
            Dim jobColumnIndex As Integer

            ' Open the Mage Extractor data file so that we can validate and cache the header row
            Using reader = New StreamReader(New FileStream(fiInputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                headerLine = reader.ReadLine()
                Dim lstColumns = headerLine.Split(ControlChars.Tab).ToList()
                jobColumnIndex = lstColumns.IndexOf("Job")
                If jobColumnIndex < 0 Then
                    ShowErrorMessage("Input file is not a valid Mage Extractor results file; it must contain a ""Job"" column: " & fiInputFile.FullName)
                    Return False
                End If
            End Using

            Dim scanStatsHeaders = GetScanStatsHeaders()
            Dim sicStatsHeaders = GetSICStatsHeaders()

            ' Populate blankAdditionalScanStatsColumns with tab characters based on the number of items in scanStatsHeaders
            Dim blankAdditionalScanStatsColumns = New String(ControlChars.Tab, scanStatsHeaders.Count - 1)

            Dim blankAdditionalSICColumns = New String(ControlChars.Tab, sicStatsHeaders.Count)

            Dim outputFileName = Path.GetFileNameWithoutExtension(fiInputFile.Name) & RESULTS_SUFFIX
            Dim outputFilePath = Path.Combine(mOutputDirectoryPath, outputFileName)

            Dim jobsSuccessfullyMerged = 0


            ' Initialize the output file
            Using writer = New StreamWriter(New FileStream(outputFilePath, FileMode.Create, FileAccess.Write, FileShare.Read))

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

                Dim masicDataLoaded = False
                Dim headerLineWritten = False
                Dim writeReporterIonStats As Boolean
                Dim reporterIonHeaders = String.Empty
                Dim blankAdditionalReporterIonColumns = String.Empty

                Do While mPHRPReader.MoveNext()

                    Dim psm As clsPSM = mPHRPReader.CurrentPSM

                    ' Parse out the job from the current line
                    Dim lstColumns = psm.DataLineText.Split(ControlChars.Tab).ToList()

                    Dim job As Integer = -1
                    If Not Integer.TryParse(lstColumns(jobColumnIndex), job) Then
                        ShowMessage("Warning: Job column does not contain a job number; skipping this entry: " & psm.DataLineText)
                        Continue Do
                    End If

                    If job <> lastJob Then

                        ' New job; read and cache the MASIC data
                        masicDataLoaded = False

                        Dim datasetInfo As clsDatasetInfo = Nothing

                        If Not dctJobToDatasetMap.TryGetValue(job, datasetInfo) Then
                            ShowErrorMessage("Error: Job " & job & " was not defined in the Metadata file; unable to determine the dataset")
                        Else

                            ' Look for the corresponding MASIC files in the input directory

                            Dim masicFiles = New clsMASICFileInfo()

                            Dim datasetNameAndDirectory = "for dataset " & datasetInfo.DatasetName & " in " & masicResultsDirectory
                            Dim success = FindMASICFiles(masicResultsDirectory, datasetInfo, masicFiles, datasetNameAndDirectory, job)

                            If success Then

                                ' Read and cache the MASIC data
                                dctScanStats = New Dictionary(Of Integer, clsScanStatsData)
                                dctSICStats = New Dictionary(Of Integer, clsSICStatsData)
                                reporterIonHeaders = String.Empty

                                masicDataLoaded = ReadMASICData(masicResultsDirectory, masicFiles, dctScanStats, dctSICStats, reporterIonHeaders)

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

                            Dim addonHeaderColumns = FlattenList(scanStatsHeaders) & ControlChars.Tab & FlattenList(sicStatsHeaders)
                            If reporterIonHeaders.Length > 0 Then
                                ' Append the reporter ion stats columns
                                addonHeaderColumns &= ControlChars.Tab & reporterIonHeaders
                                writeReporterIonStats = True
                            End If

                            writer.WriteLine(headerLine & ControlChars.Tab & addonHeaderColumns)

                            headerLineWritten = True
                        End If

                        ' Look for scanNumber in dctScanStats
                        Dim scanStatsEntry As clsScanStatsData = Nothing
                        Dim addonColumns = New List(Of String)

                        If Not dctScanStats.TryGetValue(psm.ScanNumber, scanStatsEntry) Then
                            ' Match not found; use the blank columns in blankAdditionalColumns
                            addonColumns.Add(blankAdditionalScanStatsColumns)
                        Else
                            addonColumns.Add(scanStatsEntry.ElutionTime)
                            addonColumns.Add(scanStatsEntry.ScanType)
                            addonColumns.Add(scanStatsEntry.TotalIonIntensity)
                            addonColumns.Add(scanStatsEntry.BasePeakIntensity)
                            addonColumns.Add(scanStatsEntry.BasePeakMZ)
                        End If

                        Dim sicStatsEntry As clsSICStatsData = Nothing
                        If Not dctSICStats.TryGetValue(psm.ScanNumber, sicStatsEntry) Then
                            ' Match not found; use the blank columns in blankAdditionalSICColumns
                            addonColumns.Add(blankAdditionalSICColumns)
                        Else
                            addonColumns.Add(sicStatsEntry.OptimalScanNumber)
                            addonColumns.Add(sicStatsEntry.PeakMaxIntensity)
                            addonColumns.Add(sicStatsEntry.PeakSignalToNoiseRatio)
                            addonColumns.Add(sicStatsEntry.FWHMInScans)
                            addonColumns.Add(sicStatsEntry.PeakArea)
                            addonColumns.Add(sicStatsEntry.ParentIonIntensity)
                            addonColumns.Add(sicStatsEntry.ParentIonMZ)
                            addonColumns.Add(sicStatsEntry.StatMomentsArea)
                        End If

                        If writeReporterIonStats Then
                            If scanStatsEntry Is Nothing OrElse String.IsNullOrWhiteSpace(scanStatsEntry.CollisionMode) Then
                                ' Collision mode is not defined; append blank columns
                                addonColumns.Add(String.Empty)
                                addonColumns.Add(blankAdditionalReporterIonColumns)
                            Else
                                ' Collision mode is defined
                                addonColumns.Add(scanStatsEntry.CollisionMode)
                                addonColumns.Add(scanStatsEntry.ReporterIonData)
                            End If
                        End If

                        writer.WriteLine(psm.DataLineText & ControlChars.Tab & FlattenList(addonColumns))

                    Else
                        Dim blankAddonColumns = ControlChars.Tab & blankAdditionalScanStatsColumns & ControlChars.Tab & blankAdditionalSICColumns

                        If writeReporterIonStats Then
                            writer.WriteLine(psm.DataLineText & blankAddonColumns & ControlChars.Tab & ControlChars.Tab & blankAdditionalReporterIonColumns)
                        Else
                            writer.WriteLine(psm.DataLineText & blankAddonColumns)
                        End If

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

    Private Function ProcessSingleJobFile(fiInputFile As FileSystemInfo, masicResultsDirectory As String) As Boolean

        Dim dctScanStats As Dictionary(Of Integer, clsScanStatsData)
        Dim dctSICStats As Dictionary(Of Integer, clsSICStatsData)

        Dim reporterIonHeaders As String = String.Empty

        Try
            Dim datasetName = Path.GetFileNameWithoutExtension(fiInputFile.FullName)

            Dim datasetInfo = New clsDatasetInfo(datasetName, 0)

            ' Note that FindMASICFiles will first try the full filename, and if it doesn't find a match,
            ' it will start removing text from the end of the filename by looking for underscores

            ' Look for the corresponding MASIC files in the input directory
            Dim masicFiles = New clsMASICFileInfo()

            Dim masicFileSearchInfo = " in " & masicResultsDirectory
            Dim success = FindMASICFiles(masicResultsDirectory, datasetInfo, masicFiles, masicFileSearchInfo, 0)

            If Not success Then
                SetLocalErrorCode(eResultsProcessorErrorCodes.MissingMASICFiles)
                Return False
            End If

            ' Read and cache the MASIC data
            dctScanStats = New Dictionary(Of Integer, clsScanStatsData)
            dctSICStats = New Dictionary(Of Integer, clsSICStatsData)

            success = ReadMASICData(masicResultsDirectory, masicFiles, dctScanStats, dctSICStats, reporterIonHeaders)

            If success Then
                ' Merge the MASIC data with the input file
                success = MergePeptideHitAndMASICFiles(fiInputFile, mOutputDirectoryPath,
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

            Return success

        Catch ex As Exception
            HandleException("Error in ProcessSingleJobFile", ex)
            Return False
        End Try


    End Function

    Private Function ReadMASICData(
      sourceDirectory As String,
      masicFiles As clsMASICFileInfo,
      dctScanStats As IDictionary(Of Integer, clsScanStatsData),
      dctSICStats As IDictionary(Of Integer, clsSICStatsData),
      <Out> ByRef reporterIonHeaders As String) As Boolean


        Try
            Dim scanStatsRead As Boolean
            Dim sicStatsRead As Boolean

            If String.IsNullOrWhiteSpace(masicFiles.ScanStatsFileName) Then
                scanStatsRead = False
            Else
                scanStatsRead = ReadScanStatsFile(sourceDirectory, masicFiles.ScanStatsFileName, dctScanStats)
            End If

            If String.IsNullOrWhiteSpace(masicFiles.SICStatsFileName) Then
                sicStatsRead = False
            Else
                sicStatsRead = ReadSICStatsFile(sourceDirectory, masicFiles.SICStatsFileName, dctSICStats)
            End If

            If String.IsNullOrWhiteSpace(masicFiles.ReporterIonsFileName) Then
                reporterIonHeaders = String.Empty
            Else
                ReadReporterIonStatsFile(sourceDirectory, masicFiles.ReporterIonsFileName, dctScanStats, reporterIonHeaders)
            End If

            Return scanStatsRead OrElse sicStatsRead

        Catch ex As Exception
            HandleException("Error in ReadMASICData", ex)
            reporterIonHeaders = String.Empty
            Return False
        End Try

    End Function

    Private Function ReadScanStatsFile(
      sourceDirectory As String,
      scanStatsFileName As String,
      dctScanStats As IDictionary(Of Integer, clsScanStatsData)) As Boolean

        Try
            ' Initialize dctScanStats
            dctScanStats.Clear()

            ShowMessage("  Reading: " & scanStatsFileName)

            Using reader = New StreamReader(New FileStream(Path.Combine(sourceDirectory, scanStatsFileName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

                While Not reader.EndOfStream
                    Dim dataLine = reader.ReadLine()

                    If String.IsNullOrWhiteSpace(dataLine) Then Continue While

                    Dim lineParts = dataLine.Split(ControlChars.Tab)

                    If lineParts.Length < eScanStatsColumns.BasePeakMZ + 1 Then Continue While

                    Dim scanNumber As Integer
                    If Not Integer.TryParse(lineParts(eScanStatsColumns.ScanNumber), scanNumber) Then Continue While

                    ' Note: the remaining values are stored as strings to prevent the number format from changing
                    Dim scanStatsEntry = New clsScanStatsData(scanNumber) With {
                        .ElutionTime = String.Copy(lineParts(eScanStatsColumns.ScanTime)),
                        .ScanType = String.Copy(lineParts(eScanStatsColumns.ScanType)),
                        .TotalIonIntensity = String.Copy(lineParts(eScanStatsColumns.TotalIonIntensity)),
                        .BasePeakIntensity = String.Copy(lineParts(eScanStatsColumns.BasePeakIntensity)),
                        .BasePeakMZ = String.Copy(lineParts(eScanStatsColumns.BasePeakMZ)),
                        .CollisionMode = String.Empty,
                        .ReporterIonData = String.Empty
                    }

                    dctScanStats.Add(scanNumber, scanStatsEntry)

                End While

            End Using

            Return True

        Catch ex As Exception
            HandleException("Error in ReadScanStatsFile", ex)
            Return False
        End Try

    End Function

    Private Function ReadMageMetadataFile(metadataFilePath As String) As Dictionary(Of Integer, clsDatasetInfo)

        Dim dctJobToDatasetMap = New Dictionary(Of Integer, clsDatasetInfo)
        Dim headersParsed As Boolean

        Dim jobIndex As Integer = -1
        Dim datasetIndex As Integer = -1
        Dim datasetIDIndex As Integer = -1

        Try
            Using reader = New StreamReader(New FileStream(metadataFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

                While Not reader.EndOfStream
                    Dim dataLine = reader.ReadLine()

                    If String.IsNullOrWhiteSpace(dataLine) Then Continue While

                    Dim lineParts = dataLine.Split(ControlChars.Tab).ToList()

                    If Not headersParsed Then
                        ' Look for the Job and Dataset columns
                        jobIndex = lineParts.IndexOf("Job")
                        datasetIndex = lineParts.IndexOf("Dataset")
                        datasetIDIndex = lineParts.IndexOf("Dataset_ID")

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

                    If lineParts.Count > datasetIndex Then
                        Dim jobNumber As Integer
                        Dim datasetID As Integer

                        If Integer.TryParse(lineParts(jobIndex), jobNumber) Then
                            If Integer.TryParse(lineParts(datasetIDIndex), datasetID) Then
                                Dim datasetName = lineParts(datasetIndex)

                                Dim datasetInfo = New clsDatasetInfo(datasetName, datasetID)

                                dctJobToDatasetMap.Add(jobNumber, datasetInfo)
                            Else
                                ShowMessage("Warning: Dataset_ID number not numeric in metadata file, line " & dataLine)
                            End If
                        Else
                            ShowMessage("Warning: Job number not numeric in metadata file, line " & dataLine)
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

    Private Function ReadSICStatsFile(sourceDirectory As String,
      sicStatsFileName As String,
      dctSICStats As IDictionary(Of Integer, clsSICStatsData)) As Boolean

        Dim fragScanNumber As Integer

        Try
            ' Initialize dctSICStats
            dctSICStats.Clear()

            ShowMessage("  Reading: " & sicStatsFileName)

            Using reader = New StreamReader(New FileStream(Path.Combine(sourceDirectory, sicStatsFileName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

                While Not reader.EndOfStream
                    Dim dataLine = reader.ReadLine()

                    If String.IsNullOrWhiteSpace(dataLine) Then Continue While

                    Dim lineParts = dataLine.Split(ControlChars.Tab)

                    If lineParts.Length >= eSICStatsColumns.StatMomentsArea + 1 AndAlso Integer.TryParse(lineParts(eSICStatsColumns.FragScanNumber), fragScanNumber) Then

                        ' Note: the remaining values are stored as strings to prevent the number format from changing
                        Dim sicStatsEntry = New clsSICStatsData(fragScanNumber) With {
                                .OptimalScanNumber = String.Copy(lineParts(eSICStatsColumns.OptimalPeakApexScanNumber)),
                                .PeakMaxIntensity = String.Copy(lineParts(eSICStatsColumns.PeakMaxIntensity)),
                                .PeakSignalToNoiseRatio = String.Copy(lineParts(eSICStatsColumns.PeakSignalToNoiseRatio)),
                                .FWHMInScans = String.Copy(lineParts(eSICStatsColumns.FWHMInScans)),
                                .PeakArea = String.Copy(lineParts(eSICStatsColumns.PeakArea)),
                                .ParentIonIntensity = String.Copy(lineParts(eSICStatsColumns.ParentIonIntensity)),
                                .ParentIonMZ = String.Copy(lineParts(eSICStatsColumns.MZ)),
                                .StatMomentsArea = String.Copy(lineParts(eSICStatsColumns.StatMomentsArea))
                                }

                        dctSICStats.Add(fragScanNumber, sicStatsEntry)
                    End If

                End While

            End Using

            Return True

        Catch ex As Exception
            HandleException("Error in ReadSICStatsFile", ex)
            Return False
        End Try

    End Function

    Private Function ReadReporterIonStatsFile(sourceDirectory As String,
      reporterIonStatsFileName As String,
      dctScanStats As IDictionary(Of Integer, clsScanStatsData),
      <Out> ByRef reporterIonHeaders As String) As Boolean

        Dim linesRead As Integer

        Dim warningCount = 0

        reporterIonHeaders = String.Empty

        Try

            ShowMessage("  Reading: " & reporterIonStatsFileName)

            Using reader = New StreamReader(New FileStream(Path.Combine(sourceDirectory, reporterIonStatsFileName), FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

                linesRead = 0
                While Not reader.EndOfStream
                    Dim dataLine = reader.ReadLine()

                    If String.IsNullOrWhiteSpace(dataLine) Then Continue While

                    linesRead += 1
                    Dim lineParts = dataLine.Split(ControlChars.Tab)

                    If linesRead = 1 Then
                        ' This is the header line; we need to cache it

                        If lineParts.Length >= eReporterIonStatsColumns.ReporterIonIntensityMax + 1 Then
                            reporterIonHeaders = lineParts(eReporterIonStatsColumns.CollisionMode)
                            reporterIonHeaders &= ControlChars.Tab & FlattenArray(lineParts, eReporterIonStatsColumns.ReporterIonIntensityMax)
                        Else
                            ' There aren't enough columns in the header line; this is unexpected
                            reporterIonHeaders = "Collision Mode" & ControlChars.Tab & "AdditionalReporterIonColumns"
                        End If

                    End If

                    If lineParts.Length < eReporterIonStatsColumns.ReporterIonIntensityMax + 1 Then Continue While

                    Dim scanNumber As Integer
                    If Not Integer.TryParse(lineParts(eReporterIonStatsColumns.ScanNumber), scanNumber) Then Continue While

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
                            scanStatsEntry.CollisionMode = String.Copy(lineParts(eReporterIonStatsColumns.CollisionMode))
                            scanStatsEntry.ReporterIonData = FlattenArray(lineParts, eReporterIonStatsColumns.ReporterIonIntensityMax)
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
                If MyBase.ErrorCode = ProcessFilesErrorCodes.LocalizedError Then
                    MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.NoError)
                End If
            Else
                MyBase.SetBaseClassErrorCode(ProcessFilesErrorCodes.LocalizedError)
            End If
        End If

    End Sub

    Private Function SummarizeCollisionModes(
      inputFile As FileSystemInfo,
      baseFileName As String,
      outputDirectoryPath As String,
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

            ' Try to load the collision mode info from the input file
            ' MS-GF+ results report this in the FragMethod column

            dctCollisionModeFileMap.Clear()
            collisionModeTypeCount = 0

            Try

                Using reader = New StreamReader(New FileStream(inputFile.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))

                    Dim linesRead = 0
                    Dim fragMethodColNumber = 0

                    While Not reader.EndOfStream
                        Dim dataLine = reader.ReadLine()

                        If String.IsNullOrWhiteSpace(dataLine) Then Continue While

                        linesRead += 1

                        Dim lineParts = dataLine.Split(ControlChars.Tab)
                        If linesRead = 1 Then
                            ' Header line; look for the FragMethod column
                            For colIndex = 0 To lineParts.Length - 1
                                If String.Equals(lineParts(colIndex), "FragMethod", StringComparison.OrdinalIgnoreCase) Then
                                    fragMethodColNumber = colIndex + 1
                                    Exit For
                                End If
                            Next

                            If fragMethodColNumber = 0 Then
                                ' Fragmentation method column not found
                                ShowWarning("Unable to determine the collision mode for results being merged. " &
                                            "This is typically obtained from a MASIC _ReporterIons.txt file " &
                                            "or from the FragMethod column in the MS-GF+ results file")
                                Exit While
                            End If

                            ' Also look for the scan number column, auto-updating ScanNumberColumn if necessary
                            FindScanNumColumn(inputFile, lineParts)

                            Continue While
                        End If

                        Dim scanNumber As Integer
                        If lineParts.Length < ScanNumberColumn OrElse Not Integer.TryParse(lineParts(ScanNumberColumn - 1), scanNumber) Then
                            Continue While
                        End If

                        If lineParts.Length < fragMethodColNumber Then
                            Continue While
                        End If

                        Dim collisionMode = lineParts(fragMethodColNumber - 1)

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

                    End While
                End Using

            Catch ex As Exception
                HandleException("Error extraction collision mode information from the input file", ex)
                Return New KeyValuePair(Of String, String)(collisionModeTypeCount - 1) {}
            End Try

        End If

        If collisionModeTypeCount = 0 Then collisionModeTypeCount = 1

        Dim outputFilePaths = New KeyValuePair(Of String, String)(collisionModeTypeCount - 1) {}

        If dctCollisionModeFileMap.Count = 0 Then
            outputFilePaths(0) = New KeyValuePair(Of String, String)("na", Path.Combine(outputDirectoryPath, baseFileName & "_na" & RESULTS_SUFFIX))
        Else
            For Each oItem In dctCollisionModeFileMap
                Dim collisionMode As String = oItem.Key
                If String.IsNullOrWhiteSpace(collisionMode) Then collisionMode = "na"
                outputFilePaths(oItem.Value) = New KeyValuePair(Of String, String)(
                    collisionMode, Path.Combine(outputDirectoryPath, baseFileName & "_" & collisionMode & RESULTS_SUFFIX))
            Next
        End If

        Return outputFilePaths

    End Function

End Class
