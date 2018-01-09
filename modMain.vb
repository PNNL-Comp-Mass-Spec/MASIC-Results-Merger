Option Strict On

Imports PRISM
' This program merges the contents of a tab-delimited peptide hit results file
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

Module modMain
    Public Const PROGRAM_DATE As String = "January 8, 2018"

    Private mInputFilePath As String
    Private mMageResults As Boolean
    Private mMergeWildcardResults As Boolean

    Private mMASICResultsFolderPath As String                   ' Optional
    Private mOutputFolderPath As String                         ' Optional

    Private mOutputFolderAlternatePath As String                ' Optional
    Private mRecreateFolderHierarchyInAlternatePath As Boolean  ' Optional

    Private mRecurseFolders As Boolean
    Private mRecurseFoldersMaxLevels As Integer

    Private mLogMessagesToFile As Boolean

    Private mScanNumberColumn As Integer
    Private mSeparateByCollisionMode As Boolean

    Private mMASICResultsMerger As clsMASICResultsMerger
    Private mLastProgressReportTime As DateTime
    Private mLastProgressReportValue As Integer

    Private Sub DisplayProgressPercent(intPercentComplete As Integer, blnAddCarriageReturn As Boolean)
        If blnAddCarriageReturn Then
            Console.WriteLine()
        End If

        If intPercentComplete > 100 Then intPercentComplete = 100

        Console.Write("Processing: " & intPercentComplete.ToString & "% ")
        If blnAddCarriageReturn Then
            Console.WriteLine()
        End If
    End Sub

    Public Function Main() As Integer
        ' Returns 0 if no error, error code if an error

        Dim intReturnCode As Integer
        Dim objParseCommandLine As New clsParseCommandLine
        Dim blnProceed As Boolean

        mInputFilePath = String.Empty
        mMageResults = False
        mMergeWildcardResults = False

        mMASICResultsFolderPath = String.Empty
        mOutputFolderPath = String.Empty

        mRecurseFolders = False
        mRecurseFoldersMaxLevels = 0

        mLogMessagesToFile = False

        mScanNumberColumn = clsMASICResultsMerger.DEFAULT_SCAN_NUMBER_COLUMN
        mSeparateByCollisionMode = False

        Try
            blnProceed = False
            If objParseCommandLine.ParseCommandLine Then
                If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
            End If

            If Not blnProceed OrElse
               objParseCommandLine.NeedToShowHelp OrElse
               objParseCommandLine.ParameterCount + objParseCommandLine.NonSwitchParameterCount = 0 OrElse
               mInputFilePath.Length = 0 Then
                ShowProgramHelp()
                intReturnCode = -1
            Else
                mMASICResultsMerger = New clsMASICResultsMerger()
                AddHandler mMASICResultsMerger.ErrorEvent, AddressOf mMASICResultsMerger_ErrorEvent
                AddHandler mMASICResultsMerger.WarningEvent, AddressOf mMASICResultsMerger_WarningEvent
                AddHandler mMASICResultsMerger.StatusEvent, AddressOf mMASICResultsMerger_StatusEvent
                AddHandler mMASICResultsMerger.DebugEvent, AddressOf mMASICResultsMerger_DebugEvent
                AddHandler mMASICResultsMerger.ProgressUpdate, AddressOf mMASICResultsMerger_ProgressUpdate
                AddHandler mMASICResultsMerger.ProgressReset, AddressOf mMASICResultsMerger_ProgressReset

                With mMASICResultsMerger
                    .LogMessagesToFile = mLogMessagesToFile

                    ' Note: Define other options here; they will get overridden if defined in the parameter file
                    .MASICResultsFolderPath = mMASICResultsFolderPath
                    .ScanNumberColumn = mScanNumberColumn
                    .SeparateByCollisionMode = mSeparateByCollisionMode

                    .MageResults = mMageResults
                End With

                If mRecurseFolders Then
                    If mMASICResultsMerger.ProcessFilesAndRecurseFolders(mInputFilePath, mOutputFolderPath, mOutputFolderAlternatePath, mRecreateFolderHierarchyInAlternatePath, "", mRecurseFoldersMaxLevels) Then
                        intReturnCode = 0
                    Else
                        intReturnCode = mMASICResultsMerger.ErrorCode
                    End If
                Else
                    If mMASICResultsMerger.ProcessFilesWildcard(mInputFilePath, mOutputFolderPath) Then
                        intReturnCode = 0
                    Else
                        intReturnCode = mMASICResultsMerger.ErrorCode
                        If intReturnCode <> 0 Then
                            ShowErrorMessage("Error while processing: " & mMASICResultsMerger.GetErrorMessage())
                        End If
                    End If
                End If

                If mMergeWildcardResults AndAlso mMASICResultsMerger.ProcessedDatasets.Count > 0 Then
                    mMASICResultsMerger.MergeProcessedDatasets()
                End If

                If mLastProgressReportValue > 0 Then
                    DisplayProgressPercent(mLastProgressReportValue, True)
                End If

            End If

        Catch ex As Exception
            ShowErrorMessage("Error occurred in modMain->Main: " & Environment.NewLine & ex.Message)
            intReturnCode = -1
        End Try

        Return intReturnCode

    End Function

    Private Function GetAppVersion() As String
        Return Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString() & " (" & PROGRAM_DATE & ")"
    End Function

    Private Function SetOptionsUsingCommandLineParameters(objParseCommandLine As clsParseCommandLine) As Boolean
        ' Returns True if no problems; otherwise, returns false

        Dim strValue As String = String.Empty
        Dim lstValidParameters = New List(Of String) From {"I", "M", "O", "N", "C", "Mage", "Append", "S", "A", "R"}
        Dim intValue As Integer

        Try
            ' Make sure no invalid parameters are present
            If objParseCommandLine.InvalidParametersPresent(lstValidParameters) Then
                ShowErrorMessage("Invalid commmand line parameters",
                  (From item In objParseCommandLine.InvalidParameters(lstValidParameters) Select "/" + item).ToList())
                Return False
            Else
                With objParseCommandLine
                    ' Query objParseCommandLine to see if various parameters are present
                    If .RetrieveValueForParameter("I", strValue) Then
                        mInputFilePath = strValue
                    ElseIf .NonSwitchParameterCount > 0 Then
                        mInputFilePath = .RetrieveNonSwitchParameter(0)
                    End If

                    If .RetrieveValueForParameter("M", strValue) Then mMASICResultsFolderPath = strValue
                    If .RetrieveValueForParameter("O", strValue) Then mOutputFolderPath = strValue

                    If .RetrieveValueForParameter("N", strValue) Then
                        If IsNumeric(strValue) Then
                            mScanNumberColumn = CInt(strValue)
                        End If
                    End If

                    If .IsParameterPresent("C") Then mSeparateByCollisionMode = True
                    If .IsParameterPresent("Mage") Then mMageResults = True

                    If .IsParameterPresent("Append") Then mMergeWildcardResults = True

                    If .RetrieveValueForParameter("S", strValue) Then
                        mRecurseFolders = True
                        If Integer.TryParse(strValue, intValue) Then
                            mRecurseFoldersMaxLevels = intValue
                        End If
                    End If
                    If .RetrieveValueForParameter("A", strValue) Then mOutputFolderAlternatePath = strValue
                    If .RetrieveValueForParameter("R", strValue) Then mRecreateFolderHierarchyInAlternatePath = True
                End With

                Return True
            End If

        Catch ex As Exception
            ShowErrorMessage("Error parsing the command line parameters: " & Environment.NewLine & ex.Message)
        End Try

        Return False

    End Function

    Private Sub ShowErrorMessage(message As String)
        ConsoleMsgUtils.ShowError(message)
    End Sub

    Private Sub ShowErrorMessage(title As String, errorMessages As IEnumerable(Of String))
        ConsoleMsgUtils.ShowErrors(title, errorMessages)
    End Sub

    Private Sub ShowProgramHelp()

        Try

            Console.WriteLine("This program merges the contents of a tab-delimited peptide hit results file " &
                              "(e.g. from X!Tandem or MSGF+) with the corresponding MASIC results files, " &
                              "appending the relevant MASIC stats for each peptide hit result, " &
                              "writing the merged data to a new tab-delimited text file.")
            Console.WriteLine()
            Console.WriteLine("It also supports TSV files, e.g. as created by the MzidToTsvConverter")
            Console.WriteLine()
            Console.WriteLine("If the input directory includes a MASIC _ReporterIons.txt file, " &
                              "the reporter ion intensities will also be included in the new text file.")
            Console.WriteLine()
            Console.WriteLine("Program syntax:" & ControlChars.NewLine & IO.Path.GetFileName(Reflection.Assembly.GetExecutingAssembly().Location) &
              " InputFilePathSpec [/M:MASICResultsFolderPath] [/O:OutputFolderPath]")
            Console.WriteLine(" [/N:ScanNumberColumn] [/C] [/Mage] [/Append]")
            Console.WriteLine(" [/S:[MaxLevel]] [/A:AlternateOutputFolderPath] [/R]")
            Console.WriteLine()
            Console.WriteLine("The input file should be a tab-delimited file where one column has scan numbers. " &
                              "By default, this program assumes the second column has scan number, but the " &
                              "/N switch can be used to change this (see below).")
            Console.WriteLine()
            Console.WriteLine("Common input files are:")
            Console.WriteLine("- Peptide Hit Results Processor (https://github.com/PNNL-Comp-Mass-Spec/PHRP) tab-delimited files")
            Console.WriteLine("  - MSGF+ syn/fht file (_msgfplus_syn.txt or _msgfplus_fht.txt)")
            Console.WriteLine("  - Sequest Synopsis or First-Hits file (_syn.txt or _fht.txt)")
            Console.WriteLine("  - XTandem _xt.txt file")
            Console.WriteLine("- MzidToTSVConverter (https://github.com/PNNL-Comp-Mass-Spec/Mzid-To-Tsv-Converter) .TSV files")
            Console.WriteLine("  - This is a tab-delimited text file created from a .mzid file (e.g. from MSGF+)")
            Console.WriteLine()
            Console.WriteLine("If the MASIC result files are not in the same folder as the input file, use /M to define the path to the correct folder.")
            Console.WriteLine("The output folder switch is optional.  If omitted, the output file will be created in the same folder as the input file. ")
            Console.WriteLine("")

            Console.WriteLine("Use /N to change the column number that contains scan number in the input file. " &
                              "The default is 2 (meaning /N:2). ")
            Console.WriteLine()
            Console.WriteLine("When reading data with _ReporterIons.txt files, you can use /C to specify " &
                              "that a separate output file be created for each collision mode type " &
                              "in the input file (typically pqd, cid, and etd).")
            Console.WriteLine()
            Console.WriteLine("Use /Mage to specify that the input file is a results file from Mage Extractor. " &
                              "This file will contain results from several analysis jobs; the first column " &
                              "in this file must be Job and the remaining columns must be the standard " &
                              "Synopsis or First-Hits columns supported by PHRPReader. " &
                              "In addition, the input folder must have a file named InputFile_metadata.txt " &
                              "(this file will have been auto-created by Mage Extractor).")
            Console.WriteLine()
            Console.WriteLine("Use /Append to merge results from multiple datasets together as a single file; " &
                              "this is only applicable when the InputFilePathSpec includes a * wildcard and multiple files are matched.")
            Console.WriteLine("The merged results file will have DatasetID values of 1, 2, 3, etc. " &
                              "along with a second file mapping DatasetID to Dataset Name")
            Console.WriteLine()
            Console.WriteLine("Use /S to process all valid files in the input folder and subfolders. " &
                              "Include a number after /S (like /S:2) to limit the level of subfolders to examine." &
                              "When using /S, you can redirect the output of the results using /A to specify an alternate output folder." &
                              "When using /S, you can use /R to re-create the input folder hierarchy in the alternate output folder (if defined).")
            Console.WriteLine()

            Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2008; updated in 2017")
            Console.WriteLine("Version: " & GetAppVersion())
            Console.WriteLine()

            Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com")
            Console.WriteLine("Website: https://omics.pnl.gov/ or https://panomics.pnnl.gov/")
            Console.WriteLine()

            ' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
            Threading.Thread.Sleep(750)

        Catch ex As Exception
            ShowErrorMessage("Error displaying the program syntax: " & ex.Message)
        End Try

    End Sub

    Private Sub mMASICResultsMerger_ErrorEvent(message As String, ex As Exception)
        ConsoleMsgUtils.ShowError(message, ex)
    End Sub

    Private Sub mMASICResultsMerger_WarningEvent(message As String)
        ConsoleMsgUtils.ShowWarning(message)
    End Sub

    Private Sub mMASICResultsMerger_StatusEvent(message As String)
        Console.WriteLine(message)
    End Sub

    Private Sub mMASICResultsMerger_DebugEvent(message As String)
        ConsoleMsgUtils.ShowDebug(message)
    End Sub

    Private Sub mMASICResultsMerger_ProgressUpdate(taskDescription As String, percentComplete As Single)
        Const PERCENT_REPORT_INTERVAL = 25
        Const PROGRESS_DOT_INTERVAL_MSEC = 250

        If percentComplete >= mLastProgressReportValue Then
            If mMageResults Then
                If mLastProgressReportValue > 0 AndAlso mLastProgressReportValue < 100 Then
                    Console.WriteLine()
                    DisplayProgressPercent(mLastProgressReportValue, False)
                    Console.WriteLine()
                End If
            Else
                If mLastProgressReportValue > 0 Then
                    Console.WriteLine()
                End If
                DisplayProgressPercent(mLastProgressReportValue, False)
            End If

            mLastProgressReportValue += PERCENT_REPORT_INTERVAL
            mLastProgressReportTime = DateTime.UtcNow
        Else
            If DateTime.UtcNow.Subtract(mLastProgressReportTime).TotalMilliseconds > PROGRESS_DOT_INTERVAL_MSEC Then
                mLastProgressReportTime = DateTime.UtcNow
                If Not mMageResults Then
                    Console.Write(".")
                End If
            End If
        End If
    End Sub

    Private Sub mMASICResultsMerger_ProgressReset()
        mLastProgressReportTime = DateTime.UtcNow
        mLastProgressReportValue = 0
    End Sub
End Module
