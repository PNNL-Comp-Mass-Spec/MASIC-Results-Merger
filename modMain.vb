Option Strict On

' This program merges the contents of a tab-delimited peptide hit results file
' (e.g. from Sequest, XTandem, or MSGF+) with the corresponding MASIC results files, 
' appending the relevant MASIC stats for each peptide hit result
'
' -------------------------------------------------------------------------------
' Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
' Program started November 26, 2008
'
' E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com
' Website: http://omics.pnl.gov/ or http://www.sysbio.org/resources/staff/ or http://panomics.pnnl.gov/
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

Module modMain
	Public Const PROGRAM_DATE As String = "August 20, 2014"

	Private mInputFilePath As String
	Private mMageResults As Boolean
	Private mMergeWildcardResults As Boolean

	Private mMASICResultsFolderPath As String					' Optional
	Private mOutputFolderPath As String				' Optional
	Private mParameterFilePath As String						' Optional

	Private mOutputFolderAlternatePath As String				' Optional
	Private mRecreateFolderHierarchyInAlternatePath As Boolean	' Optional

	Private mRecurseFolders As Boolean
	Private mRecurseFoldersMaxLevels As Integer

	Private mLogMessagesToFile As Boolean
	Private mQuietMode As Boolean

	Private mScanNumberColumn As Integer
	Private mSeparateByCollisionMode As Boolean

	Private WithEvents mMASICResultsMerger As clsMASICResultsMerger
	Private mLastProgressReportTime As DateTime
	Private mLastProgressReportValue As Integer

	Private Sub DisplayProgressPercent(ByVal intPercentComplete As Integer, ByVal blnAddCarriageReturn As Boolean)
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
		mParameterFilePath = String.Empty

		mRecurseFolders = False
		mRecurseFoldersMaxLevels = 0

		mQuietMode = False
		mLogMessagesToFile = False

		mScanNumberColumn = clsMASICResultsMerger.DEFAULT_SCAN_NUMBER_COLUMN
		mSeparateByCollisionMode = False

		Try
			blnProceed = False
			If objParseCommandLine.ParseCommandLine Then
				If SetOptionsUsingCommandLineParameters(objParseCommandLine) Then blnProceed = True
			End If

			If Not blnProceed OrElse _
			   objParseCommandLine.NeedToShowHelp OrElse _
			   objParseCommandLine.ParameterCount + objParseCommandLine.NonSwitchParameterCount = 0 OrElse _
			   mInputFilePath.Length = 0 Then
				ShowProgramHelp()
				intReturnCode = -1
			Else
				mMASICResultsMerger = New clsMASICResultsMerger

				With mMASICResultsMerger
					.ShowMessages = Not mQuietMode
					.LogMessagesToFile = mLogMessagesToFile

					' Note: Define other options here; they will get overridden if defined in the parameter file
					.MASICResultsFolderPath = mMASICResultsFolderPath
					.ScanNumberColumn = mScanNumberColumn
					.SeparateByCollisionMode = mSeparateByCollisionMode

					.MageResults = mMageResults					
				End With

				If mRecurseFolders Then
					If mMASICResultsMerger.ProcessFilesAndRecurseFolders(mInputFilePath, mOutputFolderPath, mOutputFolderAlternatePath, mRecreateFolderHierarchyInAlternatePath, mParameterFilePath, mRecurseFoldersMaxLevels) Then
						intReturnCode = 0
					Else
						intReturnCode = mMASICResultsMerger.ErrorCode
					End If
				Else
					If mMASICResultsMerger.ProcessFilesWildcard(mInputFilePath, mOutputFolderPath, mParameterFilePath) Then
						intReturnCode = 0
					Else
						intReturnCode = mMASICResultsMerger.ErrorCode
						If intReturnCode <> 0 AndAlso Not mQuietMode Then
							Console.WriteLine("Error while processing: " & mMASICResultsMerger.GetErrorMessage())
						End If
					End If
				End If

				If mMergeWildcardResults AndAlso mMASICResultsMerger.ProcessedDatasets.Count > 0 Then
					mMASICResultsMerger.MergeProcessedDatasets()
				End If

				DisplayProgressPercent(mLastProgressReportValue, True)
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

	Private Function SetOptionsUsingCommandLineParameters(ByVal objParseCommandLine As clsParseCommandLine) As Boolean
		' Returns True if no problems; otherwise, returns false

		Dim strValue As String = String.Empty
		Dim lstValidParameters As List(Of String) = New List(Of String) From {"I", "M", "O", "P", "N", "C", "Mage", "Append", "S", "A", "R", "Q"}
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
					If .RetrieveValueForParameter("P", strValue) Then mParameterFilePath = strValue

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

					If .RetrieveValueForParameter("Q", strValue) Then mQuietMode = True
				End With

				Return True
			End If

		Catch ex As Exception
			ShowErrorMessage("Error parsing the command line parameters: " & Environment.NewLine & ex.Message)
		End Try

		Return False

	End Function

	Private Sub ShowErrorMessage(ByVal strMessage As String)
		Const strSeparator As String = "------------------------------------------------------------------------------"

		Console.WriteLine()
		Console.WriteLine(strSeparator)
		Console.WriteLine(strMessage)
		Console.WriteLine(strSeparator)
		Console.WriteLine()

		WriteToErrorStream(strMessage)
	End Sub

	Private Sub ShowErrorMessage(ByVal strTitle As String, ByVal items As IEnumerable(Of String))
		Const strSeparator As String = "------------------------------------------------------------------------------"
		Dim strMessage As String

		Console.WriteLine()
		Console.WriteLine(strSeparator)
		Console.WriteLine(strTitle)
		strMessage = strTitle & ":"

		For Each item As String In items
			Console.WriteLine("   " + item)
			strMessage &= " " & item
		Next
		Console.WriteLine(strSeparator)
		Console.WriteLine()

		WriteToErrorStream(strMessage)
	End Sub

	Private Sub ShowProgramHelp()

		Try

			Console.WriteLine("This program merges the contents of a tab-delimited peptide hit results file (e.g. from Sequest, XTandem, or MSGF+) with the corresponding MASIC results files, appending the relevant MASIC stats for each peptide hit result.")
			Console.WriteLine()
			Console.WriteLine("Program syntax:" & ControlChars.NewLine & IO.Path.GetFileName(Reflection.Assembly.GetExecutingAssembly().Location) & _
			  " InputFilePathSpec [/M:MASICResultsFolderPath] [/O:OutputFolderPath]")
			Console.WriteLine(" [/P:ParameterFilePath]")
			Console.WriteLine(" [/N:ScanNumberColumn] [/C] [/Mage] [/Append]")
			Console.WriteLine(" [/S:[MaxLevel]] [/A:AlternateOutputFolderPath] [/R] [/Q]")
			Console.WriteLine()
			Console.WriteLine("The input file should be a tab-delimited file with scan number in the second column (e.g. Sequest Synopsis or First-Hits file (_syn.txt or _fht.txt), XTandem _xt.txt file, MSGF+ syn/fht file (_msgfdb_syn.txt or _msgfdb_fht.txt), or Inspect syn/fht file (_inspect_syn.txt or _inspect_fht.txt). " & _
			  "If the MASIC result files are not in the same folder as the input file, then use /M to define the path to the correct folder. " & _
			  "The output folder switch is optional.  If omitted, the output file will be created in the same folder as the input file. " & _
			  "The parameter file path is optional.  If included, it should point to a valid XML parameter file.")

			Console.WriteLine()
			Console.WriteLine("Use /N to change the column number that contains scan number in the input file.  The default is 2 (meaning /N:2)." & _
			  "When reading data with _ReporterIons.txt files, you can use /C to specify that a separate output file be created for each collision mode type in the input file (typically pqd, cid, and etd).")
			Console.WriteLine()
			Console.WriteLine("Use /Mage to specify that the input file is a results from from Mage Extractor.  This file will contain results from several analysis jobs; the first column in this file must be Job and the remaining columns must be the standard Synopsis or First-Hits columns supported by PHRPReader.  In addition, the input folder must have a column named InputFile_metadata.txt (this file was auto-created by Mage Extractor).")
			Console.WriteLine()
			Console.WriteLine("Use /Append to merge results from multiple datasets together as a single file; this is only applicable when the InputFilePathSpec includes a * wildcard and multiple files are matched")
			Console.WriteLine("The merged results file will have DatasetID values of 1, 2, 3, etc. along with a second file mapping DatasetID to Dataset Name")
			Console.WriteLine()
			Console.WriteLine("Use /S to process all valid files in the input folder and subfolders. Include a number after /S (like /S:2) to limit the level of subfolders to examine." & _
			  "When using /S, you can redirect the output of the results using /A." & _
			  "When using /S, you can use /R to re-create the input folder hierarchy in the alternate output folder (if defined)." & _
			  "The optional /Q switch will suppress all error messages.")
			Console.WriteLine()

			Console.WriteLine("Program written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA) in 2008")
			Console.WriteLine("Version: " & GetAppVersion())
			Console.WriteLine()

			Console.WriteLine("E-mail: matthew.monroe@pnnl.gov or matt@alchemistmatt.com")
			Console.WriteLine("Website: http://omics.pnl.gov/ or http://panomics.pnnl.gov/")
			Console.WriteLine()
			' Delay for 750 msec in case the user double clicked this file from within Windows Explorer (or started the program via a shortcut)
			Threading.Thread.Sleep(750)

		Catch ex As Exception
			ShowErrorMessage("Error displaying the program syntax: " & ex.Message)
		End Try

	End Sub

	Private Sub WriteToErrorStream(strErrorMessage As String)
		Try
			Using swErrorStream As IO.StreamWriter = New IO.StreamWriter(Console.OpenStandardError())
				swErrorStream.WriteLine(strErrorMessage)
			End Using
		Catch ex As Exception
			' Ignore errors here
		End Try
	End Sub

	Private Sub mMASICResultsMerger_ProgressChanged(ByVal taskDescription As String, ByVal percentComplete As Single) Handles mMASICResultsMerger.ProgressChanged
		Const PERCENT_REPORT_INTERVAL As Integer = 25
		Const PROGRESS_DOT_INTERVAL_MSEC As Integer = 250

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

	Private Sub mMASICResultsMerger_ProgressReset() Handles mMASICResultsMerger.ProgressReset
		mLastProgressReportTime = DateTime.UtcNow
		mLastProgressReportValue = 0
	End Sub
End Module
