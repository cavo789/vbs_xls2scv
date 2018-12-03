' -----------------------------------------------------------------
'
' Author: Christophe Avonture
' Date	: November 2018
'
' Get the list of files in a [Input] folder (suppose these files are
' Excel files -there is no control-), open files one by one and, export
' each visible sheet to CSV and store them in a [Ouput] folder.
'
' @src https://github.com/cavo789/vbs_xls2scv
'
' -----------------------------------------------------------------

Option Explicit

Class clsMSExcel

	Private oApplication
	Private sFileName
	Private bVerbose, bEnableEvents, bDisplayAlerts

	Private bAppHasBeenStarted

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo
	End Property

	Public Property Let EnableEvents(bYesNo)
		bEnableEvents = bYesNo

		If Not (oApplication Is Nothing) Then
			oApplication.EnableEvents = bYesNo
		End if
	End Property

	Public Property Let DisplayAlerts(bYesNo)
		bDisplayAlerts = bYesNo

		If Not (oApplication Is Nothing) Then
			oApplication.DisplayAlerts = bYesNo
		End if

	End Property

	Public Property Let FileName(ByVal sName)
		sFileName = sName
	End Property

	Public Property Get FileName
		FileName = sFileName
	End Property

	' Make oApplication accessible
	Public Property Get app
		Set app = oApplication
	End Property

	Private Sub Class_Initialize()
		bVerbose = False
		bAppHasBeenStarted = False
		bEnableEvents = False
		bDisplayAlerts = False
		Set oApplication = Nothing
	End Sub

	Private Sub Class_Terminate()
		Set oApplication = Nothing
	End Sub

	' --------------------------------------------------------
	' Initialize the oApplication object variable : get a pointer
	' to the current Excel.exe app if already in memory or start
	' a new instance.
	'
	' If a new instance has been started, initialize the variable
	' bAppHasBeenStarted to True so the rest of the script knows
	' that Excel should then be closed by the script.
	' --------------------------------------------------------
	Public Function Instantiate()

		If (oApplication Is Nothing) Then

			On error Resume Next

			Set oApplication = GetObject(,"Excel.Application")

			If (Err.number <> 0) or (oApplication Is Nothing) Then
				Set oApplication = CreateObject("Excel.Application")
				' Remember that Excel has been started by
				' this script ==> should be released
				bAppHasBeenStarted = True
			End If

			oApplication.EnableEvents = bEnableEvents
			oApplication.DisplayAlerts = bDisplayAlerts

			Err.clear

			On error Goto 0

		End If

		' Return True if the application was created right
		' now
		Instantiate = bAppHasBeenStarted

	End Function

	' --------------------------------------------------------
	' Be sure Excel is visible
	' --------------------------------------------------------
	Public Sub MakeVisible

		Dim objShell

		If Not (oApplication Is Nothing) Then

			With oApplication

				.Application.ScreenUpdating = True
				.Application.Visible = True
				.Application.DisplayFullScreen = False

				.WindowState = -4137 ' xlMaximized

			End With

			Set objShell = CreateObject("WScript.Shell")
			objShell.appActivate oApplication.Caption
			Set objShell = Nothing

		End If

	End Sub

	Public Sub Quit()
		If not (oApplication Is Nothing) Then
			oApplication.Quit
		End If
	End Sub

	' --------------------------------------------------------
	' Open a standard Excel file and allow to specify if the
	' file should be opened in a read-only mode or not
	' --------------------------------------------------------
	Public Sub Open(bReadOnly)

		If not (oApplication Is nothing) Then

			If bVerbose Then
				wScript.echo "Open " & sFileName & _
					" (clsMSExcel::Open)"
			End If

			' False = UpdateLinks
			oApplication.Workbooks.Open sFileName, False, _
				bReadOnly

		End If

	End sub

	' --------------------------------------------------------
	' Close the active workbook
	' --------------------------------------------------------
	Public Sub CloseFile(sFileName)

		Dim wb
		Dim I
		Dim objFSO
		Dim sBaseName

		If Not (oApplication Is Nothing) Then

			Set objFSO = CreateObject("Scripting.FileSystemObject")

			If (sFileName = "") Then
				If Not (oApplication.ActiveWorkbook Is Nothing) Then
					sFileName = oApplication.ActiveWorkbook.FullName
				End If
			End If

			If (sFileName <> "") Then

				If bVerbose Then
					wScript.echo "Close " & sFileName & _
						" (clsMSExcel::CloseFile)"
				End if

				' Only the basename and not the full path
				sBaseName = objFSO.GetFileName(sFileName)

				On Error Resume Next
				Set wb = oApplication.Workbooks(sBaseName)
				If Not (err.number = 0) Then
					' Not found, workbook not loaded
					Set wb = Nothing
				Else
					If bVerbose Then
						wScript.echo "	Closing " & sBaseName & _
							" (clsMSExcel::CloseFile)"
					End if
					' Close without saving
					wb.Close False
				End if

				On Error Goto 0

			End If

			Set objFSO = Nothing

		End If

	End Sub

	' --------------------------------------------------------
	' Process each sheets and export as CSV
	' --------------------------------------------------------
	Public Sub ExportSheetsToCSV(sOutputFolder)

		Dim sh
		Dim sFileName

		With oApplication.ActiveWorkbook

			For each sh in .Worksheets

				' Process only visible sheets
				If (sh.Visible = true) Then 

					wScript.Echo "   Process sheet " & sh.Name 

					sh.Activate

					' Force formula's calculation
					sh.Calculate

					' Copy/Paste values
					sh.Cells.Copy

					' -4163 = xlPasteValues
					sh.Cells.PasteSpecial -4163 

					' No space in the filename
					sFileName = Replace(sh.Name, " ", "_")
					
					' 6 = xlCSV
					.SaveAs sOutputFolder & sFileName & ".csv", 6

				End If

			Next

		End With

	End Sub

End Class

Dim objFSO, objFolder, objFiles, objFile
Dim cMSExcel
Dim strFileName, sInputFolder, sOutputFolder

If Wscript.Arguments.count <> 2 Then 

	wScript.Echo "Usage: xls2csv.vbs [input_folder] [output_folder]"
	wScript.Echo ""
	wScript.Echo "Scan the [input_folder] and process every .xlsx files, "
	wScript.Echo "open every workbooks and export each visible worksheet to "
	wScript.Echo "a .csv file, stored in [output_folder]"
	wScript.Echo ""
	wScript.Echo "Note: [input_folder] and [output_folder] should be absolute "
	wScript.Echo "paths like C:\Christophe\Files\Data and not just Data."

	' And quit
	wScript.Quit 0
	
Else
	sInputFolder = Trim(Wscript.Arguments(0))
	sOutputFolder = Trim(Wscript.Arguments(1))

	If Right(sInputFolder, 1) <> "\" Then
		sInputFolder = sInputFolder & "\"
	End If

	If Right(sOutputFolder, 1) <> "\" Then
		sOutputFolder = sOutputFolder & "\"
	End If
	
End If 

' -----------------------------
' Start

wScript.Echo "Process folder: " & sInputFolder
wScript.Echo "   and export visible sheets as .csv, stored in: " & sOutputFolder
wScript.Echo " "

Set objFSO = CreateObject("Scripting.FileSystemObject")

On Error Resume Next

Set objFolder = objFSO.GetFolder(sInputFolder)

If Err.Number <> 0 Then

	wScript.Echo "(" & Err.Number & ") " & Err.Description
	wScript.Echo ""
	wScript.Echo "The path you entered is invalid, please " & _
		"select specify an existing folder."
		
	Wscript.Quit Err.Number
	
End If

On Error Goto 0

Set objFiles = objFolder.Files

If objFiles.Count > 0 Then 

	' Process every files in the INPUT folder (don't process subfolders)
	For Each objFile In objFiles

		' Process only .xlsx files
		If (LCase(objFSO.GetExtensionName(objFile.Name)) = "xlsx") Then

			If Not (isobject(cMSExcel)) Then

				' On the first processed file, instantiate Excel
				Set cMSExcel = New clsMSExcel

				cMSExcel.Instantiate
				
				' Don't show Excel's warnings
				cMSExcel.DisplayAlerts = False

				' Uncomment this line to see Excel's interface
				cMSExcel.MakeVisible

				' Allow the class to echo information on the DOS window
				cMSExcel.Verbose = True

			End if

			' Get the fullname of the file and inform Excel we'll work with
			' that file
			cMSExcel.FileName = objFile.Path

			' Open the file in Excel in a read-only mode (we will not modify
			' the source file)
			cMSExcel.Open (true)

			' Export visible sheets as CSV file, store files in the 
			' OUTPUT folder
			cMSExcel.ExportSheetsToCSV(sOutputFolder)
		
		End If

	Next
	
	Set objFile = Nothing

	If Not (cMSExcel is Nothing) Then
		cMSExcel.Quit
		Set cMSExcel = Nothing
	End If

End if 

Set objFiles = Nothing
Set objFolder = Nothing
Set objFSO = Nothing



