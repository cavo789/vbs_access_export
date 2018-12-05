' ========================================================
'
' Author: Christophe Avonture
' Date	: December 2018
'
' Open a MS Access database and export every forms, macros, 
' classes, modules and reports code
'
' Include the MS Access VBS class from 
' https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md
'
' ========================================================

Option Explicit

Class clsMSAccess

	Private oApplication
	Private bVerbose

	Private sDatabaseName

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo
	End Property

	Public Property Let DatabaseName(ByVal sFileName)
		sDatabaseName = sFileName
	End Property

	Private Sub Class_Initialize()

		bVerbose = False
		sDatabaseName = ""

		Set oApplication = Nothing

	End Sub

	Private Sub Class_Terminate()

		If Not (oApplication Is Nothing) Then

			' Quit MS Access only when it was opened by this script
			' This is the case when UserControl is equal to False
			If (oApplication.UserControl = false) then
				oApplication.Quit
			End If

			Set oApplication = Nothing

		End If

	End Sub

	Private Function CheckIfFileExists(sFileName)

		Dim objFSO
		Dim bReturn

		bReturn = True

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If Not (objFSO.FileExists(sFileName)) Then
		 	wScript.echo "Error, the file " & sFileName & " is not found"
			bReturn = False
		End If

		Set objFSO = Nothing

		CheckIfFileExists = bReturn

	End function

	' -----------------------------------------------------------
	'
	' Create a folder structure; create parents folder if not found
	' CreateFolderStructure("c:\temp\a\b\c\d\e") will create the
	' full structure in one call
	'
	' -----------------------------------------------------------
	Private Sub CreateFolderStructure(ByVal sFolderName)

		Dim arrPart, sBaseName, sDirName
		Dim objFSO

		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

		If Not (objFSO.FolderExists(sFolderName)) Then

			' Explode the folder name in parts
			arrPart = split(sFolderName, "\")
			sDirName = ""

			For Each sBaseName In arrPart

				If sDirName <> "" Then
					sDirName = sDirName & "\"
				End If

				sDirName = sDirName & sBaseName

				If (objFSO.FolderExists(sDirName) = False) Then
					objFSO.CreateFolder(sDirName & "\")
				End If

			Next

		End If

	End Sub

	' -----------------------------------------------------------
	' Open the database
	' -----------------------------------------------------------
	Public Sub OpenDatabase()

		If (oApplication is Nothing) Then

			On Error Resume Next

			' Perhaps the database is already opened (manually opened by the user)
			Set oApplication = GetObject(sDatabaseName)

			If Err.Number <> 0 Then

				' No, so start Access
				Set oApplication = CreateObject("Access.Application")

				If (Right(sDatabaseName,4) = ".adp") Then
					oApplication.OpenAccessProject sDatabaseName, false
				Else
					oApplication.OpenCurrentDatabase sDatabaseName, false
				End If

				oApplication.visible = true

				' MS Access has been opened by code, not under
				' the control of the user.
				' So UserControl = False will be used to detect that
				' we can close the DB and Quit Access when the job is done
				oApplication.UserControl = False

			End If

			On Error Goto 0

			If (oApplication is Nothing) then 
				' Ouch! Something goes wrong with the automation
				wScript.echo "ERROR - It was impossible to open the database by automation"
				wScript.echo "Please manually open MS Access and open the database, do not close "
				wScript.echo "Access and start this script again."
				wScript.quit
			End If

		End If

	End Sub

	' -----------------------------------------------------------
	' Close the database
	' -----------------------------------------------------------
	Public Sub CloseDatabase()

		If Not (oApplication is Nothing) then

			' Close the database only when it was opened by this script
			' This is the case when UserControl is equal to False
			If (oApplication.UserControl = false) then
				oApplication.CloseCurrentDatabase
			End If

		End If

	End Sub

	' -----------------------------------------------------------
	' Open a database and export every forms, macros, modules
	' and reports code to flat files
	'
	' Example
	' 	cMSAccess.Decompose("c:\temp\db1.accdb")
	'
	' -----------------------------------------------------------
	Public Sub Decompose(sDBName, sExportPath)

		Dim j, k
		Dim objFSO, obj, objShell
		Dim myComponent
		Dim sModuleType
		Dim sTempName, sOutFileName
		Dim sDBExtension, sDBParentFolder

		' Before starting, just verify that files exists
		' If no, show an error message and stop
		If CheckIfFileExists(sDBName) Then

			Set objFSO = CreateObject("Scripting.FileSystemObject")

			sDBExtension = objFSO.GetExtensionName(sDBName)
			sDBName = objFSO.GetBaseName(sDBName) & "." & sDBExtension
			sDBParentFolder = objFSO.GetParentFolderName(sDBName)

			' sDBParentFolder is empty when the script was called
			' with only the name of the file like f.i. "db1.accdb"
			' So, in that case, get the current folder and concatenate
			If (sDBParentFolder = "") Then
				Set objShell = WScript.CreateObject("WScript.Shell")
				sDBParentFolder = objShell.CurrentDirectory 
				Set objShell = Nothing
			End If 
	
			' Full path
			sDatabaseName = sDBParentFolder & "\" & sDBName

			If bVerbose Then
				wScript.echo "Process database " & sDatabaseName
			End If

			' Open the database, start MS Access if not yet opened
			Call OpenDatabase()

			If (sExportPath = "") then
				sExportPath = sDBParentFolder & "\src\" & sDBName & "\"
			End If

			If bVerbose Then
				wScript.echo "Exporting sources to " & sExportPath & vbCRLF
			End If

			Call CreateFolderStructure(sExportPath)

			' Export the code under each forms
			k = oApplication.CurrentProject.AllForms.Count
			If (k > 0) Then

				j = 0

				Call CreateFolderStructure(sExportPath & "Forms\")

				For Each obj In oApplication.CurrentProject.AllForms

					j = j + 1

					sOutFileName = obj.FullName & ".frm"
					sOutFileName = "Forms\" & sOutFileName

					If bVerbose Then
						wScript.echo "  Export form " & j & "/" & k & " - " & _
							obj.FullName & " to " & sOutFileName
					End If

					' 2 = acForm
					oApplication.SaveAsText 2, obj.FullName, sExportpath & sOutFileName
					oApplication.DoCmd.Close 2, obj.FullName

				Next
			End If 

			' Export macros
			k = oApplication.CurrentProject.AllMacros.Count
			If (k > 0) Then

				j = 0
			
				Call CreateFolderStructure(sExportPath & "Macros\")

				For Each obj In oApplication.CurrentProject.AllMacros

					j = j + 1

					sOutFileName = obj.FullName & ".txt"
					sOutFileName = "Macros\" & sOutFileName

					If bVerbose Then
						wScript.echo "  Export macro " & j & "/" & k & " - " & _
							obj.FullName & " to " & sOutFileName
					End If

					' 4 = acMacro
					oApplication.SaveAsText 4, obj.FullName, sExportpath & sOutFileName

				Next
			End If 

			' Export modules
			k = oApplication.CurrentProject.AllModules.Count
			If (k > 0) Then

				j = 0

				Call CreateFolderStructure(sExportPath & "Modules\")

				For Each obj In oApplication.CurrentProject.AllModules

					j = j + 1

					sOutFileName = obj.FullName & ".bas"
					sOutFileName = "Modules\" & sOutFileName

					If bVerbose Then
						wScript.echo "  Export module " & j & "/" & k & " - " & _ 
							obj.FullName & " to " & sOutFileName
					End If

					' 5 = acModule
					oApplication.SaveAsText 5, obj.FullName, sExportpath & sOutFileName

				Next
			End If 

			' Export the code under each reports
			k = oApplication.CurrentProject.AllReports.Count
			If (k > 0) Then

				j = 0

				Call CreateFolderStructure(sExportPath & "Reports\")

				For Each obj In oApplication.CurrentProject.AllReports

					j = j + 1

					sOutFileName = obj.FullName & ".bas"
					sOutFileName = "Reports\" & sOutFileName

					If bVerbose Then
						wScript.echo "  Export report " & j & "/" & k & " - " & _
							obj.FullName & " to " & sOutFileName
					End If

					' 3 = acReport
					oApplication.SaveAsText 3, obj.FullName, sExportpath & sOutFileName

				Next

			End If 
			
			' Close the database (only if it was opened by automation)
			Call CloseDatabase

		End If

	End Sub

End Class

Sub ShowHelp()

	wScript.echo " ======================================"
	wScript.echo " = Export MS Access code in flatfiles ="
	wScript.echo " ======================================"
	wScript.echo ""
	wScript.echo " Please specify the name of the database to process; f.i. : "
	wScript.echo " " & Wscript.ScriptName & " 'C:\Temp\db1.accdb'"
	wScript.echo ""
	wScript.echo "To get more info, please read https://github.com/cavo789/vbs_access_export"
	wScript.echo ""

	wScript.quit

End sub

Dim cMSAccess
Dim sFile
Dim arrDBNames(0)

	' Get the first argument (f.i. "C:\Temp\db1.accdb")
	If (wScript.Arguments.Count = 0) Then

		Call ShowHelp

	Else

		' Get the path specified on the command line
		sFile = Wscript.Arguments.Item(0)

		Set cMSAccess = New clsMSAccess

		cMSAccess.Verbose = True

		' The second parameter is where source files should be stored
		' If not specified, will be in the same folder where the
		' database is stored, in the /src subfolder (will be created if
		' needed)
		Call cMSAccess.Decompose(sFile, "")

		Set cMSAccess = Nothing

	End If
