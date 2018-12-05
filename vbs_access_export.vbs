' ========================================================
'
' Author: Christophe Avonture
' Date	: December 2018
'
' Open a MS Access database and export every forms, macros, 
' modules, queries and reports code
'
' Include the somes VBS classes from 
' https://github.com/cavo789/vbs_scripts
'
' ========================================================

Option Explicit

Class clsFiles

	Dim objFSO, objFile

	Private bVerbose

	Public Property Let verbose(ByVal bYesNo)
		bVerbose = bYesNo
	End Property

	Private Sub Class_Initialize()
		bVerbose = False
		Set objFSO = CreateObject("Scripting.FileSystemObject")
	End Sub

	Private Sub Class_Terminate()
		Set objFSO = Nothing
	End Sub

	' --------------------------------------------------
	' Create a text file
	' --------------------------------------------------
	Public Sub CreateText(ByVal sFileName, ByVal sContent)

		If bVerbose Then
		'	wScript.echo "Create file " & sFileName 
		End If

		Set objFile = objFSO.CreateTextFile(sFileName, 2, True)
		objFile.Write sContent
		objFile.Close
		Set objFile = Nothing

	End Sub

	' --------------------------------------------------
	' Verifies the existence of the file 
	' --------------------------------------------------
	Public Function Exists(ByVal sFileName)
		Exists = objFSO.FileExists(sFileName)
	End Function

	' --------------------------------------------------
	' Remove some characters like a wildcard (*) if 
	' present in the suggested filename
	' --------------------------------------------------
	Public Function MakeSafe(ByVal sFileName) 
	
		' Don't allow * in a filename
		sFileName = Replace(sFileName, "*", "_")

		MakeSafe = sFileName

	End Function

	' --------------------------------------------------
	' Return the file extension (f.i. "accdb")
	' --------------------------------------------------
	Public Function GetExtensionName(ByVal sFileName)
		GetExtensionName = objFSO.GetExtensionName(sFileName)
	End Function 

	' --------------------------------------------------
	' Return only the file name (f.i. "db1")
	' --------------------------------------------------
	Public Function GetBaseName(ByVal sFileName)
		GetBaseName = objFSO.GetBaseName(sFileName)
	End Function 

	' --------------------------------------------------
	' Return the folder where the file is stored (f.i. c:\temp)
	' --------------------------------------------------
	Public Function GetParentFolderName(ByVal sFileName)

		Dim sPath
		Dim objShell

		sPath = ""

		If (Exists(sFileName)) Then

			sPath = objFSO.GetParentFolderName(sFileName)

			' sPath is empty when sFileName was just a filename 
			' like f.i. "db1.accdb". So, in that case, get the current 
			' folder and concatenate
			If (sPath = "") Then
				Set objShell = WScript.CreateObject("WScript.Shell")
				sPath = objShell.CurrentDirectory 
				Set objShell = Nothing
			End If 

		End If

		GetParentFolderName = sPath

	End Function 

End Class

Class clsFolders

	Dim objFSO
	Private bVerbose

	Private Sub Class_Initialize()
		bVerbose = False
		Set objFSO = CreateObject("Scripting.FileSystemObject")
	End Sub

	Private Sub Class_Terminate()
		Set objFSO = Nothing
	End Sub

	Public Property Let verbose(ByVal bYesNo)
		bVerbose = bYesNo
	End Property

	' -----------------------------------------------------------
	'
	' Create a folder structure; create parents folder if not found
	' MakeFolder("c:\temp\a\b\c\d\e") will create the
	' full structure in one call
	'
	' -----------------------------------------------------------
	Public Sub MakeFolder(ByVal sFolderName)

		Dim arrPart, sBaseName, sDirName

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

End Class

Class clsMSAccess

	Private oApplication
	Private bVerbose
	Private sExportPath
	Private cFiles
	Private cFolders

	Private sDatabaseName

	Public Property Let verbose(ByVal bYesNo)
		bVerbose = bYesNo
		cFiles.Verbose = bYesNo
		cFolders.Verbose = bYesNo
	End Property

	Public Property Let DatabaseName(ByVal sFileName)
		sDatabaseName = sFileName
	End Property

	' -----------------------------------------------------------
	' Folder where files will be generated
	' -----------------------------------------------------------
	Public Property Let ExportPath(ByVal sPath)

		If bVerbose Then
			wScript.echo "Exporting sources to " & sPath & vbCRLF
		End If
		
		sExportPath = sPath

	End Property

	Private Sub Class_Initialize()

		Set cFiles = new clsFiles 
		Set cFolders = new clsFolders 
		Set oApplication = Nothing

		bVerbose = False
		sDatabaseName = ""
		sExportPath = ""

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

		Set cFiles = Nothing
		Set cFolders = Nothing

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
	' Export forms as .frm files
	' -----------------------------------------------------------
	Private Sub ExportForms()
	
		Dim j, k
		Dim sOutFileName, sSQL
		Dim obj 
	
		k = oApplication.CurrentProject.AllForms.Count

		If (k > 0) Then

			If bVerbose Then
				wScript.Echo vbCrLf & "Export all forms"
			End if

			j = 0

			Call cFolders.MakeFolder(sExportPath & "Forms\")

			For Each obj In oApplication.CurrentProject.AllForms

				j = j + 1

				sOutFileName = cFiles.MakeSafe(obj.FullName) & ".frm"
				sOutFileName = "Forms\" & sOutFileName

				If bVerbose Then
					wScript.echo "	Export form " & j & "/" & k & " - " & _
						obj.FullName & " to " & sOutFileName
				End If

				' 2 = acForm
				oApplication.SaveAsText 2, obj.FullName, sExportpath & sOutFileName
				oApplication.DoCmd.Close 2, obj.FullName

			Next

		End If 

	End Sub

	' -----------------------------------------------------------
	' Export macros as .txt files
	' -----------------------------------------------------------
	Private Sub ExportMacros()
	
		Dim j, k
		Dim sOutFileName, sSQL
		Dim obj 

		k = oApplication.CurrentProject.AllMacros.Count

		If (k > 0) Then

			If bVerbose Then
				wScript.Echo vbCrLf & "Export all macros"
			End if

			j = 0
		
			Call cFolders.MakeFolder(sExportPath & "Macros\")

			For Each obj In oApplication.CurrentProject.AllMacros

				j = j + 1

				sOutFileName = cFiles.MakeSafe(obj.FullName) & ".txt"
				sOutFileName = "Macros\" & sOutFileName

				If bVerbose Then
					wScript.echo "	Export macro " & j & "/" & k & " - " & _
						obj.FullName & " to " & sOutFileName
				End If

				' 4 = acMacro
				oApplication.SaveAsText 4, obj.FullName, sExportpath & sOutFileName

			Next
		End If 

	End Sub

	' -----------------------------------------------------------
	' Export modules as .bas files
	' -----------------------------------------------------------
	Private Sub ExportModules()
	
		Dim j, k
		Dim sOutFileName, sSQL
		Dim obj 
		
		k = oApplication.CurrentProject.AllModules.Count

		If (k > 0) Then

			If bVerbose Then
				wScript.Echo vbCrLf & "Export all modules"
			End if

			j = 0

			Call cFolders.MakeFolder(sExportPath & "Modules\")

			For Each obj In oApplication.CurrentProject.AllModules

				j = j + 1

				' Don't allow * in a filename
				sOutFileName = cFiles.MakeSafe(obj.FullName) & ".bas"
				sOutFileName = "Modules\" & sOutFileName

				If bVerbose Then
					wScript.echo "	Export module " & j & "/" & k & " - " & _ 
						obj.FullName & " to " & sOutFileName
				End If

				' 5 = acModule
				On Error Resume Next

				oApplication.SaveAsText 5, obj.FullName, sExportpath & sOutFileName

				If Err.number <> 0 Then
					wScript.echo "      An error has occured: " & Err.Description
					wScript.echo "      If a password is needed to see the code, please open the "
					wScript.echo "      database manually and open a module so you can specify the "
					wScript.echo "      password. Then, without closing the database, restart this script"
					Err.clear

					' And stop 
					Exit For

				End If

				On Error Goto 0

			Next

		End If 

	End sub

	' -----------------------------------------------------------
	' Export queries as .sql files
	' -----------------------------------------------------------
	Private Sub ExportQueries()

		Dim j, k
		Dim sOutFileName, sSQL
		Dim obj 

		k = oApplication.CurrentDb.QueryDefs.Count

		If (k > 0) Then

			If bVerbose Then
				wScript.Echo vbCrLf & "Export all queries"
			End if

			j = 0

			Call cFolders.MakeFolder(sExportPath & "Queries\")

			' Process all objects
			For Each obj In oApplication.CurrentDb.QueryDefs

				j = j + 1

				sOutFileName = cFiles.MakeSafe(obj.Name) & ".sql"
				sOutFileName = "Queries\" & sOutFileName

				If bVerbose Then
					wScript.echo "	Export query " & j & "/" & k & " - " & _
						obj.Name & " to " & sOutFileName
				End If

				' Get the query SQL statement
				sSQL = Trim(obj.SQL)

				' replace carriage return by a single space
				sSQL = Replace(sSQL, vbCrLf, " ")

				Call cFiles.CreateText(sExportPath & sOutFileName, sSQL)

			Next

		End If

	End Sub
	
	' -----------------------------------------------------------
	' Export reports as .txt files
	' -----------------------------------------------------------
	Private Sub ExportReports()
	
		Dim j, k
		Dim sOutFileName, sSQL
		Dim obj 

		k = oApplication.CurrentProject.AllReports.Count

		If (k > 0) Then

			If bVerbose Then
				wScript.Echo vbCrLf & "Export all reports"
			End if

			j = 0

			Call cFolders.MakeFolder(sExportPath & "Reports\")

			For Each obj In oApplication.CurrentProject.AllReports

				j = j + 1

				sOutFileName = cFiles.MakeSafe(obj.FullName) & ".txt"
				sOutFileName = "Reports\" & sOutFileName

				If bVerbose Then
					wScript.echo "	Export report " & j & "/" & k & " - " & _
						obj.FullName & " to " & sOutFileName
				End If

				' 3 = acReport
				oApplication.SaveAsText 3, obj.FullName, sExportpath & sOutFileName

			Next

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
	Public Sub Decompose(sDBName)

		Dim sDBExtension, sDBParentFolder

		' Before starting, just verify that files exists
		' If no, show an error message and stop
		If cFiles.Exists(sDBName) Then

			sDBExtension = cFiles.GetExtensionName(sDBName)
			sDBName = cFiles.GetBaseName(sDBName) & "." & sDBExtension
			sDBParentFolder = cFiles.GetParentFolderName(sDBName)

			' Full path
			sDatabaseName = sDBParentFolder & "\" & sDBName

			If bVerbose Then
				wScript.echo "Process database " & sDatabaseName
			End If

			' Open the database, start MS Access if not yet opened
			Call OpenDatabase()

			If (sExportPath = "") then
				ExportPath = sDBParentFolder & "\src\" & sDBName & "\"
			End If

			Call cFolders.MakeFolder(sExportPath)

			If (oApplication.CurrentProject.AllForms.Count > 0) Then
				Call ExportForms()
			End if

			If (oApplication.CurrentProject.AllMacros.Count > 0) Then
				Call ExportMacros()
			End if

			If (oApplication.CurrentProject.AllModules.Count > 0) Then
				Call ExportModules()
			End if

			If (oApplication.CurrentDb.QueryDefs.Count > 0) Then
				Call ExportQueries()
			End if

			If (oApplication.CurrentProject.AllReports.Count > 0) Then
				Call ExportReports()
			End if

			' Close the database (only if it was opened by automation)
			Call CloseDatabase

		Else

			wScript.echo "Error, the file " & sDBName & " is not found"

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

		Call cMSAccess.Decompose(sFile)

		Set cMSAccess = Nothing

	End If
