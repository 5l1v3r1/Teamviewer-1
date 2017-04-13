Dim WshProcEnv, process_architecture, str1, str2, str3, str4, objShell, arrStr, UserName, count2, allComp(500,4), strDirectory, strFile, folder
Set objShell = WScript.CreateObject("WScript.Shell")
Set WshProcEnv = objShell.Environment("Process")
process_architecture= WshProcEnv("PROCESSOR_ARCHITECTURE") 
Const PATH = "C:\Program Files\TeamViewer"
Const PATH1 = "C:\Program Files (x86)\TeamViewer"
const file = "C:\Program Files\TeamViewer\TeamViewer.exe"
const file1 = "C:\Program Files (x86)\TeamViewer\TeamViewer.exe"
'Filepath = "\\share location\TeamViewer\TeamViewerID-2.csv" 

str1 = "10.0.43879.0" 'find the version number by using "sigcheck.exe -q -n ""%ProgramFiles%\TeamViewer\TeamViewer.exe" command line
str3 = "company name" ' find the custom Teamviewer host file
	'Edit this to point to a shared drive where all users have write access
		strDirectory = "\\share location\TeamViewer\team"
	'Edit this to point to a shared drive where all users have write access
		strFile = "\TeamViewerID.csv" 'Outputs to "TeamViewerID.csv" 
	objShell.Run "C:\windows\system32\reg.exe ADD ""HKCU\Software\Sysinternals\Sigcheck"" /v EulaAccepted /t REG_DWORD /d 1 /f", 0, True
	objShell.Run "C:\windows\system32\reg.exe ADD ""HKCU\Software\Sysinternals\PsLoggedon"" /v EulaAccepted /t REG_DWORD /d 1 /f", 0, True
	objShell.Run "taskkill /f /im teamviewer.exe", 0, True
	objShell.Run "taskkill /f /im teamviewer.exe", 0, True
	'str1 = "Sigcheck v2.52 - File version and signature viewerCopyright (C) 2004-2016 Mark RussinovichSysinternals - www.sysinternals.com"&version

UserName = runCMD("\\share locationTeamViewer\Teamviewer_Host_X86\PsLoggedon.exe -x")

' command line output 
	Function runCMD(strRunCmd)
		Set objShell = WScript.CreateObject("WScript.Shell")
		Set objExec = objShell.Exec(strRunCmd)	
		 strOut = ""
		 Do While Not objExec.StdOut.AtEndOfStream
		  strOut = strOut & objExec.StdOut.ReadLine()
		 Loop
		 Set objShell = Nothing
		 Set objExec = Nothing
		 
		if InStrRev(lcase(strOut), "igcheck") > 1 Then 
				runCMD = mid(strOut,InStrRev(lcase(strOut), ".com")+4,(len(strOut)-(InStrRev(strOut, "-"))))
		elseif InStrRev(lcase(strOut), ".json") > 1 Then
			    runCMD = mid(strOut,InStrRev(strOut, ":")+3,(len(strOut)-(InStrRev(strOut,":"))-4))
		elseif InStrRev(lcase(strOut), "loggedon") > 1 Then
				runCMD = mid(strOut,InStrRev(strOut, ":")+10,(len(strOut)-(InStrRev(strOut," "))))
		else
				runCMD = strOut
		end if
	
		'Wscript.Echo runCMD
	End Function
	
	
' Find the Last Logon User ID
	Function GetLastLoggedon()
		Const HKEY_CURRENT_USER = &H80000001
		Const HKEY_LOCAL_MACHINE = &H80000002
		strComputer = "."
		Set objRegistry = GetObject("winmgmts:\\" & strComputer & "\root\default:StdRegProv")
		strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\"
		strValueName = "LastLoggedOnUser"
		objRegistry.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
		'GetLastLoggedon = strValue
		'str1 = len(strValue)
		'variable = InStrRev(strValue, "\")
		GetLastLoggedon = Right(strValue, (LEN(strValue)-(InStrRev(strValue, "\")))) &"; " &UserName
		
		'Wscript.Echo " Last User logged on = " & GetLastLoggedon & "," & variable '& strValue & str1
	End Function

' Remove old verion and install Teamviewer 11 Host
	Function Teamvieweruninstall()
		On Error Resume Next
		Const DeleteReadOnly = TRUE
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strComputer & "\root\default:StdRegProv")
		objShell.Run("""%ProgramFiles%\TeamViewer\Version6\uninstall.exe"" /S"), 0, True
		objShell.Run("""%ProgramFiles%\TeamViewer\Version7\uninstall.exe"" /S"), 0, True
		objShell.Run("""%ProgramFiles%\TeamViewer\Version8\uninstall.exe"" /S"), 0, True
		objShell.Run("""%ProgramFiles%\TeamViewer\Version9\uninstall.exe"" /S"), 0, True
		objShell.Run("""%ProgramFiles%\TeamViewer\uninstall.exe"" /S"), 0, True
		objShell.Run("""%ProgramFiles(x86)%\TeamViewer\Version6\uninstall.exe"" /S"), 0, True
		objShell.Run("""%ProgramFiles(x86)%\TeamViewer\Version7\uninstall.exe"" /S"), 0, True
		objShell.Run("""%ProgramFiles(x86)%\TeamViewer\Version8\uninstall.exe"" /S"), 0, True
		objShell.Run("""%ProgramFiles(x86)%\TeamViewer\Version9\uninstall.exe"" /S"), 0, True
		objShell.Run("""%ProgramFiles(x86)%\TeamViewer\uninstall.exe"" /S"), 0, True	
		WScript.Sleep 3000
		objShell.RegDelete "HKLM\SOFTWARE\Wow6432Node\TeamViewer\"
		objShell.RegDelete "HKLM\SOFTWARE\TeamViewer\"
		objFSO.DeleteFile("C:\Program Files\TeamViewer\*"), DeleteReadOnly
		objFSO.DeleteFolder("C:\Program Files\TeamViewer"),DeleteReadOnly
		objFSO.DeleteFile("C:\Program Files (x86)\TeamViewer\*"), DeleteReadOnly	
		objFSO.DeleteFolder("C:\Program Files (x86)\TeamViewer"),DeleteReadOnly				
		If process_architecture = "x86" Then
			'if (objFSO.FolderExists(PATH)) Then
			if (objFSO.FileExists(file)) Then
					Teamvieweruninstall()
				'wscript.echo "file 	Teamvieweruninstall() in x86 Function Teamvieweruninstall()"
			else
				Teamviewer()
			end if
		else
			'if (objFSO.FolderExists(PATH1)) Then
			if (objFSO.FileExists(file1)) Then
				Teamvieweruninstall()
				'wscript.echo "file 	Teamvieweruninstall() in x64 Function Teamvieweruninstall()"
			else
				Teamviewer()
				'wscript.echo "file  Teamviewer() PATH1 in x64 Function Teamvieweruninstall()"  '& objFSO.getFolder(PATH1).files.Count & objFSO.FolderExists(PATH1) & objFSO.getFolder(PATH1) & objFSO.getFolder(PATH1).subfolders.Count
			end if
		end if
	End function
' TeamViewer installation
	Function Teamviewer()
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objShell = WScript.CreateObject("WScript.Shell")
		'objShell.Run("msiexec /fpecms \\share locationTeamViewer\Teamviewer_Host_X86\TeamViewer_Host-idcuv6hpt7.msi /qn"), 0, True
		'objShell.Run("msiexec /i \\share locationTeamViewer\Teamviewer_Host_X86\TeamViewer_Host-idcftb8ctj.msi /qn"), 0, True	
		If process_architecture = "x86" Then
			'if (objFSO.FolderExists(PATH)) Then
			if (objFSO.FileExists(file)) Then
				Teamvieweruninstall()
				'wscript.echo "Teamvieweruninstall() in Function Teamviewer()"
			else 
				objShell.Run("msiexec /i \\share locationTeamViewer\Teamviewer_Host_X86\TeamViewer_Host-idcuv6hpt7.msi /qn"), 0, True	
				'WScript.Sleep 1000
				'wscript.echo "file installx86 PATH in Function Teamviewer()" & objFSO.getFolder(PATH).files.Count & objFSO.FolderExists(PATH) & objFSO.getFolder(PATH) & objFSO.getFolder(PATH1).subfolders.Count
				'objShell.Run("""%ProgramFiles(x86)%\TeamViewer\teamviewer.exe"""), 0, True
				objShell.Run("""%ProgramFiles%\TeamViewer\teamviewer.exe"""), 0, True
				objShell.Run("net stop teamviewer"), 0, True		
				objShell.Run "taskkill /f /im teamviewer.exe", 0, True
				objShell.Run "taskkill /f /im teamviewer.exe", 0, True
			end if
		else
			'if (objFSO.FolderExists(PATH1)) Then
			if (objFSO.FileExists(file1)) Then
				Teamvieweruninstall()
				'wscript.echo "Teamvieweruninstall() in Function Teamviewer()" '& objFSO.getFolder(PATH1).files.Count & objFSO.FolderExists(PATH1) & objFSO.getFolder(PATH1).subfolders.Count
			else 
				'wscript.echo "file install PATH1 in Function Teamviewer()" '& objFSO.getFolder(PATH1).files.Count & objFSO.FolderExists(PATH1) '& objFSO.getFolder(PATH1)& objFSO.getFolder(PATH1).subfolders.Count & PATH1
				objShell.Run("msiexec /i \\share locationTeamViewer\Teamviewer_Host_X86\TeamViewer_Host-idcuv6hpt7.msi /qn"), 0, True
				'WScript.Sleep 1000				
				objShell.Run("""%ProgramFiles(x86)%\TeamViewer\teamviewer.exe"""), 0, True
				'objShell.Run("""%ProgramFiles%\TeamViewer\teamviewer.exe"""), 0, True
				objShell.Run("net stop teamviewer"), 0, True		
				objShell.Run "taskkill /f /im teamviewer.exe", 0, True
				objShell.Run "taskkill /f /im teamviewer.exe", 0, True
			end if
		end if

		objShell.Run("msiexec /i \\share locationTeamViewer\Teamviewer_Host_X86\TeamViewer_Host.msi /qn"), 0, True		
		if ((objFSO.fileExists(file1)) or (objFSO.fileExists(file))) Then
				objShell.Run("net stop teamviewer"), 0, True		
				objShell.Run "taskkill /f /im teamviewer.exe", 0, True
				objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		else
		objShell.Run("msiexec /i \\share locationTeamViewer\Teamviewer_Host_X86\TeamViewer_Host-idcuv6hpt7.msi /qn"), 0, True
		objShell.Run("net stop teamviewer"), 0, True		
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		end if
		'objShell.Run("c:\Windows\System32\xcopy.exe ""\\share locationTeamViewer\Teamviewer_Host_X86\Settings\*"" ""C:\Program Files\TeamViewer"" /s /i /E /c /y /r"), 0, True	
		objShell.Run("net start teamviewer"), 0, True
		'objShell.Run("""%ProgramFiles(x86)%\TeamViewer\teamviewer.exe"""), 0, True
		'objShell.Run("""%ProgramFiles%\TeamViewer\teamviewer.exe"""), 0, True	
        'Validation 'Check the Teamviewer ID
		Team() 'store the Teamviewer ID
		logo()
		'Set objShell = Nothing
	End Function
'Find the Bin Aweidha Logo Files check
	Function logo()
		Const HKEY_CURRENT_USER = &H80000001
		Const HKEY_LOCAL_MACHINE = &H80000002
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		strComputer = "."
		Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strComputer & "\root\default:StdRegProv")
		If process_architecture = "x86" Then
		strKeyPath = "SOFTWARE\TeamViewer\"
		strValueName = "OwningManagerAccountName"
		oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwTVID
		Else
		strKeyPath = "SOFTWARE\Wow6432Node\TeamViewer\"
		strValueName = "OwningManagerAccountName"
		oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwTVID
		End If
		Set oReg = Nothing
		logo = strValueName
		if logo = "company name" Then
			'if str3 = str4 then
			'wscript.echo logo
		end if	
	End Function


' Find the Teamviewer ID
	Function TeamviewerID()
		Const HKEY_CURRENT_USER = &H80000001
		Const HKEY_LOCAL_MACHINE = &H80000002
		strComputer = "."
		'Set WshShell = WScript.CreateObject("WScript.Shell")
		'Set objShell = WScript.CreateObject( "WScript.Shell")
		'OSArchCheck = objShell.RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE")
		Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
		strComputer & "\root\default:StdRegProv")
		If process_architecture = "x86" Then
		strKeyPath = "SOFTWARE\TeamViewer\"
		strValueName = "ClientID"
		oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwTVID
		if len(dwValue) >= 9 then strText2 = domain & ";" & comp & ";" & user & ";" & dwValue end if
		Else
		strKeyPath = "SOFTWARE\Wow6432Node\TeamViewer\"
		strValueName = "ClientID"
		oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwTVID
		if len(dwValue) >= 9 then strText2 = domain & ";" & comp & ";" & user & ";" & dwValue end if
		End If
		Set oReg = Nothing
		TeamviewerID = dwTVID
	End Function

	Function Validation()
		Const ForReading = 1
		'Dim strSearchFor, length
		'strSearchFor = TeamviewerID
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objTextFile = objFSO.OpenTextFile(strDirectory & strFile, ForReading)
		do until objTextFile.AtEndOfStream
		   strLine = objTextFile.ReadLine()
		If InStr(strLine, TeamviewerID) <> 0 then
			Validation = InStr(strLine, TeamviewerID)
			'Wscript.Echo Validation 'ID 'strSearchFor 
		 End If
		loop   
	End Function

	
	
	Function EnvString(variable)
		set objShell = WScript.CreateObject( "WScript.Shell" )
		variable = "%" & variable & "%"
		EnvString = objShell.ExpandEnvironmentStrings(variable)    
		Set objShell = Nothing    
	End Function
	
	Function Team
			On Error Resume Next
			'Dim allComp(500,4)
			'Dim arrStr, count2

			' ForAppending = 8 ForReading = 1, ForWriting = 2
			Const ForAppending = 8
			Const ForReading = 1
			Const ForWriting = 2
			Const HKEY_CURRENT_USER = &H80000001
			Const HKEY_LOCAL_MACHINE = &H80000002

			'Get Enviromental Vars
			'user=lcase(EnvString("username"))
			user=lcase(GetLastLoggedon)
			comp=lcase(EnvString("ComputerName"))
			domain=lcase(EnvString("UserDomain"))

			'Edit this to point to a shared drive where all users have write access
			'strDirectory = "\\....\TeamViewer"
			'Edit this to point to a shared drive where all users have write access
			'strFile = "\TeamViewerID.csv" 'Outputs to "TeamViewerID.csv" 


			strComputer = "."
			 
			Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" &_
			 strComputer & "\root\default:StdRegProv")
			
			strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\"
			strValueName = "LastLoggedOnUser"
			oReg.GetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, strValue
			 
			strKeyPath = "SOFTWARE\Wow6432Node\TeamViewer\"
			strValueName = "ClientID"
			oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
			if len(dwValue) >= 9 then strText2 = domain & ";" & comp & ";" & user & ";" & dwValue & ";" & strValue end if

			strKeyPath = "SOFTWARE\TeamViewer\"
			strValueName = "ClientID"
			oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,dwValue
			if len(dwValue) >= 9 then strText2 = domain & ";" & comp & ";" & user & ";" & dwValue & ";" & strValue end if	'Domain;Computer;User;ID

			Set objFSO = CreateObject("Scripting.FileSystemObject")
			If objFSO.FolderExists(strDirectory) Then
			   Set objFolder = objFSO.GetFolder(strDirectory)
			   'wscript.echo("Folder Exists!")
			Else
				Set objFolder = objFSO.CreateFolder(strDirectory)  
				If not objFSO.FolderExists(strDirectory) Then
					wscript.quit
				End If
				'wscript.echo("Folder Does Not Exist!")
			End If

			If objFSO.FileExists(strDirectory & strFile) Then
				'wscript.echo("File Exists!")
				Set objFolder = objFSO.GetFolder(strDirectory)
				Set objTextFile = objFSO.OpenTextFile(strDirectory & strFile)
				count = 0
				Do while NOT objTextFile.AtEndOfStream
					arrStr = split(objTextFile.ReadLine,";")
					'wscript.echo(Count & "," & arrStr)
					if arrStr(1) <> comp Then
						count = count + 1
						allcomp(count,0)=arrStr(0) 'Domain / Company
						allcomp(count,1)=arrStr(1) 'Computer Name
						allcomp(count,2)=arrStr(2) 'Username
						allcomp(count,3)=arrStr(3) 'Teamviewer ID
					End If
				Loop
					
					'wscript.echo("EOF!")
					objTextFile.Close
					set objTextFile = nothing
					
					objFSO.DeleteFile(strDirectory & strFile)
					
					Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
					set objFile = nothing
					
					Set objTextFile = objFSO.OpenTextFile(strDirectory & strFile, ForAppending, True)
					count2 = Count
					'wscript.echo(count2)
					for count = 1 to count2
						strtext=allcomp(count,0)&";"&allcomp(count,1)&";"&allcomp(count,2)&";"&allcomp(count,3)
						objTextFile.WriteLine(strText)
						'wscript.echo(strText)
					Next
					objTextFile.WriteLine(strText2)
					'wscript.echo(strText2)
					objTextFile.Close
					set objTextFile = nothing
			Else
				Set objFile = objFSO.CreateTextFile(strDirectory & strFile)
				set objFile = nothing
				Set objTextFile = objFSO.OpenTextFile(strDirectory & strFile, ForAppending, True)
				objTextFile.WriteLine(strText2)
				objTextFile.Close
				set objTextFile = nothing
			End If

			set objFolder = nothing
			set objFSO = Nothing
			wscript.quit
			
	End Function					
if process_architecture = "x86" Then 
												'Windows 32 bit


		str2 = runCMD("\\share locationTeamViewer\Teamviewer_Host_X86\sigcheck.exe -q -n ""%ProgramFiles%\TeamViewer\TeamViewer.exe""")
		str4 = runCMD("find ""company name"" ""%ProgramFiles%\TeamViewer\TeamViewer.json""")


	If str1 = str2 Then
		
		if str3 = str4 then 
		Set objShell = WScript.CreateObject( "WScript.Shell")
		objShell.Run("""%ProgramFiles%\TeamViewer\teamviewer.exe"""), 0, True	
		'objShell.Run("""%ProgramFiles(x86)%\TeamViewer\teamviewer.exe""")
				if Validation >= 1 Then
					'Wscript.Echo "ID Found"
				Else
				team()
					'Wscript.Echo "ID Not Found" & Validation 'TeamviewerID '
				end if
				
		'Wscript.Echo("corporate version" & Validation)
				
		else 
		Set objShell = WScript.CreateObject( "WScript.Shell")
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		Teamviewer()
		

		'Wscript.Echo("not corporate")

		End If
	Else
 		Set objShell = WScript.CreateObject( "WScript.Shell")
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		Teamviewer()
	
		'Wscript.Echo("old version")
		
	End If
	
'wscript.echo "Windows 32 bit OS"
														'Windows 64 Bit
 elseif process_architecture = "AMD64" Then
	
'	Set objShell = WScript.CreateObject( "WScript.Shell")
'	objShell.Run "C:\windows\system32\reg.exe ADD ""HKCU\Software\Sysinternals\Sigcheck"" /v EulaAccepted /t REG_DWORD /d 1 /f", 0, True
'	objShell.Run "taskkill /f /im teamviewer.exe", 0, True
'	objShell.Run "taskkill /f /im teamviewer.exe", 0, True

		'str1 = "Sigcheck v2.52 - File version and signature viewerCopyright (C) 2004-2016 Mark RussinovichSysinternals - www.sysinternals.com11.0.59518.0"
		str2 = runCMD("\\share locationTeamViewer\Teamviewer_Host_X86\sigcheck.exe -q -n ""%ProgramFiles(x86)%\TeamViewer\TeamViewer.exe""")
		'str3 = "---------- C:\PROGRAM FILES (X86)\TEAMVIEWER\TEAMVIEWER.JSON    ""title"": ""company name"","
		str4 = runCMD("find ""company name"" ""%ProgramFiles(x86)%\TeamViewer\TeamViewer.json""")

'	Function runCMD(strRunCmd)
'	 Set objShell = WScript.CreateObject("WScript.Shell")
'	 Set objExec = objShell.Exec(strRunCmd)
'	 strOut = ""
'	 Do While Not objExec.StdOut.AtEndOfStream
'	  strOut = strOut & objExec.StdOut.ReadLine()
'	 Loop
'	 Set objShell = Nothing
'	 Set objExec = Nothing
'	 runCMD = strOut
'	End Function

	If str1 = str2 Then
		
		if str3 = str4 then 
		Set objShell = WScript.CreateObject( "WScript.Shell")
		'objShell.Run("""%ProgramFiles%\TeamViewer\teamviewer.exe""")
		objShell.Run("""%ProgramFiles(x86)%\TeamViewer\teamviewer.exe"""), 0, True	
		
		if Validation >= 1 Then
					'Wscript.Echo "ID Found"
		Else
				team()
					'Wscript.Echo "ID Not Found" & Validation 'TeamviewerID '
		end if
				
		'Wscript.Echo("corporate version" & Validation)
		
		else 
		Set objShell = WScript.CreateObject( "WScript.Shell")
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		Teamviewer()	
		'Wscript.Echo("not corporate" )

		End If
	Else
 		Set objShell = WScript.CreateObject( "WScript.Shell")
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		objShell.Run "taskkill /f /im teamviewer.exe", 0, True
		Teamviewer()
	
		
	'Wscript.Echo("old version")

	End If

  'wscript.echo "Windows 64 bit OS"
 else 
 'wscript.echo "Unable to find system type"
end if
