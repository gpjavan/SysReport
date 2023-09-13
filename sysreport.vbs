'  MVC System Response VBscript v1.0 by Eblis01 (gpjavan@gmail.com) 05/18/2024
' ================================================================================
' DESCRIPTION: 
'  Basic script to retrieve system information and provide your PC with a response
'  to "How are you feeling?" where certain sysinfo parameters results in a program
'  "doing fine" or "not feeling well due to..."
'  Requires: No requirements, designed to fetch information from local system
'			 However, to get defrag status, script must run with elevated privs
'			 or UAC must be turned off, hence defrag check is off by default
'			 Network test requires LAN, Internet connectivity
' ================================================================================
' INSTRUCTIONS:
'  1. Copy this script wherever you want 
'  2. Edit the SETTINGS area below
'  3. set a custom command in JARVIS to run this
'   i.e. under Shell -
'    Your Command: How are you feeling? (or How are you doing? etc.)
'    Response: (leave it blank) .. The response is generated
'    Action: C:\Wherever\syscheck.vbs
'    Profile: (whatever)
' ================================================================================
' NOTE: The voice tts below is spelled improperly to facilitate proper enunciation
'  of those words. Please do NOT inform me I spelled things incorrectly, it is
'  INTENTIONAL for proper text-to-speech in my particular case!
' --------------------------------------------------------------------------------
' REVISIONS:
'  05/17/2014 - v1.0 - ported to vbs from powershell for portability/simplicity
'				added in tts functions, parameter selections via variables, etc.
' --------------------------------------------------------------------------------
' NO warranties are provided should you use this script. Although this script is
' very harmless, the author is not responsible for any and all mishaps that may
' occur during/from its use and your usage indicates acceptance of such.
' ================================================================================

'[ SETTINGS ]======================================================================
' Comment settings you want disabled with an apostrophe, remove to enable setting
SysMode = "R" ' C - Check Mode, R - Report Mode
'            Check mode will give humanized condition response based on parameters
'             you specify.. ie. I am feeling fine.. or I'm not feeling well because
'             I'm getting full since I'm down to 5% free hard drive space, etc..
'             NOTE: You can change the humanized responses below
'            Report mode just runs down all enabled parameters and report them 
'             back to you.. ie. System report is as follows, C drive down to 15%
'             free space.. etc.. Report mode will popup a report window too
'            This parameter will be OVERRIDDEN if passed as an argument via cmdline
'===----------------------------------------------------------------------------===
DSKcheck = 1 ' Uncomment to perform disk checks, commented out ignores Disk params
'----------------------------------------------------------------------------------
Diskspace = 1 ' Uncomment to perform disk space check
freeoutput = 3 ' 1 - In percentage, 2 - In space notation, 3 - Both % and GB
freewthreshold = 20 ' Specify free space WARNING percentage threshold (low = worse)
                    ' ie. if free space below 25% of total space, report warning
freewresp = "I'm starting to run out of room here! Delete some of that porn or buy me a new hard drive soon, please!"
freeathreshold = 10 ' Specify free space ALERT/ALARM percentage threshold
freearesp = "I'm out of room here! Delete some of that porn or buy me a new hard drive now!"
'----------------------------------------------------------------------------------
' WARNING: Disk defrag status ONLY works if running with elevated privs or with
'          UAC turned OFF *and* there is a slight delay while it retrieves defrag
'          status, so this option is turned off by default (commented out)
'Diskdefrag = 1 ' Uncomment to check disk defragmentation status
fragwthreshold = 10 ' Specify disk defrag WARNING threshold (higher # = worse)
fragathreshold = 15 ' Specify disk defrag ALERT/ALARM threshold
'===----------------------------------------------------------------------------===
CPUcheck = 1 ' Uncomment to perform CPU usage check
procwthreshold = 60 ' Specify cpu usage WARNING threshold (higher # = worse)
procwresp = "I feel a head-ache coming on! Seems my average CPU load is high! It's aspirin time or time to whip some processes!"
procathreshold = 80 ' Specify cpu usage ALERT/ALARM threshold
procaresp = "My head is killing me! Seems my average CPU load is very high! Give me an aspirin or kill some processes!"
'===----------------------------------------------------------------------------===
RAMcheck = 1 ' Uncomment to perform Memory usage check
memoutput = 3 ' 1 - Percentage, 2 - memory notation, 3 - both % and GB
memwthreshold = 75 ' Specify ram usage WARNING threshold (higher = worse)
memwresp = "I must be getting old. My memory isn't what it used to be, especially since something is hogging a bunch of it!"
memathreshold = 90 ' Specify ram usage ALERT/ALARM threshold
memaresp = "I feel Alzzheimer's kicking in. All my memory seems to be nearly gone because some process is sucking it all up!"
'NOTE: My sincerest apologies if anyone in their life has Alzheimers.. was the only thing I could equate to no memory left.
'===----------------------------------------------------------------------------===
NETcheck = 1 ' Uncomment to perform network checks (currently only IPv4 checks)
netAUTO = 1 ' Uncomment to automatically pull/use network adapter information
			' NOTE: uncomment below will override that particular auto-detect entry
			' Turning off netAUTO and uncomment below will ONLY use those entries.
'netlocal = "X.X.X.X" ' Uncomment/change to specify localhost address to check
'netlan = "X.X.X.X" ' Uncomment/change to specify local LAN IP address
'netgate = "X.X.X.X" ' Uncomment/change to specify local gateway
'netdns = "X.X.X.X" ' Uncomment/change to specify local/set DNS
'inetdns = "X.X.X.X" ' Uncomment/change to specify an Internet based DNS
'inetwan = "Example.com" ' Uncomment to specify an Internet based domain for test
netfresp = "Feeling a little cut off from the world since I'm detecting network issues. Give me my googly and facey-booky, please!"
'===----------------------------------------------------------------------------===
AllClear = "I'm feeling fine! All systems are operating efficiently! Thanks for asking!"
NotClear = "Well.. since you asked, I'm not feeling well, I don't want to complain but,"
'[ END SETTINGS ]==================================================================

' grab any command line arguments that might be passed
If WScript.Arguments.Count = 1 Then
	SysMode = WScript.Arguments.Item(0)
Else	
	SysMode = SysMode
End If

' Initialize some variables
Dim SysReport
Vers = "1.0"
errtoggle = 0
nerrtoggle = 0

' Initialize tts
set sapi = CreateObject("sapi.spvoice")
Set Sapi.Voice = sapi.GetVoices.Item(1)
' Error toggle alert
Function ToggleError()
	If (errtoggle = 0) AND (SysMode = "C") Then
		sapi.Speak NotClear
	    errtoggle = 1
	End If
	sapi.Speak "<break time='300ms' />" 'needed to end a delay between consecutive error messages
End Function

' Network Error toggle alert
Function ToggleNError()
	If (nerrtoggle = 0) AND (SysMode = "C") Then
		sapi.Speak NotClear
	    nerrtoggle = 1
	End If
End Function

' Free Space Disk Check function
Function dfchk()
	If SysMode = "R" Then
		sapi.Speak "Free Disk Space Status Report."
	End If

	' connect to local WMI
	strComputer = "." 
	Set objWMIService = GetObject("winmgmts:" _ 
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
	
	' physical harddrive query, hard-coded to one drive, the system drive, for now
	Set oShell = CreateObject("Wscript.Shell")
	sysdrive = oShell.ExpandEnvironmentStrings("%SystemDrive%")
	Set colDisks = objWMIService.ExecQuery _ 
			("Select * from Win32_LogicalDisk Where DriveType = 3 and DeviceID = '" & sysdrive & "'") 
	
	' structure for looping through multiple drives in place
	For Each objDisk in colDisks 
		'format space accordingly, gigabytes down to megabytes
		intFreeSpace = objDisk.FreeSpace
		intTotalSpace = objDisk.Size
		formFreeSpace = objDisk.FreeSpace / 1048576
		If formFreeSpace > 0 Then
		  formFreeSpace = formFreeSpace / 1024
		  formFreeSpace = Int(formFreeSpace)
		  formFreeAmt = " Gigabytes"
		Else
		  formFreeSpace = Int(formFreeSpace)		
		  formFreeAmt = " Megabytes"
		  ' Really hope your free space never gets down to the megs!
		End If
		
		'calc percentage of free space based on total space
		pctFreeSpace = intFreeSpace / intTotalSpace	
		pctFreeSpace = Int(FormatNumber(pctFreeSpace*100))
		
		'alert at thresholds
		If pctFreeSpace < freeathreshold Then
			FreeReport = "ALERT! ALERT! "
			ToggleError
			If SysMode = "C" Then sapi.Speak freearesp End If
			
		ElseIf pctFreeSpace < freewthreshold Then
			FreeReport = "WARNING! "
			ToggleError
			If SysMode = "C" Then sapi.Speak freewresp End If
		End If
		
		' Generate output based on set parameters
		FreeReport = FreeReport & sysdrive & " drive is at "
		If (freeoutput = 3) Or (freeoutput = 1) Then
			FreeReport = FreeReport & pctFreeSpace & "%"
		End If
		If (freeoutput = 3) Or (freeoutput = 2) Then
		 If freeoutput = 3 Then FreeReport = FreeReport & " or " End If
		 FreeReport = FreeReport & "roughly " & formFreeSpace & "" & formFreeAmt
		End If
		FreeReport = FreeReport & " of free space left."
		
		'generate/read report
		If SysMode = "R" Then
			sapi.Speak FreeReport
			SysReport = SysReport & "(Free Disk Space Report)--------------------------------------" & vbCrlf & vbCrlf
			SysReport = SysReport & FreeReport & vbCrlf &vbCrlf
		End If
	Next
End Function

' Disk Defrag System Check function
Function ddchk()
	If SysMode = "R" Then
		sapi.Speak "Disk Defrag Status Report."
	End If
	
	' connect to local WMI
	strComputer = "." 
	Set objWMIService = GetObject("winmgmts:" _ 
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 
	
	' physical harddrive query, hard-coded to one drive, the system drive, for now
	Set oShell = CreateObject("Wscript.Shell")
	sysdrive = oShell.ExpandEnvironmentStrings("%SystemDrive%")
	Set colVolumes = objWMIService.ExecQuery _ 
			("Select * from Win32_Volume Where DriveType = 3 and DriveLetter = '" & sysdrive & "'")  
 
 	' structure for looping through multiple drives in place
	For Each objVolume in colVolumes 
		If SysMode = "R" Then
			sapi.Speak "Analyzing defrag status of the " & objVolume.DriveLetter & " drive, this may take a moment..."
		End If

		FragReport = objVolume.DriveLetter & " drive analysis : "

		errResult = objVolume.DefragAnalysis(blnRecommended, objReport) 
		
		' no UAC or elevated permissions issues detected
		If errResult = 0 then 
			'alert at thresholds
			If objReport.FilePercentFragmentation > fragathreshold Then
				FragReport = "ALERT! ALERT! "
				ToggleError
			ElseIf objReport.FilePercentFragmentation > fragwthreshold Then
				FragReport = "WARNING! "
				ToggleError
			End If
		
			' generate output
			FragReport = FragReport & " the " & objVolume.DriveLetter & " drive is at " & objReport.FilePercentFragmentation & "% fragmentation"
           	If blnRecommended = True Then 
				FragReport = FragReport & " and should be defragged soon!"
				If SysMode = "C" Then
					ToggleError
					sapi.Speak "I'm a mess, my " & objVolume.DriveLetter & " drive is all fragmented. Defrag me!"
				End If
			Else 
				FragReport = FragReport & " but does not need to be defragged yet." 
			End If 
		' UAC or elevated permissions issue detected
		End If
		
		If errResult = 11 Then
			If SysMode = "R" Then
				sapi.Speak "Sorry, but the Defrag status check requires elevated permissions or U A C turned off to run properly!"
				sapi.Speak "Please correct this or disable the disk defrag option in the script settings."
			End If
			FragReport = FragReport & "Could not run, U A C error."
			If SysMode = "C" Then
				sapi.Speak "I couldn't check the defrag status because of a permissions issue. Fix it or turn defrag status check off!"
			End If
		End If

		' generate/read report
		If SysMode = "R" Then
			sapi.Speak FragReport
			SysReport = SysReport & "(Disk Defrag Status Report)--------------------------------------" & vbCrlf & vbCrlf
			SysReport = SysReport & FragReport & vbCrlf & vbCrlf
		End If
	Next 
End Function

' CPU Load Check Function
Function cpuchk()
	If SysMode = "R" Then
		sapi.Speak "C P U Load Status Report. (please wait, this may take a moment to query)"
	End If
	
	' connect to local WMI for overall CPU load
	On Error Resume Next
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set procItems = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor WHERE Name = '_Total'")
	For Each loadItem in procItems
		ProcTime = loadItem.PercentProcessorTime
	next
		
	' physical processor query
	Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
	
	' structure for looping through multiple processors in place
	For Each objItem in colItems
			PCore = objItem.NumberofCores
	next
	
	'alert at thresholds
	If Int(ProcTime) > procathreshold Then
		CLoadReport = "ALERT! ALERT! "
		ToggleError
		If SysMode = "C" Then sapi.Speak procaresp End If
	ElseIf Int(ProcTime) > procwthreshold Then
		CLoadReport = "WARNING! "
		ToggleError
		If SysMode = "C" Then sapi.Speak procwresp End If
	End If
		
	CLoadReport = CLoadReport & "The current " & PCore & "-core C P U, Load average is at " & ProcTime & "%"
		
	' generate/read report
	If SysMode = "R" Then
		sapi.Speak CLoadReport
		SysReport = SysReport & "(CPU Load Status Report)--------------------------------------" & vbCrlf & vbCrlf
		SysReport = SysReport & CLoadReport & vbCrlf & vbCrlf
	End If
End Function

' RAM Usage Check Function
Function ramchk()
	If SysMode = "R" Then
		sapi.Speak "System Memory Usage Status Report. (please wait, this may take a moment to generate)"
	End If
		
	'connect to local WMI for memory information
	On Error Resume Next
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set perfData = objWMIService.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfOS_Memory")
	
	For Each memItem in perfData
		MemUse = memItem.AvailableBytes
	Next
	
	'physical memory query
	Set sysData = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
	For Each memTot in sysData
		TotalMem = memTot.TotalPhysicalMemory
	Next
		
	'format ram accordingly, gigabytes down to megabytes
	intFreeMem = MemUse
	intTotalMem = TotalMem
	intUsedMem = TotalMem - MemUse
	formFreeMem = intFreeMem / 1048576
	If formFreeMem > 0 Then
	  formFreeMem = FormatNumber(formFreeMem/1024,1)
	  formFreeMAmt = " Gigabytes"
	Else
	  formFreeMem = FormatNumber(formFreeMem,2)
	  formFreeMAmt = " Megabytes"
	  ' Really hope your free ram never gets down to megs!
	End If
		
	'calc percentage of free/used ram based on total ram
	pctFreeMem = intFreeMem / intTotalMem	
	pctUsedMem = (intTotalMem - intFreeMem) / intTotalMem
	pctFreeMem = FormatNumber(pctFreeMem*100,1)
	pctUsedMem = FormatNumber(pctUsedMem*100,1)
	
	'format ram accordingly
	formUsedMem = intUsedMem / 1048576
	If formUsedMem > 0 Then
		formUsedMem = FormatNumber(formUsedMem/1024,1)
		formUsedMAmt = " Gigabytes"
	Else
		formUsedMem = FormatNumber(formUsedMem,2)
		formUsedMAmt = " Megabytes"
	End If
	
	'alert at thresholds
	If Round(pctUsedMem) > memathreshold Then
		MUseReport = "ALERT! ALERT! "
		ToggleError
		If SysMode = "C" Then sapi.Speak memaresp End If
	ElseIf Round(pctUsedMem) > memwthreshold Then
		MUseReport = "WARNING! "
		ToggleError
		If SysMode = "C" Then sapi.Speak memwresp End If
	End If

	' Generate output based on set parameters
	MUseReport = MUseReport & "The system memory is at "
	If (memoutput = 3) Or (memoutput = 1) Then
		MUseReport = MUseReport & pctUsedMem & "%"
	End If
	If (memoutput = 3) Or (memoutput = 2) Then
	If memoutput = 3 Then MUseReport = MUseReport & " or, " End If	
		MUseReport = MUseReport & "roughly " & formUsedMem & "" & formUsedMAmt
	End If
	MUseReport = MUseReport & " used up, and, "
	If (memoutput = 3) Or (memoutput = 1) Then
		MUseReport = MUseReport & pctFreeMem & "%"
	End If
	If (memoutput = 3) Or (memoutput = 2) Then
	If memoutput = 3 Then MUseReport = MUseReport & " or, " End If
	 MUseReport = MUseReport & "roughly " & formFreeMem & "" & formFreeMAmt
	End If
	MUseReport = MUseReport & " of free memory left."
		
	' generate/read report
	If SysMode = "R" Then
		sapi.Speak MUseReport
		SysReport = SysReport & "(System Memory Usage Status Report)--------------------------------------" & vbCrlf & vbCrlf
		SysReport = SysReport & MUseReport & vbCrlf & vbCrlf
	End If
End Function

' Ping test function
Function Ping(sNET) 
	On Error Resume Next 
	Set oWMI = GetObject("winMgmts:") 
	Set cPing = oWMI.ExecQuery("Select * from Win32_PingStatus Where Address = '" & sNET & "'") 
	For Each oPing In cPing 
		If oPing.StatusCode = 0 Then 
			Ping = "Passed." 
		Else 
			Ping = "Failed."
		End If 
	Next  
End Function 
	
' Network Connectivity Status Function
Function netchk()
	If SysMode = "R" Then
		sapi.Speak "Network Connectivity Status Report. (this may take a moment to complete)"
	End If
	
	If netAUTO = 1 Then
		If SysMode = "R" Then
			sapi.Speak "Auto-Network detection engaged."
		End If
		
		' connect to local WMI for active network adapter(s) information
		strComputer = "."
		Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
		Set colAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
		n = 1
 
		Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

		Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")

		For Each objItem in colItems
			Select Case objItem.AdapterTypeID
				Case 0 strAdapterType = "Ethernet" 
				Case 1 strAdapterType = "TokenRing" 
				Case 2 strAdapterType = "FDDI" 
				Case 3 strAdapterType = "WAN" 
				Case 4 strAdapterType = "LocalTalk" 
				Case 5 strAdapterType = "Ethernet with DIX" 
				Case 6 strAdapterType = "ARCNET" 
				Case 7 strAdapterType = "ARCNET 878.2" 
				Case 8 strAdapterType = "ATM" 
				Case 9 strAdapterType = "Wireless" 
				Case 10 strAdapterType = "Infrared Wireless" 
				Case 11 strAdapterType = "Bpc" 
				Case 12 strAdapterType = "CoWan" 
				Case 13 strAdapterType = "1394"
			End Select
 		Next
		
		' Auto pull/set network variables
		If IsEmpty(netlocal) Then netlocal = "127.0.0.1" End If
		'NetReport = NetReport & "Network Adapter " & n & " type is : " & strAdapterType & "." & vbCrlf
		For Each objAdapter in colAdapters
			If Not IsNull(objAdapter.IPAddress) Then
				For i = 0 To UBound(objAdapter.IPAddress)
					If IsEmpty(netlan) Then netlan = objAdapter.IPAddress(i) End If
				Next
			End If
			If Not IsNull(objAdapter.DefaultIPGateway) Then
				For i = 0 To UBound(objAdapter.DefaultIPGateway)
					If IsEmpty(netgate) Then netgate = objAdapter.DefaultIPGateway(i) End If
				Next
			End If
			If Not IsNull(objAdapter.DNSServerSearchOrder) Then
				For i = 0 To UBound(objAdapter.DNSServerSearchOrder)
					If IsEmpty(netdns) Then netdns = objAdapter.DNSServerSearchOrder(i) End If
				Next
			End If
			n = n + 1
		Next
		If IsEmpty(inetdns) Then inetdns = "8.8.8.8" End If
		If IsEmpty(inetwan) Then inetwan = "Google.com" End If
	Else
		If SysMode = "R" Then
			sapi.Speak "Network manual override engaged."
		End If
	End If
	
	' test down the line, skip entry if manual override in effect and commented
	If Not IsEmpty(netlocal) Then 
		PingTemp = Ping(netlocal)
		NetReport = NetReport & "Loopback test (" & netlocal & ") " & PingTemp & vbCrlf
		If (PingTemp = "Failed.") Then 
			ToggleNError 
		End If		
	End If
	If Not IsEmpty(netlan) Then 
		PingTemp = Ping(netlan)
		NetReport = NetReport & "LAN test (" & netlan & ") " & PingTemp & vbCrlf
		If (PingTemp = "Failed.") Then 
			ToggleNError 
		End If		
	End If
	If Not IsEmpty(netgate) Then 
		PingTemp = Ping(netgate)
		NetReport = NetReport & "Gateway test (" & netgate & ") " & PingTemp & vbCrlf
		If (PingTemp = "Failed.") Then 
			ToggleNError 
		End If
	End If
	If Not IsEmpty(netdns) Then 
		PingTemp = Ping(netdns)
		NetReport = NetReport & "D N S test (" & netdns & ") " & PingTemp & vbCrlf
		If (PingTemp = "Failed.") Then 
			ToggleNError 
		End If
	End If
	If Not IsEmpty(inetdns) Then 
		PingTemp = Ping(inetdns)
		NetReport = NetReport & "Internet D N S test (" & inetdns & ") " & PingTemp & vbCrlf
		If (PingTemp = "Failed.") Then 
			ToggleNError 
		End If
	End If
	If Not IsEmpty(inetwan) Then 
		PingTemp = Ping (inetwan)
		NetReport = NetReport & "Internet domain test (" & inetwan & ") " & PingTemp & vbCrlf
		If (PingTemp = "Failed.") Then
			ToggleNError
		End If
	End If
	
	If (nerrtoggle = 1) AND (SysMode = "C") Then
		sapi.Speak netfresp 
	End If
	
	' generate/read report
	If SysMode = "R" Then
		sapi.Speak NetReport
		SysReport = SysReport & "(Network Connectivity Status Report)--------------------------------------" & vbCrlf & vbCrlf
		SysReport = SysReport & NetReport & vbCrlf & vbCrlf
	End If
 End Function

 ' load disk check subroutines, if enabled
Function diskchk()
	If Diskspace = 1 Then dfchk
	If Diskdefrag = 1 Then ddchk
End Function

' If system mode is set to report, init the report popup string
If SysMode = "R" Then
	SysReport = "        ---=>>> SYSTEM REPORT <<<=--- ( by Eblis01 )" & vbCrlf
	SysReport = SysReport & "==================================================" & vbCrlf
	SysReport = SysReport & " Report summary follows for enabled parameters: " & vbCrlf & vbCrlf
	sapi.Speak "Initializing System Report Mode..."
	sapi.Speak "System Report Mode version " & Vers & " created by Eblis01."
	sapi.Speak "The System Report summary is, as specified and follows..."
End If

' load the cpu check function, if enabled
If CPUcheck = 1 Then
	cpuchk
End If

' load the ram check function, if enabled
If RAMcheck = 1 Then
	ramchk
End If

' load the disk check function, if enabled
If DSKcheck = 1 Then
	diskchk
End If

' load the network check function, if enabled
If NETcheck = 1 Then
	netchk
End If

' check for any error toggles
If SysMode = "C" Then
	If (errtoggle <> 1) AND (nerrtoggle <> 1) Then
		sapi.Speak AllClear
	End If
End If

' output report popup window at end, if enabled
If SysMode = "R" Then
	sapi.Speak "System Report completed."
	SysReport = SysREport & "        ---=>>> END SYSTEM REPORT <<<=---" & vbCrlf
	SysReport = SysReport & "==================================================" & vbCrlf
	Wscript.Echo SysReport
End If
