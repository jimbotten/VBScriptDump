'loop machines
'	ping a machine
'	if response
'		check on remote registry service
'		if its stopped
'			log it
'			start it
'		if started
'			log it
'	if not
'		log it
		
set objFSO = CreateObject("Scripting.FileSystemObject")
strInputFile = inputbox("What is the path to the input file containing computer names?","Input","C:\comps.txt")		
set objFile = objFSO.OpenTextFile(strInputFile)
strLogFile = inputbox("What is the path to the log file?","Log","C:\enremreg.log")		
set objOutput = objFSO.CreateTextFile(strLogFile)
strServiceName = inputbox("What is the name of the service you are enabling?" & vbCrLf & "(Service Name on the service properties sheet)","Service Name","RemoteRegistry")		

Do Until objFile.AtEndOfStream
	strComputer= objFile.ReadLine
	strLog = strComputer
	if pingComp(strComputer) then
		strLog = strLog & vbTab & "Ping response"
		if GetServiceState(strServiceName, strComputer) = "Running" then
			strLog = strLog & vbtab & strServiceName & " already running"
		else
			strLog = strLog & vbtab & strServiceName & " stopped"
			if StartService(strServiceName, strComputer) then
				strLog = strLog & vbtab & "Startup Failed"
			else
				strLog = strLog & vbtab & "Started"
			end if
		end if  
	else
		strLog = strLog & vbTab & "No ping response"
	end if
	objOutput.writeline strLog
Loop
objOutput.close
objFile.close

function pingComp(strComputer)
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colPings = objWMIService.ExecQuery ("Select * From Win32_PingStatus where Address = '" & strComputer & "'")

    For Each objStatus in colPings
        If IsNull(objStatus.StatusCode) _
            or objStatus.StatusCode<>0 Then 
            pingComp = 0
        Else
            pingComp = 1
        End If
    Next
End function
Function GetServiceState(strServiceName, strComputer)
 	ON error resume next
	Err.Clear
   	set osvcRemote = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	if (Err.Number <> 0) then
		GetServiceState= "Can't Connect: " & Err.Description
	else
    	osType = DetectOS(osvcRemote) 
		'msgbox osType & " and left 3 is " & left(osType,3)
		strType = left(osType,3)
		if (instr(strType, "W2K") or instr(strType, "WXP") or instr(strType, "2K3"))then		
			Set colServices = osvcRemote.ExecQuery ("Select * from Win32_Service where Name='" & strServiceName & "'")
			'msgbox "count = " & colServices.count
			For Each objService in colServices
				GetServiceState = objService.State
			Next
		end if
	end if
	set osvcRemote = nothing
end function

function detectOS(osvcRemote)
     set oOSInfo = osvcRemote.InstancesOf("Win32_OperatingSystem")
     'Only one instance is ever returned (the currently active OS), even though the following is a foreach.
     for each objOperatingSystem in oOSInfo
          if (objOperatingSystem.OSType <> 18) then
               ' Make sure that this computer is Windows NT-based.
               systemType = "OLD"
          else
               if (objOperatingSystem.Version = "5.0.2195") then
                    ' Windows 2000 SP2, SP3, SP4.
                    if (objOperatingSystem.ServicePackMajorVersion = 2) then 
                    	systemType = "OLD"		' SP2 can't be sussed, therefore its old
                    end if
                    if (objOperatingSystem.ServicePackMajorVersion = 3) then 
                    	systemType = "W2KSP3"
                    end if
                    if (objOperatingSystem.ServicePackMajorVersion = 4) then
                        systemType = "W2KSP4"
                    end if
               elseif (objOperatingSystem.Version = "5.1.2600") then
                    ' Windows XP RTM, SP1.					
                    if (objOperatingSystem.ServicePackMajorVersion = 0) or (objOperatingSystem.ServicePackMajorVersion = 1) then
                         systemType = "WXPSP1"
                    end if
					if (objOperatingSystem.ServicePackMajorVersion = 2) then
						 systemType = "WXPSP2"
					end if
               elseif (objOperatingSystem.Version = "5.2.3790") then
                    ' Windows Server 2003 RTM
                    if (objOperatingSystem.ServicePackMajorVersion = 0) then
                         systemType = "2K3SP0"
                    end if
               end if
               if (systemType = "") then
                    'This was a Windows NT-based computer, but not with a valid service pack.
                    systemType = "OLD"
               end if
          end if
     next
     detectOS = systemType
end function

function StartService(strServiceName, strComputer) 
'strServiceName = inputBox("What is the Name of the service you seek? (eg. Albd)",,"WZCSVC")
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colServices = objWMIService.ExecQuery ("Select * from Win32_Service where Name='" & strServiceName & "'")
	StartService = 1
	For Each objService in colServices
		objService.StartService()
		StartService = objService.ChangeStartMode("Automatic")
	Next
end function