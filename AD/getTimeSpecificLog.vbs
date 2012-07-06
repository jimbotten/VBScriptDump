'  November 5, 2007
'  Jim Montgomery
'  Pull Events
'  Pull time specific security events, format them.
strServer = InputBox("What server are the logs on?","Server")
dteTime = InputBox("What time did the event happen?" & vbCrLf & "Format:" & date & " " & time,"Time",date & " " & time)

Set StartDate = CreateObject("WbemScripting.SWbemDateTime")
Set EndDate = CreateObject("WbemScripting.SWbemDateTime")
Set RealTime = CreateObject("WbemScripting.SWbemDateTime")

StartDate.SetVarDate (CDate (dateadd("s",-60,dteTime)))
EndDate.SetVarDate (CDate (dateadd("s",60,dteTime)))

' wscript.echo date & " " & time
' wscript.echo startdate
' wscript.echo enddate

set WMI = GetObject("winmgmts:\\" & strServer & "\root\cimv2")

qry = "Select * from win32_ntlogevent WHERE logfile='Security' and (EventCode= 644 or eventcode = 539) and TimeWritten>='" & startDate & "' and TimeWritten<'" & endDate & "'"
set Events = WMI.ExecQuery(qry)

for each line in Events 
    RealTime.Value = line.TimeWritten
    wscript.echo RealTime.GetVarDate & vbtab _
				& line.EventCode & vbTab _
				& line.Message
    'objOut.WriteLine line.RecordNumber & "º" & line.SourceName & "º" & line.SourceName & "º" & _
     '           line.Type & "º" & line.ComputerName & "º" & _
      '          strTime & "º" & line.EventCode & "º" & _
       '         line.User & "º" & strMessage
next
