sub SendLogMail (strTo, strFrom, strSubject, strBody)

	Set cdoConfig = CreateObject("CDO.Configuration") 
	sch = "http://schemas.microsoft.com/cdo/configuration/" 
	cdoConfig.Fields.Item(sch & "sendusing") = 2 ' cdoSendUsingPort 
	cdoConfig.Fields.Item(sch & "smtpserver") = "smtp.adc.com" 
	cdoConfig.Fields.update 
	
	Set objMessage = CreateObject("CDO.Message")
	objMessage.Configuration = cdoConfig 
	objMessage.Subject = strSubject
	objMessage.Sender = strFrom
	objMessage.To = strTo
	objMessage.TextBody = strBody

	objMessage.Send

	Set objMessage = Nothing
	Set cdoConfig=nothing

end sub