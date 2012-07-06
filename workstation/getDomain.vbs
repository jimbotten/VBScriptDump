strIn = inputbox("What is the remote machine name?")

msgbox "Remote Domain: " & GetComputerDomain(strIn)

function GetComputerDomain(ip)
  GetComputerDomain = "Function Can't find Domain"
  set wmi = getobject("winmgmts:{authority=ntlmdomain:" & strIn & "}\\" & ip & "\root\cimv2")
  wql = "select domain from Win32_ComputerSystem"
  set results = wmi.execquery(wql)
  for each obj in results
    GetComputerDomain = obj.Domain
  next
end function 