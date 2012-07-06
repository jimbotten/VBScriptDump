OPTION EXPLICIT

DIM CRLF, TAB
DIM strServer
DIM objWebService

TAB  = CHR( 9 )
CRLF = CHR( 13 ) & CHR( 10 )

IF WScript.Arguments.Length = 1 THEN
    strServer = WScript.Arguments( 0 )
ELSE
    strServer = "localhost"
END IF

WScript.Echo "Enumerating websites on " & strServer & CRLF
SET objWebService = GetObject( "IIS://" & strServer & "/W3SVC" )
EnumWebsites objWebService


SUB EnumWebsites( objWebService )
    DIM objWebServer, objVirtualDir, strBindings, secondDir

    FOR EACH objWebServer IN objWebService
        IF objWebserver.Class = "IIsWebServer" THEN
            WScript.Echo _
                "Site ID = " & objWebserver.Name & CRLF & _
                "Comment = """ & objWebServer.ServerComment & """ " & CRLF & _
                "State   = " & State2Desc( objWebserver.ServerState ) & CRLF & _
                "LogDir  = " & objWebServer.LogFileDirectory & _
                ""
			
			IterateVDirs objWebserver

            ' Enumerate the HTTP bindings (ServerBindings) and
            ' SSL bindings (SecureBindings)
            strBindings = EnumBindings( objWebServer.ServerBindings ) & _
                          EnumBindings( objWebServer.SecureBindings )
            IF NOT strBindings = "" THEN
                WScript.Echo "IP Address" & TAB & _
                             "Port" & TAB & _
                             "Host" & CRLF & _
                             strBindings
            END IF
        END IF
    NEXT

END SUB

SUB IterateVDirs( objRoot )
	DIM Child
	
	FOR EACH Child IN objRoot
		IF Child.Class = "IIsWebVirtualDir" THEN		
			Wscript.Echo vbTab & Child.Name & " -- " & Child.AuthFlags
			if (Child.AuthFlags AND 1) then wscript.echo vbtab & vbTab & Child.AnonymousUserName
			IterateVDirs Child
		END IF
		IF Child.Class = "IIsWebDirectory" THEN		
			Wscript.Echo vbTab & Child.Name & " -- " & Child.AuthFlags
			if (Child.AuthFlags AND 1) then wscript.echo vbtab & vbTab & Child.AnonymousUserName
			IterateVDirs Child
		END IF
	NEXT
	
END SUB


FUNCTION EnumBindings( objBindingList )
    DIM i, strIP, strPort, strHost
    DIM reBinding, reMatch, reMatches
    SET reBinding = NEW RegExp
    reBinding.Pattern = "([^:]*):([^:]*):(.*)"

    FOR i = LBOUND( objBindingList ) TO UBOUND( objBindingList )
        ' objBindingList( i ) is a string looking like IP:Port:Host
        SET reMatches = reBinding.Execute( objBindingList( i ) )
        FOR EACH reMatch IN reMatches
            strIP = reMatch.SubMatches( 0 )
            strPort = reMatch.SubMatches( 1 )
            strHost = reMatch.SubMatches( 2 )

            ' Do some pretty processing
            IF strIP = "" THEN strIP = "All Unassigned"
            IF strHost = "" THEN strHost = "*"
            IF LEN( strIP ) < 8 THEN strIP = strIP & TAB

            EnumBindings = EnumBindings & _
                           strIP & TAB & _
                           strPort & TAB & _
                           strHost & TAB & _
                           ""
        NEXT

        EnumBindings = EnumBindings & CRLF
    NEXT

END FUNCTION

FUNCTION State2Desc( nState )
    SELECT CASE nState
    CASE 1
        State2Desc = "Starting (MD_SERVER_STATE_STARTING)"
    CASE 2
        State2Desc = "Started (MD_SERVER_STATE_STARTED)"
    CASE 3
        State2Desc = "Stopping (MD_SERVER_STATE_STOPPING)"
    CASE 4
        State2Desc = "Stopped (MD_SERVER_STATE_STOPPED)"
    CASE 5
        State2Desc = "Pausing (MD_SERVER_STATE_PAUSING)"
    CASE 6
        State2Desc = "Paused (MD_SERVER_STATE_PAUSED)"
    CASE 7
        State2Desc = "Continuing (MD_SERVER_STATE_CONTINUING)"
    CASE ELSE
        State2Desc = "Unknown state"
    END SELECT

END FUNCTION

