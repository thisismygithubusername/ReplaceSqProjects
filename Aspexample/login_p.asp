<%@ CodePage=65001 %>

<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1

' Careful this will require COM_INTEROP since inc_dbconn_str.asp isn't loaded with the DISABLE_COM_INTEROP() var
'logIt "loading login_p...."
%>
<!-- #include file="init.asp" -->
		<!-- #include file="inc_dbconn_wsMaster.asp" -->
		<!-- #include file="inc_dbconn_wsSession.asp" -->
		<!-- #include file="inc_dbconn_regions.asp" -->
		<!-- #include file="inc_mb_alert.asp" -->
		<!-- #include file="adm/inc_accpriv.asp" -->
		<!-- #include file="inc_i18n.asp" -->
		<!-- #include file="adm/inc_chk_ss.asp" -->
		<!-- #include file="adm/inc_ws_stats.asp" -->
		<!-- #include file="adm/inc_crypt.asp" -->
        <!-- #include file="adm/inc_masterclients_util.asp" -->
		<!-- #include file="adm/inc_chk_membership.asp" -->
<%
dim cmdWS, rsUser, rsEntry, rsEntry2, rsEntry3, cltLoc, pHomeStudio, varIPAddress, smodeLoc, ap_selsaleloc, ap_restrictIP, ss_RestrictIP, smodeLoginOk, tmpStr, tmpArr, tmpCounter, promptForLoc, ss_CltSecVerLoginOk, loginPendingSecVer, IPRestrictMultiLocs, promptForRegister, locationActive, badIP, invalidLoc, strPWDEnc, lockOut, lockOutNotice, pwdExpired, mbnetPost, strUsername, ss_ConsumerModeVersion, newCRClientID, VARRestricted, studioVARID
dim fbid, fbAccessToken, noLoginPermision, siteDeactivated

dim cnML : set cnML = Server.CreateObject("ADODB.Connection")
cnML.CommandTimeout = 90
cnML.Open = getMasterLogPath()

noLoginPermision = false
badIP = false
lockOut = false
lockOutNotice = false
pwdExpired = false
VARRestricted = false
studioVARID = 1
cltLoc = "98"
IPRestrictMultiLocs = 0
set cmdWS = Server.CreateObject("ADODB.Command")
set rsUser = Server.CreateObject("ADODB.Recordset")
set rsEntry = Server.CreateObject("ADODB.Recordset")
set rsEntry2 = Server.CreateObject("ADODB.Recordset")
set rsEntry3 = Server.CreateObject("ADODB.Recordset")
smodeLoginOk = true
locationActive = true
promptForLoc = false
loginPendingSecVer = false
session("noTracking") = false
strUsername = Request.Form("requiredtxtUserName")


' BJD: 11/24/08 - was the page posted from mbnet?
if request.form("mbnetPost")<>"" then
	mbnetPost = true
end if
strSQL = "SELECT CltSecVerLoginOk, TrackCashRegisters, UpdateCltHomeStudioOnLogin, ConsumerModeVersion, Studios.SiteDeactivated FROM tblGenOpts INNER JOIN Studios ON Studios.StudioID = tblGenOpts.StudioID WHERE Studios.StudioID=" & session("StudioID")
rsUser.CursorLocation = 3
rsUser.open strSQL, cnWS
Set rsUser.ActiveConnection = Nothing
	ss_CltSecVerLoginOk = rsUser("CltSecVerLoginOk")
	promptForRegister = rsUser("TrackCashRegisters")
	ss_UpdateCltHomeStudioOnLogin = rsUser("UpdateCltHomeStudioOnLogin")
	ss_ConsumerModeVersion = rsUser("ConsumerModeVersion")
	siteDeactivated = cbool(rsUser("SiteDeactivated"))
rsUser.close


'This is called on every successful Consumer Mode login to clear the flags
'involved with a password reset
'
'@param clientID The clientID of the client who just logged into Consumer mode
Function ResetPasswordChangeRequest (clientID)
	dim resetSQL	
	resetSQL =	"UPDATE CLIENTS " &_
				" SET PasswordChangeKey = NULL, ChangePassword = 0 " &_
				" WHERE ClientID = " & clientID		
	cnWS.Execute resetSQL	
End Function


if strUsername<>"" then
	'Lockout Check
	strSQL = "SELECT COUNT(*) AS NumFailedLogins FROM tblEntryLogs WITH (NOLOCK) WHERE (StudioID = ?) AND (EntryDateTime > DATEADD(minute, - 30, GETDATE())) AND (IPaddr = ?) AND (LoginName = ?) AND (FailedLogin = 1)"	
	cmdWS.ActiveConnection = cnML
	cmdWS.CommandText = strSQL
	cmdWS.CommandType = 1	'adCmdText
	Set rsUser = cmdWS.Execute(,array(session("StudioID"),getIPAddress,strUsername))
	Set cmdWS = Nothing
	if NOT rsUser.EOF then
		if rsUser("NumFailedLogins")>= 6 then
			lockOut = true
		end if
		if rsUser("NumFailedLogins")= 4 then
			lockOutNotice = true
		end if
	end if
	rsUser.close


    
    'VAR Check
    strSQL = "SELECT Studios.VarID FROM Studios WHERE Deleted=0 AND StudioID = " & session("StudioID")
	rsUser.CursorLocation = 3
	rsUser.open strSQL, cnMB
	Set rsUser.ActiveConnection = Nothing
	if NOT rsUser.EOF then
        studioVARID = rsUser("VarID")
    end if
	rsUser.close

end if

if request.form("requiredtxtPassword")<>"" then
	strPWDEnc = request.form("requiredtxtPassword")
	'	logIt "attempting login with " & strUsername & "/" & strPWDEnc
	strPWDEnc = DES_Encrypt(strPWDEnc, false, cnMB)
    'response.write(strPWDEnc)
    'response.end
end if

if Not lockOut then
	if strUsername<>"" and strPWDEnc<>"" then ' normal login (Business mode)
		strSQL = "SELECT CLIENTS.FirstName, CLIENTS.LastName, CLIENTS.Status, CLIENTS.ClientID, CLIENTS.LoginName, CLIENTS.LocationID, CLIENTS.LocIDStr, CLIENTS.MBORestrictIP, CLIENTS.MBOAccessGroup, CLIENTS.PWDChangeDate, tblMBOAccessPriv.VarID FROM CLIENTS LEFT OUTER JOIN tblMBOAccessPriv ON CLIENTS.MBOAccessGroup = tblMBOAccessPriv.MBOAccessGroup "
		strSQL = strSQL & "WHERE (CLIENTS.LoginName = N'" & sqlInjectStr(strUsername) & "')"
		strSQL = strSQL & " AND (  CONVERT(varbinary(100), CONVERT(char(100), CLIENTS.Password1)) = CONVERT(varbinary(100), CONVERT(char(100), N'" & sqlInjectStr(strPWDEnc)  & "')) "
		strSQL = strSQL & " AND CLIENTS.[Deleted]=0) AND ((CLIENTS.StudioID = " & Session("StudioID") & ") OR (NOT CLIENTS.MBOAccessGroup IS NULL))"
		rsUser.CursorLocation = 3
		rsUser.open strSQL, cnMB
		Set rsUser.ActiveConnection = Nothing
		
		session("MBOAccessGroup") = ""
		session("smodeLocationID") = ""

		if Not rsUser.EOF then	'Found Biz / MBO Admin Login

			session("MBOAccessGroup") = rsUser("MBOAccessGroup")
			session("smodeLocationID") = rsUser("LocationID")

			if rsUser("status")<>"sa" AND rsUser("status")<>"owner" then	'biz mode login
				smodeLoc = rsUser("LocationID")
				''query for vals from tblAccessPriv here
				strSQL = "SELECT TB_SALESELLOC, RestrictIP, CanLogin FROM tblAccessPriv WHERE Status=N'" & sqlInjectStr(rsUser("status")) & "' AND StudioID=" & Session("StudioID")
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS, 0, 1
				Set rsEntry.ActiveConnection = Nothing
				if NOT rsEntry.EOF then
					if NOT rsEntry("CanLogin") then
						noLoginPermision = true
					end if
					ap_selsaleloc = rsEntry("TB_SALESELLOC")
					ap_restrictIP = rsEntry("RestrictIP")
				else
					ap_selsaleloc = false
					ap_restrictIP = false
				end if
				rsEntry.close
				ss_RestrictIP = checkStudioSetting("tblGenOpts", "RestrictIP")
				if ss_RestrictIP AND ap_restrictIP  then
					badIP = true
					smodeLoginOk = false
					'REVERTED DUE TO LOGIN ISSUE IN PURE YOGA TWN CB 3/18/09 Updated to always query for active locations regardless of session("numLocations")
					if session("numLocations")<=1 then
						strSQL = "SELECT IPaddress, LocationID FROM tblIPs WHERE StudioID=" & session("StudioID") & " AND IPaddress=N'" & getIPAddress & "'"
					else
						strSQL = "SELECT tblIPs.IPaddress, tblIPs.LocationID FROM tblIPs INNER JOIN Location ON Location.LocationID = tblIPs.LocationID WHERE StudioID=" & session("StudioID") & " AND IPaddress=N'" & getIPAddress & "' AND Location.Active = 1"
					end if
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing

					if NOT rsEntry.EOF then
						do while NOT rsEntry.EOF
							if ap_selsaleloc OR session("numLocations")<=1 then
								smodeLoginOk = true
								session("UserLoc") = 0
							else	'NOT AP_SALELOC AND Multi-Loc
								tmpStr = rsUser("LocIDStr")
								tmpArr = Split(tmpStr,",")
								tmpCounter = 0
								do While tmpCounter < UBound(tmpArr)+1
									if TRIM(tmpArr(tmpCounter))=CSTR(rsEntry("LocationID")) then
										smodeLoginOk = true
										session("UserLoc") = rsEntry("LocationID")
                                        if rsEntry("LocationID") <> 0 then
                                            session("curLocation") = rsEntry("LocationID")
                                        end if
										IPRestrictMultiLocs = IPRestrictMultiLocs + 1
										badIP = false
									end if
									tmpCounter = tmpCounter + 1
								loop
								if NOT smodeLoginOk then
									'loginFailed(Unauthorized Location)
								end if
							end if
							rsEntry.MoveNext
						loop
						if smodeLoginOk AND IPRestrictMultiLocs>1 then
							promptForLoc = true
							Session("mvarUserId") = rsUser("ClientID")
							Session("mvarLoginName") = rsUser("LoginName")

							rsEntry.close
							'Check for relate to Trainer
							strSQL = "SELECT TrainerID FROM TRAINERS WHERE smodeID=" & session("mvaruserID")
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS, 0, 1
							Set rsEntry.ActiveConnection = Nothing
							if NOT rsEntry.EOF then
								session("empID") = rsEntry("TrainerID")
							else
								session("noTracking") = true
							end if
						end if
					else	'User IP NOT in IPList or location not active
						smodeLoginOk=false
						locationActive = false
						'loginFailed(BadIP)
					end if
					rsEntry.close
				else	'NOT Restricted IP
					if ap_selsaleloc OR session("numLocations")<=1 then
						session("UserLoc") = 0
					else	'NOT AP_SALELOC AND Multi-Loc
						tmpStr = rsUser("LocIDStr")
						tmpArr = Split(tmpStr,",")

						strSQL = "SELECT LocationID FROM Location WHERE LocationID IN ( " & rsUser("LocIDStr") & ") AND Location.Active = 1 "
						rsEntry2.CursorLocation = 3
						rsEntry2.open strSQL, cnWS, 0, 1
						Set rsEntry2.ActiveConnection = Nothing

						if rsEntry2.recordCount > 1 then 'More than 1 work location
							'''PROMPT USER TO SELECT LOCATION
							Session("mvarUserId") = rsUser("ClientID")
							Session("mvarLoginName") = rsUser("LoginName")
							promptForLoc = true
							'Check for relate to Trainer
							strSQL = "SELECT TrainerID FROM TRAINERS WHERE smodeID=" & session("mvaruserID")
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS, 0, 1
							Set rsEntry.ActiveConnection = Nothing
							if NOT rsEntry.EOF then
								session("empID") = rsEntry("TrainerID")
							else
								session("noTracking") = true
							end if
							rsEntry.close
						elseif rsEntry2.recordCount = 1 then
							'1 work location
							if TRIM(tmpArr(0))="0" then
								session("noTracking") = true
								session("UserLoc") = 0
							else
								session("UserLoc") = TRIM(tmpArr(0))
                                if TRIM(tmpArr(0)) <> 0 then
                                    session("curLocation") = TRIM(tmpArr(0))
                                end if
							end if
						else
							smodeLoginOk=false
							locationActive = false
							'loginFailed(No Active work locations)
						end if
						rsEntry2.close
					end if
				end if	'End Restricted IP Check
			else	'owner - mbo admin login
			
				session("MBOAdmin") = rsUser("MBOAccessGroup")

				if NOT isNULL(rsUser("MBOAccessGroup")) AND rsUser("MBORestrictIP") then
				''MBO Admin Login with Restricted IP
					if getIPAddress<>"64.4.159.33" AND getIPAddress<>"72.29.164.26" AND getIPAddress<>"209.203.114.154" AND Left(getIPAddress,9) <> "72.29.184" then
						smodeLoginOk = false
					end if
				end if

				if rsUser("status")<>"owner"then	'mbo admin - check for expired password
					if isNULL(rsUser("PWDChangeDate")) then
						smodeLoginOk = false
						pwdExpired  = true
					elseif rsUser("PWDChangeDate") < DateAdd("y", -90, Date) then
						smodeLoginOk = false
						pwdExpired  = true
					else

						'Check VAR Restriction
						if NOT isNULL(rsUser("VarID")) then 'null = full access
							if rsUser("VarID")<>studioVARID then
								smodeLoginOk = false
								VARRestricted  = true
							end if
						end if
	                
					end if
				end if

			end if	'NOT sa/owner
			
			'if smodeLoginOk AND NOT promptForLoc then	'CB - UPDATED 11/22 Release was letting promptForLoc thru when site deactivated
			if smodeLoginOk then
				if siteDeactivated AND isNULL(rsUser("MBOAccessGroup")) then

				'removed alert since only super admin can login if site is disactivated, added Site Deactivated message

				elseif NOT promptForLoc then
					Session("mvarNameFirst") = rsUser("FirstName")
					Session("mvarNameLast") = rsUser("LastName")
					Session("Pass") = true
					Session("mvarUserId") = rsUser("ClientID")
					Session("mvarLoginName") = rsUser("LoginName")
					Session("Admin") = rsUser("status")

					'hack for Chet Brandenburg, Blake Davis, Susan Figueroa, and Nicole Sell :: wsMaster dbo.Clients.ClientID
                    'mbsw, mbswEU, mbswUK, mbswAU
					if ((session("studioID")="-111" OR session("studioID")="-110" OR session("studioID")="-109" OR session("studioID")="-108") AND (session("mvaruserID")="3545" OR session("mvaruserID")="13716" OR session("mvaruserID")="38914" OR session("mvaruserID")="31160"))  then
						Session("Admin") = "sa"
						session("MBOAdmin") = "SA"
					end if

					'Get last login date/time
					Session("LastLogon") = "n/a"
					if Session("mvarUserId")<>"" then
						strSQL = "SELECT MAX(EntryDateTime) AS LastLogon FROM tblEntryLogs WHERE (ClientID = " & Session("mvarUserId") & ") AND (StudioID = " & Session("StudioID") & ")"
						rsEntry.open strSQL, cnML
						if NOT rsEntry.EOF then
							if NOT isNULL(rsEntry("LastLogon")) then
								Session("LastLogon") = FmtDateTime(DateAdd("n", Session("tzOffset"),rsEntry("LastLogon")))
							end if
						end if
						rsEntry.close
					end if

					if Session("Admin")<>"sa" AND environmentName<>"BU" then
						strSQL = "INSERT INTO EntryTimes (ClientID, LogInName, EntryDateTime) VALUES ("
						strSQL = strSQL & -1
						strSQL = strSQL & ", N'" & sqlInjectStr(Session("mvarNameFirst")) & " " & sqlInjectStr(Session("mvarNameLast")) & "'"
						strSQL = strSQL & ", " & DateSep & DateAdd("n", Session("tzOffset"),Now) & DateSep
						strSQL = strSQL & ")"
						cnWS.Execute strSQL 
					end if	'''NOT SA
					if environmentName<>"BU" then
						strSQL = "INSERT INTO tblEntryLogs (ClientID, Location, LogInName, EntryDateTime, StudioID, IPAddr, browser, FailedLogin, AppIDCode, AccessGroup) VALUES ("
						if Session("mvarUserId")<>"" then
							strSQL = strSQL & Session("mvarUserId")
						else
							strSQL = strSQL & "0"
						end if
						strSQL = strSQL & ", N'" & cltLoc & "'"
						strSQL = strSQL & ", N'" & sqlInjectStr(left(trim(strUsername), 140)) & "'"
						strSQL = strSQL & ", " & DateSep & Now & DateSep
						strSQL = strSQL & ", " & Session("studioID")
						strSQL = strSQL & ", N'" & sqlInjectStr(getIPAddress) & "'"
						strSQL = strSQL & ", N'" & sqlInjectStr(Request.ServerVariables("HTTP_USER_AGENT")) & "'"
						strSQL = strSQL & ", 0"	'failedLogin - false
						strSQL = strSQL & ", 1"	'code 1 for core sw site
						strSQL = strSQL & ", N'" & sqlInjectStr(session("Admin")) & "'"
						strSQL = strSQL & ")"
						'response.write debugSQL(strSQL, "SQL")
						cnML.Execute strSQL 
						
						if session("Admin")="sa" AND environmentName="PRD" then	'mbo admin - set session expire to 20 minutes
							strSQL = "UPDATE sessions SET ttl = 20 WHERE (guid = '" & getSessionGUID() & "')"
							cnSession.Execute strSQL
						end if
						
					end if	'''NOT SA
				end if	'Site Deactivated
			end if 	'SmodeLoginOk

			'BJD: 5/5/08 - set remember me cookie
			if request.form("optRememberMe")="on" then
				response.cookies("username") = strUsername
				response.cookies("username").domain = "mindbodyonline.com"
				response.cookies("username").expires = Date() + 999
			end if
		end if	''rs.EOF
		rsUser.close
	end if 'normal login
	
	'BJD: 11/25/08 - GUID check
	if request.form("launchGUID")<>"" then
		strSQL = "SELECT Clients.ClientID, Clients.LoginName, Clients.PasswordEnc, Clients.FirstName, Clients.LastName FROM tblClientAutoLogin INNER JOIN CLIENTS ON tblClientAutoLogin.ClientID=CLIENTS.ClientID WHERE Cast(GUID as nvarchar(36))=N'" & sqlInjectStr(request.form("launchGUID")) & "' AND (DATEDIFF(mi, TimeStamp, { fn NOW() }) < 30) "
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
	
		if NOT rsEntry.EOF then
			Session("mvarUserId") = rsEntry("ClientID")
			Session("mvarNameFirst") = rsEntry("FirstName")
			Session("mvarNameLast") = rsEntry("LastName")
			Session("mvarLoginName") = strUsername = rsEntry("LoginName")
			Session("Pass") = true
			Session("Admin") = "false"
			strPWDEnc = rsEntry("PasswordEnc")
		else
			strUsername = ""
			strPWDEnc = ""
		end if
		rsEntry.close
	else ' normal - check form vars
		strUsername = request.form("requiredtxtUserName")
		strPWDEnc = request.form("requiredtxtPassword")
		strPWDEnc = DES_Encrypt(strPWDEnc, false, null)
	end if
	
	'''''''''''''''''''NOT Business Mode, Check for Consumer MODE ''''''''''''''''''''''''''''''''''''''''''''''''''''

	' get facebook fields from form
	fbid = null : fbAccessToken = null
	if trim(request.Form("fbid")) <> "" and trim(request.Form("fbAccessToken")) <> "" then
		fbid = sqlInjectStr(trim(request.Form("fbid")))
		fbAccessToken = sqlInjectStr(trim(request.Form("fbAccessToken")))
	end if

	dim canDoCMLogin
	canDoCMLogin = (strUsername<>"" and strPWDEnc<>"") OR (not isNull(fbid) and not isNull(fbAccessToken))
	canDoCMLogin = canDoCMLogin and not session("Pass") AND NOT promptForLoc AND environmentName<>"BU" AND NOT pwdExpired AND NOT VARRestricted

	if canDoCMLogin AND NOT siteDeactivated then
		dim doingFaceBookLogin, doingRegularMBLogin, doFbExternalLoginConnect

		doFbExternalLoginConnect = (not isNull(fbid) and not isNull(fbAccessToken))
		doingFaceBookLogin = (not (strUsername<>"" and strPWDEnc<>"")) AND not IsNull(fbid) AND not IsNull(fbAccessToken)
		doingRegularMBLogin = (strUsername<>"" and strPWDEnc<>"")

		'create SQL select query string
		strSQL = "SELECT CLIENTS.ClientID, CLIENTS.Location, CLIENTS.HomeStudio, CLIENTS.LoginName, CLIENTS.FirstName, CLIENTS.LastName, CLIENTS.wspending,  Clients.dear " &_
				" FROM CLIENTS "
				if doingFaceBookLogin then
					strSQL = strSQL & " INNER JOIN tblClientFaceBook ON tblClientFaceBook.ClientID = CLIENTS.ClientID AND tblClientFaceBook.FBID = '" & fbid & "' "&_
											"AND tblClientFaceBook.fbAccessToken = '" & fbAccessToken & "' "
				end if
				'conditional where based on method of login regular, facebook, etc
				strSQL = strSQL & " WHERE CLIENTS.Deleted=0 "
				if doingRegularMBLogin then
					strSQL = strSQL & " AND (CLIENTS.LoginName=N'" & sqlInjectStr(strUsername) & "' OR CLIENTS.EmailName=N'" & sqlInjectStr(strUsername) & "') " &_
					" AND (CONVERT(varbinary(100), CONVERT(char(100), PasswordEnc)) = CONVERT(varbinary(100), CONVERT(char(100), N'" & sqlInjectStr(strPWDEnc) & "'))) "
				end if
		rsUser.CursorLocation = 3
		rsUser.open strSQL, cnWS, 0, 1
		Set rsUser.ActiveConnection = Nothing

		'CB 49_1125 - Username not found in local region and CR Logins Enabled
		if rsUser.EOF AND Session("CR_Login")<>0 then

			'(1) Verifiy Username Does Not Exisit in Current Site - this catches case where user mistypes username and prevents copying a client withe same username
			strSQL = "SELECT ClientID FROM CLIENTS WHERE (LoginName=N'" & sqlInjectStr(strUsername) & "' OR EmailName=N'" & sqlInjectStr(strUsername) & "') AND Deleted=0"
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			if rsEntry.EOF then	'username does not exist in local region - continue to search other regions
				rsEntry.close
				
				'(2) - Get List CR Sites where Username Exists
				strSQL = "SELECT Studios.StudioID, Studios.StudioShort "&_
					" FROM Studios "&_
					" INNER JOIN tblMasterClientsEmail ON tblMasterClientsEmail.StudioID = Studios.StudioID "&_
					" WHERE Studios.Deleted=0 AND (tblMasterClientsEmail.Email = N'" & sqlInjectStr(strUsername) &"' ) "&_
					" AND (Studios.RegionID = " & Session("CR_Login") & ") AND (Studios.StudioID<>" & Session("StudioID") & ") "
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnMB
				Set rsEntry.ActiveConnection = Nothing
				if NOT rsEntry.EOF then
	%>
		<!-- #include file="adm\inc_username.asp" -->
		<!-- #include file="adm\inc_masterclients.asp" -->
<%
					do while NOT rsEntry.EOF

						'(3) Connect to xRegDatabase
						connectToRegionalDB(rsEntry("StudioShort"))

						'(4) CR Decrypt password and authenticate
						strPWDEnc = request.form("requiredtxtPassword")
						strPWDEnc = DES_Encrypt(strPWDEnc, false, cnWSReg)

						strSQL = "SELECT ClientID, Location, HomeStudio, LoginName, FirstName, LastName, wspending FROM Clients WHERE EmailName=N'" & sqlInjectStr(strUsername) & "' AND Deleted=0"
						strSQL = strSQL & " AND (CONVERT(varbinary(100), CONVERT(char(100), PasswordEnc)) = CONVERT(varbinary(100), CONVERT(char(100), N'" & sqlInjectStr(strPWDEnc) & "')))"
						rsEntry2.CursorLocation = 3
						rsEntry2.open strSQL, cnWSReg
						Set rsEntry2.ActiveConnection = Nothing
						if NOT rsEntry2.EOF then	'Found CR Login

                            'CB 3/25/2010 - Determined this is buggy where it can create duplicate client profiles
                            'Need to check the master list here to see if this client already exists locally and then update/add their login to the existing profile rather than creating a new profile
                            
							'(5) Copy Client to Local Database and reset rsUser

							newCRClientID = addLocalClient(rsEntry2("ClientID"), rsEntry("StudioID"), "", "", "")

							if newCRClientID<>-1 then	'Client Added Successfully
								rsUser.close
								strSQL = "SELECT ClientID, Location, HomeStudio, LoginName, FirstName, LastName, wspending, Dear FROM Clients WHERE ClientID=" & newCRClientID
								rsUser.CursorLocation = 3
								rsUser.open strSQL, cnWS, 0, 1
								Set rsUser.ActiveConnection = Nothing
								exit do
							end if
						end if

						rsEntry2.close
						rsEntry.MoveNext
					loop
				end if
				rsEntry.close

			else	'username found in local region but password did not match
				rsEntry.close
			end if

		end if	'CB 49_1125 CR_Login

		'check rsUser for valid consumer mode user
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        dim multipleLogins : multipleLogins = false
		if NOT rsUser.EOF then
            if rsUser.RecordCount = 1 then
			    if NOT rsUser("wspending") OR ss_CltSecVerLoginOk then
				    Session("Pass") = true
				    Session("Admin") = "false"
				    if NOT rsUser("ClientId") = "" Then
					    Session("mvarUserId") = rsUser("ClientId")
					    Session("mvarLoginName") = rsUser("LoginName")

						if Request.Form("newLoginCreated") <> "true" then
							call ResetPasswordChangeRequest(Session("mvarUserId"))						
						end if
				    else
					    Session("mvarUserId") = ""
					    Session("mvarLoginName") = ""
				    end If

				    if rsUser("HomeStudio")<>"" AND session("numlocations")>1 then
					    pHomeStudio = rsUser("HomeStudio")
				    else
					    pHomeStudio = "1"
				    end If
				    if NOT rsUser("FirstName") = "" then
					    Session("mvarNameFirst") = rsUser("FirstName")
				    else
					    Session("mvarNameFirst") = ""
				    end if
				    if NOT rsUser("LastName") = "" then
					    Session("mvarNameLast") = rsUser("LastName")
				    else
					    Session("mvarNameLast") = ""
				    end if
				    'JM-55_3298
				    if NOT rsUser("Dear") = "" then
					    Session("mvarClientDear") = rsUser("dear")
				    else
					    Session("mvarClientDear") = ""
				    end if
				    rsUser.Close

				    'CB 7/25/08 - Membership Tier Day of Month Scheduling Restriction
				    Session("SchOpenDOMCloseDate") = ""
				    Session("SchOpenDOM") = getDayOfMonthMemberTierScheduleOpens(getClientMembershipTier(Session("mvarUserId")))
				    Dim tmpDOM : tmpDOM = Session("SchOpenDOM")
				    if tmpDOM <> "" then
					    if Day(DateAdd("n", Session("tzOffset"),Now)) >= CINT(tmpDOM) then	'restricted to end of next month
						    Session("SchOpenDOMCloseDate") = CDATE(Month(DateAdd("m",2, DateAdd("n", Session("tzOffset"),Now))) & "/1/" & Year(DateAdd("m",2, DateAdd("n", Session("tzOffset"),Now))) ) - 1
					    else	'restricted to end of current month
						    Session("SchOpenDOMCloseDate") = CDATE(Month(DateAdd("m",1, DateAdd("n", Session("tzOffset"),Now))) & "/1/" & Year(DateAdd("m",1, DateAdd("n", Session("tzOffset"),Now))) ) - 1
					    end if
				    end if

				    'response.write "tmpDOM:" & tmpDOM & "-" & Day(DateAdd("n", Session("tzOffset"),Now))
				    'response.write "<br />" & Session("SchOpenDOMCloseDate")
				    'response.end

				
				    'BJD: 5/5/08 - set remember me cookie
				    if request.form("optRememberMe")="on" then
					    response.cookies("username") = strUsername
					    response.cookies("username").domain = "mindbodyonline.com"
					    response.cookies("username").expires = Date() + 999
				    end if
		
					'************************
					varIPAddress = getIPAddress
					strSQL = "UPDATE Clients SET IPaddress =N'" & varIPAddress & "'"
					strSQL = strSQL & " WHERE ClientId=" & Session("mvarUserId")
					cnWS.Execute strSQL

				    'Update Home Studio
				    if ss_UpdateCltHomeStudioOnLogin AND request.form("optLocation")<>"" AND request.form("optLocation")<>"0" then
					    strSQL = "UPDATE Clients SET HomeStudio=" & request.form("optLocation")
					    strSQL = strSQL & " WHERE ClientId=" & Session("mvarUserId") & " AND HomeStudio=0"
					    cnWS.Execute strSQL 
				    end if
			
				    'Get last login date/time
				    Session("LastLogon") = "n/a"
				    if Session("mvarUserId")<>"" then
					    strSQL = "SELECT MAX(EntryDateTime) AS LastLogon FROM EntryTimes WHERE (ClientID = " & Session("mvarUserId") & ")"
					    rsEntry.open strSQL, cnWS
					    if NOT rsEntry.EOF then
						    if NOT isNULL(rsEntry("LastLogon")) then
							    Session("LastLogon") = FmtDateTime(rsEntry("LastLogon"))
						    end if
					    end if
					    rsEntry.close
				    end if

				    strSQL = "UPDATE Clients SET Inactive=0, ReactivatedTime=" & DateSep & DateAdd("n", Session("tzOffset"),Now) & DateSep & " WHERE ClientID=" & Session("mvarUserID") & " AND Inactive=1"
				    cnWS.Execute strSQL
						
				    if environmentName<>"BU" then
					    'updates customers entry time to masterDB
					    '#####################################################
					    strSQL = "INSERT INTO EntryTimes (ClientID, Location, LogInName, EntryDateTime) VALUES ("
					    strSQL = strSQL & Session("mvarUserId")
					    strSQL = strSQL & ", N'" & cltLoc &"'"
					    strSQL = strSQL & ", N'" & sqlInjectStr(strUsername) & "'"
					    strSQL = strSQL & ", " & DateSep & DateAdd("n", Session("tzOffset"),Now) & DateSep
					    strSQL = strSQL & ")"
					    cnWS.Execute strSQL
		
						strSQL = "INSERT INTO tblEntryLogs (ClientID, Location, LogInName, EntryDateTime, StudioID, IPAddr, browser, FailedLogin, AppIDCode, AccessGroup, fbLogin) VALUES ("
						strSQL = strSQL & Session("mvarUserId")
						strSQL = strSQL & ", N'" & cltLoc & "'"
						strSQL = strSQL & ", N'" & sqlInjectStr(left(trim(strUsername), 140)) & "'"
						strSQL = strSQL & ", " & DateSepSQL & Now & DateSepSQL
						strSQL = strSQL & ", " & Session("studioID")
						strSQL = strSQL & ", N'" & sqlInjectStr(getIPAddress) & "'"
						strSQL = strSQL & ", N'" & sqlInjectStr(Request.ServerVariables("HTTP_USER_AGENT")) & "'"
						strSQL = strSQL & ", 0"	'failedLogin - false
						strSQL = strSQL & ", 1"	'code 1 for core sw site
						strSQL = strSQL & ", N'" & sqlInjectStr(session("Admin")) & "'"
						if doFbExternalLoginConnect or doingFaceBookLogin then
							strSQL = strSQL & ", 1 "
						else
							strSQL = strSQL & ", 0 "
						end if
						strSQL = strSQL & ")"
						cnML.Execute strSQL 
					end if	'NOT environmentName<>"BU"
		
				    '''Update WS Metrics
				    updateStat 2, 1		''2 - Client Entries
					if doFbExternalLoginConnect then ' external facebook connect
						set jsonParams = JSON.parse("{}")
						jsonParams.set "ClientID", Session("mvarUserId")
						jsonParams.set "fbid", fbid
						jsonParams.set "fbAccessToken", fbAccessToken
						CallMethodWithJSON "mb.Core.Integrations.Facebook.FacebookConnectIntegration.ConnectFBUser", jsonParams
					end if
			    else	'Security Verification Pending - Studio doesn't allow to login
				    Session("Pass") =  false
				    Session("Admin") = "false"
				    loginPendingSecVer = true
			    end if
		    else ' More than one match
                Session("Pass") =  false
			    Session("Admin") = "false"
                multipleLogins = true
		    end if 'pending
        else ' rs.EOF - Login Not Matched
			Session("Pass") =  false
			Session("Admin") = "false"
        end if 
	end if ' EOF

	'END CONSUMER MODE LOGIN FUNCTION

	Dim curSchType, schDate, schLoc, formAction, formTarget, curView
	curSchType = ""


	schDate = ""
	Call SetLocale(session("mvarLocaleStr"))			
		if isDate(request.form("txtDate")) then
			schDate = CDATE(request.form("txtDate"))
		end if
	Call SetLocale("en-us")
	
	if session("numLocations")>1 AND session("Pass") AND session("Admin")="false" then	''multiloc consumer mode
		'CB 1/20/09 ClientModeLockLoc appears to be legacy
		'if checkStudioSetting("tblGenOpts","ClientModeLockLoc") OR request.form("optLocation")="0"  then
			schLoc = pHomeStudio
		'else
		'	schLoc = request.form("optLocation")				
		'end if
	else	'Single Location
		schLoc = session("curLocation")
	end if

	'Check for notracking / Set Sess EmpID
	if session("Admin")<>"false" AND session("Admin")<>"sa" AND session("Admin")<>"owner" then
		'Check for relate to Trainer
		strSQL = "SELECT TrainerID FROM TRAINERS WHERE smodeID=" & session("mvaruserID")
		rsUser.CursorLocation = 3
		rsUser.open strSQL, cnWS, 0, 1
		Set rsUser.ActiveConnection = Nothing

		if NOT rsUser.EOF then
			session("empID") = rsUser("TrainerID")
		else
			session("noTracking") = true
		end if
		rsUser.close
	end if
end if	'Lockout

if (session("studioid")="-111" OR session("studioid")="-110" OR session("studioid")="-109" OR session("studioid")="-1022" OR session("studioid")="-108") AND session("MBOAccessGroup")<>"" then
	Session("mvarNameFirst") = ""
	Session("mvarNameLast") = ""
	Session("Pass") = false
	Session("mvarUserId") = ""		'''Current ID Assigned to SA login
	Session("mvarLoginName") = ""		'''Current ID Assigned to SA login
	Session("Admin") = "false"
end if

' Check if this is the Owner of the site, if they have agreed to their contract, 
' and if not, if they have been prompted to agree to their contract in the last week
promptForContract = false
' Check the kill switch
if enableOwnerContracts then
    if session("Admin") = "owner" then
        ' Check the last query time
        strSQL = "SELECT lastOwnerContractCheck, BillMBOCltID FROM Studios WHERE Deleted=0 AND StudioID = " & session("StudioID")
	    rsEntry.CursorLocation = 3
	    rsEntry.open strSQL, cnMB
	    Set rsEntry.ActiveConnection = Nothing
        
        ' If there is a studio entry
	    if NOT rsEntry.EOF then
	        ' And that studio has a non-null billing ID
	        if NOT isNull(rsEntry("BillMBOCltID")) then
	            ' And that studio has a non-null last checked date
	            if NOT isNull(rsEntry("lastOwnerContractCheck")) then
	                'if the difference is 1 week or more since last prompt
		            if DateDiff("ww", rsEntry("lastOwnerContractCheck"), Now) > 0 then 
			            promptForContract = checkAgreedToContracts(session("StudioID"))
		            end if
		        ' Null contract date, prompt if you need to
		        else
		            ' contract check date is null, always prompt
		            promptForContract = checkAgreedToContracts(session("StudioID"))
		        end if
		    end if
	    end if
	    rsEntry.close    
    end if
end if

	'''Logged In
	if (session("Pass") OR promptForLoc) AND NOT noLoginPermision then

		'CB 49_2575
		session("username") = strUsername
		
		'BJD: 11/24/08 - determine wsLaunch form action and use below
		if promptForContract then
		    formAction = "/ASP/adm/adm_mb_contract.asp"
		    formTarget = "_parent"
		elseif promptForLoc then
			formAction = "/ASP/login_sel_loc.asp"
			formTarget = "mainFrame"
		elseif promptForRegister and session("Admin")<>"false" then
			strSQL = "SELECT CashRegisterName, CashRegisterID FROM tblCashRegister WHERE [Delete] = 0 "
			if session("UserLoc")<>"0" and session("UserLoc")<>"" then
				strSQL = strSQL & " AND LocationID = " & session("UserLoc")
			end if
			'JM - 48_2506
			strSQL = strSQL & " ORDER BY tblCashRegister.SortOrder, tblCashRegister.CashRegisterName"
			'response.write debugSQL(strSQL, "SQL")
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing

			if NOT rsEntry.EOF then
				rsEntry.close
				set rsEntry = nothing
				'session variable added for bug 6991
				session("trackCashRegisters") = true
				formAction = "/ASP/cash_register_sel.asp"
				formTarget = "mainFrame"
			else
				rsEntry.close
				set rsEntry = nothing
				formAction = "/ASP/adm/home.asp?studioid=" & session("StudioID")
				formTarget = "_parent"
			end if
' bug 3624		
'		elseif request.form("launchGUID")<>"" AND session("Pass")=true then
'			formAction = "home.asp?studioid=" & session("StudioID")
		else '' normal login
			if session("Admin")<>"false" then
				formAction = "/ASP/adm/home.asp?studioid=" & session("StudioID")
				formTarget = "_parent"
			else ' consumer mode login
				if mbnetPost then 'if posted by mbnet - send it page to the page that posted
					formAction = Request.ServerVariables("HTTP_REFERER")
				else
				
					formAction = "/ASP/home.asp?studioid=" & session("StudioID")

					formTarget = "_parent"
						
				end if
			end if
			
		end if
		
		if request.QueryString("isLibAsync") = "true" then
			dim url
			url = formAction

			dim target
			if (formTarget="_parent") then
				target = "parent"
			end if

			response.ContentType = "application/json"
			response.Write("{")
			response.Write("""success"": true,")
			response.Write("""json"": {""success"": true, ""url"":""" & jsEscDouble(url) & """, ""target"": """ & target & """, ""buildLinkVars"":""" & jsEscDouble(buildLinkVars(true)) & """}")
			response.Write("}")
			response.End
		end if
		
		'logIt "logged in as " & Session("mvarLoginName") & "...."
		
		
		if NOT DISABLE_3531() then
			'FIXME
			'dim initializeCOM : set initializeCOM = Server.CreateObject("mb.Core.Tools.InitializeCOM")
			'initializeCOM.SetLoginInfo()
		end if

%>
		<html>
		<head>
		<title><%=Session("StudioName")%> Online</title>
		<meta http-equiv="Content-Type" content="text/html">
<!-- #include file="inc_jquery.asp" -->		
<!-- #include file="common/common_auto_login_js.asp" -->		
		<script type="text/javascript">
			function launchHome() {
<% if NOT DISABLE_3531() then %>
alert('look for FIXME in login_p.asp');
window.currentUser = User();
currentUser.create(
				'<%= Session("studioID") %>',
				'<%= Session("mvarUserId") %>',
				'<% if Session("Admin")<>"false" AND Session("Admin")<>"" then Response.write("B") else Response.write("C") end if %>'
			);
updatePageInfoCookie();
updateSessionInfoCookie();
				getRootWindow().DONT_UPDATE_PAGE_INFO_COOKIE = true;

<% end if %>		
/*			
					setLoginInfo(
		'<%= Session("studioID") %>',
		'<%= Session("mvarUserId") %>',
		<% if Session("Admin")<>"false" AND Session("Admin")<>"" then Response.write("BIZ_MODE") else Response.write("CONSUMER_MODE") end if %>, 
		<% if Session("Admin")<>"sa" then Response.write(120) else Response.write(15) end if %>
	);	
*/
                document.wsLaunch.submit();
			}
		</script>
		</head>
		<body onLoad="launchHome();" style="background-color:#FFFFFF;" text="#000000">
		<form name="wsLaunch" action="<%=formAction%>" method="post" target="<%=formTarget%>">
		<%=buildLinkVars(true) %>

		<input type="hidden" name="justloggedin" value="true">
		<input type="hidden" name="newLoginCreated" value="<%=xssStr(request.form("newLoginCreated"))%>" />

		</form>
		</body>
		</html>
<% 
		'''Login Failed	
		else
			'Write failed login
			if environmentName<>"BU" and strUsername<>"" then
				strSQL = "INSERT INTO tblEntryLogs (ClientID, LogInName, EntryDateTime, StudioID, IPAddr, browser, FailedLogin, AppIDCode, AccessGroup) VALUES ("
				strSQL = strSQL & "null"
				strSQL = strSQL & ", N'" & sqlInjectStr(left(trim(strUsername), 140)) & "'"
				strSQL = strSQL & ", " & DateSepSQL & Now & DateSepSQL
				strSQL = strSQL & ", " & Session("studioID")
				strSQL = strSQL & ", N'" & sqlInjectStr(getIPAddress) & "'"
				strSQL = strSQL & ", N'" & sqlInjectStr(Request.ServerVariables("HTTP_USER_AGENT")) & "'"
				strSQL = strSQL & ", 1"	'failedLogin - true
				strSQL = strSQL & ", 1"	'code 1 for core sw site
				strSQL = strSQL & ", null"
				strSQL = strSQL & ")"
				session("strSQL") = strSQL
				cnML.Execute strSQL 
				session("strSQL") = ""
			end if

			Dim qstrparam,reason

			qstrparam = "?" &buildLinkVars(false)
			if loginPendingSecVer then
				reason = "secver"
			elseif lockOutNotice then
				reason = "lockNotice"
			elseif lockOut then
				reason = "locked"
			elseif pwdExpired then
				reason = "exp"
			elseif VARRestricted then
				reason = "var"
			elseif badIP then
				reason = "badIP"
			elseif not locationActive then
				reason = "badLoc"
			elseif multipleLogins then
				reason = "seeStudio"
			elseif noLoginPermision then
				reason = "loginPermision"
			elseif siteDeactivated then
				reason = "siteDeactivated"
			else
				reason = "false"
			end if
			qstrparam = qstrparam & "&login=" & reason

			if request.QueryString("isLibAsync") = "true" then
				response.ContentType = "application/json"
				response.Write("{")
				response.Write("""success"": true,")
				response.Write("""json"": {""success"": false, ""failureReason"":""" & reason & """}")
				response.Write("}")
				response.End
			elseif mbnetPost then ' mbnet post failed
				Response.Redirect(Request.ServerVariables("HTTP_REFERER")&qstrparam)
			elseif request.form("launchGUID")<>"" then ' guid failed - redirect to site not logged in
				Response.Redirect("/ws.asp?studioid=" & session("StudioID"))
			elseif request.form("sSU") ="true" then%>
				<script type="text/javascript">
				//change top frame location to top_page w/qstrparams
					parent.mainFrame.location= "SU1.asp<%=qstrparam%>";
				</script>
			<%else ' normal
%>
				<script type="text/javascript">
				//change top frame location to top_page w/qstrparams
					parent.mainFrame.location= "main<%=session("pageID")%>.asp<%=qstrparam%>";
				</script>
<%	
				
			end if ' not mbnet
		end if	''session("pass")

	' close down the connection to wsMaster
	cnWS.close
	set cnWS = nothing
	response.end
	cnMB.close
	Set cnMB = Nothing
	cnML.close
	Set cnML = Nothing
%>
