<%@ CodePage=65001 %>

<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
if Session("StudioID") = "" then
%>
<script type="text/javascript">
	parent.resetSession();
</script>
<%
else
%>
		<!-- #include file="inc_init_functions.asp" -->
		<!-- #include file="inc_dbconn_wsMaster.asp" -->
		<!-- #include file="adm/inc_accpriv.asp" -->
		<!-- #include file="inc_i18n.asp" -->
		<!-- #include file="adm/inc_chk_ss.asp" -->
<%
	dim cnML : set cnML = Server.CreateObject("ADODB.Connection")
	cnML.CommandTimeout = 90
	cnML.Open = getMasterLogPath()

	Dim tmpUserName, curSchType
	session("UserLoc") = request.form("optUserLocation")
    if session("UserLoc") <> 0 then
        session("curLocation") = session("UserLoc")
    end if

	set rsEntry = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT CLIENTS.FirstName, CLIENTS.LastName, CLIENTS.Status, CLIENTS.LoginName FROM CLIENTS WHERE CLIENTS.[Deleted]=0 AND CLIENTS.ClientID=" & Session("mvarUserId")
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnMB
	Set rsEntry.ActiveConnection = Nothing


		Session("mvarNameFirst") = rsEntry("FirstName")
		Session("mvarNameLast") = rsEntry("LastName")
		Session("Pass") = true
		Session("Admin") = rsEntry("status")
		tmpUserName = rsEntry("LoginName")
	
	rsEntry.close
	
	' BJD: 3/30/09 - Check default tab here
	' BJD: 51_2816 - Default Launch Tab ID - IF not prompting for loc
	if Session("Admin")<>"" then
		strSQL = "SELECT DefaultLaunchTabID FROM tblAccessPriv WHERE status=N'" & session("Admin") & "'"
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing

		if NOT rsEntry.EOF then
			if rsEntry("DefaultLaunchTabID")<>0 then
				curSchType = -1*rsEntry("DefaultLaunchTabID")
			end if
		else ' no default found - straight through
			curSchType = request.form("stype")
		end if
		rsEntry.close
	end if


	if Session("Admin")<>"sa" AND environmentName<>"BU" then
		strSQL = "INSERT INTO EntryTimes (ClientID, LogInName, EntryDateTime) VALUES ("
		strSQL = strSQL & -1
		strSQL = strSQL & ", N'" & sqlInjectStr(Session("mvarNameFirst")) & " " & sqlInjectStr(Session("mvarNameLast")) & "', " & DateSep & DateAdd("n", Session("tzOffset"),Now) & DateSep
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
		strSQL = strSQL & ", N'" & sqlInjectStr(tmpUserName) & "', " & DateSep & Now & DateSep
		strSQL = strSQL & ", " & Session("studioID")
		strSQL = strSQL & ", N'" & getIPAddress & "'"
		strSQL = strSQL & ", N'" & Request.ServerVariables("HTTP_USER_AGENT") & "'"
		strSQL = strSQL & ", 0"	'failedLogin - false
		strSQL = strSQL & ", 1"	'code 1 for core sw site
		strSQL = strSQL & ", N'" & sqlInjectStr(session("Admin")) & "'"
		strSQL = strSQL & ")"
		cnML.Execute strSQL 
	end if	'''NOT SA

%>
		<html>
		<head>
		<title><%=Session("StudioName")%> Online</title>
		<meta http-equiv="Content-Type" content="text/html">
		<script type="text/javascript">
			function launchHome() {
				document.wsLaunch.submit();
			}
		</script>
		</head>
		<body onLoad="launchHome();" style="background-color:#FFFFFF;" text="#000000">
		<% if checkStudioSetting("tblGenOpts", "TrackCashRegisters") then %>
		<form name="wsLaunch" action="cash_register_sel.asp" method="post">
		<% else %>
		<form name="wsLaunch" action="adm/home.asp" method="post" target="_top">
		<% end if %>
		<%=buildLinkVars(true) %>

		</form>
		</body>
		</html>
<%	
	cnWS.close
	set cnWS = nothing
	cnMB.close
	Set cnMB = Nothing
	cnML.close
	Set cnML = Nothing

	end if '''session expired
	response.end
%>
