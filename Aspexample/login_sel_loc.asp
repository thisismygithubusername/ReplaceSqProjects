<%@ CodePage=65001 %>
<%
Option Explicit
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
if Session("StudioID") = "" then
%>
<script type="text/javascript">
	history.go(-1);
</script>
<%
else
%>
		<!-- #include file="init.asp" -->
		<!-- #include file="inc_dbconn.asp" -->
		<!-- #include file="inc_dbconn_wsMaster.asp" -->
		<!-- #include file="inc_i18n.asp" -->
		<!-- #include file="adm/inc_chk_ss.asp" -->
		<!-- #include file="inc_localization.asp" -->
<%
Dim rsEntry, tmpStr, tmpArr, tmpCounter, tmpfirst
tmpfirst = true

dim phraseDictionary
set phraseDictionary = LoadPhrases("LoginlocationselectionPage", 17)


	set rsEntry = Server.CreateObject("ADODB.Recordset")
	dim ss_BBenabled, ss_RestrictIP, ss_CltModeSigup, ss_HideCltHelp, ss_HideCltForgotPwd, ss_ClientModeLockTGRemoveBuy, mvStudioURL, mvStudioLinkTab, logoH, logoW, topBGClr, lockLoc
	mvStudioURL = ""
	mvStudioLinkTab = ""
	strSQL = "SELECT Studios.StudioURL, tblGenOpts.StudioLinkTab, tblGenOpts.HideCltHelp, tblGenOpts.HideCltForgotPwd, tblGenOpts.RestrictIP, tblAppearance.LogoHeight, tblAppearance.LogoWidth, tblAppearance.topBGColor, tblGenOpts.ClientModeLockTGRemoveBuy, tblGenOpts.CltModeSigup, tblGenOpts.CltModeSigupExisting, tblGenOpts.BBenabled, tblGenOpts.ClientModeLockLoc FROM Studios INNER JOIN tblGenOpts ON Studios.StudioID = tblGenOpts.StudioID INNER JOIN tblAppearance ON Studios.StudioID = tblAppearance.StudioID INNER JOIN tblApptOpts ON Studios.StudioID = tblApptOpts.StudioID WHERE (Studios.StudioID = " & session("StudioID") & ")"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
		''Standard
		mvStudioURL = TRIM(rsEntry("StudioURL")) & session("rtnURLqs")
		mvStudioLinkTab = TRIM(rsEntry("StudioLinkTab"))
		logoH = rsEntry("LogoHeight")
		logoW = rsEntry("LogoWidth")
		topBGClr = rsEntry("topBGColor")
		ss_BBenabled = rsEntry("BBenabled")
		ss_CltModeSigup = rsEntry("CltModeSigup") OR rsEntry("CltModeSigupExisting")
		ss_HideCltHelp = rsEntry("HideCltHelp")
		ss_HideCltForgotPwd = rsEntry("HideCltForgotPwd")
		ss_ClientModeLockTGRemoveBuy = rsEntry("ClientModeLockTGRemoveBuy")
		ss_RestrictIP = rsEntry("RestrictIP")
		lockLoc = rsEntry("ClientModeLockLoc")
	rsEntry.close
%>
<html>
<head>
<title><%=Session("StudioName")%> Online</title>
<meta http-equiv="Content-Type" content="text/html">
	<!-- #include file="inc_top_js.asp" -->
	<style type="text/css">
	#overlayWrap
	{
		height:100%;
		z-index: 1002;
		width: 100%;
	}
	
	#darkDiv:before {
		content: ".";
		display: block;
		height: 0;
		visibility: hidden;
	}

	#darkDiv
	{
		z-index:1;
		width:100%;
		background-color:#000000;
		position:absolute;
		filter:alpha(opacity=50);
		-moz-opacity:0.50;
		overflow:hidden;
		opacity:0.50;
		height:100%;
	}
	#alertdivWrap
	{
		padding: 20px 0px;
	}
	#alertdiv {
		position:relative;
		z-index:2;
		overflow:auto;
		background-color: #FFFFFF;
		border:1px solid #666; 
		padding: 20px 30px 5px 30px;
		width:600px;
		margin: 0px auto ;
	}
</style>
</head>
<body style="background-color:#FFFFFF;" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div id="overlayWrap">

<div id="darkDiv">
</div>
<div id="alertdivWrap">
<div id="alertdiv">
<div align="left">
  <table width="<%=strPageWidth%>" cellspacing="0" height="100%">
    <form name="frmSelectLoc" method="post" target="mainFrame" action="login_sel_loc_p.asp">
		<%=buildLinkVars(true) %>
      <tr> 
        <td align="left" valign="top">
		<table class="mainText" width="100%"  height="60" cellspacing="0">
		  <tr valign="middle" height="100%">
			<td class="center-ch"><strong><%=DisplayPhrase(phraseDictionary,"Selectloc")%>: 
			    <select name="optUserLocation">
<%
	Dim tmpStatus, ipRestricted
	''check if user group has restricted IP and if so then filter avail locs to that IP
	strSQL = "SELECT CLIENTS.Status FROM CLIENTS WHERE CLIENTS.[Deleted]=0 AND CLIENTS.ClientID=" & Session("mvarUserId")
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnMB
	Set rsEntry.ActiveConnection = Nothing
	if NOT rsEntry.EOF then
		tmpStatus = rsEntry("Status")
	else
		response.end
	end if
	rsEntry.close

	strSQL = "SELECT RestrictIP FROM tblAccessPriv WHERE Status=N'" & tmpStatus & "'"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	if NOT rsEntry.EOF then
		ipRestricted = rsEntry("RestrictIP")
	else
		response.end
	end if
	rsEntry.close

	strSQL = "SELECT LocIDStr FROM CLIENTS WHERE ClientID=" & Session("mvarUserId")
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnMB
	Set rsEntry.ActiveConnection = Nothing
	if NOT rsEntry.EOF then
		tmpStr = rsEntry("LocIDStr")
		tmpArr = Split(tmpStr,",")
		tmpCounter = 0
	else
		response.end
	end if
	rsEntry.close
	
	if ss_RestrictIP AND ipRestricted  then
		strSQL = "SELECT Location.LocationName, Location.LocationID FROM tblIPs INNER JOIN Location ON tblIPs.LocationID = Location.LocationID WHERE tblIPs.IPaddress=N'" & getIPAddress & "' AND " 
	else
		strSQL = "SELECT LocationName, LocationID FROM Location WHERE "
	end if

	do While tmpCounter < UBound(tmpArr)+1
		if tmpfirst then
			strSQL = strSQL & " (Location.LocationID=" & TRIM(tmpArr(tmpCounter))
			tmpfirst = false		
		else
			strSQL = strSQL & " OR Location.LocationID=" & TRIM(tmpArr(tmpCounter))
		end if
		tmpCounter = tmpCounter + 1
	loop
	strSQL = strSQL & ") ORDER BY Location.LocationName"

	'response.write debugSQL(strSQL, "SQL")

	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing

	do while not rsEntry.EOF	
%>			
					<option value="<%=rsEntry("LocationID")%>"><%=rsEntry("LocationName")%></option>
<%
		rsEntry.MoveNext
	loop
%>			
			</select>
                <input name="Continue" type="submit" value="Continue">
			</strong></td>
		  </tr>
		</table>
		</td>
      </tr>
    </form>
  </table></div>
    </div>
  </div>
  </div>
</body>
</html>
<%
	set rsEntry = nothing
	cnWS.close
	set cnWS = nothing
	cnMB.close
	Set cnMB = Nothing
	end if 	''' Session Expired '''
%>
