<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
	dim rsEntry
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	dim rsEntry2
	set rsEntry2 = Server.CreateObject("ADODB.Recordset")
	%>
	<!-- #include file="inc_accpriv.asp" -->
	<%
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_CONLOG") then 
	%>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
	<%else%>
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_rpt_tagging.asp" -->
		<!-- #include file="inc_utilities.asp" -->
		<!-- #include file="inc_rpt_save.asp" -->
        <% dim doRefresh : doRefresh = false %>
        <!-- #include file="inc_date_arrows.asp" -->
	<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

	Dim showDetails, cSDate, cEDate, cCSDate, cCEDate, PercentAttend, cLoc, onlyOwnStudio
	Dim showHeader, rowcolor, barcolor, curTrainer, TotAttend, GTotAttend, first, cont, tmpContactLogID
	dim conLogTypeIndex, firstCLT, displaycSDate, displaycEDate, ss_UseContactLogForecasting, ss_ContactLogSubtypes, resultCount
	
	dim category : category = ""
	if (RQ("category"))<>"" then
		category = RQ("category")
	elseif (RF("category"))<>"" then
		category = RF("category") 
	end if

	dim strStudioShort, strServerName, strLink
	strStudioShort = ""
	strLink = ""
	strServerName = ""
	onlyOwnStudio = true
	if validAccessPriv("TB_V_RPT_ALL_LOC") OR validAccessPriv("TB_SALESELLOC") then
		onlyOwnStudio = false
	end if
	ss_UseContactLogForecasting = checkStudioSetting("tblGenOpts", "UseContactLogForecasting")
	ss_ContactLogSubtypes = checkStudioSetting("tblGenOpts", "ContactLogSubtypes") 

	if request.QueryString("studioshort") <> "" then
		strStudioShort = request.QueryString("studioshort")
	else
		if Session("StudioShort") <> "" then
			strStudioShort = Session("studioShort")
		end if
	end if
	
	if request.querystring("clearTaggedClts")="true" then
		strSQL = "DELETE FROM tblClientTag "
		if session("mvarUserID")<>"" then
			strSQL = strSQL & "WHERE smodeID = " & session("mvarUserID") & ""
		else
			strSQL = strSQL & "WHERE smodeID = 0"
		end if
		cnWS.execute strSQL
		%>
		<script>
			alert('<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(12))%>s tagged by this login removed from tag list.");
		</script>
		<%
	end if

	
	If request.form("optSaleLoc")<>"" then
		cLoc = CINT(sqlInjectStr(request.form("optSaleLoc")))
	else
		cLoc = 0 'temp fix for QQ
		'if session("numLocations")>1 then
		'	if session("UserLoc") <> 0 then
		'		cLoc = CINT(session("UserLoc"))
		'	else
		'		cLoc = CINT(session("curLocation"))
		'	end if
		'else
		'	cLoc = 0
		'end if
		'if onlyOwnStudio then
		'	cLoc = 0
		'end if 
	end if
	
	if request.form("requiredtxtFDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSDate = CDATE(request.form("requiredtxtFDateStart"))
		Call SetLocale("en-us")
	else
		cSDate = DateAdd("m",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
	end if

	if request.form("requiredtxtFDateEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(request.form("requiredtxtFDateEnd"))
		Call SetLocale("en-us")
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if

	if request.form("requiredtxtDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cCSDate = CDATE(request.form("requiredtxtDateStart"))
		Call SetLocale("en-us")
	else
	    if session("studioShort")="mbsw" then
    		cCSDate = DateAdd("d",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
        else
    		cCSDate = DateAdd("m",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
    	end if
	end if

	if request.form("requiredtxtDateEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cCEDate = CDATE(request.form("requiredtxtDateEnd"))
		Call SetLocale("en-us")
	else
		cCEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if

	displaycSDate = cSDate
	displaycEDate = cEDate

	if request.Form("frmGenReport")="true" or request.Form("frmExpReport")="true" then
		cSDate = DateAdd("d", -1, cSDate)
		cEDate = DateAdd("d", 1, cEDate)
	end if

	
	showDetails = request.form("optTrainer")

	if NOT request.form("frmExpReport")="true" then
		%>
			
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			<%= css(array("calendar")) %>
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %>
			<%= js(array("adm/adm_rpt_conlogfollowup")) %>
			<script type="text/javascript">
			function exportReport() {
				document.frmParameter.frmExpReport.value = "true";
				document.frmParameter.frmGenReport.value = "true";
				<% iframeSubmit "frmParameter", "adm_rpt_conlogfollowup.asp" %>
			}
			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
			<!-- #include file="../inc_ajax.asp" -->
			<!-- #include file="../inc_val_date.asp" -->
		<%
		end if
		
		%>
		
		
		<% if NOT request.form("frmExpReport")="true" then %>
<% pageStart %>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div class="headText breadcrumbs-old">
			<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
			<%if category<>"" then%>
			<span class="breadcrumb-item">&raquo;</span>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<% end if %>
			<span class="breadcrumb-item">&raquo;</span>
			<%= DisplayPhrase(reportPageTitlesDictionary,"Contactlogs") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
			</div>
<%end if %>
			<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
				<tr> 
				<td valign="top" height="100%" width="100%"> 
				<table class="center" cellspacing="0" width="90%" height="100%">
				<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<tr>
					<td class="headText" align="left" valign="top">
					<table width="100%" cellspacing="0">
						<tr>
						<td class="headText" valign="bottom"><b><%= pp_PageTitle("Contact Logs") %></b></td>
						<td valign="bottom" class="right" height="26"> </td>
						</tr>
					</table>
					</td>
					</tr>
				<%end if %>
					<tr> 
					<td height="30"  valign="bottom" class="headText">
					<table class="mainText border4 center" cellspacing="0">
						<form name="frmParameter" action="adm_rpt_conlogfollowup.asp" method="POST">
						<input type="hidden" name="frmGenReport" value="">
						<input type="hidden" name="frmExpReport" value="">
						<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
							<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
							<input type="hidden" name="category" value="<%=category%>">
						<% end if %>
						<tr> 
						<td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b><span style="color:<%=session("pageColor4")%>;">&nbsp;</span>
								Staff Member:
		                     	<select name="optTrainer" onChange="refreshReport();">
    	            	        <option value="-2" <%if request.form("optTrainer")="-2" then response.write "selected" end if%>>All Staff Members - Summary</option>
    	            	        <option value="-3" <%if request.form("optTrainer")="-3" then response.write "selected" end if%>>Auto Emails Sent</option>
								<%
								set rsEntry2 = Server.CreateObject("ADODB.Recordset")
								strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName "
								strSQL = strSQL & "FROM TRAINERS "
								
								'CB 5/14/09 - Removed for Performance
								'strSQL = strSQL & "INNER JOIN tblContactLogs ON (Trainers.TrainerID=tblContactLogs.TrainerID OR Trainers.TrainerID=tblContactLogs.AssignedTo)"
								
								strSQL = strSQL & "WHERE TRAINERS.[Active]=1 AND TRAINERS.[Delete]=0 AND TRAINERS.TrainerID>0 AND TRAINERS.isSystem=0 "
								'if request.Form("optFollowup")="on" then
								'	strSQL = strSQL & "AND ([tblContactLogs].FollowupDate>=" & DateSep & cSDate & DateSep & ") "
								'	strSQL = strSQL & "AND ([tblContactLogs].FollowupDate<=" & DateSep & cEDate & DateSep & ") "
								'	strSQL = strSQL & "AND RequiresFollowup=0 AND tblContactLogs.Deleted=0 "
								'else
								'	strSQL = strSQL & "AND ([tblContactLogs].ContactDate>=" & DateSep & cSDate & DateSep & ") "
								'	strSQL = strSQL & "AND ([tblContactLogs].ContactDate<=" & DateSep & cEDate & DateSep & ") "
								'	strSQL = strSQL & "AND RequiresFollowup=0 AND tblContactLogs.Deleted=0 "
								'end if
								strSQL = strSQL & "ORDER BY TRAINERS.TrLastName"
								rsEntry2.CursorLocation = 3
								rsEntry2.open strSQL, cnWS
								Set rsEntry2.ActiveConnection = Nothing
	
								Do While NOT rsEntry2.EOF
							%>
								<option value="<%=rsEntry2("TrainerID")%>" <%if request.form("optTrainer")=CSTR(rsEntry2("TrainerID")) then response.write "selected" end if%>><% response.write rsEntry2("TrLastName") & ", " & rsEntry2("TrFirstName") %></option>
							<%
									rsEntry2.MoveNext
								Loop	
								rsEntry2.close
							%>
    	                  </select>&nbsp;&nbsp;&nbsp;
    	                  
    	                  Creation Date Between: 
							&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(77))%>: 
								<input onBlur="validateDate(this, '<%=FmtDateShort(cCSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cCSDate)%>" class="date">
						<script type="text/javascript">
						var cal3 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStart'});
						cal3.a_tpl.yearscroll = true;
						</script>
								&nbsp;<%=xssStr(allHotWords(79))%>: 
								<input onBlur="validateDate(this, '<%=FmtDateShort(cCEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cCEDate)%>" class="date">
						<script type="text/javascript">
						var cal4 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
						cal4.a_tpl.yearscroll = true;
						</script>
								&nbsp;
							<span title="Location of contact log is determined by <br />1) the staff member that created it <br /> 2) what location that staff member is assigned to. <br /> Location does not apply to system generated contact logs.">Location: </span><select name="optSaleLoc">
									<option value="0" <% if cLoc = 0 then response.write "selected" end if %>>All</option>
		<%
					strSQL = "SELECT Location.LocationID, Location.LocationName FROM Location WHERE (Location.Active = 1) ORDER BY Location.LocationName"
					rsEntry2.CursorLocation = 3
					rsEntry2.open strSQL, cnWS
					Set rsEntry2.ActiveConnection = Nothing

					do While NOT rsEntry2.EOF
						if rsEntry2("LocationID") <> 98 then
		%>
							<option value="<%=rsEntry2("LocationID")%>" <%if cLoc=rsEntry2("LocationID") then response.write "selected" end if%>><%=rsEntry2("LocationName")%></option>
		<%
						end if
						rsEntry2.MoveNext
					loop
					rsEntry2.close
		%>
									</select>							


						  &nbsp;&nbsp;Group By:&nbsp;
							<select name="optGroupBy">
								<option value="0" <% if request.form("optGroupBy")="0" then response.write "selected" end if %>>Created By</option>
								<option value="1" <% if request.form("optGroupBy")="1" then response.write "selected" end if %>>Assigned To</option>
							</select>
						</td>
						</tr>
						<tr style="background-color:#F2F2F2;">
							<td class="center">
								<strong>View:&nbsp;</strong>
								<select name="optView" onchange="refreshReport();">
									<option value="0" <% if request.form("optView")="0" then response.write "selected" end if %>>Detail</option>
									<option value="1" <% if request.form("optView")="1" then response.write "selected" end if %>>Summary</option>
								</select>
							</td>
						</tr>

<tr style="background-color:#F2F2F2;">
<td class="center">
<%showDateArrows "frmParameter" %>
</td>
</tr>

						<tr style="background-color:#F2F2F2;">
						    <td class="center"><strong>
							&nbsp;&nbsp;<input type="checkbox" name="optDetailedFilters" <%if request.Form("optDetailedFilters")="on" then response.write(" checked") end if%> onClick="refreshReport();"> Use Detailed Filters
<% if request.form("optDetailedFilters")="on" then %>
							 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Log Types:
                                <select id="optContactLogType" name="optContactLogType" <% if ss_ContactLogSubtypes then response.write "onChange=""refreshReport();"""%>>
							      <option value="0" <%if request.form("optContactLogType")="" OR request.form("optContactLogType")="0" then response.write "selected" end if%>>All Log Types</option>
                                  <%
						strSQL = "SELECT ContactTypeID, ContactType FROM tblContactTypes WHERE (Active = 1) ORDER BY ContactType"
						rsEntry2.CursorLocation = 3
						rsEntry2.open strSQL, cnWS
						Set rsEntry2.ActiveConnection = Nothing
						do while NOT rsEntry2.EOF
                                  %>								
							      <option value="<%=rsEntry2("ContactTypeID")%>" <%if request.form("optContactLogType")=CSTR(rsEntry2("ContactTypeID")) then response.write "selected" end if%>><%=rsEntry2("ContactType")%></option>
                                  <%
							rsEntry2.MoveNext
						loop
						rsEntry2.close
%>
		                    </select>
						<% if ss_ContactLogSubtypes AND request.Form("optContactLogType")<>"" AND request.Form("optContactLogType")<>"0" then %>
                            <select name="optContactLogSubtype">
                                <option value="0" <%if request.form("optContactLogSubtype")="" OR request.form("optContactLogSubtype")="0" then response.write "selected" end if%>>All Subtypes</option>
                                <%
								if request.form("optContactLogType")<>"" AND request.form("optContactLogType")<>"0" then
									strSQL = "SELECT ContactLogSubtypeID, SubtypeName FROM tblContactLogSubtypes WHERE (Active = 1) AND (Deleted = 0) AND ContactLogTypeID = " & request.form("optContactLogType") & " ORDER BY SubtypeName"
									rsEntry2.CursorLocation = 3
									rsEntry2.open strSQL, cnWS
									Set rsEntry2.ActiveConnection = Nothing
									do while NOT rsEntry2.EOF
										%>								
										<option value="<%=rsEntry2("ContactLogSubtypeID")%>" <%if request.form("optContactLogSubtype")=CSTR(rsEntry2("ContactLogSubtypeID")) then response.write "selected" end if%>><%=rsEntry2("SubtypeName")%></option>
										<%
										rsEntry2.MoveNext
									loop
									rsEntry2.close
								end if
                                %>
                            </select>
                        <% end if %>
                        <!-- JM-54_2440-->
    	                  Alerts:
		                     	
    	            	        
								<%
								set rsEntry2 = Server.CreateObject("ADODB.Recordset")
								strSQL = "SELECT DISTINCT tblContactLogs.AlertID, tblAlert.AlertName FROM tblContactLogs INNER JOIN tblAlert ON tblAlert.AlertID=tblContactLogs.AlertID WHERE Deleted = 0 "
								if request.form("optTrainer")<>"-3" AND request.form("optTrainer")<>"-2" AND request.form("optTrainer")<> "" then
								    ' SQL injection test
								    if isNum(request.form("optTrainer")) then
								        strSQL = strSQL & " AND tblContactLogs.TrainerID=" & request.form("optTrainer")
								    else 
								        strSQL = strSQL & " AND tblContactLogs.TrainerID=-1 "
								    end if
								end if
								if cCSDate <> " " then
								    strSQL = strSQL & " AND tblContactLogs.ContactDate>=" & DateSep & sqlInjectStr(cCSDate) & DateSep 
								end if
								if cCEDate <> " "  then
								    strSQL = strSQL & " AND tblContactLogs.ContactDate<=" & DateSep & sqlInjectStr(cCEDate) & DateSep 
								end if
							'response.write debugSQL(strSQL, "SQL")
								rsEntry2.CursorLocation = 3
								rsEntry2.open strSQL, cnWS
								Set rsEntry2.ActiveConnection = Nothing%>
	<select name="optAlerts" >
    	            	        <option value="" <%if request.form("optAlerts")="" then response.write "selected" end if%>>Select</option>
    	            	        <option value="0" <%if request.form("optAlerts")="0" then response.write "selected" end if%>>All Alerts</option>
								<%Do While NOT rsEntry2.EOF		%>
								<option value="<%=xssStr(rsEntry2("AlertID"))%>" <%if request.form("optAlerts")=CSTR(rsEntry2("AlertID")) then response.write "selected" end if%>><% response.write xssStr(rsEntry2("AlertName")) %></option>
							<%
									rsEntry2.MoveNext
								Loop	
								rsEntry2.close
							%>
    	                  </select>&nbsp;&nbsp;&nbsp;
                        &nbsp;&nbsp;
						<select name="optContactMethod">
							<option value="all">All Contact Methods</option>
							<option value="E-mail"<% if request.form("optContactMethod")="E-mail" then response.write " selected" end if %>>E-mail</option>
							<option value="In Person"<% if request.form("optContactMethod")="In Person" then response.write " selected" end if %>>In Person</option>
							<option value="Mail"<% if request.form("optContactMethod")="Mail" then response.write " selected" end if %>>Mail</option>
							<option value="Note"<% if request.form("optContactMethod")="Note" then response.write " selected" end if %>>Note</option>
							<option value="Phone"<% if request.form("optContactMethod")="Phone" then response.write " selected" end if %>>Phone</option>
							<option value="SMS"<% if request.form("optContactMethod")="SMS" then response.write " selected" end if %>>SMS</option>
						</select>
							
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="checkbox" name="optNagging" <%if request.Form("optNagging")="on" then response.write(" checked") end if%>> Include 'Top Ten <%=session("ClientHW")%>s' chart 
						</td>
						</tr>
						<tr style="background-color:#F2F2F2;">
						    <td class="center"><strong>
							&nbsp;Assigned to: 
							<select name="optAssignedFilter">
								<option value="0" <%if request.form("optContactLogType")="" OR request.form("optContactLogType")="0" then response.write "selected" end if%>>Any</option>
<%
						strSQL = "SELECT tblContactLogs.AssignedTo, TRAINERS.TrFirstName, TRAINERS.TrLastName FROM tblContactLogs INNER JOIN TRAINERS ON tblContactLogs.AssignedTo = TRAINERS.TrainerID WHERE (TRAINERS.Active = 1) AND (TRAINERS.[Delete] = 0) GROUP BY tblContactLogs.AssignedTo, TRAINERS.TrFirstName, TRAINERS.TrLastName HAVING (tblContactLogs.AssignedTo IS NOT NULL) ORDER BY TRAINERS.TrLastName"
						rsEntry2.CursorLocation = 3
						rsEntry2.open strSQL, cnWS
						Set rsEntry2.ActiveConnection = Nothing
						do while NOT rsEntry2.EOF
%>								
								<option value="<%=rsEntry2("AssignedTo")%>" <%if request.form("optAssignedFilter")=CSTR(rsEntry2("AssignedTo")) then response.write "selected" end if%>><%=rsEntry2("TrLastName")%>, <%=rsEntry2("TrFirstName")%></option>
<%
							rsEntry2.MoveNext
						loop
						rsEntry2.close
%>
							</select>
							<script type="text/javascript">
								document.frmParameter.optAssignedFilter.options[0].text = "Any Employee";
							</script>
							&nbsp;<input type="checkbox" name="optFollowups" onClick="refreshReport();" <%if request.Form("optFollowups")="on" then response.write(" checked") end if%>>Filter followups by date
							<% if request.Form("optFollowups")="on" then %>
							&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(77))%>: 
								<input onBlur="validateDate(this, '<%=FmtDateShort(displaycSDate)%>', true);" type="text"  name="requiredtxtFDateStart" value="<%=FmtDateShort(displaycSDate)%>" class="date">
						<script type="text/javascript">
						var cal1 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtFDateStart'});
						cal1.a_tpl.yearscroll = true;
						</script>
								&nbsp;<%=xssStr(allHotWords(79))%>: 
								<input onBlur="validateDate(this, '<%=FmtDateShort(displaycEDate)%>', true);" type="text"  name="requiredtxtFDateEnd" value="<%=FmtDateShort(displaycEDate)%>" class="date">
						<script type="text/javascript">
						var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtFDateEnd'});
						cal2.a_tpl.yearscroll = true;
						</script>
								&nbsp;
								&nbsp;<input type="checkbox" name="optOverDues" <%if request.Form("optOverDues")="on" then response.write(" checked") end if%>> Include Overdue Followups

							<% end if %>
							<% if ss_UseContactLogForecasting then %>
							&nbsp;&nbsp;<input type="checkbox" name="optShowForecasting" <%if request.Form("optShowForecasting")="on" then response.write " checked" end if%>>
							Show Forecasting&nbsp;&nbsp;
							<% end if %>
<% end if %>				

						
						</td></tr>

						<tr style="background-color:#F2F2F2;"><td class="center-ch" valign="middle">
							&nbsp;&nbsp;<% taggingFilter %>&nbsp;&nbsp;
							&nbsp;
							<input type="button" name="Button" value="Generate" onClick="genReport();"></b>
							<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
							else%>
								<% exportToExcelButton %>
							<%end if%>
							<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
							else%>
								<% taggingButtons("frmParameter") %>
							<%end if%>
							<% savingButtons "frmParameter", "Contact Logs" %>
						  </td>
						</tr>
						</form>
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" id="contactLogsGenTag" class="mainTextBig center-ch"> 
					<% 

						if request.form("frmGenReport")="true" or request.Form("frmExpReport")="true" then
							if request.Form("optNagging")="on" then 
								
								strSQL = "SELECT COUNT(*) AS LogCount, t.ClientID, Min(t.ContactDate) as FirstContact, (SELECT d.LastName FROM Clients d WHERE d.ClientID = t.ClientID) AS LastName "
								strSQL = strSQL & "FROM tblContactLogs t INNER JOIN CLIENTS c ON t.ClientID = c.ClientID "
								if isNum(request.form("optContactLogType")) AND request.form("optContactLogType")<>"0" then
									strSQL = strSQL & "INNER JOIN tblContactLogsContactTypes ON t.ContactLogID = tblContactLogsContactTypes.ContactLogID AND tblContactLogsContactTypes.ContactTypeID = " & request.form("optContactLogType")
								end if
								if isNum(request.form("optContactLogSubtype")) AND request.form("optContactLogSubtype")<>"0" then 
									strSQL = strSQL & "INNER JOIN tblContactLogsContactSubtypes ON t.ContactLogID = tblContactLogsContactSubtypes.ContactLogID AND tblContactLogsContactSubtypes.ContactLogSubtypeID = " & request.form("optContactLogSubtype")
								end if
								strSQL = strSQL & "WHERE (t.ContactDate >= " & DateSep & cSDate & DateSep & ") "
								strSQL = strSQL & "AND (t.ContactDate <= " & DateSep & cEDate & DateSep & ") "
								strSQL = strSQL & "AND t.Deleted=0 "
								strSQL = strSQL & "GROUP BY t.ClientID "
								strSQL = strSQL & "ORDER BY LogCount DESC"
							
							    response.write debugSQL(strSQL, "SQL")
								
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								Set rsEntry.ActiveConnection = Nothing
								%>
							<br /><br />
							<table class="mainText center"  cellspacing="0" width="39%">
							<tr>
							  <td colspan="4" class="whiteHeader" style="background-color:<%=session("pageColor4")%>;" align="left"><center><strong>Top Ten <%=session("ClientHW")%>s In Date Range</strong></center></td>
							 </tr>
<% if NOT request.form("frmExpReport")="true"  then %>
							 <tr height="2">
								<td colspan="6" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
							</tr>
<% end if %>
							<tr>
									<td width="8%"><strong>Rank</strong></td>
									<td width="15%"><strong><%=session("ClientHW")%></strong></td>
									<td width="8%"><strong>Log Count</strong></td>
									<td width="8%"><strong>Logs Per Day</strong></td>
							</tr>
							<%
								dim i
								dim logsperday
								i = 0
								 if NOT rsEntry.EOF then			'EOF
									do while NOT rsEntry.EOF AND i < 10
									i = i + 1
							%>
							<tr>
								<td><%=i%></td>
								<%strLink = "adm_clt_conlog.asp?studioshort=" & strStudioShort & "&amp;clientid=" & rsEntry("ClientID")%>
								<td><a href="<%=strLink%>"><%=rsEntry("LastName")%></a></td>
								<td><%=rsEntry("LogCount")%></td>
								<% dim daydiff
									daydiff = DateDiff("d",cSDate,cEDate)
									if daydiff <= 0 then 
									   daydiff = 1 
									end if
								%>
								<td><%=formatnumber((rsEntry("LogCount") / daydiff),3)%></td>
								<!--<td><% 'logsperday=(rsEntry("LogCount") / DateDiff("d",rsEntry("FirstContact"),Now()))%><% '=formatnumber(logsperday, 2)%></td>-->
								<!-- <% '=rsEntry("LogCount")%> | <% '=rsEntry("CallsPerDay")%><%'	=rsEntry("LogCount")%><% 'logsperday=(DateDiff("d",rsEntry("FirstContact"),Now()) / rsEntry("LogCount")) %><% '= logsperday%> </td> -->
							</tr>
							<%		
										rsEntry.MoveNext
									loop
								 end if
								 rsEntry.close
							%>
<% if NOT request.form("frmExpReport")="true"  then %>
							<tr height="2">
								<td colspan="6" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
							</tr>
<% end if %>
							</table>
						<%end if%>
						<%end if%>
						
					<table class="mainText" width="100%" cellspacing="0">
						<tr>
						<td  colspan="2" valign="top" class="mainTextBig center-ch">&nbsp;</td>
						</tr>
						<tr > 
						<td  colspan="2" valign="top" class="mainTextBig center-ch">
		<% 
		end if			'end of frmExpreport value check before /head line	  
							if request.form("frmGenReport")="true" then 
								if request.form("frmExpReport")="true" then
									Dim stFilename
									if showDetails="-2" then
										stFilename="attachment; filename=Contact Log Analysis - All Staff " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									else
										stFilename="attachment; filename=Contact Log Analysis - " & FmtTrnName(CLNG(showDetails)) & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									end if
									Response.ContentType = "application/vnd.ms-excel" 
									Response.AddHeader "Content-Disposition", stFilename 
								end if

								showHeader = "false"
								curTrainer=0
				
								dim strOptText(100)
								dim maxStrOptText
								maxStrOptText = 0
								strSQL = "SELECT ContactTypeID, ContactType FROM tblContactTypes"
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								set rsEntry.ActiveConnection = Nothing
								if NOT rsEntry.EOF then
									Do While NOT rsEntry.EOF
										strOptText(rsEntry("ContactTypeID")) = rsEntry("ContactType")
										maxStrOptText = rsEntry("ContactTypeID")
										rsEntry.MoveNext
									loop
								end if
								rsEntry.Close

                                ''*****************************************************
                                '' MAIN REPORT SQL
                                ''*****************************************************							
								if request.form("frmTagClients")<>"true" AND request.Form("optView")<>"1" then
									strSQL = "SELECT TOP 500 "
								else
									strSQL = "SELECT "
								end if 
								strSQL = strSQL & "tblContactLogs.FollowupDate, AssignedTr.TrFirstName as AssignedFirstName, AssignedTr.TrLastName as AssignedLastName, tblContactLogs.AssignedTo, tblContactLogs.ContactMethod, tblContactLogs.RequiresFollowup, tblContactLogs.ContactDate, tblContactLogs.ContactName, tblContactLogs.ContactLog, tblContactLogs.ContactLogType, t.TrFirstName, t.TrLastName, tblContactLogs.TrainerID, cl.LastName, cl.FirstName, tblContactLogs.ContactLogID, tblContactLogs.ClientID, tblContactLogs.ForecastAmount, Categories.CategoryID, Categories.CategoryName, tblContactTypes.ContactType "
								if request.Form("optView") = "1" then
									strSQL = "SELECT tblContactLogs.ContactMethod, COUNT(DISTINCT tblContactLogs.ContactLogID) AS Total "
									if request.Form("frmTagClients")="true" then
										strSQL = strSQL & ", tblContactLogs.ClientID "
									end if
								end if
								'strSQL = strSQL & ", c.NumTypes, c.ContactType "

								strSQL = strSQL & "FROM tblContactLogs "

								'strSQL = strSQL & "LEFT OUTER JOIN (SELECT tblContactLogs.ContactLogID, COUNT(tblContactLogsContactTypes.ContactLogID) AS NumTypes, MIN(tblContactTypes.ContactType) AS ContactType FROM tblContactLogs LEFT OUTER JOIN tblContactLogsContactTypes ON tblContactLogs.ContactLogID = tblContactLogsContactTypes.ContactLogID LEFT OUTER JOIN tblContactTypes ON tblContactLogsContactTypes.ContactTypeID = tblContactTypes.ContactTypeID GROUP BY tblContactLogs.ContactLogID) AS c ON c.ContactLogID = tblContactLogs2.ContactLogID "

								strSQL = strSQL & "INNER JOIN Clients cl ON tblContactLogs.ClientID = cl.ClientID "
								strSQL = strSQL & "INNER JOIN Trainers t ON tblContactLogs.TrainerID=t.TrainerID "
								strSQL = strSQL & "LEFT OUTER JOIN Trainers AssignedTr ON AssignedTr.TrainerID = tblContactLogs.AssignedTo "
								strSQL = strSQL & "LEFT OUTER JOIN Categories ON tblContactLogs.ForecastCategoryID = Categories.CategoryID "
								if request.form("optFilterTagged")="on" then
									strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = cl.ClientID "
									if session("mVarUserID")<>"" then
										strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
									end if
									strSQL = strSQL & " ) "
								end if
								if isNum(request.form("optContactLogSubtype")) AND request.form("optContactLogSubtype")<>"0" then 
									strSQL = strSQL & "INNER JOIN tblContactLogsContactSubtypes ON tblContactLogs.ContactLogID = tblContactLogsContactSubtypes.ContactLogID AND tblContactLogsContactSubtypes.ContactLogSubtypeID = " & request.form("optContactLogSubtype")
								end if

								strSQL = strSQL & "LEFT OUTER JOIN tblContactLogsContactTypes ON tblContactLogs.ContactLogID = tblContactLogsContactTypes.ContactLogID "
								
								if isNum(request.form("optContactLogType")) AND request.form("optContactLogType")<>"0" then
								  strSQL = strSQL & "INNER JOIN tblContactTypes ON tblContactTypes.ContactTypeID = tblContactLogsContactTypes.ContactTypeID " &_
									                  "AND tblContactLogsContactTypes.ContactTypeID = " & request.form("optContactLogType") & " "
							  else
							    strSQL = strSQL & "LEFT OUTER JOIN tblContactTypes ON tblContactTypes.ContactTypeID = tblContactLogsContactTypes.ContactTypeID "
								end if

								strSQL = strSQL & "WHERE 1=1 "
								'JM-54_2440
								if request.form("optAlerts")= "0" AND request.form("optAlerts")<>"" then
								    strSQL = strSQL & " AND tblContactLogs.AlertID is NOT NULL AND tblContactLogs.SystemGenerated=1 "
								elseif request.form("optAlerts")<>"0" AND request.form("optAlerts")<>"" AND isNum(request.Form("optAlerts")) then
								    strSQL = strSQL & " AND tblContactLogs.AlertID= " & request.form("optAlerts")&" AND tblContactLogs.SystemGenerated=1 "
								elseif request.form("optAlerts")= "" then
								    strSQL = strSQL & " AND tblContactLogs.AlertID IS NULL "
								end if
								'if request.form("optContactLogType")<>"" AND request.form("optContactLogType")<>"0" then
								'	strSQL = strSQL & "INNER JOIN tblContactLogsContactTypes ON tblContactLogs2.ContactLogID = tblContactLogsContactTypes.ContactLogID AND tblContactLogsContactTypes.ContactTypeID = " & request.form("optContactLogType") & " "
								'end if
								if request.Form("optFollowups")="on" and request.form("optDetailedFilters")="on" then
									if NOT request.form("optOverDues")="on" then ' include overdue
										strSQL = strSQL & " AND (tblContactLogs.FollowupDate >=" & DateSep & cSDate & DateSep & ") "
									end if
									strSQL = strSQL & " AND (tblContactLogs.FollowupDate <=" & DateSep & cEDate & DateSep & ") "
									strSQL = strSQL & " AND tblContactLogs.RequiresFollowup = 0 AND tblContactLogs.Deleted=0 "
								end if
								
								strSQL = strSQL & " AND (tblContactLogs.ContactDate >= " & DateSep & cCSDate & DateSep & ") "
								strSQL = strSQL & " AND (tblContactLogs.ContactDate < " & DateSep & (cdate(cCEDate) + 1) & DateSep & ") "
								strSQL = strSQL & " AND tblContactLogs.Deleted=0 "
								'RI-55_3250
								if request.form("optSaleLoc")<>"0" AND request.form("optSaleLoc")<>"" then
									strSQL = strSQL & "AND tblContactLogs.LocationID = " & request.Form("optSaleLoc") & " "
								end if

								if showdetails<>"-2" and request.form("optDetailedFilters")="on" then
									strSQL = strSQL & "AND tblContactLogs.TrainerID=" & clng(showDetails) & " " 
								end if

								if isNum(request.form("optAssignedFilter")) and request.form("optAssignedFilter")<>"0" and request.form("optDetailedFilters")="on" then
									strSQL = strSQL & " AND AssignedTo = " & request.form("optAssignedFilter")
								end if
								
														
								if request.form("optContactMethod")<>"all" and request.form("optDetailedFilters")="on" then
									strSQL = strSQL & " AND tblContactLogs.ContactMethod	= N'" & sqlInjectStr(request.form("optContactMethod")) & "' "
								end if

								if NOT request.form("optDetailedFilters")="on" then
									if not showdetails="-2" then
										strSQL = strSQL & " AND (tblContactLogs.TrainerID=" & clng(showDetails) & " OR AssignedTo = " & clng(showDetails) & " )"
									end if
								end if
								
								if request.form("optGroupBy")="1" then
									strSQL = strSQL & " AND (NOT AssignedTr.TrainerID IS NULL) "
								end if
								if request.form("optShowForecasting")="on" then
									strSQL = strSQL & " AND (tblContactLogs.ForecastAmount IS NOT NULL) "
								end if
								
								if request.Form("optView") = "1" AND request.form("frmTagClients")="true" then
									strSQL = strSQL & " GROUP BY ContactMethod, tblContactLogs.ClientID "
								elseif request.Form("optView") = "1" AND request.form("frmTagClients")<>"true" then
									strSQL = strSQL & " GROUP BY ContactMethod ORDER BY ContactMethod"
								elseif request.form("frmTagClients")<>"true" then
									strSQL = strSQL & " ORDER BY "
									if request.form("optGroupBy")="0" then
										strSQL = strSQL & " t.TrLastName, "
									elseif request.form("optGroupBy")="1" then
										strSQL = strSQL & " tblContactLogs.AssignedTo, "
									end if
									
									if request.Form("optFollowups")="on" and request.form("optDetailedFilters")="on" then
										strSQL = strSQL & " tblContactLogs.FollowupDate DESC "
									else
										strSQL = strSQL & " tblContactLogs.ContactDate DESC "
									end if
								end if

            					response.write debugSQL(strSQL, "SQL")

								if request.form("frmTagClients")="true" then
									if request.form("frmTagClientsNew")="true" then
										clearAndTagQuery(strSQL)
									else    
										tagQuery(strSQL)
									end if
									strSQL = "SELECT StudioID FROM Studios WHERE 1=0 "
								end if
								
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								Set rsEntry.ActiveConnection = Nothing
								dim newTrn, trnCount, forecastTotal, grandTotal
								resultCount = rsEntry.RecordCount
								trnCount = 0
								forecastTotal = 0
								grandTotal = 0
								newTrn = false
								%>
									<table class="mainText center" cellspacing="0">
								<%   if NOT rsEntry.EOF then			'EOF
										do while NOT rsEntry.EOF
										
										    response.flush
											if request.Form("optView")<>"1" then
												newTrn = false
												if request.form("optGroupBy")="0" then  	'if this is a new trainer then write the header cells
													if curTrainer<>clng(rsEntry("TrainerID")) then
														newTrn = true
													end if
												end if
												if request.form("optGroupBy")="1" then
													if curTrainer<>clng(rsEntry("AssignedTo")) then
														newTrn = true
													end if
												end if
												if newTrn then
													if curTrainer<> "" AND curTrainer <> 0 then	'if this isn't the first record total cells
										%>
	<% if NOT request.form("frmExpReport")="true"  then %>
													<tr height="2">
														<td colspan="9" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
	<% end if %>
													<tr>
														<td colspan="9" class="right"><strong>Total Logs: <%=trnCount%></strong>&nbsp;</td>
													</tr>

										<%				grandTotal = grandTotal + trnCount
																		trnCount = 0
													end if	
										%>
													<tr>
														<td colspan="9">&nbsp;</td>
													</tr>
													<tr>
														<% if request.form("optGroupBy")="0" then %>
														<td colspan="9" class="whiteHeader" style="background-color:<%=session("pageColor4")%>;" align="left">&nbsp;Created By: <%=rsEntry("TrFirstName")%>&nbsp;<%=rsEntry("TrLastName")%></td>
												  		<% else %>
														<td colspan="9" class="whiteHeader" style="background-color:<%=session("pageColor4")%>;" align="left">&nbsp;Assigned To: <%=rsEntry("AssignedFirstName")%>&nbsp;<%=rsEntry("AssignedLastName")%></td>
														<% end if %>
													</tr>
													<tr align="left" valign="bottom" height="20">
														<td><strong>Followup Due</strong>&nbsp;&nbsp;</td>
														<td><strong>Log Date</strong>&nbsp;&nbsp;</td>
														<td><strong>Log Type</strong>&nbsp;&nbsp;</td>
														<% if ss_ContactLogSubtypes then %>
                                                    		<td><strong>Sub Type</strong>&nbsp;&nbsp;</td>
														<%end if%>
														<td><strong>Contact</strong>&nbsp;&nbsp;</td>
														<td><strong>Contact Method</strong>&nbsp;&nbsp;</td>
														<td><strong>Contact Log</strong>&nbsp;&nbsp;</td>
													<% if request.form("optShowForecasting")="on" then %>
														<td><strong>Forecast Amount</strong>&nbsp;&nbsp;</td>
														<td><strong>Forecast Category</strong>&nbsp;&nbsp;</td>
													<% end if %>
														<td><strong>Link</strong></td>
													</tr>
	<% if NOT request.form("frmExpReport")="true"  then %>
													<tr height="2">
															<td colspan="9" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
													</tr>
	<% end if %>
												<%
												end if
												if request.form("optGroupBy")="0" then
													curTrainer = CLNG(rsEntry("TrainerID"))
												else
													curTrainer = CLNG(rsEntry("AssignedTo"))
												end if
												
												if rowColor = "#F2F2F2" then
													rowColor = "#FAFAFA"
												else
													rowColor = "#F2F2F2"
												end if
												%>
													<tr align="left" valign="top" style="background-color:<%=rowColor%>;">
														<%if not isnull(rsEntry("FollowupDate")) then%>
															<td><%if CDate(rsEntry("FollowupDate")) < Now AND rsEntry("RequiresFollowup")=0 then%><span style="color:#FF0000;"><%=trim(FmtDateTime(rsEntry("FollowUpDate")))%></span><%else%><% if rsEntry("RequiresFollowup")=0 then response.write(trim(FmtDateTime(rsEntry("FollowUpDate")))) end if %><%end if%>&nbsp;</td>
														<%else%>
															<td></td>
														<%end if%>
														<td><%=trim(FmtDateTime(rsEntry("ContactDate")))%>&nbsp;</td>
														<td align="left">
	<%
												'if rsEntry("NumTypes")>1 then
												'	'BJD: 47_2365 - updated type logic - queries only if there is more than 1 type on the contact log
												'	first = true
													
												'	strSQL = "SELECT tblContactTypes.ContactType FROM tblContactLogsContactTypes INNER JOIN tblContactTypes ON tblContactLogsContactTypes.ContactTypeID = tblContactTypes.ContactTypeID WHERE (tblContactLogsContactTypes.ContactLogID = " & rsEntry("ContactLogID") & ") "
												'	rsEntry2.CursorLocation = 3
												'	rsEntry2.open strSQL, cnWS
												'	Set rsEntry2.ActiveConnection = Nothing
													
												'	do while NOT rsEntry2.EOF
												'		if not first then
												'			response.write ", "
												'		end if
												'		first = false
												'		response.write TRIM(rsEntry2("ContactType"))
													
												'		rsEntry2.moveNext
												'	loop
												'	rsEntry2.close
												'elseif rsEntry("NumTypes")=1 then
												'		response.write rsEntry("ContactType")
												'end if

														response.write rsEntry("ContactType")
														tmpContactLogID = CLNG(rsEntry("ContactLogID"))

																rsEntry.MoveNext
														cont = true

														do while cont
																cont = false
																if NOT rsEntry.EOF then
																		if CLNG(tmpContactLogID) = CLNG(rsEntry("ContactLogID")) then
																				response.write ", " & rsEntry("ContactType")
    																		tmpContactLogID = CLNG(rsEntry("ContactLogID"))
																				cont = true
																				rsEntry.MoveNext
																															end if                                                        
																													end if
														loop
														
														rsEntry.MovePrevious
	%>													
														</td>
														<% if ss_ContactLogSubtypes then %>
																													<td align="left">
																															<%
																															first = true
	                                                            
																															strSQL = "SELECT tblContactLogSubtypes.SubtypeName FROM tblContactLogsContactSubtypes INNER JOIN tblContactLogSubtypes ON tblContactLogsContactSubtypes.ContactLogSubtypeID = tblContactLogSubtypes.ContactLogSubtypeID WHERE (tblContactLogsContactSubtypes.ContactLogID = " & rsEntry("ContactLogID") & ") "
																															rsEntry2.CursorLocation = 3
																															rsEntry2.open strSQL, cnWS
																															Set rsEntry2.ActiveConnection = Nothing
	                                                            
																															do while NOT rsEntry2.EOF
																																	if not first then
																																			response.write ", "
																																	end if
																																	first = false
																																	response.write TRIM(rsEntry2("SubtypeName"))
	                                                                
																																	rsEntry2.moveNext
																															loop
																															rsEntry2.close
																															%>													
																													</td>
																											<% end if %>
														<td width="180"><a href="main_info.asp?id=<%=rsEntry("ClientID")%>&fl=true"><%=trim(rsEntry("LastName"))%>, <%=trim(rsEntry("FirstName"))%></a></td>
														<td><%=rsEntry("ContactMethod")%></td>
														<td width="350">
															<%if Len(trim(rsEntry("ContactLog"))) > 122 and Request.Form("frmExpReport")<>"true" then 
																response.write(Left(trim(replace(stripHTML(rsEntry("ContactLog")),VbCrLf,"<br />")),122) & "...") 
															else 
																response.Write(trim(replace(stripHTML(rsEntry("ContactLog")),VbCrLf,"<br />"))) 
															end if%>
														</td>
														<%strLink = "adm_clt_conlog.asp?studioshort=" & strStudioShort & "&amp;clientid=" & rsEntry("ClientID") & "&windowed=yes" & "&amp;conlogid=" & rsEntry("ContactLogID")%>
	<%
													if request.form("optShowForecasting")="on" then
														forecastTotal = forecastTotal + rsEntry("ForecastAmount")
	%>
														<td nowrap class="right"><%=FmtCurrency(rsEntry("ForecastAmount"))%>&nbsp;&nbsp;</td>
														<td><strong><%=rsEntry("CategoryName")%></strong>&nbsp;&nbsp;</td>
													<% end if %>
														<td><a href="<%=strLink%>" target="_blank">[Link]</a></td>
													</tr>
												<%
												trnCount=trnCount+1
											else
                        if newTrn then
													if NOT request.form("frmExpReport")="true"  then %>
														<tr height="2">
															<td colspan="9" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
														</tr>
													<% end if %>
														<tr>
															<td colspan="9" class="right"><strong>Total Logs: <%=trnCount%></strong>&nbsp;</td>
														</tr>
													<%				
													grandTotal = grandTotal + trnCount
												end if
												trnCount = rsEntry("Total")
												newTrn = true
												%>
													<tr>
														<td colspan="9">&nbsp;</td>
													</tr>
													<tr>
														<td colspan="9" class="whiteHeader" style="background-color:<%=session("pageColor4")%>;" align="left">&nbsp;<%= rsEntry("ContactMethod") %></td>
													</tr>
												<%
											end if	
											rsEntry.MoveNext
										loop
										%>
<% if NOT request.form("frmExpReport")="true"  then %>
											<tr height="2">
												<td colspan="9" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
												<tr>
												<% if request.form("optShowForecasting")="on" then %>
													<td class="right" colspan="7"><strong><%=FmtCurrency(forecastTotal)%></strong>&nbsp;&nbsp;</td>
												<% end if %>
												  <td colspan="<% if request.form("optShowForecasting")="on" then response.write "2" else response.write "9" end if %>" class="right"><strong>Total Logs: <%=trnCount%></strong>&nbsp;</td>
												</tr>

									<%				grandTotal = grandTotal + trnCount
									                trnCount = 0 %>
											<tr>
											  <td colspan="9">&nbsp;</td>
											</tr>
									<%end if	'eof
									
									if request.Form("optTrainer") = "-2" then
									    response.Write("<tr><td colspan=""9"" class=""nowrap bold right"">Grand Total Logs: " & grandTotal & "&nbsp</td></tr>")
									end if
									if resultCount >=500 then 
										%>
										<td class="center" colspan="13"><span style="color:#990000;"><br /><b>-----&nbsp;&nbsp;Only the first 500 contact logs have been listed. Please refine your search criteria, and try again.&nbsp;&nbsp;-----</b></span></td>
										<%		
									end if
%>
						  </table>
									<%
									'end if
								rsEntry.close
							end if		'end of generate report if statement
							%>
					  </table></table>
				</td>
				</tr>
				</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%'end of frmexport value check
	
end if
%>
