<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
Server.ScriptTimeout = 300    '5 min (value in seconds)
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
	dim rsEntry
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	%>
	<!-- #include file="inc_accpriv.asp" -->
	<%
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_ATTND_REV") then 
	%>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
	<%
	else
	
	%>
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_rpt_tagging.asp" -->
		<!-- #include file="inc_utilities.asp" -->
		<!-- #include file="inc_rpt_save.asp" -->
		<!-- #include file="inc_hotword.asp" -->
<%	'BQL - 45_1395 added doRefresh, set to true if changing the date should refresh the report %>
<%	dim doRefresh
	doRefresh = false %>
		<!-- #include file="inc_date_arrows.asp" -->
		<!-- #include file="../inc_ajax.asp" --> 
		<!-- #include file="../inc_val_date.asp" --> 
	<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

	Dim showDetails, cSDate, cEDate, cLoc, cLoc1, cTrainer, cTime, varGen, curRsrcName
	Dim curDate, curTime, curVisitType, curTrainer, curTrainerID, curLoc, curTG, curClassname, curClassID, curApptID
	Dim tmpRevPerVisit, tmpRev, tmpMemRev, TotRev, TotMemRev, GTMemRev, GTRev, tmpWeb, TotWeb, GrandTotWeb, tmpHCMem, TotHCMem, GrandTotHCMem
	Dim tmpHCPaid, tmpHCComp, tmpHCNoShow
	Dim TotHCPaid, TotHCComp, TotHCNoShow
	Dim GrandTotHCPaid, GrandTotHCComp, GrandTotHCNoShow
	Dim ss_UseAsst1, ss_UseAsst2, curAsst1, curAsst2, cTG, visitTGChk, visitTGList, ap_view_all_locs, pmtTGChk, pmtTGList

	ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")

	visitTGList = sqlInjectStr(request.form("optAttTG"))
	pmtTGList = sqlInjectStr(request.form("optPmtTG"))
	visitTGChk = "," & Replace(visitTGList, " ", "") & ","
	pmtTGChk = "," & Replace(pmtTGList, " ", "") & ","
				
	dim category : category = ""
	if (RQ("category"))<>"" then
		category = RQ("category")
	elseif (RF("category"))<>"" then
		category = RF("category") 
	end if
	
	If request.querystring("qGen") = "true" then
		varGen = True
	elseIf request.form("frmGenReport") = "true" then
		varGen = True
	else
		varGen = False
	end if
	
	If request.querystring("qSDate")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSDate = CDATE(request.querystring("qSDate"))
		Call SetLocale("en-us")
	elseif request.form("requiredtxtDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSDate = CDATE(sqlInjectStr(request.form("requiredtxtDateStart")))
		Call SetLocale("en-us")
	else
		cSDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if

	If request.querystring("qEDate")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(request.querystring("qEDate"))
		Call SetLocale("en-us")
	elseif request.form("requiredtxtDateEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(sqlInjectStr(request.form("requiredtxtDateEnd")))
		Call SetLocale("en-us")
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if
	
	If request.querystring("qDisMode")<>"" then
		showDetails = request.querystring("qDisMode")
	elseIf request.form("optDisMode")<>"" then
		showDetails = sqlInjectStr(request.form("optDisMode"))
	else
		showDetails = "0"
	end if

	If request.querystring("qLoc")<>"" then
		cLoc = CINT(request.querystring("qLoc"))
	elseIf request.form("optSaleLoc")<>"" then
		cLoc = CINT(sqlInjectStr(request.form("optSaleLoc")))
	else
		if session("numLocations")>1 then
			if session("UserLoc") <> 0 then
				cLoc = CINT(session("UserLoc"))
			else
				cLoc = CINT(session("curLocation"))
			end if
		else
			cLoc = 0
		end if
	end if

	'CB 5/28/2008 Updated to default the purchased at location to all locations
	If request.form("optPurLoc")<>"" then		
		cLoc1 = CINT(sqlInjectStr(request.form("optPurLoc")))
	else
		'if session("numLocations")>1 then
		'	cLoc1 = CINT(session("curLocation"))
		'else
			cLoc1 = 0
		'end if
	end if

	If request.querystring("qTrnID")<>"" and request.querystring("qTrnID")<>"0" then
		cTrainer = cstr(request.querystring("qTrnID"))
	elseIf request.form("optTrainer")<>"" and request.form("optTrainer")<>"0" then
		cTrainer = Cstr(sqlInjectStr(request.form("optTrainer")))
	else
		cTrainer="0"
	end if

	If request.querystring("qTime")<>"" and request.querystring("qTime")<>"0" then
		cTime = cDate(request.querystring("qTime"))
	elseIf request.form("optTime")<>"" and request.form("optTime")<>"0" then
		if request.Form("optTime") = "null" then
		  cTime = "null"
		else
		  cTime = cDate(sqlInjectStr(request.form("optTime")))
		end if
	else
		cTime = "0"
	end if

	'cTG = "0"
	'if request.form("optAttTG")<>"" AND request.form("optAttTG")<>"0" then
		'cTG = request.form("optAttTG")	
	'end if

	dim stType, StudTypeArr, stTypeChk

	if request.form("optPayMeth")<>"" then
		stType = sqlInjectStr(request.form("optPayMeth"))
		stTypeChk = "," & Replace(stType, " ", "") & ","
	end if

	dim hw113, hw61, hw7, hw6, hw0, hw13, hw15
	hw113 = getHotWord(113)
	hw61 = getHotWord(61)
	hw7 = getHotWord(7)
	hw6 = getHotWord(6)
	hw0 = getHotWord(0)
	hw13 = getHotWord(13)
	hw15 = getHotWord(15)

	ss_UseAsst1 = checkStudioSetting("tblGenOpts","UseAsst1")
	ss_UseAsst2 = checkStudioSetting("tblGenOpts","UseAsst2")

		if NOT request.form("frmExpReport")="true" then
		%>
			<style type="text/css">
			select.textSmall {vertical-align:middle}
			</style>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_attendance_rev", "reportFavorites", "plugins/jquery.SimpleLightBox" )) %>
<%= css(array("SimpleLightBox")) %> 

			<script type="text/javascript">
			function exportReport() {
				document.frmParameter.frmExpReport.value = "true";
				document.frmParameter.frmGenReport.value = "true";
				<% iframeSubmit "frmParameter", "adm_rpt_attendance_rev.asp" %>
			}
			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
			<!-- #include file="inc_help_content.asp" -->
		<%
		end if
		
		%>
		
		<% if NOT request.form("frmExpReport")="true" then %>
<% pageStart %>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div class="headText breadcrumbs-old" align="left">
			<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
			<%if category<>"" then%>
			<span class="breadcrumb-item">&raquo;</span>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<% end if %>
			<span class="breadcrumb-item">&raquo;</span>
			<%= DisplayPhrase(reportPageTitlesDictionary, "Attendancewithrevenue") %>&nbsp;&nbsp;&nbsp;
			<span class="mainText"><a href="javascript:switchToNoRev();">[No Revenue]</a></span>
			<% showNewHelpContentIcon("attendance-wrevenue-report") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
			</div>
<%end if %>
			<table height="100%" width="<%=strPageWidth%>" border="0" cellspacing="0" cellpadding="0">    
				<tr> 
				<td valign="top" height="100%" width="100%"> 
				<table class="center" border="0" cellspacing="0" cellpadding="0" width="90%" height="100%">
				<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<tr>
					<td class="headText left" valign="top">
					<table width="100%" border="0" cellspacing="0" cellpadding="0">
						<tr>
						<td class="headText" valign="top"><b> <%= pp_PageTitle("Attendance w Revenue") %>&nbsp;&nbsp;&nbsp;</b><span class="mainText"><a href="javascript:switchToNoRev();">[No Revenue]</a></span>
						<!--JM - 49_2447-->
						<% showNewHelpContentIcon("attendance-wrevenue-report") %>

						</td>
						<td valign="bottom" class="right" height="26"> </td>
						</tr>
					</table>
					</td>
					</tr>
				<%end if %>
					<tr> 
					<td height="30" valign="top" class="center headText">
						<form name="frmParameter" action="adm_rpt_attendance_rev.asp" method="POST">
							<input type="hidden" name="frmGenReport" value="<%=varGen%>">
							<input type="hidden" name="frmExpReport" value="">
							<input type="hidden" name="frmSubmitted" value="<%=session("StudioID")%>">
							<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
								<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
								<input type="hidden" name="category" value="<%=category%>">
							<% end if %>
							<table class="mainText center border4">
								<tr> 
									<td class="center" valign="middle" style="background-color:#F2F2F2;" nowrap>
										<b>
											<span style="color:<%=session("pageColor4")%>">&nbsp;</span>
											<%=xssStr(allHotWords(77))%>: 
											<input onchange="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
											<script type="text/javascript">
												var cal1 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStart'});
												cal1.a_tpl.yearscroll = true;
											</script>
											&nbsp;
											<%=xssStr(allHotWords(79))%>: 
											<input onchange="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
											<script type="text/javascript">
												var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
												cal2.a_tpl.yearscroll = true;
											</script>
											&nbsp;
											Used At:
											<%
											strSQL = "SELECT LocationID, LocationName FROM Location WHERE [Active]=1 AND wsShow=1 ORDER BY LocationName "
											rsEntry.CursorLocation = 3
											rsEntry.open strSQL, cnWS
											Set rsEntry.ActiveConnection = Nothing
											%>
											<select name="optSaleLoc" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
												<option value="0" <% if cLoc=0 then response.write "selected" end if %>>All</option>
												<%
												do While NOT rsEntry.EOF
													%>
													<option value="<%=rsEntry("LocationID")%>" <%if cLoc=CINT(rsEntry("LocationID")) then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
													<%
													rsEntry.MoveNext
												loop
												rsEntry.close
												%>
											</select>
						<script type="text/javascript">
							document.frmParameter.optSaleLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
						</script>
						&nbsp;

						Purchased At:
						<%strSQL = "SELECT LocationID, LocationName FROM Location WHERE [Active]=1 AND (wsShow=1 OR LocationID=98) ORDER BY LocationName "
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						Set rsEntry.ActiveConnection = Nothing%>
						<select name="optPurLoc">
						<option value="0" <% if cLoc1=0 then response.write "selected" end if %>>All</option>
						<%do While NOT rsEntry.EOF%>
								<option value="<%=rsEntry("LocationID")%>" <%if cLoc1=CINT(rsEntry("LocationID")) then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
						<%rsEntry.MoveNext
						loop
						rsEntry.close
						%>
						</select>
						<script type="text/javascript">
							document.frmParameter.optPurLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
						</script>
						&nbsp;<%=xssStr(allHotWords(6))%>:&nbsp;
						<select name="optTrainer"><option value="0" <%if cTrainer="0" then response.write "selected" end if%>>All</option>
						<%
							strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName "
							strSQL = strSQL & "FROM TRAINERS "
							strSQL = strSQL & "INNER JOIN [VISIT DATA] ON [Visit Data].TrainerID = Trainers.TrainerID "
							strSQL = strSQL & "WHERE (TRAINERS.AppointmentTrn=1 OR TRAINERS.ReservationTrn=1) AND ([VISIT DATA].ClassDate >=  " & DateSep & cSDate & DateSep & ") "
							strSQL = strSQL & "AND ([VISIT DATA].ClassDate <=  " & DateSep & cEDate & DateSep & ") "
							strSQL = strSQL & "ORDER BY " & GetTrnOrderBy()
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>	
									<option value="<%=rsEntry("TrainerID")%>" <%if cTrainer=CSTR(rsEntry("TrainerID")) then response.write "selected" end if%>><%=FmtTrnNameNew(rsEntry, true)%></option>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
						</select>		
						<script type="text/javascript">
							document.frmParameter.optTrainer.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" " + '<%=jsEscSingle(allHotWords(113))%>';
						</script>
						<br />
						&nbsp;<%=xssStr(allHotWords(76))%>:&nbsp;
						<select name="optTime"><option value="0" <%if cTime="0" then response.write "selected" end if%>><%=xssStr(allHotWords(149))%>&nbsp;<%=xssStr(allHotWords(76))%>:</option>
						<%
							strSQL = "SELECT DISTINCT ClassTime "
							strSQL = strSQL & "FROM [VISIT DATA] "
							strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >=  " & DateSep & cSDate & DateSep & ") "
							strSQL = strSQL & "AND ([VISIT DATA].ClassDate <=  " & DateSep & cEDate & DateSep & ") "
							strSQL = strSQL & "ORDER BY Classtime"
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>	
									<% if isNull(rsEntry("ClassTime")) then %>
									     <option value="null" <%if cTime="null" then response.write "selected" end if%>>TBD</option>
									<% else %>
									     <option value="<%=rsEntry("ClassTime")%>" <%if cTime=Cdate(rsEntry("Classtime")) then response.write "selected" end if%>><%=FmtTimeShort(rsEntry("Classtime"))%></option>
								  <% end if %>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
						</select>&nbsp;
						
						<%=xssStr(allHotWords(152))%>: <input type="checkbox" name="chkDaySun" <% if request.form("chkDaySun")="on" OR request.form("frmSubmitted")="" then response.write "checked" end if %>>&nbsp;
						<%=xssStr(allHotWords(153))%>: <input type="checkbox" name="chkDayMon" <% if request.form("chkDayMon")="on" OR request.form("frmSubmitted")="" then response.write "checked" end if %>>&nbsp;
						<%=xssStr(allHotWords(154))%>: <input type="checkbox" name="chkDayTue" <% if request.form("chkDayTue")="on" OR request.form("frmSubmitted")="" then response.write "checked" end if %>>&nbsp;
						<%=xssStr(allHotWords(155))%>: <input type="checkbox" name="chkDayWed" <% if request.form("chkDayWed")="on" OR request.form("frmSubmitted")="" then response.write "checked" end if %>>&nbsp;
						<%=xssStr(allHotWords(156))%>: <input type="checkbox" name="chkDayThu" <% if request.form("chkDayThu")="on" OR request.form("frmSubmitted")="" then response.write "checked" end if %>>&nbsp;
						<%=xssStr(allHotWords(157))%>: <input type="checkbox" name="chkDayFri" <% if request.form("chkDayFri")="on" OR request.form("frmSubmitted")="" then response.write "checked" end if %>>&nbsp;
						<%=xssStr(allHotWords(158))%> <input type="checkbox" name="chkDaySat" <% if request.form("chkDaySat")="on" OR request.form("frmSubmitted")="" then response.write "checked" end if %>>&nbsp;
						<%if session("CR_Class")<>0 OR session("CR_Appt")<>0 then 'CB 3/3/2009%>
						&nbsp;Cross Regional Visits:<input type="checkbox" name="optCRVisits" <%if request.form("optCRVisits")="on" then response.write "checked" end if%>>
						<%end if%>
						&nbsp;<% taggingFilter %>
						<br />
						&nbsp;Visit&nbsp;<%=xssStr(allHotWords(7))%>:
						<select multiple size="3" name="optAttTG" <%showMultiSelectTitle() %>>
							<option value="0" <%if inStr(visitTGChk, ",0,")>0 OR request.form("optAttTG")="" then response.write "selected" end if%>>All</option>
<%
							strSQL = "SELECT TypeGroupID, TypeGroup FROM tblTypeGroup WHERE (Active = 1) AND ((wsReservation = 1) OR (wsAppointment = 1) OR (wsResource=1) OR (wsEnrollment=1) OR (wsArrival=1) ) ORDER BY TypeGroup"
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>	
									<option value="<%=rsEntry("TypeGroupID")%>" <%if inStr(visitTGChk, "," & TRIM(rsEntry("TypeGroupID")) & ",") > 0 then response.write "selected" end if%>><%=rsEntry("TypeGroup")%></option>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
						</select>
						<script type="text/javascript">
							document.frmParameter.optAttTG.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(503))%>";
						</script>

<%
					dim useRelatedTGs

					strSQL = "SELECT TG1 FROM tblTGRelate "
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing
					
					if NOT rsEntry.EOF then
						useRelatedTGs = true
					end if
					
					rsEntry.close
										
					if useRelatedTGs then 
%>
						&nbsp;Payment&nbsp;<%=xssStr(allHotWords(7))%>:
						<select multiple size="3" name="optPmtTG" <%showMultiSelectTitle() %>>
							<option value="0" <%if inStr(pmtTGChk, ",0,")>0 OR request.form("optPmtTG")="" then response.write "selected" end if%>>All</option>
<%
							strSQL = "SELECT TypeGroupID, TypeGroup FROM tblTypeGroup WHERE (Active = 1) AND ((wsReservation = 1) OR (wsAppointment = 1) OR (wsResource=1) OR (wsEnrollment=1) OR (wsArrival=1) ) ORDER BY TypeGroup"
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>	
									<option value="<%=rsEntry("TypeGroupID")%>" <%if inStr(pmtTGChk, "," & TRIM(rsEntry("TypeGroupID")) & ",") > 0 then response.write "selected" end if%>><%=rsEntry("TypeGroup")%></option>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
						</select>
						<script type="text/javascript">
							document.frmParameter.optPmtTG.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(503))%>";
						</script>
					<% end if 'use related TGs %>

<%
			            strSQL = "SELECT DISTINCT Item#, PmtTypes FROM [Payment Types] WHERE (Active = 1) AND (Deleted = 0)"
			           response.write debugSQL(strSQL, "SQL")
			            rsEntry.CursorLocation = 3
			            rsEntry.open strSQL, cnWS
			            Set rsEntry.ActiveConnection = Nothing

%>     
					&nbsp;Payment Method:
					<select name="optPayMeth" size="1">
						<option value="0" <%if inStr(stTypeChk, ",0,")>0 OR request.form("optPayMeth")="" then response.write "selected" end if%>><%=xssStr(allHotWords(149))%>&nbsp;<%= getHotWord(59)%></option>
<%
			            do while NOT rsEntry.EOF
%>	
			            <option value="<%=rsEntry("Item#")%>" <%if inStr(stTypeChk, "," & TRIM(rsEntry("Item#")) & ",") > 0 then response.write "selected" end if%>><%=rsEntry("PmtTypes")%></option>
<%
				            rsEntry.MoveNext
			            loop
			            rsEntry.close
%>
					</select>

						&nbsp;View:
						<select name="optDisMode">
						  <option value="0" <%if showDetails="0" then response.write "selected" end if%>>By Sessions</option>
						  <option value="2" <%if showDetails="2" then response.write "selected" end if%>>By Instructor</option>
						  <option value="3" <%if showDetails="3" then response.write "selected" end if%>>By Date</option>
						  <option value="4" <%if showDetails="4" then response.write "selected" end if%>>By Type Group</option>
						  <option value="5" <%if showDetails="5" then response.write "selected" end if%>>By Visit Type</option>
						  <option value="6" <%if showDetails="6" then response.write "selected" end if%>>Roll Sheet</option>
						  <option value="1" <%if showDetails="1" then response.write "selected" end if%>>Summary</option>
						</select>
						<script type="text/javascript">
							document.frmParameter.optDisMode.options[0].text = "By " + '<%=jsEscSingle(allHotWords(5))%>'; // classes
							document.frmParameter.optDisMode.options[1].text = "By " + '<%=jsEscSingle(allHotWords(6))%>'; // instructor
							document.frmParameter.optDisMode.options[3].text = "By " + '<%=jsEscSingle(allHotWords(7))%>'; // program
							document.frmParameter.optDisMode.options[4].text = "By " + '<%=jsEscSingle(allHotWords(1))%>'; // session type
						</script>
						

						<br />
					<% showDateArrows("frmParameter") %> 
					<br />
						<input type="button" name="Button" value="Generate" onClick="genReport();"></b>
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
				else%>
						<% exportToExcelButton %>
				<%end if%>
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
				 else%>
						<% 'JM - 47_2381
						useTagSubtract = true 
						taggingButtons("frmParameter") %>
				<%end if%>
						<% savingButtons "frmParameter", "Attendance with Revenue" %>
						</td>
						</tr>
						
						</form>
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" class="mainTextBig"> 
					
					<table class="mainText center" id="attendanceWithRevenueGenTag" width="95%" border="0" cellspacing="0" cellpadding="0">
						<tr>
						<td class="mainTextBig center" colspan="2" valign="top">&nbsp;</td>
						</tr>
						<tr > 
						<td class="mainTextBig center" colspan="2" valign="top">
		<% 
		end if			'end of frmExpreport value check before /head line	  
				'

							if varGen=True then 
								if request.form("frmExpReport")="true" then
									Dim stFilename
									if showDetails="0" then 
										stFilename="attachment; filename=Attendance " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									elseif showDetails="1" then
										stFilename="attachment; filename=Attendance-Summary " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									else
										stFilename="attachment; filename=Attendance " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									end if
									Response.ContentType = "application/vnd.ms-excel" 
									Response.AddHeader "Content-Disposition", stFilename 
								end if
								
								curDate=""
								curTime=""
								curVisitType=""
								curTrainer=""
								curLoc=""
								curTG=""
								curAsst1=""
								curAsst2=""
								curClassName=""
								curClassID=0
								tmpRevPerVisit=0
								tmpRev=0
								tmpMemRev=0
								TotMemRev=0
								TotRev=0
								GTMemRev=0
								GTRev=0
								tmpHCPaid=0
								tmpHCComp=0
								tmpHCNoShow=0
								tmpHCMem=0
								TotHCPaid=0
								TotHCComp=0
								TotHCNoShow=0
								TotHCMem=0
								GrandTotHCPaid=0
								GrandTotHCComp=0
								GrandTotHCNoShow=0
								GrandTotHCMem=0
								tmpWeb=0
								TotWeb=0
								GrandTotWeb=0
								
								strSQL = "SELECT [VISIT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, [VISIT DATA].ClassDate, [VISIT DATA].ClassTime, " 
								strSQL = strSQL & " [VISIT DATA].TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName, [VISIT DATA].TypeTaken, tblVisitTypes.TypeName, [PAYMENT DATA].Returned, " 
								strSQL = strSQL & " [VISIT DATA].Cancelled, [VISIT DATA].Missed, Location.LocationName, [VISIT DATA].Location, [VISIT DATA].PmtRefNo, " 
								strSQL = strSQL & " [VISIT DATA].VisitRefNo, [VISIT DATA].[Value], [VISIT DATA].NumDeducted, [PAYMENT DATA].ExpDate, TRAINERS.DisplayName, [PAYMENT DATA].Remaining, [PAYMENT DATA].RealRemaining, " 
								strSQL = strSQL & " [PAYMENT DATA].Type, [PAYMENT DATA].NumClasses, CLIENTS.Deleted, [VISIT DATA].TrainerID2, [VISIT DATA].TrainerID3, " 
								strSQL = strSQL & " tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup, tblClassSch.ClassID, tblClassDescriptions.ClassName, [VISIT DATA].Webscheduler, tblResources.ResourceName, tblReservation.VisitRefNo AS ApptVisitRefNo " 

								if request.form("optPayMeth")<>"" and request.form("optPayMeth")<>"0" then
									strSQL = strSQL & " , SUM(ISNULL(tblSDPayments.SDPaymentAmount, 0)) as PaymentAmount "
								else
									strSQL = strSQL & " , (ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0)) as PaymentAmount "
								end if
								if showdetails="1" then ' select membership info if in summary
									strSQL = strSQL & ", Membership.ActiveDate "
								end if
								strSQL = strSQL & "FROM tblResourceSchedules INNER JOIN tblResources ON tblResourceSchedules.ResourceID = tblResources.ResourceID RIGHT OUTER JOIN tblClassDescriptions INNER JOIN tblClassSch ON tblClassDescriptions.ClassDescriptionID = tblClassSch.DescriptionID ON tblResourceSchedules.RefClass = tblClassSch.ClassID AND tblResourceSchedules.StartDate <= tblClassSch.ClassDate AND tblResourceSchedules.EndDate >= tblClassSch.ClassDate RIGHT OUTER JOIN TRAINERS RIGHT OUTER JOIN tblVisitTypes RIGHT OUTER JOIN "
								strSQL = strSQL & "[VISIT DATA] INNER JOIN CLIENTS ON CLIENTS.ClientID = [VISIT DATA].ClientID INNER JOIN tblTypeGroup ON [VISIT DATA].TypeGroup = tblTypeGroup.TypeGroupID INNER JOIN Location ON [VISIT DATA].Location = Location.LocationID ON tblVisitTypes.TypeID = [VISIT DATA].VisitType LEFT OUTER JOIN [PAYMENT DATA] ON [VISIT DATA].PmtRefNo = [PAYMENT DATA].PmtRefNo LEFT OUTER JOIN [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo ON TRAINERS.TrainerID = [VISIT DATA].TrainerID ON  tblClassSch.ClassID = [VISIT DATA].ClassID AND tblClassSch.ClassDate = [VISIT DATA].ClassDate LEFT OUTER JOIN tblReservation ON [VISIT DATA].VisitRefNo = tblReservation.VDID "
								' added condition around join to improve report when filtering by all payment type groups BQL 2/29/8
								if pmtTGList<>"" AND pmtTGList<>"0" then
									strSQL = strSQL & " INNER JOIN tblTypeGroup PaymentTypeGroup ON PaymentTypeGroup.TypeGroupID = [PAYMENT DATA].TypeGroup "
								end if
								' end 2/29/8
								if request.form("optPayMeth")<>"" and request.form("optPayMeth")<>"0" then
									strSQL = strSQL & " INNER JOIN tblSDPayments ON [Sales Details].SDID = tblSDPayments.SDID "
									strSQL = strSQL & " INNER JOIN tblPayments ON tblPayments.PaymentID = tblSDPayments.PaymentID "
			                    end if		
								if request.form("optFilterTagged")="on" then
									strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
									if session("mVarUserID")<>"" then
										strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
									end if
									strSQL = strSQL & " ) "
								end if
								if showdetails="1" then ' check membership if in summary
									strSQL = strSQL & "LEFT OUTER JOIN " & strMemberSubSQL & " Membership ON Clients.ClientID = Membership.ClientID "
								end if
								strSQL = strSQL & "WHERE (CLIENTS.Deleted = 0) AND ([VISIT DATA].ClassDate >=  " & DateSep & cSDate & DateSep & ") AND ([VISIT DATA].ClassDate <=  " & DateSep & cEDate & DateSep & ") "
								' BQL 11/5/2008 - added to filter out XR payment rows for visits at another location
								if request.form("optCRVisits")<>"on" then
									strSQL = strSQL & " AND (tblVisitTypes.TypeID<>-1 OR tblVisitTypes.TypeID IS NULL) "
								end if
								if cLoc<>0 then
									strSQL = strSQL & " AND ([VISIT DATA].Location=" & cLoc & ") "
								end if
								'CB 5/28/2008 - ADDED OR Type=9 so unpaids are returned
								if cLoc1<>0 then
									strSQL = strSQL & " AND ([Sales Details].Location=" & cLoc1 & " OR [PAYMENT DATA].Type=9) "
								end if									
								if cTrainer<>"0" and cTrainer<>"" then
									strSQL = strSQL & " AND [VISIT DATA].TrainerID=" & CLNG(cTrainer)
								end if
								if cTime<>"0" and cTime<>"" then
								  if cTime="null" then 
								    strSQL = strSQL & " AND [VISIT DATA].ClassTime IS NULL "
								  else
									  strSQL = strSQL & " AND [VISIT DATA].ClassTime=" & TimeSepB & cTime & TimeSepA
								  end if
								end if
								if visitTGList<>"" AND visitTGList<>"0" then
									strSQL = strSQL & "AND [tblTypeGroup].TypeGroupID IN (" & visitTGList & ") "
								else ' BJD 4/23/07 - Add null trainerID check for media visits
									strSQL = strSQL & " AND (tblTypegroup.wsResource=0) AND (tblTypegroup.wsMedia=0) AND ([Visit Data].TrainerID IS NOT NULL) "
								end if
								if pmtTGList<>"" AND pmtTGList<>"0" then
									strSQL = strSQL & "AND PaymentTypeGroup.TypeGroupID IN (" & pmtTGList & ") "
								' removed to improve report when filtering by all payment type groups BQL 2/29/8
								'else
								'	strSQL = strSQL & " AND (PaymentTypeGroup.wsResource=0) "
								end if
								if request.form("optPayMeth")<>"" and request.form("optPayMeth")<>"0" then
				                    strSQL = strSQL & " AND (tblPayments.PaymentMethod=" & stType & ")"
			                    end if		
								' opening for days of week
								strSQL = strSQL & "AND (1 = 0 "
								if request.form("chkDaySun")="on" then
									strSQL = strSQL & "OR DATEPART(WEEKDAY, [VISIT DATA].ClassDate) = 1 "
								end if
								if request.form("chkDayMon")="on" then
									strSQL = strSQL & "OR DATEPART(WEEKDAY, [VISIT DATA].ClassDate) = 2 "
								end if
								if request.form("chkDayTue")="on" then
									strSQL = strSQL & "OR DATEPART(WEEKDAY, [VISIT DATA].ClassDate) = 3 "
								end if
								if request.form("chkDayWed")="on" then
									strSQL = strSQL & "OR DATEPART(WEEKDAY, [VISIT DATA].ClassDate) = 4 "
								end if
								if request.form("chkDayThu")="on" then
									strSQL = strSQL & "OR DATEPART(WEEKDAY, [VISIT DATA].ClassDate) = 5 "
								end if
								if request.form("chkDayFri")="on" then
									strSQL = strSQL & "OR DATEPART(WEEKDAY, [VISIT DATA].ClassDate) = 6 "
								end if
								if request.form("chkDaySat")="on" then
									strSQL = strSQL & "OR DATEPART(WEEKDAY, [VISIT DATA].ClassDate) = 7 "
								end if
								strSQL = strSQL & ") "
								if request.form("optPayMeth")<>"" and request.form("optPayMeth")<>"0" then
									strSQL = strSQL & " GROUP BY [VISIT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, [VISIT DATA].ClassDate, [VISIT DATA].ClassTime, " 
									strSQL = strSQL & " [VISIT DATA].TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName, [VISIT DATA].TypeTaken, tblVisitTypes.TypeName, " 
									strSQL = strSQL & " [VISIT DATA].Cancelled, [VISIT DATA].Missed, Location.LocationName, [VISIT DATA].Location, [VISIT DATA].PmtRefNo, " 
									strSQL = strSQL & " [VISIT DATA].VisitRefNo, [VISIT DATA].[Value], [VISIT DATA].NumDeducted, [PAYMENT DATA].ExpDate, [PAYMENT DATA].Returned, [PAYMENT DATA].Remaining, [PAYMENT DATA].RealRemaining, " 
									strSQL = strSQL & " [PAYMENT DATA].Type, [PAYMENT DATA].NumClasses, CLIENTS.Deleted, [VISIT DATA].TrainerID2, [VISIT DATA].TrainerID3, " 
									strSQL = strSQL & " tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup, tblClassSch.ClassID, tblClassDescriptions.ClassName, [VISIT DATA].Webscheduler, tblResources.ResourceName, tblReservation.VisitRefNo " 
									if showdetails="1" then ' select membership info if in summary
										strSQL = strSQL & ", Membership.ActiveDate "
									end if
								end if
								if request.form("frmTagClients")<>"true" then
									if showdetails="0" then
										strSQL = strSQL & " ORDER BY tblTypeGroup.TypeGroup, [VISIT DATA].Classdate DESC, [VISIT DATA].ClassTime, tblClassSch.ClassID, " & GetTrnOrderBy() & ", tblVisitTypes.TypeName, Clients.LastName"
									elseif showdetails="1" then
										strSQL = strSQL & " ORDER BY [VISIT DATA].ClassDate DESC, [VISIT DATA].ClassTime, tblClassSch.ClassID, " & GetTrnOrderBy() & ", [VISIT DATA].TrainerID, tblTypeGroup.TypeGroup, tblVisitTypes.TypeName "
									elseif showdetails="2" then
										strSQL = strSQL & " ORDER BY " & GetTrnOrderBy() & ", tblTypeGroup.TypeGroup, [VISIT DATA].Classdate DESC, [VISIT DATA].ClassTime, tblClassSch.ClassID, tblVisitTypes.TypeName, Clients.LastName"
									elseif showdetails="3" then
										strSQL = strSQL & " ORDER BY [VISIT DATA].Classdate DESC, [VISIT DATA].ClassTime, tblClassSch.ClassID, " & GetTrnOrderBy() & ", tblVisitTypes.TypeName, Clients.LastName"
									elseif showdetails="4" then
										strSQL = strSQL & " ORDER BY tblTypeGroup.TypeGroup, " & GetTrnOrderBy() & ", [VISIT DATA].Classdate DESC, [VISIT DATA].ClassTime, tblClassSch.ClassID, tblVisitTypes.TypeName, Clients.LastName"
									elseif showdetails="5" then
										strSQL = strSQL & " ORDER BY tblVisitTypes.TypeName, [VISIT DATA].Classdate DESC, [VISIT DATA].ClassTime, tblClassSch.ClassID, tblTypeGroup.TypeGroup, " & GetTrnOrderBy() & ", Clients.LastName"
									elseif showdetails="6" then
										strSQL = strSQL & " ORDER BY [VISIT DATA].Classdate, [VISIT DATA].ClassTime, tblClassSch.ClassID,  tblClassDescriptions.ClassName, tblVisitTypes.TypeName, " & GetTrnOrderBy()
									end if
								end if
							
							response.write debugSQL(strSQL, "Attendance Export") 
								
								if request.form("frmTagClients")="true" then
								        ' -- CCP Removed limitation, Bug 360, 6/30/09
									'if showdetails="1" or showdetails="6" then %>
										<!--<script>
											alert("Summary report results can't be tagged");
										</script>-->
									<% 'else
										if request.form("frmTagClientsNew")="true" then
											clearAndTagQuery(strSQL)
										'JM - 47_2381
										elseif request.form("frmUnTagClients")="true" then
											tagSubtract(strSQL)
										else
											tagQuery(strSQL)
										end if
									'end if 'summary or roll sheet
									strSQL = "SELECT StudioID FROM Studios WHERE 1=0 "
								end if
								
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								Set rsEntry.ActiveConnection = Nothing

	
								if showdetails="0" then										
									curDate=""
									curTime=""
									curVisitType=""
									curTrainer=""
									curLoc=""
									curTG=""
								'*********************  REPORT By CLASS  *****************************
								%>
									<table class="mainText" width="100%"  border="0" cellpadding="0" cellspacing="0">
								<% 
									if NOT rsEntry.EOF then			'EOF
										do while NOT rsEntry.EOF
											if curTG<>rsEntry("Typegroup") and curTG="" then
												%>
												<tr>
													<td height="35" colspan="8" class="maintextbig"><strong><%=UCASE(rsEntry("TypeGroup"))%></strong></td>
												</tr>
												<%
											end if
											if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) or curClassID<>rsEntry("ClassID") then	'if this is a new class then write the header cells
												if curDate<> "" then	'if this isn't the first record total cells
													%>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
													<tr>
														<td nowrap><div class="right"><strong> HEAD COUNT&nbsp;</strong></div></td>
														<td class="right nowrap"><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
														<td class="left nowrap"><strong><%=tmpHCPaid%></strong></td>
														<td class="right nowrap"><strong>COMP:&nbsp;</strong></td>
														<td class="left nowrap"><strong><%=tmpHCComp%></strong></td>
														<td class="right nowrap"><strong>NO SHOW:&nbsp;</strong></td>
														<td class="left nowrap"><strong><%=tmpHCNoShow%></strong></td>
														<td class="right nowrap"><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
														<td class="left nowrap"><strong><%=tmpWeb%></strong></td>
													<% if NOT request.form("frmExpReport")="true" then %>
														<td class="right nowrap"><strong><%=FmtCurrency(tmpRev)%></strong></td>
													<% else %>
														<td class="right nowrap"><strong><%=FmtNumber(tmpRev)%></strong></td>
													<% end if %>
													</tr>
													<tr>
													  <td colspan="11" class="nowrap">&nbsp;</td>
													</tr>
												<%	
												end if
												if curTG<>rsEntry("Typegroup") and curTG<>"" then
												%>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
													<tr class="maintextbig">
														<td nowrap><div class="right"><strong><%=UCASE(curTG)%>&nbsp;HEAD COUNT&nbsp;</strong></div></td>
														<td class="right nowrap"><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
														<td class="left nowrap"><strong><%=TotHCPaid%></strong></td>
														<td class="right nowrap"><strong>COMP:&nbsp;</strong></td>
														<td class="left nowrap"><strong><%=TotHCComp%></strong></td>
														<td class="right nowrap"><strong>NO SHOW:&nbsp;</strong></td>
														<td class="left nowrap"><strong><%=TotHCNoShow%></strong></td>
														<td class="right nowrap"><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
														<td class="left nowrap"><strong><%=TotWeb%></strong></td>
<% if NOT request.form("frmExpReport")="true" then %>
														<td class="right nowrap"><strong><%=FmtCurrency(TotRev)%></strong></td>
<% else %>
														<td class="right nowrap"><strong><%=FmtNumber(TotRev)%></strong></td>
<% end if %>
													</tr>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
													<tr>
													  <td colspan="11" nowrap>&nbsp;</td>
													</tr>
													<tr>
														<td height="35" colspan="10" class="maintextbig"><strong><%=UCASE(rsEntry("TypeGroup"))%></strong></td>
													</tr>
												<%	
													TotHCPaid=0
													TotHCComp=0
													TotHCNoShow=0
													TotWeb=0
													TotRev=0
												end if
												%>
												<tr class="blackHeader">
												  <td>&nbsp;<strong><%=FmtDateTime(rsEntry("ClassDate"))%></strong></td>
												  <td>&nbsp;<strong><%if isNull(rsEntry("ClassTime")) then response.Write "TBD" else response.Write FmtTimeShort(rsEntry("ClassTime")) end if%></strong></td>
												  <td colspan="4">&nbsp;<strong>
<% '******************************************************************************** %>
<% if NOT request.form("frmExpReport")="true" then %>
													<% if NOT isNull(rsEntry("ClassID")) then %>
												  		<a href="adm_cls_list.asp?pDate=<%=DateValue(rsEntry("ClassDate"))%>&pClsID=<%=rsEntry("ClassID")%>">
													<% elseif NOT isNull(rsEntry("ApptVisitRefNo"))then %>
														<a href="adm_appt_e.asp?id=<%=rsEntry("ApptVisitRefNo")%>">
													<% end if %>
<% end if %>
												  		<%If isnull(rsEntry("ClassName")) then response.write rsEntry("TypeName") else response.write rsEntry("ClassName") end if%>
<% if NOT request.form("frmExpReport")="true" then %>
													</a>
<% end if %>
												  </strong></td>
												  <td><strong><% If request.form("optSaleLoc")="0" then%><%=rsEntry("LocationName")%><% end if %></strong>&nbsp;</td>
												  <td colspan="3" class="right"><strong><%=FmtTrnNameNew(rsEntry, false)%></strong>&nbsp;</td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
												<tr height="2">
													  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
												</tr>
<% end if %>
												<tr>
												  <td width="20%" class="left" colspan="2" nowrap><strong><%=session("ClientHW")%> </strong></td>
												  <td width="17%" class="left" colspan="2" nowrap><strong><%=hw61%></strong></td>
												  <td width="10%" class="right" nowrap><strong><%= getHotWord(117)%> </strong></td>
												  <td width="10%" class="right" nowrap><strong>Rem </strong></td>
												  <td width="10%" class="center" nowrap><strong><%=hw113%> <%= getHotWord(36)%>&nbsp;</strong></td>
												  <td width="10%" class="center" nowrap><strong><%= getHotWord(172)%>&nbsp;</strong></td>
												  <td width="10%" class="center" nowrap><strong>No Show&nbsp;</strong></td>
												  <td width="13%" class="right" nowrap><strong>Rev. Per Visit</strong></td>
												</tr>
												<%
												tmpHCPaid=0
												tmpHCComp=0
												tmpHCNoShow=0
												tmpWeb=0
												tmpRev=0
											end if	'new description ended. Continue with next for both new and old descriptions
											curDate=rsEntry("ClassDate")
											curTime=rsEntry("ClassTime")
											curVisitType=rsEntry("TypeName")
											curTrainerID=CLNG(rsEntry("TrainerID"))
											curTrainer=FmtTrnNameNew(rsEntry,false)
											curLoc=rsEntry("LocationName")
											curTG=rsEntry("Typegroup")
											curClassID=rsEntry("ClassID")
									'response.write tmpRevPerVisit & " -- 1<br />"
											if rsEntry("NumDeducted")="0" OR rsEntry("Returned") then
												tmpRevPerVisit=0
											else
												if rsEntry("Type")=1 then
													if rsEntry("NumClasses")=0 or isnull(rsEntry("NumClasses")) then
														tmpRevPerVisit=rsEntry("PaymentAmount")*rsEntry("NumDeducted")
													else
														'CB - updated 9_25_2007 to factor in NumDeducted
									'response.write tmpRevPerVisit & " -- 1.3<br />"
									'response.write rsEntry("PaymentAmount") & " -- " & rsEntry("NumClasses") & " -- " & rsEntry("NumDeducted") & "<br />"
														tmpRevPerVisit=(rsEntry("PaymentAmount")/rsEntry("NumClasses")) * rsEntry("NumDeducted")
									'response.write tmpRevPerVisit & " -- 1.6<br />"
													end if
													
												'elseif rsEntry("Type")=2 OR rsEntry("Type")=3 then
												elseif rsEntry("Type")<>9 then	'CB 10/21/2008
													if (rsEntry("NumClasses")-rsEntry("Remaining"))=0 or isnull((rsEntry("NumClasses")-rsEntry("Remaining"))) then
														tmpRevPerVisit=rsEntry("PaymentAmount")*rsEntry("NumDeducted")
													else
														tmpRevPerVisit=rsEntry("PaymentAmount")/(rsEntry("NumClasses")-rsEntry("Remaining")) *rsEntry("NumDeducted")
													end if
												else
													tmpRevPerVisit=0
												end if
											end if
									'response.write tmpRevPerVisit & " -- 2<br />"
											tmpRev = tmpRev + tmpRevPerVisit
											TotRev = TotRev + tmpRevPerVisit
											GTRev = GTRev + tmpRevPerVisit
											
											If rsEntry("Webscheduler") then
												tmpWeb=tmpWeb+1
												TotWeb=TotWeb+1
												GrandTotWeb=GrandTotWeb+1
											end if

											If rsEntry("Value")=1 then
												tmpHCPaid=tmpHCPaid+1
												TotHCPaid=TotHCPaid+1
												GrandTotHCPaid=GrandTotHCPaid+1
											else
												tmpHCComp=tmpHCComp+1
												TotHCComp=TotHCComp+1
												GrandTotHCComp=GrandTotHCComp+1
											end if
											
											If rsEntry("Missed") then
												tmpHCNoShow=tmpHCNoShow+1
												TotHCNoShow=TotHCNoShow+1
												GrandTotHCNoShow=GrandTotHCNoShow+1
											end if
									'response.write tmpRevPerVisit & " -- 3<br />"

											%>
												<tr>
<% if NOT request.form("frmExpReport")="true" then %>
												  <td colspan="2" class="left" nowrap><a href="main_info.asp?ID=<%=rsEntry("ClientID")%>&fl=true"><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></a></td>
<% else %>
												  <td colspan="2" class="left" nowrap><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></td>
<% end if %>
												  <td colspan="2" class="left" nowrap><%=rsEntry("TypeTaken")%></td>
												  <td class="right" nowrap><%=FmtDateShort(rsEntry("ExpDate"))%></td>
												  <td class="right" nowrap>
<%
											if rsEntry("RealRemaining")=rsEntry("Remaining") then
												if rsEntry("Type") = 9 then
													response.write "<span style=""color:#990000;"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write rsEntry("RealRemaining")
												end if
											else
												if rsEntry("Type") = 9 then
													response.write "<span style=""color:#990000;"" title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write "<span title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")
												end if
											end if
%>												  
												  </td>
												  <td class="center" nowrap><% if rsEntry("Value") then response.write "Yes" else response.write "No" end if %>&nbsp;</td>
												  <%if true=false then %>
												  <td class="center" nowrap><input type="checkbox" name="chkCancelled" <%if rsEntry("Cancelled") then response.write "checked" end if %>></td>
												  <td class="center" nowrap><input type="checkbox" name="chkMissed" <%if rsEntry("Missed") then response.write "checked" end if %>></td>
												  <%else %>
												  <td class="center" nowrap><%if rsEntry("Cancelled") then response.write "Yes" else response.write "No" end if %></td>
												  <td class="center" nowrap><%if rsEntry("Missed") then response.write "Yes" else response.write "No" end if %></td>
												  <%end if%>
<% if NOT request.form("frmExpReport")="true" then %>
												  <td class="right" nowrap><% if rsEntry("Returned") then response.write"<span style=""color:red"">(returned)</span> " end if %><%=FmtCurrency(tmpRevPerVisit)%></td>
<% else %>
												  <td class="right" nowrap><%=FmtNumber(tmpRevPerVisit)%></td>
<% end if %>
												</tr>
											<%		
											rsEntry.MoveNext
										loop
										%>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
												  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr>
												<td nowrap><div class="right"><strong> HEAD COUNT&nbsp;</strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
<% end if %>
											</tr>
											<tr>
											  <td colspan="10" nowrap>&nbsp;</td>
											</tr>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
												  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr class="maintextbig">
												<td nowrap><div class="right"><strong><%=UCASE(curTG)%>&nbsp;HEAD COUNT&nbsp;</strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotWeb%></strong></td>
<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
<% end if %>
											</tr>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
												  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
										<%
									end if	'eof
									%>
									<tr>
									  <td colspan="10" nowrap>&nbsp;</td>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
									<tr class="maintextbig">
										<td nowrap><div class="right"><strong>HEAD COUNT GRAND TOTAL&nbsp;</strong></div></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCPaid%></strong></td>
										<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCComp%></strong></td>
										<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCNoShow%></strong></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotWeb%></strong></td>
<% if NOT request.form("frmExpReport")="true" then %>
										<td class="right" nowrap><strong><%=FmtCurrency(GTRev)%></strong></td>
<% else %>
										<td class="right" nowrap><strong><%=FmtNumber(GTRev)%></strong></td>
<% end if %>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
									</table>
								<%
								elseif showdetails="1" then										
								'********************* REPORT-SUMMARY  *****************************
									curDate=""
									curClassID=""
									curApptID=""
									curTime=""
									curVisitType=""
									curTrainer=""
									curLoc=""
									curTG=""
								%>
									<table class="mainText" width="100%"  border="0" cellpadding="0" cellspacing="0">
								<% 
									if NOT rsEntry.EOF then			'EOF
										do while NOT rsEntry.EOF
											if curDate<>rsEntry("ClassDate") then	'if this is a new day then write the header cells
												if curDate<> "" then	'if this isn't the first record total cells
													if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or curTG<>rsEntry("Typegroup") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) or curClassID<>rsEntry("ClassID") then	'if this is a new class then write the header cells
														if curTime<> "" then	'if this isn't the first record total cells
														%>
															<tr>
															  <td class="left" nowrap><%=FmtTimeShort(curTime)%></td>
															  <td class="left" nowrap><%=curTG%></td>
															  <td class="left" nowrap><%=curTrainer%></td>
															  <td class="left" nowrap>									  
<% if NOT request.form("frmExpReport")="true" then %>
															<% if curClassID<>"" then %>
																<a href="adm_cls_list.asp?pDate=<%=curDate%>&pClsID=<%=curClassID%>">
															<% elseif curApptID<>"" then %>
																<a href="adm_appt_e.asp?id=<%=curApptID%>">
															<% end if %>
<% end if %>
												  			<%=curVisitType%>
<% if NOT request.form("frmExpReport")="true" then %>
															</a>
<% end if %>
															  </td>
															  <td class="right" nowrap><%=tmpHCPaid%></td>
															  <td class="right" nowrap><%=tmpHCComp%></td>
															  <td class="right" nowrap><%=tmpHCNoShow%></td>
															  <td class="right" nowrap><%=tmpWeb%></td>
															  <td class="right" nowrap><%=tmpHCMem%></td>
															<% if NOT request.form("frmExpReport")="true" then %>
															  <td class="right" nowrap><%=FmtCurrency(tmpMemRev)%></td>
															  <td class="right" nowrap><%=FmtCurrency(tmpRev)%></td>
															<% else %>
															  <td class="right" nowrap><%=FmtNumber(tmpMemRev)%></td>
															  <td class="right" nowrap><%=FmtNumber(tmpRev)%></td>
															<% end if %>
															</tr>
														<%
														end if	'end of curtime<>""
														tmpHCPaid=0
														tmpHCComp=0
														tmpHCNoShow=0
														tmpWeb=0
														tmpHCMem=0
														tmpRev=0
														tmpMemRev=0
													end if	'end of curTime<>rsentry("classtime")...
												%>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														  <td colspan="11" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
													</tr>
<% end if %>
													<tr>
														<td colspan="4" nowrap><div class="right"><strong><%=FmtDateLong(curDate)%> <%= getHotWord(22)%>:</strong></div></td>
														<td class="right" nowrap><strong><%=TotHCPaid%></strong></td>
														<td class="right" nowrap><strong><%=TotHCComp%></strong></td>
														<td class="right" nowrap><strong><%=TotHCNoShow%></strong></td>
														<td class="right" nowrap><strong><%=TotWeb%></strong></td>
														<td class="right" nowrap><strong><%=TotHCMem%></strong></td>
													<% if NOT request.form("frmExpReport")="true" then %>
														<td class="right" nowrap><strong><%=FmtCurrency(TotMemRev)%></strong></td>
														<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
													<% else %>
														<td class="right" nowrap><strong><%=FmtNumber(TotMemRev)%></strong></td>
														<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
													<% end if %>
													</tr>
													<tr>
													  <td colspan="11" nowrap>&nbsp;</td>
													</tr>
												<%	
												end if	'end of curDate<>""
												%>
												<tr>
												  <td colspan="11" class="blackHeader left"><strong><%=FmtDateLong(rsEntry("ClassDate"))%>&nbsp;</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
												<tr height="2">
													  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
												</tr>
<% end if %>
												<tr>
												  <td width="8%" class="left" nowrap><strong><%= getHotWord(58)%> </strong></td>
												  <td width="8%" class="left" nowrap><strong><%=hw7%> </strong></td>
												  <td width="10%" class="left" nowrap><strong><%=hw6%></strong></td>
												  <td width="10%" class="left" nowrap><strong>Visit Type </strong></td>
												  <td width="10%" class="right" nowrap><strong>&nbsp;&nbsp;<%= getHotWord(36)%> <%=session("ClientHW")%>s</strong></td>
												  <td width="10%" class="right" nowrap><strong>&nbsp;&nbsp;Comp <%=session("ClientHW")%>s</strong></td>
												  <td width="8%" class="right" nowrap><strong>No Shows</strong></td>
												  <td width="8%" class="right" nowrap><strong>Web</strong></td>
												  <td width="8%" class="right" nowrap><strong><%= getHotWord(220)%></strong></td>
												  <td width="10%" class="right" nowrap><strong>&nbsp;Mem. Revenue</strong></td>
												  <td width="10%" class="right" nowrap><strong>Total Revenue</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
												<tr height="1">
													<td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
												</tr>
<% end if %>
												<%
												TotHCPaid=0
												TotHCComp=0
												TotHCNoShow=0
												TotWeb=0
												TotHCMem=0
												TotMemRev=0
												TotRev=0
											else	'If curdate is the same as rsEntry(ClassDate) but the time is different
												if curDate<> "" then	'if this isn't the first record total cells
													if curTime<>rsEntry("ClassTime") or curTG<>rsEntry("Typegroup") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
														if curTime<> "" then	'if this isn't the first record total cells
														%>
															<tr>
															  <td class="left" nowrap><%=FmtTimeShort(curTime)%></td>
															  <td class="left" nowrap><%=curTG%></td>
															  <td class="left" nowrap><%=curTrainer%></td>
															  <td class="left" nowrap>
<% if NOT request.form("frmExpReport")="true" then %>
																<% if curClassID<>"" then %>
																	<a href="adm_cls_list.asp?pDate=<%=curDate%>&pClsID=<%=curClassID%>">
																<% elseif curApptID<>"" then %>
																	<a href="adm_appt_e.asp?id=<%=curApptID%>">
																<% end if %>
<% end if %>
																<%=curVisitType%>
<% if NOT request.form("frmExpReport")="true" then %>
																</a>
<% end if %>
															  </td>
															  <td class="right" nowrap><%=tmpHCPaid%></td>
															  <td class="right" nowrap><%=tmpHCComp%></td>
															  <td class="right" nowrap><%=tmpHCNoShow%></td>
															  <td class="right" nowrap><%=tmpWeb%></td>
															  <td class="right" nowrap><%=tmpHCMem%></td>
														<% if NOT request.form("frmExpReport")="true" then %>
															  <td class="right" nowrap><%=FmtCurrency(tmpMemRev)%></td>
															  <td class="right" nowrap><%=FmtCurrency(tmpRev)%></td>
														<% else %>
															  <td class="right" nowrap><%=FmtNumber(tmpMemRev)%></td>
															  <td class="right" nowrap><%=FmtNumber(tmpRev)%></td>
														<% end if %>
															</tr>
														<%
														end if	'end of curtime<>""
														tmpHCPaid=0
														tmpHCComp=0
														tmpHCNoShow=0
														tmpWeb=0
														tmpHCMem=0
														tmpMemRev=0
														tmpRev=0
													end if	'end of curTime<>rsentry("classtime")...
												end if	'end of curDate<>""
											end if	'end of curDate<>rsEntry("ClassDate")
											curDate=rsEntry("ClassDate")
											curTime=rsEntry("ClassTime")
											curVisitType=rsEntry("TypeName")
											curTrainerID=CLNG(rsEntry("TrainerID"))
											curTrainer=FmtTrnNameNew(rsEntry,false)
											curLoc=rsEntry("LocationName")
											curTG=rsEntry("Typegroup")
											curClassID=rsEntry("ClassID")
											curApptID=rsEntry("ApptVisitRefNo")
											
											if rsEntry("NumDeducted")=0 then
												tmpRevPerVisit=0
											else
												If rsEntry("Type")=1 then
													If rsEntry("NumClasses")=0 or isnull(rsEntry("NumClasses")) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=(rsEntry("PaymentAmount")/rsEntry("NumClasses")) * rsEntry("NumDeducted")
													end if
												'elseif rsEntry("Type")=2 OR rsEntry("Type")=3 then
												elseif rsEntry("Type")<>9 then	'CB 10/21/2008
													If (rsEntry("NumClasses")-rsEntry("Remaining"))=0 or isnull((rsEntry("NumClasses")-rsEntry("Remaining"))) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=rsEntry("PaymentAmount")/(rsEntry("NumClasses")-rsEntry("Remaining"))
													end if
												else
													tmpRevPerVisit=0
												end if
											end if
											if NOT isNull(rsEntry("ActiveDate")) then ' sum mem total if client is member
												tmpMemRev = tmpMemRev + tmpRevPerVisit
												TotMemRev = TotMemRev + tmpRevPerVisit
												GTMemRev = GTMemRev + tmpRevPerVisit
												tmpHCMem = tmpHCMem + 1
												TotHCMem = TotHCMem + 1
												GrandTotHCMem = GrandTotHCMem + 1
											end if
											tmpRev = tmpRev + tmpRevPerVisit
											TotRev = TotRev + tmpRevPerVisit
											GTRev = GTRev + tmpRevPerVisit
											
											If rsEntry("Webscheduler") then
												tmpWeb=tmpWeb+1
												TotWeb=TotWeb+1
												GrandTotWeb=GrandTotWeb+1
											end if

											If rsEntry("Value")=1 then
												tmpHCPaid=tmpHCPaid+1
												TotHCPaid=TotHCPaid+1
												GrandTotHCPaid=GrandTotHCPaid+1
											else
												tmpHCComp=tmpHCComp+1
												TotHCComp=TotHCComp+1
												GrandTotHCComp=GrandTotHCComp+1
											end if
											
											If rsEntry("Missed") then
												tmpHCNoShow=tmpHCNoShow+1
												TotHCNoShow=TotHCNoShow+1
												GrandTotHCNoShow=GrandTotHCNoShow+1
											end if
											rsEntry.MoveNext
										loop
										%>
										<tr>
										  <td class="left" nowrap><%=FmtTimeShort(curTime)%></td>
										  <td class="left" nowrap><%=curTG%></td>
										  <td class="left" nowrap><%=curTrainer%></td>
										  <td class="left" nowrap>
<% if NOT request.form("frmExpReport")="true" then %>
											<% if curClassID<>"" then %>
												<a href="adm_cls_list.asp?pDate=<%=curDate%>&pClsID=<%=curClassID%>">
											<% elseif curApptID<>"" then %>
												<a href="adm_appt_e.asp?id=<%=curApptID%>">
											<% end if %>
<% end if %>
											<%=curVisitType%>
<% if NOT request.form("frmExpReport")="true" then %>
											</a>
<% end if %>
										  </td>
										  <td class="right" nowrap><%=tmpHCPaid%></td>
										  <td class="right" nowrap><%=tmpHCComp%></td>
										  <td class="right" nowrap><%=tmpHCNoShow%></td>
										  <td class="right" nowrap><%=tmpWeb%></td>
										  <td class="right" nowrap><%=tmpHCMem%></td>
									<% if NOT request.form("frmExpReport")="true" then %>
										  <td class="right" nowrap><%=FmtCurrency(tmpMemRev)%></td>
										  <td class="right" nowrap><%=FmtCurrency(tmpRev)%></td>
									<% else %>
										  <td class="right" nowrap><%=FmtNumber(tmpMemRev)%></td>
										  <td class="right" nowrap><%=FmtNumber(tmpRev)%></td>
									<% end if %>
										</tr>
<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="1">
											  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
										</tr>
<% end if %>
										<tr>
											<td colspan="4" nowrap><div class="right"><strong><%=FmtDateLong(curDate)%> <%= getHotWord(22)%>:</strong></div></td>
											<td class="right" nowrap><strong><%=TotHCPaid%></strong></td>
											<td class="right" nowrap><strong><%=TotHCComp%></strong></td>
											<td class="right" nowrap><strong><%=TotHCNoShow%></strong></td>
											<td class="right" nowrap><strong><%=TotWeb%></strong></td>
											<td class="right" nowrap><strong><%=TotHCMem%></strong></td>
									<% if NOT request.form("frmExpReport")="true" then %>
											<td class="right" nowrap><strong><%=FmtCurrency(TotMemRev)%></strong></td>
											<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
									<% else %>
											<td class="right" nowrap><strong><%=FmtNumber(TotMemRev)%></strong></td>
											<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
									<% end if %>
										</tr>
										<tr>
										  <td colspan="11" nowrap>&nbsp;</td>
										</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
									<tr>
									  <td colspan="4" class="maintextbig" nowrap><div class="right"><strong>GRAND TOTAL&nbsp;</strong></div></td>
									  <td class="right" nowrap><strong><%=GrandTotHCPaid%></strong></td>
									  <td class="right" nowrap><strong><%=GrandTotHCComp%></strong></td>
									  <td class="right" nowrap><strong><%=GrandTotHCNoShow%></strong></td>
									  <td class="right" nowrap><strong><%=GrandTotWeb%></strong></td>
									  <td class="right" nowrap><strong><%=GrandTotHCMem%></strong></td>
									<% if NOT request.form("frmExpReport")="true" then %>
									  <td class="right" nowrap><strong><%=FmtCurrency(GTMemRev)%></strong></td>
									  <td class="right" nowrap><strong><%=FmtCurrency(GTRev)%></strong></td>
									<% else %>
									  <td class="right" nowrap><strong><%=FmtNumber(GTMemRev)%></strong></td>
									  <td class="right" nowrap><strong><%=FmtNumber(GTRev)%></strong></td>
									<% end if %>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
									</table>
										<%
									end if	'eof
									%>
								<%
								elseif showdetails="2" then										
								'*********************  REPORT By INSTRUCTOR  *****************************
									curDate=""
									curTime=""
									curVisitType=""
									curTrainer=""
									curLoc=""
									curTG=""
									%>
									<table class="mainText" width="100%"  border="0" cellpadding="0" cellspacing="0">
									<%
									if NOT rsEntry.EOF then			'EOF
										do while NOT rsEntry.EOF
											if curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new trainer then write the header cells
												if curTrainerID<> "" then	'if this isn't the first record total cells
													if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or curTG<>rsEntry("Typegroup") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") then	'if this is a new class then write the header cells
														if not isnull(curClassname) then
															if curDate<> "" then	'if this isn't the first record total cells
															%>
<% if NOT request.form("frmExpReport")="true" then %>
															<tr height="1">
															<td colspan="1" style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
															<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
															</tr>
<% end if %>
															<tr>
																<td>&nbsp;</td>
																<td nowrap><div class="left"><strong><%=FmtDateTime(curDate)%>&nbsp;<%=FmtTimeShort(curTime)%>&nbsp;<%If isnull(curClassName) then response.write  LEFT(curVisitType,18) else response.write LEFT(curClassName,18) end if%></strong></div></td>
																<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
																<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
																<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
																<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
																<% if NOT request.form("frmExpReport")="true" then %>
																<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
																<% else %>
																<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
																<% end if %>
															</tr>
															<tr>
															  <td colspan="11" nowrap>&nbsp;</td>
															</tr>
															<%
															end if	
														else%>
														<tr>
														  <td colspan="11" nowrap>&nbsp;</td>
														</tr>
													<%
														end if
														tmpHCPaid=0
														tmpHCComp=0
														tmpHCNoShow=0
														tmpWeb=0
														tmpRev=0

													end if
													%>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
													<tr class="right">
														<td>&nbsp;</td>
														<td nowrap class="left"><strong><%=UCASE(Replace(curTrainer,"&nbsp;"," "))%> - HEAD COUNT</strong></td>
														<td nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCPaid%></strong></td>
														<td nowrap><strong>COMP:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCComp%></strong></td>
														<td nowrap><strong>NO SHOW:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCNoShow%></strong></td>
														<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotWeb%></strong></td>
													<% if NOT request.form("frmExpReport")="true" then %>
														<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
													<% else %>
														<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
													<% end if %>
													</tr>
													<tr>
													  <td colspan="11" nowrap>&nbsp;</td>
													</tr>
												<%
												end if
												%>
												<tr class="blackHeader">
												  <td colspan="11" class="left"><strong><%=FmtTrnNameNew(rsEntry, false)%>&nbsp;</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="2">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
												<%
												TotHCPaid=0
												TotHCComp=0
												TotHCNoShow=0
												TotWeb=0
												TotRev=0
											else
												if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or curTG<>rsEntry("Typegroup") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") then	'if this is a new class then write the header cells
													if not isnull(curClassname) then
														if curDate<> "" then	'if this isn't the first record total cells
														%>
<% if NOT request.form("frmExpReport")="true" then %>
														<tr height="1">
															<td colspan="1" style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
															<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
														</tr>
<% end if %>
														<tr>
															<td>&nbsp;</td>
															<td nowrap><div class="left"><strong><%=FmtDateTime(curDate)%>&nbsp;<%=FmtTimeShort(curTime)%>&nbsp;<%If isnull(curClassName) then response.write  LEFT(curVisitType,18) else response.write LEFT(curClassName,18) end if%></strong></div></td>
															<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
															<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
															<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
															<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
														<% if NOT request.form("frmExpReport")="true" then %>
															<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
														<% else %>
															<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
														<% end if %>
														</tr>
														<tr>
														  <td colspan="11" nowrap>&nbsp;</td>
														</tr>
														<%	
														end if
													else%>
														<tr>
														  <td colspan="11" nowrap>&nbsp;</td>
														</tr>
													<%
													end if
													tmpHCPaid=0
													tmpHCComp=0
													tmpHCNoShow=0
													tmpWeb=0
													tmpRev=0
												end if
											end if
											
											if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or curTG<>rsEntry("Typegroup") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") then	'if this is a new class then write the header cells
												%>
												<tr>
												  <td>&nbsp;</td>
												  <td><strong>&nbsp;<%=FmtDateTime(rsEntry("ClassDate"))%></strong></td>
												  <td colspan="2"><strong>&nbsp;<%=FmtTimeShort(rsEntry("ClassTime"))%></strong></td>
												  <td colspan="5"><strong>
												  &nbsp;												  
<% if NOT request.form("frmExpReport")="true" then %>
													<% if NOT isNull(rsEntry("ClassID")) then %>
												  		<a href="adm_cls_list.asp?pDate=<%=DateValue(rsEntry("ClassDate"))%>&pClsID=<%=rsEntry("ClassID")%>">
													<% elseif NOT isNull(rsEntry("ApptVisitRefNo"))then %>
														<a href="adm_appt_e.asp?id=<%=rsEntry("ApptVisitRefNo")%>">
													<% end if %>
<% end if %>
												  		<%If isnull(rsEntry("ClassName")) then response.write rsEntry("TypeName") else response.write rsEntry("ClassName") end if%>
<% if NOT request.form("frmExpReport")="true" then %>
													</a>
<% end if %>
												  
												  </strong></td>
												  <td colspan="2"><strong><%If request.form("optSaleLoc")="0" then%><%=rsEntry("LocationName")%><% end if %>&nbsp;</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
												<tr height="1">
													  <td colspan="1" style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
												</tr>
<% end if %>
												<tr>
												  <td width="5%">&nbsp;</td>
												  <td colspan="2" width="20%" class="left" nowrap><strong><%=session("ClientHW")%> </strong></td>
												  <td colspan="2" width="17%" class="left" nowrap><strong><%=hw61%></strong></td>
												  <td width="10%" class="right" nowrap><strong><%= getHotWord(117)%> </strong></td>
												  <td width="10%" class="center" nowrap><strong>Rem </strong></td>
												  <td width="15%" class="center" nowrap><strong><%= getHotWord(172)%>&nbsp;</strong></td>
												  <td width="10%" class="center" nowrap><strong>No Show&nbsp;</strong></td>
												  <td colspan="2" width="13%" class="right" nowrap><strong>Rev. Per Visit</strong></td>
												</tr>
												<%
												tmpHCPaid=0
												tmpHCComp=0
												tmpHCNoShow=0
												tmpWeb=0
												tmpRev=0
											end if	'new description ended. Continue with next for both new and old descriptions
											curDate=rsEntry("ClassDate")
											curTime=rsEntry("ClassTime")
											curVisitType=rsEntry("TypeName")
											curTrainerID=CLNG(rsEntry("TrainerID"))
											curTrainer=FmtTrnNameNew(rsEntry,false)
											curLoc=rsEntry("LocationName")
											curClassname=rsEntry("Classname")
											curTG=rsEntry("Typegroup")
											
											if rsEntry("NumDeducted")=0 then
												tmpRevPerVisit=0
											else
												If rsEntry("Type")=1 then
													If rsEntry("NumClasses")=0 or isnull(rsEntry("NumClasses")) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=(rsEntry("PaymentAmount")/rsEntry("NumClasses")) * rsEntry("NumDeducted")
													end if
												'elseif rsEntry("Type")=2 OR rsEntry("Type")=3 then
												elseif rsEntry("Type")<>9 then	'CB 10/21/2008
													If (rsEntry("NumClasses")-rsEntry("Remaining"))=0 or isnull((rsEntry("NumClasses")-rsEntry("Remaining"))) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=rsEntry("PaymentAmount")/(rsEntry("NumClasses")-rsEntry("Remaining"))
													end if
												else
													tmpRevPerVisit=0
												end if
											end if
											tmpRev = tmpRev + tmpRevPerVisit
											TotRev = TotRev + tmpRevPerVisit
											GTRev = GTRev + tmpRevPerVisit
											
											If rsEntry("Webscheduler") then
												tmpWeb=tmpWeb+1
												TotWeb=TotWeb+1
												GrandTotWeb=GrandTotWeb+1
											end if

											
											If rsEntry("Value")=1 then
												tmpHCPaid=tmpHCPaid+1
												TotHCPaid=TotHCPaid+1
												GrandTotHCPaid=GrandTotHCPaid+1
											else
												tmpHCComp=tmpHCComp+1
												TotHCComp=TotHCComp+1
												GrandTotHCComp=GrandTotHCComp+1
											end if
											
											If rsEntry("Missed") then
												tmpHCNoShow=tmpHCNoShow+1
												TotHCNoShow=TotHCNoShow+1
												GrandTotHCNoShow=GrandTotHCNoShow+1
											end if
											%>
												<tr class="left">
												  <td>&nbsp;</td>
											<% if NOT request.form("frmExpReport")="true" then %>
												  <td colspan="2" nowrap><a href="main_info.asp?ID=<%=rsEntry("ClientID")%>&fl=true"><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></a></td>
											<% else %>
												  <td colspan="2" nowrap><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></td>
											<% end if %>
												  <td colspan="2" nowrap><%=rsEntry("TypeTaken")%></td>
												  <td class="right" nowrap><%=FmtDateShort(rsEntry("ExpDate"))%></td>
												  <td class="center" nowrap>
<%
											if rsEntry("RealRemaining")=rsEntry("Remaining") then
												if rsEntry("Type") = 9 then
													response.write "<span style=""color:#990000;"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write rsEntry("RealRemaining")
												end if
											else
												if rsEntry("Type") = 9 then
													response.write "<span style=""color:#990000;"" title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write "<span title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")
												end if
											end if
%>												  
												  </td>
												  <%if true=false then %>
												  <td class="center" nowrap><input type="checkbox" name="chkCancelled" <%if rsEntry("Cancelled") then response.write "checked" end if %>></td>
												  <td class="center" nowrap><input type="checkbox" name="chkMissed" <%if rsEntry("Missed") then response.write "checked" end if %>></td>
												  <%else %>
												  <td class="center" nowrap><%if rsEntry("Cancelled") then response.write "Yes" else response.write "No" end if %></td>
												  <td class="center" nowrap><%if rsEntry("Missed") then response.write "Yes" else response.write "No" end if %></td>
												  <%end if%>
											<% if NOT request.form("frmExpReport")="true" then %>
												  <td colspan="2" class="right" nowrap><%=FmtCurrency(tmpRevPerVisit)%></td>
											<% else %>
												  <td colspan="2" class="right" nowrap><%=FmtNumber(tmpRevPerVisit)%></td>
											<% end if %>
												</tr>
											<%		
											rsEntry.MoveNext
										loop
										if not isnull(curClassname) then
										%>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
													<td colspan="1" style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
												  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr>
												  <td>&nbsp;</td>
												<td nowrap><div class="left"><strong><%=FmtDateTime(curDate)%>&nbsp;<%=FmtTimeShort(curTime)%>&nbsp;<%If isnull(curClassName) then response.write  LEFT(curVisitType,18) else response.write LEFT(curClassName,18) end if%></strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
											<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
											<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
											<% end if %>
											</tr>
										<%
										else%>
											<tr>
											  <td colspan="11" nowrap>&nbsp;</td>
											</tr>
										<%end if%>
											<tr>
											  <td colspan="11" nowrap>&nbsp;</td>
											</tr>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
												  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr>
												  <td>&nbsp;</td>
												<td nowrap><div class="left"><strong><%=UCASE(Replace(curTrainer,"&nbsp;"," "))%> - HEAD COUNT</strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotWeb%></strong></td>
											<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
											<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
											<% end if %>
											</tr>
										<%
									end if	'eof
									%>
									<tr>
									  <td colspan="11" nowrap>&nbsp;</td>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
									<tr class="maintextbig">
										<td>&nbsp;</td>
										<td nowrap><div class="left"><strong>HEAD COUNT GRAND TOTAL</strong></div></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCPaid%></strong></td>
										<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCComp%></strong></td>
										<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCNoShow%></strong></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotWeb%></strong></td>
									<% if NOT request.form("frmExpReport")="true" then %>
										<td class="right" nowrap><strong><%=FmtCurrency(GTRev)%></strong></td>
									<% else %>
										<td class="right" nowrap><strong><%=FmtNumber(GTRev)%></strong></td>
									<% end if %>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
						  </table>
								<%
								elseif showdetails="3" then										
								'*********************  REPORT By DATE  *****************************
									curDate=""
									curTime=""
									curVisitType=""
									curTrainer=""
									curLoc=""
									curTG=""
									%>
									<table class="mainText" width="100%"  border="0" cellpadding="0" cellspacing="0">
									<%
									if NOT rsEntry.EOF then			'EOF
										do while NOT rsEntry.EOF
											if curDate<>rsEntry("ClassDate") then	'if this is a new trainer then write the header cells
												if curDate<> "" then	'if this isn't the first record total cells
													if curTime<>rsEntry("ClassTime") or curTG<>rsEntry("Typegroup") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
														if not isnull(curClassname) then
															if curTime<> "" then	'if this isn't the first record total cells
															%>
<% if NOT request.form("frmExpReport")="true" then %>
															<tr height="1">
																<td colspan="1" style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
																<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
															</tr>
<% end if %>
															<tr>
																<td>&nbsp;</td>
																<td nowrap><div class="left"><strong><%=FmtTimeShort(curTime)%>&nbsp;<%=curTrainer%>&nbsp;<%If isnull(curClassName) then response.write LEFT(curVisitType,12) else response.write LEFT(curClassName,12) end if%></strong></div></td>
																<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
																<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
																<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
																<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
															<% if NOT request.form("frmExpReport")="true" then %>
																<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
															<% else %>
																<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
															<% end if %>
															</tr>
															<tr>
															  <td colspan="11" nowrap>&nbsp;</td>
															</tr>
															<%
															end if	
														else%>
															<tr>
															  <td colspan="11" nowrap>&nbsp;</td>
															</tr>
														<%
														end if
														tmpHCPaid=0
														tmpHCComp=0
														tmpHCNoShow=0
														tmpWeb=0
														tmpRev=0
													end if
													%>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
													<tr class="right">
														<td>&nbsp;</td>
														<td nowrap class="left"><strong><%=FmtDateLong(curDate)%> - HEAD COUNT</strong></td>
														<td nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCPaid%></strong></td>
														<td nowrap><strong>COMP:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCComp%></strong></td>
														<td nowrap><strong>NO SHOW:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCNoShow%></strong></td>
														<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotWeb%></strong></td>
													<% if NOT request.form("frmExpReport")="true" then %>
														<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
													<% else %>
														<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
													<% end if %>
													</tr>
													<tr>
													  <td colspan="11" nowrap>&nbsp;</td>
													</tr>
												<%	
												end if
												%>
												<tr>
												  <td colspan="11" class="blackHeader left"><strong><%=FmtDateLong(rsEntry("ClassDate"))%>&nbsp;</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="2">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
												<%
												TotHCPaid=0
												TotHCComp=0
												TotHCNoShow=0
												TotWeb=0
												TotRev=0
											else
												if curTime<>rsEntry("ClassTime") or curTG<>rsEntry("Typegroup") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
													if not isnull(curClassname) then
														if curTime<> "" then	'if this isn't the first record total cells
														%>
<% if NOT request.form("frmExpReport")="true" then %>
														<tr height="1">
														<td colspan="1" style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
														<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
														</tr>
<% end if %>
														<tr>
															<td>&nbsp;</td>
															<td nowrap><div class="left"><strong><%=FmtTimeShort(curTime)%>&nbsp;<%=curTrainer%>&nbsp;<%If isnull(curClassName) then response.write LEFT(curVisitType,12) else response.write LEFT(curClassName, 12) end if%></strong></div></td>
															<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
															<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
															<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
															<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
														<% if NOT request.form("frmExpReport")="true" then %>
															<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
														<% else %>
															<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
														<% end if %>
														</tr>
														<tr>
														  <td colspan="11" nowrap>&nbsp;</td>
														</tr>
														<%
														end if	
													else%>
														<tr>
														  <td colspan="11" nowrap>&nbsp;</td>
														</tr>
													<%
													end if
													tmpHCPaid=0
													tmpHCComp=0
													tmpHCNoShow=0
													tmpWeb=0
													tmpRev=0
												end if
											end if	'new description ended. Continue with next for both new and old descriptions
											if curTime<>rsEntry("ClassTime") or curTG<>rsEntry("Typegroup") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) or curDate<>rsEntry("ClassDate") then	'if this is a new class then write the header cells
												%>
												<tr>
												  <td>&nbsp;</td>
												  <td><strong>&nbsp;<%=FmtTimeShort(rsEntry("ClassTime"))%></strong></td>
												  <td colspan="2"><strong>&nbsp;<%=FmtTrnNameNew(rsEntry,false)%></strong></td>
												  <td colspan="5"><strong>&nbsp;
<% if NOT request.form("frmExpReport")="true" then %>
													<% if NOT isNull(rsEntry("ClassID")) then %>
												  		<a href="adm_cls_list.asp?pDate=<%=DateValue(rsEntry("ClassDate"))%>&pClsID=<%=rsEntry("ClassID")%>">
													<% elseif NOT isNull(rsEntry("ApptVisitRefNo"))then %>
														<a href="adm_appt_e.asp?id=<%=rsEntry("ApptVisitRefNo")%>">
													<% end if %>
<% end if %>
												  		<%If isnull(rsEntry("ClassName")) then response.write rsEntry("TypeName") else response.write rsEntry("ClassName") end if%>
<% if NOT request.form("frmExpReport")="true" then %>
													</a>
<% end if %>
												  </strong></td>
												  <td colspan="2"><strong><%If request.form("optSaleLoc")="0" then%><%=rsEntry("LocationName")%><% end if %>&nbsp;</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
												<tr height="1">
												  	<td colspan="1" style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
												</tr>
<% end if %>
												<tr>
												<td width="5%">&nbsp;</td>
												  <td colspan="2" width="20%" class="left" nowrap><strong><%=session("ClientHW")%> </strong></td>
												  <td colspan="2" width="17%" class="left" nowrap><strong><%=hw61%></strong></td>
												  <td width="10%" class="right" nowrap><strong><%= getHotWord(117)%> </strong></td>
												  <td width="10%" class="center" nowrap><strong>Rem </strong></td>
												  <td width="15%" class="center" nowrap><strong><%= getHotWord(172)%>&nbsp;</strong></td>
												  <td width="10%" class="center" nowrap><strong>No Show&nbsp;</strong></td>
												  <td colspan="2" width="13%" class="right" nowrap><strong>Rev. Per Visit</strong></td>
												</tr>
												<%
												tmpHCPaid=0
												tmpHCComp=0
												tmpHCNoShow=0
												tmpWeb=0
												tmpRev=0
											end if	'new description ended. Continue with next for both new and old descriptions
											curDate=rsEntry("ClassDate")
											curTime=rsEntry("ClassTime")
											curVisitType=rsEntry("TypeName")
											curTrainerID=CLNG(rsEntry("TrainerID"))
											curTrainer=FmtTrnNameNew(rsEntry,false)
											curLoc=rsEntry("LocationName")
											curTG=rsEntry("Typegroup")
											curClassName=rsEntry("ClassName")
											
											if rsEntry("NumDeducted")=0 then
												tmpRevPerVisit=0
											else
												If rsEntry("Type")=1 then
													If rsEntry("NumClasses")=0 or isnull(rsEntry("NumClasses")) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=(rsEntry("PaymentAmount")/rsEntry("NumClasses")) * rsEntry("NumDeducted")
													end if
												'elseif rsEntry("Type")=2 OR rsEntry("Type")=3 then
												elseif rsEntry("Type")<>9 then	'CB 10/21/2008
													If (rsEntry("NumClasses")-rsEntry("Remaining"))=0 or isnull((rsEntry("NumClasses")-rsEntry("Remaining"))) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=rsEntry("PaymentAmount")/(rsEntry("NumClasses")-rsEntry("Remaining"))
													end if
												else
													tmpRevPerVisit=0
												end if
											end if											
											tmpRev = tmpRev + tmpRevPerVisit
											TotRev = TotRev + tmpRevPerVisit
											GTRev = GTRev + tmpRevPerVisit
											
											If rsEntry("Webscheduler") then
												tmpWeb=tmpWeb+1
												TotWeb=TotWeb+1
												GrandTotWeb=GrandTotWeb+1
											end if

											If rsEntry("Value")=1 then
												tmpHCPaid=tmpHCPaid+1
												TotHCPaid=TotHCPaid+1
												GrandTotHCPaid=GrandTotHCPaid+1
											else
												tmpHCComp=tmpHCComp+1
												TotHCComp=TotHCComp+1
												GrandTotHCComp=GrandTotHCComp+1
											end if
											
											If rsEntry("Missed") then
												tmpHCNoShow=tmpHCNoShow+1
												TotHCNoShow=TotHCNoShow+1
												GrandTotHCNoShow=GrandTotHCNoShow+1
											end if
											%>
												<tr class="left">
													<td>&nbsp;</td>
											<% if NOT request.form("frmExpReport")="true" then %>
												  <td colspan="2" nowrap><a href="main_info.asp?ID=<%=rsEntry("ClientID")%>&fl=true"><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></a></td>
											<% else %>
												  <td colspan="2" nowrap><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></td>
											<% end if %>
												  <td colspan="2" nowrap><%=rsEntry("TypeTaken")%></td>
												  <td class="right" nowrap><%=FmtDateShort(rsEntry("ExpDate"))%></td>
												  <td class="center" nowrap>
<%
											if rsEntry("RealRemaining")=rsEntry("Remaining") then
												if rsEntry("Type") = 9 then
													response.write "<span style=""color:#990000;"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write rsEntry("RealRemaining")
												end if
											else
												if rsEntry("Type") = 9 then
													response.write "<span style=""color:#990000;"" title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write "<span title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")
												end if
											end if
%>												  
												  </td>
												  <%if true=false then %>
												  <td class="center" nowrap><input type="checkbox" name="chkCancelled" <%if rsEntry("Cancelled") then response.write "checked" end if %>></td>
												  <td class="center" nowrap><input type="checkbox" name="chkMissed" <%if rsEntry("Missed") then response.write "checked" end if %>></td>
												  <%else %>
												  <td class="center" nowrap><%if rsEntry("Cancelled") then response.write "Yes" else response.write "No" end if %></td>
												  <td class="center" nowrap><%if rsEntry("Missed") then response.write "Yes" else response.write "No" end if %></td>
												  <%end if%>
											<% if NOT request.form("frmExpReport")="true" then %>
												  <td colspan="2" class="right" nowrap><%=FmtCurrency(tmpRevPerVisit)%></td>
											<% else %>
												  <td colspan="2" class="right" nowrap><%=FmtNumber(tmpRevPerVisit)%></td>
											<% end if %>
												</tr>
											<%		
											rsEntry.MoveNext
										loop
										if not isnull(curClassname) then
										%>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
										  	<td colspan="1" style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr>
												<td>&nbsp;</td>
												<td nowrap><div class="left"><strong><%=FmtTimeShort(curTime)%>&nbsp;<%=curTrainer%>&nbsp;<%If isnull(curClassName) then response.write LEFT(curVisitType,12) else response.write LEFT(curClassName, 12) end if%></strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
											<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
											<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
											<% end if %>
											</tr>
											<tr>
											  <td colspan="11" nowrap>&nbsp;</td>
											</tr>
										<%else%>
											<tr>
											  <td colspan="11" nowrap>&nbsp;</td>
											</tr>
										<%end if%>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
												  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr>
												<td>&nbsp;</td>
												<td nowrap><div class="left"><strong><%=FmtDateLong(curDate)%> - HEAD COUNT</strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotWeb%></strong></td>
											<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
											<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
											<% end if %>
											</tr>
										<%
										TotHCPaid=0
										TotHCComp=0
										TotHCNoShow=0
									end if	'eof
									%>
									<tr>
									  <td colspan="11" nowrap>&nbsp;</td>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
									<tr class="maintextbig">
										<td>&nbsp;</td>
										<td nowrap><div class="left"><strong>HEAD COUNT GRAND TOTAL</strong></div></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCPaid%></strong></td>
										<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCComp%></strong></td>
										<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCNoShow%></strong></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotWeb%></strong></td>
									<% if NOT request.form("frmExpReport")="true" then %>
										<td class="right" nowrap><strong><%=FmtCurrency(GTRev)%></strong></td>
									<% else %>
										<td class="right" nowrap><strong><%=FmtNumber(GTRev)%></strong></td>
									<% end if %>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
						  </table>
								<%
								elseif showdetails="4" then										
								'*********************  REPORT By TYPE GROUP  *****************************
									curDate=""
									curTime=""
									curVisitType=""
									curTrainer=""
									curLoc=""
									curTG=""
									%>
									<table class="mainText" width="100%"  border="0" cellpadding="0" cellspacing="0">
									<%
									if NOT rsEntry.EOF then			'EOF
										do while NOT rsEntry.EOF
											if curTG<>rsEntry("Typegroup") then	'if this is a new trainer then write the header cells
												if curTG<> "" then	'if this isn't the first record total cells
													if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or  curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
														if not isnull(curClassname) then
															if curDate<> "" then	'if this isn't the first record total cells
															%>
<% if NOT request.form("frmExpReport")="true" then %>
															<tr height="1">
													 	 		<td style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
																<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
															</tr>
<% end if %>
															<tr>
																<td>&nbsp;</td>
																<td nowrap><div class="left"><strong><%=FmtDateTime(curDate)%>&nbsp;<%=FmtTimeShort(curTime)%>&nbsp;<%=curTrainer%>&nbsp;<%If isnull(curClassName) then response.write LEFT(curVisitType,12) else response.write LEFT(curClassName,12) end if%></strong></div></td>
																<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
																<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
																<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
																<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
															<% if NOT request.form("frmExpReport")="true" then %>
																<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
															<% else %>
																<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
															<% end if %>
															</tr>
															<tr>
															  <td colspan="11" nowrap>&nbsp;</td>
															</tr>
															<%
															end if	
														else%>
															<tr>
															  <td colspan="11" nowrap>&nbsp;</td>
															</tr>
														<%
														end if
														tmpHCPaid=0
														tmpHCComp=0
														tmpHCNoShow=0
														tmpWeb=0
														tmpRev=0
													end if
													%>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
													<tr class="right">
														<td>&nbsp;</td>
														<td class="left" nowrap><strong><%=UCASE(curTG)%> - HEAD COUNT</strong></td>
														<td nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCPaid%></strong></td>
														<td nowrap><strong>COMP:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCComp%></strong></td>
														<td nowrap><strong>NO SHOW:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCNoShow%></strong></td>
														<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotWeb%></strong></td>
													<% if NOT request.form("frmExpReport")="true" then %>
														<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
													<% else %>
														<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
													<% end if %>
													</tr>
													<tr>
													  <td colspan="11" nowrap>&nbsp;</td>
													</tr>
												<%	
												end if
												%>
												<tr>
												  <td colspan="11" class="blackHeader left"><strong><%=UCASE(rsEntry("Typegroup"))%>&nbsp;</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="2">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
												<%
												TotHCPaid=0
												TotHCComp=0
												TotHCNoShow=0
												TotWeb=0
												TotRev=0
											else
												if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or  curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
													if not isnull(curClassname) then
														if curDate<> "" then	'if this isn't the first record total cells
														%>
<% if NOT request.form("frmExpReport")="true" then %>
														<tr height="1">
														<td style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
														<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
														</tr>
<% end if %>
														<tr>
															<td>&nbsp;</td>
															<td nowrap><div class="left"><strong><%=FmtDateTime(curDate)%>&nbsp;<%=FmtTimeShort(curTime)%>&nbsp;<%=curTrainer%>&nbsp;<%If isnull(curClassName) then response.write LEFT(curVisitType,12) else response.write LEFT(curClassName,12) end if%></strong></div></td>
															<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
															<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
															<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
															<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
														<% if NOT request.form("frmExpReport")="true" then %>
															<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
														<% else %>
															<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
														<% end if %>
														</tr>
														<tr>
														  <td colspan="11" nowrap>&nbsp;</td>
														</tr>
														<%
														end if	
													else%>
														<tr>
														  <td colspan="11" nowrap>&nbsp;</td>
														</tr>
													<%
													end if
													tmpHCPaid=0
													tmpHCComp=0
													tmpHCNoShow=0
													tmpWeb=0
													tmpRev=0
												end if
											end if	'new description ended. Continue with next for both new and old descriptions
											if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
												%>
												<tr>
												  <td>&nbsp;</td>
												  <td><strong>&nbsp;<%=FmtDateTime(rsEntry("ClassDate"))%></strong></td>
												  <td><strong>&nbsp;<%=FmtTimeShort(rsEntry("ClassTime"))%></strong></td>
												  <td colspan="2"><strong>&nbsp;<%=FmtTrnNameNew(rsEntry,false)%></strong></td>
												  <td colspan="3"><strong>&nbsp;
<% if NOT request.form("frmExpReport")="true" then %>
													<% if NOT isNull(rsEntry("ClassID")) then %>
												  		<a href="adm_cls_list.asp?pDate=<%=DateValue(rsEntry("ClassDate"))%>&pClsID=<%=rsEntry("ClassID")%>">
													<% elseif NOT isNull(rsEntry("ApptVisitRefNo"))then %>
														<a href="adm_appt_e.asp?id=<%=rsEntry("ApptVisitRefNo")%>">
													<% end if %>
<% end if %>
												  		<%If isnull(rsEntry("ClassName")) then response.write rsEntry("TypeName") else response.write rsEntry("ClassName") end if%>
<% if NOT request.form("frmExpReport")="true" then %>
													</a>
<% end if %>
												  </strong></td>
												  <td colspan="2"><strong><%If request.form("optSaleLoc")="0" then%><%=rsEntry("LocationName")%><% end if %>&nbsp;</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
												<tr height="1">
													  <td style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
												</tr>
<% end if %>
												<tr>
												  <td width="5%">&nbsp;</td>
												  <td colspan="2" width="20%" class="left" nowrap><strong><%=session("ClientHW")%> </strong></td>
												  <td colspan="2" width="17%" class="left" nowrap><strong><%=hw61%></strong></td>
												  <td width="10%" class="right" nowrap><strong><%= getHotWord(117)%> </strong></td>
												  <td width="10%" class="center" nowrap><strong>Rem </strong></td>
												  <td width="15%" class="center" nowrap><strong><%= getHotWord(172)%>&nbsp;</strong></td>
												  <td width="10%" class="center" nowrap><strong>No Show&nbsp;</strong></td>
												  <td colspan="2" width="13%" class="right" nowrap><strong>Rev. Per Visit</strong></td>
												</tr>
												<%
												tmpHCPaid=0
												tmpHCComp=0
												tmpHCNoShow=0
												tmpWeb=0
												tmpRev=0
											end if	'new description ended. Continue with next for both new and old descriptions
											curDate=rsEntry("ClassDate")
											curTime=rsEntry("ClassTime")
											curVisitType=rsEntry("TypeName")
											curTrainerID=CLNG(rsEntry("TrainerID"))
											curTrainer=FmtTrnNameNew(rsEntry,false)
											curLoc=rsEntry("LocationName")
											curTG=rsEntry("Typegroup")
											curClassName=rsEntry("ClassName")
											
											if rsEntry("NumDeducted")=0 then
												tmpRevPerVisit=0
											else
												If rsEntry("Type")=1 then
													If rsEntry("NumClasses")=0 or isnull(rsEntry("NumClasses")) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=(rsEntry("PaymentAmount")/rsEntry("NumClasses")) * rsEntry("NumDeducted")
													end if
												'elseif rsEntry("Type")=2 OR rsEntry("Type")=3 then
												elseif rsEntry("Type")<>9 then	'CB 10/21/2008
													If (rsEntry("NumClasses")-rsEntry("Remaining"))=0 or isnull((rsEntry("NumClasses")-rsEntry("Remaining"))) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=rsEntry("PaymentAmount")/(rsEntry("NumClasses")-rsEntry("Remaining"))
													end if
												else
													tmpRevPerVisit=0
												end if
											end if
											
											tmpRev = tmpRev + tmpRevPerVisit
											TotRev = TotRev + tmpRevPerVisit
											GTRev = GTRev + tmpRevPerVisit
											
											If rsEntry("Webscheduler") then
												tmpWeb=tmpWeb+1
												TotWeb=TotWeb+1
												GrandTotWeb=GrandTotWeb+1
											end if

											If rsEntry("Value")=1 then
												tmpHCPaid=tmpHCPaid+1
												TotHCPaid=TotHCPaid+1
												GrandTotHCPaid=GrandTotHCPaid+1
											else
												tmpHCComp=tmpHCComp+1
												TotHCComp=TotHCComp+1
												GrandTotHCComp=GrandTotHCComp+1
											end if
											
											If rsEntry("Missed") then
												tmpHCNoShow=tmpHCNoShow+1
												TotHCNoShow=TotHCNoShow+1
												GrandTotHCNoShow=GrandTotHCNoShow+1
											end if
											%>
												<tr class="left">
												  <td>&nbsp;</td>
												<% if NOT request.form("frmExpReport")="true" then %>
												  <td colspan="2" nowrap><a href="main_info.asp?ID=<%=rsEntry("ClientID")%>&fl=true"><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></a></td>
												<% else %>
												  <td colspan="2" nowrap><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></td>
												<% end if %>
												  <td colspan="2" nowrap><%=rsEntry("TypeTaken")%></td>
												  <td class="right" nowrap><%=FmtDateShort(rsEntry("ExpDate"))%></td>
												  <td class="center" nowrap>
<%
											if rsEntry("RealRemaining")=rsEntry("Remaining") then
												if rsEntry("Type") = 9 then
													response.write "<span style=""color:#990000;"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write rsEntry("RealRemaining")
												end if
											else
												if rsEntry("Type") = 9 then
													response.write "<span style=""color:#990000;"" title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write "<span title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")
												end if
											end if
%>												  
												  </td>
												  <%if true=false then %>
												  <td class="center" nowrap><input type="checkbox" name="chkCancelled" <%if rsEntry("Cancelled") then response.write "checked" end if %>></td>
												  <td class="center" nowrap><input type="checkbox" name="chkMissed" <%if rsEntry("Missed") then response.write "checked" end if %>></td>
												  <%else %>
												  <td class="center" nowrap><%if rsEntry("Cancelled") then response.write "Yes" else response.write "No" end if %></td>
												  <td class="center" nowrap><%if rsEntry("Missed") then response.write "Yes" else response.write "No" end if %></td>
												  <%end if%>
												<% if NOT request.form("frmExpReport")="true" then %>
												  <td colspan="2" class="right" nowrap><%=FmtCurrency(tmpRevPerVisit)%></td>
												<% else %>
												  <td colspan="2" class="right" nowrap><%=FmtNumber(tmpRevPerVisit)%></td>
												<% end if %>
												</tr>
											<%		
											rsEntry.MoveNext
										loop
										if not isnull(curClassname) then
										%>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
											<td style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr>
												<td>&nbsp;</td>
												<td nowrap><div class="left"><strong><%=FmtDateTime(curDate)%>&nbsp;<%=FmtTimeShort(curTime)%>&nbsp;<%=curTrainer%>&nbsp;<%If isnull(curClassName) then response.write LEFT(curVisitType,12) else response.write LEFT(curClassName,12) end if%></strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
											<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
											<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
											<% end if %>
											</tr>
											<tr>
											  <td colspan="11" nowrap>&nbsp;</td>
											</tr>
										<%else%>
											<tr>
											  <td colspan="11" nowrap>&nbsp;</td>
											</tr>
										<%end if%>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
												  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr>
												<td>&nbsp;</td>
												<td nowrap><div class="left"><strong><%=UCASE(curTG)%> - HEAD COUNT</strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=totHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=totHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=totHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotWeb%></strong></td>
											<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
											<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
											<% end if %>
											</tr>
										<%
									end if	'eof
									%>
									<tr>
									  <td colspan="11" nowrap>&nbsp;</td>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
									<tr class="maintextbig">
										<td>&nbsp;</td>
										<td nowrap><div class="left"><strong>HEAD COUNT GRAND TOTAL</strong></div></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCPaid%></strong></td>
										<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCComp%></strong></td>
										<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCNoShow%></strong></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotWeb%></strong></td>
									<% if NOT request.form("frmExpReport")="true" then %>
										<td class="right" nowrap><strong><%=FmtCurrency(GTRev)%></strong></td>
									<% else %>
										<td class="right" nowrap><strong><%=(GTRev)%></strong></td>
									<% end if %>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
						  </table>
								<%
								elseif showdetails="5" then										
									curDate=""
									curTime=""
									curVisitType=""
									curTrainer=""
									curLoc=""
									curTG=""
								'*********************  REPORT By VISIT TYPE  *****************************
									%>
									<table class="mainText" width="100%"  border="0" cellpadding="0" cellspacing="0">
									<%
									if NOT rsEntry.EOF then			'EOF
										do while NOT rsEntry.EOF
											if curVisitType<>rsEntry("TypeName") then	'if this is a new trainer then write the header cells
												if curVisitType<> "" then	'if this isn't the first record total cells
													if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or  curTG<>rsEntry("TypeGroup") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
														if not isnull(curClassname) then
															if curDate<> "" then	'if this isn't the first record total cells
															%>
<% if NOT request.form("frmExpReport")="true" then %>
															<tr height="1">
														  		<td style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
																<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
															</tr>
<% end if %>
															<tr>
																<td>&nbsp;</td>
																<td nowrap><div class="left"><strong><%=FmtDateTime(curDate)%>&nbsp;<%=FmtTimeShort(curTime)%>&nbsp;<%=curTG%>&nbsp;<%=curTrainer%></strong></div></td>
																<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
																<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
																<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
																<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
																<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
															<% if NOT request.form("frmExpReport")="true" then %>
																<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
															<% else %>
																<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
															<% end if %>
															</tr>
															<tr>
															  <td colspan="11" nowrap>&nbsp;</td>
															</tr>
															<%
															end if	
														else%>
															<tr>
															  <td colspan="11" nowrap>&nbsp;</td>
															</tr>
														<%
														end if
														tmpHCPaid=0
														tmpHCComp=0
														tmpHCNoShow=0
														tmpWeb=0
														tmpRev=0
													end if
													%>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
													<tr class="right">
														<td>&nbsp;</td>
														<td class="left" nowrap><strong><%=UCASE(curVisitType)%> - HEAD COUNT</strong></td>
														<td nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCPaid%></strong></td>
														<td nowrap><strong>COMP:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCComp%></strong></td>
														<td nowrap><strong>NO SHOW:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotHCNoShow%></strong></td>
														<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
														<td class="left" nowrap><strong><%=TotWeb%></strong></td>
													<% if NOT request.form("frmExpReport")="true" then %>
														<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
													<% else %>
														<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
													<% end if %>
													</tr>
													<tr>
													  <td colspan="11" nowrap>&nbsp;</td>
													</tr>
												<%	
												end if
												%>
												<tr>
												  <td colspan="11" class="blackHeader left"><strong><%=UCASE(rsEntry("TypeName"))%>&nbsp;</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="2">
														  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
												<%
												TotHCPaid=0
												TotHCComp=0
												TotHCNoShow=0
												TotWeb=0
												TotRev=0
											else
												if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or  curTG<>rsEntry("TypeGroup") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
													if not isnull(curClassname) then
														if curDate<> "" then	'if this isn't the first record total cells
														%>
<% if NOT request.form("frmExpReport")="true" then %>
														<tr height="1">
															<td style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
															<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
														</tr>
<% end if %>
														<tr>
															<td>&nbsp;</td>
															<td nowrap><div class="left"><strong><%=FmtDateTime(curDate)%>&nbsp;<%=FmtTimeShort(curTime)%>&nbsp;<%=curTG%>&nbsp;<%=curTrainer%></strong></div></td>
															<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
															<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
															<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
															<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
															<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
														<% if NOT request.form("frmExpReport")="true" then %>
															<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
														<% else %>
															<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
														<% end if %>
														</tr>
														<tr>
														  <td colspan="11" nowrap>&nbsp;</td>
														</tr>
														<%
														end if	
													else%>
														<tr>
														  <td colspan="11" nowrap>&nbsp;</td>
														</tr>
													<%
													end if
													tmpHCPaid=0
													tmpHCComp=0
													tmpHCNoShow=0
													tmpWeb=0
													tmpRev=0
												end if
											end if	'new description ended. Continue with next for both new and old descriptions
											if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or curVisitType<>rsEntry("TypeName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
												%>
												<tr>
												  <td>&nbsp;</td>
												  <td><strong>&nbsp;<%=FmtDateTime(rsEntry("ClassDate"))%></strong></td>
												  <td><strong>&nbsp;<%=FmtTimeShort(rsEntry("ClassTime"))%></strong></td>
												  <td colspan="3"><strong>&nbsp;
<% if NOT request.form("frmExpReport")="true" then %>
													<% if NOT isNull(rsEntry("ClassID")) then %>
												  		<a href="adm_cls_list.asp?pDate=<%=DateValue(rsEntry("ClassDate"))%>&pClsID=<%=rsEntry("ClassID")%>">
													<% elseif NOT isNull(rsEntry("ApptVisitRefNo"))then %>
														<a href="adm_appt_e.asp?id=<%=rsEntry("ApptVisitRefNo")%>">
													<% end if %>
<% end if %>
												  		<%=rsEntry("TypeGroup")%>
<% if NOT request.form("frmExpReport")="true" then %>
													</a>
<% end if %>
												  </strong></td>
												  <td colspan="2"><strong>&nbsp;<%=FmtTrnNameNew(rsEntry,false)%></strong></td>
												  <td colspan="2"><strong><%If request.form("optSaleLoc")="0" then%><%=rsEntry("LocationName")%><% end if %>&nbsp;</strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true" then %>
												<tr height="1">
													  <td style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
												</tr>
<% end if %>
												<tr>
												  <td width="5%">&nbsp;</td>
												  <td colspan="2" width="20%" class="left" nowrap><strong><%=session("ClientHW")%> </strong></td>
												  <td colspan="2" width="17%" class="left" nowrap><strong><%=hw61%></strong></td>
												  <td width="10%" class="right" nowrap><strong><%= getHotWord(117)%> </strong></td>
												  <td width="10%" class="center" nowrap><strong>Rem </strong></td>
												  <td width="15%" class="center" nowrap><strong><%= getHotWord(172)%>&nbsp;</strong></td>
												  <td width="10%" class="center" nowrap><strong>No Show&nbsp;</strong></td>
												  <td colspan="2" width="13%" class="right" nowrap><strong>Rev. Per Visit</strong></td>
												</tr>
												<%
												tmpHCPaid=0
												tmpHCComp=0
												tmpHCNoShow=0
												tmpWeb=0
												tmpRev=0
											end if	'new description ended. Continue with next for both new and old descriptions
											curDate=rsEntry("ClassDate")
											curTime=rsEntry("ClassTime")
											curVisitType=rsEntry("TypeName")
											curTrainerID=CLNG(rsEntry("TrainerID"))
											curTrainer=FmtTrnNameNew(rsEntry,false)
											curLoc=rsEntry("LocationName")
											curTG=rsEntry("Typegroup")
											curClassName=rsEntry("ClassName")
											
											if rsEntry("NumDeducted")=0 then
												tmpRevPerVisit=0
											else
												If rsEntry("Type")=1 then
													If rsEntry("NumClasses")=0 or isnull(rsEntry("NumClasses")) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=(rsEntry("PaymentAmount")/rsEntry("NumClasses")) * rsEntry("NumDeducted")
													end if
												'elseif rsEntry("Type")=2 OR rsEntry("Type")=3 then
												elseif rsEntry("Type")<>9 then	'CB 10/21/2008
													If (rsEntry("NumClasses")-rsEntry("Remaining"))=0 or isnull((rsEntry("NumClasses")-rsEntry("Remaining"))) then
														tmpRevPerVisit=rsEntry("PaymentAmount")
													else
														tmpRevPerVisit=rsEntry("PaymentAmount")/(rsEntry("NumClasses")-rsEntry("Remaining"))
													end if
												else
													tmpRevPerVisit=0
												end if
											end if
											
											tmpRev = tmpRev + tmpRevPerVisit
											TotRev = TotRev + tmpRevPerVisit
											GTRev = GTRev + tmpRevPerVisit
											
											If rsEntry("Webscheduler") then
												tmpWeb=tmpWeb+1
												TotWeb=TotWeb+1
												GrandTotWeb=GrandTotWeb+1
											end if

											If rsEntry("Value")=1 then
												tmpHCPaid=tmpHCPaid+1
												TotHCPaid=TotHCPaid+1
												GrandTotHCPaid=GrandTotHCPaid+1
											else
												tmpHCComp=tmpHCComp+1
												TotHCComp=TotHCComp+1
												GrandTotHCComp=GrandTotHCComp+1
											end if
											
											If rsEntry("Missed") then
												tmpHCNoShow=tmpHCNoShow+1
												TotHCNoShow=TotHCNoShow+1
												GrandTotHCNoShow=GrandTotHCNoShow+1
											end if
											%>
												<tr class="left">
												  <td>&nbsp;</td>
											<% if NOT request.form("frmExpReport")="true" then %>
												  <td colspan="2" nowrap><a href="main_info.asp?ID=<%=rsEntry("ClientID")%>&fl=true"><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></a></td>
											<% else %>
												  <td colspan="2" nowrap><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></td>
											<% end if %>

												  <td colspan="2" nowrap><%=rsEntry("TypeTaken")%></td>
												  <td class="right" nowrap><%=FmtDateShort(rsEntry("ExpDate"))%></td>
												  <td class="center" nowrap>
<%
											if rsEntry("RealRemaining")=rsEntry("Remaining") then
												if rsEntry("Remaining") < 0 then
													response.write "<span style=""color:#990000;"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write rsEntry("RealRemaining")
												end if
											else
												if rsEntry("Remaining") < 0 then
													response.write "<span style=""color:#990000;"" title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")*-1 & "&nbsp;owed</b></span>"
												else
													response.write "<span title=""" & rsEntry("RealRemaining")-rsEntry("Remaining") & " additional scheduled"">" & rsEntry("RealRemaining")
												end if
											end if
%>												  
												  </td>
												  <%if true=false then %>
												  <td class="center" nowrap><input type="checkbox" name="chkCancelled" <%if rsEntry("Cancelled") then response.write "checked" end if %>></td>
												  <td class="center" nowrap><input type="checkbox" name="chkMissed" <%if rsEntry("Missed") then response.write "checked" end if %>></td>
												  <%else %>
												  <td class="center" nowrap><%if rsEntry("Cancelled") then response.write "Yes" else response.write "No" end if %></td>
												  <td class="center" nowrap><%if rsEntry("Missed") then response.write "Yes" else response.write "No" end if %></td>
												  <%end if%>
											<% if NOT request.form("frmExpReport")="true" then %>
												  <td colspan="2" class="right" nowrap><%=FmtCurrency(tmpRevPerVisit)%></td>
											<% else %>
												  <td colspan="2" class="right" nowrap><%=FmtNumber(tmpRevPerVisit)%></td>
											<% end if %>
												</tr>
											<%		
											rsEntry.MoveNext
										loop
										if not isnull(curClassname) then
										%>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
											<td style="background-color:#FFFFFF;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											<td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr>
												<td>&nbsp;</td>
												<td nowrap><div class="left"><strong><%=FmtDateTime(curDate)%>&nbsp;<%=FmtTimeShort(curTime)%>&nbsp;<%=curTG%>&nbsp;<%=curTrainer%></strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=tmpWeb%></strong></td>
											<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(tmpRev)%></strong></td>
											<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(tmpRev)%></strong></td>
											<% end if %>
											</tr>
											<tr>
											  <td colspan="11" nowrap>&nbsp;</td>
											</tr>
										<%else%>
											<tr>
											  <td colspan="11" nowrap>&nbsp;</td>
											</tr>
										<%end if%>
<% if NOT request.form("frmExpReport")="true" then %>
											<tr height="1">
												  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											</tr>
<% end if %>
											<tr>
												<td>&nbsp;</td>
												<td nowrap><div class="left"><strong><%=UCASE(curVisitType)%> - HEAD COUNT</strong></div></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=totHCPaid%></strong></td>
												<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=totHCComp%></strong></td>
												<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=totHCNoShow%></strong></td>
												<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
												<td class="left" nowrap><strong><%=TotWeb%></strong></td>
											<% if NOT request.form("frmExpReport")="true" then %>
												<td class="right" nowrap><strong><%=FmtCurrency(TotRev)%></strong></td>
											<% else %>
												<td class="right" nowrap><strong><%=FmtNumber(TotRev)%></strong></td>
											<% end if %>
											</tr>
										<%
									end if	'eof
									%>
									<tr>
									  <td colspan="11" nowrap>&nbsp;</td>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
									<tr class="maintextbig">
										<td>&nbsp;</td>
										<td nowrap><div class="left"><strong>HEAD COUNT GRAND TOTAL</strong></div></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(36))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCPaid%></strong></td>
										<td class="right" nowrap><strong>COMP:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCComp%></strong></td>
										<td class="right" nowrap><strong>NO SHOW:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotHCNoShow%></strong></td>
										<td class="right" nowrap><strong><%= UCASE(getHotWord(54))%>:&nbsp;</strong></td>
										<td class="left" nowrap><strong><%=GrandTotWeb%></strong></td>
									<% if NOT request.form("frmExpReport")="true" then %>
										<td class="right" nowrap><strong><%=FmtCurrency(GTRev)%></strong></td>
									<% else %>
										<td class="right" nowrap><strong><%=FmtNumber(GTRev)%></strong></td>
									<% end if %>
									</tr>
<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
<% end if %>
						  </table>
								<%
								elseif showdetails="6" then										
								'*********************  ROLL SHEET  *****************************
									%>
									<table class="mainText" width="650"  border="0" cellpadding="0" cellspacing="0">
									<%
									if NOT rsEntry.EOF then			'EOF
									%>
										<tr>
										<td colspan="11" class="mainTextBig" nowrap><b><%=WeekDayName(WeekDay(cSDate))%> - <%=MonthName(Month(cSDate))%>&nbsp;<%=Day(cSDate)%>,&nbsp;<%=Year(cSDate)%></b></td>
										</tr>
<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="2">
											  <td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
										</tr>
<% end if %>
										<tr>
										  <%If cSDate<>cEDate then%><td width="60" class="left" nowrap><strong><%= getHotWord(57)%> </strong></td><%end if%>
										  <td width="50" class="left" nowrap><strong><%= getHotWord(58)%> </strong></td>
										  <td width="100" class="Left" nowrap><strong>Class Name </strong></td>
										  <%if session("useResrcResv") then%>
										  <td width="60" class="Left" nowrap><strong><%=hw0%></strong></td>
										  <% end if %>
										  <td width="80" class="Left" nowrap><strong><%=hw6%></strong></td>
										  <%If ss_UseAsst1 then%><td width="80" class="left" nowrap><strong><%=hw13%> </strong></td><%end if%>
										  <%If ss_UseAsst2 then%><td width="80" class="left" nowrap><strong><%=hw15%> </strong></td><%end if%>
										  <td width="50" class="center" nowrap><strong>#Enrolled</strong></td>
										  <td width="50" class="center" nowrap><strong>#<%= getHotWord(36)%></strong></td>
										  <td width="50" class="center" nowrap><strong>#Comps&nbsp;</strong></td>
										  <td class="center" nowrap><strong>Head Count</strong></td>
										</tr>
<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="1">
											<td colspan=<%If (ss_UseAsst1 and ss_UseAsst2) then response.write "11" else if (ss_UseAsst1 or ss_UseAsst2) then response.write "10" else response.write "9" end if%> style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
										</tr>
<% end if %>
										<%
										do while NOT rsEntry.EOF
											if curDate<>rsEntry("ClassDate") or curTime<>rsEntry("ClassTime") or curClassName<>rsEntry("ClassName") or curLoc<>rsEntry("LocationName") or curTrainerID<>CLNG(rsEntry("TrainerID")) then	'if this is a new class then write the header cells
												if curTime<> "" then	'if this isn't the first record total cells
												%>
													<tr height="30">
													  <%If cSDate<>cEDate then%><td class="left" nowrap><%=FmtDateShort(curDate)%></td><%end if%>
													  <td class="left" nowrap><%=FmtTimeShort(curTime)%></td>
													  <td class="left" nowrap><%=curClassName%>&nbsp;</td>
												  <%if session("useResrcResv") then%>
												  <td class="Left" nowrap><%=curRsrcName%>&nbsp;</td>
												  <% end if %>
													  <td class="left" nowrap><%=curTrainer%></td>
													  <%If ss_UseAsst1 then%><td class="left" nowrap><%If isnull(curAsst1) then Response.write "" else response.write FmtTrnName(curAsst1) end if%></td><%end if%>
													  <%If ss_UseAsst2 then%><td class="left" nowrap><%If isnull(curAsst2) then Response.write "" else response.write FmtTrnName(curAsst2) end if%></td><%end if%>
													  <td class="center" nowrap></td>
													  <td class="center" nowrap><%=tmpHCPaid%></td>
													  <td class="center" nowrap><%=tmpHCComp%></td>
													  <td class="center" nowrap></td>
													</tr>
													<tr height="30">
													  <%If cSDate<>cEDate then%><td class="left" nowrap>&nbsp;</td><%end if%>
													  <td class="left" nowrap>&nbsp;</td>
													  <td class="left" nowrap>&nbsp;</td>
													  <%if session("useResrcResv") then%>
													  <td class="left" nowrap>&nbsp;</td>
													  <% end if %>
													  <td class="left" nowrap>_____________________________&nbsp;&nbsp;</td>
													  <%If ss_UseAsst1 then%><td class="left" nowrap>_____________________________&nbsp;&nbsp;</td><%end if%>
													  <%If ss_UseAsst2 then%><td class="left" nowrap>_____________________________</td><%end if%>
													  <td class="center" nowrap>________</td>
													  <td class="center" nowrap>&nbsp;</td>
													  <td class="center" nowrap>&nbsp;</td>
													  <td class="center" nowrap>________</td>
													</tr>
<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														<td colspan=<%If (ss_UseAsst1 and ss_UseAsst2) then response.write "11" else if (ss_UseAsst1 or ss_UseAsst2) then response.write "10" else response.write "9" end if%> style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
													</tr>
<% end if %>
												<%
												end if	'end of curtime<>""
												tmpHCPaid=0
												tmpHCComp=0
												tmpWeb=0
											end if	'end of curTime<>rsentry("classtime")...
											curDate=rsEntry("ClassDate")
											curTime=rsEntry("ClassTime")
											curVisitType=rsEntry("TypeName")
											curTrainerID=CLNG(rsEntry("TrainerID"))
											curTrainer=FmtTrnNameNew(rsEntry,false)
											curLoc=rsEntry("LocationName")
											curTG=rsEntry("Typegroup")
											curClassName=rsEntry("ClassName")
											if not isNULL(rsEntry("ResourceName")) then
												curRsrcName = rsEntry("ResourceName")
											else
												curRsrcName = "not assigned"
											end if
											If not isnull(rsEntry("TrainerID2")) then
												curAsst1=CLng(rsEntry("TrainerID2"))
											else
												curAsst1=Null
											end if
											If not isnull(rsEntry("TrainerID3")) then
												curAsst2=Clng(rsEntry("TrainerID3"))
											else
												curAsst2=Null
											end if
											
											
											If rsEntry("Webscheduler") then
												tmpWeb=tmpWeb+1
												TotWeb=TotWeb+1
												GrandTotWeb=GrandTotWeb+1
											end if

											If rsEntry("Value")=1 then
												tmpHCPaid=tmpHCPaid+1
												TotHCPaid=TotHCPaid+1
												GrandTotHCPaid=GrandTotHCPaid+1
											else
												tmpHCComp=tmpHCComp+1
												TotHCComp=TotHCComp+1
												GrandTotHCComp=GrandTotHCComp+1
											end if
											rsEntry.MoveNext
										loop
										%>
										<tr height="30">
										  <%If cSDate<>cEDate then%><td class="left" nowrap><%=FmtDateShort(curDate)%></td><%end if%>
										  <td class="left" nowrap><%=FmtTimeShort(curTime)%></td>
										  <td class="left" nowrap><%=curClassName%>&nbsp;</td>
										  <%if session("useResrcResv") then%>
										  <td class="Left" nowrap><%=curRsrcName%>&nbsp;</td>
										  <% end if %>
										  <td class="left" nowrap><%=curTrainer%></td>
										  <%If ss_UseAsst1 then%><td class="left" nowrap><%If isnull(curAsst1) then Response.write "" else response.write FmtTrnName(curAsst1) end if%></td><%end if%>
										  <%If ss_UseAsst2 then%><td class="left" nowrap><%If isnull(curAsst2) then Response.write "" else response.write FmtTrnName(curAsst2) end if%></td><%end if%>
										  <td class="center" nowrap></td>
										  <td class="center" nowrap><%=tmpHCPaid%></td>
										  <td class="center" nowrap><%=tmpHCComp%></td>
										  <td class="center" nowrap></td>
										</tr>
										<tr height="30">
										  <%If cSDate<>cEDate then%><td class="left" nowrap>&nbsp;</td><%end if%>
										  <td class="left" nowrap>&nbsp;</td>
										  <td class="left" nowrap>&nbsp;</td>
										  <%if session("useResrcResv") then%>
										  <td class="left" nowrap>&nbsp;</td>
										  <% end if %>
										  <td class="left" nowrap>_____________________________&nbsp;&nbsp;</td>
										  <%If ss_UseAsst1 then%><td class="left" nowrap>_____________________________&nbsp;&nbsp;</td><%end if%>
										  <%If ss_UseAsst2 then%><td class="left" nowrap>_____________________________</td><%end if%>
										  <td class="center" nowrap>________</td>
										  <td class="center" nowrap>&nbsp;</td>
										  <td class="center" nowrap>&nbsp;</td>
										  <td class="center" nowrap>________</td>
										</tr>
<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="2">
											<td colspan=<%If (ss_UseAsst1 and ss_UseAsst2) then response.write "11" else if (ss_UseAsst1 or ss_UseAsst2) then response.write "10" else response.write "9" end if%> style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
										</tr>
<% end if %>
										<tr>
										  <td colspan="11" nowrap>&nbsp;</td>
										</tr>
									</table>
										<%
									end if	'eof
								end if		' end of showdetails=true if statement
								rsEntry.close
								set rsEntry = nothing
							end if		'end of generate report if statement
							%>
						  </td>
							</tr>
						</table>
						</td>
						</tr></table>
				</td>
				</tr>
				</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%'end of frmexport value check
	
	
end if
%>
