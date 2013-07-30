<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
%>
		<!-- #include file="inc_accpriv.asp" -->
		<!-- #include file="inc_rpt_tagging.asp" -->
<%
if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EVENT_INVOICE") then 
%>
<script type="text/javascript">
	alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
	javascript:history.go(-1);
</script>
<%
else
%>

<!-- #include file="../inc_i18n.asp" -->
<!-- #include file="inc_row_colors.asp" -->
<!-- #include file="inc_help_content.asp" -->
<!-- #include file="inc_hotword.asp" -->
<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

if request.form("requiredtxtDateStart")<>"" then
	Call SetLocale(session("mvarLocaleStr"))
		pSDate = CDATE(request.form("requiredtxtDateStart"))
	Call SetLocale("en-us")
else
	pSDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
end if

if request.form("requiredtxtDateEnd")<>"" then
	Call SetLocale(session("mvarLocaleStr"))
		pEDate = CDATE(request.form("requiredtxtDateEnd"))
	Call SetLocale("en-us")
else
	pEDate = DateAdd("m", 3, DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
end if


dim rsTimes, rsTrainers, strSQLtime, strSQLtrainers, fldName, selectedClassDescription, cont
dim disDate, disTime, disType, disTrn, disClt, cLoc, ap_view_all_locs, strSQLSumTag

ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
		
dim category : category = ""
if (RQ("category"))<>"" then
	category = RQ("category")
elseif (RF("category"))<>"" then
	category = RF("category") 
end if

if request.form("optRptLoc")<>"" then
	cLoc = CINT(request.form("optRptLoc"))
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

set rsEvent = Server.CreateObject("ADODB.Recordset")
set rsEvent2 = Server.CreateObject("ADODB.Recordset")

if NOT request.form("frmExpReport")="true" then
%>

<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->

<!-- American/Canada=2 format mm/dd/yyyy -->
<!-- European/Rest of the world=1 format dd-mm-yyyy -->
			
<%= js(array("mb", "calendar" & dateFormatCode, "adm/adm_rpt_event_invoice", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %> 
<script type="text/javascript">
function exportReport() {
	document.frmEvtInvoice.frmExpReport.value = "true";
	document.frmEvtInvoice.displayResults.value = "true";
	//document.frmEvtInvoice.frmTagClients.value = "false";
	<% iframeSubmit "frmEvtInvoice", "adm_rpt_event_invoice.asp" %>
}
</script>

	
	<!-- #include file="../inc_date_ctrl.asp" -->
	<!-- #include file="../inc_ajax.asp" -->
	<!-- #include file="../inc_val_date.asp" -->



<% pageStart %>
	<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div class="headText breadcrumbs-old">
			<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
			<%if category<>"" then%>
			<span class="breadcrumb-item">&raquo;</span>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<% end if %>
			<span class="breadcrumb-item">&raquo;</span>
			<%= DisplayPhrase(reportPageTitlesDictionary,"Eventpayments") %>
			<% showNewHelpContentIcon("event-payments-report") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
			</div>
	<%end if %>
<table height="100%" width="<%=strPageWidth%>" cellspacing="0">
	<tr> 
		<td valign="top" width="100%"> <br />
			<table cellspacing="0" width="90%" style="margin: 0 auto;">
			<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<tr> 
					<td class="headText" align="left"><b><%= pp_PageTitle("Event Payments") %> </b>
					<!--JM - 49_2447-->
					<% showNewHelpContentIcon("event-payments-report") %>

					</td>
				</tr>
			<%end if %>
				<tr> 
					<td valign="top" class="mainText right">&nbsp;</td>
				</tr>
				<tr> 
					<td valign="top" class="mainTextBig"> 
						<table class="mainText" width="95%" cellspacing="0" style="margin: 0 auto;">

							<form name="frmEvtInvoice" method="POST" action="adm_rpt_event_invoice.asp">
								<input type="hidden" name="frmGenReport" value="">
								<input type="hidden" name="frmExpReport" value="">
								<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
									<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
									<input type="hidden" name="category" value="<%=category%>">
								<% end if %>
							<tr> 
								<td class="mainTextBig" colspan="2" valign="top"> 
<table class="mainText" width="92%" cellspacing="0" style="margin: 0 auto;">
	<tr> 
		<td colspan=2 style="background-color:<%=session("pageColor2")%>;" style="height:2px;line-height:2px;font-size:2px;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
	</tr>
	<tr style="background-color:<%=session("pageColor4")%>;"> 
		<td class="whiteHeader"><b>&nbsp;Select Event</b></td>
		<td class="whiteHeader" nowrap style="background-color:<%=session("pageColor4")%>;" align="right" colspan="-1">&nbsp;</td>
	</tr>
	<tr> 
		<td colspan=2 style="background-color:<%=session("pageColor2")%>;" style="height:1px;line-height:1px;font-size:1px;"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
	</tr>
	<tr> 
		<td nowrap height="26" style="background-color:#F2F2F2;" colspan="2"><b>&nbsp;<%=xssStr(allHotWords(77))%>:&nbsp;</b> 
                            <b> 
                            <input onBlur="validateDate(this, '<%=FmtDateShort(pSDate)%>', true);"  type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(pSDate)%>" class="date" />
                        <script type="text/javascript">
			var cal1 = new tcal({'formname':'frmEvtInvoice', 'controlname':'requiredtxtDateStart'});
			cal1.a_tpl.yearscroll = true;
		</script>
                            &nbsp;&nbsp;&nbsp; &nbsp;<%=xssStr(allHotWords(79))%>:</b><b> 
                            <input onBlur="validateDate(this, '<%=FmtDateShort(pEDate)%>', true);"  type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(pEDate)%>" class="date" />
                        <script type="text/javascript">
			var cal2 = new tcal({'formname':'frmEvtInvoice', 'controlname':'requiredtxtDateEnd'});
			cal2.a_tpl.yearscroll = true;
		</script>

				<% 	if session("numLocations")>1 then %>
                            &nbsp;&nbsp;<%=xssStr(allHotWords(8))%>: <select name="optRptLoc" onChange="submitForm();" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
							<option value="0"><%=xssStr(allHotWords(149))%></option>
				<%
						strSQL = "SELECT LocationID, LocationName from Location WHERE wsShow=1 ORDER BY LocationName ASC" 
						rsEvent.CursorLocation = 3
						rsEvent.open strSQL, cnWS
						Set rsEvent.ActiveConnection = Nothing

						Do While NOT rsEvent.EOF
				
				%>
					<option value="<%=rsEvent("LocationID")%>" <% if cLoc=rsEvent("LocationID") then response.write " selected"%>><%=Trim(rsEvent("LocationName"))%></option>
				<%
							rsEvent.MoveNext
						Loop
						rsEvent.close
				%>
                            </select>
				<script type="text/javascript">
					document.frmEvtInvoice.optRptLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
				</script>
				<% end if ' numLocations > 1 %>							
                            </b>
		</td>
	</tr>
	<tr style="background-color:#FAFAFA;"> 
		<td nowrap height="26" colspan="2"><b>&nbsp;<%=xssStr(allHotWords(3))%>:&nbsp; 
<%
dim tmpCourseID

strSQL = "SELECT ClassID, ClassName, ClassDateStart, ClassDateEnd, STARTDATE.courseDateStart, tblCourses.CourseID, tblCourses.CourseName FROM tblClasses INNER JOIN tblClassDescriptions ON tblClasses.DescriptionID=tblClassDescriptions.ClassDescriptionID INNER JOIN tblTypeGroup ON tblTypeGroup.TypeGroupID=tblClassDescriptions.ClassPayment "
strSQL = strSQL & " LEFT OUTER JOIN tblCourses ON tblCourses.CourseID = tblClasses.CourseID "
strSQL = strSQL & " LEFT OUTER JOIN (SELECT MIN(classDateStart) AS courseDateStart, CourseID FROM tblClasses WHERE CourseID IS NOT NULL GROUP BY CourseID) STARTDATE ON STARTDATE.CourseID = tblClasses.CourseID "
strSQL = strSQL & " WHERE tblTypeGroup.wsEnrollment=1 AND ClassDateStart >= " & DateSep & pSDate & DateSep
strSQL = strSQL & " AND ClassDateStart <= " & DateSep & pEDate & DateSep
if NOT checkStudioSetting("tblResvOpts","EnrollReqResource") then
	strSQL = strSQL & " AND ISNULL(tblCourses.PmtPlan, tblClasses.PmtPlan) = 1 "
end if
if cLoc<>0 then
	strSQL = strSQL & " AND ISNULL(tblCourses.LocationID, tblClasses.LocationID) = " & cLoc & " "
end if
strSQL = strSQL & " ORDER BY CASE WHEN tblClasses.CourseID IS NULL THEN tblClasses.ClassDateStart ELSE STARTDATE.CourseDateStart END "

response.write debugSQL(strSQL, "SQL")
rsEvent.CursorLocation = 3
rsEvent.open strSQL, cnWS
Set rsEvent.ActiveConnection = Nothing

%>
                            <select name="optEvent" onChange="submitForm();">
								<option value="0">Summary</option>
                              <%
Do While NOT rsEvent.EOF
	if isNull(rsEvent("CourseID")) then %>
                              <option value="<%=rsEvent("ClassID")%>" 
					<%	if request.form("optEvent")=CSTR(rsEvent("ClassID")) then 
							 		response.write "selected"
									selectedClassDescription = rsEvent("ClassName")
						end if%>><%=rsEvent("ClassName")%> 
                              - <%=FmtDateShort(rsEvent("ClassDateStart"))%></option>
                              <%
		rsEvent.MoveNext
	else
		tmpCourseID = CLNG(rsEvent("CourseID")) %>
                              <option value="<%=rsEvent("ClassID")%>" 
					<%	if request.form("optEvent")=CSTR(rsEvent("ClassID")) then 
							 		response.write "selected"
									selectedClassDescription = rsEvent("CourseName")
						end if%>><%=rsEvent("CourseName")%> 
                              - <%=FmtDateShort(rsEvent("courseDateStart"))%></option>
	<%	cont = true
		do while cont
			rsEvent.MoveNext
			if rsEvent.EOF then
				cont = false
			else
				if isNull(rsEvent("CourseID")) then
					cont = false
				else
					if tmpCourseID<>CLNG(rsEvent("CourseID")) then
						cont = false
					end if
				end if
			end if
		loop
	end if
Loop
rsEvent.close
%>
                            </select>
							
						<br /><% taggingFilter %>&nbsp;&nbsp;
                            <input type="button" name="Button" value="Show Payments" onClick="showResults();">
							<%	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
								else%>
                            <% exportToExcelButton %>
							<% 	end if %>
							<% 	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
							 	else%>
									<% taggingButtons("frmEvtInvoice") %>
							<%	end if%>
                            </b>
		</td>
	</tr>
	<tr> 
		<td colspan=2 style="background-color:<%=session("pageColor4")%>;" style="height:2px;line-height:2px;font-size:2px;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
	</tr>
	<tr> 
		<td nowrap colspan="2">&nbsp;</td>
	</tr>
	<tr valign="middle"> 
		<td nowrap colspan="2"><b> </b></td>
	</tr>
	<tr style="background-color:#F2F2F2;" valign="middle"> 
		<td colspan="2">
<%
else 
	if request.form("frmExpReport")="true" then
		Dim stFilename
		stFilename="attachment; filename=Event Payments Report " & Replace(pSDate,"/","-") & " to " & Replace(pEDate,"/","-") & ".xls" 
		Response.ContentType = "application/vnd.ms-excel" 
		Response.AddHeader "Content-Disposition", stFilename 
	end if
end if ' NOT request.form("frmExpReport")="true"

setRowColors "#F2F2F2", "#FAFAFA" 

if request.form("optEvent")="0" OR request.form("optEvent")="" then 'summary
	if request.form("DisplayResults")="true" OR request.form("frmTagClients")="true" then 

		'summary sql
		'strSQL = "SELECT tblClasses.ClassID, tblClasses.CourseID*-1 as CourseID, tblClassDescriptions.ClassName, tblClasses.ClassDateStart, tblClasses.ClassDateEnd, SUM([PAYMENT DATA].ClientCredit) AS EventFees, Balance.SumOfClientCredits, COUNT([PAYMENT DATA].ClientID) AS NumClients, Location.LocationName, TRAINERS.TrFirstName, TRAINERS.TrLastName "
		'strSQL = strSQL & "FROM tblClasses INNER JOIN tblClassDescriptions ON tblClasses.DescriptionID = tblClassDescriptions.ClassDescriptionID INNER JOIN [PAYMENT DATA] ON [PAYMENT DATA].ClassID = tblClasses.ClassID INNER JOIN "
		'strSQL = strSQL & "(SELECT ClassID, SUM(ClientCredit) AS SumOfClientCredits FROM [Payment Data] WHERE [PAYMENT DATA].Returned = 0 AND ClientCredit > 0 GROUP BY ClassID) Balance ON Balance.ClassID = tblClasses.ClassID INNER JOIN Location ON tblClasses.LocationID = Location.LocationID INNER JOIN TRAINERS ON tblClasses.ClassTrainerID = TRAINERS.TrainerID "
		'if request.form("optFilterTagged")="on" then
		'	strSQL = strSQL & "INNER JOIN tblClientTag ON [Payment Data].ClientID = tblClientTag.clientID "
		'end if
		'strSQL = strSQL & "WHERE [PAYMENT DATA].Returned = 0 AND (tblClasses.ClassDateStart >= " & DateSep & pSDate & DateSep & ") AND (tblClasses.ClassDateStart <= " & DateSep & pEDate & DateSep & ") AND (tblClasses.PmtPlan = 1) AND ([PAYMENT DATA].ClientCredit < 0) "
		'if request.form("optFilterTagged")="on" then
		'	if session("mvaruserID")<>"" then
		'		strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
		'	else
		'		strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
		'	end if
		'end if
		'if cLoc<>0 then
		'	strSQL = strSQL & " AND Location.LocationID=" & cLoc & " "
		'end if
		'strSQL = strSQL & " GROUP BY tblClasses.ClassID, tblClassDescriptions.ClassName, tblClasses.ClassDateStart, tblClasses.ClassDateEnd, tblClasses.CourseID, Balance.SumOfClientCredits, Location.LocationName, TRAINERS.TrFirstName, TRAINERS.TrLastName "
		'strSQL = strSQL & " ORDER BY tblClasses.ClassDateStart, tblClasses.ClassID  "

		strSQL = " SELECT ISNULL(tblClasses.CourseID * - 1, tblClasses.ClassID) AS CID, ISNULL(tblCourses.CourseName, tblClassDescriptions.ClassName) AS CName, "
		strSQL = strSQL & " MIN(tblClasses.ClassDateStart) AS CStart, MAX(tblClasses.ClassDateEnd) AS CEnd, CASE WHEN tblClasses.CourseID IS NULL THEN TRAINERS.TrFirstName ELSE ISNULL(CourseTrainers.TrFirstName, '') END "
		strSQL = strSQL & " AS TrFirstName, CASE WHEN tblClasses.CourseID IS NULL THEN TRAINERS.TrLastName ELSE ISNULL(CourseTrainers.TrLastName, '') END AS TrLastName, ISNULL(CourseLocation.LocationName, "
		strSQL = strSQL & " Location.LocationName) AS LocationName, COUNT(DISTINCT [VISIT DATA].ClientID) AS NumClients "
		strSQL = strSQL & " FROM tblClasses INNER JOIN Location ON tblClasses.LocationID = Location.LocationID "
		strSQL = strSQL & " INNER JOIN TRAINERS ON TRAINERS.TrainerID = tblClasses.ClassTrainerID "
		strSQL = strSQL & " INNER JOIN [VISIT DATA] ON [VISIT DATA].ClassID = tblClasses.ClassID "
		strSQL = strSQL & " INNER JOIN tblClassDescriptions ON tblClasses.DescriptionID = tblClassDescriptions.ClassDescriptionID "
		if request.form("optFilterTagged")="on" then
			strSQL = strSQL & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
		end if
		strSQL = strSQL & " LEFT OUTER JOIN tblCourses ON tblClasses.CourseID = tblCourses.CourseID "
		strSQL = strSQL & " LEFT OUTER JOIN Location AS CourseLocation ON CourseLocation.LocationID = tblCourses.LocationID "
		strSQL = strSQL & " LEFT OUTER JOIN TRAINERS AS CourseTrainers ON tblCourses.TrainerID = CourseTrainers.TrainerID "
		strSQL = strSQL & " WHERE (tblClasses.PmtPlan = 1 OR tblCourses.PmtPlan = 1) "
		if request.form("optFilterTagged")="on" then
			if session("mvaruserID")<>"" then
				strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
			else
				strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
			end if
		end if
		if cLoc<>0 then
			strSQL = strSQL & " AND ISNULL(CourseLocation.LocationID, Location.LocationID) = " & cLoc
		end if
		strSQL = strSQL & " GROUP BY ISNULL(tblClasses.CourseID * - 1, tblClasses.ClassID), ISNULL(tblCourses.CourseName, tblClassDescriptions.ClassName), "
		strSQL = strSQL & " CASE WHEN tblClasses.CourseID IS NULL THEN TRAINERS.TrFirstName ELSE ISNULL(CourseTrainers.TrFirstName, '') END, CASE WHEN tblClasses.CourseID IS NULL THEN TRAINERS.TrLastName ELSE ISNULL(CourseTrainers.TrLastName, '') END, "
		strSQL = strSQL & " ISNULL(CourseLocation.LocationName, Location.LocationName) "
		strSQL = strSQL & " HAVING (MIN(tblClasses.ClassDateStart) >= " & DateSep & pSDate & DateSep & ") AND (MIN(tblClasses.ClassDateStart) <= " & DateSep & pEDate & DateSep & ") "
	response.write debugSQL(strSQL, "SQL")
		rsEvent.CursorLocation = 3
		rsEvent.open strSQL, cnWS
		Set rsEvent.ActiveConnection = Nothing

		'summary sql  old query - left in for troubleshooting
		'strSQL = "SELECT tblClasses.ClassID, tblClasses.CourseID*-1 as CourseID, tblClassDescriptions.ClassName, tblClasses.ClassDateStart, tblClasses.ClassDateEnd, SUM([PAYMENT DATA].PaymentAmount * -1) AS EventFees, SUM([PAYMENT DATA].PaymentAmount) AS SumOfClientCredits, COUNT([PAYMENT DATA].ClientID) AS NumClients, Location.LocationName, TRAINERS.TrFirstName, TRAINERS.TrLastName "
		'strSQL = strSQL & " FROM [PAYMENT DATA] INNER JOIN "
		'strSQL = strSQL & " (SELECT DISTINCT PmtRefNo, ClassID "
		'strSQL = strSQL & " FROM [VISIT DATA]) VD2 ON VD2.PmtRefNo = [PAYMENT DATA].PmtRefNo INNER JOIN "
		'strSQL = strSQL & " CLIENTS ON [PAYMENT DATA].ClientID = CLIENTS.ClientID INNER JOIN "
		'strSQL = strSQL & " tblClasses ON tblClasses.ClassID = VD2.ClassID INNER JOIN tblClassDescriptions ON tblClasses.DescriptionID = tblClassDescriptions.ClassDescriptionID "
		'strSQL = strSQL & " INNER JOIN Location ON tblClasses.LocationID = Location.LocationID INNER JOIN TRAINERS ON tblClasses.ClassTrainerID = TRAINERS.TrainerID LEFT OUTER JOIN "
		'strSQL = strSQL & " (SELECT ClassID, ClientID "
		'strSQL = strSQL & " FROM [PAYMENT DATA] "
		'strSQL = strSQL & " WHERE [PAYMENT DATA].Returned = 0 AND (ISNULL(ClassID, 0) <> 0)) PmtPlanClients ON PmtPlanClients.ClientID = CLIENTS.ClientID AND  "
		'strSQL = strSQL & " PmtPlanClients.ClassID = tblClasses.ClassID "
		'if request.form("optFilterTagged")="on" then
		'	strSQL = strSQL & " INNER JOIN tblClientTag ON [Payment Data].ClientID = tblClientTag.clientID "
		'end if
		'strSQL = strSQL & " WHERE [PAYMENT DATA].Returned = 0 AND (PmtPlanClients.ClientID IS NULL) "
		'strSQL = strSQL & " AND (tblClasses.ClassDateStart >= " & DateSep & pSDate & DateSep & ") AND (tblClasses.ClassDateStart <= " & DateSep & pEDate & DateSep & ") AND (tblClasses.PmtPlan = 1) "
		'if request.form("optFilterTagged")="on" then
		'	if session("mvaruserID")<>"" then
		'		strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
		'	else
		'		strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
		'	end if
		'end if
		'if cLoc<>0 then
		'	strSQL = strSQL & " AND Location.LocationID=" & cLoc & " "
		'end if
		'strSQL = strSQL & " GROUP BY tblClasses.ClassID, tblClassDescriptions.ClassName, tblClasses.ClassDateStart, tblClasses.ClassDateEnd, tblClasses.CourseID, Location.LocationName, TRAINERS.TrFirstName, TRAINERS.TrLastName "
		'strSQL = strSQL & " ORDER BY tblClasses.ClassDateStart, tblClasses.ClassID "

		
		strSQLSumTag = "SELECT DISTINCT [VISIT DATA].ClientID FROM [VISIT DATA] INNER JOIN "
		strSQLSumTag = strSQLSumTag & " tblClasses ON [VISIT DATA].ClassID = tblClasses.ClassID INNER JOIN Location ON tblClasses.LocationID = Location.LocationID "
		if request.form("optFilterTagged")="on" then
			strSQLSumTag = strSQLSumTag & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
		end if
		strSQLSumTag = strSQLSumTag & "WHERE (tblClasses.ClassDateStart >= " & DateSep & pSDate & DateSep & ") AND (tblClasses.ClassDateStart <= " & DateSep & pEDate & DateSep & ") AND (tblClasses.PmtPlan = 1) "
		if request.form("optFilterTagged")="on" then
			if session("mvaruserID")<>"" then
				strSQLSumTag = strSQLSumTag & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
			else
				strSQLSumTag = strSQLSumTag & "AND (tblClientTag.smodeID = 0) " 
			end if
		end if
		if cLoc<>0 then
			strSQLSumTag = strSQLSumTag & " AND Location.LocationID=" & cLoc & " "
		end if

		if request.form("frmTagClients")="true" then
			if request.form("frmTagClientsNew")="true" then
				clearAndTagQuery(strSQLSumTag)
			else
				tagQuery(strSQLSumTag)
			end if
		end if 'tag clients

%>
			<table class="mainText" id="eventPaymentsGenTag" width="100%" cellspacing="0">

	<% 	if NOT request.form("frmExpReport")="true" then %>
				<tr style="background-color:<%=session("pageColor2")%>;"> 
					<td colspan="8" style="height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
			  	</tr>
	<% 	else %>
				<tr>
					<td colspan="8"><%=selectedClassDescription%></td>
				</tr>
	<% 	end if %>
			
				<tr style="background-color:<%=session("pageColor4")%>;">
					<td  width="10%" nowrap class="whiteHeader right"><b>&nbsp;<%= getHotWord(57)%>&nbsp;</b></td>
					<td class="whiteHeader" width="15%"><b>&nbsp;&nbsp;<%= getHotWord(28)%>&nbsp;<%= getHotWord(40)%>&nbsp;</b></td>
	<% 	if NOT request.form("frmExpReport")="true" then %>
					<td class="whiteHeader" width="15%"><b>&nbsp;<script type="text/javascript">document.write('<%=jsEscSingle(allHotWords(6))%>')</script>&nbsp;</b></td>
	<% 	else %>
					<td class="whiteHeader" width="15%"><b>&nbsp;<%= getHotWord(6)%>&nbsp;</b></td>
	<% 	end if %>
					<td class="whiteHeader" width="12%"><b>&nbsp;<%= getHotWord(8)%>&nbsp;</b></td>
					<td  width="12%" nowrap class="whiteHeader center-ch"><b>&nbsp;<%= getHotWord(22)%>&nbsp;<%=session("ClientHW")%>&nbsp;</b></td>
					<td  width="12%" nowrap class="whiteHeader right"><b>&nbsp;<%= getHotWord(22)%>&nbsp;<%= getHotWord(165)%>&nbsp;</b></td>
					<td  width="12%" nowrap class="whiteHeader right"><b>&nbsp;<%= getHotWord(22)%>&nbsp;Received&nbsp;</b></td>
					<td  width="12%" nowrap class="whiteHeader right"><b>&nbsp;<%= getHotWord(22)%>&nbsp;<%= getHotWord(29)%>&nbsp;&nbsp;</b></td>
				</tr>
	<% 	if NOT request.form("frmExpReport")="true" then %>
				<tr style="background-color:<%=session("pageColor2")%>;"> 
					<td colspan="8" style="height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
				</tr>
	<% 	end if %>
<%
		cltCount = 0
		totalFees = 0
		totalPayments = 0
		totalBalances = 0
		cont = true
		useRS1 = false
		useRS2 = false
		
		do while NOT rsEvent.EOF
	
			' def to 0 or blank
			sumDate = ""
			sumName = ""
			sumTrName = ""
			sumLocName = ""
			sumFees = 0
			sumCreds = 0
			sumClients = 0
			sumBalance = 0
			
			sumDate = rsEvent("CStart")
			sumName = rsEvent("CName")
			sumTrName = rsEvent("TrFirstName") & "&nbsp;" & rsEvent("TrLastName")
			sumLocName = rsEvent("LocationName")
			sumClients = sumClients + rsEvent("NumClients")

			if CLNG(rsEvent("CID")) < 0 then ' Course
				strSQL = " SELECT ISNULL(PAIDINFULL.Fees, 0) + ISNULL(PmtPlan.Fees, 0) AS EventFees, ISNULL(PAIDINFULL.Paid, 0) + ISNULL(PmtPlan.Paid, 0) AS SumOfClientCredits "
				strSQL = strSQL & " FROM (SELECT SUM([PAYMENT DATA].PaymentAmount) * - 1 AS Fees, SUM([PAYMENT DATA].PaymentAmount) AS Paid "
					strSQL = strSQL & " FROM [PAYMENT DATA] INNER JOIN "
						strSQL = strSQL & " (SELECT DISTINCT [PAYMENT DATA_2].PmtRefNo FROM [PAYMENT DATA] AS [PAYMENT DATA_2]  "
						strSQL = strSQL & " INNER JOIN [VISIT DATA] ON [PAYMENT DATA_2].PmtRefNo = [VISIT DATA].PmtRefNo "
						strSQL = strSQL & " INNER JOIN tblClasses ON tblClasses.ClassID = [VISIT DATA].ClassID  "
						strSQL = strSQL & " LEFT OUTER JOIN " 
							strSQL = strSQL & " (SELECT [PAYMENT DATA_1].ClientID FROM [PAYMENT DATA] AS [PAYMENT DATA_1] "
							strSQL = strSQL & " INNER JOIN tblClasses AS tblClasses_1 ON tblClasses_1.ClassID = [PAYMENT DATA_1].ClassID "
							strSQL = strSQL & " WHERE ([PAYMENT DATA_1].Returned = 0) AND (ISNULL([PAYMENT DATA_1].ClassID, 0) <> 0) AND tblClasses_1.CourseID = " & CLNG(rsEvent("CID")) * -1 & ") " 
						strSQL = strSQL & " AS PmtPlanClients ON PmtPlanClients.ClientID = [PAYMENT DATA_2].ClientID "
						strSQL = strSQL & " WHERE (tblClasses.CourseID = " & CLNG(rsEvent("CID")) * -1 & ") AND (PmtPlanClients.ClientID IS NULL)) " 
					strSQL = strSQL & " AS PD1 ON PD1.PmtRefNo = [PAYMENT DATA].PmtRefNo) "
				strSQL = strSQL & " AS PAIDINFULL CROSS JOIN "
					strSQL = strSQL & " (SELECT SUM(CASE WHEN [PAYMENT DATA_3].ClientCredit < 0 THEN [PAYMENT DATA_3].ClientCredit ELSE 0 END) AS Fees,  "
					strSQL = strSQL & " SUM(CASE WHEN [PAYMENT DATA_3].ClientCredit > 0 THEN [PAYMENT DATA_3].ClientCredit ELSE 0 END) AS Paid "
					strSQL = strSQL & " FROM [PAYMENT DATA] AS [PAYMENT DATA_3] INNER JOIN "
					strSQL = strSQL & " tblClasses AS tblClasses_2 ON [PAYMENT DATA_3].ClassID = tblClasses_2.ClassID AND tblClasses_2.CourseID = " & CLNG(rsEvent("CID")) * -1 & ") " 
				strSQL = strSQL & " AS PmtPlan "
			else ' Class
				strSQL = " SELECT ISNULL(PAIDINFULL.Fees, 0) + ISNULL(PmtPlan.Fees, 0) AS EventFees, ISNULL(PAIDINFULL.Paid, 0) + ISNULL(PmtPlan.Paid, 0) AS SumOfClientCredits "
				strSQL = strSQL & " FROM (SELECT SUM([PAYMENT DATA].PaymentAmount) * - 1 AS Fees, SUM([PAYMENT DATA].PaymentAmount) AS Paid "
					strSQL = strSQL & " FROM [PAYMENT DATA] INNER JOIN "
						strSQL = strSQL & " (SELECT DISTINCT [PAYMENT DATA_2].PmtRefNo FROM [PAYMENT DATA] AS [PAYMENT DATA_2]  "
						strSQL = strSQL & " INNER JOIN [VISIT DATA] ON [PAYMENT DATA_2].PmtRefNo = [VISIT DATA].PmtRefNo "
						strSQL = strSQL & " LEFT OUTER JOIN " 
							strSQL = strSQL & " (SELECT [PAYMENT DATA_1].ClientID FROM [PAYMENT DATA] AS [PAYMENT DATA_1] "
							strSQL = strSQL & " WHERE ([PAYMENT DATA_1].Returned = 0) AND [PAYMENT DATA_1].ClassID = " & rsEvent("CID") & ") " 
						strSQL = strSQL & " AS PmtPlanClients ON PmtPlanClients.ClientID = [PAYMENT DATA_2].ClientID "
						strSQL = strSQL & " WHERE ([VISIT DATA].ClassID = " & rsEvent("CID") & ")) " 
					strSQL = strSQL & " AS PD1 ON PD1.PmtRefNo = [PAYMENT DATA].PmtRefNo) "
				strSQL = strSQL & " AS PAIDINFULL CROSS JOIN "
					strSQL = strSQL & " (SELECT 0 AS Fees,  "
					strSQL = strSQL & " SUM(ISNULL([PAYMENT DATA_3].ClientCredit, 0)) AS Paid "
					strSQL = strSQL & " FROM [PAYMENT DATA] AS [PAYMENT DATA_3] "
					strSQL = strSQL & " WHERE [PAYMENT DATA_3].ClassID = " & rsEvent("CID") & " "
					strSQL = strSQL & "AND Returned = 0) " 
				strSQL = strSQL & " AS PmtPlan "				
			end if
		response.write debugSQL(strSQL, "SQL")
			rsEvent2.CursorLocation = 3
			rsEvent2.open strSQL, cnWS
			Set rsEvent2.ActiveConnection = Nothing
			
			if NOT rsEvent2.EOF then				
				sumFees = rsEvent2("EventFees")
				sumCreds = rsEvent2("SumOfClientCredits")
				sumBalance = rsEvent2("SumOfClientCredits") + rsEvent2("EventFees")
			else
				sumFees = 0
				sumCreds = 0
				sumBalance = 0
			end if
			rsEvent2.close

%>
			<tr bgcolor=<%=getRowColor(true)%>>
				<td class="right">&nbsp;<%=sumDate%>&nbsp;</td>
				<td>&nbsp;&nbsp;<%=sumName%>&nbsp;</td>
				<td>&nbsp;<%=sumTrName%>&nbsp;</td>
				<td>&nbsp;<%=sumLocName%>&nbsp;</td>
				<td nowrap class="center-ch">&nbsp;<%=sumClients%>&nbsp;</td>
			  <% if NOT request.form("frmExpReport")="true" then %>
				<td nowrap class="right">&nbsp;<%=FmtCurrency(sumFees)%><% if clng(sumFees)>=0 then response.write "&nbsp;" end if %>&nbsp;</td>
				<td nowrap class="right">&nbsp;<%=FmtCurrency(sumCreds)%><% if clng(sumCreds)>=0 then response.write "&nbsp;" end if %>&nbsp;</td>
				<td nowrap class="right">&nbsp;<%=FmtCurrency(sumBalance)%><% if clng(sumBalance)>=0 then response.write "&nbsp;" end if %>&nbsp;</td>
			  <% else %>
				<td nowrap class="right"><%=FmtNumber(sumFees)%></td>
				<td nowrap class="right"><%=FmtNumber(sumCreds)%></td>
				<td nowrap class="right"><%=FmtNumber(sumBalance)%></td>
			  <% end if %>
			</tr>

<%				cltCount = cltCount + sumClients
			totalFees = totalFees + sumFees
			totalPayments = totalPayments + sumCreds
			totalBalances = totalBalances + sumBalance

			rsEvent.MoveNext							
		loop
%>
			  <% if NOT request.form("frmExpReport")="true" then %>
				<tr style="background-color:<%=session("pageColor4")%>;"> 
					<td colspan="8" style="height:2px;line-height:2px;font-size:2px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="2" width="100%"></td>
				</tr>
			  <% end if %>
				<tr style="background-color:#FFFFFF;">
					<td class="right" colspan="4"><b>&nbsp;<%= getHotWord(22)%>:</b></td>
					<td class="center-ch"><b>&nbsp;<%=cltCount%>&nbsp;</b></td>
				  <%	if NOT request.form("frmExpReport")="true" then %>
					<td class="right"><b>&nbsp;<%=FmtCurrency(totalFees)%><% if clng(totalFees)>=0 then response.write "&nbsp;" end if %>&nbsp;</b></td>
					<td class="right"><b>&nbsp;<%=FmtCurrency(totalPayments)%><% if clng(totalPayments)>=0 then response.write "&nbsp;" end if %>&nbsp;</b></td>
					<td class="right"><b>&nbsp;<%=FmtCurrency(totalBalances)%><% if clng(totalBalances)>=0 then response.write "&nbsp;" end if %>&nbsp;</b></td>
				  <% else %>
					<td class="right"><b><%=FmtNumber(totalFees)%></b></td>
					<td class="right"><b><%=FmtNumber(totalPayments)%></b></td>
					<td class="right"><b><%=FmtNumber(totalBalances)%></b></td>
				  <% end if %>
				</tr>
<%		'end if 'end not rsEvent.EOF %>
			</table>
<%		rsEvent.close

	end if 'display results
else ' detailed view
	
	if request.form("optEvent")<>"0" and request.form("optEvent")<>"" then ' detail (a specific class chosen)
		'' Get the class description		
		strSQL = "SELECT ISNULL(tblCourses.CourseName, tblClassDescriptions.ClassName) as ClassName, tblClasses.CourseID FROM tblClasses INNER JOIN tblClassDescriptions ON tblClasses.DescriptionID=tblClassDescriptions.ClassDescriptionID INNER JOIN tblTypeGroup ON tblTypeGroup.TypeGroupID=tblClassDescriptions.ClassPayment "
		strSQL = strSQL & " LEFT OUTER JOIN tblCourses ON tblCourses.CourseID = tblClasses.CourseID "
		strSQL = strSQL & " WHERE tblTypeGroup.wsEnrollment=1 "
		strSQL = strSQL & " AND tblClasses.ClassID = " & request.form("optEvent")
		if NOT checkStudioSetting("tblResvOpts","EnrollReqResource") then
			strSQL = strSQL & " AND ISNULL(tblCourses.PmtPlan, tblClasses.PmtPlan) = 1 "
		end if
		if cLoc<>0 then
			strSQL = strSQL & " AND ISNULL(tblCourses.LocationID, tblClasses.LocationID) = " & cLoc & " "
		end if
		strSQL = strSQL & " ORDER BY tblClasses.ClassDateStart"
	response.write debugSQL(strSQL, "SQL")
		rsEvent.CursorLocation = 3
		rsEvent.open strSQL, cnWS
		Set rsEvent.ActiveConnection = Nothing
		if NOT rsEvent.EOF then
			selectedClassDescription = rsEvent("ClassName")
			tmpCourseID = rsEvent("CourseID")
		end if
		rsEvent.close
	end if
	
	if request.form("DisplayResults")="true" OR request.form("frmTagClients")="true" then 
		if request.form("optEvent")<>"" then
			'strSQL = "SELECT [PAYMENT DATA].ClassID, [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, Sum([PAYMENT DATA].ClientCredit) AS SumOfClientCredit, Min([PAYMENT DATA].ClientCredit) AS EventFee "
			'strSQL = strSQL & "FROM CLIENTS INNER JOIN [PAYMENT DATA] ON CLIENTS.ClientID = [PAYMENT DATA].ClientID "
			'strSQL = strSQL & "WHERE (((CLIENTS.ClientID) In (SELECT CLIENTID FROM [VISIT DATA] WHERE ClassID=" & request.form("optEvent") & "))) "
			'strSQL = strSQL & "GROUP BY [PAYMENT DATA].ClassID, [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName "
			'strSQL = strSQL & "HAVING ((([PAYMENT DATA].ClassID)=" & request.form("optEvent") & ")) "
			'strSQL = strSQL & "ORDER BY CLIENTS.LastName;"

			'strSQL = "SELECT [PAYMENT DATA].ClassID, [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, SUM([PAYMENT DATA].ClientCredit) AS EventFee, Balance.EventBalance AS SumOfClientCredit "
			'strSQL = strSQL & " FROM CLIENTS INNER JOIN [PAYMENT DATA] ON CLIENTS.ClientID = [PAYMENT DATA].ClientID INNER JOIN (SELECT ClassID, ClientID, SUM(ClientCredit) AS EventBalance FROM [PAYMENT DATA] GROUP BY ClassID, ClientID) Balance ON CLIENTS.ClientID = Balance.ClientID AND [PAYMENT DATA].ClassID = Balance.ClassID "
			'if request.form("optFilterTagged")="on" then
			'	strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
			'end if
			'strSQL = strSQL & "WHERE (CLIENTS.ClientID IN (SELECT CLIENTID FROM [VISIT DATA] WHERE ClassID = " & request.form("optEvent") & ")) AND [PAYMENT DATA].ClientCredit < 0 "
			'if request.form("optFilterTagged")="on" then
			'	if session("mvaruserID")<>"" then
			'		strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
			'	else
			'		strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
			'	end if
			'end if
			'strSQL = strSQL & "GROUP BY [PAYMENT DATA].ClassID, [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, Balance.EventBalance HAVING [PAYMENT DATA].ClassID = " & request.form("optEvent") & " ORDER BY CLIENTS.LastName"

			strSQL = "SELECT DISTINCT CLIENTS.ClientID, CLIENTS.LastName, CLIENTS.FirstName, CASE WHEN ISNULL(EF.EventFee, 0) = 0 THEN ISNULL(EF.EventFee, 0) - ISNULL(PA.PackageAmt, 0) ELSE ISNULL(EF.EventFee, 0) END AS EventFee,  "
			strSQL = strSQL & " ISNULL(EB.EventBalance, 0) AS SumOfClientCredit "
			strSQL = strSQL & " FROM [VISIT DATA] INNER JOIN CLIENTS ON [VISIT DATA].ClientID = CLIENTS.ClientID "
			strSQL = strSQL & " INNER JOIN tblClasses ON [VISIT DATA].ClassID = tblClasses.ClassID AND [VISIT DATA].ClassID IS NOT NULL "
			if request.form("optFilterTagged")="on" then
				strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
			end if

			strSQL = strSQL & " LEFT OUTER JOIN (SELECT PmtRefNo, SUM(ISNULL(PaymentAmount, 0)) AS PackageAmt FROM [PAYMENT DATA] "
			strSQL = strSQL & " WHERE Returned = 0 AND PaymentAmount <> 0 AND ISNULL(ClassID, 0) = 0 "
			strSQL = strSQL & " GROUP BY PmtRefNo) PA ON PA.PmtRefNo = [VISIT DATA].PmtRefNo  "

			strSQL = strSQL & " LEFT OUTER JOIN (SELECT ClientID, SUM(Amount) AS EventFee FROM tblClientAccount INNER JOIN tblClasses ON tblClasses.ClassID = tblClientAccount.ClassID "
			strSQL = strSQL & " WHERE PaymentID IS NOT NULL AND (tblClientAccount.ClassID = " & request.form("optEvent")
			if tmpCourseID<>"" then
				strSQL = strSQL & " OR tblClasses.CourseID = " & tmpCourseID
			end if
			strSQL = strSQL & ") GROUP BY ClientID) EF ON EF.ClientID = CLIENTS.ClientID "

			strSQL = strSQL & " LEFT OUTER JOIN (SELECT ClientID, SUM(Amount) AS EventBalance FROM tblClientAccount INNER JOIN tblClasses ON tblClasses.ClassID = tblClientAccount.ClassID "
			strSQL = strSQL & " WHERE (tblClientAccount.ClassID = " & request.form("optEvent")
			if tmpCourseID<>"" then
				strSQL = strSQL & " OR tblClasses.CourseID = " & tmpCourseID
			end if
			strSQL = strSQL & ") GROUP BY ClientID) EB ON EB.ClientID = CLIENTS.ClientID "
			
			strSQL = strSQL & " WHERE ([VISIT DATA].ClassID = " & request.form("optEvent")
			if tmpCourseID<>"" then
				strSQL = strSQL & " OR tblClasses.CourseID = " & tmpCourseID
			end if
			strSQL = strSQL & ") "
			if request.form("optFilterTagged")="on" then
				if session("mvaruserID")<>"" then
					strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
				else
					strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
				end if
			end if

			if request.form("frmTagClients")="true" then
				if request.form("frmTagClientsNew")="true" then
					clearAndTagQuery(strSQL)
				else
					tagQuery(strSQL)
				end if
			end if 'tag clients


			'strSQL = strSQL & " GROUP BY CLIENTS.ClientID, CLIENTS.LastName, CLIENTS.FirstName, [VISIT DATA].ClassID, EF.EventFee, PA.PackageAmt
			strSQL = strSQL & " ORDER BY CLIENTS.LastName "
		response.write debugSQL(strSQL, "SQL")
			logIt strSQL
			rsEvent.CursorLocation = 3
			rsEvent.open strSQL, cnWS
			Set rsEvent.ActiveConnection = Nothing
%>

			<table class="mainText" width="100%" cellspacing="0">
			  <% if NOT request.form("frmExpReport")="true" then %>
				<tr style="background-color:<%=session("pageColor2")%>;"> 
					<td colspan="5" style="height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
				</tr>
			  <% else %>
				<tr>
					<td colspan="5"><%=selectedClassDescription%></td>
				</tr>
			  <% end if %>
				<tr style="background-color:<%=session("pageColor4")%>;"> 
					<td class="whiteHeader" width="40%" colspan="2"><b>&nbsp;<%=session("ClientHW")%>&nbsp;<%= getHotWord(40)%></b></td>
					<td class="whiteHeader right" width="20%"><b><%= getHotWord(28)%>&nbsp;<%= getHotWord(165)%></b></td>
					<td class="whiteHeader right" width="20%"><b><%= getHotWord(28)%>&nbsp;<%= getHotWord(27)%></b></td>
					<td class="whiteHeader right" width="20%"><b><%= getHotWord(28)%>&nbsp;<%= getHotWord(29)%>&nbsp;</b></td>
				</tr>
			<% 	if NOT request.form("frmExpReport")="true" then %>
				<tr style="background-color:<%=session("pageColor2")%>;"> 
					<td colspan="5" style="height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
				</tr>
			<% 	end if %>
<%
			if NOT rsEvent.EOF then

				cltCount = 1
				rowCount = 0
				totalFees = 0
				totalPayments = 0
				totalBalances = 0
				do while NOT rsEvent.EOF		
%>
				<tr bgcolor=<%=getRowColor(true)%>>
							  	<% if NOT request.form("frmExpReport")="true" then %>
					<td nowrap width="1%">&nbsp;<%=cltCount%>.&nbsp;</td>
					<td width="39%">&nbsp;<a href="adm_clt_ph.asp?ID=<%=rsEvent("ClientID")%>&qParam=ph"><%=rsEvent("LastName")%>,&nbsp;<%=rsEvent("FirstName")%></a></td>
								<% else %>
					<td colspan="2" width="40%"><a href="adm_clt_ph.asp?ID=<%=rsEvent("ClientID")%>&qParam=ph"><%=rsEvent("LastName")%>,&nbsp;<%=rsEvent("FirstName")%></a></td>						
								<% end if %>
					<%	if NOT request.form("frmExpReport")="true" then %>
					<td width="20%" class="right">&nbsp;<%=FmtCurrency(rsEvent("EventFee"))%></td>
					<td width="20%" class="right">&nbsp;<%=FmtCurrency(rsEvent("SumOfClientCredit") - rsEvent("EventFee") )%></td>
					<td width="20%" class="right">&nbsp;<%=FmtCurrency(rsEvent("SumOfClientCredit"))%></td>
					<% else %>
					<td width="20%" class="right"><%=FmtNumber(rsEvent("EventFee"))%></td>
					<td width="20%" class="right"><%=FmtNumber(rsEvent("SumOfClientCredit") - rsEvent("EventFee") )%></td>
					<td width="20%" class="right"><%=FmtNumber(rsEvent("SumOfClientCredit"))%></td>
					<% end if %>
				</tr>
<%
						cltCount = cltCount + 1
						totalFees = totalFees + rsEvent("EventFee")
						totalPayments = totalPayments + (rsEvent("SumOfClientCredit") - rsEvent("EventFee"))
						totalBalances = totalBalances + rsEvent("SumOfClientCredit")
						rsEvent.MoveNext
				loop
%>
					<% if NOT request.form("frmExpReport")="true" then %>
				<tr style="background-color:<%=session("pageColor4")%>;"> 
					<td colspan="5" style="height:2px;line-height:2px;font-size:2px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="2" width="100%"></td>
				</tr>
					<% end if %>
				<tr style="background-color:#FFFFFF;"> 
					<td nowrap width="1%">&nbsp;<b><%= getHotWord(22)%>:</b></td>
					<td width="39%">&nbsp;</td>
					<%	if NOT request.form("frmExpReport")="true" then %>
					<td width="20%" class="right">&nbsp;<b><%=FmtCurrency(totalFees)%></b></td>
					<td width="20%" class="right">&nbsp;<b><%=FmtCurrency(totalPayments)%></b></td>
					<td width="20%" class="right">&nbsp;<b><%=FmtCurrency(totalBalances)%></b></td>
					<% else %>
					<td width="20%" class="right"><b><%=FmtNumber(totalFees)%></b></td>
					<td width="20%" class="right"><b><%=FmtNumber(totalPayments)%></b></td>
					<td width="20%" class="right"><b><%=FmtNumber(totalBalances)%></b></td>
					<% end if %>
				</tr>
			</table>


<%
			else	''EOF
%>
				<tr> 
					<td nowrap width="1%">&nbsp;</td>
					<td colspan="4">No&nbsp;Payments</td>
				</tr>
			</table>
<%
			end if ''EOF
			rsEvent.close
		end if '''no event
%>

<%			
	end if '''end display results 
%>			&nbsp;
		</td>
	</tr>
	<tr style="background-color:#F2F2F2;" valign="middle"> 
		<td colspan="5">

<%
	if request.form("DisplayResults")="true" then 
		if request.form("optEvent")<>"" then
			strSQL = "SELECT [PAYMENT DATA].ClassID, [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, Sum([PAYMENT DATA].ClientCredit) AS SumOfClientCredit, Min([PAYMENT DATA].ClientCredit) AS EventFee "
			strSQL = strSQL & "FROM CLIENTS INNER JOIN [PAYMENT DATA] ON CLIENTS.ClientID = [PAYMENT DATA].ClientID INNER JOIN tblClasses ON [PAYMENT DATA].ClassID = tblClasses.ClassID "
			strSQL = strSQL & "WHERE [PAYMENT DATA].Returned = 0 AND ((CLIENTS.ClientID Not In (SELECT CLIENTID FROM [VISIT DATA] INNER JOIN tblClasses ON tblClasses.ClassID = [VISIT DATA].ClassID WHERE ([VISIT DATA].ClassID=" & request.form("optEvent")
			if tmpCourseID<>"" then
				strSQL = strSQL & " OR tblClasses.CourseID = " & tmpCourseID
			end if
			strSQL = strSQL & ")))) "
			strSQL = strSQL & "GROUP BY [PAYMENT DATA].ClassID, tblClasses.CourseID, [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName "
			strSQL = strSQL & "HAVING ([PAYMENT DATA].ClassID=" & request.form("optEvent")
			if tmpCourseID<>"" then
				strSQL = strSQL & " OR tblClasses.CourseID = " & tmpCourseID
			end if
			strSQL = strSQL & ") ORDER BY CLIENTS.LastName;"
		response.write debugSQL(strSQL, "SQL")
			rsEvent.CursorLocation = 3
			rsEvent.open strSQL, cnWS
			Set rsEvent.ActiveConnection = Nothing
%>
			<table class="mainText" width="100%" cellspacing="0">
				<tr style="background-color:#FFFFFF;"> 
					<td class="mainText" width="100%" colspan="5"><b>&nbsp;<%=session("ClientHW")%>s without Reservations</b></td>
				</tr>
				<tr> 
					<td colspan="5" style="background-color:<%=session("pageColor2")%>;" style="height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
				</tr>
				<tr> 
					<td class="whiteHeader" width="25%" colspan="2" nowrap style="background-color:<%=session("pageColor4")%>;"><b>&nbsp;<%=session("ClientHW")%>&nbsp;<%= getHotWord(40)%></b></td>
					<td class="whiteHeader" width="25%" nowrap style="background-color:<%=session("pageColor4")%>;"><b><%= getHotWord(28)%>&nbsp;<%= getHotWord(165)%></b></td>
					<td class="whiteHeader" width="25%" nowrap style="background-color:<%=session("pageColor4")%>;"><b><%= getHotWord(28)%>&nbsp;<%= getHotWord(27)%></b></td>
					<td class="whiteHeader" width="25%" nowrap style="background-color:<%=session("pageColor4")%>;"><b><%= getHotWord(28)%>&nbsp;<%= getHotWord(29)%>&nbsp;</b></td>
				</tr>
				<%	if NOT request.form("frmExpReport")="true" then %>
				<tr> 
					<td colspan="5" style="background-color:<%=session("pageColor2")%>;" style="height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
				</tr>
				<% 	end if %>
<%			if NOT rsEvent.EOF then

				Dim cltCount, totalFees, totalPayments, totalBalances
					cltCount = 1
					'rowCount = 0
					'totalFees = 0
					'totalPayments = 0
					'totalBalances = 0
				do while NOT rsEvent.EOF		
%>
				<tr
<% 
					if rowcount=1 then
						rowcount = 0
						response.write "style=""background-color:#FAFAFA;"""
					else
						rowcount = 1
						response.write "style=""background-color:#F2F2F2;"""
					end if
%>	
							  > 
							  	<% if NOT request.form("frmExpReport")="true" then %>
					<td nowrap width="1%">&nbsp;<%=cltCount%>.&nbsp;</td>
					<td width="25%">&nbsp;<a href="adm_clt_ph.asp?ID=<%=rsEvent("ClientID")%>&qParam=ph"><%=rsEvent("LastName")%>,&nbsp;<%=rsEvent("FirstName")%></a></td>
								<% else %>
					<td colspan="2" width="25%"><a href="adm_clt_ph.asp?ID=<%=rsEvent("ClientID")%>&qParam=ph"><%=rsEvent("LastName")%>,&nbsp;<%=rsEvent("FirstName")%></a></td>						
								<% end if %>
                                <!--td width="30%">&nbsp;<a href="adm_clt_ph.asp?ID=<%=rsEvent("ClientID")%>&qParam=ph"><%=rsEvent("LastName")%>,&nbsp;<%=rsEvent("FirstName")%></a></td-->
						<%	if NOT request.form("frmExpReport")="true" then %>
					<td width="25%">&nbsp;<%=FmtCurrency(rsEvent("EventFee"))%></td>
					<td width="25%">&nbsp;<%=FmtCurrency(rsEvent("SumOfClientCredit") - rsEvent("EventFee") )%></td>
					<td width="25%">&nbsp;<%=FmtCurrency(rsEvent("SumOfClientCredit"))%></td>
						<% else %>
					<td width="25%"><%=FmtNumber(rsEvent("EventFee"))%></td>
					<td width="25%"><%=FmtNumber(rsEvent("SumOfClientCredit") - rsEvent("EventFee") )%></td>
					<td width="25%"><%=FmtNumber(rsEvent("SumOfClientCredit"))%></td>
						<% end if %>
				</tr>
                              <%
						cltCount = cltCount + 1
						totalFees = totalFees + rsEvent("EventFee")
						totalPayments = totalPayments + (rsEvent("SumOfClientCredit") - rsEvent("EventFee"))
						totalBalances = totalBalances + rsEvent("SumOfClientCredit")
						rsEvent.MoveNext
				loop
%>
						<%	if NOT request.form("frmExpReport")="true" then %>
				<tr> 
					<td colspan="5" style="background-color:<%=session("pageColor4")%>;" style="height:2px;line-height:2px;font-size:2px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="2" width="100%"></td>
				</tr>
						<% end if %>
				<tr> 
					<td nowrap width="1%" style="background-color:#FFFFFF;">&nbsp;<b>&nbsp;<%= getHotWord(22)%>:</b></td>
					<td width="30%" style="background-color:#FFFFFF;">&nbsp;</td>
						<%	if NOT request.form("frmExpReport")="true" then %>
					<td width="30%" style="background-color:#FFFFFF;">&nbsp;<b><%=FmtCurrency(totalFees)%></b></td>
					<td width="30%" style="background-color:#FFFFFF;">&nbsp;<b><%=FmtCurrency(totalPayments)%></b></td>
					<td width="30%" style="background-color:#FFFFFF;">&nbsp;<b><%=FmtCurrency(totalBalances)%></b></td>
						<% else %>
					<td width="30%" style="background-color:#FFFFFF;"><b><%=FmtNumber(totalFees)%></b></td>
					<td width="30%" style="background-color:#FFFFFF;"><b><%=FmtNumber(totalPayments)%></b></td>
					<td width="30%" style="background-color:#FFFFFF;"><b><%=FmtNumber(totalBalances)%></b></td>
						<% end if %>
				</tr>
<%
			else	''EOF
%>
				<tr> 
					<td nowrap width="1%">&nbsp;</td>
					<td colspan="4">No&nbsp;Payments</td>
				</tr>
<%			end if ''EOF
%>	
			</table> <%
		end if '''no event
	end if '''end display results 
end if ' detail vs summary
%>		</td>
	</tr> 
</table>
                  				</td>
							</tr>
							<input type="hidden" name="displayResults" value="<%=request.form("displayResults")%>">
							</form>
						</table>
					</td>
				</tr>
			</table>
		</td>
	</tr>

<%
if NOT request.form("frmExpReport")="true" then %>
				
<% 	
end if %>
</table>
<% pageEnd %>
<!-- #include file="post.asp" -->


<%
end if
%>
