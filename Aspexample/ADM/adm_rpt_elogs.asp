<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
%>
		<!-- #include file="../inc_dbconn_wsMaster.asp" -->
		<!-- #include file="inc_accpriv.asp" -->
<%
if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_LOGS") then 
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
		<!-- #include file="inc_rpt_tagging.asp" -->
		<!-- #include file="inc_utilities.asp" -->
		<!-- #include file="inc_rpt_save.asp" -->
		<!-- #include file="inc_hotword.asp" -->
<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

	Dim cSDate, cEDate, ap_RPT_DATAACCESS

	ap_RPT_DATAACCESS = validAccessPriv("RPT_DATAACCESS")
				
	dim category : category = ""
	if (RQ("category"))<>"" then
		category = RQ("category")
	elseif (RF("category"))<>"" then
		category = RF("category") 
	end if

	if request.form("requiredtxtDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSDate = CDATE(request.form("requiredtxtDateStart"))
		Call SetLocale("en-us")
	else
		cSDate = DateAdd("ww",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
	end if

	if request.form("requiredtxtDateEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(request.form("requiredtxtDateEnd"))
		Call SetLocale("en-us")
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if

	if request.form("optFilter")<>"" then
		cLoc = CINT(request.form("optFilter"))
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

	if cLoc = 0 then ' removed All Locations option
		cLoc = -1
	end if

	' set the row colors
	setRowColors "#FAFAFA", "#F2F2F2"
%>
<% if request.form("frmExpReport")<>"true" then %>

<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
<%= js(array("mb", "adm/adm_rpt_elogs", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %>
<script type="text/javascript">
function exportReport() {
	document.frmELog.frmExpReport.value = "true";
	document.frmELog.frmGenReport.value = "true";
	<% iframeSubmit "frmELog", "adm_rpt_elogs.asp" %>
}
</script>

<%= js(array("calendar" & dateFormatCode, "reportFavorites")) %>
<!-- #include file="../inc_date_ctrl.asp" -->
<!-- #include file="../inc_ajax.asp" -->
<!-- #include file="inc_date_arrows.asp" -->
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
			<%= DisplayPhrase(reportPageTitlesDictionary,"Entrylogs") %>
			<% if ap_RPT_DATAACCESS then %>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class="textSmall" href="adm_rpt_elogs_activity.asp">[Staff Activity]</a>
			<% end if %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
			</div>
		<%end if %>
<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
<tr> 
     <td valign="top" height="100%" width="100%"> <br />
        <table class="center" cellspacing="0" width="90%" height="100%">
		<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
		<tr> 
            <td class="headText" align="left"><b><%= pp_PageTitle("Entry Logs") %></b>
			
				<% if ap_RPT_DATAACCESS then %>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class="textSmall" href="adm_rpt_elogs_activity.asp">[Staff Activity]</a>
				<% end if %>

			</td>
          </tr>
		<%end if %>
          <tr> 
            <td valign="top" class="mainText"> 
                <form name="frmELog" action="adm_rpt_elogs.asp" method="POST">
					<input type="hidden" name="frmGenReport" value="">
					<input type="hidden" name="frmExpReport" value="">
					<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
						<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
						<input type="hidden" name="category" value="<%=category%>">
					<% end if %>
              <table class="mainText border4 center" cellspacing="0"> 
                  <tr> 
                    <td class="center-ch" valign="bottom" style="background-color:#F2F2F2;"><b>
						&nbsp;<%=xssStr(allHotWords(77))%>: 
                        <input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text" size="11" name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
        <script type="text/javascript">
			var cal1 = new tcal({'formname':'frmELog', 'controlname':'requiredtxtDateStart'});
			cal1.a_tpl.yearscroll = true;
		</script>
						&nbsp;<%=xssStr(allHotWords(79))%>: 
						<input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text" size="11" name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
        <script type="text/javascript">
			var cal2 = new tcal({'formname':'frmELog', 'controlname':'requiredtxtDateEnd'});
			cal2.a_tpl.yearscroll = true;
		</script>
						&nbsp; 

						<%' BY JM dev list 2138%>
						&nbsp;<%=xssStr(allHotWords(76))%>:&nbsp;
						<select name="optStartTime">
							<option value="-1" <% if request.form("optStartTime")="-1" then response.write "selected" end if %>>Any Time</option>
							<option value="00:00:00" <% if request.form("optStartTime")="00:00:00" then response.write "selected" end if %>>Start of Day</option>
							<option value="01:00:00" <% if request.form("optStartTime")="01:00:00" then response.write "selected" end if %>>1 am</option>
							<option value="02:00:00" <% if request.form("optStartTime")="02:00:00" then response.write "selected" end if %>>2 am</option>
							<option value="03:00:00" <% if request.form("optStartTime")="03:00:00" then response.write "selected" end if %>>3 am</option>
							<option value="04:00:00" <% if request.form("optStartTime")="04:00:00" then response.write "selected" end if %>>4 am</option>
							<option value="05:00:00" <% if request.form("optStartTime")="05:00:00" then response.write "selected" end if %>>5 am </option>
							<option value="06:00:00" <% if request.form("optStartTime")="06:00:00" then response.write "selected" end if %>>6 am</option>
							<option value="07:00:00" <% if request.form("optStartTime")="07:00:00" then response.write "selected" end if %>>7 am</option>
							<option value="08:00:00" <% if request.form("optStartTime")="08:00:00" then response.write "selected" end if %>>8 am</option>
							<option value="09:00:00" <% if request.form("optStartTime")="09:00:00" then response.write "selected" end if %>>9 am</option>
							<option value="10:00:00" <% if request.form("optStartTime")="10:00:00" then response.write "selected" end if %>>10 am</option>
							<option value="11:00:00" <% if request.form("optStartTime")="11:00:00" then response.write "selected" end if %>>11 am</option>
							<option value="12:00:00" <% if request.form("optStartTime")="12:00:00" then response.write "selected" end if %>>12 pm (Noon)</option>
							<option value="13:00:00" <% if request.form("optStartTime")="13:00:00" then response.write "selected" end if %>>1 pm</option>
							<option value="14:00:00" <% if request.form("optStartTime")="14:00:00" then response.write "selected" end if %>>2 pm</option>
							<option value="15:00:00" <% if request.form("optStartTime")="15:00:00" then response.write "selected" end if %>>3 pm</option>
							<option value="16:00:00" <% if request.form("optStartTime")="16:00:00" then response.write "selected" end if %>>4 pm</option>
							<option value="17:00:00" <% if request.form("optStartTime")="17:00:00" then response.write "selected" end if %>>5 pm</option>
							<option value="18:00:00" <% if request.form("optStartTime")="18:00:00" then response.write "selected" end if %>>6 pm</option>
							<option value="19:00:00" <% if request.form("optStartTime")="19:00:00" then response.write "selected" end if %>>7 pm</option>
							<option value="20:00:00" <% if request.form("optStartTime")="20:00:00" then response.write "selected" end if %>>8 pm</option>
							<option value="21:00:00" <% if request.form("optStartTime")="21:00:00" then response.write "selected" end if %>>9 pm</option>
							<option value="22:00:00" <% if request.form("optStartTime")="22:00:00" then response.write "selected" end if %>>10 pm</option>
							<option value="23:00:00" <% if request.form("optStartTime")="23:00:00" then response.write "selected" end if %>>11 pm</option>
							<option value="23:59:59" <% if request.form("optStartTime")="23:59:59" then response.write "selected" end if %>>End of Day</option>
						</select>
						&nbsp;&nbsp;<%=xssStr(allHotWords(78))%>:&nbsp;
						<select name="optEndTime">
							<option value="-1" <% if request.form("optEndTime")="-1" or request.form("optEnDTime")="" then response.write "selected" end if %>>Any Time</option>
							<option value="00:00:00" <% if request.form("optEndTime")="00:00:00" then response.write "selected" end if %>>Start of Day</option>
							<option value="01:00:00" <% if request.form("optEndTime")="01:00:00" then response.write "selected" end if %>>1 am</option>
							<option value="02:00:00" <% if request.form("optEndTime")="02:00:00" then response.write "selected" end if %>>2 am</option>
							<option value="03:00:00" <% if request.form("optEndTime")="03:00:00" then response.write "selected" end if %>>3 am</option>
							<option value="04:00:00" <% if request.form("optEndTime")="04:00:00" then response.write "selected" end if %>>4 am</option>
							<option value="05:00:00" <% if request.form("optEndTime")="05:00:00" then response.write "selected" end if %>>5 am </option>
							<option value="06:00:00" <% if request.form("optEndTime")="06:00:00" then response.write "selected" end if %>>6 am</option>
							<option value="07:00:00" <% if request.form("optEndTime")="07:00:00" then response.write "selected" end if %>>7 am</option>
							<option value="08:00:00" <% if request.form("optEndTime")="08:00:00" then response.write "selected" end if %>>8 am</option>
							<option value="09:00:00" <% if request.form("optEndTime")="09:00:00" then response.write "selected" end if %>>9 am</option>
							<option value="10:00:00" <% if request.form("optEndTime")="10:00:00" then response.write "selected" end if %>>10 am</option>
							<option value="11:00:00" <% if request.form("optEndTime")="11:00:00" then response.write "selected" end if %>>11 am</option>
							<option value="12:00:00" <% if request.form("optEndTime")="12:00:00" then response.write "selected" end if %>>12 pm (Noon)</option>
							<option value="13:00:00" <% if request.form("optEndTime")="13:00:00" then response.write "selected" end if %>>1 pm</option>
							<option value="14:00:00" <% if request.form("optEndTime")="14:00:00" then response.write "selected" end if %>>2 pm</option>
							<option value="15:00:00" <% if request.form("optEndTime")="15:00:00" then response.write "selected" end if %>>3 pm</option>
							<option value="16:00:00" <% if request.form("optEndTime")="16:00:00" then response.write "selected" end if %>>4 pm</option>
							<option value="17:00:00" <% if request.form("optEndTime")="17:00:00" then response.write "selected" end if %>>5 pm</option>
							<option value="18:00:00" <% if request.form("optEndTime")="18:00:00" then response.write "selected" end if %>>6 pm</option>
							<option value="19:00:00" <% if request.form("optEndTime")="19:00:00" then response.write "selected" end if %>>7 pm</option>
							<option value="20:00:00" <% if request.form("optEndTime")="20:00:00" then response.write "selected" end if %>>8 pm</option>
							<option value="21:00:00" <% if request.form("optEndTime")="21:00:00" then response.write "selected" end if %>>9 pm</option>
							<option value="22:00:00" <% if request.form("optEndTime")="22:00:00" then response.write "selected" end if %>>10 pm</option>
							<option value="23:00:00" <% if request.form("optEndTime")="23:00:00" then response.write "selected" end if %>>11 pm</option>
							<option value="23:59:59" <% if request.form("optEndTime")="23:59:59" then response.write "selected" end if %>>End of Day</option>
						</select>																
					  <br />
					    <% showDateArrows("frmELog") %>
						<input type="button" name="Button" value="Generate" onClick="genReport();">
						<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
						else
						    exportToExcelButton
							savingButtons "frmELog", "Entry Logs"
						end if%>
                      </b></td>
                  </tr>
                </form>
              </table>
              
              <br />

			</td>
          </tr>
          <tr> 
            <td valign="top" class="mainTextBig" > 
<% end if ' hide in export %>

<% if request.form("frmGenReport")="true" then %>
              <table id="entryLogsReport" class="mainText" width="100%" cellspacing="0">
                <tr>
                  <td class="mainText" colspan="2" valign="top" >
				    <table class="mainText" width="60%" cellspacing="0" style="margin : 0 auto">
<%
	if request.form("frmExpReport")="true" then
		Dim stFilename
		stFilename="attachment; filename=Entry_Logs_" & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
		Response.ContentType = "application/vnd.ms-excel" 
		Response.AddHeader "Content-Disposition", stFilename 
	end if

    ' Create the recordset
    set rsEntry = Server.CreateObject("ADODB.Recordset")
    
    Dim clientLogins, staffLogins
    
    ' Query to get the total number of staff and client logins. It is just two subselects put together with slightly
    ' different where clauses
    strSQL = "SELECT "
    
    ' Client Count
    strSQL = strSQL & "(SELECT COUNT(1) as c FROM EntryTimes WHERE ClientID >= 0 AND EmpID IS NULL AND Deleted = 0"
    ' Add the full dates
    strSQL = strSQL & " AND EntryDateTime >= " & DateSep & cSDate & DateSep & " "
	strSQL = strSQL & " AND EntryDateTime <= " & DateSep & DateAdd("y", 1, cEDate) & DateSep & " "
	' Add the start/stop times
	if request.form("optStartTime") <> "-1" and request.form("optStartTime") <> "" then
		enttime = split(request.form("optStartTime"), ":")
		strSQL = strSQL & " AND DATENAME(hh, EntryDateTime) >= " & enttime(0) &" AND DATENAME(mi, EntryDateTime) >= " & enttime(1)
	end if
	if request.form("optEndTime") <> "-1" and request.form("optEndTime") <> "" then
		outtime = split(request.form("optEndTime"), ":")
		strSQL = strSQL & " AND DATENAME(hh, EntryDateTime) <= " &  outtime(0) - 1 &" AND DATENAME(n, EntryDateTime) <= " & 60
	end if
    strSQL = strSQL & ") AS ClientCount, "
    
    ' Staff count
    strSQL = strSQL & "(SELECT COUNT(1) as c FROM EntryTimes WHERE ClientID=-1 AND Deleted = 0 "
    ' Add the full dates
    strSQL = strSQL & " AND EntryDateTime >= " & DateSep & cSDate & DateSep & " "
	strSQL = strSQL & " AND EntryDateTime <= " & DateSep & DateAdd("y", 1, cEDate) & DateSep & " "
	' Add the start/stop times
	if request.form("optStartTime") <> "-1" and request.form("optStartTime") <> "" then
		enttime = split(request.form("optStartTime"), ":")
		strSQL = strSQL & " AND DATENAME(hh, EntryDateTime) >= " & enttime(0) &" AND DATENAME(mi, EntryDateTime) >= " & enttime(1)
	end if
	if request.form("optEndTime") <> "-1" and request.form("optEndTime") <> "" then
		outtime = split(request.form("optEndTime"), ":")
		strSQL = strSQL & " AND DATENAME(hh, EntryDateTime) <= " &  outtime(0) - 1 &" AND DATENAME(n, EntryDateTime) <= " & 60
	end if
    strSQL = strSQL & ") AS StaffCount"
    
response.write debugSQL(strSQL, "SQL")
    'response.end
    rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	
	clientLogins = rsEntry("ClientCount")
	staffLogins = rsEntry("StaffCount")
	rsEntry.Close
	
	' Pull out some metrics
	dim webBookings, purchases
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT SUM(tblWSMetrics.Stat3 + tblWSMetrics.Stat4) as WebBookings, SUM(tblWSMetrics.Stat5) as Purchases "
	strSQL = strSQL & "FROM tblWSMetrics  "
	strSQL = strSQL & "WHERE  StatDate >= " & DateSep & cSDate & DateSep & " AND "
	strSQL = strSQL & " StatDate <= " & DateSep & cEDate & DateSep & " "
	strSQL = strSQL & "AND StudioID = " & session("studioID")
	
response.write debugSQL(strSQL, "SQL")
    'response.end
    
    rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnMB
	Set rsEntry.ActiveConnection = Nothing
    
    webBookings = rsEntry("WebBookings")
    purchases = rsEntry("Purchases")
    
    ' Setup for the main query
    rsEntry.Close
    
    ' Select all the EntryTimes that fit our time criteria
	Dim strTempName   
	Dim intCount
	'create SQL select query string
	'strSQL = "SELECT ClientID, LoginName, EntryDateTime FROM EntryTimes "
	strSQL = "SELECT EntryTimes.TimeID, EntryTimes.ClientID, EntryTimes.LoginName, EntryTimes.EntryDateTime, CLIENTS.LastName, CLIENTS.FirstName FROM EntryTimes LEFT OUTER JOIN CLIENTS ON EntryTimes.ClientID = CLIENTS.ClientID "
	if request.form("optFilterTagged")="on" then
		strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
		if session("mvarUserID")<>"" then
			strSQL = strSQL & " AND smodeID = " & session("mvarUserID") & " "
		else
			strSQL = strSQL & " AND smodeID = 0 "
		end if
	end if
	strSQL = strSQL & "WHERE 1=1 "
	strSQL = strSQL & " AND EntryDateTime >= " & DateSep & cSDate & DateSep & " "
	strSQL = strSQL & " AND EntryDateTime <= " & DateSep & DateAdd("y", 1, cEDate) & DateSep & " "
	' BY JM dev list 2138
	if request.form("optStartTime") <> "-1" and request.form("optStartTime") <> "" then
		enttime = split(request.form("optStartTime"), ":")
		strSQL = strSQL & " AND DATENAME(hh, EntryDateTime) >= " & enttime(0) &" AND DATENAME(mi, EntryDateTime) >= " & enttime(1)
	end if
	if request.form("optEndTime") <> "-1" and request.form("optEndTime") <> "" then
		outtime = split(request.form("optEndTime"), ":")
		strSQL = strSQL & " AND DATENAME(hh, EntryDateTime) <= " &  outtime(0) - 1 &" AND DATENAME(n, EntryDateTime) <= " & 60
	end if
	strSQL = strSQL & " AND NOT EntryTimes.LoginName IS NULL "
	strSQL = strSQL & "	AND EntryTimes.Deleted = 0 ORDER BY EntryDateTime DESC"
	
response.write debugSQL(strSQL, "SQL")
    'response.end
	
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	intCount = 0
	
	if NOT rsEntry.EOF then
%>
						<% if request.form("frmExpReport")<>"true" then %>
							<tr style="background-color:<%=session("pageColor2")%>;"> 
							  <td colspan="5" style="height:2px;line-height:2px;font-size:2px;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
							</tr>
						<% end if %>
						    <tr class="whiteHeader" style="background-color:<%=session("pageColor4")%>;"> 
							  <td>&nbsp;</td>
							  <td>&nbsp;<strong>Total Consumer Logins: <%=clientLogins %></strong></td>
							  <td>&nbsp;<strong>Total Staff Logins: <%=staffLogins %></strong></td>
							  <td>&nbsp;</td>
							</tr>
							<tr class="whiteHeader" style="background-color:<%=session("pageColor4")%>;"> 
							  <td>&nbsp;</td>
							  <td>&nbsp;<strong><%= getHotWord(54)%>&nbsp;Bookings: <%=webBookings %></strong></td>
							  <td>&nbsp;<strong><%= getHotWord(164)%>: <%=purchases %></strong></td>
							  <td>&nbsp;</td>
							</tr>
						<% if request.form("frmExpReport")<>"true" then %>
							<tr style="background-color:<%=session("pageColor2")%>;"> 
							  <td colspan="5" style="height:2px;line-height:2px;font-size:2px;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
							</tr>
						<% end if %>
							<tr class="whiteHeader" style="background-color:<%=session("pageColor4")%>;"> 
							  <td>&nbsp;</td>
							  <td>&nbsp;<strong><%= getHotWord(57)%>/<%= getHotWord(58)%></strong></td>
							  <td>&nbsp;<strong><%= getHotWord(41)%></strong></td>
							  <td>&nbsp;<strong><%= getHotWord(40)%></strong></td>
							</tr>
						<% if request.form("frmExpReport")<>"true" then %>
							<tr style="background-color:<%=session("pageColor2")%>;"> 
							  <td colspan="5" style="height:2px;line-height:2px;font-size:2px;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
							</tr>
						<% end if %>
<%
		do While NOT rsEntry.EOF
			intCount = intCount + 1
%>
                      <tr bgcolor=<%=getRowColor(true)%>> 
						<td width="1%">&nbsp;<%=intCount%>.&nbsp;</td>
					    <td><%=FmtDateTime(rsEntry("EntryDateTime"))%></td>
						<td>
				<%if CLNG(rsEntry("ClientID"))<>-1 then%>
								<a href="main_info.asp?ID=<%=rsEntry("ClientID")%>&fl=true"><%=rsEntry("LoginName")%> </a></td>
				<%else%>
								<%=rsEntry("LoginName")%>
						</td>
				<%end if%>
				<%if CLNG(rsEntry("ClientID"))<>-1 AND NOT isNULL(rsEntry("FirstName")) then%>
						<td>&nbsp;<%=rsEntry("FirstName")%>&nbsp;<%=rsEntry("LastName")%></td>
				<%else%>
						<td>&nbsp;</td>
				<%end if%>
                      </tr>
                    <% if request.form("frmExpReport")<>"true" then %>
					  <tr> 
                        <td colspan="5" height="1"   style="background-color:#CCCCCC;height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" width="100%" height="1"></td>
                      </tr>
					<% end if %>
<%
			rsEntry.MoveNext
		loop
	else
%>
                      <tr> 
                        <td colspan="5">&nbsp;No Results.</td>
                      </tr>
                      <%
	end if	''rs.EOF
	rsEntry.Close
	Set rsEntry = Nothing
%>
                      <tr> 
                        <td colspan="5">&nbsp;</td>
                      </tr>
                    </table>
				  </td>
                </tr>
                <tr > 
                  <td class="mainTextBig" colspan="2" valign="top"> 
					</td>
                </tr>
              </table>
	<% end if ' genReport %>
	<% if request.form("frmExpReport")<>"true" then %>
            </td>
          </tr>
          
        </table>
    </td>
    </tr>

		</table>
<% pageEnd %>
<!-- #include file="post.asp" -->


<%
	end if
	
	end if
%>
