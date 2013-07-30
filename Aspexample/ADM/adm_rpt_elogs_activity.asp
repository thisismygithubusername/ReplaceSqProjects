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
<%
if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_DATAACCESS") then 
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
		<!-- #include file="inc_hotword.asp" -->
<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

	dim category : category = ""
	if RQ("category") <> "" then
		category = RQ("category")
	elseif RF("category") <> "" then
		category = RF("category")
	end if

	Dim cSDate, cEDate

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

	disMode = "range"
	' set the row colors
	setRowColors "#FAFAFA", "#F2F2F2"
%>
<% if request.form("frmExpReport")<>"true" then %>

<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
<%= js(array("mb", "adm/adm_rpt_elogs_activity")) %>
<script type="text/javascript">
function exportReport() {
	document.frmELog.frmExpReport.value = "true";
	document.frmELog.frmGenReport.value = "true";
	<% iframeSubmit "frmELog", "adm_rpt_elogs_activity.asp" %>
}
</script>

<%= js(array("calendar" & dateFormatCode, "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %> 

<!-- #include file="../inc_date_ctrl.asp" -->

<% pageStart %>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
<div class="headText breadcrumbs-old" align="left" id="staffActivityTag">
<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
<span class="breadcrumb-item">&raquo;</span>
<%if category <> "" then%>
<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
<span class="breadcrumb-item">&raquo;</span>
<%end if %>
<%=DisplayPhrase(reportPageTitlesDictionary, "Staffactivity")%> &nbsp;&nbsp;&nbsp;
<%if validAccessPriv("RPT_LOGS") then%>
	<a class="textSmall" href="adm_rpt_elogs.asp">[Entry Logs]</a>
	<%end if %>
	<div id="add-to-favorites">
		<span id="" class="favorite"></span><span class="text"><%=DisplayPhrase(reportPageTitlesDictionary, "Addtofavorites")%></span>
	</div>
</div>
<%end if %>

<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
<tr> 
     <td valign="top" height="100%" width="100%"> <br />
        <table cellspacing="0" width="90%" height="100%" style="margin: 0 auto;">
	<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
		  <tr> 
            <td id="Td1" class="headText" align="left"><b>Staff Activity</b>
				&nbsp;&nbsp;&nbsp;<a class="textSmall" href="adm_rpt_elogs.asp">[Entry Logs]</a>
			</td>
          </tr>
	<%end if %>
          <tr> 
            <td valign="top" class="mainText"> 
              <table class="mainText border4 center-block" cellspacing="0">
                <form name="frmELog" action="adm_rpt_elogs_activity.asp" method="POST">
				<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
					<% if category <> "" then %>
						<input type="hidden" name="category" id="category" value="<%=category %>" />
					<%end if %>
					<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
				<%end if %>
					<input type="hidden" name="frmGenReport" value="">
					<input type="hidden" name="frmExpReport" value="">
                  <tr> 
                    <td class="center-ch" valign="bottom" style="background-color:#F2F2F2;"><b>
						&nbsp;<%=xssStr(allHotWords(77))%>: 
                        <input type="text" size="11" name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
                <script type="text/javascript">
			var cal1 = new tcal({'formname':'frmELog', 'controlname':'requiredtxtDateStart'});
			cal1.a_tpl.yearscroll = true;
		</script>
						&nbsp;<%=xssStr(allHotWords(79))%>: 
						<input type="text" size="11" name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
                <script type="text/javascript">
			var cal2 = new tcal({'formname':'frmELog', 'controlname':'requiredtxtDateEnd'});
			cal2.a_tpl.yearscroll = true;
		</script>
						&nbsp; 
					    &nbsp;
					<select name="optUsername">
						<option value="">All Users</option></option>
<%
				strSQL = "SELECT DISTINCT Username FROM tblSavedReport "
				strSQL = strSQL & "WHERE (SavedReport = 0) AND (NOT Username IS NULL) AND (Username <> '') AND (SaveDate >= " & DateSep & cSDate & DateSep & " AND SaveDate <= " & DateSep & cEDate & DateSep & ")	"
				strSQL = strSQL & "ORDER BY Username"
				set rsEntry = Server.CreateObject("ADODB.Recordset")
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing
				do while NOT rsEntry.EOF
%>
					<option value="<%=rsEntry("Username")%>" <%if request.form("optUsername")=CSTR(rsEntry("Username")) then response.write "selected" end if%>><%=rsEntry("Username")%></option>
<%
					rsEntry.MoveNext
				loop
				rsEntry.close
%>
					</select>
					  <br />
						
						<input type="button" name="Button" value="Generate" onClick="genReport();">
						<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
						else%>
								<% exportToExcelButton %>
						<%end if%>
                      </b></td>
                  </tr>
                </form>
              </table>
			</td>
          </tr>
          <tr> 
            <td valign="top" class="mainTextBig"> 
<br />
<% end if ' hide in export %>

<% if request.form("frmGenReport")="true" then %>
              <table id="staffActivityReport" class="mainText" width="100%" cellspacing="0" style="margin: 0 auto;">
                <tr>
                  <td class="mainText" colspan="2" valign="top">
				    <table class="mainText" width="70%" cellspacing="0" style="margin: 0 auto;">
<%
	if request.form("frmExpReport")="true" then
		Dim stFilename
		stFilename="attachment; filename=Staff_Activity_" & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
		Response.ContentType = "application/vnd.ms-excel" 
		Response.AddHeader "Content-Disposition", stFilename 
	end if

	


	strSQL = "SELECT Username, SaveDateTime, SaveReportName, IPAddress, SavedReportID FROM tblSavedReport WHERE (SavedReport = 0) AND (SaveDate >= " & DateSep & cSDate & DateSep & " AND SaveDate <= " & DateSep & cEDate & DateSep & ") "
	if request.form("optUsername")<>"" then
		strSQL = strSQL & " AND (Username = N'" & sqlInjectStr(request.form("optUsername")) & "') "
	end if
	strSQL = strSQL & "ORDER BY SaveDateTime DESC"
response.write debugSQL(strSQL, "SQL")
	dim rsEntry, intCount
	intCount = 0
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing

	if NOT rsEntry.EOF then
%>
						<% if request.form("frmExpReport")<>"true" then %>
							<tr style="background-color:<%=session("pageColor2")%>;"> 
							  <td colspan="7" style="height:2px;line-height:2px;font-size:2px;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
							</tr>
						<% end if %>
							<tr class="whiteHeader" style="background-color:<%=session("pageColor4")%>;"> 
							  <td width="25">&nbsp;</td>
							  <td class="whiteHeader" nowrap style="background-color:<%=session("pageColor4")%>;">&nbsp;<strong><%= getHotWord(57)%>/<%= getHotWord(58)%></strong></td>
							  <td class="whiteHeader">&nbsp;<strong>Report</strong></td>
							  <td class="whiteHeader">&nbsp;<strong><%= getHotWord(41)%></strong></td>
							  <td class="whiteHeader">&nbsp;<strong>IP Address</strong></td>
							  <td>&nbsp;</td>
							  <td>&nbsp;</td>
							</tr>
						<% if request.form("frmExpReport")<>"true" then %>
							<tr style="background-color:<%=session("pageColor2")%>;"> 
							  <td colspan="7" style="height:2px;line-height:2px;font-size:2px;"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
							</tr>
						<% end if %>
<%
		do While NOT rsEntry.EOF
			intCount = intCount + 1
%>
                      <tr bgcolor=<%=getRowColor(true)%>> 
					  <% if request.form("frmExpReport")<>"true" then %>
                        <td align="left" width="1%">&nbsp;<%=intCount%>.&nbsp;</td>
					  <% end if %>
                        <td><% if request.form("frmExpReport")<>"true" then %>&nbsp;<% end if %><%=FmtDateTime(rsEntry("SaveDateTime"))%></td>
                        <td><% if request.form("frmExpReport")<>"true" then %>&nbsp;<% end if %><%=rsEntry("SaveReportName")%></td>
                        <td><% if request.form("frmExpReport")<>"true" then %>&nbsp;<% end if %><%=rsEntry("Username")%></td>
                        <td><% if request.form("frmExpReport")<>"true" then %>&nbsp;<% end if %><%=rsEntry("IPAddress")%></td>
                        <td><% if request.form("frmExpReport")<>"true" then %>&nbsp;&nbsp;<a href="load_report.asp?srid=<%=rsEntry("SavedReportID")%>">[<%= getHotWord(159)%>]</a><% end if %></td>
                        <td>&nbsp;</td>
                      </tr>
                    <% if request.form("frmExpReport")<>"true" then %>
					  <tr> 
                        <td colspan="7" height="1"   style="background-color:#CCCCCC;height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" width="100%" height="1"></td>
                      </tr>
					<% end if %>
<%
			rsEntry.MoveNext
		loop
	else
%>
                      <tr> 
                        <td colspan="7">&nbsp;No Results.</td>
                      </tr>
                      <%
	end if	''rs.EOF
	rsEntry.Close
	Set rsEntry = Nothing
%>
                      <tr> 
                        <td colspan="7">&nbsp;</td>
                      </tr>
                    </table>
				  </td>
                </tr>
                <tr > 
                  <td  colspan="2" valign="top" class="mainTextBig center-ch"> 
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
