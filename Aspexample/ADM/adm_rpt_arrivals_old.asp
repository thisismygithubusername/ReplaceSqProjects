<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
	dim rsEntry, rsDayTotals, rsHourTotals
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	set rsDayTotals = Server.CreateObject("ADODB.Recordset")
	set rsHourTotals = Server.CreateObject("ADODB.Recordset")
%>
	<!-- #include file="inc_accpriv.asp" -->
	<!-- #include file="inc_rpt_tagging.asp" -->
	<!-- #include file="inc_hotword.asp" -->
	<!-- #include file="inc_row_colors.asp" -->
<%
	if not Session("Pass") OR Session("Admin")="false" then 'OR NOT validAccessPriv("RPT_DAY") then 
%>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
	<% else ' has permission to view page %>
		<!-- #include file="../inc_i18n.asp" -->
<%
		dim cLoc, ap_view_all_locs, disMode
		ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")

		If request.form("optLoc")<>"" then
			cLoc = CINT(request.form("optLoc"))
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

		if request.form("optDate")="all" then
			disMode = "all"
		else
			disMode = "range"
		end if

		if NOT request.form("frmExpReport")="true" then
%>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_arrivals_old")) %>
			<script type="text/javascript">
			function exportReport() {
				document.frmParameter.frmGenReport.value = "true";
				document.frmParameter.frmExpReport.value = "true";
				<% iframeSubmit "frmParameter", "adm_rpt_arrivals_old.asp" %>
			}
			</script>
<%
		end if
		
		
		%>
		
		
		<% if NOT request.form("frmExpReport")="true" then %>
<% pageStart %>
			<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
				<tr> 
				<td class="center-ch" valign="top" height="100%" width="100%"> 
				<table cellspacing="0" width="90%" height="100%">
					<tr height="30">
						<td class="headText" align="left" valign="top">
							<table width="100%" cellspacing="0">
								<tr>
								<td class="headText" valign="bottom"><b> <%=ucase(session("ClientHW"))%> ARRIVALS REPORT - Prior to <%=FmtDateShort("8/27/2008")%></b></td>
								<td valign="bottom" class="right" height="26"> </td>
								</tr>
							</table>
						</td>
					</tr>
					<tr> 
					<td height="30"  valign="bottom" class="center-ch headText">
					<table class="mainText border4" cellspacing="0">
						<form name="frmParameter" action="adm_rpt_arrivals_old.asp" method="POST">
						<input type="hidden" name="frmGenReport" value="">
						<input type="hidden" name="frmExpReport" value="">
						<tr> 
						<td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b>&nbsp;&nbsp;

&nbsp;Date&nbsp;Range:</span> 
                            <input onClick="document.frmParameter.submit();" type="radio" name="optDate" value="all" <%if disMode="all" then response.write "checked" end if%>><%=xssStr(allHotWords(149))%>&nbsp;&nbsp; 
                            <input onClick="document.frmParameter.submit();" type="radio" name="optDate" value="range" <%if disMode="range" then response.write "checked" end if%>><%=xssStr(allHotWords(77))%>: 
                            <input onBlur="document.frmParameter.submit();" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
                      <script type="text/javascript">
			var cal1 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStart'});
			cal1.a_tpl.yearscroll = true;
		</script>
                            <input onBlur="document.frmParameter.submit();" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
                      <script type="text/javascript">
			var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
			cal2.a_tpl.yearscroll = true;
		</script>
                            &nbsp; </b>
							&nbsp;&nbsp;
							<br />
						<%= getHotWord(8)%>:&nbsp;
						<select name="optLoc" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
							<option value="0" <% if cLoc=0 then response.write "selected" end if %>>All</option>
						<%
						strSQL = "SELECT LocationID, LocationName FROM Location WHERE [Active]=1 AND wsShow=1 ORDER BY LocationName "
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						Set rsEntry.ActiveConnection = Nothing
	
						do While NOT rsEntry.EOF			
							%>
								<option value="<%=rsEntry("LocationID")%>" <%if cLoc=rsEntry("LocationID") then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
							<%
							rsEntry.MoveNext
						loop
						rsEntry.close
						%>
						</select>
						<script type="text/javascript">
							document.frmParameter.optLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>'+ " <%=jsEscDouble(allHotWords(8))%>s";
						</script>
						&nbsp;&nbsp;
						<% taggingFilter %>
						<input type="button" name="Button" value="Generate" onClick="genReport();">
						<% exportToExcelButton %>
						</b>&nbsp;&nbsp;
						</td>
						</tr>
						
						</form>
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" class="mainTextBig center-ch"> 
					
					<table class="mainText" width="100%" cellspacing="0">
						<tr>
						<td  colspan="2" valign="top" class="mainTextBig center-ch">&nbsp;</td>
						</tr>
						<tr > 
						<td  colspan="2" valign="top" class="mainTextBig center-ch">
<% 
		end if			'end of frmExpreport value check before /head line	  
	
		dim startTime, endTime, numHours, numDays, i, j, startDate, startHour, endHour, tmpDate, tmpDoy, hourTotal, hourTotalsExist
		setRowColors "#F2F2F2", "#FAFAFA"
		hourTotalsExist = false
	
		if request.form("frmGenReport")="true" then 
			if request.form("frmTagClients")="true" then
				
				' tagging sql
				
			end if	'End Tag Clients
			
			'***************** QUERY STUFF **********************
			
			if disMode = "all" then
				numDays = 30
			else
				numDays = datediff("d", cSDate, cEDate)
			end if

			startDate = DATEVALUE(DateAdd("d", -numDays, DateAdd("n", Session("tzOffset"),Now)))

			strSQL = "SELECT MIN(DATEPART(hh, EntryDateTime)) AS MinTime, MAX(DATEPART(hh, EntryDateTime)) AS MaxTime "
			strSQL = strSQL & "FROM EntryTimes WHERE (Location IS NOT NULL) AND (Location <> 98) "
			if cLoc <> "" AND cLoc <> "0" then
				strSQL = strSQL & " AND (Location = " & cLoc & ") "
			end if
			if disMode <> "all" then
				strSQL = strSQL & " AND EntryDateTime >= " & DateSep & cSDate & DateSep & " "
				strSQL = strSQL & " AND EntryDateTime <= " & DateSep & DateAdd("y", 1, cEDate) & DateSep & " "
			else
				strSQL = strSQL & " AND EntryDateTime >= " & DateSep & startDate & DateSep & " "
			end if
			
		response.write debugSQL(strSQL, "SQL")
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			
			if NOT isNull(rsEntry("MinTime")) then
				startHour = rsEntry("MinTime")
				startTime = CDATE(rsEntry("MinTime") & ":00:00")
			else
				startTime = CDATE("12:00:00 AM")
				startHour = 0
				
			end if
			
			if NOT isNULL(rsEntry("MaxTime")) then
				endHour = rsEntry("MaxTime")
			else
				endHour = 0
			end if
				
			'response.write startTime
			rsEntry.close
			
			numHours = endHour - startHour
			
			
			strSQL = "SELECT DATEPART(y, EntryDateTime) AS Doy, DATEPART(hh, EntryDateTime) AS Hour, COUNT(*) AS Arrivals FROM EntryTimes "
			if request.form("optFilterTagged")="on" then
				strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = EntryTimes.ClientID "
				if session("mVarUserID")<>"" then
					strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
				end if
				strSQL = strSQL & " ) "
			end if
			strSQL = strSQL & "WHERE (Location IS NOT NULL) AND (Location <> 98) "
			if cLoc <> "" AND cLoc <> "0" then
				strSQL = strSQL & " AND (Location = " & cLoc & ") "
			end if
			if disMode <> "all" then
				strSQL = strSQL & " AND EntryDateTime >= " & DateSep & cSDate & DateSep & " "
				strSQL = strSQL & " AND EntryDateTime <= " & DateSep & DateAdd("y", 1, cEDate) & DateSep & " "
			else
				strSQL = strSQL & " AND EntryDateTime >= " & DateSep & startDate & DateSep & " "
			end if
			strSQL = strSQL & "GROUP BY DATEPART(y, EntryDateTime), DATEPART(hh, EntryDateTime) ORDER BY Doy, Hour "

		response.write debugSQL(strSQL, "SQL")
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			
			' Query Day Totals
			strSQL = "SELECT DATEPART(y, EntryDateTime) AS Doy, COUNT(*) AS Arrivals FROM EntryTimes "
			if request.form("optFilterTagged")="on" then
				strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = EntryTimes.ClientID "
				if session("mVarUserID")<>"" then
					strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
				end if
				strSQL = strSQL & " ) "
			end if
			strSQL = strSQL & "WHERE (Location IS NOT NULL) AND (Location <> 98) "
			if cLoc <> "" AND cLoc <> "0" then
				strSQL = strSQL & " AND (Location = " & cLoc & ") "
			end if
			if disMode <> "all" then
				strSQL = strSQL & " AND EntryDateTime >= " & DateSep & cSDate & DateSep & " "
				strSQL = strSQL & " AND EntryDateTime <= " & DateSep & DateAdd("y", 1, cEDate) & DateSep & " "
			else
				strSQL = strSQL & " AND EntryDateTime >= " & DateSep & startDate & DateSep & " "
			end if
			strSQL = strSQL & "GROUP BY DATEPART(y, EntryDateTime) ORDER BY Doy "
			
		response.write debugSQL(strSQL, "SQL")
			rsDayTotals.CursorLocation = 3
			rsDayTotals.open strSQL, cnWS
			Set rsDayTotals.ActiveConnection = Nothing
			
			' Query Hour Totals
			strSQL = "SELECT DATEPART(hh, EntryDateTime) AS Hour, COUNT(*) AS Arrivals FROM EntryTimes "
			if request.form("optFilterTagged")="on" then
				strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = EntryTimes.ClientID "
				if session("mVarUserID")<>"" then
					strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
				end if
				strSQL = strSQL & " ) "
			end if
			strSQL = strSQL & "WHERE (Location IS NOT NULL) AND (Location <> 98) "
			if cLoc <> "" AND cLoc <> "0" then
				strSQL = strSQL & " AND (Location = " & cLoc & ") "
			end if
			if disMode <> "all" then
				strSQL = strSQL & " AND EntryDateTime >= " & DateSep & cSDate & DateSep & " "
				strSQL = strSQL & " AND EntryDateTime <= " & DateSep & DateAdd("y", 1, cEDate) & DateSep & " "
			else
				strSQL = strSQL & " AND EntryDateTime >= " & DateSep & startDate & DateSep & " "
			end if
			strSQL = strSQL & "GROUP BY DATEPART(hh, EntryDateTime) ORDER BY Hour "
			
		response.write debugSQL(strSQL, "SQL")
			rsHourTotals.CursorLocation = 3
			rsHourTotals.open strSQL, cnWS
			Set rsHourTotals.ActiveConnection = Nothing
			
			' ****** have to check for records here for later
			if NOT rsHourTotals.EOF then
				hourTotalsExist = true
			end if
	
			if request.form("frmExpReport")="true" then
				Dim stFilename
				
				stFilename = "attachment; filename=Client Arrival Report.xls"
				Response.ContentType = "application/vnd.ms-excel" 
				Response.AddHeader "Content-Disposition", stFilename 
			end if
	
%>
			<table class="mainText"  cellspacing="0">	
				<tr>
					<td>&nbsp;</td>
				<% for i=0 to numHours %>
					<td class="right"><% if NOT request.form("frmExpReport")="true" then %>&nbsp;&nbsp;<% end if %><%=FmtTimeShorter(DATEADD("h", i, startTime))%>- <%=FmtTimeShort(DATEADD("n", i * 60 + 59, startTime))%></td>
				<% next %>
					<td class="right"><% if NOT request.form("frmExpReport")="true" then %>&nbsp;&nbsp;<% end if %><b>Day Total</b></td>
				</tr>
				
				<% if NOT request.form("frmExpReport")="true" then %>
					<tr height="2">
						<td colspan="<%=numHours + 3%>" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
					</tr>
				<% end if %>	
				
				<% for i=0 to numDays
					if disMode<>"all" then
						tmpDate = DATEADD("d", i, cSDate)
					else
						tmpDate = DATEADD("d", i, startDate)
					end if
					tmpDoy = DATEPART("y", tmpDate)
					%>
					<tr style="background-color:<%=getRowColor(true)%>;">
						<td><%=FmtDateShort(tmpDate)%></td>
						<% for j=startHour to endHour %>
							<td class="right">
<% 
							if NOT rsEntry.EOF then
								'response.write "(" & rsEntry("Doy") & " " & tmpDoy & ")<br />"
								'response.write "(" & rsEntry("Hour") & " " & j & ")"
								if rsEntry("Doy") = tmpDoy AND j = rsEntry("Hour") then
									response.write rsEntry("Arrivals")
									rsEntry.moveNext
								else
									response.write "0"
								end if 
							else
								response.write "0"
							end if
%>
							</td>
						<% next 'hour loop %>
						<td class="right"><b>
<%
						if NOT rsDayTotals.EOF then
							if rsDayTotals("Doy") = tmpDoy then
								response.write rsDayTotals("Arrivals")
								rsDayTotals.moveNext
							else
								response.write "0"
							end if
						else
							response.write "0"
						end if %>
						</b></td>
					</tr>
				<% next 'day loop%>
				<tr style="background-color:<%=getRowColor(true)%>;">
					<td><b>Hour Total</b></td> 
<% 
					hourTotal = 0
					for i=startHour to endHour %>
						<td class="right"><b>
<%
						if NOT rsHourTotals.EOF then
							if rsHourTotals("Hour") = i then
								response.write rsHourTotals("Arrivals")
								hourTotal = hourTotal + rsHourTotals("Arrivals")
								rsHourTotals.moveNext
							else
								response.write "0"
							end if
						else
							response.write "0"
						end if
%>
						</b></td>
					<% next %>
					<td class="right"><b><%=hourTotal%></b></td>
				</tr>
				
				<tr style="background-color:<%=getRowColor(true)%>;">
					<td>Hour Average</td> 
<%
					if hourTotalsExist then
						rsHourTotals.moveFirst
					end if
						
					for i=startHour to endHour %>
						<td class="right">
<%
						if NOT rsHourTotals.EOF then
							if rsHourTotals("Hour") = i AND numDays > 0 then
								response.write FmtNumber(rsHourTotals("Arrivals") / numDays)
								rsHourTotals.moveNext
							else
								response.write "0"
							end if
						else
							response.write "0"
						end if
%>
						</td>
					<% next %>
					<td class="right"><b>
<%
					if numDays > 0 then
						response.write FmtNumber(hourTotal / numDays)
					else
						response.write "0"
					end if
%>
					</b></td>
				</tr>
				
			</table>
<%
			'rsEntry.close
			'set rsEntry = nothing
		end if		'end of generate report if statement
%>
					</table>
				</table>
			</td>
		</tr>
</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%
	
end if
%>
