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
Dim ap_ipay, ap_ipay_trn
ap_ipay = validAccessPriv("RPT_IPAY")
ap_ipay_trn = validAccessPriv("RPT_IPAY_TRN")
if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_DAY") OR (NOT ap_ipay AND NOT ap_ipay_trn) then 
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
		<!-- #include file="inc_hotword.asp" -->
		<!-- #include file="../inc_ajax.asp" -->
		<!-- #include file="../inc_val_date.asp" -->
<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

	Dim cSDate, cEDate, trnTotCls, trnTotVD, totCls, totVD, ss_asst1, ss_asst2, pView, trnIDStr
	ss_asst1 = checkStudioSetting("tblGenOpts", "UseAsst1")
	ss_asst2 = checkStudioSetting("tblGenOpts", "UseAsst2")

	dim category : category = ""
	if (RQ("category"))<>"" then
		category = RQ("category")
	elseif (RF("category"))<>"" then
		category = RF("category") 
	end if

	Dim strTG, tgArr, tgChk, numTGs, firstTG

		strTG = sqlInjectStr(request.form("optTG"))
		tgArr = Split(strTG,",")
		tgChk = "," & Replace(strTG, " ", "") & ","


	if request.form("optView")<>"" then
		pView = sqlInjectStr(request.form("optView"))
	else
		pView = "2"
	end if
	
	if request.form("requiredtxtDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSDate = CDATE(sqlInjectStr(request.form("requiredtxtDateStart")))
		Call SetLocale("en-us")
	else
		'cSDate = DateAdd("y",-14,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
		cSDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if

	if request.form("requiredtxtDateEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(sqlInjectStr(request.form("requiredtxtDateEnd")))
		Call SetLocale("en-us")
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if
	trnTotCls = 0
	trnTotVD = 0
	totCls = 0
	totVD = 0
	totTrnPay = 0

	set rsEntry = Server.CreateObject("ADODB.Recordset")


	if NOT request.form("frmExpReport")="true" then
%>

<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "calendar" & dateFormatCode, "adm/adm_rpt_asst", "reportFavorites", "plugins/jquery.SimpleLightBox" )) %>
<%= css(array("SimpleLightBox")) %> 

<script type="text/javascript">
function exportReport() {
	document.frmPayroll.frmExpReport.value = "true";
	document.frmPayroll.frmShowReport.value = "true";
	<% iframeSubmit "frmPayroll", "adm_rpt_asst.asp" %>
}
</script>

<!-- #include file="../inc_date_ctrl.asp" -->
<%
end if 'excel

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
	<%= DisplayPhrase(reportPageTitlesDictionary, "Assistants") %>
	<div id="add-to-favorites">
		<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
	</div>
	</div>
<% end if %>

<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
<tr> 
     <td valign="top" height="100%" width="100%"> <br />
        <table class="center" cellspacing="0" width="90%" height="100%">
		<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
		<tr> 
            <td class="headText" align="left"><b>
			<% if ss_asst1 AND ss_asst2 then %>gen
				<%=xssStr(allHotWords(13))%> / <%=xssStr(allHotWords(15))%>
			<% else %>
				<%=xssStr(allHotWords(13))%>
			<% end if %>
			 </b></td>
          </tr>
		<% end if %>
          <tr>
            <td valign="top" class="mainText">
              <form name="frmPayroll" action="adm_rpt_asst.asp" method="POST">
                <input type="hidden" name="frmShowReport" value="request.form("frmShowReport")">
				<input type="hidden" name="frmExpReport" value="">
				<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
				<input type="hidden" name="category" value="<%=category%>">
				<% end if %>
            <table class="mainText border4 center" cellspacing="0">
                <tr>
                  <td align="left" valign="bottom" style="background-color:#F2F2F2;">
				  	<table class="mainText" cellspacing="0">
						<tr>
							<td>
				  <b>&nbsp;<%=xssStr(allHotWords(159))%>:
				<% if ss_asst1 AND ss_asst2 then %>				  
				  <select name="optView" onChange="document.frmPayroll.submit();">
					<option value="2" <%if pView="0" then response.write "selected" end if%>><%=xssStr(allHotWords(149))%></option>
					<option value="0" <%if pView="0" then response.write "selected" end if%>><%=xssStr(allHotWords(13))%>s</option>
					<option value="1" <%if pView="1" then response.write "selected" end if%>><%=xssStr(allHotWords(15))%>s</option>
                    </select> 

				<% end if %>
				<% if ap_ipay then %>
                    <select name="optAssistant" id="optAssistant">
					<option value="-1"><%=xssStr(allHotWords(149))%></option>
<%
		if pView="0" then
			strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName FROM TRAINERS INNER JOIN [VISIT DATA] ON TRAINERS.TrainerID = [VISIT DATA].TrainerID2 WHERE ([VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & ") "
		elseif pView="1" then
			strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName FROM TRAINERS INNER JOIN [VISIT DATA] ON TRAINERS.TrainerID = [VISIT DATA].TrainerID3 WHERE ([VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & ") "
		else	'pView = "2"
			strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName FROM TRAINERS LEFT OUTER JOIN [VISIT DATA] ON TRAINERS.TrainerID = [VISIT DATA].TrainerID2 AND NOT ([VISIT DATA].TrainerID2 IS NULL) AND [VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & " LEFT OUTER JOIN [VISIT DATA] [VISIT DATA_1] ON TRAINERS.TrainerID = [VISIT DATA_1].TrainerID3 AND NOT ([VISIT DATA_1].TrainerID3 IS NULL) AND [VISIT DATA_1].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & " WHERE ((NOT ([VISIT DATA].TrainerID2 IS NULL)) OR (NOT ([VISIT DATA_1].TrainerID3 IS NULL)))"
		end if
		strSQL = strSQL & " ORDER BY "
		strSQL = strSQL & GetTrnOrderBy()
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing

		do while not rsEntry.EOF
%>
			<option value="<%=rsEntry("TrainerID")%>" <%if CSTR(request.form("optAssistant"))=CSTR(rsEntry("TrainerID")) then response.write "selected" end if%>><%=FmtTrnNameNew(rsEntry, true)%></option>
<%
			rsEntry.MoveNext
		loop
		rsEntry.close
%>
		</select> 
		<% end if %>
		<select name="optDetail">
			<option value="" <%if request.form("optDetail")="" then response.write "selected" end if%>>Detail</option>
			<option value="2" <%if request.form("optDetail")="2" then response.write "selected" end if%>>Summary</option>
		</select>
	<!--
		<select name="optTG" onchange="document.frmSales.submit();">
						<option value="0" <%if request.form("optTG")="0" then response.write "selected" end if%>>All Type Groups</option>
						<%
							strSQL = "SELECT TypegroupID, Typegroup FROM tblTypegroup "
							strSQL = strSQL & "WHERE [Active]=1 "
							strSQL = strSQL & "ORDER BY Typegroup"
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>	
									<option value="<%=rsEntry("TypegroupID")%>" <%if request.form("optTG")=CSTR(rsEntry("TypegroupID")) then response.write "selected" end if%>><%=rsEntry("Typegroup")%></option>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
					  </select>
		--></td><td rowspan="2" valign="top">
			<select name="optTG" size="3" multiple <%showMultiSelectTitle() %>>
				<option value="0"  <%if inStr(tgChk, ",0,")>0 OR request.form("optTG")="" then response.write "selected" end if%>>All</option>
<%
						strSQL = "SELECT DISTINCT TypegroupID, Typegroup FROM tblTypegroup WHERE [Active]=1 AND wsResource=0 ORDER BY Typegroup"
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						Set rsEntry.ActiveConnection = Nothing

						Do While NOT rsEntry.EOF
%>
						<option value="<%=rsEntry("TypegroupID")%>" <%if inStr(tgChk, "," & rsEntry("TypeGroupID") & ",") > 0 then response.write "selected" end if%>><%=rsEntry("Typegroup")%></option>
<%
							rsEntry.MoveNext
						Loop	
						rsEntry.close
%>
               </select>
				<script type="text/javascript">
					document.frmPayroll.optTG.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' + " <%=jsEscDouble(allHotWords(503))%>";
				</script>							  
					  </td>
					 </tr>
					<tr>
						<td class="center-ch">
                    &nbsp;<%=xssStr(allHotWords(77))%>:
                      <input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
              <script type="text/javascript">
			var cal1 = new tcal({'formname':'frmPayroll', 'controlname':'requiredtxtDateStart'});
			cal1.a_tpl.yearscroll = true;
		    </script>
        &nbsp;<%=xssStr(allHotWords(79))%>:
        <input  onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
              <script type="text/javascript">
			var cal2 = new tcal({'formname':'frmPayroll', 'controlname':'requiredtxtDateEnd'});
			cal2.a_tpl.yearscroll = true;
		    </script>
		</td></tr><tr><td class="center-ch" colspan="2">
		        <input type="button" name="Button" value="Generate" onClick="showReport();">
				<%if NOT validAccessPriv("RPT_EXPORT") then
				else%>
						 <span class="icon-button" style="vertical-align: middle;" title="Export to Excel" ><a onClick="exportReport();" ><img src="<%= contentUrl("/asp/adm/images/export-to-excel-20px.png") %>" /></a></span> 
				<%end if%>
                  &nbsp;</b>
				  	</td>
					</tr>
					</table>
				  </td>
                </tr>
              </form>
            </table></td>
          </tr>
          <tr> 
            <td valign="top" class="mainTextBig center-ch"> 
				<br />
	<% end if 'excel %>
	<% if request.form("frmShowReport")="true" then %>
<%
			if request.form("frmExpReport")="true" then
				Dim stFilename
				if showDetails then 
					stFilename="attachment; filename=Services Sold-Detail " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
				else 
					stFilename="attachment; filename=Services Sold-Summary " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
				end if
				Response.ContentType = "application/vnd.ms-excel" 
				Response.AddHeader "Content-Disposition", stFilename 
			end if
%>	
              <table class="mainText center" width="80%" cellspacing="0">
			<% if request.form("optDetail")="2" then ' summary view %>
                <tr>
                  <td width="50%" colspan="2" align="left" nowrap class="mainText center">&nbsp;</td>
                  <td  nowrap class="center-ch mainText"><strong>&nbsp;# <%=session("ClientHW")%>s</strong></td>
                  <td  nowrap class="center-ch mainText"><strong>&nbsp;# <%= getHotWord(5)%></strong></td>
                  <td  nowrap class="center-ch mainText"><strong>&nbsp;Rate</strong></td>
                  <td  nowrap class="right mainText"><strong>&nbsp;Earnings&nbsp;&nbsp;</strong></td>
				</tr>
				<%if NOT request.form("frmExpReport")="true" then%>
                <tr style="background-color:#CCCCCC;"><td colspan="8" height="2" width="100%"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td></tr>
				<% end if %>
			<% end if %>
<%
		set rsEntry2 = Server.CreateObject("ADODB.Recordset")
		if request.form("optAssistant")<>"" AND request.form("optAssistant")<>"-1" then	'Single Trainer
			strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.EmpID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName, tblPayRates.Rate1 FROM TRAINERS LEFT OUTER JOIN tblPayRates ON TRAINERS.TrainerID = tblPayRates.TrainerID WHERE (TRAINERS.TrainerID = " & sqlInjectStr(request.form("optAssistant")) & ")"
		elseif NOT ap_ipay then
		  strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.EmpID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName, tblPayRates.Rate1 FROM TRAINERS LEFT OUTER JOIN tblPayRates ON TRAINERS.TrainerID = tblPayRates.TrainerID WHERE (TRAINERS.TrainerID = " & session("empID") & ")"
		else
			if pView="0" then
				strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.EmpID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName, tblPayRates.Rate1 FROM TRAINERS INNER JOIN [VISIT DATA] ON TRAINERS.TrainerID = [VISIT DATA].TrainerID2 LEFT OUTER JOIN tblPayRates ON TRAINERS.TrainerID = tblPayRates.TrainerID WHERE ([VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & ") "
				if inStr(tgChk, ",0,")>0 OR request.form("optTG")="" then
					''All TGs
				else
					firstTG = true
					numTGs = 0
					do While numTGs < UBound(tgArr)+1
						if tgArr(numTGs)<>"0" then
							if firstTG then
								strSQL = strSQL & " AND ([VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
								firstTG = false
							else
								strSQL = strSQL & " OR [VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
							end if
						end if
						numTGs = numTGs + 1
					loop
					strSQL = strSQL & ") "
				end if
				strSQL = strSQL & " ORDER BY "
				strSQL = strSQL & GetTrnOrderBy()
			elseif pView="1" then
				strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.EmpID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName, tblPayRates.Rate1 FROM TRAINERS INNER JOIN [VISIT DATA] ON TRAINERS.TrainerID = [VISIT DATA].TrainerID3 LEFT OUTER JOIN tblPayRates ON TRAINERS.TrainerID = tblPayRates.TrainerID WHERE ([VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & ") "
				if inStr(tgChk, ",0,")>0 OR request.form("optTG")="" then
					''All TGs
				else
					firstTG = true
					numTGs = 0
					do While numTGs < UBound(tgArr)+1
						if tgArr(numTGs)<>"0" then
							if firstTG then
								strSQL = strSQL & " AND ([VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
								firstTG = false
							else
								strSQL = strSQL & " OR [VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
							end if
						end if
						numTGs = numTGs + 1
					loop
					strSQL = strSQL & ") "
				end if
				strSQL = strSQL & " ORDER BY "
				strSQL = strSQL & GetTrnOrderBy()
			else	'pView="2" all asst1 & asst2
				strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.EmpID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName, tblPayRates.Rate1 FROM TRAINERS LEFT OUTER JOIN [VISIT DATA] ON TRAINERS.TrainerID = [VISIT DATA].TrainerID2 AND NOT ([VISIT DATA].TrainerID2 IS NULL) AND [VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & " LEFT OUTER JOIN [VISIT DATA] [VISIT DATA_1] ON TRAINERS.TrainerID = [VISIT DATA_1].TrainerID3 AND NOT ([VISIT DATA_1].TrainerID3 IS NULL) AND [VISIT DATA_1].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & " LEFT OUTER JOIN tblPayRates ON TRAINERS.TrainerID = tblPayRates.TrainerID WHERE (NOT ([VISIT DATA].TrainerID2 IS NULL)) OR (NOT ([VISIT DATA_1].TrainerID3 IS NULL)) "
				if inStr(tgChk, ",0,")>0 OR request.form("optTG")="" then
					''All TGs
				else
					firstTG = true
					numTGs = 0
					do While numTGs < UBound(tgArr)+1
						if tgArr(numTGs)<>"0" then
							if firstTG then
								strSQL = strSQL & " AND ([VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
								firstTG = false
							else
								strSQL = strSQL & " OR [VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
							end if
						end if
						numTGs = numTGs + 1
					loop
					strSQL = strSQL & ") "
				end if
				strSQL = strSQL & " ORDER BY "
				strSQL = strSQL & GetTrnOrderBy()
			end if
		end if
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing

		do while NOT rsEntry.EOF
			trnTotCls = 0
			trnTotVD = 0
			if request.form("optDetail")<>"2" then	'Detail View
				tmpTrnStr = "<span class=""mainTextBig"">" & rsEntry("TrLastName") & ",&nbsp;" & rsEntry("TrFirstName") & "</span>"
				if NOT isNull(rsEntry("EmpID")) then
					tmpTrnStr = tmpTrnStr & "&nbsp;&nbsp;&nbsp;&nbsp;ID:" & rsEntry("EmpID")
				end if
			else
				tmpTrnStr = rsEntry("TrLastName") & ",&nbsp;" & rsEntry("TrFirstName")
				if NOT isNull(rsEntry("EmpID")) then
					trnIDStr = rsEntry("EmpID")
				else
					trnIDStr = ""
				end if
			end if
%>
			<% if request.form("optDetail")<>"2" then ' detail view %>
                <tr>
                  <td align="left" valign="top" class="mainText"><strong><%=tmpTrnStr%></strong></td>
                  <td align="left" valign="bottom" class="smallTextBlack"><%= getHotWord(57)%></td>
                  <td class="smallTextBlack" valign="bottom" align="left"><%= getHotWord(58)%></td>
                  <td class="smallTextBlack" valign="bottom" align="left">Class&nbsp;Name</td>
				<% if session("useResrcResv") then %>
					  <td class="smallTextBlack" valign="bottom" align="left"><%= getHotWord(0)%></td>
				<% end if %>
                  <td  valign="bottom" class="smallTextBlack center-ch">#<%=session("ClientHW")%>s</td>
                  <td  valign="bottom" class="smallTextBlack center-ch">First Date</td>
                  <td  valign="bottom" class="smallTextBlack center-ch">Last Date</td>
                </tr>
				<%if NOT request.form("frmExpReport")="true" then%>
                <tr style="background-color:<%=session("pageColor4")%>;"><td colspan="8" height="1" width="100%"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td></tr>
				<%end if%>
			<% end if %>
<%
		if session("useResrcResv") then	'use resource
			if pView="0" then	'TrainerID2
				strSQL = "SELECT COUNT([VISIT DATA].VisitRefNo) AS NumVisits, [TRAINERS].TrainerID, [VISIT DATA].ClassDate, [VISIT DATA].ClassTime, [VISIT DATA].ClassID, tblClassDescriptions.ClassName, tblResources.ResourceName, FLDate.LastDate, FLDate.FirstDate "
				strSQL = strSQL & "FROM (SELECT ClassID, MAX(ClassDate) AS LastDate, MIN(ClassDate) AS FirstDate FROM tblClassSch WHERE (TrainerID2 = " & rsEntry("TrainerID") & ") OR (TrainerID3 = " & rsEntry("TrainerID") & ") GROUP BY ClassID) FLDate INNER JOIN tblClassSch INNER JOIN tblClassDescriptions ON tblClassSch.DescriptionID = tblClassDescriptions.ClassDescriptionID INNER JOIN [VISIT DATA] ON tblClassSch.ClassDate = [VISIT DATA].ClassDate AND tblClassSch.ClassID = [VISIT DATA].ClassID ON FLDate.ClassID = tblClassSch.ClassID INNER JOIN TRAINERS ON ([VISIT DATA].TrainerID3 = TRAINERS.TrainerID OR TRAINERS.TrainerID = [VISIT DATA].TrainerID2) "
				strSQL = strSQL & "AND [VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & " LEFT OUTER JOIN tblResources INNER JOIN tblResourceSchedules ON tblResources.ResourceID = tblResourceSchedules.ResourceID ON tblClassSch.ClassID = tblResourceSchedules.RefClass AND tblClassSch.ClassDate >= tblResourceSchedules.StartDate AND tblClassSch.ClassDate <= tblResourceSchedules.EndDate "
				strSQL = strSQL & " WHERE (1=1 "
				if inStr(tgChk, ",0,")>0 OR request.form("optTG")="" then
					''All TGs
				else
					firstTG = true
					numTGs = 0
					do While numTGs < UBound(tgArr)+1
						if tgArr(numTGs)<>"0" then
							if firstTG then
								strSQL = strSQL & " AND ([VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
								firstTG = false
							else
								strSQL = strSQL & " OR [VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
							end if
						end if
						numTGs = numTGs + 1
					loop
					strSQL = strSQL & ") "
				end if
				strSQL = strSQL & ") GROUP BY [VISIT DATA].ClassDate, [VISIT DATA].ClassID, [TRAINERS].TrainerID, tblClassDescriptions.ClassName, [VISIT DATA].ClassTime, tblResources.ResourceName, FLDate.LastDate, FLDate.FirstDate HAVING (NOT ([VISIT DATA].ClassID IS NULL)) AND ([TRAINERS].TrainerID = " & rsEntry("TrainerID") & ") ORDER BY [VISIT DATA].ClassDate, [VISIT DATA].ClassTime"
			else	'TrainerID3
				strSQL = "SELECT COUNT([VISIT DATA].VisitRefNo) AS NumVisits, [TRAINERS].TrainerID, [VISIT DATA].ClassDate, [VISIT DATA].ClassTime, [VISIT DATA].ClassID, tblClassDescriptions.ClassName, tblResources.ResourceName, FLDate.LastDate, FLDate.FirstDate "
				strSQL = strSQL & "FROM (SELECT ClassID, MAX(ClassDate) AS LastDate, MIN(ClassDate) AS FirstDate FROM tblClassSch WHERE (TrainerID2 = " & rsEntry("TrainerID") & ") OR (TrainerID3 = " & rsEntry("TrainerID") & ") GROUP BY ClassID) FLDate INNER JOIN tblClassSch INNER JOIN tblClassDescriptions ON tblClassSch.DescriptionID = tblClassDescriptions.ClassDescriptionID INNER JOIN [VISIT DATA] ON tblClassSch.ClassDate = [VISIT DATA].ClassDate AND tblClassSch.ClassID = [VISIT DATA].ClassID ON FLDate.ClassID = tblClassSch.ClassID INNER JOIN TRAINERS ON ([VISIT DATA].TrainerID3 = TRAINERS.TrainerID OR TRAINERS.TrainerID = [VISIT DATA].TrainerID2) "
				strSQL = strSQL & "AND [VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & " LEFT OUTER JOIN tblResources INNER JOIN tblResourceSchedules ON tblResources.ResourceID = tblResourceSchedules.ResourceID ON tblClassSch.ClassID = tblResourceSchedules.RefClass AND tblClassSch.ClassDate >= tblResourceSchedules.StartDate AND tblClassSch.ClassDate <= tblResourceSchedules.EndDate "
				strSQL = strSQL & " WHERE (1=1 "
				if inStr(tgChk, ",0,")>0 OR request.form("optTG")="" then
					''All TGs
				else
					firstTG = true
					numTGs = 0
					do While numTGs < UBound(tgArr)+1
						if tgArr(numTGs)<>"0" then
							if firstTG then
								strSQL = strSQL & " AND ([VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
								firstTG = false
							else
								strSQL = strSQL & " OR [VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
							end if
						end if
						numTGs = numTGs + 1
					loop
					strSQL = strSQL & ") "
				end if
				strSQL = strSQL & ") GROUP BY [VISIT DATA].ClassDate, [VISIT DATA].ClassID, [TRAINERS].TrainerID, tblClassDescriptions.ClassName, [VISIT DATA].ClassTime, tblResources.ResourceName, FLDate.LastDate, FLDate.FirstDate HAVING (NOT ([VISIT DATA].ClassID IS NULL)) AND ([TRAINERS].TrainerID = " & rsEntry("TrainerID") & ") ORDER BY [VISIT DATA].ClassDate, [VISIT DATA].ClassTime"
			end if
		else	'no resources
			if pView="0" then
				strSQL = "SELECT COUNT([VISIT DATA].VisitRefNo) AS NumVisits, [TRAINERS].TrainerID, [VISIT DATA].ClassDate, [VISIT DATA].ClassTime, [VISIT DATA].ClassID, tblClassDescriptions.ClassName, FLDate.LastDate, FLDate.FirstDate "
				strSQL = strSQL & "FROM (SELECT ClassID, MAX(ClassDate) AS LastDate, MIN(ClassDate) AS FirstDate FROM tblClassSch WHERE (TrainerID2 = " & rsEntry("TrainerID") & ") OR (TrainerID3 = " & rsEntry("TrainerID") & ") GROUP BY ClassID) FLDate INNER JOIN tblClassSch INNER JOIN tblClassDescriptions ON tblClassSch.DescriptionID = tblClassDescriptions.ClassDescriptionID INNER JOIN [VISIT DATA] ON tblClassSch.ClassDate = [VISIT DATA].ClassDate AND tblClassSch.ClassID = [VISIT DATA].ClassID ON FLDate.ClassID = tblClassSch.ClassID INNER JOIN TRAINERS ON ([VISIT DATA].TrainerID3 = TRAINERS.TrainerID OR TRAINERS.TrainerID = [VISIT DATA].TrainerID2) "
				strSQL = strSQL & "AND [VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & " LEFT OUTER JOIN tblResources INNER JOIN tblResourceSchedules ON tblResources.ResourceID = tblResourceSchedules.ResourceID ON tblClassSch.ClassID = tblResourceSchedules.RefClass AND tblClassSch.ClassDate >= tblResourceSchedules.StartDate AND tblClassSch.ClassDate <= tblResourceSchedules.EndDate "
				strSQL = strSQL & " WHERE (1=1 "
				if inStr(tgChk, ",0,")>0 OR request.form("optTG")="" then
					''All TGs
				else
					firstTG = true
					numTGs = 0
					do While numTGs < UBound(tgArr)+1
						if tgArr(numTGs)<>"0" then
							if firstTG then
								strSQL = strSQL & " AND ([VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
								firstTG = false
							else
								strSQL = strSQL & " OR [VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
							end if
						end if
						numTGs = numTGs + 1
					loop
					strSQL = strSQL & ") "
				end if
				strSQL = strSQL & ") GROUP BY [VISIT DATA].ClassDate, [VISIT DATA].ClassID, [TRAINERS].TrainerID, tblClassDescriptions.ClassName, [VISIT DATA].ClassTime, FLDate.LastDate, FLDate.FirstDate HAVING (NOT ([VISIT DATA].ClassID IS NULL)) AND ([TRAINERS].TrainerID = " & rsEntry("TrainerID") & ") ORDER BY [VISIT DATA].ClassDate, [VISIT DATA].ClassTime"
			else
				strSQL = "SELECT COUNT([VISIT DATA].VisitRefNo) AS NumVisits, [TRAINERS].TrainerID, [VISIT DATA].ClassDate, [VISIT DATA].ClassTime, [VISIT DATA].ClassID, tblClassDescriptions.ClassName, FLDate.LastDate, FLDate.FirstDate "
				strSQL = strSQL & "FROM (SELECT ClassID, MAX(ClassDate) AS LastDate, MIN(ClassDate) AS FirstDate FROM tblClassSch WHERE (TrainerID2 = " & rsEntry("TrainerID") & ") OR (TrainerID3 = " & rsEntry("TrainerID") & ") GROUP BY ClassID) FLDate INNER JOIN tblClassSch INNER JOIN tblClassDescriptions ON tblClassSch.DescriptionID = tblClassDescriptions.ClassDescriptionID INNER JOIN [VISIT DATA] ON tblClassSch.ClassDate = [VISIT DATA].ClassDate AND tblClassSch.ClassID = [VISIT DATA].ClassID ON FLDate.ClassID = tblClassSch.ClassID INNER JOIN TRAINERS ON ([VISIT DATA].TrainerID3 = TRAINERS.TrainerID OR TRAINERS.TrainerID = [VISIT DATA].TrainerID2) "
				strSQL = strSQL & "AND [VISIT DATA].ClassDate BETWEEN " & DateSep & cSDate & DateSep & " AND " & DateSep & cEDate & DateSep & " LEFT OUTER JOIN tblResources INNER JOIN tblResourceSchedules ON tblResources.ResourceID = tblResourceSchedules.ResourceID ON tblClassSch.ClassID = tblResourceSchedules.RefClass AND tblClassSch.ClassDate >= tblResourceSchedules.StartDate AND tblClassSch.ClassDate <= tblResourceSchedules.EndDate "
				strSQL = strSQL & " WHERE (1=1 "
				if inStr(tgChk, ",0,")>0 OR request.form("optTG")="" then
					''All TGs
				else
					firstTG = true
					numTGs = 0
					do While numTGs < UBound(tgArr)+1
						if tgArr(numTGs)<>"0" then
							if firstTG then
								strSQL = strSQL & " AND ([VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
								firstTG = false
							else
								strSQL = strSQL & " OR [VISIT DATA].TypeGroup=" & TRIM(tgArr(numTGs))
							end if
						end if
						numTGs = numTGs + 1
					loop
					strSQL = strSQL & ") "
				end if
				strSQL = strSQL & ") GROUP BY [VISIT DATA].ClassDate, [VISIT DATA].ClassID, [TRAINERS].TrainerID, tblClassDescriptions.ClassName, [VISIT DATA].ClassTime, FLDate.LastDate, FLDate.FirstDate HAVING (NOT ([VISIT DATA].ClassID IS NULL)) AND ([TRAINERS].TrainerID = " & rsEntry("TrainerID") & ") ORDER BY [VISIT DATA].ClassDate, [VISIT DATA].ClassTime"
			end if
		end if
response.write debugSQL(strSQL, "SQL")
		rsEntry2.CursorLocation = 3
		rsEntry2.open strSQL, cnWS
		Set rsEntry2.ActiveConnection = Nothing

		do while NOT rsEntry2.EOF
			if request.form("optDetail")<>"2" then ' summary view
				if rowCount=1 then
					rowCount = 2
					rowColor = "#FAFAFA"
				else
					rowCount = 1
					rowColor = "#F2F2F2"
				end if
%>
                <tr style="background-color:<%=rowColor%>;">
                  <td  valign="top" class="mainText center-ch">&nbsp;</td>
                  <td class="mainText" valign="top" align="left"><%=FmtDateShort(rsEntry2("ClassDate"))%>&nbsp;</td>
                  <td class="mainText" valign="top" align="left"><%if isNull(rsEntry2("ClassTime")) then response.Write "TBD" else response.Write FmtTimeShort(rsEntry2("ClassTime")) end if%>&nbsp;</td>
                  <td width="30%" class="mainText" valign="top" align="left"><%=rsEntry2("ClassName")%></td> 
			<% if session("useResrcResv") then %>
                  <td class="mainText" valign="top" align="left"><%=rsEntry2("ResourceName")%></td>
			<% end if %>
                  <td  valign="top" class="mainText center-ch"><%=rsEntry2("NumVisits")%></td>
                  <td  valign="top" class="mainText center-ch"><%if NOT isNULL(rsEntry2("FirstDate")) then response.write FmtDateShort(rsEntry2("FirstDate")) end if%></td>
                  <td  valign="top" class="mainText center-ch"><%if NOT isNULL(rsEntry2("LastDate")) then response.write FmtDateShort(rsEntry2("LastDate")) end if%></td>
                </tr>
				<%if NOT request.form("frmExpReport")="true" then%>
                <tr style="background-color:#CCCCCC;"><td colspan="8" height="1" width="100%"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td></tr>
				<%end if%>
<%
				end if 'summary view
				trnTotCls = trnTotCls + 1
				trnTotVD = trnTotVD + rsEntry2("NumVisits")
				rsEntry2.MoveNext
		loop
		rsEntry2.close
		'Trainer Totals
		if NOT isNULL(rsEntry("Rate1")) then
			totTrnPay = totTrnPay + rsEntry("Rate1")*trnTotCls
		end if
%>
			<% if request.form("optDetail")<>"2" then ' summary view %>
				<%if NOT request.form("frmExpReport")="true" then%>
                <tr style="background-color:<%=session("pageColor4")%>;"><td colspan="8" height="1" width="100%"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td></tr>
				<%end if%>
                <tr align="left">
                  <td colspan="3"><b>
				  #<%= getHotWord(5)%>: <%=trnTotCls%>&nbsp;&nbsp;&nbsp;#<%=session("ClientHW")%>s: <%=trnTotVD%>&nbsp;
<%
					if NOT isNULL(rsEntry("Rate1")) then
						if NOT request.form("frmExpReport")="true" then
							response.write "PayRate: " & FmtCurrency(rsEntry("Rate1")) & "&nbsp;&nbsp;"
						else 
							response.write "PayRate: " & rsEntry("Rate1")
						end if
					end if
%>
<%
					if NOT isNULL(rsEntry("Rate1")) then
						if NOT request.form("frmExpReport")="true" then
							response.write "<td colspan=""3"" align=""right""><b>TOTAL PAY: " & FmtCurrency(rsEntry("Rate1")*trnTotCls) & "&nbsp;&nbsp;</b></td>"
						else
							response.write "<td colspan=""3"" align=""right""><b>" & FmtNumber(rsEntry("Rate1")*trnTotCls) & "</b></td>"
						end if 
					end if
%>
                </tr>
              <tr><td colspan="8">&nbsp;</td></tr>
			<% else 'summary view%>
<%
				if rowCount=1 then
					rowCount = 2
					rowColor = "#FAFAFA"
				else
					rowCount = 1
					rowColor = "#F2F2F2"
				end if			
%>
                <tr style="background-color:<%=rowColor%>;">
                  <td width="1%" align="left" valign="top" class="mainText" colspan="1"><strong>&nbsp;<%if trnIDStr<>"" then response.write "ID:&nbsp;" end if%><%=trnIDStr%></strong></td>
                  <td align="left" valign="top" class="mainText" colspan="1"><strong>&nbsp;&nbsp;&nbsp;<%=tmpTrnStr%></strong></td>
                  <td  nowrap class="center-ch mainText"><strong><%=trnTotVD%></strong></td>
                  <td  nowrap class="center-ch mainText"><strong><%=trnTotCls%></strong></td>
                  <td  nowrap class="center-ch mainText"><strong>
				<%if NOT request.form("frmExpReport")="true" then%>
					<%=FmtCurrency(rsEntry("Rate1"))%></strong>
				<% else %>
					<%=FmtNumber(rsEntry("Rate1"))%></strong>
				<% end if %>
				  </td>
                  <td  nowrap class="right mainText"><strong>
  					<%if NOT isNULL(rsEntry("Rate1")) then%>
						<%if NOT request.form("frmExpReport")="true" then%>
						  <%=FmtCurrency(rsEntry("Rate1")*trnTotCls)%>
						<% else %>
						  <%=FmtNumber(rsEntry("Rate1")*trnTotCls)%>
						<% end if %>
					<% else %>
						<%if NOT request.form("frmExpReport")="true" then%>
							<%=FmtCurrency(0)%>&nbsp;&nbsp;
						<% else %>
							<%=FmtNumber(0)%>
						<% end if %>
					<% end if %>
				  </strong></td>
				</tr>
				<%if NOT request.form("frmExpReport")="true" then%>
                <tr style="background-color:#CCCCCC;"><td colspan="8" height="1" width="100%"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td></tr>
				<%end if%>
			<% end if %>
<%
			rsEntry.MoveNext
		loop
		rsEntry.close
		if totTrnPay > 0 then
%>
                <tr>
                  <td colspan="8" width="100%">&nbsp;</td>
                </tr>
				<%if NOT request.form("frmExpReport")="true" then%>
                <tr style="background-color:<%=session("pageColor4")%>;"><td colspan="8"width="100%"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="2" width="100%"></td></tr>
				<%end if%>
                <tr >
                  	<td width="100%" height="1" colspan="8" nowrap class="mainTextBig">
					<% if NOT request.form("frmExpReport")="true" then %>		
				  		<b>GRAND TOTAL: <%=FmtCurrency(totTrnPay)%></b>
					<% else %>
				  		<b>GRAND TOTAL: <%=FmtNumber(totTrnPay)%></b>
					<% end if %>
					</td>
                </tr>
				<%if NOT request.form("frmExpReport")="true" then%>
                <tr style="background-color:<%=session("pageColor4")%>;"><td colspan="8" height="1" width="100%"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="2" width="100%"></td></tr>
				<%end if%>
<%		
		end if
%>
              </table>
	<% end if %>
	<% if NOT request.form("frmExpReport")="true" then %>
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
