<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
Server.ScriptTimeout = 300    '5 min (value in seconds)

if request.querystring("pdf") = "true" Then
'  if request.form("sid")<>"" then
'    Response.Cookies("SessionFarmGUID") = request.form("sid")
'  end if
end if

'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
	 dim rsEntry, rsEntry2, rsTipA, rsTipB, rsTipDayA, rsTipDayB, ss_AP_TodayOnly, ss_AP_DateRange
	 set rsEntry = Server.CreateObject("ADODB.Recordset")
	 set rsEntry2 = Server.CreateObject("ADODB.Recordset")
	 set rsTipA = Server.CreateObject("ADODB.Recordset")
	 set rsTipB = Server.CreateObject("ADODB.Recordset")
	 set rsTipDayA = Server.CreateObject("ADODB.Recordset")
	 set rsTipDayB = Server.CreateObject("ADODB.Recordset")
	 Dim useTips, includeTips
%>
	<!-- #include file="inc_rpt_pdf.asp" -->
	<!-- #include file="inc_accpriv.asp" -->
	<!-- #include file="inc_rpt_tagging.asp" -->
	<!-- #include file="inc_utilities.asp" -->
	<!-- #include file="inc_rpt_save.asp" -->
	<!-- #include file="inc_hotword.asp" -->
	
	<% if not Session("Pass") OR Session("Admin")="false" OR (NOT validAccessPriv("RPT_CASH_DATES") AND NOT validAccessPriv("RPT_CASHDRAW")) then %>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
	<% else %>
		<!-- #include file="../inc_i18n.asp" -->
		<%
		Dim exportExcel, cSDate, cEDate, cLoc, tmpCurDate, tmpCurEmpID, tmpCurPayType, tmpCurPayTypeName, cont, useSplit, rsSaleID, rsSaleDate, rsEmpID, rsPayMethID, rsPayMeth, rsClient, rsDesc, rsLoc, rsTax, rsPaid, rsCheckNo
		Dim tmpPMTotal, rowColor, dayTotCash, dayTotCheck, dayTotCredit, dayTotOther, TotalCash, TotalCheck, TotalCredit, TotalOther, dayPMRecCount, dayRecCount, TotalRecCount, cashBasis, curSplit, splitFactor, cSTime, cETime, cSDate2, cEDate2, cSDateTime, cEDateTime
		Dim TotalTip, TotalCashTip, TotalCheckTip, TotalCreditTip, TotalOtherTip, dayTotalTip, dayTotalCashTip, dayTotalCheckTip, dayTotalCreditTip, dayTotalOtherTip, byEmployee, ap_view_all_locs, useSaleTimeS, useSaleTimeE

		useSaleTimeS = false
		useSaleTimeE = false
		useTips = checkStudioSetting("tblGenOpts","useTips")
		includeTips = checkStudioSetting("tblGenOpts","IncludeTipsInPayroll")
		ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
		ss_AP_DateRange = validAccessPriv("RPT_CASH_DATES")
		if not ss_AP_DateRange then
	      ss_AP_TodayOnly = validAccessPriv("RPT_CASHDRAW")
		end if

		if request.form("requiredtxtDateStart")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cSDate = CDATE(request.form("requiredtxtDateStart"))
				cSDate2 = CDATE(request.form("requiredtxtDateStart"))
			Call SetLocale("en-us")
		else
			cSDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
			cSDate2 = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		end if
		if request.form("requiredtxtDateEnd")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cEDate = CDATE(request.form("requiredtxtDateEnd"))
				'cEDate2 = CDATE(request.form("requiredtxtDateEnd"))
			Call SetLocale("en-us")
		else
			cEDate = DateValue(cSDate)
			'cEDate2 = DateValue(cSDate)
		end if
		
		dim STimeField, ETimeField
		STimeField = ""
		ETimeField = ""
		'Set up cSDateTime
		if request.form("optStartTime") = "-1" then
			STimeField = "SaleDate"
			cSDateTime = cSDate
		else
			useSaleTimeS = true
			STimeField = "SaleTime"
			cSDateTime = cSDate & " " & request.form("optStartTime")
		end if
		
		'Set up cEDateTime
		if request.form("optEndTime") = "-1" then
			ETimeField = "SaleDate"
			cEDateTime = cEDate
		else
			useSaleTimeE = true
			ETimeField = "SaleTime"
			cEDateTime = cEDate & " " & request.form("optEndTime")
		end if


		If request.form("optSaleLoc")<>"" then
			cLoc = CINT(request.form("optSaleLoc"))
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
		if request.form("optByEmployee")="on" then
			byEmployee = true
		else
			byEmployee = false
		end if
		cashBasis = true
		if NOT request.form("frmExpReport")="true" then
		%>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_cashdrawer")) %>
			<script type="text/javascript">
			function genReport() {
				document.frmSales.frmGenReport.value = "true";
				document.frmSales.frmExpReport.value = "false";
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
				else%>
				document.frmSales.frmGenPdf.value = "false"; 
				<% end if %>
				document.frmSales.submit();
			}
			
			<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
				else%>
			function exportReport() {
				document.frmSales.frmExpReport.value = "true";
				document.frmSales.frmGenReport.value = "true";
				document.frmSales.frmGenPdf.value = "false";
				<% iframeSubmit "frmSales", "adm_rpt_cashdrawer.asp" %>
			}
			<% end if %>
			
			</script>		
			<!-- #include file="inc_help_content.asp" -->
			
			<!-- #include file="../inc_date_ctrl.asp" -->
			<!-- #include file="../inc_ajax.asp" -->
			<!-- #include file="../inc_val_date.asp" -->
		<%
		end if
		if NOT request.form("frmExpReport")="true" then %>
<%= pageStart %>
			<table width="<%=strPageWidth%>" cellspacing="0">
			<tr> 
				<td valign="top"> 
				<table cellspacing="0" width="100%" style="margin: 0 auto;">
					<tr>
						<td class="headText" align="left" valign="top">
						<table width="100%" cellspacing="0">
							<tr>
								<td width="3%" height="25" valign="bottom" class="headText"><b> </b> </td>
								<td width="97%" valign="bottom" class="headText"><b><%= pp_PageTitle("Cash Drawer") %> </b>
									<%if session("Admin")="sa" then %>
									<a class="mainText" href="/Report/Sales/CashDrawer">Current version</a>
									<%end if %>
									<!--JM - 48_2448-->
									<% showTrainingMovieIcon("21024123-the-cash-drawer-report") %>
									<!--JM - 49_2447-->
									<% showNewHelpContentIcon("cash-drawer-report") %>

								</td>
							</tr>
						</table>
						</td>
					</tr>
					<tr> 
						<td height="30" valign="bottom" class="headText">
						<table class="mainText border4 center-block" cellspacing="0">
							<form name="frmSales" action="adm_rpt_cashdrawer.asp" method="POST">
							<input type="hidden" name="frmGenReport" value="">
							<input type="hidden" name="frmExpReport" value="">
							<tr> 
								<td width="90%" valign="middle" nowrap style="background-color:#F2F2F2;" class="center-ch">
                                                                  <b>&nbsp;
                                                                    <% if NOT ss_AP_TodayOnly then response.write "Start " end if %> <%=xssStr(allHotWords(57))%>: 
                                                                    <input onBlur="validateDate(this, '<%=FmtDateShort(cSDate2)%>', true);" type="text" name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate2)%>" class="date" <% if ss_AP_TodayOnly then response.write(" disabled") end if %>>
																						  <% if NOT ss_AP_TodayOnly then %>
																		<script type="text/javascript">
																			var cal1 = new tcal({'formname':'frmSales', 'controlname':'requiredtxtDateStart'});
																			cal1.a_tpl.yearscroll = true;
																		</script>
																						  <% end if %>
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
                                                                    <% if ss_AP_DateRange then %>
                                                                      &nbsp;&nbsp;through&nbsp;&nbsp;<br />
                                                                      <%=xssStr(allHotWords(79))%> 
                                                                      <input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
																		<script type="text/javascript">
																			var cal2 = new tcal({'formname':'frmSales', 'controlname':'requiredtxtDateEnd'});
																			cal2.a_tpl.yearscroll = true;
																		</script>
                                                                    <% else 'BJD 3/24/08: added hidden form var to allow query to behave the same in all cases %>
                                                                      <input type="hidden" name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>">
                                                                    <% end if %>
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
                                                                    &nbsp;&nbsp;<% taggingFilter %>
							<%	if checkStudioSetting("tblGenOpts", "TrackCashRegisters") then 					
									strSQL = "SELECT CashRegisterName, CashRegisterID FROM tblCashRegister WHERE [Delete] = 0 "
									if cLoc<>0 AND cLoc<>98 then
										strSQL = strSQL & " AND LocationID = " & cLoc
									end if
								response.write debugSQL(strSQL, "SQL")
									rsEntry2.CursorLocation = 3
									rsEntry2.open strSQL, cnWS
									Set rsEntry2.ActiveConnection = Nothing %>
								
											<select name="optCashRegister">
												<option value="-1">All Cash Registers</option>
							
							<%		do while not rsEntry2.EOF	
							%>			
												<option value="<%=rsEntry2("CashRegisterID")%>"<% if rsEntry2("CashRegisterID")=cint(request.form("optCashRegister")) then response.write " selected" end if %>><%=rsEntry2("CashRegisterName")%></option>
							<%
										rsEntry2.MoveNext
									loop
							%>			
										</select>&nbsp;&nbsp;
							<%	
									rsEntry2.close
								end if %>
									
									<select name="optSaleLoc" onChange="document.frmSales.submit();" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
									<option value="0" <%if cLoc=0 then response.write "selected" end if %>>All</option>
									<option value="98" <%if cLoc=98 then response.write "selected" end if%>>Online Store</option>
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
									document.frmSales.optSaleLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
								</script>					  
											
									By Employee:&nbsp;<input type="checkbox" name="optByEmployee" <%if byEmployee then response.write "checked" end if%>>
									&nbsp;<select name="optEmployee">
									<option value="-1" <%if request.form("optEmployee")="-1" then response.write "selected" end if%>>All Employees</option>
									<option value="0" <%if request.form("optEmployee")="0" then response.write "selected" end if%>>Owner</option>
								<%
									strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.SmodeID, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName FROM TRAINERS INNER JOIN Sales ON TRAINERS.TrainerID = Sales.EmployeeID INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID "
									strSQL = strSQL & "WHERE (Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ")"
									if cLoc<>0 then
										strSQL = strSQL & " AND [Sales Details].Location=" & cLoc
									end if
									strSQL = strSQL & " ORDER BY " & GetTrnOrderBy()
									rsEntry.CursorLocation = 3
									rsEntry.open strSQL, cnWS
									Set rsEntry.ActiveConnection = Nothing

									do while NOT rsEntry.EOF
								%>
									<option value="<%=rsEntry("TrainerID")%>" <%if request.form("optEmployee")=CSTR(rsEntry("TrainerID")) then response.write "selected" end if%>><%=FmtTrnNameNew(rsEntry, true)%></option>
								<%
										rsEntry.MoveNext
									loop
									rsEntry.close
								%>
								</select> &nbsp;&nbsp;
								<br />&nbsp;
								<input type="button" name="Button" value="Generate" onClick="genReport();">
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
				else%>
						<% exportToExcelButton %><% pdfExportButton "frmSales", "CashDrawer_" & Replace(cSDate, "/", "-") & "_to_" & Replace(cEDate, "/", "-") & ".pdf" %>
				<%end if%>
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
				 else%>
						<% taggingButtons("frmSales") %>
				<%end if%>
						<% savingButtons "frmSales", "Cash Drawer" %>
								</td>
							</tr>
							</form>
						</table>			
						</td>
					</tr>
					
					<tr> 
						<td valign="top" class="mainTextBig"> 
						<table class="mainText" width="95%" cellspacing="0" style="margin: 0 auto;">
							<tr> 
								<td class="mainTextBig" colspan="2" valign="top">
								<table width="100%" cellspacing="0" class="mainText">
									<tr>
										<td valign="top">
		<% end if	%>
		<table class="mainText" width="100%"  cellspacing="0">
		                  <tr>
			                  <td colspan="18"> <!-- Header for exports -->
                <%        dIM locName
                          if cLoc = 0 then
                            locName = "All Locations"
                          elseif cLoc = 98 then
                            locName = "Online Store"
                          else
                            strSQL = "SELECT LocationName FROM Location WHERE LocationID =  " & cLoc
		                    rsEntry.CursorLocation = 3
		                    rsEntry.open strSQL, cnWS
		                    Set rsEntry.ActiveConnection = Nothing

		                    if NOT rsEntry.EOF then
                                locName = rsEntry("LocationName")
                            end if
                            rsEntry.close
                          end if
                          dim TrnName, trid
                          if request.form("optEmployee") = "-1" then
                            TrnName = "All Employees"
                          elseif request.form("optEmployee") = "0" then
                            TrnName = "Owner"
                          else
                            trid = CDBL(request.form("optEmployee"))
                            strSQL = "SELECT TRAINERS.TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.SmodeID, TRAINERS.DisplayName FROM TRAINERS "
						    strSQL = strSQL & "WHERE TrainerID = " & trid
						    rsEntry.CursorLocation = 3
						    rsEntry.open strSQL, cnWS
						    Set rsEntry.ActiveConnection = Nothing

						    if NOT rsEntry.EOF then
							    TrnName = FmtTrnNameNew(rsEntry, true)
						    end if
						    rsEntry.close
                          end if
                          
                          if request.Form("frmExpReport")="true" then 
                %>          <strong><%=FmtDateShort(cSDate)%> - <%=FmtDateShort(cEDate)%>&nbsp;&nbsp;
                            &nbsp;&nbsp;<%=locName%>
                            &nbsp;&nbsp;<%=TrnName%></strong>
                <%        end if 
                %>        &nbsp;
                        </td>
			                </tr>
			                </table>			  
		<table class="mainText" width="100%"  cellspacing="0">
		<%
				if request.form("frmTagClients")="true" then
					strSQL = "SELECT Sales.ClientID "
					strSQL = strSQL & "FROM Sales "
					strSQL = strSQL & " INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID "
					strSQL = strSQL & " INNER JOIN tblPayments ON tblPayments.SaleID = Sales.SaleID "
					strSQL = strSQL & " INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID "
					strSQL = strSQL & " INNER JOIN [Payment Types] ON tblPayments.PaymentMethod = [Payment Types].[Item#] "
					if request.form("optFilterTagged")="on" then
						strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
						if session("mVarUserID")<>"" then
							strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
						end if
						strSQL = strSQL & " ) "
					end if
					'BJD 3/24/08: made query incorporate start/end dates & times in all cases
					'if ss_AP_DateRange then
						strSQL = strSQL & "WHERE (Sales.SaleDate >= " & DateSep & cdate(cSDate) & DateSep & ") AND (Sales.SaleDate <= " & DateSep & cstr(cEDate) & DateSep & ") "
						' ADDED 3/10/08 by Brad
						' Added a check vs the time ... either we're past the starting day (at any time) OR we're on the starting day of the range AND past the starting time
						' In other words, edge dates of the range are the only days that the time matters.
						if useSaleTimeS then
							strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate > " & DateSep & cdate(cSDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102)  >= " & TimeSepB & request.form("optStartTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
						end if
		
						' ADDED 3/10/08 by Brad
						' Added a check vs the time ... either we're before the ending day (at any time) OR we're on the ending day of the range AND before the ending time
						' In other words, edge dates of the range are the only days that the time matters.
						if useSaleTimeE then
							strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate < " & DateSep & cdate(cEDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102) <= " & TimeSepB & request.form("optEndTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
						end if
					'else
					'	strSQL = strSQL & "WHERE (Sales.SaleDate= " & DateSep & cdate(cSDate2) & DateSep & ") "
					'end if
					if request.form("optCashRegister")<>"" and request.form("optCashRegister")<>"-1" then
						strSQL = strSQL & "  AND CashRegisterID = " & request.form("optCashRegister")
					end if
					strSQL = strSQL & " AND [Payment Types].[CashEQ]=1 "
					if cLoc<>0 then
						strSQL = strSQL & " AND [Sales Details].Location=" & cLoc & " "
					end if
					if request.form("optEmployee")<>"-1" then
						strSQL = strSQL & " AND [Sales].EmployeeID=" & request.form("optEmployee") & " "
					end if
				'response.write "<br /><br />" & strSQL
				
					if request.form("frmTagClientsNew")="true" then
						clearAndTagQuery(strSQL)
					else
						tagQuery(strSQL)
					end if
				
				end if
			
			if request.form("frmGenReport")="true" then 
				if request.form("frmExpReport")="true" then
					Dim stFilename
					stFilename="attachment; filename=CashDrawer " & Replace(cSDate,"/","-") & ".xls" 
					Response.ContentType = "application/vnd.ms-excel" 
					Response.AddHeader "Content-Disposition", stFilename 
				end if
			
			'NOT 1st Load''''
			'Query for Reg Sales
			'CB 6_12_07 Changed join to tblEFTSchedule to sub query as a sale can have more than one record in tblEFTSchedule causing multiple records to get returned
			'CB 6_12_07 Now uses sub query on tblEFTSchedule (AutoPay) that groups by the sale ID such that each sale is only returned once
			strSQL = "SELECT Sales.SaleID, Sales.ClientID, [Sales Details].Location, SUM(tblSDPayments.SDPaymentAmount) AS PmtAmt, Sales.EmployeeID, Sales.SaleDate, Sales.SaleTime, CLIENTS.LastName, CLIENTS.FirstName, [Payment Types].PmtTypes, [Payment Types].Item#, Location.LocationName, AutoPays.EFTSaleID, tblPayments.PaymentNotes "

			'strSQL = strSQL & "FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID INNER JOIN [Payment Types] ON Sales.PaymentMethod = [Payment Types].Item# INNER JOIN Location ON [Sales Details].Location = Location.LocationID LEFT OUTER JOIN (SELECT SaleID AS EFTSaleID FROM tblEFTSchedule WHERE (RunAtPOS = 0) GROUP BY SaleID) AutoPays ON Sales.SaleID = AutoPays.EFTSaleID "
            'CB 54_3168 - New Sales Tables
            strSQL = strSQL & "FROM tblSDPayments INNER JOIN Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID INNER JOIN Location ON [Sales Details].Location = Location.LocationID ON tblSDPayments.SDID = [Sales Details].SDID INNER JOIN [Payment Types] INNER JOIN tblPayments ON [Payment Types].Item# = tblPayments.PaymentMethod ON tblSDPayments.PaymentID = tblPayments.PaymentID LEFT OUTER JOIN (SELECT SaleID AS EFTSaleID FROM tblEFTSchedule WHERE (RunAtPOS = 0) GROUP BY SaleID) AS AutoPays ON Sales.SaleID = AutoPays.EFTSaleID "

			if request.form("optFilterTagged")="on" then
				strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
			end if
			if request.form("optFilterTagged")="on" then
				if session("mvaruserID")<>"" then
					strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
				else
					strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
				end if
			end if
			strSQL = strSQL & "GROUP BY Sales.SaleID, Sales.ClientID, [Sales Details].Location, Sales.EmployeeID, Sales.SaleDate, Sales.SaleTime, CLIENTS.LastName, CLIENTS.FirstName, [Payment Types].PmtTypes, [Payment Types].Item#, Location.LocationName, [Payment Types].CashEQ, Sales.CashRegisterID, AutoPays.EFTSaleID, tblPayments.PaymentNotes "
			' ([Sales Details].CategoryID <> 21) AND 
			'BJD 3/24/08: made query incorporate start/end dates & times in all cases
			'if ss_AP_DateRange then
				strSQL = strSQL & "HAVING (Sales.SaleDate >= " & DateSep & cdate(cSDate) & DateSep & ") AND (Sales.SaleDate <= " & DateSep & cstr(cEDate) & DateSep & ") "

				' ADDED 3/10/08 by Brad
				' Added a check vs the time ... either we're past the starting day (at any time) OR we're on the starting day of the range AND past the starting time
				' In other words, edge dates of the range are the only days that the time matters.
				if useSaleTimeS then
					strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate > " & DateSep & cdate(cSDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102)  >= " & TimeSepB & request.form("optStartTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
				end if

				' ADDED 3/10/08 by Brad
				' Added a check vs the time ... either we're before the ending day (at any time) OR we're on the ending day of the range AND before the ending time
				' In other words, edge dates of the range are the only days that the time matters.
				if useSaleTimeE then
					strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate < " & DateSep & cdate(cEDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102) <= " & TimeSepB & request.form("optEndTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
				end if
			'else
				'strSQL = strSQL & "HAVING (Sales.SaleDate= " & DateSep & cdate(cSDate2) & DateSep & ") "
			'end if
			strSQL = strSQL & "AND (AutoPays.EFTSaleID IS NULL) "
			strSQL = strSQL & "AND [Payment Types].[CashEQ]=1 "
			if request.form("optCashRegister")<>"" and request.form("optCashRegister")<>"-1" then
				strSQL = strSQL & "  AND CashRegisterID = " & request.form("optCashRegister")
			end if
			if cLoc<>0 then
				strSQL = strSQL & " AND [Sales Details].Location=" & cLoc & " "
			end if
			if request.form("optEmployee")<>"-1" then
				strSQL = strSQL & " AND [Sales].EmployeeID=" & request.form("optEmployee") & " "
			end if
			if byEmployee then
				strSQL = strSQL & " ORDER BY Sales.SaleDate, Sales.EmployeeID, [Payment Types].[Item#], Sales.SaleTime, CLIENTS.LastName"
			else
				strSQL = strSQL & " ORDER BY Sales.SaleDate, [Payment Types].[Item#], Sales.SaleTime, CLIENTS.LastName"
			end if
			response.write debugSQL(strSQL, "SQL")
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing


            'CB 54_3168 - New Sales Tables
			strSQL = "SELECT '-1' AS SaleID WHERE 1=0"
			
			rsEntry2.CursorLocation = 3
			rsEntry2.open strSQL, cnWS
			Set rsEntry2.ActiveConnection = Nothing

			If useTips then
				'Query for Tip Totals Primary Payment Method
				strSQL = "SELECT [Sales Details].CategoryID, [Sales Details].Location, [Sales].EmployeeID, SUM(tblSDPayments.SDPaymentAmount) AS TipAmt, Sales.SaleDate, [Payment Types].PmtTypes, [Payment Types].Item#, [Payment Types].[CashEQ] "

				'strSQL = strSQL & "FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN [Payment Types] ON Sales.PaymentMethod = [Payment Types].Item# "
                'CB 54_3168 - New Sales Tables
                strSQL = strSQL & "FROM tblSDPayments INNER JOIN Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID ON tblSDPayments.SDID = [Sales Details].SDID INNER JOIN [Payment Types] INNER JOIN tblPayments ON [Payment Types].Item# = tblPayments.PaymentMethod ON tblSDPayments.PaymentID = tblPayments.PaymentID "

				strSQL = strSQL & "GROUP BY [Sales Details].CategoryID, [Sales Details].Location, [Sales].EmployeeID, Sales.SaleDate, Sales.SaleTime, [Payment Types].PmtTypes, [Payment Types].Item#, [Payment Types].[CashEQ], Sales.CashRegisterID "
				strSQL = strSQL & "HAVING ([Sales Details].CategoryID = 21) "
				'if ss_AP_DateRange then
					strSQL = strSQL & "AND (Sales.SaleDate >= " & DateSep & cdate(cSDate) & DateSep & ") AND (Sales.SaleDate <= " & DateSep & cstr(cEDate) & DateSep & ") "
	
					' ADDED 3/10/08 by Brad
					' Added a check vs the time ... either we're past the starting day (at any time) OR we're on the starting day of the range AND past the starting time
					' In other words, edge dates of the range are the only days that the time matters.
					if useSaleTimeS then
						strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate > " & DateSep & cdate(cSDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102)  >= " & TimeSepB & request.form("optStartTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
					end if
	
					' ADDED 3/10/08 by Brad
					' Added a check vs the time ... either we're before the ending day (at any time) OR we're on the ending day of the range AND before the ending time
					' In other words, edge dates of the range are the only days that the time matters.
					if useSaleTimeE then
						strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate < " & DateSep & cdate(cEDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102) <= " & TimeSepB & request.form("optEndTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
					end if
				'else
				'	strSQL = strSQL & "AND (Sales.SaleDate= " & DateSep & cdate(cSDate2) & DateSep & ") "
				'end if
				if request.form("optCashRegister")<>"" and request.form("optCashRegister")<>"-1" then
					strSQL = strSQL & "  AND CashRegisterID = " & request.form("optCashRegister")
				end if
				strSQL = strSQL & " AND [Payment Types].[CashEQ]=1 "
				if cLoc<>0 then
					strSQL = strSQL & " AND [Sales Details].Location=" & cLoc & " "
				end if
				if request.form("optEmployee")<>"-1" then
					strSQL = strSQL & " AND [Sales].EmployeeID=" & request.form("optEmployee") & " "
				end if
			    'response.write "<br /><br />" & strSQL
				rsTipA.CursorLocation = 3
				rsTipA.open strSQL, cnWS
				Set rsTipA.ActiveConnection = Nothing

				
			
                'CB 54_3168 - New Sales Tables
			    strSQL = "SELECT '-1' AS SaleID WHERE 1=0"
		
				rsTipB.CursorLocation = 3
				rsTipB.open strSQL, cnWS
				Set rsTipB.ActiveConnection = Nothing

			end if	'end of useTips check
			
			tmpCurDate = dateadd("d", -1, cSDate)
			tmpCurEmpID = -2222
			tmpCurPayType = ""
			tmpCurPayTypeName = ""
			dayTotCash = 0
			dayTotCheck = 0
			dayTotCredit = 0
			dayTotOther = 0
			TotalCash = 0
			TotalCheck = 0
			TotalCredit = 0
			TotalOther = 0
			dayPMRecCount = 0
			dayRecCount = 0
			TotalRecCount = 0
			tmpPMTotal	= 0
		
			dayTotalTip = 0
			dayTotalCashTip = 0
			dayTotalCheckTip = 0
			dayTotalCreditTip = 0
			dayTotalOtherTip = 0
			TotalTip = 0
			TotalCashTip = 0
			TotalCheckTip = 0
			TotalCreditTip = 0
			TotalOtherTip = 0
		
			If false then   ' if useTips
			'**********  CALCULATION OF TIP TOTALS SK 02/09/05  *******************
				if rsTipA.EOF then
				else
					Do Until rsTipA.EOF
						if rsTipA("Item#") = 1 then
							dayTotalCashTip = dayTotalCashTip + rsTipA("TipAmt")
							dayTotalTip=dayTotalTip+rsTipA("TipAmt")
							TotalCashTip = TotalCashTip + rsTipA("TipAmt")
							TotalTip=TotalTip+rsTipA("TipAmt")
						elseif rsTipA("Item#") = 2 then
							dayTotalCheckTip = dayTotalCheckTip + rsTipA("TipAmt")
							dayTotalTip=dayTotalTip+rsTipA("TipAmt")
							TotalCheckTip = TotalCheckTip + rsTipA("TipAmt")
							TotalTip=TotalTip+rsTipA("TipAmt")
						elseif rsTipA("Item#") >= 3 AND rsTipA("Item#") <= 6 then
							dayTotalCreditTip = dayTotalCreditTip + rsTipA("TipAmt")
							dayTotalTip=dayTotalTip+rsTipA("TipAmt")
							TotalCreditTip = TotalCreditTip + rsTipA("TipAmt")
							TotalTip=TotalTip+rsTipA("TipAmt")
						else
							dayTotalOtherTip = dayTotalOtherTip + rsTipA("TipAmt")		
							dayTotalTip=dayTotalTip+rsTipA("TipAmt")
							TotalOtherTip = TotalOtherTip + rsTipA("TipAmt")		
							TotalTip=TotalTip+rsTipA("TipAmt")
						end if
						rsTipA.movenext
					Loop
				end if
			
				if rsTipB.EOF then
				else
					Do Until rsTipB.EOF
						if rsTipB("Item#") = 1 then
							dayTotalCashTip = dayTotalCashTip + rsTipB("TipAmt")
							dayTotalTip=dayTotalTip+rsTipB("TipAmt")
							TotalCashTip = TotalCashTip + rsTipB("TipAmt")
							TotalTip=TotalTip+rsTipB("TipAmt")
						elseif rsTipB("Item#") = 2 then
							dayTotalCheckTip = dayTotalCheckTip + rsTipB("TipAmt")
							dayTotalTip=dayTotalTip+rsTipB("TipAmt")
							TotalCheckTip = TotalCheckTip + rsTipB("TipAmt")
							TotalTip=TotalTip+rsTipB("TipAmt")
						elseif rsTipB("Item#") >= 3 AND rsTipB("Item#") <= 6 then
							dayTotalCreditTip = dayTotalCreditTip + rsTipB("TipAmt")
							dayTotalTip=dayTotalTip+rsTipB("TipAmt")
							TotalCreditTip = TotalCreditTip + rsTipB("TipAmt")
							TotalTip=TotalTip+rsTipB("TipAmt")
						else
							dayTotalOtherTip = dayTotalOtherTip + rsTipB("TipAmt")		
							dayTotalTip=dayTotalTip+rsTipB("TipAmt")
							TotalOtherTip = TotalOtherTip + rsTipB("TipAmt")		
							TotalTip=TotalTip+rsTipB("TipAmt")
						end if
						rsTipB.movenext
					Loop
				end if
				'TotalCashTip = TotalCashTip - TotalCheckTip - TotalCreditTip - TotalOtherTip
				'TotalTip =	 TotalCashTip + TotalCheckTip + TotalCreditTip + TotalOtherTip
				rsTipA.close
				rsTipB.close
			'********************************************************************
			end if	'end of useTips check
			
			if rsEntry.EOF AND rsEntry2.EOF then
				cont = false
			else
				cont = true
			end if
		
			Do While cont
				'Determine rs to work with and pop local vars
				if NOT rsEntry.EOF AND NOT rsEntry2.EOF then
					' 04/03/06 - Removed the date criteria from the or/and clause and made it its own exterior loop
					' Code left below:
						' 02/08/06 Added the OR clause to correct the case when recordset2 was on a different date,
						' but not a different payment type - this groups by the highest order (date) first, THEN by
						' payment type, which is by design - old code left commented below
						'if (rsEntry2("Item#")<rsEntry("Item#")) then
						'if (rsEntry2("Item#")<rsEntry("Item#")) AND (rsEntry2("SaleDate") <= rsEntry("SaleDate")) then
							'useSplit = true
						'else
							'useSplit = false
						'end if
					if (rsEntry2("SaleDate") < rsEntry("SaleDate")) then
						useSplit = true
					else 
					    if byEmployee then
						    if (rsEntry2("Item#")<rsEntry("Item#")) AND (CLNG(rsEntry2("EmployeeID"))=CLNG(rsEntry("EmployeeID"))) AND (rsEntry2("SaleDate") = rsEntry("SaleDate")) then
							    useSplit = true
						    else
							    useSplit = false
						    end if					    
					    else
						    if (rsEntry2("Item#")<rsEntry("Item#")) AND (rsEntry2("SaleDate") = rsEntry("SaleDate")) then
							    useSplit = true
						    else
							    useSplit = false
						    end if
						end if
					end if
				elseif NOT rsEntry.EOF then
						useSplit = false
				else	''NOT rsEntry2.EOF
					useSplit = true
				end if
		
				curSplit = false
				'''Populate local Vars
				if NOT useSplit then
					rsSaleID = rsEntry("SaleID")
					rsSaleDate = rsEntry("SaleDate")
					if NOT isNULL(rsEntry("EmployeeID")) then
						rsEmpID = rsEntry("EmployeeID")
					else
						rsEmpID = 0
					end if
					rsPayMethID = rsEntry("Item#")
					rsCheckNo=Null
					if NOT isNULL(rsEntry("PaymentNotes")) then
						if rsEntry("PaymentNotes")<>"0" then
							rsCheckNo=rsEntry("PaymentNotes")
						end if
					end if
					rsPayMeth = rsEntry("PmtTypes")
					rsClient = TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))
					rsLoc = rsEntry("LocationName")
					rsPaid = rsEntry("PmtAmt")
		
					rsEntry.MoveNext
				else	''Split Sale
					rsSaleID = rsEntry2("SaleID")
					rsSaleDate = rsEntry2("SaleDate")
					if NOT isNULL(rsEntry2("EmployeeID")) then
						rsEmpID = rsEntry2("EmployeeID")
					else
						rsEmpID = 0
					end if
					rsPayMethID = rsEntry2("Item#")
					rsCheckNo=Null
					if NOT isNULL(rsEntry2("PaymentNotes")) then
						if rsEntry2("PaymentNotes")<>"0" then
							rsCheckNo=rsEntry2("PaymentNotes")
						end if
					end if
					rsPayMeth = rsEntry2("PmtTypes")
					rsClient = TRIM(rsEntry2("LastName")) & ",&nbsp;" & TRIM(rsEntry2("FirstName"))
					rsLoc = rsEntry2("LocationName")
					rsPaid = rsEntry2("PmtAmt")
					rsEntry2.MoveNext
				end if
		
		
				if datevalue(rsSaleDate) <> datevalue(tmpCurDate) OR tmpCurPayType<>rsPayMethID OR (byEmployee AND CLNG(tmpCurEmpID)<>CLNG(rsEmpID)) then
					if dayPMRecCount <> 0 then

		%>
						<tr>
						  <td colspan="4" class="right"><strong>Total <%=tmpCurPayTypeName%> Sales (<%=dayPMRecCount%> item<%if dayPMRecCount>1 then response.write "s" end if%>):</strong></td>
<% if NOT request.form("frmExpReport")="true" then %>
						  <td class="right"><strong><%=FmtCurrency(tmpPMTotal)%></strong>&nbsp;&nbsp;</td>
<% else %>
						  <td class="right"><strong><%=FmtNumber(tmpPMTotal)%></strong></td>
<% end if %>
						</tr>
<% if NOT request.form("frmExpReport")="true" then %>
						<tr   style="height:1px; font-size: 1px; line-height: 1px;background-color:#666666;">
						  <td colspan="5"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"><% end if %></td>
						</tr>
<% end if %>
		<%
					end if
					dayPMRecCount = 0
					tmpPMTotal = 0
				end if
				
				if datevalue(rsSaleDate) <> datevalue(tmpCurDate) OR (byEmployee AND CLNG(tmpCurEmpID)<>CLNG(rsEmpID)) then

					if dayRecCount <> 0 then
		%>
						<% if NOT request.form("frmExpReport")="true" then %>
							<tr   style="height:2px; font-size: 2px; line-height: 2px;background-color:#666666;">
								  <td colspan="5"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" style="height:2px; font-size: 2px; line-height: 2px;" width="100%"><% end if %></td>
							</tr>
						<% end if %>
							<tr>
							  <td colspan="12"><table class="mainText" width="100%"  cellspacing="0">
								<tr>
								  <td >&nbsp;</td>
								  <td height="30" class="right"><strong><%= getHotWord(59)%>:&nbsp;&nbsp; </strong></td>
								  <td class="right"><strong>Cash</strong></td>
								  <td class="right"><strong>Checks</strong></td>
								  <td class="right"><strong><%= getHotWord(30)%></strong></td>
								  <td class="right"><strong>Other</strong></td>
								  <td class="right"><strong><%= ucase(getHotWord(22))%>&nbsp;&nbsp;</strong></td>
								<td colspan="">					
								</tr>
<%		
		if useTips then
				strSQL = "SELECT [Sales Details].CategoryID, [Sales Details].Location, [Sales].EmployeeID, SUM(tblSDPayments.SDPaymentAmount) AS TipAmt, Sales.SaleDate, [Payment Types].PmtTypes, [Payment Types].Item#, [Payment Types].[CashEQ] "

                'strSQL = strSQL & "FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN [Payment Types] ON Sales.PaymentMethod = [Payment Types].Item# "
                'CB 54_3168 - New Sales Tables
                strSQL = strSQL & "FROM tblSDPayments INNER JOIN Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID ON tblSDPayments.SDID = [Sales Details].SDID INNER JOIN tblPayments INNER JOIN [Payment Types] ON tblPayments.PaymentMethod = [Payment Types].Item# ON tblSDPayments.PaymentID = tblPayments.PaymentID "

				strSQL = strSQL & "WHERE 1=1 "
				' ADDED 3/10/08 by Brad
				' Added a check vs the time ... either we're past the starting day (at any time) OR we're on the starting day of the range AND past the starting time
				' In other words, edge dates of the range are the only days that the time matters.
				if useSaleTimeS then
					strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate > " & DateSep & cdate(cSDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102)  >= " & TimeSepB & request.form("optStartTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
				end if

				' ADDED 3/10/08 by Brad
				' Added a check vs the time ... either we're before the ending day (at any time) OR we're on the ending day of the range AND before the ending time
				' In other words, edge dates of the range are the only days that the time matters.
				if useSaleTimeE then
					strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate < " & DateSep & cdate(cEDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102) <= " & TimeSepB & request.form("optEndTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
				end if
				strSQL = strSQL & "GROUP BY [Sales Details].CategoryID, [Sales Details].Location, [Sales].EmployeeID, Sales.SaleDate, [Payment Types].PmtTypes, [Payment Types].Item#, [Payment Types].[CashEQ] "
				strSQL = strSQL & "HAVING ([Sales Details].CategoryID = 21) "
				strSQL = strSQL & "AND (Sales.SaleDate = " & DateSep & cstr(tmpCurDate) & DateSep & ") "
				strSQL = strSQL & "AND [Payment Types].[CashEQ]=1 "
				if cLoc<>0 then
					strSQL = strSQL & "AND [Sales Details].Location=" & cLoc & " "
				end if
				if request.form("optEmployee")<>"-1" then
					strSQL = strSQL & " AND [Sales].EmployeeID=" & request.form("optEmployee") & " "
				end if
				if byEmployee then
					if CLNG(tmpCurEmpID)=0 OR CLNG(tmpCurEmpID)=-2222 then
						strSQL = strSQL & " AND [Sales].EmployeeID=0"
					else
						strSQL = strSQL & " AND [Sales].EmployeeID=" & tmpCurEmpID & " "
					end if
				end if
				rsTipDayA.CursorLocation = 3
				rsTipDayA.open strSQL, cnWS
				'response.Write(strSQL) & "<br />"
				Set rsTipDayA.ActiveConnection = Nothing
				
                'CB 54_3168 - New Sales Tables
			    strSQL = "SELECT '-1' AS SaleID WHERE 1=0"

				rsTipDayB.CursorLocation = 3
				rsTipDayB.open strSQL, cnWS
				Set rsTipDayB.ActiveConnection = Nothing

			
			'**********  CALCULATION OF TIP TOTALS SK 02/09/05  *******************
				if NOT rsTipDayA.EOF then
					Do Until rsTipDayA.EOF
						if rsTipDayA("Item#") = 1 then
							dayTotalCashTip = dayTotalCashTip + rsTipDayA("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayA("TipAmt")
							TotalCashTip = TotalCashTip + rsTipDayA("TipAmt")
							TotalTip=TotalTip+rsTipDayA("TipAmt")
						elseif rsTipDayA("Item#") = 2 then
							dayTotalCheckTip = dayTotalCheckTip + rsTipDayA("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayA("TipAmt")
							TotalCheckTip = TotalCheckTip + rsTipDayA("TipAmt")
							TotalTip=TotalTip+rsTipDayA("TipAmt")
						elseif rsTipDayA("Item#") >= 3 AND rsTipDayA("Item#") <= 6 then
							dayTotalCreditTip = dayTotalCreditTip + rsTipDayA("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayA("TipAmt")
							TotalCreditTip = TotalCreditTip + rsTipDayA("TipAmt")
							TotalTip=TotalTip+rsTipDayA("TipAmt")
						else
							dayTotalOtherTip = dayTotalOtherTip + rsTipDayA("TipAmt")		
							dayTotalTip=dayTotalTip+rsTipDayA("TipAmt")
							TotalOtherTip = TotalOtherTip + rsTipDayA("TipAmt")		
							TotalTip=TotalTip+rsTipDayA("TipAmt")
						end if
						rsTipDayA.movenext
					Loop
				end if
			
				if NOT rsTipDayB.EOF then
					Do Until rsTipDayB.EOF
						if rsTipDayB("Item#") = 1 then
							dayTotalCashTip = dayTotalCashTip + rsTipDayB("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayB("TipAmt")
							TotalCashTip = TotalCashTip + rsTipDayB("TipAmt")
							TotalTip=TotalTip+rsTipDayB("TipAmt")
						elseif rsTipDayB("Item#") = 2 then
							dayTotalCheckTip = dayTotalCheckTip + rsTipDayB("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayB("TipAmt")
							TotalCheckTip = TotalCheckTip + rsTipDayB("TipAmt")
							TotalTip=TotalTip+rsTipDayB("TipAmt")
						elseif rsTipDayB("Item#") >= 3 AND rsTipDayB("Item#") <= 6 then
							dayTotalCreditTip = dayTotalCreditTip + rsTipDayB("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayB("TipAmt")
							TotalCreditTip = TotalCreditTip + rsTipDayB("TipAmt")
							TotalTip=TotalTip+rsTipDayB("TipAmt")
						else
							dayTotalOtherTip = dayTotalOtherTip + rsTipDayB("TipAmt")		
							dayTotalTip=dayTotalTip+rsTipDayB("TipAmt")
							TotalOtherTip = TotalOtherTip + rsTipDayB("TipAmt")		
							TotalTip=TotalTip+rsTipDayB("TipAmt")
						end if
						rsTipDayB.movenext
					Loop
				end if


				rsTipDayA.close
				rsTipDayB.close
			end if
%>
								<tr>
								  <td align="left" ><strong>&nbsp;Sale Total for <%=FmtDateShort(tmpCurDate)%></strong>&nbsp;&nbsp; </strong> </td>
								  <td class="right"><strong>Total Received:&nbsp;&nbsp; </strong></td>
							<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotCash-dayTotalCashTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotCheck-dayTotalCheckTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotCredit-dayTotalCreditTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotOther-dayTotalOtherTip)%></strong></td>
								  <td width="100" class="right"><strong><%=FmtCurrency((dayTotCash-dayTotalCashTip)+(dayTotCheck-dayTotalCheckTip)+(dayTotCredit-dayTotalCreditTip)+(dayTotOther-dayTotalOtherTip))%>&nbsp;&nbsp;</strong></td>
							<% else %>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotCash-dayTotalCashTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotCheck-dayTotalCheckTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotCredit-dayTotalCreditTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotOther-dayTotalOtherTip)%></strong></td>
								  <td width="100" class="right"><strong><%=FmtNumber((dayTotCash-dayTotalCashTip)+(dayTotCheck-dayTotalCheckTip)+(dayTotCredit-dayTotalCreditTip)+(dayTotOther-dayTotalOtherTip))%></strong></td>
							<% end if %>
								</tr>
								<tr class="right">
								  <td align="left" ><strong>&nbsp;(<%=dayRecCount%>&nbsp;<%= getHotWord(217)%><%if dayRecCount > 1 then response.write "s" end if%>)</strong></td>
								  <td class="right"><strong>Tips:&nbsp;&nbsp;</strong></td>
							<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotalCashTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotalCheckTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotalCreditTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotalOtherTip)%></strong></td>
								  <td width="100" class="right"><strong><%=FmtCurrency(dayTotalTip)%>&nbsp;&nbsp;</strong></td>
							<% else %>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotalCashTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotalCheckTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotalCreditTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotalOtherTip)%></strong></td>
								  <td width="100" class="right"><strong><%=FmtNumber(dayTotalTip)%></strong></td>
							<% end if %>
								</tr>
							<%if NOT includeTips and dayTotalTip<>0 then %>
								<tr class="right">
								  <td >&nbsp;</td>
								  <td class="right"><strong>Tips Paid Out:&nbsp;&nbsp;</strong></td>
							<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><span style="color:#990000;">-<%=FmtCurrency(dayTotalTip)%></span></strong></td>
								  <td class="right">-</td>
								  <td class="right">-</td>
								  <td class="right">-</td>
								  <td class="right"><strong><span style="color:#990000;">-<%=FmtCurrency(dayTotalTip)%>&nbsp;&nbsp;</span></strong></td>
							<% else %>
								  <td class="right"><strong><span style="color:#990000;">-<%=FmtNumber(dayTotalTip)%></span></strong></td>
								  <td class="right">-</td>
								  <td class="right">-</td>
								  <td class="right">-</td>
								  <td class="right"><strong><span style="color:#990000;">-<%=FmtNumber(dayTotalTip)%></span></strong></td>
							<% end if %>
								</tr>
							<% end if %>
							<% if NOT request.form("frmExpReport")="true" then %>
								<tr>
								  <td colspan="2"></td>
								  <td    style="background-color:#CCCCCC;height:1px; font-size: 1px; line-height: 1px;" colspan="5"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
							<% end if %>
								<tr class="right">
								  <td >&nbsp;</td>
								  <td class="right"><strong>In Drawer:&nbsp;&nbsp;</strong></td>
							<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotCash) else response.write FmtCurrency((dayTotCash-dayTotalCashTip)-dayTotalTip) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotCheck) else response.write FmtCurrency(dayTotCheck) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotCredit) else response.write FmtCurrency(dayTotCredit) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotOther) else response.write FmtCurrency(dayTotOther) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotCash+dayTotCheck+dayTotCredit+dayTotOther) else response.write FmtCurrency((dayTotCash-dayTotalCashTip)-dayTotalTip+dayTotCheck+dayTotCredit+dayTotOther) end if%>&nbsp;&nbsp;</strong></td>
							<% else %>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotCash) else response.write FmtNumber((dayTotCash-dayTotalCashTip)-dayTotalTip) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotCheck) else response.write FmtNumber(dayTotCheck) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotCredit) else response.write FmtNumber(dayTotCredit) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotOther) else response.write FmtNumber(dayTotOther) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotCash+dayTotCheck+dayTotCredit+dayTotOther) else response.write FmtNumber((dayTotCash-dayTotalCashTip)-dayTotalTip+dayTotCheck+dayTotCredit+dayTotOther) end if%></strong></td>
							<% end if %>
								</tr>
							  </table></td>
							  </tr>
		<%
					end if
					tmpCurDate = rsSaleDate
					tmpCurEmpID = rsEmpID
					tmpCurPayType = ""
		%>
							<tr>
							  <td colspan="5">&nbsp;</td>
							</tr>
							<tr>
							  <td class="whiteHeader" colspan="5" style="background-color:<%=session("pageColor4")%>;">&nbsp;<strong><%=WeekDayName(WeekDay(tmpCurDate))%>,&nbsp;<%=MonthName(Month(tmpCurDate))%>&nbsp;<%=Day(tmpCurDate)%>,&nbsp;<%=Year(tmpCurDate)%> 
								<%if byEmployee then%>
								- 
									<%if CLNG(tmpCurEmpID)<>0 then%>
										<%=FmtTrnName(tmpCurEmpID)%>
									<%else%>
										<%="Owner"%>
									<%end if%>
								<%end if%>							  
							  </strong></td>
							</tr>
		<%
					TotalCash = TotalCash + dayTotCash
					TotalCheck = TotalCheck + dayTotCheck
					TotalCredit = TotalCredit + dayTotCredit
					TotalOther = TotalOther + dayTotOther

					dayTotalTip = 0
					dayTotalCashTip = 0
					dayTotalCheckTip = 0
					dayTotalCreditTip = 0
					dayTotalOtherTip = 0
		
					dayRecCount = 0
					dayTotCash = 0
					dayTotCheck = 0
					dayTotCredit = 0
					dayTotOther = 0
				end if
				if tmpCurPayType<>rsPayMethID then
					tmpCurPayType = rsPayMethID
					tmpCurPayTypeName = rsPayMeth
		%>
							<tr>
							  <td height="35" colspan="5">&nbsp;&nbsp;<strong><%=rsPayMeth%></strong></td>
							  </tr>
							<tr class="smalltextBlack">
							  <td width="26%">&nbsp;<%=session("ClientHW")%></td>
							  <td width="13%"><%= getHotWord(115)%>&nbsp;</td>
							  <td width="22%"><%= getHotWord(8)%>&nbsp;</td>
							  <td width="17%" class="right">&nbsp;<%= getHotWord(90)%>/Check#&nbsp;&nbsp;</td>
							  <td class="right">Payment Amt&nbsp;&nbsp;</td>
							</tr>
						<% if NOT request.form("frmExpReport")="true" then %>
							<tr   style="height:1px; font-size: 1px; line-height: 1px;background-color:#666666;">
								  <td colspan="5"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" style="height:1px; font-size: 1px; line-height: 1px;" width="100%"><% end if %></td>
							</tr>
						<% end if %>
		<%
				end if
				if rowColor = "#F2F2F2" then
					rowColor = "#FAFAFA"
				else
					rowColor = "#F2F2F2"
				end if
		%>
							<tr style="background-color:<%=rowColor%>;">
							  <td>&nbsp;<%=rsClient%>&nbsp;&nbsp;</td>
					<% if NOT request.form("frmExpReport")="true" then %>
							  <td><a href="adm_tlbx_voidedit.asp?saleno=<%=rsSaleID%>"><%=Right(rsSaleID,4)%></a>&nbsp;</td>
					<% else %>
							   <td><%=Right(rsSaleID,4)%>&nbsp;</td>
					<% end if %>
							 <td><%=rsLoc%>&nbsp;&nbsp;</td>
							  <td class="right"><% if NOT isNULL(rscheckNo) then response.write rsCheckNo end if %>&nbsp;&nbsp;</td>
					<% if NOT request.form("frmExpReport")="true" then %>
							  <td class="right"><%=FmtCurrency(rsPaid)%>&nbsp;&nbsp;</td>
					<% else %>
								<td class="right"><%=FmtNumber(rsPaid)%></td>
					<% end if %>
							</tr>
							<% if NOT request.form("frmExpReport")="true" then %>
								<tr   style="height:1px; font-size: 1px; line-height: 1px;background-color:#666666;">
								  <td colspan="5"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" style="height:1px; font-size: 1px; line-height: 1px;" width="100%"><% end if %></td>
								</tr>
							<% end if %>
		<%
				if rsEntry.EOF AND rsEntry2.EOF then
					cont = false
				end if
		
				if rsPayMethID = 1 then
					dayTotCash = dayTotCash + FormatNumber(rsPaid, 2)
				elseif rsPayMethID = 2 then
					dayTotCheck = dayTotCheck + FormatNumber(rsPaid, 2)
				elseif rsPayMethID >= 3 AND rsPayMethID <= 6 then
					dayTotCredit = dayTotCredit + FormatNumber(rsPaid, 2)
				else
					dayTotOther = dayTotOther + FormatNumber(rsPaid, 2)		
				end if
				dayPMRecCount = dayPMRecCount + 1
				dayRecCount = dayRecCount + 1
				TotalRecCount = TotalRecCount + 1
				tmpPMTotal = tmpPMTotal + FormatNumber(rsPaid, 2)
			Loop
			rsEntry.close
			rsEntry2.close
			set rsEntry = nothing
			set rsEntry2 = nothing
		
			if TotalRecCount > 0 then
				TotalCash = TotalCash + dayTotCash
				TotalCheck = TotalCheck + dayTotCheck
				TotalCredit = TotalCredit + dayTotCredit
				TotalOther = TotalOther + dayTotOther
		%>
						<tr >
						  <td colspan="4" class="right"><strong>Total <%=tmpCurPayTypeName%> Sales (<%=dayPMRecCount%> item<%if dayPMRecCount>1 then response.write "s" end if%>):</strong></td>
						<% if NOT request.form("frmExpReport")="true" then %>
						  <td class="right"><strong><%=FmtCurrency(tmpPMTotal)%></strong>&nbsp;&nbsp;</td>
						<% else %>
						  <td class="right"><strong><%=FmtNumber(tmpPMTotal)%></strong></td>
						<% end if %>
						</tr>
						<% if NOT request.form("frmExpReport")="true" then %>
						<tr    style="height:2px; font-size: 2px; line-height: 2px;background-color:#666666;">
						  <td colspan="5"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" style="height:2px; font-size: 2px; line-height: 2px;" width="100%"><% end if %></td>
						</tr>
						<% end if %>
<%
		if useTips then
				strSQL = "SELECT [Sales Details].CategoryID, [Sales Details].Location, [Sales].EmployeeID, SUM(tblSDPayments.SDPaymentAmount) AS TipAmt, Sales.SaleDate, [Payment Types].PmtTypes, [Payment Types].Item#, [Payment Types].[CashEQ] "

				'strSQL = strSQL & "FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN [Payment Types] ON Sales.PaymentMethod = [Payment Types].Item# "
                'CB 54_3168 - New Sales Tables
                strSQL = strSQL & "FROM tblSDPayments INNER JOIN Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID ON tblSDPayments.SDID = [Sales Details].SDID INNER JOIN tblPayments INNER JOIN [Payment Types] ON tblPayments.PaymentMethod = [Payment Types].Item# ON tblSDPayments.PaymentID = tblPayments.PaymentID "

				strSQL = strSQL & "WHERE 1=1 "
				' ADDED 3/10/08 by Brad
				' Added a check vs the time ... either we're past the starting day (at any time) OR we're on the starting day of the range AND past the starting time
				' In other words, edge dates of the range are the only days that the time matters.
				if useSaleTimeS then
					strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate > " & DateSep & cdate(cSDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102)  >= " & TimeSepB & request.form("optStartTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
				end if

				' ADDED 3/10/08 by Brad
				' Added a check vs the time ... either we're before the ending day (at any time) OR we're on the ending day of the range AND before the ending time
				' In other words, edge dates of the range are the only days that the time matters.
				if useSaleTimeE then
					strSQL = strSQL & " AND (1 = CASE WHEN Sales.SaleDate < " & DateSep & cdate(cEDate) & DateSep & " OR (CONVERT(DATETIME, '1899-12-30 ' + CONVERT(varchar, Sales.SaleTime, 108), 102) <= " & TimeSepB & request.form("optEndTime") & TimeSepA & ") THEN 1 ELSE 0 END) "
				end if
				if request.form("optCashRegister")<>"" and request.form("optCashRegister")<>"-1" then
					strSQL = strSQL & "  AND CashRegisterID = " & request.form("optCashRegister")
				end if
				strSQL = strSQL & "GROUP BY [Sales Details].CategoryID, [Sales Details].Location, [Sales].EmployeeID, Sales.SaleDate, [Payment Types].PmtTypes, [Payment Types].Item#, [Payment Types].[CashEQ] "
				strSQL = strSQL & "HAVING ([Sales Details].CategoryID = 21) "
				strSQL = strSQL & "AND (Sales.SaleDate = " & DateSep & cstr(tmpCurDate) & DateSep & ") "
				strSQL = strSQL & "AND [Payment Types].[CashEQ]=1 "
				if cLoc<>0 then
					strSQL = strSQL & "AND [Sales Details].Location=" & cLoc & " "
				end if
				if request.form("optEmployee")<>"-1" then
					strSQL = strSQL & " AND [Sales].EmployeeID=" & request.form("optEmployee") & " "
				end if
				if byEmployee then
					if CLNG(tmpCurEmpID)=0 OR CLNG(tmpCurEmpID)=-2222 then
						strSQL = strSQL & " AND [Sales].EmployeeID=0"
					else
						strSQL = strSQL & " AND [Sales].EmployeeID=" & tmpCurEmpID & " "
					end if
				end if
				rsTipDayA.CursorLocation = 3
				rsTipDayA.open strSQL, cnWS
				'response.Write(strSQL) & "<br />"
				Set rsTipDayA.ActiveConnection = Nothing

                'CB 54_3168 - New Sales Tables
			    strSQL = "SELECT '-1' AS SaleID WHERE 1=0"

				rsTipDayB.CursorLocation = 3
				rsTipDayB.open strSQL, cnWS
				Set rsTipDayB.ActiveConnection = Nothing

			
			'**********  CALCULATION OF TIP TOTALS SK 02/09/05  *******************
				if NOT rsTipDayA.EOF then
					Do Until rsTipDayA.EOF
						if rsTipDayA("Item#") = 1 then
							dayTotalCashTip = dayTotalCashTip + rsTipDayA("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayA("TipAmt")
							TotalCashTip = TotalCashTip + rsTipDayA("TipAmt")
							TotalTip=TotalTip+rsTipDayA("TipAmt")
						elseif rsTipDayA("Item#") = 2 then
							dayTotalCheckTip = dayTotalCheckTip + rsTipDayA("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayA("TipAmt")
							TotalCheckTip = TotalCheckTip + rsTipDayA("TipAmt")
							TotalTip=TotalTip+rsTipDayA("TipAmt")
						elseif rsTipDayA("Item#") >= 3 AND rsTipDayA("Item#") <= 6 then
							dayTotalCreditTip = dayTotalCreditTip + rsTipDayA("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayA("TipAmt")
							TotalCreditTip = TotalCreditTip + rsTipDayA("TipAmt")
							TotalTip=TotalTip+rsTipDayA("TipAmt")
						else
							dayTotalOtherTip = dayTotalOtherTip + rsTipDayA("TipAmt")		
							dayTotalTip=dayTotalTip+rsTipDayA("TipAmt")
							TotalOtherTip = TotalOtherTip + rsTipDayA("TipAmt")		
							TotalTip=TotalTip+rsTipDayA("TipAmt")
						end if
						rsTipDayA.movenext
					Loop
				end if
			
				if NOT rsTipDayB.EOF then
					Do Until rsTipDayB.EOF
						if rsTipDayB("Item#") = 1 then
							dayTotalCashTip = dayTotalCashTip + rsTipDayB("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayB("TipAmt")
							TotalCashTip = TotalCashTip + rsTipDayB("TipAmt")
							TotalTip=TotalTip+rsTipDayB("TipAmt")
						elseif rsTipDayB("Item#") = 2 then
							dayTotalCheckTip = dayTotalCheckTip + rsTipDayB("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayB("TipAmt")
							TotalCheckTip = TotalCheckTip + rsTipDayB("TipAmt")
							TotalTip=TotalTip+rsTipDayB("TipAmt")
						elseif rsTipDayB("Item#") >= 3 AND rsTipDayB("Item#") <= 6 then
							dayTotalCreditTip = dayTotalCreditTip + rsTipDayB("TipAmt")
							dayTotalTip=dayTotalTip+rsTipDayB("TipAmt")
							TotalCreditTip = TotalCreditTip + rsTipDayB("TipAmt")
							TotalTip=TotalTip+rsTipDayB("TipAmt")
						else
							dayTotalOtherTip = dayTotalOtherTip + rsTipDayB("TipAmt")		
							dayTotalTip=dayTotalTip+rsTipDayB("TipAmt")
							TotalOtherTip = TotalOtherTip + rsTipDayB("TipAmt")		
							TotalTip=TotalTip+rsTipDayB("TipAmt")
						end if
						rsTipDayB.movenext
					Loop
				end if


				rsTipDayA.close
				rsTipDayB.close
			end if


%>						
						<% 'if cSDate <> cEDate OR (byEmployee) then %>
								<% if NOT request.form("frmExpReport")="true" then %>
							<tr   style="height:2px; font-size: 2px; line-height: 2px;background-color:#666666;">
								  <td colspan="5"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" style="height:2px; font-size: 2px; line-height: 2px;" width="100%"><% end if %></td>
							</tr>
<% end if %>
							<tr>
							  <td colspan="12"><table class="mainText" width="100%" cellspacing="0">
								<tr>
								  <td >&nbsp;</td>
								  <td height="30" class="right"><strong><%= getHotWord(59)%>:&nbsp;&nbsp; </strong></td>
								  <td class="right"><strong>Cash</strong></td>
								  <td class="right"><strong>Checks</strong></td>
								  <td class="right"><strong><%= getHotWord(30)%></strong></td>
								  <td class="right"><strong>Other</strong></td>
								  <td class="right"><strong><%= ucase(getHotWord(22))%>&nbsp;&nbsp;</strong></td>
								<td colspan="">					
								</tr>
								<tr>
								  <td align="left" ><strong>&nbsp;Sale Total for <%=FmtDateShort(tmpCurDate)%></strong>&nbsp;&nbsp; </strong> </td>
								  <td class="right"><strong>Total Received:&nbsp;&nbsp; </strong></td>
							<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotCash-dayTotalCashTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotCheck-dayTotalCheckTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotCredit-dayTotalCreditTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotOther-dayTotalOtherTip)%></strong></td>
								  <td width="100" class="right"><strong><%=FmtCurrency((dayTotCash-dayTotalCashTip)+(dayTotCheck-dayTotalCheckTip)+(dayTotCredit-dayTotalCreditTip)+(dayTotOther-dayTotalOtherTip))%>&nbsp;&nbsp;</strong></td>
							<% else %>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotCash-dayTotalCashTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotCheck-dayTotalCheckTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotCredit-dayTotalCreditTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotOther-dayTotalOtherTip)%></strong></td>
								  <td width="100" class="right"><strong><%=FmtNumber((dayTotCash-dayTotalCashTip)+(dayTotCheck-dayTotalCheckTip)+(dayTotCredit-dayTotalCreditTip)+(dayTotOther-dayTotalOtherTip))%></strong></td>
							<% end if %>
								</tr>
								<tr class="right">
								  <td><strong>&nbsp;(<%=dayRecCount%>&nbsp;<%=getHotWord(217)%><%if dayRecCount > 1 then response.write "s" end if%>)</strong></td>
								  <td class="right"><strong>Tips:&nbsp;&nbsp;</strong></td>
							<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotalCashTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotalCheckTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotalCreditTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtCurrency(dayTotalOtherTip)%></strong></td>
								  <td width="100" class="right"><strong><%=FmtCurrency(dayTotalTip)%>&nbsp;&nbsp;</strong></td>
							<% else %>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotalCashTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotalCheckTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotalCreditTip)%></strong></td>
								  <td width="80" class="right"><strong><%=FmtNumber(dayTotalOtherTip)%></strong></td>
								  <td width="100" class="right"><strong><%=FmtNumber(dayTotalTip)%></strong></td>
							<% end if %>
								</tr>
							<%if NOT includeTips and dayTotalTip<>0 then %>
								<tr class="right">
								  <td >&nbsp;</td>
								  <td><strong>Tips Paid Out:&nbsp;&nbsp;</strong></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td><strong><span style="color:#990000;">-<%=FmtCurrency(dayTotalTip)%></span></strong></td>
								  <td>-</td>
								  <td>-</td>
								  <td>-</td>
								  <td><strong><span style="color:#990000;">-<%=FmtCurrency(dayTotalTip)%>&nbsp;&nbsp;</span></strong></td>
								<% else %>
								  <td><strong><span style="color:#990000;">-<%=FmtNumber(dayTotalTip)%></span></strong></td>
								  <td>-</td>
								  <td>-</td>
								  <td>-</td>
								  <td><strong><span style="color:#990000;">-<%=FmtNumber(dayTotalTip)%></span></strong></td>
								<% end if %>
								</tr>
							<% end if %>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr>
								  <td colspan="2"></td>
								  <td style="background-color:#CCCCCC;" colspan="5"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" width="100%" style="height:1px; font-size: 1px; line-height: 1px;"></td>
								</tr>
								<% end if %>		
								<tr class="right">
								  <td >&nbsp;</td>
								  <td class="right"><strong>In Drawer:&nbsp;&nbsp;</strong></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotCash) else response.write FmtCurrency((dayTotCash)-dayTotalTip) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotCheck) else response.write FmtCurrency(dayTotCheck) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotCredit) else response.write FmtCurrency(dayTotCredit) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotOther) else response.write FmtCurrency(dayTotOther) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(dayTotCash+dayTotCheck+dayTotCredit+dayTotOther) else response.write FmtCurrency((dayTotCash)-dayTotalTip+dayTotCheck+dayTotCredit+dayTotOther) end if%>&nbsp;&nbsp;</strong></td>
								<% else %>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotCash) else response.write FmtNumber((dayTotCash)-dayTotalTip) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotCheck) else response.write FmtNumber(dayTotCheck) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotCredit) else response.write FmtNumber(dayTotCredit) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotOther) else response.write FmtNumber(dayTotOther) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(dayTotCash+dayTotCheck+dayTotCredit+dayTotOther) else response.write FmtNumber((dayTotCash)-dayTotalTip+dayTotCheck+dayTotCredit+dayTotOther) end if%></strong></td>
								<% end if %>
								</tr>
				<% 'end if  ' cSDate <> cEDate %>
							<tr>
							  <td colspan="12">
							  <table class="mainText" width="100%"  cellspacing="0">
							  </table></td>
							  </tr>
		
							<tr>
							  <td colspan="12">
							  <table class="mainText" width="100%"  cellspacing="0">
							  <% if cSDate <> cEDate OR (byEmployee) then %>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr  style="height:2px; font-size: 2px; line-height: 2px;background-color:<%=session("pageColor4")%>;">
								  <td colspan="7"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" style="height:2px; font-size: 2px; line-height: 2px;" width="100%"><%end if%></td>
								</tr>
								<% end if %>
							<% end if %>
								<tr>
								  <td colspan="7"><strong>&nbsp;</strong></td>
								  </tr>
								<tr>
								  <td ><strong>&nbsp;GRAND TOTAL</strong></td>
								  <td height="19" class="right"><strong><%= getHotWord(59)%>:&nbsp;&nbsp; </strong></td>
								  <td width="80" class="right"><strong>Cash</strong></td>
								  <td width="80" class="right"><strong>Checks</strong></td>
								  <td width="80" class="right"><strong><%= getHotWord(30)%></strong></td>
								  <td width="80" class="right"><strong>Other</strong></td>
								  <td width="100" class="right"><strong><%= UCase(getHotWord(22))%>&nbsp;&nbsp;</strong></td>
								</tr>
								<tr class="right">
								  <td align="left" >&nbsp;<strong>(<%=TotalRecCount%>&nbsp;<%=getHotWord(217)%><%if TotalRecCount > 1 then response.write "s" end if%>) </strong></td>
								  <td class="right"><strong>Total Received:&nbsp;&nbsp; </strong></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><%=FmtCurrency(TotalCash-TotalCashTip)%></strong></td>
								  <td class="right"><strong><%=FmtCurrency(TotalCheck-TotalCheckTip)%></strong></td>
								  <td class="right"><strong><%=FmtCurrency(TotalCredit-TotalCreditTip)%></strong></td>
								  <td class="right"><strong><%=FmtCurrency(TotalOther-TotalOtherTip)%></strong></td>
								  <td class="right"><strong><%=FmtCurrency((TotalCash-TotalCashTip)+(TotalCheck-TotalCheckTip)+(TotalCredit-TotalCreditTip)+(TotalOther-TotalOtherTip))%>&nbsp;&nbsp;</strong></td>
								<% else %>
								  <td class="right"><strong><%=FmtNumber(TotalCash-TotalCashTip)%></strong></td>
								  <td class="right"><strong><%=FmtNumber(TotalCheck-TotalCheckTip)%></strong></td>
								  <td class="right"><strong><%=FmtNumber(TotalCredit-TotalCreditTip)%></strong></td>
								  <td class="right"><strong><%=FmtNumber(TotalOther-TotalOtherTip)%></strong></td>
								  <td class="right"><strong><%=FmtNumber((TotalCash-TotalCashTip)+(TotalCheck-TotalCheckTip)+(TotalCredit-TotalCreditTip)+(TotalOther-TotalOtherTip))%></strong></td>
								<% end if %>
								</tr>
					<%if useTips then %>
								<tr class="right">
								  <td >&nbsp;</td>
								  <td class="right"><strong>Tips Received:&nbsp;&nbsp;</strong></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><%=FmtCurrency(TotalCashTip)%></strong></td>
								  <td class="right"><strong><%=FmtCurrency(TotalCheckTip)%></strong></td>
								  <td class="right"><strong><%=FmtCurrency(TotalCreditTip)%></strong></td>
								  <td class="right"><strong><%=FmtCurrency(TotalOtherTip)%></strong></td>
								  <td class="right"><strong><%=FmtCurrency(TotalTip)%>&nbsp;&nbsp;</strong></td>
								<% else %>
								  <td class="right"><strong><%=FmtNumber(TotalCashTip)%></strong></td>
								  <td class="right"><strong><%=FmtNumber(TotalCheckTip)%></strong></td>
								  <td class="right"><strong><%=FmtNumber(TotalCreditTip)%></strong></td>
								  <td class="right"><strong><%=FmtNumber(TotalOtherTip)%></strong></td>
								  <td class="right"><strong><%=FmtNumber(TotalTip)%></strong></td>
								<% end if %>
								</tr>
							<%if NOT includeTips and TotalTip<>0 then %>
								<tr class="right">
								  <td >&nbsp;</td>
								  <td class="right"><strong>Tips Paid Out:&nbsp;&nbsp;</strong></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><span style="color:#990000;">-<%=FmtCurrency(TotalTip)%></span></strong></td>
								  <td class="right">-</td>
								  <td class="right">-</td>
								  <td class="right">-</td>
								  <td class="right"><strong><span style="color:#990000;">-<%=FmtCurrency(TotalTip)%>&nbsp;&nbsp;</span></strong></td>
								<% else %>
								  <td class="right"><strong><span style="color:#990000;">-<%=FmtNumber(TotalTip)%></span></strong></td>
								  <td class="right">-</td>
								  <td class="right">-</td>
								  <td class="right">-</td>
								  <td class="right"><strong><span style="color:#990000;">-<%=FmtNumber(TotalTip)%></span></strong></td>
								<% end if %>
								</tr>
							<% end if %>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr>
								  <td colspan="2"></td>
								  <td style="background-color:#CCCCCC;" colspan="5"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" width="100%" style="height:1px; font-size: 1px; line-height: 1px;"></td>
								</tr>
								<% end if %>
								<tr class="right">
								  <td >&nbsp;</td>
								  <td class="right"><strong>In Drawer:&nbsp;&nbsp;</strong></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(TotalCash) else response.write FmtCurrency(TotalCash-TotalTip) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(TotalCheck) else response.write FmtCurrency(TotalCheck) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(TotalCredit) else response.write FmtCurrency(TotalCredit) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(TotalOther) else response.write FmtCurrency(TotalOther) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtCurrency(TotalCash+TotalCheck+TotalCredit+TotalOther) else response.write FmtCurrency(TotalCash+TotalCheck+TotalCredit+TotalOther-TotalTip) end if%>
								<% else %>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(TotalCash) else response.write FmtNumber(TotalCash-TotalTip) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(TotalCheck) else response.write FmtNumber(TotalCheck) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(TotalCredit) else response.write FmtNumber(TotalCredit) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(TotalOther) else response.write FmtNumber(TotalOther) end if%></strong></td>
								  <td class="right"><strong><%if includeTips then response.write FmtNumber(TotalCash+TotalCheck+TotalCredit+TotalOther) else response.write FmtNumber(TotalCash+TotalCheck+TotalCredit+TotalOther-TotalTip) end if%>
								<% end if %>
								  <% if NOT request.form("frmExpReport")="true" then %>&nbsp;&nbsp;<%end if %></strong></td>
								  </tr>
					<% end if 'UseTips%>
							  </table></td>
							  </tr>
							<% if NOT request.form("frmExpReport")="true" then %>
							<tr style="height:2px; font-size: 2px; line-height: 2px;background-color:<%=session("pageColor4")%>;">
							  <td colspan="7"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" style="height:2px; font-size: 2px; line-height: 2px;" width="100%"><% end if %></td>
							</tr>
							<% end if %>
<%
            else 
                response.Write "<br />There are no Cash Equivalent sales within the specified date and time."
			end if ''No Data
		end if '1st Load	
%>
							
						  </table></td>
						</tr>
					</table>
						  </td>
						</tr>
					  </table>
					</td>
				  </tr>
				 <% if NOT request.form("frmExpReport")="true" then %>
						 
				<% end if %>
				</table>
				</td>
				</tr>
				  </table>
				  </td>
				  </tr>
				  </table>
				<% if NOT request.form("frmExpReport")="true" then %>
<%= pageEnd %>				
<!-- #include file="post.asp" -->

				<%
		end if
		
	
	end if
%>
