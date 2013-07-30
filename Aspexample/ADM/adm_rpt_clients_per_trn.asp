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
	<!-- #include file="inc_rpt_tagging.asp" -->
	<!-- #include file="inc_utilities.asp" -->
	<!-- #include file="inc_rpt_save.asp" -->
	<%	dim doRefresh : doRefresh = false %>
	<!-- #include file="inc_date_arrows.asp" -->
	<!-- #include file="../inc_ajax.asp" --> 
	<!-- #include file="../inc_val_date.asp" --> 
	<%
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_DAY") then 
	%>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
	<%else%>
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_hotword.asp" -->
		
	<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

	Dim showDetails, cSDate, cEDate, cLoc, PercentAttend
	Dim showHeader, rowcolor, barcolor, curTrainer, TotAttend, GTotAttend, ap_view_all_locs
	
	ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
	
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
		cSDate = DateAdd("m",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
	end if

	if request.form("requiredtxtDateEnd")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(request.form("requiredtxtDateEnd"))
		Call SetLocale("en-us")
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if
	
	showDetails = request.form("optTrainer")

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

		if NOT request.form("frmExpReport")="true" then
		%>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_clients_per_trn", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %>
			<script type="text/javascript">
			function exportReport() {
				document.frmParameter.frmExpReport.value = "true";
				document.frmParameter.frmGenReport.value = "true";
				<% iframeSubmit "frmParameter", "adm_rpt_clients_per_trn.asp" %>
			}
			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
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
			<%= DisplayPhrase(reportPageTitlesDictionary,"Clientsperteacher") %>
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
						<td class="headText" valign="bottom"><b><%= pp_PageTitle("Clients Per Teacher") %></b></td>
						<td valign="bottom" class="right" height="26"> </td>
						</tr>
					</table>
					</td>
					</tr>
				<%end if %>
					<tr> 
					<td height="30"  valign="bottom" class="headText">
					<table class="mainText border4 center" cellspacing="0">
						<form name="frmParameter" action="adm_rpt_clients_per_trn.asp" method="POST">
						<input type="hidden" name="frmGenReport" value="">
						<input type="hidden" name="frmExpReport" value="">
						<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
							<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
							<input type="hidden" name="category" value="<%=category%>">
						<% end if %>
						<tr> 
						<td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b><span style="color:<%=session("pageColor4")%>">&nbsp;</span><%=xssStr(allHotWords(77))%>: 
						<input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
						<script type="text/javascript">
						var cal1 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStart'});
						cal1.a_tpl.yearscroll = true;
						</script>
						&nbsp;<%=xssStr(allHotWords(79))%>: 
						<input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
						<script type="text/javascript">
						var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
						cal2.a_tpl.yearscroll = true;
						</script>
						&nbsp;
						<%=xssStr(allHotWords(8))%>:<select name="optLoc" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
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
							document.frmParameter.optLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
						</script>
						
                      	<select name="optTrainer" >
                        <option value="-2" <%if request.form("optTrainer")="-2" then response.write "selected" end if%>>All Instructors - Summary</option>
						<%
							set rsEntry2 = Server.CreateObject("ADODB.Recordset")
							strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName "
							strSQL = strSQL & "FROM TRAINERS "
							strSQL = strSQL & "INNER JOIN [VISIT DATA] ON TRAINERS.TrainerID=[VISIT DATA].TrainerID "
							strSQL = strSQL & "WHERE TRAINERS.[Active]=1 AND TRAINERS.[Delete]=0 AND TRAINERS.TrainerID>0 AND TRAINERS.isSystem=0 "
							strSQL = strSQL & "AND ([VISIT DATA].ClassDate>=" & DateSep & cSDate & DateSep & ") "
							strSQL = strSQL & "AND ([VISIT DATA].ClassDate<=" & DateSep & cEDate & DateSep & ") "
							strSQL = strSQL & "ORDER BY TRAINERS.TrLastName"
							rsEntry2.CursorLocation = 3
							rsEntry2.open strSQL, cnWS
							Set rsEntry2.ActiveConnection = Nothing

							Do While NOT rsEntry2.EOF
						%>
							<option value="<%=rsEntry2("TrainerID")%>" <%if request.form("optTrainer")=CSTR(rsEntry2("TrainerID")) then response.write "selected" end if%>><%=FmtTrnNameNew(rsEntry2, true)%></option>
						<%
								rsEntry2.MoveNext
							Loop	
							rsEntry2.close
						%>
                      </select>
						<script type="text/javascript">
							document.frmParameter.optTrainer.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(113))%> - Summary";
						</script>
						&nbsp;
						<br />
						<% showDateArrows("frmParameter") %>
						&nbsp;&nbsp;<% taggingFilter %>&nbsp;&nbsp;
						<input type="button" name="Button" value="Generate" onClick="genReport();"></b>
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
				else%>
						<% exportToExcelButton %>
				<%end if%>
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
				 else%>
						<% taggingButtons("frmParameter") %>
				<%end if%>
						<% savingButtons "frmParameter", session("ClientHW") & "s Per Staff" %>
						</td>
						</tr>
						</form>
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" id="clientsPerTeacherGenTag" class="mainTextBig center-ch"> 
					
					<table class="mainText" width="100%" cellspacing="0">
						<tr>
						<td  colspan="2" valign="top" class="mainTextBig center-ch">&nbsp;</td>
						</tr>
						<tr > 
						<td  colspan="2" valign="top" class="mainTextBig center-ch">
		<% 
		end if			'end of frmExpreport value check before /head line	  
				if request.form("frmTagClients")="true" then
					if showdetails<>"-2" then
						strSQL = "SELECT [VISIT DATA].ClientID "
						strSQL = strSQL & "FROM CLIENTS "
						strSQL = strSQL & "INNER JOIN (TRAINERS INNER JOIN [VISIT DATA] ON TRAINERS.TrainerID = [VISIT DATA].TrainerID)ON CLIENTS.ClientID = [VISIT DATA].ClientID "
						if request.form("optFilterTagged")="on" then
							strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
							if session("mVarUserID")<>"" then
								strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
							end if
							strSQL = strSQL & " ) "
						end if
						strSQL = strSQL & "WHERE (CLIENTS.Deleted=0) AND [VISIT DATA].TrainerID <> 1 "
						if cLoc<>0 then
							strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
						end if
						strSQL = strSQL & "	AND ([VISIT DATA].ClassDate>=" & DateSep & cSDate & DateSep & ") "
						strSQL = strSQL & "	AND ([VISIT DATA].ClassDate<=" & DateSep & cEDate & DateSep & ") "
						strSQL = strSQL & " AND ([VISIT DATA].TrainerID = " & cLNG(showDetails) & ") "
						
						if request.form("frmTagClientsNew")="true" then
							clearAndTagQuery(strSQL)
						else
							tagQuery(strSQL)
						end if
						
					else 
					%>
						<script>
							alert("Summary results can't be tagged");
						</script>
					<%
					end if
					
					
				end if
							if request.form("frmGenReport")="true" then 
								if request.form("frmExpReport")="true" then
									Dim stFilename
									if showDetails="-2" then
										stFilename="attachment; filename=" & session("ClientHW") & "s Per Instructor- Summary " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									else
										stFilename="attachment; filename=" & session("ClientHW") & "s Per Instructor- " & FmtTrnName(CLNG(showDetails)) & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									end if
									Response.ContentType = "application/vnd.ms-excel" 
									Response.AddHeader "Content-Disposition", stFilename 
								end if

									strSQL = "SELECT Count([VISIT DATA].VisitRefNo) AS CountOfVisitRefNo "
									strSQL = strSQL & "FROM CLIENTS "
									strSQL = strSQL & "INNER JOIN [VISIT DATA] ON CLIENTS.ClientID = [VISIT DATA].ClientID "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
									end if
									strSQL = strSQL & "WHERE (CLIENTS.Deleted=0) AND [VISIT DATA].TrainerID <> 1 "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if
									strSQL = strSQL & "	AND ([VISIT DATA].ClassDate>=" & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "	AND ([VISIT DATA].ClassDate<=" & DateSep & cEDate & DateSep & ") "
									if request.form("optFilterTagged")="on" then
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
								rsEntry2.CursorLocation = 3
rsEntry2.open strSQL, cnWS
Set rsEntry2.ActiveConnection = Nothing

								if not rsEntry2.eof then
									GTotAttend=rsEntry2("CountofVisitRefNo")
								end if
								
								showHeader = "false"
								curTrainer=0
								TotAttend=0

								If showdetails="-2" then		'All Instructors Summary
									strSQL = "SELECT TRAINERS.TrLastName, TRAINERS.TrFirstName, 	[VISIT DATA].TrainerID, "
									strSQL = strSQL & "TRAINERS.DisplayName, Count([VISIT DATA].VisitRefNo) AS CountOfVisitRefNo "
									strSQL = strSQL & "FROM CLIENTS "
									strSQL = strSQL & "INNER JOIN (TRAINERS INNER JOIN [VISIT DATA] ON TRAINERS.TrainerID = [VISIT DATA].TrainerID) "
									strSQL = strSQL & "ON CLIENTS.ClientID = [VISIT DATA].ClientID "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
									end if
									strSQL = strSQL & "WHERE (CLIENTS.Deleted=0) "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if
									strSQL = strSQL & "	AND ([VISIT DATA].ClassDate>=" & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "	AND ([VISIT DATA].ClassDate<=" & DateSep & cEDate & DateSep & ") "
									if request.form("optFilterTagged")="on" then
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
									strSQL = strSQL & "GROUP BY TRAINERS.TrLastName, TRAINERS.TrFirstName, [VISIT DATA].TrainerID, "
									strSQL = strSQL & "TRAINERS.DisplayName "
									strSQL = strSQL & "ORDER BY CountofVisitRefNo DESC, TRAINERS.TrLastName, TRAINERS.TrFirstName"
								else		'Single Instructor 
									strSQL = "SELECT TRAINERS.TrLastName, TRAINERS.TrFirstName, [VISIT DATA].ClientID,CLIENTS.LastName, "
									strSQL = strSQL & "	CLIENTS.FirstName, CLIENTS.PostalCode, CLIENTS.HomePhone, CLIENTS.EmailName, "
									strSQL = strSQL & "	CLIENTS.Address, CLIENTS.City, CLIENTS.State, "
									strSQL = strSQL & "	TRAINERS.DisplayName, Count([VISIT DATA].VisitRefNo) AS CountOfVisitRefNo, "
									strSQL = strSQL & "	[VISIT DATA].TrainerID "
									strSQL = strSQL & "FROM CLIENTS "
									strSQL = strSQL & "INNER JOIN (TRAINERS INNER JOIN [VISIT DATA] ON TRAINERS.TrainerID = [VISIT DATA].TrainerID) "
									strSQL = strSQL & "	ON CLIENTS.ClientID = [VISIT DATA].ClientID "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
									end if
									strSQL = strSQL & "WHERE (CLIENTS.Deleted=0) "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if
									strSQL = strSQL & "	AND ([VISIT DATA].ClassDate>=" & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "	AND ([VISIT DATA].ClassDate<=" & DateSep & cEDate & DateSep & ") "
									if request.form("optFilterTagged")="on" then
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
									strSQL = strSQL & "GROUP BY TRAINERS.TrLastName, TRAINERS.TrFirstName, [VISIT DATA].ClientID,CLIENTS.LastName, "
									strSQL = strSQL & "	CLIENTS.FirstName, CLIENTS.PostalCode, CLIENTS.HomePhone, CLIENTS.EmailName, "
									strSQL = strSQL & "	CLIENTS.Address, CLIENTS.City, CLIENTS.State, "
									strSQL = strSQL & "	TRAINERS.DisplayName, [VISIT DATA].TrainerID "
									strSQL = strSQL & "	HAVING ([VISIT DATA].TrainerID = " & cLNG(showDetails) & ") "
									strSQL = strSQL & "ORDER BY TRAINERS.TrLastName, TRAINERS.TrFirstName, CLIENTS.LastName, CLIENTS.FirstName"
								end if
								rsEntry.CursorLocation = 3
rsEntry.open strSQL, cnWS
Set rsEntry.ActiveConnection = Nothing

								
									If showdetails="-2" then  '******* ALL INSTRUCTORS SUMMARY *******************
									%>
									<table class="mainText"  cellspacing="0" width="80%">
									<%if NOT rsEntry.EOF then			'EOF
										do while NOT rsEntry.EOF
											If showHeader = "false" then
										%>
											<tr>
											  <td colspan="2">&nbsp;</td>
											</tr>
											<tr>
											  <td align="Left" width="20%"><strong><%=xssStr(allHotWords(6))%></strong></td>
											  <td class="right" width="10%"><strong>Attendance</strong></td>
											  <td>&nbsp;</td>
											</tr>
<% if NOT request.form("frmExpReport")="true"  then %>
										<tr height="2">
											  <td colspan="3" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
<% end if %>
										<%	
											end if
											showHeader = "true"
											PercentAttend=FormatNumber((rsEntry("CountofVisitRefNo")/GTotAttend), 3) * 100
											
											if rowColor = "#F2F2F2" then
												rowColor = "#FAFAFA"
											else
												rowColor = "#F2F2F2"
											end if
											if barColor = session("pageColor4") then
												barColor = session("pageColor3")
											elseif barColor = session("pageColor3") then
												barColor = session("pageColor2")
											else
												barColor = session("pageColor4")
											end if
											%>
												<tr style="background-color:<%=rowColor%>;">
											    <td align="Left"><strong><%=rsEntry("TrLastName") & ",&nbsp;" & rsEntry("TrFirstName")%></strong></td>
												<td class="right"nowrap><%=rsEntry("CountofVisitRefNo")%></td>
												<% if NOT request.form("frmExpReport")="true" then%>												
													<td>
													 <table align="left" style="background-color:<%=barcolor%>;" width="<%=PercentAttend*10%>">
													 <tr height="9">
													 <td><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
													 </tr>
													</table>
													</td>											  
												<%end if%>
												</tr>
											<%		
											rsEntry.MoveNext
										loop
									end if	'eof
									%>
<% if NOT request.form("frmExpReport")="true"  then %>
									<tr height="2">
										<td colspan="3" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
									</tr>
<% end if %>
									<tr>
										<td><div class="right"><strong>Total Attendance:&nbsp;</strong></div></td>
										<td class="right"><strong><%=GTotAttend%></strong></td>
									</tr>
									</table>
									<%
									else		'************ Single Instructor  ******************
									%>
									<table class="mainText"  cellspacing="0" width="90%">
									<%if NOT rsEntry.EOF then			'EOF
										do while NOT rsEntry.EOF
											if curTrainer<>clng(rsEntry("TrainerID")) then	'if this is a new trainer then write the header cells
												if curTrainer<> "" AND curTrainer <> 0 then	'if this isn't the first record total cells
												%>
<% if NOT request.form("frmExpReport")="true"  then %>
												<tr height="2">
													<td colspan="6" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
												</tr>
<% end if %>
												<tr>
													<td colspan="5"><div class="right"><strong>Total Attendance For <%=UCASE(Replace(FmtTrnName(curTrainer),"&nbsp;"," "))%>:&nbsp;</strong></div></td>
													<td align="left"><strong><%=TotAttend%></strong></td>
												</tr>
												<%
												end if	
												%>
												<tr>
												  <td colspan="6">&nbsp;</td>
												</tr>
												<tr>
												  <td colspan="6" class="whiteHeader" style="background-color:<%=session("pageColor4")%>;" align="left"><%=FmtTrnNameNew(rsEntry, false)%>&nbsp;</td>
												</tr>
												<tr align="left" height="20">
													<td width="5%" class="center-ch"><strong># Visits</strong></td>
													<td width="17%"><strong><%=session("ClientHW")%></strong></td>
													<td width="10%"><strong><%= getHotWord(47)%></strong></td>
													<td width="8%" class="center-ch"><strong><%= getHotWord(48)%></strong></td>
													<td width="12%"><strong><%= getHotWord(82)%></strong></td>
													<td width="12%"><strong><%= getHotWord(39)%></strong></td>
												</tr>
<% if NOT request.form("frmExpReport")="true"  then %>
												<tr height="2">
													  <td colspan="6" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
												</tr>
<% end if %>
												<%
												TotAttend=0
											end if
											curTrainer =CLNG(rsEntry("TrainerID"))
											TotAttend = TotAttend + rsEntry("CountofVisitRefNo")
											
											if rowColor = "#F2F2F2" then
												rowColor = "#FAFAFA"
											else
												rowColor = "#F2F2F2"
											end if
											%>
												<tr align="left" style="background-color:<%=rowColor%>;">
													<td class="center-ch"><%=rsEntry("CountofVisitRefNo")%></td>
													<% if NOT request.form("frmExpReport")="true" then %>
														<td><a href="main_info.asp?id=<%=rsEntry("ClientID")%>&fl=true"><%=TRIM(rsEntry("LastName")) & ", " & TRIM(rsEntry("FirstName"))%></a></td>
													<% else %>
														<td><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></td>
													<% end if %>
													<td><%=rsEntry("City")%></td>
													<td class="center-ch"><%=rsEntry("State")%></td>
													<td><%=FmtPhoneNum(rsEntry("HomePhone"))%></td>
													<td><%=rsEntry("EmailName")%></td>
												</tr>
											<%		
											rsEntry.MoveNext
										loop
										%>
<% if NOT request.form("frmExpReport")="true"  then %>
											<tr height="2">
												<td colspan="6" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
											</tr>
<% end if %>
											<tr>
												<td colspan="5"><div class="right"><strong>Total Attendance For <%=UCASE(Replace(FmtTrnName(curTrainer),"&nbsp;"," "))%>:&nbsp;</strong></div></td>
												<td align="left"><strong><%=TotAttend%></strong></td>
											</tr>
											<tr>
											  <td colspan="6">&nbsp;</td>
											</tr>
									<%end if	'eof%>
									</table>
									<%
									end if
								rsEntry.close
								set rsEntry = nothing
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
