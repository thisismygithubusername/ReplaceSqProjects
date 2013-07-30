<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
	dim rsEntry, rsARC, strSQLARC, rsClassCount, strSQLClass
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	set rsARC = Server.CreateObject("ADODB.Recordset")
	set rsClassCount = Server.CreateObject("ADODB.Recordset")
	%>
	<!-- #include file="inc_accpriv.asp" -->
	<!-- #include file="inc_utilities.asp" -->
	<!-- #include file="inc_rpt_save.asp" -->
	<!-- #include file="inc_rpt_tagging.asp" -->
	<%	dim doRefresh : doRefresh = false %>
	<!-- #include file="inc_date_arrows.asp" -->
	<%
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_ANALYSIS") then 
	%>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
	<%
	else
	%>
			<!-- #include file="../inc_i18n.asp" -->
			<!-- #include file="../inc_val_date.asp" --> 
			<!-- #include file="../inc_ajax.asp" --> 
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

		Dim showDetails, cSDate, cEDate, cLoc, cTime, curARC, curPackage, curTime, curTG, curVT, curTGID, curTrn, curClassDescription, curTrnName
		Dim rowColor, tmpVisitCount, tmpRevPackage, tmpNumClass, tmpRevClass, tmpAvgRevClass, tmpAvgClientsClass
		Dim totNumClass, AvgRevClass, AvgClientsClass, totRevClass, totVisitCount, i, curClassID, daysOfWeek, ap_view_all_locs
		
		ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
		
		Function ARCDate()
			strSQLARC = "SELECT Max(DateCreated) as LastArcDate from dyntblAvgRevContr"
			rsARC.CursorLocation = 3
			rsARC.open strSQLARC, cnWS
			Set rsARC.ActiveConnection = Nothing
			ARCDate="N/A"
			If not rsARC.eof then
				if NOT isNULL(rsARC("LastArcDate")) then
					ARCDate=Cstr(rsARC("LastArcDate"))
				end if
			end if
			if ARCDate="N/A" then
			%>
			<script type="text/javascript">
				alert("Please run Average Revenue Contribution by Series report first to create contribution numbers.");
				javascript:history.go(-1);
			</script>
			<%
			end if
		end function
		
		Function ClassCount()
			if curTrn="" OR curClassID="" then
				ClassCount = 0
			else
				strSQLClass = "SELECT Classtime, ClassDate, Location FROM [VISIT DATA] "
				strSQLClass = strSQLClass & "WHERE ClassTime = " & TimeSepB & curTime & TimeSepA 
				strSQLClass = strSQLClass & " AND ([VISIT DATA].ClassDate >=  " & DateSep & cSDate & DateSep & ") "
				strSQLClass = strSQLClass & "AND ([VISIT DATA].ClassDate <=  " & DateSep & cEDate & DateSep & ") "  
				strSQLClass = strSQLClass & "AND ([VISIT DATA].TrainerID =  " & curTrn & ") "  
				strSQLClass = strSQLClass & "AND ([VISIT DATA].ClassID =  " & curClassID & ") "  
				if cLoc<>0 then
					strSQLClass = strSQLClass & " AND ([VISIT DATA].Location=" & cLoc & ") "
				end if
				if request.form("optTG")<>"0" AND request.form("optTG")<>"" then
					strSQLClass = strSQLClass & " AND [VISIT DATA].Typegroup=" & Cint(sqlInjectStr(request.form("optTG"))) & " "
				end if
				strSQLClass = strSQLClass & "GROUP BY ClassTime, Classdate, Location"
				'response.write "<br />" & strSQLClass & "<br />" 
				rsClassCount.CursorLocation = 3
				rsClassCount.open strSQLClass, cnWS
				Set rsClassCount.ActiveConnection = Nothing
				ClassCount=0
				'i=0
				If not rsClassCount.eof then
					'Do Until rsClassCount.eof
						'i = i + 1
						'rsClassCount.movenext
					'loop
					ClassCount = rsClassCount.recordCount
				end if
				rsClassCount.close
			end if
		end function

		if request.form("requiredtxtDateStart")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cSDate = CDATE(sqlInjectStr(request.form("requiredtxtDateStart")))
			Call SetLocale("en-us")
		else
			cSDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		end if
		if request.form("requiredtxtDateEnd")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cEDate = CDATE(sqlInjectStr(request.form("requiredtxtDateEnd")))
			Call SetLocale("en-us")
		else
			cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		end if
		If request.form("optLoc")<>"" then
			cLoc = CINT(sqlInjectStr(request.form("optLoc")))
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
		If request.form("optTime")<>"" and request.form("optTime")<>"0" then
		  if request.Form("optTime") = "null" then
		    cTime = "null"
		  else
			  cTime = cDate(sqlInjectStr(request.form("optTime")))
			end if
		else
			cTime = "0"
		end if
		showDetails = true
		if NOT request.form("frmExpReport")="true" then
		%>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_avg_rev_analysis", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %> 
			<script type="text/javascript">
			function exportReport() {
				document.frmParameter.frmExpReport.value = "true";
				document.frmParameter.frmGenReport.value = "true";
				<% iframeSubmit "frmParameter", "adm_rpt_avg_rev_analysis.asp" %>
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
			<span class="breadcrumb-item">&raquo;</span>
			<%if category <> "" then%>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<span class="breadcrumb-item">&raquo;</span>
			<%end if %>
		    <%=DisplayPhrase(reportPageTitlesDictionary, "Revenuebyclass")%>
			<% showNewHelpContentIcon("revenue-class-report") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%=DisplayPhrase(reportPageTitlesDictionary, "Addtofavorites")%></span>
			</div>
			</div>
		<%end if %>
			<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
				<tr> 
				<td valign="top" height="100%" width="100%"> 
				<table class="center" cellspacing="0" width="95%" height="100%">
					<tr>
					<td class="headText" align="left" valign="top">
					<table width="100%" cellspacing="0">
						<tr class="headText" height="30"  valign="middle">
						<% if NOT  UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
						<td><b><%= pp_PageTitle("Revenue By Class") %> </b>
						<!--JM - 49_2447-->
						<% showNewHelpContentIcon("revenue-class-report") %>

						</td>
						<%end if %>
						<td class="right"><b>Revenue Contribution based on Average Revenue Analysis as of: <%=ARCDate()%></b></td>
						</tr>
					</table>
					</td>
					</tr>
					<tr> 
					<td height="30"  valign="bottom" class="headText">
					<table class="mainText border4 center" cellspacing="0">
						<form name="frmParameter" action="adm_rpt_avg_rev_analysis.asp" method="POST">
						<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
							<% if category <> "" then %>
								<input type="hidden" name="category" id="category" value="<%=category %>" />
							<%end if %>
							<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
						<%end if %>
						<input type="hidden" name="frmGenReport" value="">
						<input type="hidden" name="frmExpReport" value="">
						<tr> 
						<td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b><span style="color:<%=session("pageColor4")%>;">&nbsp;</span>For Visits Between:&nbsp;
						<input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
						<script type="text/javascript">
						var cal1 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStart'});
						cal1.a_tpl.yearscroll = true;
						</script>
						&nbsp;And&nbsp; 
						<input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
						<script type="text/javascript">
						var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
						cal2.a_tpl.yearscroll = true;
						</script>
						&nbsp;
						<%=xssStr(allHotWords(8))%>:&nbsp;
						<select name="optLoc" onchange="document.frmParameter.submit();" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
						<option value="0" <% if cLoc=0 then response.write "selected" end if %>>All</option>
						<%
						strSQL = "SELECT LocationID, LocationName FROM Location WHERE [Active]=1 AND wsShow=1"
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
						&nbsp;Include $0 contributions&nbsp;<input type="checkbox" name="optIncludeZero" <%if request.form("optIncludeZero")="on" then response.write " checked" end if %>>
						<br />
						<select name="optDisMode">
							<option value="detail" <%if request.form("optDisMode")="detail" then response.write "selected" end if%>>Detail</option>
							<option value="summary" <%if request.form("optDisMode")="summary" then response.write "selected" end if%>>Summary</option>
						</select>
<%
						strSQL = "SELECT [VISIT DATA].TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName "
						strSQL = strSQL & "FROM dyntblAvgRevContr RIGHT OUTER JOIN "
						strSQL = strSQL & "[PAYMENT DATA] ON dyntblAvgRevContr.ProductID = [PAYMENT DATA].ProductID RIGHT OUTER JOIN "
						strSQL = strSQL & "[VISIT DATA] INNER JOIN "
						strSQL = strSQL & "tblClasses ON tblClasses.ClassID = [VISIT DATA].ClassID INNER JOIN "
						strSQL = strSQL & "tblClassDescriptions ON tblClasses.DescriptionID = tblClassDescriptions.ClassDescriptionID INNER JOIN "
						strSQL = strSQL & "TRAINERS ON TRAINERS.TrainerID = [VISIT DATA].TrainerID INNER JOIN "
						strSQL = strSQL & "tblTypeGroup ON tblTypeGroup.TypeGroupID = [VISIT DATA].TypeGroup ON [PAYMENT DATA].PmtRefNo = [VISIT DATA].PmtRefNo "
						strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >=  " & DateSep & cSDate & DateSep & ") AND ([VISIT DATA].ClassDate <=  " & DateSep & cEDate & DateSep & ") " 
						if cLoc<>0 then
							strSQL = strSQL & " AND ([VISIT DATA].Location=" & cLoc & ") "
						end if
						if cTime<>"0" and cTime<>"" then
						  if cTime = "null" then
							  strSQL = strSQL & " AND [VISIT DATA].ClassTime is NULL "
							else
							  strSQL = strSQL & " AND [VISIT DATA].ClassTime=" & TimeSepB & cTime & TimeSepA
						  end if
						end if
						if request.form("optTG")<>"0" AND request.form("optTG")<>"" then
							strSQL = strSQL & " AND [VISIT DATA].Typegroup=" & Cint(sqlInjectStr(request.form("optTG"))) & " "
						end if
						strSQL = strSQL & "GROUP BY [VISIT DATA].TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName "
						strSQL = strSQL & "ORDER BY TRAINERS.TrLastName, TRAINERS.TrFirstName, [VISIT DATA].TrainerID "
					response.write debugSQL(strSQL, "SQL")
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS, 0, 1
						Set rsEntry.ActiveConnection = Nothing %>

						<select name="optTrainer">
								<option value="0">All Trainers</option>
					<%	if NOT rsEntry.EOF then	%>							
							<% 	do while NOT rsEntry.EOF %>
								<option value="<%=rsEntry("TrainerID")%>"<% if clng(request.form("optTrainer"))=clng(rsEntry("TrainerID")) then response.write " selected" end if %>><%=FmtTrnNameNew(rsEntry, 0)%></option>
							<%		rsEntry.MoveNext
								loop %>
					<%	end if 
						rsEntry.close %>
						</select>
						<script type="text/javascript">
							document.frmParameter.optTrainer.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" " + '<%=jsEscSingle(allHotWords(113))%>';
						</script>
                        &nbsp;<%=xssStr(allHotWords(7))%>:
						<select name="optTG" onchange="document.frmParameter.submit();">
						<option value="0" <%if request.form("optTG")="0" then response.write "selected" end if%>>All Type Groups</option>
						<%
							strSQL = "SELECT TypegroupID, Typegroup FROM tblTypegroup "
							strSQL = strSQL & "WHERE [Active]=1 AND (wsReservation = 1 OR wsEnrollment = 1) "
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
						<script type="text/javascript">
							document.frmParameter.optTG.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(503))%>";
						</script>
						&nbsp;<%=xssStr(allHotWords(76))%>:&nbsp;
						<select name="optTime"><option value="0" <%if cTime="0" then response.write "selected" end if%>><%=xssStr(allHotWords(149))%>&nbsp;<%=xssStr(allHotWords(76))%></option>
						<%
							strSQL = "SELECT DISTINCT ClassTime "
							strSQL = strSQL & "FROM [VISIT DATA] "
							strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >=  " & DateSep & cSDate & DateSep & ") "
							strSQL = strSQL & "AND ([VISIT DATA].ClassDate <=  " & DateSep & cEDate & DateSep & ") "
							if cLoc<>0 then
								strSQL = strSQL & " AND ([VISIT DATA].Location=" & cLoc & ") "
							end if
							if request.form("optTG")<>"0" AND request.form("optTG")<>"" then
								strSQL = strSQL & " AND [VISIT DATA].Typegroup=" & Cint(sqlInjectStr(request.form("optTG"))) & " "
							end if
							strSQL = strSQL & "ORDER BY Classtime"
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>
								  <% if isNull(rsEntry("ClassTime")) then %>
								       <option value="null" <%if cTime = "null" then response.write "selected" end if %>>TBD</option>
								  <% else %>
									     <option value="<%=rsEntry("ClassTime")%>" <%if cstr(cTime)=cstr(CDate(rsEntry("Classtime"))) then response.write "selected" end if%>><%=FmtTimeShort(rsEntry("Classtime"))%></option>
								  <% end if %>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
						</select>&nbsp;	
						<br />	
						<% showDateArrows("frmParameter") %>
						<input type="button" name="Button" value="Generate" onClick="genReport();"></b>
						<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
						else%>
						<% exportToExcelButton %>
						<%end if%>
						<% savingButtons "frmParameter", "Revenue by Class" %>
						</td>
						</tr>
						</form>
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" id="revenueByClassGenTag" class="mainTextBig center-ch"> 
					
					<table class="mainText" width="100%" cellspacing="0">
						<tr > 
						<td  colspan="2" valign="top" class="mainTextBig center-ch">
		<% 
		end if			'end of frmExpreport value check before /head line	  
		%>
						<table class="mainText" width="95%"  cellspacing="0">
<%
							if request.form("frmGenReport")="true" then
								if request.form("frmExpReport")="true" then
									Dim stFilename
									stFilename="attachment; filename=Revenue Analysis for the Visits between " & Replace(cSDate,"/","-") & " and " & Replace(cEDate,"/","-") & ".xls" 
									Response.ContentType = "application/vnd.ms-excel" 
									Response.AddHeader "Content-Disposition", stFilename 
								end if
								
								strSQL = "SELECT [VISIT DATA].ClassTime, [VISIT DATA].Typetaken, [VISIT DATA].TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, tblTypegroup.TypegroupID, tblTypegroup.Typegroup, "
								strSQL = strSQL & " ISNULL(tblVisitTypes.TypeName, '') as VTName, dyntblAvgRevContr.AvgRevContr, Count(DISTINCT [VISIT DATA].VisitRefNo) AS VisitCount, tblClassDescriptions.ClassName, [VISIT DATA].ClassID, tblClasses.DaySunday, tblClasses.DayMonday, tblClasses.DayTuesday, tblClasses.DayWednesday, tblClasses.DayThursday, tblClasses.DayFriday, tblClasses.DaySaturday "
								strSQL = strSQL & "FROM dyntblAvgRevContr "
								strSQL = strSQL & "RIGHT OUTER JOIN [PAYMENT DATA] ON dyntblAvgRevContr.ProductID = [PAYMENT DATA].ProductID "
								strSQL = strSQL & "RIGHT OUTER JOIN [VISIT DATA] INNER JOIN tblClasses ON tblClasses.ClassID = [VISIT DATA].ClassID INNER JOIN tblClassDescriptions ON tblClasses.DescriptionID = tblClassDescriptions.ClassDescriptionID INNER JOIN TRAINERS ON TRAINERS.TrainerID = [VISIT DATA].TrainerID "
								strSQL = strSQL & "INNER JOIN tblTypeGroup ON dbo.tblTypeGroup.TypeGroupID = dbo.[VISIT DATA].TypeGroup "
								strSQL = strSQL & "ON [PAYMENT DATA].PmtRefNo = dbo.[VISIT DATA].PmtRefNo "
								strSQL = strSQL & " LEFT OUTER JOIN tblVisitTypes ON tblVisitTypes.TypeID = [VISIT DATA].VisitType "
								strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >=  " & DateSep & cSDate & DateSep & ") AND ([VISIT DATA].ClassDate <=  " & DateSep & cEDate & DateSep & ") " 
								if request.form("optIncludeZero")<>"on" then
									strSQL = strSQL & " AND ISNULL(dyntblAvgRevContr.AvgRevContr, 0) <> 0 "
								end if
								if cLoc<>0 then
									strSQL = strSQL & " AND ([VISIT DATA].Location=" & cLoc & ") "
								end if
								if request.form("optTrainer")<>"" and request.form("optTrainer")<>"0" then
									strSQL = strSQL & " AND TRAINERS.TrainerID = " & sqlInjectStr(request.form("optTrainer"))
								end if
								if cTime<>"0" and cTime<>"" then
								  if cTime = "null" then
								    strSQL = strSQL & " AND [VISIT DATA].ClassTime is NULL "
								  else
									  strSQL = strSQL & " AND [VISIT DATA].ClassTime=" & TimeSepB & cTime & TimeSepA
								  end if
								end if
								if request.form("optTG")<>"0" AND request.form("optTG")<>"" then
									strSQL = strSQL & " AND [VISIT DATA].Typegroup=" & Cint(sqlInjectStr(request.form("optTG"))) & " "
								end if
								strSQL = strSQL & " GROUP BY  [VISIT DATA].ClassTime, [VISIT DATA].Typetaken, tblTypegroup.TypegroupID, tblTypegroup.Typegroup, dyntblAvgRevContr.AvgRevContr, [VISIT DATA].TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, tblClassDescriptions.ClassName, [VISIT DATA].ClassID,  tblClasses.DaySunday, tblClasses.DayMonday, tblClasses.DayTuesday, tblClasses.DayWednesday, tblClasses.DayThursday, tblClasses.DayFriday, tblClasses.DaySaturday, tblVisitTypes.TypeName "
								strSQL = strSQL & " ORDER BY tblTypegroup.Typegroup, [VISIT DATA].ClassTime, [VISIT DATA].TrainerID, [VISIT DATA].ClassID "
							response.write debugSQL(strSQL, "SQL")
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS, 0, 1
								Set rsEntry.ActiveConnection = Nothing
	
								curTrnName=""
								curTrn=""
								curClassDescription=""
								curClassID=""
								curtime=""
								curARC=0
								curPackage=""
								curTG="" 
								curTGID=0
								tmpVisitCount=0
								tmpRevPackage=0
								tmpNumClass=0
								tmpRevClass=0
								tmpAvgRevClass=0
								tmpAvgClientsClass=0
								totNumClass=0
								AvgRevClass=0
								AvgClientsClass=0
								totRevClass=0	
								totVisitCount=0			
		
								if NOT rsEntry.EOF then		'EOF
									if request.form("optDisMode")="detail" then
									%>
										<tr>
											<td colspan="4">&nbsp; </td>
										</tr>
									<%	do while NOT rsEntry.EOF
										if (rsEntry("AvgRevContr")<>0 AND NOT isNull(rsEntry("AvgRevContr")) OR request.form("optIncludeZero")="on") then
											If (curTrn<>cstr(rsEntry("TrainerID")) and curTrn<>"") or curTime<>rsEntry("ClassTime") or curClassID<>cstr(rsEntry("ClassID")) then 'curTG<>rsEntry("Typegroup")  and curTG<>"" then
												if (curTime<> rsEntry("ClassTime") or curTrn<>cstr(rsEntry("TrainerID")) or curClassID<>cstr(rsEntry("ClassiD")))  then
													If curTime<>"" then
														tmpNumClass=ClassCount()							
														totNumClass = totNumClass + tmpNumClass
														if tmpNumClass<>0 then
															if NOT request.form("frmExpReport")="true" then
																if NOT request.form("frmExpReport")="true" then
																	tmpAvgRevClass=FmtCurrency(Round(tmpRevClass/tmpNumClass,2))
																else
																	tmpAvgRevClass=FmtNumber(Round(tmpRevClass/tmpNumClass,2))
																end if
															else
																tmpAvgRevClass=Round(tmpRevClass/tmpNumClass,2)
															end if 
															tmpAvgClientsClass=Round(tmpVisitCount/tmpNumClass,0)
														else
															if NOT request.form("frmExpReport")="true" then
																if NOT request.form("frmExpReport")="true" then
																	tmpAvgRevClass=FmtCurrency(0)
																else
																	tmpAvgRevClass=FmtNumber(0)
																end if
															else
																tmpAvgRevClass=0
															end if 
															tmpAvgClientsClass=0
														end if
														if totNumClass<>0 then
															if NOT request.form("frmExpReport")="true" then
																AvgRevClass=FmtCurrency(Round(totRevClass/totNumClass,2))
															else
																AvgRevClass=FmtNumber(Round(totRevClass/totNumClass,2))
															end if
															AvgClientsClass=Round(totVisitCount/totNumClass,0)
														else
															if NOT request.form("frmExpReport")="true" then
																AvgRevClass=FmtCurrency(0)
															else
																AvgRevClass=FmtNumber(0)
															end if
															AvgClientsClass=0
														end if
													%>
	<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														<td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
													</tr>
	<% end if %>
														<tr class="right" >
														  <td colspan="3"><strong><%=FmtTimeShort(curTime)%>&nbsp;<%=curClassDescription%>&nbsp; w/ <%=curTrnName%> &nbsp; Total Revenue:</strong></td>
														<% if NOT request.form("frmExpReport")="true" then %>
														  <td><strong><%=FmtCurrency(tmpRevClass)%></strong></td>
														<% else %>
														  <td><strong><%=FmtNumber(tmpRevClass)%></strong></td>
														<% end if %>
														</tr>
														<tr class="right" >
														  <td colspan="3"><strong>Total Classes in Period:</strong></td>
														  <td><strong><%=tmpNumClass%></strong></td>
														</tr>
														<tr class="right" >
														  <td colspan="3"><strong>Total # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s in Class:</strong></td>
														  <td><strong><%=tmpVisitCount%></strong></td>
														</tr>
														<tr class="right" >
														  <td colspan="3"><strong>Avg Revenue per Class:</strong></td>
														  <td><strong><%=tmpAvgRevClass%></strong></td>
														</tr>
														<tr class="right" >
														  <td colspan="3"><strong>Avg # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s per Class:</strong></td>
														  <td><strong><%=tmpAvgClientsClass%></strong></td>
														</tr>
												<tr>
													<td colspan="4">&nbsp; </td>
												</tr>
													<%
													end if	'curTime<>""
													if curTG<>rsEntry("TypeGroup") and curTG<>"" then
											%>
	<% if NOT request.form("frmExpReport")="true" then %>
												<tr height="2">
													<td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
												</tr>
	<% end if %>
													<tr align="left">
													  <td rowspan="6"><strong>Average For ALL &nbsp;<%=UCASE(curTG)%>&nbsp;<%=xssStr(allHotWords(5))%></strong></td>
													</tr>
													<tr class="right" >
													  <td colspan="2"><strong>Total Revenue:</strong></td>
												<% if NOT request.form("frmExpReport")="true" then %>
													  <td><strong><%=FmtCurrency(totRevClass)%></strong></td>
												<% else %>
													  <td><strong><%=FmtNumber(totRevClass)%></strong></td>
												<% end if %>
													</tr>
													<tr class="right" >
													  <td colspan="2"><strong>Total Classes in Period:</strong></td>
													  <td><strong><%=totNumClass%></strong></td>
													</tr>
													<tr class="right" >
													  <td colspan="2"><strong>Total # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s in Classes:</strong></td>
													  <td><strong><%=totVisitCount%></strong></td>
													</tr>
													<tr class="right" >
													  <td colspan="2"><strong>Avg Revenue per Class:</strong></td>
													  <td><strong><%=AvgRevClass%></strong></td>
													</tr>
													<tr class="right" >
													  <td colspan="2"><strong>Avg # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s per Class:</strong></td>
													  <td><strong><%=AvgClientsClass%></strong></td>
													</tr>
	<% if NOT request.form("frmExpReport")="true" then %>
												<tr height="2">
													<td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
												</tr>
	<% end if %>
											<% 
												totNumClass=0
												AvgRevClass=0
												AvgClientsClass=0
												totRevClass=0	
												totVisitCount=0
	
											end if ' curTG<>rsEntry("TypeGroupID") %>
												<tr>
													<td colspan="4">&nbsp; </td>
												</tr>
													<tr>
													<% 
														daysOfWeek = ""
														if rsEntry("DaySunday") then
															daysOfWeek = daysOfWeek & "Su"
														end if
														if rsEntry("DayMonday") then
															daysOfWeek = daysOfWeek & "M"
														end if
														if rsEntry("DayTuesday") then
															daysOfWeek = daysOfWeek & "T"
														end if
														if rsEntry("DayWednesday") then
															daysOfWeek = daysOfWeek & "W"
														end if
														if rsEntry("DayThursday") then
															daysOfWeek = daysOfWeek & "Th"
														end if
														if rsEntry("DayFriday") then
															daysOfWeek = daysOfWeek & "F"
														end if
														if rsEntry("DaySaturday") then
															daysOfWeek = daysOfWeek & "Sa"
														end if
													%>
													  <td colspan="4" class="whiteHeader" style="background-color:<%=session("pageColor4")%>;" align="left">&nbsp;<%if isNull(rsEntry("ClassTime")) then response.Write "TBD" else response.Write FmtTimeShort(rsEntry("ClassTime")) end if%>&nbsp;<%=rsEntry("ClassName")%> w/ <%=FmtTrnNameNew(rsEntry, 0)%>&nbsp;&nbsp;(<%=daysOfWeek%>)</td>
													</tr>
													<tr class="right">
														<td width="30%" align="left"><strong><%=xssStr(allHotWords(61))%></strong></td>
														<td width="10%"><strong>Total <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %>Visits&nbsp;</strong></td>
														<td width="10%"><strong>Rev Contributed&nbsp;</strong></td>
														<td width="10%"><strong><%= getHotWord(118)%>&nbsp;</strong></td>
													</tr>
	<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="2">
														<td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
													</tr>
	<% end if %>
												<%
													tmpVisitCount=0
													tmpRevPackage=0
													tmpNumClass=0
													tmpRevClass=0
													tmpAvgRevClass=0
													tmpAvgClientsClass=0
												end if	'curTime<>rsEntry("ClassTime")
											else		'curTG else
												if curTime<> rsEntry("ClassTime") then
													If curTime<>"" then
														tmpNumClass=ClassCount()
														totNumClass = totNumClass + tmpNumClass
														if tmpNumClass<>0 then
															if NOT request.form("frmExpReport")="true" then
																tmpAvgRevClass=FmtCurrency(Round(tmpRevClass/tmpNumClass,2))
															else 
																tmpAvgRevClass=FmtNumber(Round(tmpRevClass/tmpNumClass,2))
															end if
															tmpAvgClientsClass=Round(tmpVisitCount/tmpNumClass,0)
														else
															if NOT request.form("frmExpReport")="true" then
																tmpAvgRevClass=FmtCurrency(0)
															else
																tmpAvgRevClass=FmtNumber(0)
															end if
															tmpAvgClientsClass=0
														end if
													%>
	<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="1">
														<td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
													</tr>
	<% end if %>
														<tr class="right" >
														  <td colspan="3"><strong><%=FmtTimeShort(curTime)%>&nbsp; <%=curClassDescription%> w/ <%=curTrnName%> &nbsp;Total Revenue:</strong></td>
													<% if NOT request.form("frmExpReport")="true" then %>
														  <td><strong><%=FmtCurrency(tmpRevClass)%></strong></td>
													<% else %>
														  <td><strong><%=FmtNumber(tmpRevClass)%></strong></td>
													<% end if %>
														</tr>
														<tr class="right" >
														  <td colspan="3"><strong>Total Classes in Period:</strong></td>
														  <td><strong><%=tmpNumClass%></strong></td>
														</tr>
														<tr class="right" >
														  <td colspan="3"><strong>Total # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s in Class:</strong></td>
														  <td><strong><%=tmpVisitCount%></strong></td>
														</tr>
														<tr class="right" >
														  <td colspan="3"><strong>Avg Revenue per Class:</strong></td>
														  <td><strong><%=tmpAvgRevClass%></strong></td>
														</tr>
														<tr class="right" >
														  <td colspan="3"><strong>Avg # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s per Class:</strong></td>
														  <td><strong><%=tmpAvgClientsClass%></strong></td>
														</tr>
												<tr>
													<td colspan="4">&nbsp; </td>
												</tr>
													<%
													end if	'curTime<>""
												%>
													<tr>
													<% 
														daysOfWeek = ""
														if rsEntry("DaySunday") then
															daysOfWeek = daysOfWeek & "Su"
														end if
														if rsEntry("DayMonday") then
															daysOfWeek = daysOfWeek & "M"
														end if
														if rsEntry("DayTuesday") then
															daysOfWeek = daysOfWeek & "T"
														end if
														if rsEntry("DayWednesday") then
															daysOfWeek = daysOfWeek & "W"
														end if
														if rsEntry("DayThursday") then
															daysOfWeek = daysOfWeek & "Th"
														end if
														if rsEntry("DayFriday") then
															daysOfWeek = daysOfWeek & "F"
														end if
														if rsEntry("DaySaturday") then
															daysOfWeek = daysOfWeek & "Sa"
														end if
													%>
													  <td colspan="4" class="whiteHeader" style="background-color:<%=session("pageColor4")%>;" align="left">&nbsp;<%=FmtTimeShort(rsEntry("ClassTime"))%>&nbsp;<%=rsEntry("ClassName")%> w/ <%=FmtTrnNameNew(rsEntry, 0)%> (<%=daysOfWeek%>)</td>
													</tr>
													<tr class="right">
														<td width="30%" align="left"><strong><%=xssStr(allHotWords(61))%></strong></td>
														<td width="10%"><strong>Total <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %>Visits&nbsp;</strong></td>
														<td width="10%"><strong>Rev Contributed&nbsp;</strong></td>
														<td width="10%"><strong><%= getHotWord(118)%>&nbsp;</strong></td>
													</tr>
	<% if NOT request.form("frmExpReport")="true" then %>
													<tr height="2">
														<td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
													</tr>
	<% end if %>
												<%
													tmpVisitCount=0
													tmpRevPackage=0
													tmpNumClass=0
													tmpRevClass=0
													tmpAvgRevClass=0
													tmpAvgClientsClass=0
												end if	'curTime<>rsEntry("ClassTime")
											end if		'curTG
											curtime = rsEntry("Classtime")
											curARC = rsEntry("AvgRevContr")
											curPackage = rsEntry("Typetaken")
											curTg = rsEntry("Typegroup")
											curTGID = rsEntry("TypegroupID")
											curTrn = cstr(rsEntry("TrainerID"))
											curClassDescription = rsEntry("ClassName")
											curTrnName = FmtTrnNameNew(rsEntry, 0)
											curClassID = cstr(rsEntry("ClassID"))
											
											tmpRevPackage = curARC * rsEntry("VisitCount")
											if NOT isNull(tmpRevPackage) then
												tmpRevClass=tmpRevClass + tmpRevPackage
											end if
											tmpVisitCount = tmpVisitCount + rsEntry("VisitCount")
											if NOT isNull(tmpRevPackage) then
												totRevClass = totRevClass + tmpRevPackage
											end if
											totVisitCount = totVisitCount + rsEntry("VisitCount")
											
											if rowColor = "#F2F2F2" then
												rowColor = "#FAFAFA"
											else
												rowColor = "#F2F2F2"
											end if
											%>
												<tr class="right" style="background-color:<%=rowcolor%>;">
												  <td align="left"><%=curPackage%></td>
												  <td><%=rsEntry("VisitCount")%></td>
												<% if NOT request.form("frmExpReport")="true" then %>
												  <td><%=FmtCurrency(curARC)%></td>
												  <td><%=FmtCurrency(tmpRevPackage)%></td>
												<% else %>
												  <td><%=FmtNumber(curARC)%></td>
												  <td><%=FmtNumber(tmpRevPackage)%></td>
												<% end if %>
												</tr>
											<%
											end if
											rsEntry.MoveNext
										loop
										tmpNumClass=ClassCount()
										totNumClass = totNumClass + tmpNumClass
										if tmpNumClass<>0 then
											if NOT request.form("frmExpReport")="true" then
												tmpAvgRevClass=FmtCurrency(Round(tmpRevClass/tmpNumClass,2))
											else
												tmpAvgRevClass=FmtNumber(Round(tmpRevClass/tmpNumClass,2))
											end if
											tmpAvgClientsClass=Round(tmpVisitCount/tmpNumClass,0)
										else
											if NOT request.form("frmExpReport")="true" then
												tmpAvgRevClass=FmtCurrency(0)
											else
												tmpAvgRevClass=FmtNumber(0)
											end if
											tmpAvgClientsClass=0
										end if
										if totNumClass<>0 then
											if NOT request.form("frmExpReport")="true" then
												AvgRevClass=FmtCurrency(Round(totRevClass/totNumClass,2))
											else
												AvgRevClass=FmtNumber(Round(totRevClass/totNumClass,2))
											end if
											AvgClientsClass=Round(totVisitCount/totNumClass,0)
										else
											if NOT request.form("frmExpReport")="true" then
												AvgRevClass=FmtCurrency(0)
											else
												AvgRevClass=FmtNumber(0)
											end if 
											AvgClientsClass=0
										end if
										%>
	<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="1">
											<td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
	<% end if %>
											<tr class="right" >
											  <td colspan="3"><strong><%=FmtTimeShort(curTime)%>&nbsp;<%=curClassDescription%> w/ <%=curTrnName%>&nbsp;Total Revenue:</strong></td>
	<% if NOT request.form("frmExpReport")="true" then %>
											  <td><strong><%=FmtCurrency(tmpRevClass)%></strong></td>
	<% else %>
											  <td><strong><%=FmtNumber(tmpRevClass)%></strong></td>
	<% end if %>
											</tr>
											<tr class="right" >
											  <td colspan="3"><strong>Total Classes in Period:</strong></td>
											  <td><strong><%=tmpNumClass%></strong></td>
											</tr>
											<tr class="right" >
											  <td colspan="3"><strong>Total # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s in Class:</strong></td>
											  <td><strong><%=tmpVisitCount%></strong></td>
											</tr>
											<tr class="right" >
											  <td colspan="3"><strong>Avg Revenue per Class:</strong></td>
											  <td><strong><%=tmpAvgRevClass%></strong></td>
											</tr>
											<tr class="right" >
											  <td colspan="3"><strong>Avg # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s per Class:</strong></td>
											  <td><strong><%=tmpAvgClientsClass%></strong></td>
											</tr>
										<tr>
											<td colspan="4">&nbsp; </td>
										</tr>
	<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="2">
											<td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
	<% end if %>
											<tr align="left" >
											  <td rowspan="6"><strong>Average For ALL &nbsp;<%=UCASE(curTG)%>&nbsp;<%=xssStr(allHotWords(5))%></strong></td>
											</tr>
											<tr class="right" >
											  <td colspan="2"><strong>Total Revenue:</strong></td>
	<% if NOT request.form("frmExpReport")="true" then %>
											  <td><strong><%=FmtCurrency(totRevClass)%></strong></td>
	<% else %>
											  <td><strong><%=FmtNumber(totRevClass)%></strong></td>
	<% end if %>
											</tr>
											<tr class="right" >
											  <td colspan="2"><strong>Total Classes in Period:</strong></td>
											  <td><strong><%=totNumClass%></strong></td>
											</tr>
											<tr class="right" >
											  <td colspan="2"><strong>Total # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s in Classes:</strong></td>
											  <td><strong><%=totVisitCount%></strong></td>
											</tr>
											<tr class="right" >
											  <td colspan="2"><strong>Avg Revenue per Class:</strong></td>
											  <td><strong><%=AvgRevClass%></strong></td>
											</tr>
											<tr class="right" >
											  <td colspan="2"><strong>Avg # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s per Class:</strong></td>
											  <td><strong><%=AvgClientsClass%></strong></td>
											</tr>
	<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="2">
											<td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
	<% end if %>
										<tr>
											<td colspan="4">&nbsp; </td>
										</tr>
										<%
									else ' optDisMode = "summary"
									%>	
										<tr><td colspan="10">&nbsp;</td></tr>
										<tr style="background-color:<%=session("pageColor4")%>;">
											<td align="left" class="whiteHeader"><%= getHotWord(58)%></td>
											<td align="left" class="whiteHeader">Day</td>
											<td>&nbsp;</td>
											<td align="left" width="20%" class="whiteHeader">Class Name</td>
											<td align="left" class="whiteHeader"><%= getHotWord(6)%></td>
											<td align="left" class="whiteHeader"><%= getHotWord(2)%></td>
											<td align="left" width="10%" class="whiteHeader"><%= getHotWord(7)%></td>
											<td class="right whiteHeader">Total Revenue</td>
											<td class="right whiteHeader">Total Classes</td>
											<td class="right whiteHeader">Total # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s</td>
											<td class="right whiteHeader">Avg Revenue per Class</td>
											<td class="right whiteHeader">Avg # of <%if request.form("optIncludeZero")<>"on" then response.write "Paid " end if %><%=session("ClientHW")%>s per Class</td>
										</tr>									
									<%	do while NOT rsEntry.EOF
										if (rsEntry("AvgRevContr")<>0 AND NOT isNull(rsEntry("AvgRevContr")) OR request.form("optIncludeZero")="on") then
											If (curTrn<>cstr(rsEntry("TrainerID")) and curTrn<>"") or curTime<>rsEntry("ClassTime") or curClassID<>cstr(rsEntry("ClassID")) then 'curTG<>rsEntry("Typegroup")  and curTG<>"" then
												if (curTime<> rsEntry("ClassTime") or curTrn<>cstr(rsEntry("TrainerID")) or curClassID<>cstr(rsEntry("ClassiD")))  then
													If curTime<>"" then
														tmpNumClass=ClassCount()							
														totNumClass = totNumClass + tmpNumClass
														if tmpNumClass<>0 then
															if NOT request.form("frmExpReport")="true" then
																if NOT request.form("frmExpReport")="true" then
																	tmpAvgRevClass=FmtCurrency(Round(tmpRevClass/tmpNumClass,2))
																else
																	tmpAvgRevClass=FmtNumber(Round(tmpRevClass/tmpNumClass,2))
																end if
															else
																tmpAvgRevClass=Round(tmpRevClass/tmpNumClass,2)
															end if 
															tmpAvgClientsClass=Round(tmpVisitCount/tmpNumClass,0)
														else
															if NOT request.form("frmExpReport")="true" then
																if NOT request.form("frmExpReport")="true" then
																	tmpAvgRevClass=FmtCurrency(0)
																else
																	tmpAvgRevClass=FmtNumber(0)
																end if
															else
																tmpAvgRevClass=0
															end if 
															tmpAvgClientsClass=0
														end if
														if totNumClass<>0 then
															if NOT request.form("frmExpReport")="true" then
																AvgRevClass=FmtCurrency(Round(totRevClass/totNumClass,2))
															else
																AvgRevClass=FmtNumber(Round(totRevClass/totNumClass,2))
															end if
															AvgClientsClass=Round(totVisitCount/totNumClass,0)
														else
															if NOT request.form("frmExpReport")="true" then
																AvgRevClass=FmtCurrency(0)
															else
																AvgRevClass=FmtNumber(0)
															end if
															AvgClientsClass=0
														end if
													%>
										<tr>
											<td class="mainText"><%=FmtTimeShort(curTime)%></td>
											<td class="center-ch mainText"><%=daysOfWeek%></td>
											<td>&nbsp;</td>
											<td class="mainText"><%=curClassDescription%></td>
											<td class="mainText"><%=curTrnName%></td>
											<td class="mainText"><%=curVT%></td>
											<td class="mainText"><%=curTG%></td>
											<% if NOT request.form("frmExpReport")="true" then %>
											  <td class="right mainText"><%=FmtCurrency(tmpRevClass)%></td>
											<% else %>
											  <td class="right mainText"><%=FmtNumber(tmpRevClass)%></td>
											<% end if %>
											<td class="right mainText"><%=tmpNumClass%></td>
											<td class="right mainText"><%=tmpVisitCount%></td>
											<td class="right mainText"><%=tmpAvgRevClass%></td>
											<td class="right mainText"><%=tmpAvgClientsClass%></td>
										</tr>
													<%
													end if	'curTime<>""
													if curTG<>rsEntry("TypeGroup") and curTG<>"" then
														totNumClass=0
														AvgRevClass=0
														AvgClientsClass=0
														totRevClass=0	
														totVisitCount=0
	
													end if ' curTG<>rsEntry("TypeGroupID")
													daysOfWeek = ""
													if rsEntry("DaySunday") then
														daysOfWeek = daysOfWeek & "Su"
													end if
													if rsEntry("DayMonday") then
														daysOfWeek = daysOfWeek & "M"
													end if
													if rsEntry("DayTuesday") then
														daysOfWeek = daysOfWeek & "T"
													end if
													if rsEntry("DayWednesday") then
														daysOfWeek = daysOfWeek & "W"
													end if
													if rsEntry("DayThursday") then
														daysOfWeek = daysOfWeek & "Th"
													end if
													if rsEntry("DayFriday") then
														daysOfWeek = daysOfWeek & "F"
													end if
													if rsEntry("DaySaturday") then
														daysOfWeek = daysOfWeek & "Sa"
													end if
													tmpVisitCount=0
													tmpRevPackage=0
													tmpNumClass=0
													tmpRevClass=0
													tmpAvgRevClass=0
													tmpAvgClientsClass=0
												end if	'curTime<>rsEntry("ClassTime")
											else		'curTG else
												if curTime<> rsEntry("ClassTime") then
													If curTime<>"" then
														tmpNumClass=ClassCount()
														totNumClass = totNumClass + tmpNumClass
														if tmpNumClass<>0 then
															if NOT request.form("frmExpReport")="true" then
																tmpAvgRevClass=FmtCurrency(Round(tmpRevClass/tmpNumClass,2))
															else 
																tmpAvgRevClass=FmtNumber(Round(tmpRevClass/tmpNumClass,2))
															end if
															tmpAvgClientsClass=Round(tmpVisitCount/tmpNumClass,0)
														else
															if NOT request.form("frmExpReport")="true" then
																tmpAvgRevClass=FmtCurrency(0)
															else
																tmpAvgRevClass=FmtNumber(0)
															end if
															tmpAvgClientsClass=0
														end if
													%>
										<tr>
											<td class="mainText"><%=FmtTimeShort(curTime)%></td>
											<td class="mainText center-ch"><%=daysOfWeek%></td>
											<td class="mainText"><%=curClassDescription%></td>
											<td class="mainText"><%=curTrnName%></td>
											<td class="mainText"><%=curVT%></td>
											<td class="mainText"><%=curTG%></td>
											<% if NOT request.form("frmExpReport")="true" then %>
											  <td class="right mainText"><%=FmtCurrency(tmpRevClass)%></td>
											<% else %>
											  <td  class="right mainText"><%=FmtNumber(tmpRevClass)%></td>
											<% end if %>
											<td class="right mainText"><%=tmpNumClass%></td>
											<td class="right mainText"><%=tmpVisitCount%></td>
											<td class="right mainText"><%=tmpAvgRevClass%></td>
											<td class="right mainText"><%=tmpAvgClientsClass%></td>
										</tr>
													<%
													end if	'curTime<>""
													daysOfWeek = ""
													if rsEntry("DaySunday") then
														daysOfWeek = daysOfWeek & "Su"
													end if
													if rsEntry("DayMonday") then
														daysOfWeek = daysOfWeek & "M"
													end if
													if rsEntry("DayTuesday") then
														daysOfWeek = daysOfWeek & "T"
													end if
													if rsEntry("DayWednesday") then
														daysOfWeek = daysOfWeek & "W"
													end if
													if rsEntry("DayThursday") then
														daysOfWeek = daysOfWeek & "Th"
													end if
													if rsEntry("DayFriday") then
														daysOfWeek = daysOfWeek & "F"
													end if
													if rsEntry("DaySaturday") then
														daysOfWeek = daysOfWeek & "Sa"
													end if
													tmpVisitCount=0
													tmpRevPackage=0
													tmpNumClass=0
													tmpRevClass=0
													tmpAvgRevClass=0
													tmpAvgClientsClass=0
												end if	'curTime<>rsEntry("ClassTime")
											end if		'curTG
											curtime = rsEntry("Classtime")
											curARC = rsEntry("AvgRevContr")
											curPackage = rsEntry("Typetaken")
											curTg = rsEntry("Typegroup")
											curVT = rsEntry("VTName")
											curTGID = rsEntry("TypegroupID")
											curTrn = cstr(rsEntry("TrainerID"))
											curClassDescription = rsEntry("ClassName")
											curTrnName = FmtTrnNameNew(rsEntry, 0)
											curClassID = cstr(rsEntry("ClassID"))
											
											tmpRevPackage = curARC * rsEntry("VisitCount")
											if NOT isNull(tmpRevPackage) then
												tmpRevClass=tmpRevClass + tmpRevPackage
											end if
											tmpVisitCount = tmpVisitCount + rsEntry("VisitCount")
											if NOT isNull(tmpRevPackage) then
												totRevClass = totRevClass + tmpRevPackage
											end if
											totVisitCount = totVisitCount + rsEntry("VisitCount")
											
											if rowColor = "#F2F2F2" then
												rowColor = "#FAFAFA"
											else
												rowColor = "#F2F2F2"
											end if
											end if
											rsEntry.MoveNext
										loop
										tmpNumClass=ClassCount()
										totNumClass = totNumClass + tmpNumClass
										if tmpNumClass<>0 then
											if NOT request.form("frmExpReport")="true" then
												tmpAvgRevClass=FmtCurrency(Round(tmpRevClass/tmpNumClass,2))
											else
												tmpAvgRevClass=FmtNumber(Round(tmpRevClass/tmpNumClass,2))
											end if
											tmpAvgClientsClass=Round(tmpVisitCount/tmpNumClass,0)
										else
											if NOT request.form("frmExpReport")="true" then
												tmpAvgRevClass=FmtCurrency(0)
											else
												tmpAvgRevClass=FmtNumber(0)
											end if
											tmpAvgClientsClass=0
										end if
										if totNumClass<>0 then
											if NOT request.form("frmExpReport")="true" then
												AvgRevClass=FmtCurrency(Round(totRevClass/totNumClass,2))
											else
												AvgRevClass=FmtNumber(Round(totRevClass/totNumClass,2))
											end if
											AvgClientsClass=Round(totVisitCount/totNumClass,0)
										else
											if NOT request.form("frmExpReport")="true" then
												AvgRevClass=FmtCurrency(0)
											else
												AvgRevClass=FmtNumber(0)
											end if 
											AvgClientsClass=0
										end if
										%>
										<tr>
											<td class="mainText"><%=FmtTimeShort(curTime)%></td>
											<td class="center-ch mainText"><%=daysOfWeek%></td>
											<td>&nbsp;</td>
											<td class="mainText"><%=curClassDescription%></td>
											<td class="mainText"><%=curTrnName%></td>
											<td class="mainText"><%=curVT%></td>
											<td class="mainText"><%=curTG%></td>
											<% if NOT request.form("frmExpReport")="true" then %>
											  <td class="right mainText"><%=FmtCurrency(tmpRevClass)%></td>
											<% else %>
											  <td class="right mainText"><%=FmtNumber(tmpRevClass)%></td>
											<% end if %>
											<td class="right mainText"><%=tmpNumClass%></td>
											<td class="right mainText"><%=tmpVisitCount%></td>
											<td class="right mainText"><%=tmpAvgRevClass%></td>
											<td class="right mainText"><%=tmpAvgClientsClass%></td>
										</tr>
									<%	end if ' summary vs details
									end if	' EOF
								rsEntry.close
								set rsEntry = nothing
							end if		'end of generate report if statement
							%>
						  </table>
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
