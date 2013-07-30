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
	%>
		<!-- #include file="inc_accpriv.asp" -->
		<%	dim doRefresh : doRefresh = false %>
		<!-- #include file="inc_date_arrows.asp" -->
		<!-- #include file="inc_rpt_tagging.asp" -->
		<!-- #include file="inc_utilities.asp" -->
		<!-- #include file="inc_rpt_save.asp" -->
<%

	set rsEntry = Server.CreateObject("ADODB.Recordset")

	dim ap_rpt_day, ap_rpt_day_self, ap_view_all_locs, clsStartTime
	ap_rpt_day = validAccessPriv("RPT_DAY")
	ap_rpt_day_self = validAccessPriv("RPT_DAY_SELF")
	ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
	
	if not Session("Pass") OR Session("Admin")="false" OR NOT (ap_rpt_day OR ap_rpt_day_self) then 
		%>
		<script type="text/javascript">
		    alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
		    javascript: history.go(-1);
		</script>
		<% 
	elseif ap_rpt_day_self AND NOT Session("Admin") = "owner" AND NOT Session("Admin") = "sa" AND session("empID") = "" AND NOT ap_rpt_day then
		%>
		<script type="text/javascript">
			alert("This report requires you to assign this username to a staff member.\nPlease have a site administrator do so in the Users and Groups page.");
			javascript: history.go(-1);
		</script>
		<% 
	else 
		%>
				<!-- #include file="inc_help_content.asp" -->
		<% if request.form("frmExpReport")<>"true" then %>
				
				<!-- #include file="../inc_loading.asp" -->
		<% end if %>
				<!-- #include file="../inc_i18n.asp" -->
				<!-- #include file="inc_acct_balance.asp" -->
				<!-- #include file="inc_visit_status.asp" -->
				<!-- #include file="../inc_ajax.asp" -->
				<!-- #include file="../inc_val_date.asp" -->
				<!-- #include file="inc_hotword.asp" -->
			<%
			Dim thisDate, curDate, curPageDate, schGlanceAcctBal, disLocationName, disTrainerID, disClientID, disPhone, disUnpaidAppointments, disBalance
			Dim disCellPhone, disWorkPhone, DisTrainerName, disClientName, disWS, disClassDate, disTG, disClassName, disClassID, disStatus
			Dim disColor, trowColor, disVT, curClassID, disNotes, disEmpID, disCrFirst, disCrLast, opener, intCount
			Dim varLink, disStartTime, disEndTime, curTrnID, curTime, trnName, cLocName, noViewOpt, intdaysave
			Dim rowcount, first, intLoopControl, cont, curTrn, curTG, curVT, rsPmtData, disAlert, disStaffAlert, disRSSID
			Dim tmpIsMakeUp, disIsNewClient, disBirthdate, intDays, strDays, trnStr, numWeeks, rsTrn, displaying, schGlanceRem
			Dim cLoc, cTG, cTrn, cView, rsEntry, Booked_HW, SignedIn_HW, Confirmed_HW, Confirmed2_HW, genReport, schGlanceShowCltPhone
			Dim ss_CltUseCompany, ss_CheckActivationDates, schGlanceCreatedBy, curStatus
			Dim cSDate, cEDate, mbfId, ss_ccs_cltPackageSharing

			displaying = ""

			Dim filterByCreated, disCreationDate, curCreationDate, nextCreationDate, nextClassDate, dateFilterSql, displayed 
			filterByCreated = request.form("optfilterByCreated") 

			if filterByCreated="1" then nextCreationDate = true

			function filterDateSQL(table) 
				if filterByCreated="1" then
					filterDateSQL = "CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, " & table & ".CreationDateTime))) " ''Converts time of CreationDateTime to 00:00:00.000 for comparisons with other dates
				else
					filterDateSQL = table & ".ClassDate "
				end if
			end function
			function orderByDateSQL(table)
				if filterByCreated="1" then
					orderByDateSQL = filterDateSQL(table) & ", "
				else
					orderByDateSQL = ""
				end if
			end function
			
			dim rsOpts, strOpts
			set rsOpts = Server.CreateObject("ADODB.Recordset")	
			
			strOpts = "SELECT schGlanceAcctBal, schGlanceRem, schGlanceShowCltPhone, CheckActivationDates, SchedGlanceCreatedBy, cltPackageSharing FROM tblGenOpts "
			rsOpts.CursorLocation = 3
			rsOpts.open strOpts, cnWS
			Set rsOpts.ActiveConnection = Nothing
			
			schGlanceAcctBal = rsOpts("schGlanceAcctBal")
			schGlanceRem = rsOpts("schGlanceRem")
			schGlanceShowCltPhone = rsOpts("schGlanceShowCltPhone")
			ss_CheckActivationDates = rsOpts("CheckActivationDates")
			schGlanceCreatedBy = rsOpts("SchedGlanceCreatedBy")
			ss_ccs_cltPackageSharing = rsOpts("cltPackageSharing")

			rsOpts.close
			
			dim hw6, hw8, hw10, hw16, hw11, hw9
			'dim arrHW : arrHW = getHotWords(array(6,8,10,16,11,9))
			hw6 = allHotWords(6)
			hw8 = allHotWords(8)
			hw10 = allHotWords(10)
			hw16 = allHotWords(16)
			hw11 = allHotWords(11)
			hw9 = allHotWords(9)
			
			
			if checkStudioSetting("tblGenOpts", "CltUseCompany") AND checkStudioSetting("tblGenOpts", "ShowCltCompanyInSch") then
				ss_CltUseCompany = true
			else
				ss_CltUseCompany = false
			end if

			function setPmtStr(cid, tid)
				set rsPmtData = Server.CreateObject("ADODB.Recordset")			
				strSQL = "SELECT Remaining, RealRemaining, Type FROM [PAYMENT DATA]"
				'Bug 3405 - fixed for shared series
				if ss_ccs_cltPackageSharing then
					strSQL = strSQL & "INNER JOIN (SELECT " & cid & " AS RelateClientID UNION SELECT ClientID1 AS RelateClientID FROM tblRelate WHERE (RelationID = - 2) AND (ClientID2 = " & cid & ") UNION SELECT ClientID2 AS RelateClientID FROM tblRelate WHERE (RelationID = - 2) AND (ClientID1 = " & cid & ")) AS RelClt ON [PAYMENT DATA].ClientID = RelClt.RelateClientID "
					strSQL = strSQL & " WHERE 1=1 "
				else 
					strSQL = strSQL & " WHERE ClientID = " & cid
				end if
				'CB 46_2409 Removed Condition [PAYMENT DATA].[Current Series]=1
				strSQL = strSQL & " AND expDate >= " & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & " "
				strSQL = strSQL & " AND (([PAYMENT DATA].TypeGroup = " & tid & ") OR  ([PAYMENT DATA].TypeGroup IN (SELECT TG2 FROM tblTGRelate WHERE TG1 = " & tid & ")))"
				strSQL = strSQL & " AND Returned = 0 "
				if ss_CheckActivationDates then
					strSQL = strSQL & " AND ([PAYMENT DATA].ActiveDate<=" & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & ") "
				end if
				'CB 46_2409 Added Condition [PAYMENT DATA].[Current Series] DESC 
				'MB #3405 added condition CASE WHEN realremaining>0 THEN 1 ELSE 2 END
				strSQL = strSQL & " ORDER BY (CASE WHEN realremaining>0 THEN 1 ELSE 2 END), [PAYMENT DATA].[Current Series] DESC, PaymentDate ASC"
				rsPmtData.CursorLocation = 3
				rsPmtData.open strSQL, cnWS
				response.write debugSQL(strSQL, "Remaining - " & cid)
				Set rsPmtData.ActiveConnection = Nothing
		
					if NOT rsPmtData.EOF then
						'if rsPmtData("Type")=9 then
						'	setPmtStr = "<span style=""color:#990000;"">" & rsPmtData("Remaining") & "</span>"
						'elseif rsPmtData("RealRemaining")=rsPmtData("Remaining") then
						'	setPmtStr = rsPmtData("remaining")
						'else
						'	setPmtStr = rsPmtData("remaining") & "&nbsp;/&nbsp;" & rsPmtData("RealRemaining")
						'end if
		
						if rsPmtData("RealRemaining")=rsPmtData("Remaining") then
							if rsPmtData("Type") = 9 then
								setPmtStr = "<span style=""color:#990000;"">" & rsPmtData("RealRemaining")*-1 & "&nbsp;owed</b></span>"
							else
								setPmtStr = rsPmtData("RealRemaining")
							end if
						else
							if rsPmtData("Type") = 9 AND rsPmtData("RealRemaining")<0  then
								setPmtStr = "<span style=""color:#990000;"" title=""" & rsPmtData("RealRemaining")-rsPmtData("Remaining") & " additional scheduled"">" & rsPmtData("RealRemaining")*-1 & "&nbsp;owed</b></span>"
							else
								setPmtStr = "<span title=""" & rsPmtData("RealRemaining")-rsPmtData("Remaining") & " additional scheduled"">" & rsPmtData("RealRemaining")
							end if
						end if
		
					else
						setPmtStr = "0"
					end if
		
				rsPmtData.close
				set rsPmtData = nothing
				exit function
			end function
			
			
			if isNum(request.form("optLocation")) then
				cLoc = CINT(request.form("optLocation"))
			elseif isNum(request.querystring("pLoc")) then
				cLoc = CINT(request.querystring("pLoc"))
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
		
			if ap_rpt_day_self AND NOT (Session("Admin") = "owner" OR Session("Admin") = "sa") AND NOT ap_rpt_day then ' can only view own schedule
				cTrn = session("empID")
			else
			    if isNum(request.form("optInstructor")) then
					cTrn = CLNG(request.form("optInstructor"))
				elseif isNum(request.querystring("pTrn")) then
					cTrn = CLNG(request.querystring("pTrn"))
				else
					cTrn = -1
				end if
			end if
		
			'if request.querystring("pView")<>"" then
			'	cView = CINT(request.querystring("pView"))
			'else
			'	cView = 3
			'end if
		
			if isNum(request.form("optTypeGroup")) then
				cTG = CINT(request.form("optTypeGroup"))
			elseif isNum(request.querystring("typeGroup")) then
				cTG = CINT(request.querystring("typeGroup"))
			else
				cTG = 0
			end if
			
			'Setup the start and end date variables (Based on code from other reports)
			if request.form("requiredtxtDateStart")<>"" then
		        Call SetLocale(session("mvarLocaleStr"))
			        cSDate = CDATE(request.form("requiredtxtDateStart"))
		        Call SetLocale("en-us")
			elseif request.querystring("pDate")<>"" then
			    Call SetLocale(session("mvarLocaleStr"))
					cSDate = CDATE(request.querystring("pDate"))
			    Call SetLocale("en-us")
	        else
		        cSDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	        end if
 
	        if request.form("requiredtxtDateEnd")<>"" then
		        Call SetLocale(session("mvarLocaleStr"))
			        cEDate = CDATE(request.form("requiredtxtDateEnd"))
		        Call SetLocale("en-us")
            elseif request.QueryString("pDate")<>"" then
                Call SetLocale(session("mvarLocaleStr"))
                    cEDate = CDATE(request.QueryString("pDate"))
                Call SetLocale("en-us")
	        else
		        cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	        end if
	        
	        ' Calculate the view type based on the date range
	        Dim dateSpan : dateSpan = DateDiff("d", cSDate, cEDate)
	        if dateSpan = 0 then
	            cView = 3
	        elseif dateSpan < 7 then
	            cView = 2
	        else
	            cView = 1
	        end if
			
			if request.form("frmGenReport")<>"" OR request.form("frmExpReport")<>"" OR request.form("requiredtxtDateStart")="" then
				genReport = true
			else
				genReport = false
			end if
			%>
		

<% if request.form("frmExpReport")<>"true" then %>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "adm/adm_rpt_dayatglance")) %>
			<script type="text/javascript">
				
				function changeLoc(strLoc) {
					var intView;
					intView = <%=cView %>
				
					if (document.frmLoc.optTypeGroup.value == "0") {
						document.frmLoc.action = strLoc + "?pTrn=" + document.frmLoc.optInstructor.value + "&pView=" + intView + "&pDate=" + document.frmLoc.requiredtxtDateStart.value<%if session("numLocations")>1 then response.write " + ""&pLoc="" + document.frmLoc.optLocation.value" end if%>;
					} else {
						document.frmLoc.action = strLoc + "?pTrn=" + document.frmLoc.optInstructor.value + "&pView=" + intView + "&typeGroup=" + document.frmLoc.optTypeGroup.value + "&pDate=" + document.frmLoc.requiredtxtDateStart.value<%if session("numLocations")>1 then response.write " + ""&pLoc="" + document.frmLoc.optLocation.value" end if%>;
					}
					document.frmLoc.frmGenReport.value = "true";
					document.frmLoc.frmExpReport.value = "false";
					document.frmLoc.submit();
				}	
				function autoSubmitWeekAdv(strLoc,nextDate) {
					var intView;
					intView = <%=cView %>
				
					if (document.frmLoc.optTypeGroup.value == "0") {
						window.location = strLoc + "?pTrn=" + document.frmLoc.optInstructor.value + 
						                           "&pView=" + intView + 
				                                   "&pDate=" + nextDate 
						                           <%if session("numLocations")>1 then response.write " + ""&pLoc="" + document.frmLoc.optLocation.value" end if%>;
					} else {
						window.location=strLoc + "?pTrn=" + document.frmLoc.optInstructor.value + 
						                         "&pView=" + intView + 
						                         "&typeGroup=" + document.frmLoc.optTypeGroup.value + 
						                         "&pDate=" + nextDate 
						                         <%if session("numLocations")>1 then response.write " + ""&pLoc="" + document.frmLoc.optLocation.value" end if%>;
					}
				}
				function exportReport() {
					document.frmLoc.frmExpReport.value = "true";
					<% iframeSubmit "frmLoc", "adm_rpt_dayatglance.asp" %>
				}
			</script>
			
<%= js(array("calendar" & dateFormatCode)) %>
					
					<!-- #include file="../inc_date_ctrl.asp" -->
					
		<% end if ''expExcel %>
		
		
		<% if request.form("frmExpReport")<>"true" then %>
		
<% pageStart %>
		<table height="100%" width="<%=strPageWidth%>" border="0" cellspacing="0" cellpadding="0">    
			<tr> 
				<td valign="top" width="100%"> <br />
					<table border="0" cellspacing="0" cellpadding="0" width="90%" height="100%" style="margin:0 auto;">
						<tr> 
							<td align="left" valign="top"> 
								<table class="mainText" border="0" cellspacing="0" cellpadding="0">
									<tr>
										<td class="headText" align="left" width="100%">
											<b>
												<%= pp_PageTitle("Schedule at a Glance") %>  
												<span class="textSmall">(Reservations &amp; <%= getHotWord(4)%>)</span></b>
												<!--JM - 49_2447-->
												<% showNewHelpContentIcon("schedule-glance-report") %>
											<br />
											<span class="textSmall"><a href="/Report/Staff/ScheduleAtAGlance">Check out the new version!</a></span>
											
										</td>
									</tr>
								</table>
							</td>
						</tr>
						<tr valign="middle">
							<td class="mainText center" align="middle" width="100%" valign="middle">&nbsp;
                                <b>&nbsp;&nbsp;
									<a class="whiteSmallText" href="javascript:autoSubmitWeekAdv('adm_rpt_dayatglance.asp','<%=FmtDateShort(DateAdd("y",-1,cSDate))%>');">
										<img border="0" src="<%= contentUrl("/asp/images/trans_arrow_grey_lt.gif") %>" width="10" height="10"/>
									</a>
									<a class="whiteSmallText" href="javascript:autoSubmitWeekAdv('adm_rpt_dayatglance.asp','<%=FmtDateShort(DateAdd("y",-1,cSDate))%>');">&nbsp;</a>
									&nbsp;Day&nbsp;
									<a class="whiteSmallText" href="javascript:autoSubmitWeekAdv('adm_rpt_dayatglance.asp','<%=FmtDateShort(DateAdd("y",1,cSDate))%>');">&nbsp;
										<img border="0" src="<%= contentUrl("/asp/images/trans_arrow_grey_rt.gif") %>" width="10" height="10"/>
									</a>&nbsp;&nbsp;
								</b> 
								<b>&nbsp;&nbsp;
									<a class="whiteSmallText" href="javascript:autoSubmitWeekAdv('adm_rpt_dayatglance.asp','<%=FmtDateShort(DateAdd("ww",-1,cSDate))%>');">
										<img border="0" src="<%= contentUrl("/asp/images/trans_arrow_grey_lt.gif") %>" width="10" height="10"/>
									</a>
									<a class="whiteSmallText" href="javascript:autoSubmitWeekAdv('adm_rpt_dayatglance.asp','<%=FmtDateShort(DateAdd("ww",-1,cSDate))%>');">&nbsp;</a>
									&nbsp;Week&nbsp;
									<a class="whiteSmallText" href="javascript:autoSubmitWeekAdv('adm_rpt_dayatglance.asp','<%=FmtDateShort(DateAdd("ww",1,cSDate))%>');">&nbsp;
										<img border="0" src="<%= contentUrl("/asp/images/trans_arrow_grey_rt.gif") %>" width="10" height="10"/>
									</a>&nbsp;&nbsp;
								</b>
							</td>
						</tr>
						<tr>
							<td valign="top" class="mainTextBig">
								<table class="mainText" width="100%" cellspacing="0" cellpadding="0" style="margin:0 auto;">
									<tr>
										<td class="mainTextBig" valign="top">
											<table class="border4 center">
											<tr>
											<td>
											<table class="mainText" border="0" cellspacing="0" cellpadding="0" style="margin:0 auto;">
													<tr valign="middle">
													<td nowrap style="background-color:#F2F2F2;">
													<b>
													<form name="frmLoc" method="post">
														<input type="hidden" name="frmExpReport" value="">
														<input type="hidden" name="frmGenReport" value="">
														
														<%=xssStr(allHotWords(57))%>:
														<select name="optfilterByCreated">
															<option value="0">Scheduled</option>
															<option value="1" <%if filterByCreated then response.write " selected" end if%>>Created</option>
														</select>

														<!--<input value="<%=FmtDateShort(curPageDate)%>" type="text" name="txtDate" size="10" maxlength="10" onBlur="validateDate(this, '<%=FmtDateShort(curPageDate)%>', true);" class="date">
														<script type="text/javascript">
															var cal1 = new tcal({'formname':'frmLoc', 'controlname':'txtDate'});
															cal1.a_tpl.yearscroll = true;
														</script> -->
														
														<%=xssStr(allHotWords(77))%>: <input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" id="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
														<script type="text/javascript">
		                                                    var cal1 = new tcal({'formname':'frmLoc', 'controlname':'requiredtxtDateStart'});
		                                                    cal1.a_tpl.yearscroll = true;
														</script>
								                        <%=xssStr(allHotWords(79))%>: <input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" id="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
														<script type="text/javascript">
		                                                    var cal2 = new tcal({'formname':'frmLoc', 'controlname':'requiredtxtDateEnd'});
		                                                    cal2.a_tpl.yearscroll = true;
		                                                </script>
                        														
														

														<% if session("numLocations")>1 then %>
															<%=xssStr(allHotWords(8))%>: 
															<select name="optLocation" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
																<option value="0" <% if cLoc=0 then response.write "selected" end if %>><b>All</b></option>
																<%
																if cLoc=0 then
																	cLocName = ""
																end if
																dim rsLoc2
																set rsLoc2 = Server.CreateObject("ADODB.Recordset")
																
																strSQL = "SELECT LocationID, LocationName FROM Location WHERE wsShow=1 ORDER BY LocationName ASC" 
																rsLoc2.CursorLocation = 3
																rsLoc2.open strSQL, cnWS
																Set rsLoc2.ActiveConnection = Nothing
																do While NOT rsLoc2.EOF
																	%>
																	<option value="<%=rsLoc2("LocationID")%>" <% if CINT(rsLoc2("LocationID"))=cLoc then response.Write "selected" end if %>><b><%=rsLoc2("LocationName")%></b></option>
																	<%
																	if CINT(rsLoc2("LocationID"))=cLoc then
																		cLocName = rsLoc2("LocationName")
																	end if	
																	rsLoc2.MoveNext
																Loop
																%>
															</select>
															<script type="text/javascript">
															    document.frmLoc.optLocation.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
															</script>
														<% else %>
															<input type="hidden" name="optLocation" value="0">
														<% end if ''''End Location dropdown %>

														<select name="optInstructor">
														<% if ap_rpt_day then ' can view all trainers %>
															<option value="-1" <%if cTrn=-1 then response.write " selected" end if%>>All</option>
														<% end if %>
														<%
														set rsTrn = Server.CreateObject("ADODB.Recordset")
														'strSQL = "SELECT TrainerID, TrFirstName, TrLastName, DisplayName "
														'strSQL = strSQL & "FROM Trainers "
														'strSQL = strSQL & "WHERE [Delete]=0 AND Active=1 AND (TRAINERS.AppointmentTrn=1 OR TRAINERS.ReservationTrn=1) "
																			
														strSQL = "SELECT DISTINCT TrainerID, TrFirstName, TrLastName, DisplayName "
														strSQL = strSQL & "FROM TRAINERS "
														strSQL = strSQL & "WHERE [Delete] = 0 AND ((TrainerID IN "
														strSQL = strSQL & "(SELECT TrainerID FROM tblReservation INNER JOIN tblTypegroup ON tblReservation.Typegroup = tblTypegroup.TypegroupID "
														strSQL = strSQL & "WHERE 1=1 "
														' filter on dates based on view
														'if CINT(cView)=1 then 'Month
														'	strSQL = strSQL & "AND (" & filterDateSQL("[tblReservation]") & ">=" & DateSep & cSDate & DateSep & " AND " & filterDateSQL("[tblReservation]") & "<=" & DateSep & cEDate & DateSep & ") "
														'elseif CINT(cView)=2 then 'Week
														'	strSQL = strSQL & "AND (" & filterDateSQL("[tblReservation]") & ">=" & DateSep & cSDate & DateSep & " AND " & filterDateSQL("[tblReservation]") & "<=" & DateSep & cEDate & DateSep & ") "
														'elseif CINT(cView)=3 then  'Day
														'	strSQL = strSQL & "AND (" & filterDateSQL("[tblReservation]") & "=" & DateSep & cSDate & DateSep & ") "
														'end if
														' Filter the dates
														strSQL = strSQL & "AND (" & filterDateSQL("[tblReservation]") & ">=" & DateSep & cSDate & DateSep & " AND " & filterDateSQL("[tblReservation]") & "<=" & DateSep & cEDate & DateSep & ") "
														' filter based on TG
														if cTG > 0 then
															strSQL = strSQL & " AND [tblReservation].TypeGroup = " & cTG & " "
														elseif cTG=-1 then	'appt only
															strSQL = strSQL & " AND tblTypeGroup.wsAppointment=1 "
														elseif cTG=-2 then
															strSQL = strSQL & " AND (tblTypeGroup.wsReservation=1 OR tblTypeGroup.wsEnrollment=1) "
														end if
														' filter on location
														if cLoc<>0 then	
															strSQL = strSQL & "AND (tblReservation.Location=" & cLoc & ") "
														end if
														strSQL = strSQL & ")) OR (TrainerID IN "
														  strSQL = strSQL & "(SELECT TrainerID FROM [VISIT DATA] INNER JOIN tblTypegroup ON [VISIT DATA].Typegroup = tblTypegroup.TypegroupID "
														strSQL = strSQL & "WHERE 1=1 "																				
														' filter on dates based on view
														'if CINT(cView)=1 then 'Month
														'	strSQL = strSQL & "AND (" & filterDateSQL("[VISIT DATA]") & ">=" & DateSep & cSDate & DateSep & " AND " & filterDateSQL("[VISIT DATA]") & "<=" & DateSep & cEDate & DateSep & ") "
														'elseif CINT(cView)=2 then 'Week
														'	strSQL = strSQL & "AND (" & filterDateSQL("[VISIT DATA]") & ">=" & DateSep & cSDate & DateSep & " AND " & filterDateSQL("[VISIT DATA]") & "<=" & DateSep & cEDate & DateSep & ") "
														'elseif CINT(cView)=3 then  'Day
														'	strSQL = strSQL & "AND (" & filterDateSQL("[VISIT DATA]") & "=" & DateSep & cSDate & DateSep & ") "
														'end if
														' Filter the dates
														strSQL = strSQL & "AND (" & filterDateSQL("[VISIT DATA]") & ">=" & DateSep & cSDate & DateSep & " AND " & filterDateSQL("[VISIT DATA]") & "<=" & DateSep & cEDate & DateSep & ") "
														
														' filter based on TG
														if cTG > 0 then
															strSQL = strSQL & " AND [VISIT DATA].TypeGroup = " & cTG & " "
														elseif cTG=-1 then	'appt only
															strSQL = strSQL & " AND tblTypeGroup.wsAppointment=1 "
														elseif cTG=-2 then
															strSQL = strSQL & " AND (tblTypeGroup.wsReservation=1 OR tblTypeGroup.wsEnrollment=1) "
														end if
														' filter on location
														if cLoc<>0 then	
															strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
														end if
														strSQL = strSQL & "))) "
																
														if not ap_rpt_day then ' can only view own schedule
															if session("empID")<>"" then
																strSQL = strSQL & " AND TrainerID = " & session("empID")
															else
																strSQL = strSQL & " AND 1 = 0 "
															end if
														end if 
														strSQL = strSQL & " ORDER BY TrLastName"
													response.write debugSQL(strSQL, "TrainerList")
														rsTrn.CursorLocation = 3
														rsTrn.open strSQL, cnWS
														Set rsTrn.ActiveConnection = Nothing
												
														do While NOT rsTrn.EOF
															trnName = Left(FmtTrnNameNew(rsTrn, true),20)
															%>
															<option value="<%=rsTrn("TrainerID")%>"<%if cTrn=CLNG(rsTrn("TrainerID")) then response.write " selected" end if%>><%=trnName%></option>
															<%
															rsTrn.MoveNext
														Loop
														rsTrn.close
														%>
														</select>
														<% if ap_rpt_day then ' can view all trainers %>
															<script type="text/javascript">
															    document.frmLoc.optInstructor.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" " + '<%=jsEscSingle(allHotWords(113))%>';
															</script>
														<% end if %>
														<select name="optTypeGroup">
															<option value="0">All Programs</option>
															<%if session("wsType")="prem" then %>
																<option value="-1" <%if cTG="-1" then response.write "selected" end if%>>Appointments Only</option>
																<option value="-2" <%if cTG="-2" then response.write "selected" end if%>>Classes/Events Only</option>
															<%end if%>
															<option value="-3" <%if cTG="-3" then response.write "selected" end if%>>Unavailability Only</option>							
															<%
															strSQL = "SELECT TypeGroupID, TypeGroup FROM tblTypeGroup WHERE (Active = 1) AND ((wsReservation = 1) OR (wsAppointment = 1) OR (wsEnrollment=1) OR (wsResource=1) ) ORDER BY TypeGroup"
															rsEntry.CursorLocation = 3
															rsEntry.open strSQL, cnWS
															Set rsEntry.ActiveConnection = Nothing
															
															do while NOT rsEntry.EOF
																%>
																	<option value="<%=rsEntry("TypeGroupID")%>" <%if cTG=rsEntry("TypeGroupID") then response.write "selected" end if%>><%=rsEntry("TypeGroup")%></option>
																<%
																rsEntry.MoveNext
															loop
															rsEntry.close
															%>
														</select>
														</b>
														  </td>
														</tr>
														<tr valign="middle">
														<td align="center" nowrap style="background-color:#F2F2F2;">
														<b>
														<%'RI 58_2842 - OrderBy%>
														Order By:
														<select name="optOrder">
															<option value="0" <% if request.form("optOrder") = "0" then response.write "selected" end if %>>Date</option>
															<option value="1" <% if request.form("optOrder") = "1" then response.write "selected" end if %>><%=xssStr(allHotWords(113))%></option>
														</select>
														&nbsp
														<%'CB 49_2655%>
														<select name="optStatus">
															<option value=""><%=xssStr(allHotWords(149))%>&nbsp;<%=xssStr(allHotWords(60))%></option>

															<%if session("wsType")<>"resv" AND cTG<>"-2" then 'show appt%><option value="<%=ivs_hw9%>" <%if request.form("optStatus")=ivs_hw9 then response.write "selected" end if%>><%=ivs_hw9%></option><%end if%>			<!-- Booked A-->
															<%if session("wsType")<>"appt" AND cTG<>"-1" then 'show class%><option value="Reserved" <%if request.form("optStatus")="Reserved" then response.write "selected" end if%>>Reserved</option><%end if%>   				<!-- Reserved C-->
															<%if session("wsType")<>"resv" AND cTG<>"-2" then 'show appt%><option value="<%=ivs_hw11%>" <%if request.form("optStatus")=ivs_hw11 then response.write "selected" end if%>><%=ivs_hw11%></option><%end if%>   		<!-- Confirmed A-->
															<%if session("wsType")<>"resv" AND cTG<>"-2" then 'show appt%><option value="<%=ivs_hw16%>" <%if request.form("optStatus")=ivs_hw16 then response.write "selected" end if%>><%=ivs_hw16%></option><%end if%>   		<!-- Arrived A-->
															<option value="Late Cancel" <%if request.form("optStatus")="Late Cancel" then response.write "selected" end if%>>Late Cancel</option>		<!-- Late Cancel B-->
															<option value="No Show" <%if request.form("optStatus")="No Show" then response.write "selected" end if%>>No Show</option>					<!-- No Show B-->
															<%if session("wsType")<>"appt" AND cTG<>"-1" then 'show class%><option value="Absent" <%if request.form("optStatus")="Absent" then response.write "selected" end if%>>Absent</option><%end if%>						<!-- Absent C-->
															<%if session("wsType")<>"appt" AND cTG<>"-1" then 'show class%><option value="Made-Up" <%if request.form("optStatus")="Made-Up" then response.write "selected" end if%>>Made-Up</option><%end if%>					<!-- Made-Up C-->
															<%if session("wsType")<>"resv" AND cTG<>"-2" then 'show appt%><option value="<%=ivs_hw10%>" <%if request.form("optStatus")=ivs_hw10 then response.write "selected" end if%>><%=ivs_hw10%></option><%end if%>   		<!-- Completed A-->
															<%if session("wsType")<>"appt" AND cTG<>"-1" then 'show class%><option value="Signed-In" <%if request.form("optStatus")="Signed-In" then response.write "selected" end if%>>Signed-In</option><%end if%>		   		<!-- Signed-In C-->
														</select>
														&nbsp;
														<% taggingFilter %>
														</b>
														</td>
														</tr>
														<tr>
											              <td nowrap style="background-color:#F2F2F2;">
											              <table cellpadding="0" class="mainText" style="margin:0 auto;">
	                                                        <tr>
		                                                        <td align="center">
                                                                    &nbsp;<img alt="Last 30" src="<%= contentUrl("/asp/adm/images/icon_last_30.png") %>" id="last30Days" title="Last 30 Days" border="0" style="cursor:pointer;" onClick="mboBack30days(frmLoc.requiredtxtDateStart, frmLoc.requiredtxtDateEnd);" />
                                                                    &nbsp;<img alt="Last 7" src="<%= contentUrl("/asp/adm/images/icon_last_7.png") %>" id="last7Days" title="Last 7 Days" border="0" style="cursor:pointer" onClick="mboBack7days(frmLoc.requiredtxtDateStart, frmLoc.requiredtxtDateEnd);" />
                                                                    &nbsp;<img alt="Next 7" src="<%= contentUrl("/asp/adm/images/icon_next_7.png") %>" id="next7Days" title="Next 7 Days" border="0" style="cursor:pointer" onClick="mboNext7days(frmLoc.requiredtxtDateStart, frmLoc.requiredtxtDateEnd);" />
                                                                    &nbsp;<img alt="Next 30" src="<%= contentUrl("/asp/adm/images/icon_next_30.png") %>" id="next30Days" title="Next 30 Days" border="0" style="cursor:pointer" onClick="mboNext30days(frmLoc.requiredtxtDateStart, frmLoc.requiredtxtDateEnd);" />
                                                                </td>
                                                                <td style="background-color:#CCCCCC;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="100%" width="1"></td>
		                                                        <td align="center">
			                                                        &nbsp;<img src="<%= contentUrl("/asp/adm/images/icon_this_d.png") %>" title="Today" border="0" style="cursor:pointer" onClick="mboToday(frmLoc.requiredtxtDateStart, frmLoc.requiredtxtDateEnd);" />
			                                                        &nbsp;<img src="<%= contentUrl("/asp/adm/images/icon_this_w.png") %>" title="This Week" border="0"  style="cursor:pointer" onClick="mboThisWeek(frmLoc.requiredtxtDateStart, frmLoc.requiredtxtDateEnd);" />
			                                                        &nbsp;<img src="<%= contentUrl("/asp/adm/images/icon_this_m.png") %>" title="This Month" border="0" style="cursor:pointer" onClick="mboThisWholeMonth(frmLoc.requiredtxtDateStart, frmLoc.requiredtxtDateEnd);" />
		                                                        </td>
		                                                      </tr>
		                                                    </table>
                                                          </td>
		                                                </tr>
														<tr valign="middle">
														  <td class="center" nowrap style="background-color:#F2F2F2;">
														  <input name="genBtn" type="button" onClick="changeLoc('adm_rpt_dayatglance.asp');" value="Generate">
														  <% exportToExcelButton %>
														  <% taggingButtons("frmLoc") %>
														  <% savingButtons "frmLoc", "Schedule at a Glance" %>
													    </td>
													  </tr>
													</form>
													</table>
													</td>
													</tr>

												</table>
										</td>
									</tr>
													<% if request.form("frmExpReport")<>"true"  then %>
														<tr><td>&nbsp;</td></tr>
													<% end if %>

									<tr valign="top" style="background-color:#FFFFFF;">
										<td valign="top" class="mainText" colspan="2" height="100%">
<% end if 'expExcel %>		  
											<%
											if request.form("frmExpReport")="true" then
												Dim stFilename
												stFilename="attachment; filename=Schedule_at_a_Glance"
												if filterByCreated="1" then
													stFilename= stFilename & "_by_Creation_date"
												end if
												if cTrn<>-1 then
													stFilename = stFileName & "_for_" & Replace(FmtTrnName(cTrn), "&nbsp;", "_") '& request.form("expTrnName")
												end if
												if CINT(cView)=3 then
													if dateFormatCode=1 then
														stFilename = stFilename & "_" & Day(cSDate) & "-" & Month(cSDate) & "-" & Year(cSDate)
													elseif dateFormatCode=2 then
														stFilename = stFilename & "_" & Month(cSDate) & "-" & Day(cSDate) & "-" & Year(cSDate)
													elseif dateFormatCode=3 then
														stFilename = stFilename & "_" & Year(cSDate) & "-" & Day(cSDate) & "-" & Month(cSDate)
													end if
												end if
												stFilename = stFilename & ".xls"
												Response.ContentType = "application/vnd.ms-excel" 
												Response.AddHeader "Content-Disposition", stFilename 
											end if
											%>
											<% if genReport then %>
												<table width="100%" class="mainText" border="0" cellspacing=0 cellpadding=0>
													<%
													numWeeks = 0
													intdays = weekday(cSDate)
													intdaysave = weekday(cSDate)
													
													set rsEntry = Server.CreateObject("ADODB.Recordset")
													
													'CB 52_2889 - New Core Query
													strSQL = "SELECT CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS C, TRAINERS.TrLastName, TRAINERS.TrFirstName, MBFBookingID, Location.LocationID, Location.LocationName, tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup AS TypeGroupName, tblTypeGroup.wsResource, tblTypeGroup.wsArrival, CLIENTS.MedicAlert, CLIENTS.StaffAlertMsg, CLIENTS.ClientID, CLIENTS.RSSID, CLIENTS.LastName, CLIENTS.FirstName, CLIENTS.CompanyName, CLIENTS.HomePhone, CLIENTS.WorkPhone, CLIENTS.CellPhone, CASE WHEN CLIENTS.FirstClassDate IS NULL OR CLIENTS.FirstClassDate=[VISIT DATA].ClassDate THEN NULL ELSE CLIENTS.FirstClassDate END AS isNewClient, CLIENTS.BirthDate, TRAINERS.TrainerID, TRAINERS.DisplayName, [VISIT DATA].VisitRefNo, [VISIT DATA].ClassDate, CN.Notes, [VISIT DATA].ClassTime, [VISIT DATA].myEndTime, [VISIT DATA].VisitType, [VISIT DATA].TypeGroup, [VISIT DATA].Webscheduler, [Visit Data].Cancelled, [VISIT DATA].EmpID, (CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, [VISIT DATA].CreationDateTime)))) AS CreationDateTime, CN.ClassName, [Visit Data].Missed, [Visit Data].ClassID, [VISIT DATA].MakeUpVisitRefNo, null AS Pending, null AS Booked, null AS Confirmed2, null AS Confirmed,  null AS TypeName, MEMBERSHIP.ActiveDate, MEMBERSHIP.ExpDate, MEMBERSHIP.TypePurch, MEMBERSHIP.IconNum, CREATEDBY.TrFirstName AS CrFirstName, CREATEDBY.TrLastName AS CrLastName, 0 AS UnpaidAppointment "
	                                                if schGlanceAcctBal then
		                                                strSQL = strSQL & ", BALANCES.Balance "
	                                                end if
													strSQL = strSQL & "FROM tblTypeGroup INNER JOIN TRAINERS INNER JOIN [VISIT DATA] INNER JOIN Location ON [VISIT DATA].Location = Location.LocationID INNER JOIN CLIENTS ON [VISIT DATA].ClientID = CLIENTS.ClientID ON TRAINERS.TrainerID = [VISIT DATA].TrainerID ON tblTypeGroup.TypeGroupID = [VISIT DATA].TypeGroup LEFT OUTER JOIN (SELECT tblClassSch.ClassDate as CNClassDate, tblClassSch.ClassID as CNClassID, tblClassSch.ClassNotes AS Notes, tblClassDescriptions.ClassName FROM tblClassSch INNER JOIN tblClassDescriptions ON tblClassSch.DescriptionID = tblClassDescriptions.ClassDescriptionID) CN ON [VISIT DATA].ClassDate = CN.CNClassDate AND [VISIT DATA].ClassID = CN.CNClassID LEFT OUTER JOIN " & memberSubSQL(0) & " MEMBERSHIP ON MEMBERSHIP.ClientID = CLIENTS.ClientID LEFT OUTER JOIN TRAINERS CREATEDBY ON [VISIT DATA].EmpID = CREATEDBY.TrainerID "
	                                                if schGlanceAcctBal then
	                                                    strSQL = strSQL & "LEFT OUTER JOIN (SELECT ClientID, SUM(Amount) AS Balance FROM tblClientAccount WHERE Amount <> 0 AND ((ClassID IS NULL) OR ClassID = 0) AND (tblClientAccount.ClientContractID IS NULL OR NOT(tblClientAccount.DepositReleaseDate IS NULL)) GROUP BY ClientID) BALANCES ON BALANCES.ClientID = CLIENTS.ClientID "
	                                                end if
                                                    if request.form("optFilterTagged")="on" then 
                                                        strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
                                                        if session("mVarUserID")<>"" then
                                                            strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
                                                        end if
                                                        strSQL = strSQL & " ) "
													end if
													strSQL = strSQL & "WHERE ([VISIT DATA].VisitType <> -2 OR [VISIT DATA].VisitType IS NULL) AND (tblTypeGroup.wsAppointment = 0) "
													
                          ' This is the SQL version of the function getVisitStatus() in inc_visit_status.asp
                          ' If changes are made to getVisitStatus they need ot be reflected here as well
                          'Classes
                          select case request.Form("optStatus")
                            case "Late Cancel"
                              strSQL = strSQL & "AND [Visit Data].Cancelled = 1 "
                            case "Signed-In"
                              strSQL = strSQL & "AND ([Visit Data].Cancelled = 0) " &_
                                                "AND (Missed <> 1) "
                            case "Reserved"
                              strSQL = strSQL & "AND ([Visit Data].Cancelled = 0) AND (Missed = 1) " &_
                                                "AND ((dateadd(mi, datepart(mi, ClassTime) + (60 * datepart(hh, ClassTime)), ClassDate) > " & DateSep & DateAdd("n", Session("tzOffset")+10,Now) & DateSep & ") OR ClassTime IS NULL) "
                            case "Made-Up"
                              strSQL = strSQL & "AND ([Visit Data].Cancelled = 0) AND (Missed = 1) " &_
                                                "AND ((dateadd(mi, datepart(mi, ClassTime) + (60 * datepart(hh, ClassTime)), ClassDate) < " & DateSep & DateAdd("n", Session("tzOffset")+10,Now) & DateSep & ") OR ClassTime IS NULL) " &_
                                                "AND (MakeUpVisitRefNo IS NOT NULL) AND (MakeUpVisitRefNo <> 0) "
                            case "Absent"
                              strSQL = strSQL & "AND ([Visit Data].Cancelled = 0) AND (Missed = 1) " &_
                                                "AND ((dateadd(mi, datepart(mi, ClassTime) + (60 * datepart(hh, ClassTime)), ClassDate) < " & DateSep & DateAdd("n", Session("tzOffset")+10,Now) & DateSep & ")  OR ClassTime IS NULL) " &_
                                                "AND ((MakeUpVisitRefNo IS NULL) OR (MakeUpVisitRefNo = 0)) "
                            case ivs_hw10, ivs_hw16, ivs_hw11, "No Show", ivs_hw9 
                              'Completed, Arrived, Confirmed, No Show, Booked
                              'This only applies to Appointments so filter out all classes
                              strSQL = strSQL & "AND (1 = 0) "
                          end select


													if cTrn<>-1 then
														strSQL = strSQL & "AND ([VISIT DATA].TrainerID = " & cTrn & ") "
													end if
													if cLoc<>0 then
														strSQL = strSQL & "AND (Location.LocationID = " & cLoc & ") "
													end if
													if cTG > 0 then
														strSQL = strSQL & "AND ([VISIT DATA].TypeGroup = " & cTG & ") "
													elseif cTG=-1 OR cTG=-3 then	'appt only or unavailability only
														strSQL = strSQL & "AND 1=0 "
													elseif cTG=-2 then  'resv only
														strSQL = strSQL & " AND (tblTypeGroup.wsReservation=1 OR tblTypeGroup.wsEnrollment=1) "
													end if
													

													' Filter by the date
													strSQL = strSQL & "AND (" & filterDateSQL("[VISIT DATA]") & ">=" & DateSep & cSDate & DateSep & " AND " & filterDateSQL("[VISIT DATA]") & "<=" & DateSep & cEDate & DateSep & ") "

                                                    strSQL = strSQL & "UNION ALL "

                                                    ' Appointment stuff starts here
                                                    strSQL = strSQL & "SELECT CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS C, TRAINERS.TrLastName, TRAINERS.TrFirstName, MBFBookingID, Location.LocationID, Location.LocationName, tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup AS TypeGroupName, tblTypeGroup.wsResource, tblTypeGroup.wsArrival, CLIENTS.MedicAlert, CLIENTS.StaffAlertMsg, CLIENTS.ClientID, CLIENTS.RSSID, CLIENTS.LastName, CLIENTS.FirstName, CLIENTS.CompanyName, CLIENTS.HomePhone, CLIENTS.WorkPhone, CLIENTS.CellPhone, CASE WHEN CLIENTS.FirstClassDate IS NULL OR CLIENTS.FirstClassDate=tblReservation.ClassDate THEN NULL ELSE CLIENTS.FirstClassDate END AS isNewClient, CLIENTS.BirthDate, TRAINERS.TrainerID, TRAINERS.DisplayName, tblReservation.VisitRefNo, tblReservation.ClassDate, tblReservation.Notes, tblReservation.ClassTime, tblReservation.myEndTime, tblReservation.VisitType, tblReservation.TypeGroup, tblReservation.Webscheduler, tblReservation.Cancelled, [tblReservation].EmpID, (CONVERT(DATETIME, FLOOR(CONVERT(FLOAT, tblReservation.CreationDateTime)))) AS CreationDateTime, null AS ClassName, null AS Missed, null AS ClassID, null AS MakeUpVisitRefNo, tblReservation.Pending, tblReservation.Booked, tblReservation.Confirmed2, tblReservation.Confirmed, tblVisitTypes.TypeName, MEMBERSHIP.ActiveDate, MEMBERSHIP.ExpDate, MEMBERSHIP.TypePurch, MEMBERSHIP.IconNum, CREATEDBY.TrFirstName AS CrFirstName, CREATEDBY.TrLastName AS CrLastName, CASE WHEN tblReservation.PmtRefNo = 0 THEN 1 ELSE 0 END AS UnpaidAppointment "
	                                                if schGlanceAcctBal then
		                                                strSQL = strSQL & ", BALANCES.Balance "
	                                                end if
                                                    strSQL = strSQL & "FROM (((CLIENTS INNER JOIN tblReservation ON CLIENTS.ClientID = tblReservation.ClientID) INNER JOIN TRAINERS ON tblReservation.TrainerID = TRAINERS.TrainerID) INNER JOIN tblTypeGroup ON tblReservation.TypeGroup = tblTypeGroup.TypeGroupID) INNER JOIN Location ON [tblReservation].Location = Location.LocationID LEFT OUTER JOIN tblVisitTypes ON tblReservation.VisitType = tblVisitTypes.TypeID LEFT OUTER JOIN " & memberSubSQL(0) & " MEMBERSHIP ON MEMBERSHIP.ClientID = CLIENTS.ClientID LEFT OUTER JOIN TRAINERS CREATEDBY ON [tblReservation].BookedBy = CREATEDBY.TrainerID "
	                                                if schGlanceAcctBal then
	                                                    strSQL = strSQL & "LEFT OUTER JOIN (SELECT ClientID, SUM(Amount) AS Balance FROM tblClientAccount WHERE Amount <> 0 AND ((ClassID IS NULL) OR ClassID = 0) AND (tblClientAccount.ClientContractID IS NULL OR NOT(tblClientAccount.DepositReleaseDate IS NULL)) GROUP BY ClientID) BALANCES ON BALANCES.ClientID = CLIENTS.ClientID "
	                                                end if
                                                    if request.form("optFilterTagged")="on" then 
                                                        strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
                                                        if session("mVarUserID")<>"" then
                                                            strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
                                                        end if
                                                        strSQL = strSQL & " ) "
													end if
                                                    strSQL = strSQL & "WHERE (tblReservation.Pending =0) "
													if cTrn<>-1 then
														strSQL = strSQL & "AND ([tblReservation].TrainerID=" & cTrn & ") "
													end if
													if cLoc<>0 then
														strSQL = strSQL & "AND (Location.LocationID = " & cLoc & ") "
													end if
													if cTG > 0 then
														strSQL = strSQL & "AND ([tblReservation].TypeGroup = " & cTG & ") "
													elseif cTG=-1 then	'appt only
														strSQL = strSQL & "AND (tblTypeGroup.wsAppointment = 1) "
													elseif cTG=-2 OR cTG=-3 then  'resv only or unavailability only
														strSQL = strSQL & " AND 1=0 "
													end if
													
                          ' This is the SQL version of the function getVisitStatus() in inc_visit_status.asp
                          ' If changes are made to getVisitStatus they need ot be reflected here as well

			                    ' Appointments
                          select case request.Form("optStatus")
                            case "Late Cancel"
                              strSQL = strSQL & "AND ([tblReservation].Cancelled = 1) "
                            case ivs_hw10 'Completed
                              strSQL = strSQL & "AND ([tblReservation].Cancelled = 0) " &_
                                                "AND (tblReservation.Booked = 1 OR tblReservation.Booked IS NULL) "
                            case ivs_hw16 ' Arrived
                              strSQL = strSQL & "AND ([tblReservation].Cancelled = 0) AND NOT (tblReservation.Booked = 1 OR tblReservation.Booked IS NULL) " &_
                                                "AND (tblReservation.Confirmed2 = 1) "
                            case "No Show"
                              strSQL = strSQL & "AND ([tblReservation].Cancelled = 0) AND NOT (tblReservation.Booked = 1 OR tblReservation.Booked IS NULL) AND (tblReservation.Confirmed2 = 0) " &_
                                                "AND (tblReservation.Booked = 0) AND (dateadd(mi, datepart(mi, ClassTime) + (60 * datepart(hh, ClassTime)), ClassDate) < " & DateSep & DateAdd("n", Session("tzOffset")-10,Now) & DateSep & ") "
                            case ivs_hw11 ' Confirmed
                              strSQL = strSQL & "AND ([tblReservation].Cancelled = 0) AND NOT (tblReservation.Booked = 1 OR tblReservation.Booked IS NULL) AND (tblReservation.Confirmed2 = 0) " &_
                                                "AND (tblReservation.Booked = 0) AND (dateadd(mi, datepart(mi, ClassTime) + (60 * datepart(hh, ClassTime)), ClassDate) > " & DateSep & DateAdd("n", Session("tzOffset")-10,Now) & DateSep & ") " &_
                                                "AND (tblReservation.Confirmed = 1) "
                            case ivs_hw9 ' Booked
                              strSQL = strSQL & "AND ([tblReservation].Cancelled = 0) AND (tblReservation.Booked IS NOT NULL) AND (tblReservation.Confirmed2 = 0) " &_
                                                "AND (dateadd(mi, datepart(mi, ClassTime) + (60 * datepart(hh, ClassTime)), ClassDate) > " & DateSep & DateAdd("n", Session("tzOffset")-10,Now) & DateSep & ") " &_
                                                "AND (tblReservation.Confirmed = 0) "
                            case "Signed-In", "Reserved", "Made-Up", "Absent"
                              ' Appointments do not have this status so filter them out
                              strSQL = strSQL & "AND (1 = 0) "
                          end select

													'Filter by Date
													strSQL = strSQL & "AND (" & filterDateSQL("[tblReservation]") & ">=" & DateSep & cSDate & DateSep & " AND " & filterDateSQL("[tblReservation]") & "<=" & DateSep & cEDate & DateSep & ") "       
													
													'RI - Dev Item 2842 : Unavailability
                          'hide unavailability if tagging clients or filtering by status
                          ' Unavailabilities do not have statuses.  They also don't relate to a client so don't select them while tagging
													if (cTG = 0 OR cTG = -3) AND filterByCreated<>"1" AND request.Form("optStatus") = "" AND request.form("frmTagClients")<>"true" then
														dim tmpDate
														dim tmpDow	
														tmpDate = cSDate
														tmpDow = WeekDayName(WeekDay(tmpDate))
														tmpDow = Left(tmpDow, 3)
														do while tmpDate <= cEDate
															strSQL = strSQL & "UNION ALL "
															strSQL = strSQL & "SELECT CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS C, TRAINERS.TrLastName, TRAINERS.TrFirstName, NULL AS MBFBookingID, NULL AS LocationID, NULL AS LocationName, NULL AS TypeGroupID, NULL AS TypeGroupName, NULL AS wsResource, NULL AS wsArrival, NULL AS MedicAlert, NULL AS StaffAlertMsg, " &_
																									"NULL AS ClientID, NULL AS RSSID, NULL AS LastName, NULL AS FirstName, NULL AS CompanyName, NULL AS HomePhone, NULL AS WorkPhone, " &_
																									"NULL AS CellPhone, NULL AS isNewClient, NULL AS Birthdate, " &_
																									"tblTrainerSchedules.TrainerID, TRAINERS.DisplayName, NULL AS VisitRefNo, " & DateSep & tmpDate & DateSep & " AS ClassDate, " &_
																									"NULL AS Notes, tblTrainerSchedules.StartTime AS ClassTime, tblTrainerSchedules.EndTime AS myEndTime, NULL AS VisitType, NULL AS TypeGroup, " &_
																									"NULL AS Webscheduler, NULL AS Cancelled, NULL AS EmpID, NULL AS CreationDateTime, tblTrainerSchedules.AvailName AS ClassName, NULL AS Missed, NULL AS ClassID, NULL AS MakeUpVisitRefNo, " &_
																									"NULL AS Pending, NULL AS Booked, NULL AS Confirmed2, NULL AS Confirmed, NULL AS TypeName, " &_
																									"NULL AS ActiveDate, NULL AS ExpDate, NULL AS TypePurch, NULL AS IconNum, NULL AS CrFirstName, " &_
																									"NULL AS CrLastName, 0 AS UnpaidAppointment "
	                                                        if schGlanceAcctBal then
		                                                        strSQL = strSQL & ", NULL AS Balance "
	                                                        end if
															strSQL = strSQL & "FROM tblTrainerSchedules INNER JOIN TRAINERS ON TRAINERS.TrainerID = tblTrainerSchedules.TrainerID " &_
																								"WHERE ClassID IS NULL AND Unavailable = 1 AND StartDate <= " & DateSep & tmpDate & DateSep & " AND EndDate >= " & DateSep & tmpDate & DateSep & " AND tblTrainerSchedules.Day" & tmpDow & " = 1 "
															if cTrn<>-1 then
																strSQL = strSQL & "AND tblTrainerSchedules.TrainerID=" & cTrn & " "
															end if
															tmpDate = DateAdd("d", 1, tmpDate)
															tmpDow = WeekDayName(WeekDay(tmpDate))
															tmpDow = Left(tmpDow, 3)
														loop
													end if

													if request.form("frmTagClients")<>"true" then 
														strSQL = strSQL & " ORDER BY "
														if request.Form("optOrder") = "1" then
															strSQL = strSQL & GetTrnOrderBy() & ", "
														end if
														if request.Form("optfilterByCreated") = "1" then
															strSQL = strSQL & "CreationDateTime, "
														end if
														strSQL = strSQL & "[VISIT DATA].ClassDate, [VISIT DATA].ClassTime, CN.ClassName, [VISIT DATA].ClassID, TRAINERS.TrainerID, [VISIT DATA].Cancelled DESC, CLIENTS.LastName, CLIENTS.FirstName;"
													end if
													'response.Write strSQL
													'response.End
													debugSQL strSQL, "Main"
													'response.flush
	
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

													curTrnID = 0
													if not rsEntry.EOF then
														curDate = rsEntry("ClassDate")
														curTime = rsEntry("ClassTime")
													end if
													
													Dim viewNumWeeks, totalNumVisits, NumVisits, NumCount
													NumCount = 0
													NumVisits = 0
													totalNumVisits = 0
													viewNumWeeks = DateDiff("w", cSDate, cEDate) + 1
													'response.Write("viewNumWeeks: " & viewNumWeeks)
													
													do while numWeeks <= viewNumWeeks
														intCount = 0
														rowcount=0 
														first = true
														
														if CINT(cView)<>3 or filterByCreated="1" then
															intLoopControl = 7
														else
															intLoopControl = intdaysave
														end if
														
														do while intDays <= intLoopControl
															if intdays>=1 and intdays<=7 then
																strDays = WeekdayName(intdays)
															end if
															
															if NOT rsEntry.EOF then
															    if nextCreationDate then 
																    'set thisDate to earliest of two ClassDates
																    cSDate = rsEntry("ClassDate")
															    end if
																if isNull(rsEntry("ClassDate")) then
																	cont=true
																elseif CDATE(rsEntry("ClassDate"))=CDATE(cSDate) then
																	cont=true
																else
																	cont=false
																end if
															end if
																													
															if cont then intCount = intCount + 1
															
															Do While cont 
																
																if isNULL(rsEntry("Pending")) then
															    displaying = "res"
																else
																	displaying = "appt"
																end if
																
																tmpIsMakeUp = false
																if NOT isNULL(rsEntry("MakeUpVisitRefNo")) then
																	if CLNG(rsEntry("MakeUpVisitRefNo"))<>0 then
																		tmpIsMakeUp = true
																	end if
																end if
																disClassDate = rsEntry("ClassDate")
																if isNull(rsEntry("ClassTime")) OR isNull(rsEntry("myEndTime")) then
																	disStartTime = "null"
																	disEndtime = ""
																	clsStartTime = ""
																else
																	disStartTime = TimeValue(rsEntry("ClassTime"))
																	disEndTime = TimeValue(rsEntry("myEndTime"))
																	clsStartTime = TimeValue(disStartTime)
																end if

																'BJD 4/15/08 - Replaced status coloring with inc function
																if displaying="res" AND NOT isNull(rsEntry("ClientID")) then
																	getVisitStatus false, CDATE(disClassDate & " " & clsStartTime), false, false, false, rsEntry("Cancelled"), rsEntry("Missed"), tmpIsMakeUp
																else
																	getVisitStatus true, CDATE(disClassDate & " " & clsStartTime), rsEntry("Booked"), rsEntry("Confirmed"), rsEntry("Confirmed2"), rsEntry("Cancelled"), false, false
																end if
																disStatus = visitStatus
																disColor = visitColor
																
                                'This shows incorrectly selected items
																'if request.form("optStatus")<>"" and request.form("optStatus")<>disStatus then 
                                if false then
																	intCount = intCount - 1
                                  %>
                                  <script type="text/javascript">
                                  try {
                                  <% if displaying="res" and not isnull(rsentry("ClientID")) then %> 
                                    console.log("getVisitStatus(false, <%=CDATE(disClassDate & " " & clsStartTime) %>, false, false, <%=rsEntry("Cancelled")%>, <%=rsEntry("Missed")%>, <%=tmpIsMakeUp %>)");
                                  <% else %>
                                    console.log("getVisitStatus(true, <%= CDATE(disClassDate & " " & clsStartTime)%>, <%=rsEntry("Booked")%>, <%=rsEntry("Confirmed")%>, <%=rsEntry("Confirmed2")%>, <%=rsEntry("Cancelled")%>, false, false)");
                                  <%end if %>
                                  }
                                  catch (ex) {}
                                  </script>
                                  <%
																end if

																disRSSID = rsEntry("RSSID")
																disLocationName = rsEntry("LocationName")
																disTrainerID = rsEntry("TrainerID")
																disClassID = rsEntry("ClassID")
																disClientID = rsEntry("ClientID")
																disPhone = rsEntry("HomePhone")
																disCellPhone = rsEntry("CellPhone")
																disWorkPhone = rsEntry("WorkPhone")
																DisTrainerName = FmtTrnNameNew(rsEntry, false)
																if ss_CltUseCompany AND NOT isNULL(rsEntry("CompanyName")) then
																	disClientName = rsEntry("CompanyName")
																else
																	disClientName = rsEntry("FirstName") & "&nbsp;" & rsEntry("LastName")
																end if
																disWS = rsEntry("WebScheduler")
																	
																disTG = rsEntry("TypeGroup")
																disAlert = rsEntry("MedicAlert")
																disStaffAlert = rsEntry("StaffAlertMsg")
																disNotes = rsEntry("Notes")
																	
																disIsNewClient = rsEntry("IsNewClient")
																disBirthdate = rsEntry("Birthdate")

																	if displaying = "res" then
																		if isNull(rsEntry("ClassName")) then
																			if rsEntry("wsResource") then
																				disClassName = LEFT(stripHTML(rsEntry("TypeGroupName")), 22)
																				disClassID = 0
																			else
																				'if NOT isNull(rsEntry("VisitType")) then
																				disClassName = LEFT(stripHTML(rsEntry("VisitType")), 22)
																				disClassID = 0
																			end if
																		else
																			disClassName = LEFT(stripHTML(rsEntry("ClassName")), 22)
																			disClassID = rsEntry("ClassID")
																		end if
																else
																	disClassName = rsEntry("TypeGroupName")
																end if

																if NOT isNULL(rsEntry("TypeName")) then
																	disVT = rsEntry("TypeName")
																else
																	disVT = "None"
																end if
																	
																MemActDate = rsEntry("ActiveDate")
																MemExpDate = rsEntry("ExpDate")
																MemService = rsEntry("TypePurch")
																MemIconNum = rsEntry("IconNum")
																disEmpID = rsEntry("EmpID")
																disCrFirst = rsEntry("CrFirstName")
																disCrLast = rsEntry("CrLastName")
																mbfId = rsEntry("MBFBookingID")
																' Client has unpaid appointments
																disUnpaidAppointments = CINT(rsEntry("UnpaidAppointment")) > 0
																if schGlanceAcctBal then
    																disBalance = rsEntry("Balance")
    															end if
																	
																	
																if filterByCreated="1" then
																	disCreationDate = FmtDateShort(rsEntry("CreationDateTime"))
																	intdays = weekday(rsEntry("ClassDate"))
																	strDays = weekdayname(intDays)
																end if


															'if request.form("optStatus")="" OR request.form("optStatus")=disStatus then 'CB 49_2655 - optStatus
																if first then
																	if request.form("frmExpReport")="true" then
																		%>
																		<tr class="mainText">
																			<%
																			if cTrn<>-1 then 
																				response.write "<td>" & Replace(FmtTrnName(cTrn), "&nbsp;", " ") & "</td>"
																			end if
																			if CINT(cView)=3 then
																				if dateFormatCode=1 then
																					response.write "<td>" & Day(cSDate) & "/" & Month(cSDate) & "/" & Year(cSDate) & "</td>"
																				elseif dateFormatCode=2 then
																					response.write "<td>" & Month(cSDate) & "/" & Day(cSDate) & "/" & Year(cSDate) & "</td>" 
																				elseif dateFormatCode=3 then
																					response.write "<td>" & Year(cSDate) & "/" & Day(cSDate) & "/" & Month(cSDate) & "</td>"
																				end if
																			end if
																			%>
																		</tr>
																	<% end if 	'if request.form("frmExpReport")="true" then %>
																	<% if filterByCreated="1" and nextCreationDate=true then %>
																		<tr><td colspan=12 nowrap><Strong>Created Date:&nbsp;<%=disCreationDate%></strong></td></tr>
																	<% end if %>
                                                                      <% if filterByCreated<>"1" OR nextCreationDate=true then %>
																	<% if filterByCreated="1" and nextCreationDate=true then %>
																		<% nextCreationDate = false %>
																	<% end if %>
																	<tr style="background-color:<%=session("pageColor4")%>;"> 
																		<td class="whiteSmallText"><b>&nbsp;</b></td>
																		<% if cView=1 or cView=2 or filterByCreated="1" then %>
																			<td class="whiteSmallText"><b>&nbsp;Day </b></td>
																			<td class="whiteSmallText" nowrap><b>&nbsp;<%= getHotWord(57)%></b></td>
																		<% end if %>
																		<td class="whiteSmallText" nowrap><b>&nbsp;<%= getHotWord(58)%></b></td>
																		<td class="whiteSmallText" nowrap><b>&nbsp;<%= getHotWord(65)%></b></td>
																		<% if cTrn=-1 then %>
																			<td class="whiteSmallText" nowrap><b><%=hw6%></b></td>
																		<% end if %>
																		<% if session("numLocations")>1 and cLoc=0 then %>
																			<td class="whiteSmallText" nowrap><b>&nbsp;<%=hw8%></b></td>
																		<% end if %>
																		<% if request.form("frmExpReport")<>"true" then %>
																			<td>&nbsp;</td>
																		<% end if %>
																		<td class="whiteSmallText" nowrap><b><%=session("ClientHW")%></b></td>
																		<% if request.form("frmExpReport")="true" then %>
																			<td class="whiteSmallText" nowrap><b>&nbsp;<%= getHotWord(134)%></b></td>
																		<% end if %>
																		<% if schGlanceAcctBal then %>
																			<td class="whiteSmallText" align="center" nowrap><b><%= getHotWord(62)%></b></td>
																		<% end if %>
																		<% if schGlanceRem then %>
																			<td class="whiteSmallText" align="center" nowrap><b>Remaining</b></td>
																		<% end if %>
																		<td class="whiteSmallText" nowrap><%if schGlanceShowCltPhone then%><b><%= getHotWord(93)%></b><%end if%></td>
																		<td class="whiteSmallText" align="center" nowrap><b><%= getHotWord(54)%>?</b></td>
																		<td class="whiteSmallText" nowrap><b><%= getHotWord(60)%></b></td>
																		<% if schGlanceCreatedBy then %>
																			<td class="whiteSmallText" align="center" nowrap><b>Created By</b></td>
																		<% end if %>
																		<% if request.form("frmExpReport")="true" then %>
																			<td class="whiteSmallText" nowrap><b>Red Alert</b></td>
																			<td class="whiteSmallText" nowrap><b>Yellow Alert</b></td>
                                                                              <td class="whiteSmallText" nowrap><b><%= getHotWord(137)%>&nbsp;<%= getHotWord(90)%></b></td>
																		<% end if %>
																	</tr>
                                                                      <% end if %>
																	<%			
																	first=false
																	displayed=true
																end if ''end if first
																
															'if request.form("optStatus")="" OR request.form("optStatus")=disStatus then 'CB 49_2655 - optStatus 
																numCount = numCount + 1
																if NOT isNull(disClientID) then 
																	totalNumVisits = totalNumVisits + 1
																	NumVisits = NumVisits + 1
																end if
			
																if CLNG(curTrnID)<>CLNG(disTrainerID) OR curDate<>disClassDate OR curTime<>disStartTime OR curClassId<>disClassID OR curTG<>disTG OR curVT<>disVT then
																	%>
																	<% if request.form("frmExpReport")<>"true" then %>
																		<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="15"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"/></td></tr>
																	<% end if %>
																	<%
																	if rowcount=0 then
																		trowColor = "#F2F2F2"
																		rowcount = 1
																	else
																		trowColor = "#FAFAFA"
																		rowcount = 0
																	end if
																	%>
																	<tr style="background-color:<%=trowColor%>;"> 
																		<td align="center" nowrap>&nbsp;<%=numCount%>.&nbsp;</td>
																		<% if cView=1 or cView=2 or filterByCreated="1" then %>
																			<td class="dayCell" width=55 align=left nowrap>&nbsp;<%=WeekDayName(WeekDay(disClassDate))%></td>
																			<td class="dateCell" align=left nowrap>&nbsp;<%=FmtDateShort(disClassDate)%>&nbsp;</td>
																		<% end if %>
																		<td class="timeCell" align=left nowrap>&nbsp;
																			<% if displaying = "res" AND NOT isNull(disClientID) then %>
																				<%if disStartTime="null" then response.Write "TBD" else response.write FmtTimeShort(disStartTime) end if%>
																			<% else %>
																				<%if disStartTime="null" then response.Write "TBD" else response.Write FmtTimeShorter(disStartTime) & "-&nbsp;" & FmtTimeShort(disEndTime) end if%>
																			<% end if %>
																		</td>
																		<td nowrap>&nbsp;
																			<%
																			if displaying = "appt" then
																				if request.form("frmExpReport")<>"true" then
																					response.write "<a class=""mainText apptLink"" title=""Click to view Appt. Schedule"" href=""/ws.asp?studio=" & session("studioShort") & "&stype=3&sview=day&sTG=" & disTG & "&sdate=" & disClassDate & "&strn=" & disTrainerID & """ target=""_parent"">" & disClassName & "&nbsp;/&nbsp;" & disVT & "</a>"
																				else
																					  response.write disClassName & "&nbsp;/&nbsp;" & disVT
																				end if
																			else
																				if disClassID <> 0 then
																					if request.form("frmExpReport")<>"true" then
																						response.write "<a class=""mainText classLink"" title=""Click to view Class List"" href=""adm_cls_list.asp?pDate=" & cSDate & "&pClsID=" & disClassID & """>" & disClassName & "</a>"
																					else
																						  response.Write disClassName
																					end if
																				else
																					if isNull(disClientID) then
																						response.write "<i>" & disClassName & "</i>"
																					else 
																						response.Write disClassName
																					end if
																				end if
																			end if
																			%>
																		</td>
																		<% if cTrn=-1 then %>
																			<td class="teacherCell" colspan=1 nowrap><%=disTrainerName%></td>
																		<% end if %>
																		<% if session("numLocations")>1 AND cLoc=0 then %>
																			<td colspan=1 align=left nowrap>&nbsp;<%=disLocationName%></td>
																		<% end if %>
																		<td align="right" nowrap>
<%if request.form("frmExpReport")<>"true" then %>
																			<% if not isNull(mbfId) then %><img src="<%= contentUrl("/asp/adm/images/pin_orange_micro.png") %>" border="0" title="MINDBODY Finder Sale" /> <% end if %>
																			<% if not isNULL(MemActDate) then %><img src="images/mem-<%=MemIconNum%>.png" height="16" width="16" title="Member with <%=MemService%> from <%=FmtDateShort(MemActDate)%> to <%=FmtDateShort(MemExpDate)%>" align="absbottom"/><% end if %>
																			<% if not isNULL(disStaffAlert) then %><img src="<%= contentUrl("/asp/adm/images/alert-red-16px.png") %>" height="18" width="16" title="Staff Alert: <%=Replace(disStaffAlert, """", "''")%>" align="absbottom"/>&nbsp;<% end if %>
																			<% if not isNULL(disAlert) then %><img src="<%= contentUrl("/asp/adm/images/alert-yellow-16px.png") %>" height="18" width="16" title="Alert: <%=Replace(disAlert, """", "''")%>" align="absbottom"/>&nbsp;<% end if %>
																			<% if not isNULL(disNotes) AND TRIM(disNotes)<>"" then%><img src="<%= contentUrl("/asp/adm/images/notes_icon_trans.gif") %>" border="0"  title="Appointment Notes: <%=Replace(disNotes, """", "''")%>" width="13" align="absmiddle"/>&nbsp;<%end if%>
																			<% if disUnpaidAppointments then %><img src="<%= contentUrl("/asp/adm/images/unpaid-10px.png") %>" height="16" width="16" border="0" align="absbottom" title="Unpaid Appointment"/><% end if %>
																			<% if IsNULL(disIsNewClient) AND NOT isNull(disClientID) then %><img src="<%= contentUrl("/asp/adm/images/green-star-16px.png") %>" height="16" width="16" border="0" align="absbottom" title="First Visit!"/><% end if %>
																			<% 
																			if NOT isNULL(disBirthdate) then
																				if isDate(disBirthdate) then
																					if Month(disBirthdate)=2 AND Day(disBirthdate)=29 then
																						''born on leap year too bad!
																						if DateAdd("y", -5, curDate)<=CDATE(Month(disBirthdate) & "/" & Day(disBirthdate)-1 & "/" & Year(curDate)) AND DateAdd("y", 5, curDate)>=CDATE(Month(disBirthdate) & "/" & Day(disBirthdate)-1 & "/" & Year(curDate)) then
																							'response.write "</span> (Birthday!&nbsp;" & FmtDateShorter(disBirthdate) & ")"
																							response.write "<img src=""" & contentUrl("/asp/adm/images/birthday-present-16px.png") & """ width=""16"" height=""16"" title=""Birthday! " & FmtDateShorter(disBirthdate) & """  align=""absbottom"">"
																						end if
																					else										
																						if DateAdd("y", -5, curDate)<=CDATE(Month(disBirthdate) & "/" & Day(disBirthdate) & "/" & Year(curDate)) AND DateAdd("y", 5, curDate)>=CDATE(Month(disBirthdate) & "/" & Day(disBirthdate) & "/" & Year(curDate)) then
																							'response.write "</span> (Birthday!&nbsp;" & FmtDateShorter(disBirthdate) & ")"
																							response.write "<img src=""" & contentUrl("/asp/adm/images/birthday-present-16px.png") & """ width=""16"" height=""16"" title=""Birthday! " & FmtDateShorter(disBirthdate) & """  align=""absbottom"">"
																						end if
																					end if
																				end if
																			end if %>
																		</td>
																		<td nowrap>
																			<% if NOT isNull(disClientID) then %>
																			<a class="clientLink" title="Click to view <%=session("ClientHW")%> Info" href="main_info.asp?ID=<%=disClientID%>&fl=true">
<% end if %>
																				<%=disClientName%>
<% if request.form("frmExpReport")<>"true" then %>
																			</a>
<% end if %>
																			<% end if 'export to excel %>
																			<%'MB bug#5985 %>
																		<% if request.form("frmExpReport")="true" then %>
																			<%=disClientName%>
																		<% end if %>	
																		</td>
																		<% if request.form("frmExpReport")="true" then %>
																			<td nowrap><%=disRSSID%></td>
																		<% end if %>
																		<% if schGlanceAcctBal then %>
																			<% if request.form("frmExpReport")<>"true" then %>
																				<td align="center" nowrap>&nbsp;<%=FmtCurrency(disBalance)%></td>
																			<% else %>
																				<td align="center" nowrap>&nbsp;<%=FmtNumber(disBalance)%></td>
																			<% end if %>
																		<% end if %>
																		<% if schGlanceRem then %>
																			<td align="center" nowrap>
																				<% if NOT isNull(disClientID) AND NOT isNull(disTG) then %>
																					<%=setPmtStr(disClientID, disTG)%>
																				<% end if %>
																			</td>
																		<% end if %>
																		<td nowrap>
																			<%
																			if schGlanceShowCltPhone then 
																				if NOT isNULL(disCellPhone) then
																					if request.form("frmExpReport")<>"true" then
																						response.write "<img src=""" & contentUrl("/asp/adm/images/smart-phone-16px.png") & """ align=""absbottom"" title=""Mobile Phone""> "
																					end if
																					response.write FmtPhoneNum(disCellPhone)
																				elseif NOT isNULL(disPhone) then
																					response.write FmtPhoneNum(disPhone)
																				else 
																					response.Write FmtPhoneNum(disWorkPhone)
																				end if
																			end if
																			%>
																		</td>
																		<td align="center" nowrap>
																			<% if NOT isNull(disWS) then %>
																				<% if request.form("frmExpReport")<>"true" then %>
																					<input type="checkbox" name="optWS" <%if CBool(disWS) then response.write "checked" end if%> disabled>
																				<% else %>
																					<%if CBool(disWS) then response.write "Yes" else response.write "No" end if%>
																				<% end if %>
																			<% end if %>
																		</td>
																		<td nowrap>
																			<% if NOT isNull(disClientID) then %>
																				<span style="color:<%=disColor%>;"><%=disStatus%></span>
																			<% end if %>
																		</td>
																		<% if schGlanceCreatedBy then %>
																			<td align="center" nowrap>
																				<% if isNull(disEmpID) AND NOT isNull(disClientID) then %>
																					<%=session("ClientHW")%>
																				<% elseif disEmpID = "0" then %>
																					Owner
																				<% else %>
																					<%=disCrFirst%>&nbsp;<%=disCrLast%>
																				<% end if %>
																			</td>
																		<% end if %>
																		<% if request.form("frmExpReport")="true" then %>
																			<td><%=disStaffAlert%></td>
																			<td><%=disAlert%></td>
                                                                              <td><%=disNotes%></td>
																		<% end if %>
																	</tr>
																	<% if request.form("frmExpReport")<>"true" then %>
																		<tr style="background-color:#CCCCCC;"><td colspan=15><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
																	<% end if %>
																	<%
																else	''''another client in this slot
																	Dim colCount
																	colCount = 2
																	if cView<3 or filterByCreated="1" then
																		colCount = colCount + 2
																	end if
																	if session("numLocations")>1 and cLoc=0 then
																		colCount = colCount + 1
																	end if
																	if cTrn=-1 then
																		colCount = colCount + 1
																	end if
		
																	if rowcount=0 then
																		trowColor = "#F2F2F2"
																		rowcount = 1
																	else
																		trowColor = "#FAFAFA"
																		rowcount = 0
																	end if
																	%>
																	<tr style="background-color:<%=trowColor%>;">
																		<td align="center" nowrap>&nbsp;<%=numCount%>.&nbsp;</td>				
																		<% if request.form("frmExpReport")<>"true" then %>
																			<td colspan="<%=colCount%>" align=left nowrap>&nbsp;</td>
																		<% else 'EXCEL %>
																			<% if cView=1 or cView=2 or filterByCreated="1" then %>
																				<td width=55 align=left nowrap>&nbsp;<%=strDays%></td>
																				<td align=left nowrap>&nbsp;<%=FmtDateShort(cSDate)%>&nbsp;</td>
																			<% end if %>
																			<td align=left nowrap>&nbsp;
																			<% if displaying = "res" then %>
																				<%if disStartTime="null" then response.Write "TBD" else response.write FmtTimeShort(disStartTime) end if%>
																			<% else %>
																				<%if disStartTime="null" then response.Write "TBD" else response.Write FmtTimeShorter(disStartTime) & "-&nbsp;" & FmtTimeShort(disEndTime) end if%>
																			<% end if %>
																			</td>
																			<td nowrap>&nbsp;
																				<%
																				if displaying = "appt" then
																					response.write disClassName & "&nbsp;/&nbsp;" & disVT
																				else
																					response.write disClassName
																				end if
																				%>
																			</td>
																			<% if cTrn=-1 then %>
																				<td colspan=1 nowrap><%=disTrainerName%></td>
																			<% end if %>
																			<% if session("numLocations")>1 AND cLoc=0 then %>
																				<td colspan=1 align=left nowrap>&nbsp;<%=disLocationName%></td>
																			<% end if %>
																		<% end if '''EXCEL%>
																		<td align="right" nowrap>
<%if request.form("frmExpReport")<>"true" then %>
																			<% if not isNull(mbfId) then %><img src="<%= contentUrl("/asp/adm/images/pin_orange_micro.png") %>"  height="14" width="9" border="0" title="MINDBODY Finder Sale" /> <% end if %>
																			<% if not isNULL(MemActDate) then %><img src="images/mem-<%=MemIconNum%>.png"  height="16" width="16" title="Member with <%=MemService%> from <%=FmtDateShort(MemActDate)%> to <%=FmtDateShort(MemExpDate)%>" align="absbottom">
																			<% end if %>
																			<% if not isNULL(disStaffAlert) then %><img src="<%= contentUrl("/asp/adm/images/alert-red-16px.png") %>"  height="18" width="16" title="Staff Alert: <%=Replace(disStaffAlert, """", "''")%>" align="absbottom">&nbsp;<% end if %>
																			<% if not isNULL(disAlert) then %><img src="<%= contentUrl("/asp/adm/images/alert-yellow-16px.png") %>"  height="18" width="16" title="Alert: <%=Replace(disAlert, """", "''")%>" align="absbottom">&nbsp;<% end if %>
																			<% if not isNULL(disNotes) AND TRIM(disNotes)<>"" then%><img src="<%= contentUrl("/asp/adm/images/notes_icon_trans.gif") %>" border="0"  title="Appointment Notes: <%=Replace(disNotes, """", "''")%>" width="13" align="absmiddle">&nbsp;<%end if%>
																			<% if IsNULL(disIsNewClient) then %>
																				<img src="<%= contentUrl("/asp/adm/images/green-star-16px.png") %>"  height="16" width="16" border="0" align="absbottom" title="First Visit!">
																			<% end if %>
																			<% if NOT isNULL(disBirthdate) then
																				if isDate(disBirthdate) then
																					if Month(disBirthdate)=2 AND Day(disBirthdate)=29 then
																						''born on leap year too bad!
																						if DateAdd("y", -5, curDate)<=CDATE(Month(disBirthdate) & "/" & Day(disBirthdate)-1 & "/" & Year(curDate)) AND DateAdd("y", 5, curDate)>=CDATE(Month(disBirthdate) & "/" & Day(disBirthdate)-1 & "/" & Year(curDate)) then
																							'response.write "</span> (Birthday!&nbsp;" & FmtDateShorter(disBirthdate) & ")"
																							response.write "<img src=""" & contentUrl("/asp/adm/images/birthday-present-16px.png") & """ width=""16"" height=""16"" title=""Birthday! " & FmtDateShorter(disBirthdate) & """  align=""absbottom"">"
																						end if
																					else										
																						if DateAdd("y", -5, curDate)<=CDATE(Month(disBirthdate) & "/" & Day(disBirthdate) & "/" & Year(curDate)) AND DateAdd("y", 5, curDate)>=CDATE(Month(disBirthdate) & "/" & Day(disBirthdate) & "/" & Year(curDate)) then
																							'response.write "</span> (Birthday!&nbsp;" & FmtDateShorter(disBirthdate) & ")"
																							response.write "<img src=""" & contentUrl("/asp/adm/images/birthday-present-16px.png") & """ width=""16"" height=""16"" title=""Birthday! " & FmtDateShorter(disBirthdate) & """  align=""absbottom"">"
																						end if
																					end if
																				end if
																			end if %>
																		</td>
																		<td nowrap>
																			<a title="Click to view <%=session("ClientHW")%> Info" href="main_info.asp?ID=<%=disClientID%>&fl=true">
<% end if 'export to excel %>
																				<%=disClientName%>
<% if request.form("frmExpReport")<>"true" then %>
																			</a>
<% end if %>																				
																		</td>
																		<% if request.form("frmExpReport")="true" then %>
																			<td nowrap><%=disRSSID%></td>
																		<% end if %>
																		<% if schGlanceAcctBal then %>
																			<% if request.form("frmExpReport")<>"true" then %>
																				<td align="center" nowrap>&nbsp;<%=FmtCurrency(disBalance)%></td>
																			<% else %>
																				<td align="center" nowrap>&nbsp;<%=FmtNumber(disBalance)%></td>
																			<% end if %>
																		<% end if %>
																		<% if schGlanceRem then %>
																			<td align="center" nowrap>
																				<% if NOT isNull(disClientID) AND NOT isNull(disTG) then %>
																				<%=setPmtStr(disClientID, disTG)%>
																				<% end if %>
																			</td>
																		<% end if %>
																		<td nowrap>
																			<%
																			if schGlanceShowCltPhone then
																				if NOT isNULL(disCellPhone) then
																						response.write "<img src=""" & contentUrl("/asp/adm/images/smart-phone-16px.png") & """ align=""absbottom"" title=""Mobile Phone""> " & FmtPhoneNum(disCellPhone)
																				elseif NOT isNULL(disPhone) then
																					response.write FmtPhoneNum(disPhone)
																				else 
																				  response.Write FmtPhoneNum(disWorkPhone)
																				end if
																			end if
																			%>	
																		</td>
																		<td align="center" nowrap>
																			<% if NOT isNull(disWS) then %>
																				<% if request.form("frmExpReport")<>"true" then %>
																					<input type="checkbox" name="optWS" <%if CBool(disWS) then response.write "checked" end if%> disabled>
																				<% else %>
																					<%if CBool(disWS) then response.write "Yes" else response.write "No" end if%>
																				<% end if %>
																			<% end if %>
																		</td>
																		<td nowrap><span style="color:<%=disColor%>;"><%=disStatus%></span></td>
																		<% if schGlanceCreatedBy then %>
																			<td align="center" nowrap>
																				<% if isNull(disEmpID) then %>
																					<%=session("ClientHW")%>
																				<% elseif disEmpID = "0" then %>
																					Owner
																				<% else %>
																					<%=disCrFirst%>&nbsp;<%=disCrLast%>
																				<% end if %>
																			</td>
																		<% end if %>
																		<% if request.form("frmExpReport")="true" then %>
																			<td><%=disStaffAlert%></td>
																			<td><%=disAlert%></td>
                                                                              <td><%=disNotes%></td>
																		<% end if %>
																	</tr>
																	<% if request.form("frmExpReport")<>"true" then %>
																		<tr style="background-color:#CCCCCC;"><td colspan=15><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
																	<% end if %>
																	<%	
																end if			'''first one or another client in slot
																

																''''''''Move next and setup vars for next loop iteration
																curTrnID = disTrainerID
																curDate = disClassDate
																curTime = disStartTime
																curClassID = disClassID
																curTG = disTG
																curVT = disVT
																curCreationDate = disCreationDate 
																curStatus = disStatus 
																			
																rsEntry.MoveNext	    
																displaying = "" 
															
																if NOT rsEntry.EOF then
																	if CDATE(rsEntry("ClassDate"))=CDATE(cSDate) OR request.Form("optOrder")="1" then
																		cont=true
																	else
																		cont=false
																	end if		
																else
																	cont=false
																end if
																
																if filterByCreated="1" then 
                                                                    if NOT rsEntry.EOF then
																		if FmtDateShort(rsEntry("CreationDateTime"))<>curCreationDate then
																			nextCreationDate = true
																		else '' rsEntry still on curCreationDate
																			nextClassDate = rsEntry("ClassDate")
																		end if
																	end if
																end if
																if nextCreationDate then 
																	intDays = 8
																	cont = false
																end if
																
															Loop 'Do While cont
															if filterByCreated<>"1" or displaying="" then 
																nextClassDate = DateAdd("y",1,cSDate)
															end if
															if not nextCreationDate or (rsEntry.EOF) then
																intDays = intDays + DateDiff("y",cSDate,nextClassDate)
															end if
															cSDate = nextClassDate

														Loop 'Do While intDays <= intLoopControl

														if (request.form("optStatus")="" or request.form("optStatus")=curStatus or not rsEntry.EOF) and displayed=true then
															if intCount = 0 then  
																%>
																<tr> 
																	<td colspan=11>&nbsp;No Reservations</td>
																</tr>
																<%
															end if
															%>
															<% if filterByCreated<>"1" then %>
															<% if request.form("frmExpReport")<>"true" then %>
																<tr style="background-color:<%=session("pageColor4")%>;"><td colspan=15><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
																<tr> 
																	<td colspan=15 align="right">&nbsp;Number of Visits: <b><%=NumVisits%></b></td>
																</tr>
															<% end if %>
															<% end if %>
															<%
														end if
														if nextCreationDate then 
															Response.Write "<tr><td colspan=15>&nbsp;</td></tr>"
														end if
														
														intDays = 1
															numWeeks = numWeeks + 1
														numVisits = 0
																														
														if filterByCreated="1" and (not rsEntry.EOF) then
															viewNumWeeks = viewNumWeeks + 1
															numWeeks = viewNumWeeks
														end if
														displayed=false
														
													Loop 'Do While numWeeks <= viewNumWeeks
													'********************
													rsEntry.Close
													Set rsEntry = Nothing
													%>
                                                    <tr><td>&nbsp;</td></tr>
																	 <% if request.form("frmTagClients")<>"true" then %>
                                                    <tr>
                                                    	<td colspan="15" align="right">
                                                        	&nbsp;Total Number of Visits: <b><%=TotalNumVisits%></b>
                                                        </td>
                                                    </tr>
																	 <% end if %>
												</table>
											<% end if 'genReport %>
<% if request.form("frmExpReport")<>"true" then %>
										</td>
									</tr>
								</table>
							</td>
						</tr>
						
					</table>
				</td>
			</tr>
			</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

		<%
	end if 'if not Session("Pass") OR Session("Admin")="false" OR NOT (ap_rpt_day OR ap_rpt_day_self) then 
end if 'if Session("StudioID") = "" then
%>
