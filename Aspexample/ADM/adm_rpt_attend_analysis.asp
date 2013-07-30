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
	<!-- #include file="inc_hotword.asp" -->
	<%	dim doRefresh : doRefresh = false %>
	<!-- #include file="inc_date_arrows.asp" -->
	<!-- #include file="../inc_ajax.asp" -->
	<!-- #include file="../inc_val_date.asp" --> 
<%  if session("CR_Memberships") <> 0 then %>
    <!-- #include file="../inc_dbconn_regions.asp" -->
    <!-- #include file="../inc_dbconn_wsMaster.asp" -->
    <!-- #include file="inc_masterclients_util.asp" -->
<%  end if 
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_DAY") then 
	%>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
	<%else%>
		<!-- #include file="../inc_i18n.asp" -->
	<%
	Dim disMode, cSDate, cEDate, cLoc, cVT, cTG, cView
	Dim showHeader, rowcolor, barcolor
	Dim TotalPaid, TotalComp, TotalVisits, PaidPercent, ap_view_all_locs, GrandTotSessions, TotalSignedIn, TotalCancelled
	GrandTotSessions = 0
	
    ' adds display:none; on export to excel
    sub hideOnExport() if request.form("frmExpReport")="true" then response.write "style=""display:none;""" end if end sub

    dim strImgSrc, strHeader
    sub drawMembershipTierHeader(MembershipName, IconNum)
    
        strHeader = MembershipName
        if request.form("frmExpReport")<>"true" then 
            strImgSrc = "<img src=""" & contentUrl("/asp/adm/images/mem-" & IconNum & ".png") & """ />"
            strHeader = MembershipName & strImgSrc
        end if

    	%>
    	<tr class="membershipHeader"><th class="mainText"><%=strHeader%></th></tr>
    	<%	
    end sub

    ' set curMembership
    dim curMembership

	if request.form("requiredtxtDateStart")<>"" then
		Call SetLocale(session("mvarLocaleStr"))
			cSDate = CDATE(request.form("requiredtxtDateStart"))
		Call SetLocale("en-us")
	else
		cSDate = DateAdd("m",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
	end if

	if request.form("requiredtxtDateEnd")<>""  then
		Call SetLocale(session("mvarLocaleStr"))
			cEDate = CDATE(request.form("requiredtxtDateEnd"))
		Call SetLocale("en-us")
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
	end if
	
	disMode = request.form("optDisMode")
	if disMode = "" AND request.querystring("cDetail")="true" then 'if coming from reports->membership
	  disMode = "5"
	end if
	
	cView = request.form("optView")
	if cView = "" AND request.querystring("cDetail")="true" then 'if coming from reports->membership
	  cView = "detail"
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
	
	ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
	
	if request.form("optVT")<>"" then
		cVT = CLNG(request.form("optVT"))
	else
		cVT = 0
	end if
	if request.form("optPmtTG")<>"" then
		cTG = CLNG(request.form("optPmtTG"))
	else
		cTG = 0
	end if
		if NOT request.form("frmExpReport")="true" then
		%>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_attend_analysis")) %>
			<script type="text/javascript">
			function exportReport() {
				document.frmParameter.frmExpReport.value = "true";
				document.frmParameter.frmGenReport.value = "true";
				<% iframeSubmit "frmParameter", "adm_rpt_attend_analysis.asp" %>
			}
			function changeFilters()
			{
				<% if disMode<>"5" then %>
				if (document.frmParameter.optDisMode.value == 5)
				<% else %>
				if (document.frmParameter.optDisMode.value != 5)
				<% end if %>
				{
					document.frmParameter.submit();
				}
			}
			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
			<!-- #include file="inc_help_content.asp" -->
		<%
		end if
		
		%>
		
        <style>
        .membershipHeader {
          font:"<%=session("pageColor4")%>";
          font-weight:bold;
          text-align:left;
        }
        .membershipHeader img {
          width:16px;
          height:16px;
          padding-left:.5em;
        }
        </style>
		
		<% if NOT request.form("frmExpReport")="true" then %>
<% pageStart %>
			<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
				<tr> 
				<td valign="top" height="100%" width="100%"> 
				<table cellspacing="0" width="90%" height="100%" style="margin: 0 auto;">
					<tr>
					<td class="headText" align="left" valign="top">
					<table width="100%" cellspacing="0">
						<tr>
						<td class="headText" valign="bottom"><b> <%= pp_PageTitle("Attendance Analysis") %> </b>
						 <%if session("Admin")="sa" then %>
                             <a class="mainText" href="/Report/Clients/AttendanceAnalysis">Current version</a>                               
                         <%end if %>
						<!--JM - 49_2447-->
						<% showNewHelpContentIcon("attendance-analysis-report") %>
						
						</td>
						<td valign="bottom" class="right" height="26"> </td>
						</tr>
					</table>
					</td>
					</tr>
					<tr> 
					<td height="30" valign="bottom" class="headText">
					<table class="mainText border4 center-block" cellspacing="0">
						<form name="frmParameter" action="adm_rpt_attend_analysis.asp" method="POST">
						<input type="hidden" name="frmGenReport" value="">
						<input type="hidden" name="frmExpReport" value="">
						<tr> 
						<td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b><span style="color:<%=session("pageColor4")%>">&nbsp;</span><%=xssStr(allHotWords(77))%>: 
						<input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text" size="11"  id="requiredtxtDateStart" name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
						<script type="text/javascript">
						var cal1 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStart'});
						cal1.a_tpl.yearscroll = true;
						</script>
						&nbsp;<%=xssStr(allHotWords(79))%>: 
						<input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text" size="11" name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
						<script type="text/javascript">
						var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
						cal2.a_tpl.yearscroll = true;
						</script>
						&nbsp;
						<%=xssStr(allHotWords(8))%>:<select name="optSaleLoc" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
						<option value="0" <% if cLoc=0 then response.write "selected" end if %>>All</option>
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
							document.frmParameter.optSaleLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
						</script>
						&nbsp;Analysis By&nbsp;<select name="optDisMode" onChange="changeFilters();">
						  <option value="0" <%if disMode="0" then response.write "selected" end if%>>Class Time</option>
						  <option value="1" <%if disMode="1" then response.write "selected" end if%>>Day of Week</option>
						  <option value="2" <%if disMode="2" then response.write "selected" end if%>><%= getHotWord(61)%></option>
						  <option value="3" <%if disMode="3" then response.write "selected" end if%>>Visit Type</option>
						  <option value="4" <%if disMode="4" then response.write "selected" end if%>><%= getHotWord(6)%></option>
						  <option value="5" <%if disMode="5" then response.write "selected" end if%>><%= getHotWord(12)%></option>
						</select>
						<script type="text/javascript">
							document.frmParameter.optDisMode.options[4].text = '<%=jsEscSingle(allHotWords(6))%>';
							document.frmParameter.optDisMode.options[5].text = '<%=jsEscSingle(allHotWords(12))%>';
						</script>
						
						&nbsp;<%= getHotWord(6)%>:&nbsp;
						<select name="optTrn">
							<option value="0" selected><%=xssStr(allHotWords(216))%></option>
						<%
						strSQL = "SELECT TrainerID, TrFirstName, TrLastName, DisplayName FROM Trainers WHERE Active=1 AND [Delete]=0 AND ((ReservationTrn=1 OR AppointmentTrn=1) OR Assistant=1 OR [Employee]=1) AND TrainerID > 1 AND isSystem=0 ORDER BY TrLastName"
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						Set rsEntry.ActiveConnection = Nothing
						do while not rsEntry.EOF
						%>
							<option value="<%= rsEntry("TrainerID")%>" <% if CLNG(request.form("optTrn"))=CLNG(rsEntry("TrainerID")) then response.write " selected" end if%>><%=rsEntry("TrLastName")%>, <%=rsEntry("TrFirstName")%> </option>
						<%
						rsEntry.MoveNext
						loop
						rsEntry.Close
						%>
						</select>
						<br />
						 <!--by JM 04_14_2008 Add Program/TypeGroup Filter-->
						&nbsp;<%=xssStr(allHotWords(7))%>:
						<%
						strSQL = "SELECT tblTypeGroup.TypegroupID as TG, tblVisitTypes.TypeID as VT,"
						strSQL = strSQL & " tblTypeGroup.TypeGroup as Program, "
						strSQL = strSQL & " tblVisitTypes.TypeName as VisitType, tblTypeGroup.wsArrival "
						strSQL = strSQL & " FROM tblVisitTypes RIGHT OUTER JOIN tblTypeGroup ON tblVisitTypes.Typegroup = tblTypeGroup.TypeGroupID "
						strSQL = strSQL & " WHERE (tblTypeGroup.Active = 1) AND (tblVisitTypes.TypeID IS NULL OR ( ((tblVisitTypes.[Delete] = 0) AND (tblVisitTypes.Active = 1))) ) "
						strSQL = strSQL & " ORDER BY tblTypeGroup.TypeGroup, tblVisitTypes.SortOrder, tblVisitTypes.TypeName"
					response.write debugSQL(strSQL, "SQL")
						rsEntry.CursorLocation = 3
						response.write debugSQL(strSQL, "SQL")
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing
							dim TGVT, LastTG, FirstTG%>
							<select name="optTGVT" >
							<option value="0">All Programs & Session Types</option>
	
							<%do while NOT rsEntry.EOF
								if rsEntry("TG") <> LastTG then 
									FirstTG = rsEntry("TG")%>
									<option value="<%=FirstTG%>" <%if cstr(request.form("optTGVT")) = cstr(rsEntry("TG")) then response.write "selected" end if%>><%=rsEntry("Program")%></option>
								<%end if%>
								<% if NOT rsEntry("wsArrival") and NOT isNull(rsEntry("VT")) then 'CB 49_2615%>
									<%TGVT = "#"&rsEntry("VT")%>
									<option value="<%=TGVT%>" <%if request.form("optTGVT")= TGVT then response.write "selected" end if%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rsEntry("VisitType")%></option> 
								<% end if %>
							<%
								rsEntry.MoveNext
								LastTG = FirstTG
							loop%>
							</select>
							<img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" title="To view statistics for your Arrivals Program, you must specifically select it with this dropdown box.">
							<%rsEntry.close%>
												
						&nbsp;&nbsp;<% taggingFilter %>&nbsp;&nbsp;
						<br />
						<% if disMode="5" then %>
						<%=xssStr(allHotWords(159))%>:&nbsp;
						<select name="optView">
							<option value="summary" <% if cView="summary" then response.write " selected" end if %>>Summary</option>
							<option value="detail" <% if cView="detail" then response.write " selected" end if %>>Detail</option>
						</select>
						&nbsp;&nbsp;&nbsp;
						Show <%= getHotWord(12)%>s with&nbsp;
						<input type="text" name="txtMinVisits" size="1" maxlength="3" value="<% if request.form("txtMinVisits")<>"" then response.write request.form("txtMinVisits") else response.write "0" end if %>">
						&nbsp;or more visits.&nbsp;&nbsp;
						<br />
						<% end if %>
						Include No Shows&nbsp;
						<input type="checkbox" name="optIncludeNoShows" <% if request.form("optIncludeNoShows")="on" then response.write "checked" end if %> />
						
						<% showDateArrows("frmParameter") %>
						<input type="button" name="Button" value="Generate" onClick="genReport();">
						<% exportToExcelButton %>
						<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
						 else%>
								<% taggingButtons("frmParameter") %>
						<%end if%>
						<% savingButtons "frmParameter", "Attendance Analysis" %>
						</b></td>
						</tr>
						
						</form>
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" class="mainTextBig"> 
					
					<table class="mainText" width="100%" cellspacing="0">
						<tr>
						<td  colspan="2" valign="top" class="mainTextBig center-ch">&nbsp;</td>
						</tr>
						<tr > 
						<td class="mainTextBig" colspan="2" valign="top">
		<% 
		end if			'end of frmExpreport value check before /head line	  

							if request.form("frmGenReport")="true" OR request.form("frmTagClients") then 
							
								if request.form("frmTagClients")="true" then
								
									strSQL = "SELECT DISTINCT [VISIT DATA].ClientID FROM [VISIT DATA]  "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
									strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if									
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if				
									if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
										strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
									end if
									
									'MB bug #4318, optIncludeNoShows was not taken into consideration for tagging query
									if request.Form("optIncludeNoShows")<>"on" then
										strSQL = strSQL & " AND ([VISIT DATA].Missed = 0) "
									end if	
																	
									if request.form("txtMinVisits")<>"" then
										strSQL = strSQL & "GROUP BY [VISIT DATA].ClientID "
										strSQL = strSQL & "HAVING (COUNT([VISIT DATA].VisitRefNo) >= " & request.form("txtMinVisits") & ") "
									end if
									if request.form("frmTagClientsNew")="true" then
										clearAndTagQuery(strSQL)
									else
										tagQuery(strSQL)
									end if
									
								end if	'End Tag Clients

								if request.form("frmExpReport")="true" then
									Dim stFilename
									if disMode="0" then 
										stFilename="attachment; filename=Attendance Analysis By Class Time " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									elseif disMode="1" then
										stFilename="attachment; filename=Attendance Analysis By Day of Week " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									elseif disMode="2" then
										stFilename="attachment; filename=Attendance Analysis By Series Used " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									elseif disMode="3" then
										stFilename="attachment; filename=Attendance Analysis By Visit Type " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls"
									elseif disMode="4" then
										stFilename="attachment; filename=Attendance Analysis By Instructor " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									else
										stFilename="attachment; filename=Attendance Analysis By Client " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls"
									end if
									Response.ContentType = "application/vnd.ms-excel" 
									Response.AddHeader "Content-Disposition", stFilename 
								end if
								
								showHeader = "false"
								TotalPaid=0
								TotalComp=0
								PaidPercent=0
								TotalVisits = 0

								'if disMode<>"5" then		'NOT By Client
									strSQL = "SELECT SUM([VISIT DATA].[Value]) as TotPaid, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 0 ELSE 1 END) AS TotComp "
									strSQL = strSQL & "FROM [VISIT DATA]   "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
									end if
									if request.form("txtMinVisits")<>"" then
										strSQL = strSQL & "INNER JOIN (SELECT ClientID FROM [VISIT DATA]  WHERE ([VISIT DATA].VisitType<>-2) AND ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") GROUP BY ClientID HAVING (COUNT([VISIT DATA].VisitRefNo) >= " & request.form("txtMinVisits") & ")) ClientGroup ON [Visit Data].ClientID = ClientGroup.ClientID "
									end if
									strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if
									
									if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
										strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
									end if
									if request.form("optFilterTagged")="on" then
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
									if request.Form("optIncludeNoShows")<>"on" then
										strSQL = strSQL & "AND ([VISIT DATA].Missed = 0) "
									end if
							response.write debugSQL(strSQL, "SQL")

								'response.end
									rsEntry2.CursorLocation = 3
									rsEntry2.open strSQL, cnWS
									Set rsEntry2.ActiveConnection = Nothing
	
									If Not rsEntry2.eof then
										TotalVisits = 0
										if NOT isNULL(rsEntry2("TotPaid")) then
											TotalPaid = rsEntry2("TotPaid")
											TotalVisits = TotalVisits + rsEntry2("TotPaid")
										end if
										if NOT isNULL(rsEntry2("TotComp")) then
											TotalComp = rsEntry2("TotComp")
											TotalVisits = TotalVisits + rsEntry2("TotComp")
										end if
	
									end if
									rsEntry2.close
								'end if	'NOT By Client
								
				
								if disMode="0" then		'By Class Time
									strSQL = "SELECT  [VISIT DATA].ClassTime, SessCountTBL.SessCount, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 1 ELSE 0 END) AS paid, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 0 ELSE 1 END) AS comp, "
									strSQL = strSQL & "COUNT(DISTINCT [VISIT DATA].ClientID) AS UniqueClient "
									strSQL = strSQL & "FROM [VISIT DATA]   "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
									end if

									''Dual Sub Query added 
									strSQL = strSQL & " INNER JOIN  (SELECT COUNT(*) AS SessCount, ClassTime FROM (SELECT ClassTime, COUNT(VisitRefNo) AS NumVisits FROM [VISIT DATA]  "
									strSQL = strSQL & " WHERE 1=1 "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if								
									if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
										strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
									end if
									strSQL = strSQL & " GROUP BY TrainerID, ClassDate, ClassTime HAVING (ClassDate >= " & DateSep & cSDate & DateSep & ") AND (ClassDate <= " & DateSep & cEDate & DateSep & ")) NumSess GROUP BY ClassTime) SessCountTBL ON [VISIT DATA].ClassTime = SessCountTBL.ClassTime "
									''Dual Sub Query added 

									strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if									
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if							
									if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
										strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
									end if
									if request.form("optFilterTagged")="on" then
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
									if request.Form("optIncludeNoShows")<>"on" then
										strSQL = strSQL & "AND ([VISIT DATA].Missed = 0) "
									end if
									strSQL = strSQL & "GROUP BY [VISIT DATA].ClassTime, SessCountTBL.SessCount  "
									strSQL = strSQL & "ORDER BY [VISIT DATA].ClassTime, Paid DESC  "
								elseif disMode="1" then		'By Day of week
									strSQL = "SELECT (CASE DatePart(dw,[VISIT DATA].ClassDate) "
									strSQL = strSQL & "WHEN 1 THEN 'SUNDAY' "
									strSQL = strSQL & "WHEN 2 THEN 'MONDAY' "
									strSQL = strSQL & "WHEN 3 THEN 'TUESDAY' "
									strSQL = strSQL & "WHEN 4 THEN 'WEDNESDAY' "
									strSQL = strSQL & "WHEN 5 THEN 'THURSDAY' "
									strSQL = strSQL & "WHEN 6 THEN 'FRIDAY' "
									strSQL = strSQL & "WHEN 7 THEN 'SATURDAY' "
									strSQL = strSQL & "END) AS DayofWeek, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 1 ELSE 0 END) AS paid, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 0 ELSE 1 END) AS comp, "
								  strSQL = strSQL & "COUNT(DISTINCT [VISIT DATA].ClientID) AS UniqueClient "
									strSQL = strSQL & "FROM [VISIT DATA]  "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
									end if
									strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if									
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if
									
									if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
										strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
									end if
									if request.form("optFilterTagged")="on" then
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
									if request.Form("optIncludeNoShows")<>"on" then
										strSQL = strSQL & "AND ([VISIT DATA].Missed = 0) "
									end if
									strSQL = strSQL & "GROUP BY DatePart(dw,[VISIT DATA].ClassDate), (CASE DatePart(dw,[VISIT DATA].ClassDate) "
									strSQL = strSQL & "WHEN 1 THEN 'SUNDAY' "
									strSQL = strSQL & "WHEN 2 THEN 'MONDAY' "
									strSQL = strSQL & "WHEN 3 THEN 'TUESDAY' "
									strSQL = strSQL & "WHEN 4 THEN 'WEDNESDAY' "
									strSQL = strSQL & "WHEN 5 THEN 'THURSDAY' "
									strSQL = strSQL & "WHEN 6 THEN 'FRIDAY' "
									strSQL = strSQL & "WHEN 7 THEN 'SATURDAY' "
									strSQL = strSQL & "END) "
									strSQL = strSQL & "ORDER BY DatePart(dw,[VISIT DATA].ClassDate) "
									
								elseif disMode="2" then		'By Series
									strSQL = "SELECT  [VISIT DATA].TypeTaken, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 1 ELSE 0 END) AS paid, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 0 ELSE 1 END) AS comp, "
									strSQL = strSQL & "COUNT(DISTINCT [VISIT DATA].ClientID) AS UniqueClient "
									strSQL = strSQL & "FROM [VISIT DATA]   "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
									end if
									strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if									
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if						
									if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
										strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
									end if
									if request.form("optFilterTagged")="on" then
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
									if request.Form("optIncludeNoShows")<>"on" then
										strSQL = strSQL & "AND ([VISIT DATA].Missed = 0) "
									end if
									strSQL = strSQL & "GROUP BY [VISIT DATA].TypeTaken "
									strSQL = strSQL & "ORDER BY Paid DESC, [VISIT DATA].TypeTaken "
								elseif disMode="3" then		'By Visit Type
									strSQL = "SELECT  [VISIT DATA].VisitType, tblVisitTypes.TypeName, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 1 ELSE 0 END) AS paid, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 0 ELSE 1 END) AS comp, "
									strSQL = strSQL & "COUNT(DISTINCT [VISIT DATA].ClientID) AS UniqueClient "
									strSQL = strSQL & "FROM [VISIT DATA]   "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
									end if
									strSQL = strSQL & "LEFT OUTER JOIN tblVisitTypes ON [VISIT DATA].VisitType = tblVisitTypes.TypeID "
									strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if									
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if						
									if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
										strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
									end if
									if request.form("optFilterTagged")="on" then
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
									if request.Form("optIncludeNoShows")<>"on" then
										strSQL = strSQL & "AND ([VISIT DATA].Missed = 0) "
									end if
									strSQL = strSQL & "GROUP BY [VISIT DATA].VisitType, tblVisitTypes.TypeName "
									strSQL = strSQL & "ORDER BY Paid DESC, tblVisitTypes.TypeName "

								elseif disMode="4" then		'By Instructor
									strSQL = "SELECT DISTINCT [VISIT DATA].TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName, SessCountTBL.SessCount, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 1 ELSE 0 END) AS paid, "
									strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value]=1 THEN 0 ELSE 1 END) AS comp, "
									strSQL = strSQL & "COUNT(DISTINCT [VISIT DATA].ClientID) AS UniqueClient "
									strSQL = strSQL & "FROM [VISIT DATA]   "
									if request.form("optFilterTagged")="on" then
										strSQL = strSQL & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
									end if
									
									''Dual Sub Query added 
									strSQL = strSQL & " INNER JOIN  (SELECT COUNT(*) AS SessCount, TrainerID FROM (SELECT TrainerID, COUNT(VisitRefNo) AS NumVisits FROM [VISIT DATA]  "
									strSQL = strSQL & " WHERE 1=1 "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if									
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if							
									if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
										strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
									end if
									strSQL = strSQL & " GROUP BY TrainerID, ClassDate, ClassTime HAVING (ClassDate >= " & DateSep & cSDate & DateSep & ") AND (ClassDate <= " & DateSep & cEDate & DateSep & ")) NumSess GROUP BY TrainerID) SessCountTBL ON [VISIT DATA].TrainerID = SessCountTBL.TrainerID "
									''Dual Sub Query added 
									
									strSQL = strSQL & "LEFT OUTER JOIN TRAINERS ON [VISIT DATA].TrainerID = TRAINERS.TrainerID "
									strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") "
									strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
									if cLoc<>0 then
										strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
									end if									
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if					
									if request.form("optFilterTagged")="on" then
										if session("mvaruserID")<>"" then
											strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
										else
											strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
										end if
									end if
									if request.Form("optIncludeNoShows")<>"on" then
										strSQL = strSQL & "AND ([VISIT DATA].Missed = 0) "
									end if
									strSQL = strSQL & "GROUP BY [VISIT DATA].TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, SessCountTBL.SessCount "
									strSQL = strSQL & "ORDER BY " & GetTrnOrderBy() & ", Paid DESC"
									
								elseif disMode="5" then		'By Client
								
									if request.form("optView")="detail" then

										'BJD: 49_2212 - added new totals to SQL
										strSQL = "SELECT [VISIT DATA].ClientID, CLIENTS.RSSID, CLIENTS.FirstName, CLIENTS.LastName, COUNT([VISIT DATA].VisitRefNo) AS NumVisits, Membership.SeriesTypeID, Membership.TypePurch, Membership.IconNum, "
										strSQL = strSQL & "SUM(CASE WHEN [VISIT DATA].[Value] = 1 THEN 1 ELSE 0 END) AS paid, SUM(CASE WHEN [VISIT DATA].[Value] = 1 THEN 0 ELSE 1 END) AS comp, SUM(CASE WHEN [VISIT DATA].[Missed] = 0 THEN 1 ELSE 0 END) AS SignedIn, SUM(CASE WHEN [VISIT DATA].[Cancelled] = 1 THEN 1 ELSE 0 END) AS Cancelled "
										strSQL = strSQL & "FROM [VISIT DATA]  INNER JOIN CLIENTS ON [VISIT DATA].ClientID = CLIENTS.ClientID "
										if request.form("optFilterTagged")="on" then
											strSQL = strSQL & "INNER JOIN tblClientTag ON [VISIT DATA].ClientID = tblClientTag.clientID "
											if session("mvaruserID")<>"" then
												strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
											else
												strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
											end if
										end if
										strSQL = strSQL & "LEFT OUTER JOIN " & memberSubSQL("") & " Membership ON CLIENTS.ClientID = Membership.ClientID AND CLIENTS.IsSystem=0 "
										strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") "
										strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
										if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
											strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
										end if
										if cLoc<>0 then
											strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
										end if
										if LEFT(request.form("optTGVT"),1) = "#" then
											cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
											strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
										elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
											cTG = CLNG(request.form("optTGVT"))
											strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
										else
											strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
										end if						
										if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
											strSQL = strSQL & "AND ([VISIT DATA].TrainerID = " & request.form("optTrn") & ") "
										end if
										if request.Form("optIncludeNoShows")<>"on" then
											strSQL = strSQL & "AND ([VISIT DATA].Missed = 0) "
										end if
										strSQL = strSQL & "GROUP BY [VISIT DATA].ClientID, CLIENTS.FirstName, CLIENTS.LastName, CLIENTS.RSSID, Membership.SeriesTypeID, Membership.TypePurch, Membership.IconNum "
										if request.form("txtMinVisits")<>"" then
											strSQL = strSQL & "HAVING (COUNT([VISIT DATA].VisitRefNo) >= " & request.form("txtMinVisits") & ") "
										end if
										strSQL = strSQL & "ORDER BY TypePurch, NumVisits DESC "
									else ' summary view
										
										strSQL = "SELECT Total.TotalClients, COUNT(VisitingClients.ClientID) AS NumClients "
										strSQL = strSQL & "FROM (SELECT COUNT(*) AS TotalClients FROM CLIENTS WHERE (Inactive = 0) AND (Deleted = 0)) Total CROSS JOIN (SELECT DISTINCT [VISIT DATA].ClientID FROM [VISIT DATA]  "
										strSQL = strSQL & "WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") "
										strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
										if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
											strSQL = strSQL & "AND [VISIT DATA].TrainerID = " & request.form("optTrn") & " "
										end if
										if cLoc<>0 then
											strSQL = strSQL & "AND ([VISIT DATA].Location=" & cLoc & ") "
										end if
									if LEFT(request.form("optTGVT"),1) = "#" then
										cVT = CLNG(REPLACE(request.form("optTGVT"),"#",""))
										strSQL = strSQL & "AND [VISIT DATA].VisitType=" & cVT & " "
									elseif LEFT(request.form("optTGVT"),1) <> "#" and request.form("optTGVT") <> "0" then
										cTG = CLNG(request.form("optTGVT"))
										strSQL = strSQL & "AND [VISIT DATA].TypeGroup=" & cTG & " "
									else
										strSQL = strSQL & "AND ([VISIT DATA].VisitType<>-2) "
									end if								
										if request.form("optTrn")<>"" AND request.form("optTrn")<>"0" then
											strSQL = strSQL & "AND ([VISIT DATA].TrainerID = " & request.form("optTrn") & ") "
										end if
										if request.Form("optIncludeNoShows")<>"on" then
											strSQL = strSQL & "AND ([VISIT DATA].Missed = 0) "
										end if
										strSQL = strSQL & "GROUP BY ClientID "
										if request.form("txtMinVisits")<>"" then
											strSQL = strSQL & "HAVING (COUNT([VISIT DATA].VisitRefNo) >= " & request.form("txtMinVisits") & ")"
										end if
										strSQL = strSQL & ") VisitingClients GROUP BY Total.TotalClients "
									
									end if
								
								end if
								
							 response.write debugSQL(strSQL, "SQL")
							'response.end
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								Set rsEntry.ActiveConnection = Nothing

								%>
									<table class="mainText"  cellspacing="0" width="100%">
								<% 
								if disMode<>"5" OR request.form("optView")="detail" then ' if normal (not by clients summary view
									if NOT rsEntry.EOF then			'EOF

										do while NOT rsEntry.EOF

                                            if request.form("optView")="detail" then
                                            	if curMembership<>rsEntry("TypePurch") then
                                                    curMembership = rsEntry("TypePurch")
                                                    showHeader = "false"
                                            	end if
                                            end if
											if showHeader = "false" then
                                                if request.form("optView")="detail" then
                                                	if rsEntry("TypePurch")<>"" then
                                                		if NOT request.form("frmExpReport")="true" then %>
                                                			<tr height="2">
                                                				<td colspan="9"   style="background-color:#666666;height:1px;line-height:1px;font-size:1px;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
                                                			</tr>
                                                		<% end if
                                                		%>
                                                		<tr><td colspan="9">&nbsp;</td></tr>
                                                		<tr><td colspan="9">&nbsp;</td></tr>
                                                		<%
                                                		drawMembershipTierHeader rsEntry("TypePurch"), rsEntry("IconNum")
                                                	end if
                                                end if
												%>

												<tr class="right">
													<td width="20%"><strong>
														<%if disMode="0" then%> Class Time
														<%elseif disMode="1" then%>Day of Week
														<%elseif disMode="2" then%><%= getHotWord(61)%>
														<%elseif disMode="3" then%>Visit Type
														<%elseif disMode="4" then%><%= getHotWord(6)%>
														<%elseif disMode="5" then%><%= getHotWord(12)%>
														<%end if%>
													</strong></td>
													<td width="10%"><strong>Paid Visits</strong></td>
													<% if disMode<>"5" then	'NOT By Client %>
														<td><strong>% of Total<span class="maintextbig">*</span></strong></td>
														<td>&nbsp;</td>
														<td><strong>&nbsp;&nbsp;Unique <%=session("ClientHW")%>s</strong></td>
														<td><strong>&nbsp;&nbsp;Comp/Guest Visits</strong></td>
													<% end if 'NOT By Client %>
													
													<% if disMode="5" then %>
														<td><strong>&nbsp;&nbsp;Signed In Visits</strong></td>
														<% if request.Form("optIncludeNoShows")="on" then %>
														<td><strong>&nbsp;&nbsp;No Shows</strong></td>
														<% end if %>
														<td><strong>&nbsp;&nbsp;Cancelled</strong></td>
														<td><strong>&nbsp;&nbsp;<%= getHotWord(134)%></strong></td>
													<% end if %>
													<td><strong>&nbsp;&nbsp;Total Visits</strong></td>
													<%if disMode="4" OR disMode="0" then%>
														<td><strong>&nbsp;&nbsp;Total Sessions</strong></td>
														<td><strong>&nbsp;&nbsp;Avg</strong></td>
													<%end if%>
												</tr>
												<% if NOT request.form("frmExpReport")="true" then %><tr height="2"><td colspan="9"   style="background-color:#666666;height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td></tr><% end if %>
												<%	
											end if
											showHeader = "true"
											
											if TotalPaid = 0 then
												PaidPercent = 0
											else
												'response.write TotalPaid & "||  " & rsEntry("Paid") & "  || <br /><br />"
												PaidPercent=FormatNumber((rsEntry("Paid")/TotalPaid), 3) * 100
											end if
											

											if rowColor = "#F2F2F2" then
												rowColor = "#FAFAFA"
											else
												rowColor = "#F2F2F2"
											end if

											'' bar graph color rotation
											if barColor = session("pageColor4") then
												barColor = session("pageColor3")
											elseif barColor = session("pageColor3") then
												barColor = session("pageColor2")
											else
												barColor = session("pageColor4")
											end if
											
											%>
                                            <tr class="right" style="background-color:<%=rowColor%>;">
                                                <td><strong>
													<%
                                                    If disMode="0" then
                                                        response.write fmtTimeShort(rsEntry("ClassTime"))
                                                    elseIf disMode="1" then
                                                        response.write rsEntry("DayofWeek")
                                                    elseIf disMode="2" then
                                                        response.write rsEntry("Typetaken")
                                                    elseIf disMode="3" then
                                                        response.write rsEntry("TypeName")
                                                    elseif disMode="4" then
                                                        response.write FmtTrnNameNew(rsEntry, false)
                                                    elseif disMode="5" then
                                                        response.write rsEntry("FirstName") & " " & rsEntry("LastName")
                                                    end if
                                                    %>
												</strong></td>
                                          		<td><%=rsEntry("Paid")%></td>
                                          		
												<% if disMode<>"5" then		'NOT By Client %>
												  	<td><%=PaidPercent%>%</td>
												  	<td>
														<% if NOT request.form("frmExpReport")="true" then%>
													 		<table align="left" style="background-color:<%=barcolor%>;" width="<%if TotalPaid = 0 then response.write "0"	else response.write (rsEntry("Paid")/TotalPaid)*360	end if%>"><tr height="9"><td style="height:9px;line-height:9px;font-size:9px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr></table>
														<% end if 'NOT By Client%>
												  	</td>				
														<td><%=rsEntry("UniqueClient")%></td>
														<td><%=rsEntry("Comp")%></td>					  
												<%end if%>
												
												<% if disMode="5" then %>
                                                    <td><%=rsEntry("SignedIn")%></td>
                                                    <% if request.Form("optIncludeNoShows")="on" then %>
	                                                    <td><%=rsEntry("NumVisits") - rsEntry("SignedIn")%></td>
                                                    <% end if %>
                                                    <td><%=rsEntry("Cancelled")%></td>
                                                    <td><a title="Go To <%=session("ClientHW")%> Info" href='main_info.asp?id=<%=rsEntry("ClientID")%>&fl=true'><%=rsEntry("RSSID")%></a></td>
            										<%TotalSignedIn = TotalSignedIn + cInt(rsEntry("SignedIn"))%>
            										<%TotalCancelled = TotalCancelled + cInt(rsEntry("Cancelled")) %>
												<% end if %>
												
												<td><%=rsEntry("Comp")+rsEntry("Paid")%></td>
												
												<%if disMode="4" OR disMode="0" then%>
													<% GrandTotSessions = GrandTotSessions + rsEntry("SessCount") %>
                                                    <td><%=rsEntry("SessCount")%></td>
                                                    <td><%=FmtNumber((rsEntry("Comp")+rsEntry("Paid"))/rsEntry("SessCount"))%></td>
												<%end if%>												
											</tr>
											<%		
											rsEntry.MoveNext
										loop 'while not rsEntry.EOF
										%>
										<% if NOT request.form("frmExpReport")="true" then %><tr height="2"><td colspan="9"   style="background-color:#666666;height:1px;line-height:1px;font-size:1px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td></tr><% end if %>
										
											<tr class="right">
												<td align="left"><strong><%= getHotWord(22)%>:</strong></td>
												<td><strong><%=TotalPaid%></strong></td>
												<% if disMode<>"5" then 'NOT By Client %>
    												<td colspan="2">&nbsp;</td>
														<td>&nbsp</td>
														<td><strong><%=TotalComp%></strong></td>
												<% end if %>
                                                <% if request.form("optView")="detail" then %>
                                                <td><%=TotalSignedIn%></td>
                                                <% if request.Form("optIncludeNoShows")="on" then %>
	                                                <td><%=TotalVisits - TotalSignedIn%></td>
                                                <% end if %>
                                                <td><%=TotalCancelled%></td>
    												<td>&nbsp;</td>
                                                <% end if %>
                                                
												<td><strong><%=TotalVisits%></strong></td>
												<%if disMode="4" OR disMode="0" then%>
                                                    <td><strong><%=GrandTotSessions%></strong></td>
                                                    <td><strong><%=FmtNumber(TotalVisits/GrandTotSessions)%></strong></td>
												<%end if%>
											</tr>
                                        <tr>
                                            <td colspan="9">&nbsp;</td>
                                        </tr>
                                        <tr>
                                            <td colspan="9">*: Total does not include comp/guests.</td>
                                        </tr>
									<% 
									end if 'eof
								else 'by clients summary view
									if NOT rsEntry.EOF then	'EOF 
										%>
										<tr>
											<td colspan="4">&nbsp;</td>
										</tr>
                                        <tr class="right">
                                            <td width="25%" class="center-ch"><b>Visiting <%= getHotWord(12)%>s</b></td>
                                            <td width="25%" class="center-ch"><strong><%= getHotWord(22)%>&nbsp;<%= getHotWord(12)%>s</strong></td>
                                            <td width="25%" class="center-ch" colspan="2"><b>% of Total</b></td>
                                        </tr>
                                        <tr height="2">
                                            <td colspan="4"   style="background-color:#666666;height:1px;line-height:1px;font-size:1px;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
                                        </tr>
                                        <tr>
                                            <td class="center-ch"><%=rsEntry("NumClients")%></td>
                                            <td class="center-ch"><%=rsEntry("TotalClients")%></td>
                                            <td class="center-ch"><%=FormatNumber((rsEntry("NumClients")/rsEntry("TotalClients")), 3) * 100%>%</td>
                                            <td>
                                                <% if NOT request.form("frmExpReport")="true" then%>
                                                     <table align="left" style="background-color:<%=session("pageColor4")%>;" width="<%if rsEntry("TotalClients") = 0 then response.write "0"	else response.write (rsEntry("NumClients")/rsEntry("TotalClients"))*360	end if%>"><tr height="9"><td style="height:9px;line-height:9px;font-size:9px;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr></table>
                                                <%end if%>
                                            </td>											
                                        </tr>
									<% end if 'EOF%>
								<% end if %>
							</table>
							<%
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
