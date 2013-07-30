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
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_CCP") then 
		%>

		<script type="text/javascript">
			alert("<%=DisplayPhraseJS(systemMessagesErrorsDictionary, "Usernamenotauthorizedtoviewthisscreen") %>");
			javascript:history.go(-1);
		</script>
		<%
	else
		%>
		<!-- #include file="../inc_i18n.asp" -->
		<!-- #include file="inc_rpt_tagging.asp" -->
		<!-- #include file="inc_hotword.asp" -->
		<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

		VoidOk = validAccessPriv("TB_VOID")
        
        dim phraseDictionary 
        set phraseDictionary = LoadPhrases("BusinessmodependingtransactionsPage", 141)

        dim reportDictionary
        set reportDictionary =  LoadPhrases("ReportmasterPage", 82)
	
		dim category : category = ""
		if (RQ("category"))<>"" then
			category = RQ("category")
		elseif (RF("category"))<>"" then
			category = RF("category") 
        end if

		Dim cSDate, cEDate, disMode, ccProcessor, ccProcessor2, isBatchProcessor, numTrx, tmpCounter, totalAmt, numFTPAccounts, ACHHotWord, ss_EnableACH
		ccProcessor = checkStudioSetting("Studios", "ccProcessor")
        if implementationSwitchIsEnabled("BluefinCanada") then
            if ccProcessor = "PMN" then
                ccProcessor2 = checkStudioSetting("Studios", "ccProcessor2")
            end if
        end if
		numTrx = 0
		tmpCounter = 0
		totalAmt = 0
        if implementationSwitchIsEnabled("BluefinCanada") then
            if ccProcessor = "MON" OR ccProcessor = "OP" OR ccProcessor = "HSBC" OR (ccProcessor = "PMN" AND ccProcessor2 = "ELV")  then
			    isBatchProcessor = true
		    else
			    isBatchProcessor = false
		    end if
        else
            if ccProcessor = "MON" OR ccProcessor = "OP" OR ccProcessor = "HSBC" then
			    isBatchProcessor = true
		    else
			    isBatchProcessor = false
		    end if
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
		if request.form("optDate")="all" then
			disMode = "all"
		else
			disMode = "range"
		end if
		if request.querystring("loc")<>"" then
			ccLoc = request.querystring("loc")
		elseif request.form("optCCLocation")<>"" then
			ccLoc = request.form("optCCLocation")
		else
			'if session("curLocation")="0" then
				ccLoc = "-2"
			'else
			'	ccLoc = session("curLocation")
			'end if
		end if

		Dim expReport
		if request.form("frmExpReport")="true" then
			expReport = "true"
		end if

		dim rsEntry
		set rsEntry = Server.CreateObject("ADODB.Recordset")
	
    	ACHHotWord = "ACH"
		ACHHotWord = xssStr(allHotWords(109))
		
		numFTPAccounts = 1
		strSQL = "SELECT COUNT(*) AS NumAccts FROM Location GROUP BY BankClientID, FTPUsername, FTPPassword, FTPHeaderRecord HAVING (NOT (BankClientID IS NULL)) AND (NOT (FTPUsername IS NULL)) AND (NOT (FTPPassword IS NULL)) AND (NOT (FTPHeaderRecord IS NULL))"
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
		if NOT rsEntry.EOF then
			if NOT isNULL(rsEntry("NumAccts")) then
				numFTPAccounts = rsEntry.RecordCount
			end if
		end if
		rsEntry.close	


		strSQL = "SELECT tblCCOpts.EnableACH FROM tblCCOpts WHERE tblCCOpts.StudioID=" & session("StudioID")
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			if NOT rsEntry.EOF then
				ss_EnableACH = rsEntry("EnableACH")
			end if
			rsEntry.close
			
        if request.Form("frmResetBatch")="true" AND isNum(request.Form("optBatchFileNum")) AND request.Form("optBatchFileNum")<>"0" then
            strSQL = "UPDATE tblCCTrans SET Status=N'Pending', BatchFileNum=null WHERE BatchFileNum=" & request.Form("optBatchFileNum")
            cnWS.execute strSQL
        end if

		
if NOT expReport then 
		
		%>
<!-- #include file="pre.asp" -->
		<!-- #include file="frame_bottom.asp" -->
<%= js(array("mb", "adm/adm_rpt_ccp_ach", "MBS", "reportFavorites")) %>
		<!-- American/Canada=2 format mm/dd/yyyy --> <!-- European/Rest of the world=1 format dd-mm-yyyy -->
		
<%= js(array("calendar" & dateFormatCode)) %>
		<!-- #include file="../inc_date_ctrl.asp" -->
		<!-- #include file="../inc_ajax.asp" -->
		<!-- #include file="../inc_val_date.asp" -->
		<!-- #include file="inc_user_options.asp" -->
		<%
end if 'expReport
		%>
		
<% if NOT expReport then %>
    <!-- #include file="css/site_setup.asp" -->
<% pageStart %>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %> 
	<div class="headText breadcrumbs-old" valign="bottom">
	<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
	<span class="breadcrumb-item">&raquo;</span>
	<%if category <> "" then%>
	<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
	<span class="breadcrumb-item">&raquo;</span>
	<%end if %>
	<%=DisplayPhrase(reportPageTitlesDictionary, "Pendingtransactions")%>

	<div id="add-to-favorites">
		<span id="" class="favorite"></span><span class="text"><%=DisplayPhrase(reportPageTitlesDictionary, "Addtofavorites")%></span>
	</div>
	</div>
<%else %>
<div class="headText" valign="bottom"><b><%= DisplayPhrase(pageTitlesDictionary, "Pendingtransactions") %></b></div>
<%end if %>
<form name="frmACH" id="frmACH" method="POST">
			<table height="100%" width="<%=strPageWidth%>" border="0" cellspacing="0" cellpadding="0">    
				<tr>
					<td valign="top" height="100%" width="100%">
						<table class="center" border="0" cellspacing="0" cellpadding="0" width="90%" height="100%">
							
                                <input type="hidden" name="frmExpReport" value="">
                                <input type="hidden" name="frmGenReport" value="">
                                <input type="hidden" name="frmResetBatch" value=""/>
								<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
									<input type="hidden" name="category" value="<%=category%>">
									<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
								<% end if %>
								<div id="topdiv">
								<tr>
									<td valign="bottom" class="mainText right" height="18">
										<!-- #include file="inc_batch_nav.asp" -->
									</td>
								</tr>
								<tr>
								<td>
								<table class="mainText" cellspacing="0" cellpadding="0" style="float:right;border: 1px solid <%=session("pageColor4")%>">
														<tr>
															<td class="center" valign="bottom" style="background-color:#F2F2F2;" nowrap>
																<b>
							                                        <select name="optDisMode">
								                                        <option value="detail"<% if request.form("optDisMode")="detail" or request.form("optDisMode")="" then response.write " selected" end if %>><%=xssStr(allHotWords(674)) %></option>
								                                        <option value="summary"<% if request.form("optDisMode")="summary" then response.write " selected" end if %>><%=xssStr(allHotWords(675)) %></option>
							                                        </select>
							                                        
							                                        <!--
																	<span style="color:<%=session("pageColor4")%>;">&nbsp;Date&nbsp;Range:</span> 
																	<input type="radio" name="optDate" value="all" <%if disMode="all" then response.write "checked" end if%>>
																	<%= xssStr(allHotWords(149))%>&nbsp;&nbsp; 
																	<input type="radio" name="optDate" value="range" <%if disMode="range" then response.write "checked" end if%>>
																	-->
																	
																	<%=xssStr(allHotWords(77))%>: 
																	<input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
														<script type="text/javascript">
															var cal1 = new tcal({ 'formname': 'frmACH', 'controlname': 'requiredtxtDateStart' });
															cal1.a_tpl.yearscroll = true;
														</script>
																	&nbsp;<%=xssStr(allHotWords(79))%>: 
																	<input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
														<script type="text/javascript">
															var cal2 = new tcal({ 'formname': 'frmACH', 'controlname': 'requiredtxtDateEnd' });
															cal2.a_tpl.yearscroll = true;
														</script>
																	&nbsp; 
							                                        <% if validAccessPriv("RPT_TAG") then 
	  							                                        taggingFilter 
							                                        end if %>
                                                                    <input name="Button" type="button" value="<%=xssStr(allHotWords(226)) %>" onClick="genCCP();">
							                                        <span class="icon-button" style="vertical-align: middle;" title="<%=xssStr(allHotWords(658))%>" ><a onClick="exportCCP();" ><img src="<%= contentUrl("/asp/adm/images/export-to-excel-20px.png") %>" /></a></span>
							                                        <% if NOT validAccessPriv("RPT_TAG") then 
							                                        else 
								                                        taggingButtons("frmACH") 
							                                        end if %>
																</b>
															</td>
														</tr>
													</table>
								</td>
								</tr>
								</div>

								<tr>
									<td valign="top" class="mainTextBig center" height="100%">
										<table class="mainText center" width="95%" border="0" cellspacing="0" cellpadding="0" height="100%">
											<tr>
												<td class="mainText center" colspan="2" valign="top">
<% end if 'expReport %>
													<table class="mainText center" border="0" cellspacing="0" cellpadding="0" width="90%">
<% if NOT expReport then %>
														<tr>
															<td class="right">
																<table class="mainText" width="100%"  border="0" cellspacing="0" cellpadding="0">
																	<tr>
															<td>
																			<% if session("numLocations")>1 then %>
																				<select name="optCCLocation">
																					<option value="-2" <%if ccLoc="-2" then response.write "selected" end if%>><%=xssStr(allHotWords(479)) %></option>
																					<%
																					strSQL = "SELECT LocationID, LocationName FROM Location WHERE  (NOT (MID IS NULL)) AND (Active = 1) ORDER BY LocationName"
																					rsEntry.CursorLocation = 3
																					rsEntry.open strSQL, cnWS
																					Set rsEntry.ActiveConnection = Nothing
																					
																					do while NOT rsEntry.EOF
																						%>
																						<option value="<%=rsEntry("LocationID")%>" <%if ccLoc=CSTR(rsEntry("LocationID")) then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
																						<%
																						rsEntry.MoveNext
																					loop
																					rsEntry.close
																					%>
																				</select>
																			<% end if 'numLocs > 1 %>
																			<% if isBatchProcessor then%>
																				<%if ccProcessor="HSBC" then%>
																					<select name="optBatchTransType">
																						<option value="">All Types</option>
																						<option value="BANK" <%if request.form("optBatchTransType")="BANK" then response.write "selected" end if%>><%=xssStr(allHotWords(109)) %></option>
																						<option value="VISA" <%if request.form("optBatchTransType")="VISA" then response.write "selected" end if%>><%=xssStr(allHotWords(660)) %></option>
																						<option value="AMEX" <%if request.form("optBatchTransType")="AMEX" then response.write "selected" end if%>><%=xssStr(allHotWords(659)) %></option>
																					</select>							  
																				<%else%>
																					<input type="hidden" name="optBatchTransType" value="BANK">
																				<%end if%>
																				<%
																				strSQL = "SELECT DISTINCT tblCCTrans.BatchFileNum "
																				strSQL = strSQL & "FROM tblCCTrans "
																				strSQL = strSQL & "WHERE (NOT (tblCCTrans.BatchFileNum IS NULL)) "
																				if ccLoc<>"-2" then
																					strSQL = strSQL & " AND tblCCTrans.LocationID=" & ccLoc
																				end if
																				if disMode = "range" then
																					strSQL = strSQL & " AND tblCCTrans.TransTime >= " & DateSep & cSDate & DateSep & " "
																					strSQL = strSQL & " AND tblCCTrans.TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep & " "
																				end if
																				strSQL = strSQL & "ORDER BY tblCCTrans.BatchFileNum DESC"
																				%>
																				<select name="optBatchFileNum">
																					<option value="0"><%=DisplayPhraseAttr(phraseDictionary, "Allbatchfiles") %></option>
																					<%
																					rsEntry.CursorLocation = 3
																					rsEntry.open strSQL, cnWS
																					Set rsEntry.ActiveConnection = Nothing
																					do While NOT rsEntry.EOF
																						%>
																						<option value="<%=rsEntry("BatchFileNum")%>" <% if request.form("optBatchFileNum")=CSTR(rsEntry("BatchFileNum")) then response.write "selected" end if %>><%=xssStr(allHotWords(676)) %> <%=rsEntry("BatchFileNum")%></option>
																						<%
																						rsEntry.MoveNext
																					loop
																					rsEntry.close
																					%>
																				</select>
																				<select name="optStatus">
																					<option value="">All Status</option>
																					<option value="Pending" <%if request.form("optStatus")="Pending" then response.write "selected" end if%>><%=xssStr(allHotWords(109))%> - <%=xssStr(allHotWords(621))%></option>
																					<option value="Sent to Bank" <%if request.form("optStatus")="Sent to Bank" then response.write "selected" end if%>><%=xssStr(allHotWords(109))%> - <%=xssStr(allHotWords(678))%></option>
																					<option value="Open" <%if request.form("optStatus")="Open" then response.write "selected" end if%>><%=DisplayPhrase(phraseDictionary, "Ccopen") %></option>
																				</select>
																				
																			    <%if ccProcessor = "HSBC" then %>
																			        &nbsp;<strong><a href="../upload/adm_rpt_ccp_ach_upload.asp">[<%=DisplayPhrase(phraseDictionary, "Uploadresults") %>]</a></strong>
																			    <%end if %>
																				
																			<% end if 	'isBatchProcessor %>
																		</td>
																		<td class="right" valign="bottom"> 
																			<%if isBatchProcessor then%>
																				<a href="javascript:checkAll(document.getElementById('frmACH'), 'filecheck', true);"><%=xssStr(allHotWords(617))%></a> | <a href="javascript:checkAll(document.getElementById('frmACH'), 'filecheck', false);"><%=xssStr(allHotWords(618))%> </a>
																			<%end if%>
																		</td>
																	</tr>
																</table>
															</td>
														</tr>
<% end if 'expReport %>
														<%

                                                    	if request.form("optDisMode")="detail" or request.form("optDisMode")="" then

														    strSQL = "SELECT tblCCTrans.SaleID, tblCCTrans.Status, tblCCTrans.Settled, tblCCTrans.ccAmt, tblCCTrans.ClientID, tblCCTrans.TransactionNumber, tblCCTrans.AuthCode, tblCCTrans.TransTime, tblCCTrans.OrderID, tblCCTrans.BatchFileNum, tblCCTrans.ACHAccountNum, tblCCTrans.ACHName, tblCCTrans.ccType, tblCCTrans.CCLastFour, CLIENTS.LastName, CLIENTS.FirstName, tblCCTrans.Cardholder FROM tblCCTrans INNER JOIN CLIENTS ON tblCCTrans.ClientID = CLIENTS.ClientID LEFT OUTER JOIN Sales ON tblCCTrans.SaleID = Sales.SaleID "
														    strSQL = strSQL & "WHERE (tblCCTrans.Settled = 0) AND (tblCCTrans.Status=N'Pending' OR tblCCTrans.Status=N'Sent to Bank' OR tblCCTrans.Status=N'Open') "
														    'CB 47_225 - Support for PAP/DDA ie BANK, VISA & AMEX
														    if request.form("optBatchTransType")="BANK" then
															    strSQL = strSQL & " AND (NOT (tblCCTrans.ACHName IS NULL)) "
														    elseif request.form("optBatchTransType")="VISA" then
															    strSQL = strSQL & "AND (tblCCTrans.ccType=N'Visa' OR tblCCTrans.ccType=N'Master Card') "
														    elseif request.form("optBatchTransType")="AMEX" then
															    strSQL = strSQL & "AND (tblCCTrans.ccType=N'American Express') "
														    end if				
														    if disMode = "range" then
															    strSQL = strSQL & " AND TransTime >= " & DateSep & cSDate & DateSep
															    strSQL = strSQL & " AND TransTime <= " & DateSep & DateAdd("d", 2, cEDate) & DateSep
														    end if
														    if ccLoc<>"-2" then
															    strSQL = strSQL & " AND (tblCCTrans.LocationID = " & ccLoc & ") "
														    end if
														    if request.form("optBatchFileNum")<>"" AND request.form("optBatchFileNum")<>"0" then
															    strSQL = strSQL & " AND (tblCCTrans.BatchFileNum = " & request.form("optBatchFileNum") & ") "
														    end if
														    if request.form("optStatus")<>"" then
															    strSQL = strSQL & " AND (tblCCTrans.Status=N'" & request.form("optStatus") & "') "
														    end if
														    
                                                        	if request.form("frmTagClients")="true" then 'tag clients sql
		                                                        if request.form("frmTagClientsNew")="true" then
			                                                        clearAndTagQuery(strSQL)
		                                                        else
			                                                        tagQuery(strSQL)
		                                                        end if
		                                                    end if

														    strSQL = strSQL & " ORDER BY tblCCTrans.TransTime DESC, TransactionNumber DESC"
                                                            
                                                        else    'summary
                                                        
                                                            strSQL = "SELECT SUM(tblCCTrans.ccAmt) AS TotalAmt, CONVERT(datetime, CONVERT(varchar, CONVERT(DATETIME, tblCCTrans.TransTime, 10), 10) + ' 00:00:00', 10) AS Date, tblCCTrans.ccType, COUNT(*) AS NumTrans FROM tblCCTrans "
                                                            strSQL = strSQL & "WHERE (tblCCTrans.Settled = 0) AND (tblCCTrans.Status = N'Pending' OR tblCCTrans.Status = N'Sent to Bank' OR tblCCTrans.Status=N'Open') "
														    if request.form("optBatchTransType")="BANK" then
															    strSQL = strSQL & " AND (NOT (tblCCTrans.ACHName IS NULL)) "
														    elseif request.form("optBatchTransType")="VISA" then
															    strSQL = strSQL & "AND (tblCCTrans.ccType=N'Visa' OR tblCCTrans.ccType=N'Master Card') "
														    elseif request.form("optBatchTransType")="AMEX" then
															    strSQL = strSQL & "AND (tblCCTrans.ccType=N'American Express') "
														    end if				
														    if disMode = "range" then
															    strSQL = strSQL & " AND TransTime >= " & DateSep & cSDate & DateSep
															    strSQL = strSQL & " AND TransTime <= " & DateSep & DateAdd("d", 2, cEDate) & DateSep
														    end if
														    if ccLoc<>"-2" then
															    strSQL = strSQL & " AND (tblCCTrans.LocationID = " & ccLoc & ") "
														    end if
														    if request.form("optBatchFileNum")<>"" AND request.form("optBatchFileNum")<>"0" then
															    strSQL = strSQL & " AND (tblCCTrans.BatchFileNum = " & request.form("optBatchFileNum") & ") "
														    end if
														    if request.form("optStatus")<>"" then
															    strSQL = strSQL & " AND (tblCCTrans.Status=N'" & request.form("optStatus") & "') "
														    end if
                                                            strSQL = strSQL & "GROUP BY CONVERT(datetime, CONVERT(varchar, CONVERT(DATETIME, tblCCTrans.TransTime, 10), 10) + ' 00:00:00', 10), tblCCTrans.ccType "
                                                            strSQL = strSQL & "ORDER BY Date DESC, tblCCTrans.ccType"
                                                        
                                                        end if  'deatil vs summary

													response.write debugSQL(strSQL, "SQL")
														rsEntry.CursorLocation = 3
														rsEntry.open strSQL, cnWS
														Set rsEntry.ActiveConnection = Nothing
														if NOT rsEntry.EOF then
                                                        	if request.form("optDisMode")="detail" or request.form("optDisMode")="" then
															%>
															<tr>
																<td colspan="2">
																	<table class="mainText" border="0" cellpadding="0" cellspacing="0" width="100%">
																		<% if NOT expReport then %>
																			<tr style="background-color:<%=session("pageColor2")%>;"><td colspan="11"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
																		<% end if %>
																		<tr bgcolor="<%Response.Write session("pageColor4")%>"> 
																			<td colspan="2" class="whiteHeader" nowrap> <b>&nbsp;<%= xssStr(allHotWords(57))%> / <%= xssStr(allHotWords(58))%></b></td>
																			<td class="whiteHeader" nowrap> <b>&nbsp;<%= xssStr(allHotWords(12))%></b></td>
																			<td class="whiteHeader" nowrap> <b>&nbsp;<%= xssStr(allHotWords(62))%></b></td>
																			<td class="whiteHeader center" nowrap><b>&nbsp;<%=xssStr(allHotWords(424)) %>&nbsp;</b></td>
																			<td class="whiteHeader right" nowrap><b>&nbsp;<%= xssStr(allHotWords(35))%>&nbsp;</b></td>
																			<% if NOT isBatchProcessor then %>
																				<td class="whiteHeader center" nowrap><b><%=xssStr(allHotWords(666)) %></b></td>
																			<% end if %>
																			<td class="whiteHeader center" nowrap><b>&nbsp;<%=xssStr(allHotWords(667)) %></b></td>
																			<% if isBatchProcessor then %>
																				<td class="whiteHeader center" nowrap><b><%=xssStr(allHotWords(676)) %></b></td>
																			<% end if %>
																			<td class="whiteHeader center" nowrap><b><%= xssStr(allHotWords(60))%></b></td>
																			<% if isBatchProcessor then %>
																				<td class="whiteHeader center" nowrap><%=xssStr(allHotWords(784)) %></td>
																			<% end if %>
																		</tr>
																		<% if NOT expReport then %>
																			<tr style="background-color:<%=session("pageColor2")%>;"><td colspan="11"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
																		<% end if %>
																		<%
																		rowcount = 0    
																		Do While NOT rsEntry.EOF
																		    tmpCounter = tmpCounter + 1
																		    totalAmt = totalAmt + rsEntry("ccAmt")
																			if rowcount=0 then
																				rowColor = "#F2F2F2"
																				rowcount = 1
																			else
																				rowColor = "#FAFAFA"
																				rowcount = 0
																			end if
																			%>
																			<tr style="background-color:<%=rowColor%>;" height="22">
																				<td nowrap>&nbsp;<%=tmpCounter%>.</td>
																				<td nowrap>&nbsp;<%=FmtDateTime(DateAdd("n", Session("tzOffset"),rsEntry("TransTime")))%></td>
																				<td nowrap>&nbsp;
																					<% if NOT expReport then %>
																						<a href="adm_clt_ph.asp?ID=<%=rsEntry("ClientID")%>&qParam=ph"><%=rsEntry("FirstName")%>&nbsp;<%=rsEntry("LastName")%></a>
																					<% else %>
																						<%=rsEntry("FirstName")%>&nbsp;<%=rsEntry("LastName")%>
																					<% end if %>
																				</td>
																				<td nowrap>&nbsp;
																					<% if NOT isNULL(rsEntry("ACHAccountNum")) then %>
																						<%=rsEntry("ACHName")%>
																					<% else 'CC Trx %>
																						<%=rsEntry("Cardholder")%>&nbsp;
																						<%=rsEntry("ccType")%>&nbsp;/&nbsp;<%=xssStr(allHotWords(670)) %>:&nbsp;
																						<%
																						if NOT isNULL(rsEntry("CCLastFour")) then
																							response.write FmtPadString(rsEntry("CCLastFour"), 4, "0", true)
																						end if
																						%>
																					<% end if %>
																				</td>
																				<td nowrap class="center">&nbsp;
																					<%
																					if NOT isNull(rsEntry("SaleID")) then
																						if NOT expReport then 
																							if VoidOk then
																								response.write "<a title=""" & DisplayPhraseAttr(phraseDictionary, "Clicktogototransaction") & """ href=""adm_tlbx_voidedit.asp?saleno=" & rsEntry("SaleID") & """>[View/Void]</a>"
																							else
																								response.write xssStr(allHotWords(159)) & "/" & xssStr(allHotWords(379))
																							end if
																						else
																							response.write rsEntry("SaleID")
																						end if
																					else
																						response.write "n/a"
																					end if
																					%>
																				</td>
																				<td nowrap class="right">&nbsp;<%=FormatNumber(rsEntry("ccAmt")*.01,2)%>&nbsp;</td>
																				<% if NOT isBatchProcessor then%>
																					<td nowrap class="center">&nbsp;<%=rsEntry("authCode")%>
																						<%if NOT isNull(rsEntry("OrderID")) then%>
																							| <%=rsEntry("OrderID")%>
																						<%end if%>
																					</td>
																				<%end if%>
																				<td nowrap class="center"><%=rsEntry("TransactionNumber")%></td>
																				<%if isBatchProcessor then%>
																					<td nowrap class="center"><%=rsEntry("BatchFileNum")%></td>
																				<%end if%>
																				<td nowrap class="center"><%=rsEntry("Status")%></td>
																				<%if isBatchProcessor then%>
																					<td class=center> 
																						<%if rsEntry("Status")="Pending" then %>
																							<input type="checkbox" name="chk_<%=rsEntry("TransactionNumber")%>" id="chk_<%=rsEntry("TransactionNumber")%>" class="filecheck">
																						<%end if%>
																					</td>
																				<%end if%>
																			</tr>
																			<% if NOT expReport then %>
																				<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="11"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
																			<% end if %>
																			<%
																			if rsEntry("Status")="Pending" then
																			numTrx = numTrx + 1
																			end if
																			rsEntry.MoveNext 
																		Loop
																		%>
																		<tr><td colspan="11"><strong><%=DisplayPhrase(phraseDictionary, "Totalamount") %>: <%=FmtCurrency(totalAmt/100)%></strong></td></tr>
																	</table>
																</td>
															</tr>
															<%
															else    'summary
															%>




															<tr>
																<td colspan="2">
																	<table class="mainText" border="0" cellpadding="0" cellspacing="0" width="100%">
																		<% if NOT expReport then %>
																			<tr style="background-color:<%=session("pageColor2")%>;"><td colspan="11"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
																		<% end if %>
																		<tr bgcolor="<%Response.Write session("pageColor4")%>"> 
																			<td class="whiteHeader" nowrap> <b>&nbsp;<%= xssStr(allHotWords(57))%></b></td>
																			<td class="whiteHeader" nowrap> <b>&nbsp;<%= xssStr(allHotWords(218))%></b></td>
																			<td class="whiteHeader" nowrap> <b>&nbsp;<%=DisplayPhrase(phraseDictionary, "Numbertransactions") %></b></td>
																			<td class="whiteHeader right" nowrap><b>&nbsp;<%= xssStr(allHotWords(35))%>&nbsp;</b></td>
																		</tr>
																		<% if NOT expReport then %>
																			<tr style="background-color:<%=session("pageColor2")%>;"><td colspan="11"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
																		<% end if %>
																		<%
																		rowcount = 0    
																		Do While NOT rsEntry.EOF
																			if rowcount=0 then
																				rowColor = "#F2F2F2"
																				rowcount = 1
																			else
																				rowColor = "#FAFAFA"
																				rowcount = 0
																			end if
																			%>
																			<tr style="background-color:<%=rowColor%>;">
																				<td nowrap>&nbsp;<a href="javascript:gotoDetail('<%=FmtDateShort(rsEntry("Date"))%>','<%=rsEntry("ccType")%>');"><%=FmtDateShort(rsEntry("Date"))%></a></td>
																				<td nowrap>&nbsp;
																				<% if isNull(rsEntry("ccType")) then 
																				        response.Write ACHHotWord
																				   elseif rsEntry("ccType")="Visa" OR rsEntry("ccType")="Master Card" then
																				        response.Write xssStr(allHotWords(660))
																				   elseif rsEntry("ccType")="American Express" then
																				        response.Write xssStr(allHotWords(659))
																				   else
																				        response.Write xssStr(allHotWords(661))
																				   end if%></td>
																				<td nowrap>&nbsp;<%=rsEntry("NumTrans")%></td>
																				<td nowrap class="right">&nbsp;<%=FmtCurrency(rsEntry("TotalAmt")*.01)%>&nbsp;</td>
																			</tr>
																			<% if NOT expReport then %>
																				<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="11"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
																			<% end if %>
																			<%
																			numTrx = numTrx + rsEntry("NumTrans")
																			rsEntry.MoveNext 
																		Loop
																		%>
																	</table>
																</td>
															</tr>



															
															
															
														<%
														    end if  'detail vs summary
														else	'EOF
															%>
															<% if NOT expReport then %>
																<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="11"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
															<% end if %>
															<tr height="30"><td colspan="11" class="center"><span style="color:#990000"><%=DisplayPhrase(phraseDictionary, "Nopendingtransactions") %></span></td></tr>
															<%
														end if
														rsEntry.Close
														Set rsEntry = Nothing
														%>
														<% if NOT expReport then %>
															<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="11"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td></tr>
														<% end if %>
													</table>
<% if NOT expReport then %>
												</td>
											</tr>
											<tr>
												<td class="center">
													<br /><br />
													<%if isBatchProcessor AND numTrx>0 AND (numFTPAccounts=1 OR ccLoc<>"-2") then %>
															<strong>Payment Date:</strong> 
															<input type="text"  name="requiredtxtPaymentDate" value="<%=FmtDateShort(cEDate)%>" class="date">
															<script type="text/javascript">
																var cal3 = new tcal({'formname':'frmACH', 'controlname':'requiredtxtPaymentDate'});
																cal3.a_tpl.yearscroll = true;
															</script>
															&nbsp;&nbsp;&nbsp;
                                                        <input onClick="SendToBank();" type="button" name="sendBtn" id="sendBtn" value="<%=DisplayPhraseAttr(phraseDictionary, "Sendselectedtransactions") %>" <%if ccProcessor="HSBC" AND request.form("optBatchTransType")="" then response.write "disabled" end if%>>
														<%if ccProcessor="HSBC" AND request.form("optBatchTransType")="" then%>
															<script type="text/javascript">
																document.getElementById("sendBtn").value = "<%=DisplayPhraseJS(phraseDictionary, "Selecttransactiontype") %>";
															</script>
														<%end if%>
														
													
														
													<%end if%>

													<%if ccProcessor="HSBC" AND request.form("optBatchFileNum")<>"" AND request.form("optBatchFileNum")<>"0" then %>
    													<input onClick="resetBatch();" type="button" name="ResetBatchBtn" id="clearBatchBtn" value="<%=DisplayPhraseAttr(phraseDictionary, "Resetbatchtopending") %>">
													<%end if %>
													
                                                    <% if implementationSwitchIsEnabled("BluefinCanada") then
                                                        Dim locstr
													    if ccProcessor="PMN" OR  ccProcessor="TCI" then
                                                            if ccProcessor2 <> "ELV" then 
															    locstr = "http"
															    if Request.ServerVariables("HTTPS")="on" then
																    locstr = locstr & "s"
															    end if
															    locstr = locstr & "://" & Request.ServerVariables("server_name") & "/ASP/adm/adm_chk_ach.asp"
													    %>
													    <script type="text/javascript">
													        $('#checkACH').live('click', function() {
													            $(this).prop('disabled', true);
													            $('#checkACHspan').html('<b> <%=DisplayPhraseJS(phraseDictionary, "Checkingstatuses") %> </b>').css('color', 'red');
													            //console.log("<%=locstr%>");
													            $.post("<%=locstr%>", function(data) {
													                $('#checkACHspan').html('<b> <%=DisplayPhraseJS(phraseDictionary, "Checkcomplete") %> </b>').css('color', 'green');
													            });
													        });
													    </script>
    													    <span id="checkACHspan"><input type="button" name="checkACH" id="checkACH" value="<%=DisplayPhraseAttr(phraseDictionary, "Refreshstatuses") %>"></span>
                                                            <%end if 
													    end if
                                                    else 
                                                        if ccProcessor="PMN" OR  ccProcessor="TCI" then                                                        
															locstr = "http"
															if Request.ServerVariables("HTTPS")="on" then
																locstr = locstr & "s"
															end if
															locstr = locstr & "://" & Request.ServerVariables("server_name") & "/ASP/adm/adm_chk_ach.asp"
													    %>
													    <script type="text/javascript">
													        $('#checkACH').live('click', function() {
													            $(this).prop('disabled', true);
													            $('#checkACHspan').html('<b> <%=DisplayPhraseJS(phraseDictionary, "Checkingstatuses") %> </b>').css('color', 'red');
													            //console.log("<%=locstr%>");
													            $.post("<%=locstr%>", function(data) {
													                $('#checkACHspan').html('<b> <%=DisplayPhraseJS(phraseDictionary, "Checkcomplete") %> </b>').css('color', 'green');
													            });
													        });
													    </script>
    													    <span id="Span1"><input type="button" name="checkACH" id="Button1" value="<%=DisplayPhraseAttr(phraseDictionary, "Refreshstatuses") %>"></span>                                                        
													    <%end if
													end if %>
													
													<%if session("admin")="sa" AND isBatchProcessor then %>
														<br /><br />
														<!--<a href="tmp_get_pap_results.asp">Download File & Process Results</a>-->
													<%end if%>
												</td>
											</tr>
										</table>
									</td>
								</tr>
							
						</table>
					</td>
				</tr>
				
				</table>
				</form>
<% pageEnd %>
<!-- #include file="post.asp" -->

		<%
else 'expReport
	Dim stFilename
'	stFilename="attachment; filename=PendingTransactions " & Replace(cSDate,"/","-") & " - " & Replace(cEDate,"/","-") & ".xls" 
	stFilename="attachment; filename=PendingTransactions.xls" 
	Response.ContentType = "application/vnd.ms-excel" 
	Response.AddHeader "Content-Disposition", stFilename 

	end if

end if
%>
