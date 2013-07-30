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
		alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
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

		dim category : category = ""
		dim tmpCat : tmpCat = ""
		if (RQ("category"))<>"" then
			tmpCat = RQ("category")
			category = Replace(tmpCat, " ", "")
		elseif (RF("category"))<>"" then
			tmpCat = RF("category") 
			category = Replace(tmpCat, " ", "")
		end if

		dim masterDictionary
		set masterDictionary = LoadPhrases("ReportmasterPage", 82)

		dim phraseDictionary
		set phraseDictionary = LoadPhrases("BusinessmodevoidedrejectedtransactionsPage", 139)

		Dim cSDate, cEDate, disMode, ss_EnableACH, VoidOk, ap_cc_settle, ACHHotWord
		dim rsEntry, ccProcessor
		set rsEntry = Server.CreateObject("ADODB.Recordset")

		ccLoc = "0"
		ss_EnableACH = checkStudioSetting("tblCCOpts", "EnableACH")
		VoidOk = validAccessPriv("TB_VOID")
		ap_cc_settle = validAccessPriv("CC_Settle")

		ACHHotWord = "ACH"
		ACHHotWord = getHotWord(109)
		

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
		
		if request.Form("optCCLoc")<>"" then
			ccLoc = sqlInjectStr(request.Form("optCCLoc"))
		end if
		
		if request.form("frmGenReport")="true" or request.querystring("saleID")<>"" then
			genReport = true
		end if
		
		if NOT request.form("frmExpReport")="true" then

	%>

			
<!-- #include file="pre.asp" -->
	<!-- #include file="frame_bottom.asp" -->



<%= js(array("mb", "adm/adm_rpt_ccp_rej", "MBS", "reportFavorites", "plugins/jquery.SimpleLightBox" )) %>
	<!-- American/Canada=2 format mm/dd/yyyy --> <!-- European/Rest of the world=1 format dd-mm-yyyy -->
	
<%= js(array("calendar" & dateFormatCode)) %>
<%= css(array("SimpleLightBox")) %> 
	<!-- #include file="../inc_date_ctrl.asp" -->
	<!-- #include file="../inc_ajax.asp" -->
	<!-- #include file="../inc_val_date.asp" -->
	<!-- #include file="inc_user_options.asp" -->  
	<!-- #include file="css/site_setup.asp" -->

<% pageStart %>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %> 
	<div class="headText breadcrumbs-old" valign="bottom">
	<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(pageTitlesDictionary, "Reports")%></a></span>
	<span class="breadcrumb-item">&raquo;</span>
	<%if category <> "" then%>
	<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
	<span class="breadcrumb-item">&raquo;</span>
	<%end if %>
	<%= DisplayPhrase(reportPageTitlesDictionary,"Voidedrejectedtransactions") %>

	<div id="add-to-favorites">
		<span id="" class="favorite"></span><span class="text"><%=DisplayPhrase(reportPageTitlesDictionary, "Addtofavorites")%></span>
	</div>
	</div>
<%else %>
	<div class="headText" valign="bottom"><b><%= DisplayPhrase(pageTitlesDictionary,"Voidedrejectedtransactions") %></b></div>
<%end if %>
	<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
	<tr> 
		 <td valign="top" height="100%" width="100%">
			<table id="formTable" class="center" cellspacing="0" width="90%" height="100%">
						  <form name="frmCCP" action="adm_rpt_ccp_rej.asp" method="POST">
						  <% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %> 
								<% if category <> "" then %>
								   <input type="hidden" name="category" id="category" value="<%=category %>" />
								<%end if %>
								<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" /> 
							<%end if %>
									<input type="hidden" name="frmExpReport" value="">
									<input type="hidden" name="frmGenReport" value="">
			  <tr> 
				  <td valign="bottom" class="mainText right" height="18"> 
					<!-- #include file="inc_batch_nav.asp" -->
				  </td>
			  </tr>
							</div>
							<tr>
							<td>
								<table class="mainText border4" cellspacing="0" style="float:right;">
							<tr> 
							  <td class="center-ch nowrap" valign="bottom" style="background-color:#F2F2F2;">
							  <b>
							  <!--
							  <span style="color:<%=session("pageColor4")%>;">&nbsp;<%=xssStr(allHotWords(245)) %></span> 
								<input onClick="document.frmCCP.submit();" class="textSmall" type="radio" name="optDate" value="all" <%if disMode="all" then response.write "checked" end if%>>
								<%=xssStr(allHotWords(149))%>&nbsp;&nbsp; 
								<input onClick="document.frmCCP.submit();" class="textSmall" type="radio" name="optDate" value="range" <%if disMode="range" then response.write "checked" end if%>>
							  -->
								&nbsp;
								<%=xssStr(allHotWords(77))%> 
								<input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" class="transForm" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
								  <script type="text/javascript">
									var cal1 = new tcal({ 'formname': 'frmCCP', 'controlname': 'requiredtxtDateStart' });
									cal1.a_tpl.yearscroll = true;			    
								  </script>
								&nbsp;<%=xssStr(allHotWords(79))%> 
								<input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" class="transForm" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
								 <script type="text/javascript">
									var cal2 = new tcal({ 'formname': 'frmCCP', 'controlname': 'requiredtxtDateEnd' });
									cal2.a_tpl.yearscroll = true;
								 </script>
								&nbsp;
								<% if validAccessPriv("RPT_TAG") then 
									taggingFilter 
								end if %> 
								<input class="textSmall" name="Button" type="button" value= <%=xssStr(allHotWords(226)) %> onClick="genCCP();">
								<span class="icon-button" style="vertical-align: middle;" title= "<%=DisplayPhraseAttr(masterDictionary,"Exporttoexcel")%>" <a onClick="exportCCP();" ><img src="<%= contentUrl("/asp/adm/images/export-to-excel-20px.png") %>" /></a></span>
								</b>
								<% if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then 
								else 
									taggingButtons("frmCCP") 
								end if %>
								</td>
							</tr>
							  </table>
							</td>
							</tr>
			  <tr> 
				  <td valign="bottom" class="textSmall left"> <br />
					<!--CB 3/2/09 removed comment, not true for non-real time integrations
					The below transactions were authorized however either did not 
					meet the CVV2 and/or AVS requirement or were manually voided.--></td>
			  </tr>
			  <tr> 
				
			<td valign="top" class="mainTextBig" height="100%"> 
			  <table class="mainText center" width="95%" cellspacing="0" height="100%">
				<tr > 
					  <td  colspan="2" valign="top" class="mainTextBig">
						  <table class="mainText center" cellspacing="0" width="90%">
						  <tr> 
						<td valign="top"> 
							<table class="mainText center" cellspacing="0" width="90%">
				  <tr> 
								<td class=left valign="top"> 
						<select class="transForm" name="optCCLoc" onChange="document.frmCCP.submit();">
							<option value="" <%if request.form("optCCLoc")="" then response.write "selected" end if%>><%=xssStr(allHotWords(479)) %></option>

	<%
				strSQL = "SELECT MID, LocationName, LocationID, Active FROM Location WHERE (NOT (MID IS NULL)) AND (Active = 1) ORDER BY LocationName"
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing
				do while NOT rsEntry.EOF
	%>
							<option value="<%=rsEntry("LocationID")%>" <%if request.form("optCCLoc")=CSTR(rsEntry("LocationID")) then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
	<%
					rsEntry.MoveNext
				loop
				rsEntry.close
	%>
						  </select>&nbsp;
	<%
				if implementationSwitchIsEnabled("BluefinCanada") then
					strSQL = "SELECT tblCCOpts.EnableACH, tblCCOpts.ccVisa, tblCCOpts.ccMasterCard, tblCCOpts.ccAmericanExpress, tblCCOpts.ccDiscover, Studios.ccProcessor "
					strSQL = strSQL & "FROM tblCCOpts, Studios WHERE tblCCOpts.StudioID=" & session("StudioID")
				else
					strSQL = "SELECT tblCCOpts.EnableACH, tblCCOpts.ccVisa, tblCCOpts.ccMasterCard, tblCCOpts.ccAmericanExpress, tblCCOpts.ccDiscover, tblCCOpts.ccProcessor "
					strSQL = strSQL & "FROM tblCCOpts WHERE tblCCOpts.StudioID=" & session("StudioID")
				end if
				rsEntry.CursorLocation = 3
				rsEntry.open strSQL, cnWS
				Set rsEntry.ActiveConnection = Nothing
				if NOT rsEntry.EOF then
					ss_EnableACH = rsEntry("EnableACH")
					tmpAmex = rsEntry("ccAmericanExpress")
					tmpVisa = rsEntry("ccVisa")
					tmpMC = rsEntry("ccMasterCard")
					tmpDisc = rsEntry("ccDiscover")
					ccProcessor = rsEntry("ccProcessor")
				end if
				rsEntry.close
	%>
									<select class="transForm" name="optCCType" onChange="document.frmCCP.submit();">
									  <option value="-1"> <%if ss_EnableACH then %><%=DisplayPhrase(phraseDictionary,"Allcreditcardsandach")%> <%else%><%=xssStr(allHotWords(680))%><%end if%></option>
	<%
				if tmpVisa or tmpMC then
	%>
									  <option value="4" <%if request.form("optCCType")="4" then response.write "selected" end if%>><%=xssStr(allHotWords(660)) %></option>
	<% 
				end if
				if tmpAmex then
	%>
									  <option value="3" <%if request.form("optCCType")="3" then response.write "selected" end if%> ><%=xssStr(allHotWords(659)) %></option>
	<%
				end if
				if tmpDisc then
	%>
									  <option value="5" <%if request.form("optCCType")="5" then response.write "selected" end if%> ><%=xssStr(allHotWords(661)) %></option>
	<%
				end if
	%>
								<%if ss_EnableACH then %>
									  <option value="100" <%if request.form("optCCType")="100" then response.write "selected" end if%> ><%=xssStr(allHotWords(681))%></option>
									  <option value="101" <%if request.form("optCCType")="101" then response.write "selected" end if%> ><%=DisplayPhrase(phraseDictionary,"Creditcardsonly") %></option>
								<%end if%>
									</select>
								</td>
			<% else ' we are exporting %>
<%
			Dim stFilename
			stFilename="attachment; filename=Voided/RejectedTransactions " & Replace(cSDate,"/","-") & " - " & Replace(cEDate,"/","-") & ".xls" 
			Response.ContentType = "application/vnd.ms-excel" 
			Response.AddHeader "Content-Disposition", stFilename 

%>


<%			end if 'export check %>
								<td class=right>
	<%
			 Dim strTempName   
			 Dim intCount
	if genReport then
			
		if request.form("frmTagClients")="true" then 'tag clients sql
				strSQL = "SELECT tblCCTrans.ClientID "
				strSQL = strSQL & "FROM tblCCTrans INNER JOIN CLIENTS ON tblCCTrans.ClientID = CLIENTS.ClientID LEFT OUTER JOIN Sales ON tblCCTrans.SaleID = Sales.SaleID "
				if request.form("optFilterTagged")="on" then '55_3198, CCP 10/9/09, Tagged Clients Only
					strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
				end if
				strSQL = strSQL & "WHERE tblCCTrans.ccAmt IS NOT NULL AND ((tblCCTrans.Status)<>'Approved') AND ((tblCCTrans.Status)<>'Pending') AND ((tblCCTrans.Status)<>'Sent to Bank') AND ((tblCCTrans.Settled)=0) "
				if ccProcessor="TCI" then
							strSQL = strSQL & " AND ((tblCCTrans.Status)<>'Credit') "
						end if
				if request.form("optFilterTagged")="on" then '55_3198, CCP 10/9/09, 
					if session("mvaruserID")<>"" then
						strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
					else
						strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
					end if
				end if
				if disMode = "range" then
					strSQL = strSQL & " AND TransTime >= " & DateSep & cSDate & DateSep
					strSQL = strSQL & " AND TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep
				end if
				if request.form("optCCLoc")<>"" then
					strSQL = strSQL & " AND tblCCTrans.LocationID=" & ccLoc
				end if
				if request.form("optCCType")<>"" AND request.form("optCCType")<>"-1" then
					if request.form("optCCType")="4" then
						strSQL = strSQL & " AND ([Sales].PaymentMethod=4 OR [Sales].PaymentMethodB=4 OR tblCCTrans.ccType = 'Visa' OR tblCCTrans.ccType = 'Master Card') "
					elseif request.form("optCCType")="5" then
						strSQL = strSQL & " AND ([Sales].PaymentMethod=5 OR [Sales].PaymentMethodB=5 OR tblCCTrans.ccType = 'Discover') "
					elseif request.form("optCCType")="3" then
						strSQL = strSQL & " AND ([Sales].PaymentMethod=3 OR [Sales].PaymentMethodB=3 OR tblCCTrans.ccType = 'American Express') "
					elseif request.form("optCCType")="6" then
						strSQL = strSQL & " AND ([Sales].PaymentMethod=6 OR [Sales].PaymentMethodB=6) "
					elseif request.form("optCCType")="100" then
						strSQL = strSQL & " AND (NOT (tblCCTrans.ACHName IS NULL)) "
					elseif request.form("optCCType")="101" then
						strSQL = strSQL & " AND (tblCCTrans.ACHName IS NULL) "
					else
						strSQL = strSQL & " AND ([Sales].PaymentMethod=" & request.form("optCCType") & " OR [Sales].PaymentMethodB=" & request.form("optCCType") & ") "
					end if
				end if
				
			   response.write debugSQL(strSQL, "SQL")
				'response.end
							
				if request.form("frmTagClientsNew")="true" then
					clearAndTagQuery(strSQL)
				else
					tagQuery(strSQL)
				end if
							
			end if
			 
			strSQL = "SELECT tblCCTrans.SaleID, tblCCTrans.ACHName, tblCCTrans.Status, tblCCTrans.Settled, tblCCTrans.ccAmt, tblCCTrans.ClientID, tblCCTrans.TransactionNumber, tblCCTrans.AuthCode, tblCCTrans.TransTime, tblCCTrans.OrderID, tblCCTrans.CCLastFour, CLIENTS.LastName, CLIENTS.FirstName "
			strSQL = strSQL & "FROM tblCCTrans INNER JOIN CLIENTS ON tblCCTrans.ClientID = CLIENTS.ClientID LEFT OUTER JOIN Sales ON tblCCTrans.SaleID = Sales.SaleID "
			if request.form("optFilterTagged")="on" then '55_3198, CCP 10/9/09, Tagged Clients Only
				strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
			end if
			strSQL = strSQL & "WHERE tblCCTrans.ccAmt IS NOT NULL AND ((tblCCTrans.Status)<>'Approved') AND ((tblCCTrans.Status)<>'Pending') AND ((tblCCTrans.Status)<>'Sent to Bank') AND ((tblCCTrans.Settled)=0) "
			if ccProcessor="TCI" then
					strSQL = strSQL & " AND ((tblCCTrans.Status)<>'Credit') "
			end if
			if request.form("optFilterTagged")="on" then '55_3198, CCP 10/9/09, 
				if session("mvaruserID")<>"" then
					strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
				else
					strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
				end if
			end if
			if disMode = "range" then
				strSQL = strSQL & " AND TransTime >= " & DateSep & cSDate & DateSep
				strSQL = strSQL & " AND TransTime <= " & DateSep & DateAdd("d", 1, cEDate) & DateSep
			end if
			if request.form("optCCLoc")<>"" then
				strSQL = strSQL & " AND tblCCTrans.LocationID=" & ccLoc
			end if
			if request.form("optCCType")<>"" AND request.form("optCCType")<>"-1" then
				if request.form("optCCType")="4" then
					strSQL = strSQL & " AND ([Sales].PaymentMethod=4 OR [Sales].PaymentMethodB=4 OR tblCCTrans.ccType = 'Visa' OR tblCCTrans.ccType = 'Master Card') "
				elseif request.form("optCCType")="5" then
					strSQL = strSQL & " AND ([Sales].PaymentMethod=5 OR [Sales].PaymentMethodB=5 OR tblCCTrans.ccType = 'Discover') "
				elseif request.form("optCCType")="3" then
					strSQL = strSQL & " AND ([Sales].PaymentMethod=3 OR [Sales].PaymentMethodB=3 OR tblCCTrans.ccType = 'American Express') "
				elseif request.form("optCCType")="6" then
					strSQL = strSQL & " AND ([Sales].PaymentMethod=6 OR [Sales].PaymentMethodB=6) "
				elseif request.form("optCCType")="100" then
					strSQL = strSQL & " AND (NOT (tblCCTrans.ACHName IS NULL)) "
				elseif request.form("optCCType")="101" then
					strSQL = strSQL & " AND (tblCCTrans.ACHName IS NULL) "
				else
					strSQL = strSQL & " AND ([Sales].PaymentMethod=" & request.form("optCCType") & " OR [Sales].PaymentMethodB=" & request.form("optCCType") & ") "
				end if
			end if

			'strSQL = strSQL & " AND tblCCTrans.MerchantID=N'" & pMID & "'"
			strSQL = strSQL & " ORDER BY tblCCTrans.TransTime DESC;"
		   response.write debugSQL(strSQL, "SQL1")
			'response.end
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
	%>
						<!--
						<a href="javascript:checkAll(document.getElementById('frmCCP'), 'filecheck', true);">Check 
						All</a> | <a href="javascript:checkAll(document.getElementById('frmCCP'), 'filecheck', false);">Uncheck 
						All </a> 
						-->
						</td></tr><tr><td colspan="2">
						<table class="mainText" cellspacing="0" width="100%">
						<% if NOT request.form("frmExpReport")="true" then  %>
							<tr style="background-color:<%=session("pageColor2")%>;"> 
							  <td colspan="10"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
							</tr>
						<%end if %>
						<tr bgcolor="<%= session("pageColor4")%>"> 
						  <td class="whiteHeader"> <b>&nbsp;<%= xssStr(allHotWords(57))%> / <%=xssStr(allHotWords(58))%></b></td>
						  <td class="whiteHeader"> <b>&nbsp;<%= xssStr(allHotWords(12))%></b></td>
						  <td class="whiteHeader"> <b>&nbsp;<%= xssStr(allHotWords(44))%></b></td>
						  <td  nowrap class="whiteHeader center-ch"><b>&nbsp;<%= xssStr(allHotWords(35))%>&nbsp;</b></td>
						  <td  nowrap class="whiteHeader center-ch"><b>&nbsp;<%= xssStr(allHotWords(115))%>&nbsp;</b></td>
						  <td  nowrap class="whiteHeader center-ch"><b>&nbsp;<%=xssStr(allHotWords(666)) %></b></td>
						  <td  nowrap class="whiteHeader center-ch"><b>&nbsp;<%= xssStr(allHotWords(667)) %></b></td>
						  <td  nowrap class="whiteHeader center-ch"><b><%= xssStr(allHotWords(60))%></b></td>
						  <td  nowrap class="whiteHeader center-ch"><!--<b>Delete</b>--></td>
						</tr>
						<% if NOT request.form("frmExpReport")="true" then  %>
						<tr style="background-color:<%=session("pageColor2")%>;"> 
						  <td colspan="10"><img height="2" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
						</tr>
						<% end if %>
						<%

		rowcount = 0    
		if NOT rsEntry.EOF then
			Do While NOT rsEntry.EOF


					if rowcount=0 then
	%>
						<tr bgcolor=#F2F2F2> 
	<%
				   rowcount = 1
				else
	%>
						<tr bgcolor=#FAFAFA> 
	<%
				   rowcount = 0
				end if
	%>
						  <td>&nbsp;<%=FmtDateTime(DateAdd("n", Session("tzOffset"),rsEntry("TransTime")))%></td>
						  <td>&nbsp;<a href="adm_clt_purch.asp?ID=<%=rsEntry("ClientID")%>&qParam=ph"><%=rsEntry("FirstName")%>&nbsp;<%=rsEntry("LastName")%></a>
						  <%if NOT isNULL(rsEntry("CCLastFour")) then%>
								&nbsp;(Last4: <%=FmtPadString(rsEntry("CCLastFour"), 4, "0", true)%>)
						  <%end if%>
						  </td>
						  <td><%if NOT isNULL(rsEntry("ACHName")) then response.write rsEntry("ACHName") end if%></td>
						  <td nowrap class="center-ch">&nbsp;<%=FormatNumber(rsEntry("ccAmt")*.01,2)%></td>
						  <td nowrap class="center-ch">&nbsp;
	<%
							if NOT isNull(rsEntry("SaleID")) then
								if VoidOk then
									response.write "<a title=""Click to go to Transaction"" href=""adm_tlbx_voidedit.asp?saleno=" & rsEntry("SaleID") & """>" & rsEntry("SaleID") & "</a>"
								else
									response.write rsEntry("SaleID")
								end if
							else
								response.write "n/a"
							end if
	%>
						  </td>
						  <td nowrap class="center-ch">&nbsp;<%=rsEntry("authCode")%>
										<%if NOT isNull(rsEntry("OrderID")) then%>
											| <%=rsEntry("OrderID")%>
										<%end if%>
						  </td>
						  <td nowrap class="center-ch"><%=rsEntry("TransactionNumber")%></td>
						  <td nowrap class="center-ch"><%=rsEntry("Status")%></td>
						  <td class="center"> 
							<!--<input type="checkbox" name="chk_<%=rsEntry("TransactionNumber")%>"  class="filecheck">-->
						  </td>
						</tr>
						<tr style="background-color:<%=session("pageColor4")%>;"> 
						  <td colspan="10"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
						</tr>
						<%
			rsEntry.MoveNext 
		
			Loop
		else
	%>
						<tr> 
						  <td colspan="10"> <%=DisplayPhrase(phraseDictionary,"Norejectedtransactions") %></td>
						</tr>
	<%	
		end if
		rsEntry.Close
		Set rsEntry = Nothing
	%>
						<% if NOT request.form("frmExpReport")="true" then  %>
							<tr style="background-color:<%=session("pageColor4")%>;"> 
							  <td colspan="10"><img height="1" width="100%" src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
							</tr>
						<%end if %>
					  </table>
	<%end if	'frmGenReport%>	
	<% if NOT request.form("frmExpReport")="true" then  %>
					</TD>
				  </TR>
				</TABLE>
				<br />
							<br />
					  </td>
					</tr>
				  </table>
				</td>
			  </tr>
			  
	</form>
	
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
	end if
	

end if
%>
