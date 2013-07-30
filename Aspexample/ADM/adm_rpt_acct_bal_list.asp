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
	dim rsUnpaidRem
	set rsUnpaidRem = Server.CreateObject("ADODB.Recordset")
	%>
	<!-- #include file="inc_accpriv.asp" -->
	<!-- #include file="inc_rpt_tagging.asp" -->
	<!-- #include file="inc_utilities.asp" -->
	<!-- #include file="inc_rpt_save.asp" -->
	<!-- #include file="inc_hotword.asp" -->
	<!-- #include file="../inc_post.asp" -->
	<%
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_ACCTBAL") then 
	%>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
	<%
	else
	%>
			<!-- #include file="../inc_i18n.asp" -->
		<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

		Dim cSDate, cEDate, tmpCurDate, rsClientID, rsRssid, rsBalance, rsUnpaid, balanceTotal, unpaidTotal, cStDate, pShow, pBilling, cBillingDate, pCltPhone
		Dim tmpBal30, tmpBal60, tmpBal90, tmpBal90Plus, totalBal30, totalBal60, totalBal90, totalBal90Plus, ShowAutoPaySchedule, is3rdParty
		
		Dim ap_RPT_SCH_BAL_AUTOPAY : ap_RPT_SCH_BAL_AUTOPAY = validAccessPriv("RPT_SCH_BAL_AUTOPAY")
		Dim ss_EnableACH : ss_EnableACH = checkStudioSetting("tblCCOpts", "EnableACH")
		Dim ss_UseEFT : ss_UseEFT = checkStudioSetting("Studios", "UseEFT")
		Dim ss_Use3rdParty : ss_Use3rdParty = checkStudioSetting("tblGenOpts", "Enable3rdPartyPayers")
		Dim ss_ApplyAccountPaymentsByLocation : ss_ApplyAccountPaymentsByLocation = checkStudioSetting("tblGenOpts", "ApplyAccountPaymentsByLocation")
				
		dim category : category = ""
		if (RQ("category"))<>"" then
			category = RQ("category")
		elseif (RF("category"))<>"" then
			category = RF("category") 
		end if

		if request.form("requiredtxtDate")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cSDate = CDATE(request.form("requiredtxtDate"))
			Call SetLocale("en-us")
		else
			cSDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		end if
		
		if request.form("requiredtxtDateStatement")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cStDate = CDATE(request.form("requiredtxtDateStatement"))
			Call SetLocale("en-us")
		else
			cStDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		end if
		
		if request.Form("optAutopayScheduleDate") <> "" then
		    Call SetLocale(session("mvarLocaleStr"))
				cBillingDate = CDATE(request.form("optAutopayScheduleDate"))
			Call SetLocale("en-us")
		else
			cBillingDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		end if

		pShow = request.form("chkNegOnly")
		pBilling = request.form("optBillingInfo")
		if request.form("frmSchAutoPay")="true" then
			pShow = "-1"
			pBilling = "with"
		end if
		ShowAutoPaySchedule = false
		'has permission, has integrated ccp, report options set correctly
		if ap_RPT_SCH_BAL_AUTOPAY AND Session("mvarMIDs")<>"0" AND ss_UseEFT AND pShow="-1" AND pBilling="with" then
			ShowAutoPaySchedule = true
		end if
		
		Dim ap_view_all_locs : ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
		Dim optSaleLocation : optSaleLocation = 0
		Dim tblClientAccountStr : tblClientAccountStr = "tblClientAccount"
		if isNumeric(request.form("optSaleLocation")) then
			optSaleLocation = CInt(request.Form("optSaleLocation"))
			if NOT optSaleLocation=0 then
				dim jsonParams : set jsonParams = JSON.parse("{}")
				jsonParams.set "LocationID",optSaleLocation
				CallMethodWithJSON "mb.Core.BLL.SitesPoco.ClientAccount.GetPerLocationView",jsonParams
'				dim clientAccountBLL : set clientAccountBLL = Server.CreateObject("mb.Core.BLL.ClientAccountBLLCOM")	
'				tblClientAccountStr = clientAccountBLL.GetPerLocationView(optSaleLocation)
			end if

		end if
		
		if NOT request.form("frmExpReport")="true" then
		%>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			<%= css(array("calendar", "SimpleLightBox")) %>
<%= js(array("mb", "MBS", "valcur", "calendar" & dateFormatCode, "reportFavorites", "plugins/jquery.SimpleLightBox" )) %>
			
			<%= js(array("adm/adm_rpt_acct_bal_list")) %>
			<script type="text/javascript">
			function updateHidden(what) {
				
				var id = what.id.replace("required","hidden");
				var val = document.getElementById(what.id).value;
				$('#'+id).val(document.getElementById(what.id).value);
				$.get("/i18n/ToUSCents", { val: val},
				 function(data){
					 $('#'+id).val(data);
				});
			}
			
			
			function exportReport() {
				document.frmParameter.frmExpReport.value = "true";
				document.frmParameter.frmGenReport.value = "true";
				<% iframeSubmit "frmParameter", "adm_rpt_acct_bal_list.asp" %>
			}
			function createInvoice(clientID)
			{   
			    <%'JM-55_3182
			    if request.form("optEventBal")="on" then%>
				    document.frmParameter.action = "adm_rpt_acct_bal_statmnt.asp?ClientID=" + clientID + "&optEventBal=on";
				<%else %>
				    document.frmParameter.action = "adm_rpt_acct_bal_statmnt.asp?ClientID=" + clientID;
				<%end if %>
				document.frmParameter.frmGenReport.value = "true";
				document.frmParameter.frmExpReport.value = "false";
				document.frmParameter.submit();
			}
			function GenerateAutoPaySchedules() {
				var total, count;
				total = 0;
				count = 0;
				for (i=0,n=document.frmParameter.elements.length;i<n;i++) {
					if (document.frmParameter.elements[i].className.indexOf('filecheck') !=-1) {
						if (document.frmParameter.elements[i].checked) {
							count++;
							total += parseFloat(document.getElementById("hiddentxtAmount" + document.frmParameter.elements[i].name.substr(10)).value);
							//console.log("total: "+ total);
						}
					}
				}
				
				if (count > 0) {
					$.get("/i18n/FmtCurrencyFromCents", { num: total },
						function(data) {
							if (confirm("You are about to schedule " + count + " AutoPay transactions totaling " + data + ".\n\n" + "Are you sure you want to do this? ") ) {
								document.frmParameter.RunBtn.disabled = true;
								document.frmParameter.frmGenAutoPaySch.value = "true";
								genReport();
							}
					});
				} else {
					alert("Please select one or more <%=LCASE(session("ClientHW"))%>'s to schedule AutoPays.");
				}
			}			

			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
			<!-- #include file="inc_help_content.asp" -->
			<!-- #include file="../inc_ajax.asp" -->
			<!-- #include file="../inc_val_date.asp" -->
		<%
		end if
		%>
		
		
		<% if NOT request.form("frmExpReport")="true" then %>
<% pageStart %>
			<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div class="headText breadcrumbs-old">
			<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
			<%if category <>"" then%>
			<span class="breadcrumb-item">&raquo;</span>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<%end if %>
			<span class="breadcrumb-item">&raquo;</span>
			<%= DisplayPhrase(reportPageTitlesDictionary, "Accountbalances") %> as of <%=FmtDateShort(cStDate)%>
			<% showNewHelpContentIcon("account-balances-report") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
			</div>
			<%end if %>
			<table height="100%" width="100%" cellspacing="0">    
				<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<tr> 
				<td id="accountBalancesHeader" class="headText" align="left" valign="top">
				<table width="100%" cellspacing="0">
					<tr>
					<td class="headText" valign="bottom" height="30"><b> <%= pp_PageTitle("Account Balances") %> as of <%=FmtDateShort(cStDate)%></b>
					<!--JM - 49_2447-->
					<% showNewHelpContentIcon("account-balances-report") %>

					</td>
					</tr>
				</table>
				</td>
				</tr>
				<%end if %>
				<tr> 
				<td valign="top" height="100%" width="100%"> 
				<table class="center" cellspacing="0" width="90%" height="100%">
					<form name="frmParameter" id="frmParameter" action="adm_rpt_acct_bal_list.asp" method="post">
					<input type="hidden" id="alertTotal" value="" />
					<tr> 
					<td height="30"  valign="bottom" class="headText">
					
					  <table class="mainText border4 center" cellspacing="0">
						<input type="hidden" name="frmGenReport" value="">
						<input type="hidden" name="frmExpReport" value="">
						<input type="hidden" name="chkPrinterFriendly" value="">
						<input type="hidden" name="frmSchAutoPay" value="">
						<input type="hidden" name="frmGenAutoPaySch" value="">
						<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
							<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
							<input type="hidden" name="category" value="<%=category%>">
						<% end if %>

						<tr> 
						<td height="30" class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b><span style="color:<%=session("pageColor4")%>;"></span>
						<input onBlur="document.frmParameter.submit();" type="hidden" size="8" name="requiredtxtDate" value="<%=FmtDateShort(cSDate)%>" >
						&nbsp;Sort By: <select name="optSortBy">
						<option value="0" <%if request.form("optSortBy")="0" then response.write "selected" end if%>><%=session("ClientHW")%> Name</option>
						<option value="1" <%if request.form("optSortBy")="1" or request.form("optSortBy")="" then response.write "selected" end if%>>Account Balance</option>
						</select>
					    &nbsp;Show: 
						<select name="chkNegOnly">
							<option value="0" <% if pShow="0" then response.write "selected" end if %>>All Balances</option>
							<option value="-1" <% if pShow="-1" then response.write "selected" end if %>>Negative Balances Only</option>
							<option value="1" <% if pShow="1" then response.write "selected" end if %>>Positive Balances Only</option>
							<!--JM - 47_2364 -->
							<option value="2" <% if pShow="2" then response.write "selected" end if %>>Zero Balances Only</option>
							
						</select>
						&nbsp;&nbsp;
                      <select name="optLocation" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
					  	<option value="0" <% if request.form("optLocation")="0" then response.write "selected" end if %>>All Client</option>
<%

								strSQL = "SELECT LocationID, LocationName from Location WHERE wsShow=1 ORDER BY LocationName " 
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								Set rsEntry.ActiveConnection = Nothing


								do While NOT rsEntry.EOF
%>
                        <option value="<%= rsEntry("LocationID")%>" <%if request.form("optLocation")=cstr(rsEntry("LocationID")) then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
                        <%
									rsEntry.MoveNext
								loop
								rsEntry.close
%>
                      </select>
                      <% if ss_ApplyAccountPaymentsByLocation then %>
                      <select name="optSaleLocation" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
					  	<option value="0" <% if optSaleLocation="0" then response.write "selected" end if %>>All Sale</option>
<%

								strSQL = "SELECT LocationID, LocationName from Location WHERE Active=1 AND LocationID < 100 ORDER BY LocationName " 
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								Set rsEntry.ActiveConnection = Nothing


								do While NOT rsEntry.EOF
%>
                        <option value="<%response.write rsEntry("LocationID")%>" <%if optSaleLocation=cint(rsEntry("LocationID")) then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
                        <%
									rsEntry.MoveNext
								loop
								rsEntry.close
%>
                      </select>
                      <% end if %>
                      
                      
					  <script type="text/javascript">
					  	document.frmParameter.optLocation.options[0].text = '<%=jsEscSingle(allHotWords(149))%>'+ " Home <%=jsEscDouble(allHotWords(8))%>s";
                        <% if ss_ApplyAccountPaymentsByLocation then %>
					  	document.frmParameter.optSaleLocation.options[0].text = '<%=jsEscSingle(allHotWords(149))%>'+ " Sale <%=jsEscDouble(allHotWords(8))%>s";
                        <% end if %>
					  </script>
&nbsp;
						<% if ss_Use3rdParty then %>
						&nbsp;3rd Party Payers Only:
						<input type="checkbox" name="optIs3rdParty" <% if request.form("optIs3rdParty")="on" then response.write " checked" end if %>>&nbsp;
						<% end if %>
						<!--JM - 45_2326 -->
						&nbsp;Show Event Balances Only:
						<input type="checkbox" name="optEventBal" <% if request.form("optEventBal")="on" then response.write " checked" end if %>>&nbsp;
						
&nbsp;&nbsp;
						<br />
						&nbsp;Show Clients Owing More Than&nbsp;$
						<input type="text" size="2" name="txtBalanceMax" value="<% if isNumeric(request.form("txtBalanceMax")) then response.write request.form("txtBalanceMax") end if %>">&nbsp;<%=xssStr(allHotWords(89))%>:&nbsp;
						<select name="optBillingInfo">
							<option value="all" <% if pBilling="all" then response.write " selected" end if %>>All Clients</option>
							<option value="with" <% if pBilling="with" then response.write " selected" end if %>>Clients w/&nbsp;<%=xssStr(allHotWords(89))%></option>
							<option value="without" <% if pBilling="without" then response.write " selected" end if %>>Clients w/o&nbsp;<%=xssStr(allHotWords(89))%></option>
						</select>
						&nbsp;<% taggingFilter %>
						&nbsp;&nbsp;Statement Date: 
						<input onBlur="validateDate(this, '<%=FmtDateShort(cStDate)%>', true); " type="text"  name="requiredtxtDateStatement" value="<%=FmtDateShort(cStDate)%>" class="date">
						<script type="text/javascript">
						var cal0 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStatement'});
						cal0.a_tpl.yearscroll = true;
						</script>
						&nbsp;<input type="button" value="Print All" onClick="printStatments();">&nbsp;&nbsp;
						<br />
						<input type="button" name="Button" value="Generate" onClick="genReport();"></b>
						<% if ap_RPT_SCH_BAL_AUTOPAY AND Session("mvarMIDs")<>"0" AND ss_UseEFT then %>
						<input type="button" name="Button" value="Schedule AutoPays" onClick="scheduleAutoPay();"></b>
						<% end if %>
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
				else%>
						<% exportToExcelButton %>
				<%end if%>
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
				 else
				 	taggingButtons("frmParameter")
				 end if%>
						<% savingButtons "frmParameter", "Account Balances" %>
						</td>
						</tr>
												
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" class="mainTextBig"> 
					
					<table id="accountBalancesReport" class="mainText center" width="95%" cellspacing="0">

						<tr> 
						<td  colspan="8" valign="top" class="mainTextBig center-ch">
		<% 
		end if			'end of frmExpreport value check before /head line	  
		%>
		<table class="mainText"  width="100%" cellspacing="0">
		<% 
							if request.form("frmTagClients")="true" then
								
								strSQL = "SELECT tblClientAccount.ClientID "
								strSQL = strSQL & "FROM " & tblClientAccountStr & " INNER JOIN CLIENTS ON tblClientAccount.ClientID = CLIENTS.ClientID " 
								if request.form("optIs3rdParty")="on" then
									strSQL = strSQL & " AND CLIENTS.Is3rdParty = 1 "
								end if
								strSQL = strSQL & " LEFT OUTER JOIN tblCCNumbers ON CLIENTS.ClientID = tblCCNumbers.ClientID "
								strSQL = strSQL & " LEFT OUTER JOIN Sales ON Sales.SaleID = tblClientAccount.SaleID "
								if request.form("optFilterTagged")="on" then
									strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
								 	if session("mVarUserID")<>"" then
										strSQL = strSQL & " AND smodeID = " & session("mVarUserID")
									end if
									strSQL = strSQL & " ) "
								end if
								strSQL = strSQL & "WHERE " 
								
								strSQL = strSQL & "tblClientAccount.EntryDate <= " & DateSep & cStDate & DateSep & " AND "
								if request.form("optLocation")<>"0" and IsNum(request.form("optLocation")) then
									strSQL = strSQL & " (Clients.HomeStudio=" & request.form("optLocation") & " OR CLIENTS.HomeStudio=0) AND "
								end if
								if request.Form("optSaleLocation")<>"0" AND IsNum(request.form("optSaleLocation")) then
									strSQL = strSQL & " Sales.LocationID = " & request.Form("optSaleLocation") & " AND "
								end if
								strSQL = strSQL & " (tblClientAccount.ClientID <> 1) "
								'JM - 45_2326
								if request.form("optEventBal")="on" then
									strSQL = strSQL & " AND NOT (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
								else
									strSQL = strSQL & " AND (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
								end if
								strSQL = strSQL & " AND (tblClientAccount.ClientContractID IS NULL OR NOT(tblClientAccount.DepositReleaseDate IS NULL)) "
								if pBilling="with" then
									strSQL = strSQL & "AND ((CreditCardNo IS NOT NULL) OR (ACHAccountNum IS NOT NULL))"
								elseif pBilling="without" then
									strSQL = strSQL & "AND (CreditCardNo IS NULL) AND (ACHAccountNum IS NULL) "
								end if
								strSQL = strSQL & "GROUP BY tblClientAccount.ClientID "
								if pShow="0" then
									strSQL = strSQL & "HAVING SUM(tblClientAccount.Amount)<>0 " 
								elseif pShow="-1" then
									strSQL = strSQL & "HAVING SUM(tblClientAccount.Amount)<0 " 
								elseif pShow="1" then
									strSQL = strSQL & "HAVING SUM(tblClientAccount.Amount)>0 " 
								' JM - 47_2364
								elseif pShow="2" then
									strSQL = strSQL & "HAVING SUM(tblClientAccount.Amount)=0 " 
 
								end if
								if isNumeric(request.form("txtBalanceMax")) then
									strSQL = strSQL & " AND SUM(tblClientAccount.Amount)<" & request.form("txtBalanceMax")*-1 & " "
								end if
								
								if request.form("frmTagClientsNew")="true" then
									clearAndTagQuery(strSQL)
								else
									tagQuery(strSQL)
								end if
							end if
				
							if request.form("frmGenReport")="true" then 
								if request.form("frmExpReport")="true" then
									Dim stFilename
									stFilename="attachment; filename=Account_Balance_as_of_" & Replace(cStDate,"/","-") & ".xls" 									
									Response.ContentType = "application/vnd.ms-excel" 
									Response.AddHeader "Content-Disposition", stFilename 
								end if
								rsClientID=0
									rsRssid=0
								rsBalance=0
								unpaidTotal=0
								rsUnpaid=""
								
								strSQL = "SELECT Clients.RSSID, CLIENTS.LastName, CLIENTS.FirstName, Clients.HomeStudio, tblCCNumbers.ACHAccountNum, tblCCNumbers.CreditCardNo, tblClientAccount.ClientID, SUM(tblClientAccount.Amount) AS AccountBal, Bal1.ClientBalance AS Bal30, "
								strSQL = strSQL & "Bal2.ClientBalance AS Bal60, Bal3.ClientBalance AS Bal90, Bal4.ClientBalance AS Bal90Plus, Unpaid.UnpaidRem, EmailName, CellPhone, HomePhone, WorkPhone  "
								strSQL = strSQL & "FROM " & tblClientAccountStr & " INNER JOIN CLIENTS ON tblClientAccount.ClientID = CLIENTS.ClientID " 
								if request.form("optIs3rdParty")="on" then
									strSQL = strSQL & " AND CLIENTS.Is3rdParty = 1 "
								end if
								strSQL = strSQL & " LEFT OUTER JOIN tblCCNumbers ON CLIENTS.ClientID = tblCCNumbers.ClientID "
								strSQL = strSQL & " LEFT OUTER JOIN Sales ON Sales.SaleID = tblClientAccount.SaleID "
								if request.form("optFilterTagged")="on" then
									strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
								 	if session("mVarUserID")<>"" then
										strSQL = strSQL & " AND smodeID = " & session("mVarUserID")
									end if
									strSQL = strSQL & " ) "
								end if
								
								strSQL = strSQL & " LEFT OUTER JOIN "
								' 0 - 29 days ago
								strSQL = strSQL & "(SELECT SUM(Amount) AS ClientBalance, tblClientAccount.ClientID FROM " & tblClientAccountStr & " "
								strSQL = strSQL & " LEFT OUTER JOIN Sales ON Sales.SaleID = tblClientAccount.SaleID "
								strSQL = strSQL & " WHERE (tblClientAccount.ClientID <> 1) "
								'JM - 45_2326
								if request.form("optEventBal")="on" then
									strSQL = strSQL & " AND NOT (ClassID IS NULL OR ClassID = 0) "
								else
									strSQL = strSQL & " AND (ClassID IS NULL OR ClassID = 0) "
								end if
								strSQL = strSQL & " AND (ClientContractID IS NULL) AND (Amount <> 0) AND (EntryDate < " & DateSep & CDATE(DateAdd("d", 1, cStDate)) & DateSep & ") AND (EntryDate > " & DateSep & CDATE(DateAdd("d", -30, cStDate)) & DateSep & ") "
								if request.Form("optSaleLocation")<>"0" AND IsNum(request.form("optSaleLocation")) then
									strSQL = strSQL & " AND Sales.LocationID = " & request.Form("optSaleLocation") & " "
								end if
								strSQL = strSQL & "GROUP BY tblClientAccount.ClientID) Bal1 ON Bal1.ClientID = CLIENTS.ClientID LEFT OUTER JOIN "
								' 30 - 59 days ago
								strSQL = strSQL & "(SELECT SUM(Amount) AS ClientBalance, tblClientAccount.ClientID FROM " & tblClientAccountStr & " "
								strSQL = strSQL & " LEFT OUTER JOIN Sales ON Sales.SaleID = tblClientAccount.SaleID "
								strSQL = strSQL & "WHERE (tblClientAccount.ClientID <> 1) "
								'JM - 45_2326
								if request.form("optEventBal")="on" then
									strSQL = strSQL & " AND NOT (ClassID IS NULL OR ClassID = 0) "
								else
									strSQL = strSQL & " AND (ClassID IS NULL OR ClassID = 0) "
								end if
								strSQL = strSQL & " AND (ClientContractID IS NULL) AND (Amount <> 0) AND (EntryDate <= " & DateSep & CDATE(DateAdd("d", -30, cStDate)) & DateSep & ") AND (EntryDate > " & DateSep & CDATE(DateAdd("d", -60, cStDate)) & DateSep & ") "
								if request.Form("optSaleLocation")<>"0" AND IsNum(request.form("optSaleLocation")) then
									strSQL = strSQL & " AND Sales.LocationID = " & request.Form("optSaleLocation") & " "
								end if
								strSQL = strSQL & "GROUP BY tblClientAccount.ClientID) Bal2 ON Bal2.ClientID = CLIENTS.ClientID LEFT OUTER JOIN "
								' 60 - 89 days ago
								strSQL = strSQL & "(SELECT SUM(Amount) AS ClientBalance, tblClientAccount.ClientID FROM " & tblClientAccountStr & " "
								strSQL = strSQL & " LEFT OUTER JOIN Sales ON Sales.SaleID = tblClientAccount.SaleID "
								strSQL = strSQL & "WHERE (tblClientAccount.ClientID <> 1) "
								'JM - 45_2326
								if request.form("optEventBal")="on" then
									strSQL = strSQL & " AND NOT (ClassID IS NULL OR ClassID = 0) "
								else
									strSQL = strSQL & " AND (ClassID IS NULL OR ClassID = 0) "
								end if
								strSQL = strSQL & " AND (ClientContractID IS NULL) AND (Amount <> 0) AND (EntryDate <= " & DateSep & CDATE(DateAdd("d", -60, cStDate)) & DateSep & ") AND (EntryDate > " & DateSep & CDATE(DateAdd("d", -90, cStDate)) & DateSep & ") "
								if request.Form("optSaleLocation")<>"0" AND IsNum(request.form("optSaleLocation")) then
									strSQL = strSQL & " AND Sales.LocationID = " & request.Form("optSaleLocation") & " "
								end if
								strSQL = strSQL & "GROUP BY tblClientAccount.ClientID) Bal3 ON Bal3.ClientID = CLIENTS.ClientID LEFT OUTER JOIN "
								' 90+ days ago
								strSQL = strSQL & "(SELECT SUM(Amount) AS ClientBalance, tblClientAccount.ClientID FROM " & tblClientAccountStr & " "
								strSQL = strSQL & " LEFT OUTER JOIN Sales ON Sales.SaleID = tblClientAccount.SaleID "
								strSQL = strSQL & "WHERE (tblClientAccount.ClientID <> 1) "
								'JM - 45_2326
								if request.form("optEventBal")="on" then
									strSQL = strSQL & " AND NOT (ClassID IS NULL OR ClassID = 0) "
								else
									strSQL = strSQL & " AND (ClassID IS NULL OR ClassID = 0) "
								end if
								strSQL = strSQL & " AND (ClientContractID IS NULL) AND (Amount <> 0) AND (EntryDate <= " & DateSep & CDATE(DateAdd("d", -90, cStDate)) & DateSep & ") "
								if request.Form("optSaleLocation")<>"0" AND IsNum(request.form("optSaleLocation")) then
									strSQL = strSQL & " AND Sales.LocationID = " & request.Form("optSaleLocation") & " "
								end if
								strSQL = strSQL & "GROUP BY tblClientAccount.ClientID) Bal4 ON Bal4.ClientID = CLIENTS.ClientID LEFT OUTER JOIN "
								' unpaids
								strSQL = strSQL & "(SELECT ClientID, SUM(Remaining) AS UnpaidRem FROM [PAYMENT DATA] "
								strSQL = strSQL & "WHERE (ExpDate > " & DateSep & Date & DateSep & ") AND (Type = 9) "
								if request.Form("optSaleLocation")<>"0" AND IsNum(request.form("optSaleLocation")) then
									strSQL = strSQL & " AND Location = " & request.Form("optSaleLocation") & " "
								end if
								strSQL = strSQL & "GROUP BY ClientID) Unpaid ON Unpaid.ClientID = CLIENTS.ClientID "
								strSQL = strSQL & "WHERE " 
								
								strSQL = strSQL & "tblClientAccount.EntryDate < " & DateSep & CDATE(DateAdd("d", 1, cStDate)) & DateSep & " AND "
								if request.form("optLocation")<>"0" AND IsNum(request.form("optLocation")) then
									strSQL = strSQL & " (Clients.HomeStudio=" & request.form("optLocation") & " OR CLIENTS.HomeStudio=0) AND "
								end if
								if request.Form("optSaleLocation")<>"0" AND IsNum(request.form("optSaleLocation")) then
									strSQL = strSQL & " Sales.LocationID = " & request.Form("optSaleLocation") & " AND "
								end if
								strSQL = strSQL & " (tblClientAccount.ClientID <> 1) "
								'JM - 45_2326
								if request.form("optEventBal")="on" then
									strSQL = strSQL & " AND NOT (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
								else
									strSQL = strSQL & " AND (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
								end if

								strSQL = strSQL & " AND (tblClientAccount.ClientContractID IS NULL OR NOT(tblClientAccount.DepositReleaseDate IS NULL)) "
								if pBilling="with" then
									strSQL = strSQL & "AND ((CreditCardNo IS NOT NULL) OR (ACHAccountNum IS NOT NULL))"
								elseif pBilling="without" then
									strSQL = strSQL & "AND (CreditCardNo IS NULL) AND (ACHAccountNum IS NULL) "
								end if
								strSQL = strSQL & "GROUP BY Clients.RSSID, CLIENTS.LastName, CLIENTS.FirstName, Clients.HomeStudio, tblCCNumbers.ACHAccountNum, tblCCNumbers.CreditCardNo, tblClientAccount.ClientID, Bal1.ClientBalance, Bal2.ClientBalance, Bal3.ClientBalance, Bal4.ClientBalance, Unpaid.UnpaidRem, EmailName, CellPhone, HomePhone, WorkPhone "
								if pShow="0" then
									strSQL = strSQL & "HAVING SUM(tblClientAccount.Amount)<>0 " 
								elseif pShow="-1" then
									strSQL = strSQL & "HAVING SUM(tblClientAccount.Amount)<0 " 
								elseif pShow="1" then
									strSQL = strSQL & "HAVING SUM(tblClientAccount.Amount)>0 " 
								elseif pShow="2" then
									strSQL = strSQL & "HAVING SUM(tblClientAccount.Amount)=0 " 
								end if
								if isNumeric(request.form("txtBalanceMax")) then
									strSQL = strSQL & " AND SUM(tblClientAccount.Amount)<" & request.form("txtBalanceMax")*-1 & " "
								end if
								
								'BJD 5/2/08 - Split up report and tagging SQL
								if request.form("optSortBy")="0" then
									strSQL = strSQL & "ORDER BY CLIENTS.LastName, SUM(tblClientAccount.Amount) "
								else
									strSQL = strSQL & "ORDER BY SUM(tblClientAccount.Amount), CLIENTS.LastName "
								end if
							response.write debugSQL(strSQL, "SQL")
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								Set rsEntry.ActiveConnection = Nothing

								if NOT rsEntry.EOF then			'EOF
%>
								<% if NOT request.form("frmExpReport")="true" then %>
									<tr>
									  <td colspan="11">&nbsp;</td>
									</tr>
								<% end if %>
									<tr>
									  <% if ShowAutoPaySchedule then %>
									  <td>&nbsp;</td>
									  <% end if %>
									  <td colspan="3">&nbsp;</td>
									  <td colspan="4" class="center-ch"><strong>A C C O U N T&nbsp;&nbsp;&nbsp;A C T I V I T Y&nbsp;&nbsp;&nbsp;S U M M A R Y</strong></td>
									  <td colspan="1">&nbsp;</td>
									</tr>
									<tr valign="bottom">
									<% if request.form("frmExpReport")="true" then %>
									  <td><strong><%= getHotWord(134)%></strong></td>
									<% end if %>
									<% if ShowAutoPaySchedule then %>
									  <td><strong>Schedule Account AutoPays</strong><br />
										<a href="javascript:checkAll(document.getElementById('frmParameter'), 'filecheck', true);"><%=xssStr(allHotWords(147))%>&nbsp;<%=xssStr(allHotWords(149))%></a> | <a href="javascript:checkAll(document.getElementById('frmParameter'), 'filecheck', false);"><%=xssStr(allHotWords(148))%>&nbsp;<%=xssStr(allHotWords(149))%></a>									  
									  </td>
									<% end if %>
									  <td width="100"><div align="left"><strong><%=session("ClientHW")%>&nbsp;</strong></div></td>
									  <td><div><strong><%=getHotWord(39)%></strong></div></td>
									  <td><div><strong><%=getHotWord(93)%></strong></div></td>
									  <td class="right" ><strong>Account Balance </strong></td>
									  <td><div><strong>&nbsp;</strong></div></td>
									  <td><div class="right"><strong>&nbsp;&nbsp;0-29 Days Ago&nbsp;&nbsp;</strong></div></td>
									  <td><div class="right"><strong>&nbsp;&nbsp;30-59 Days Ago&nbsp;&nbsp;</strong></div></td>
									  <td><div class="right"><strong>&nbsp;&nbsp;60-89 Days Ago&nbsp;&nbsp;</strong></div></td>
									  <td><div class="right"><strong>&nbsp;&nbsp;90+ Days Ago&nbsp;&nbsp;</strong></div></td>
									<% if NOT request.form("frmExpReport")="true" then %>
									  <td nowrap class="center-ch"><div><p><strong>&nbsp;Unpaid Sessions</p>
								      </div></td>
									<% else %>
									  <td nowrap class="center-ch"><div><strong>Unpaid Sessions</strong></div></td>
									<% end if %>
									</tr>
									<% if NOT request.form("frmExpReport")="true" then %>
									<tr height="2">
										<td colspan="11" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
									</tr>
									<% end if %>
<%
									do while NOT rsEntry.EOF

										'Schedule AutoPays
										if request.form("frmGenAutoPaySch")="true" AND request.form("optAutoPay"&rsEntry("ClientID"))="on" then
										dim intlAmt100
											strSQL = "INSERT INTO tblEFTSchedule ( ClientID, RecClientID, ScheduleDate, Amount, Method, ProductID, ClassID, SaleLoc, StatusCode, StatusMessage, ScheduledBy, ScheduledDateTime, AcountAutoPay) VALUES ("
											strSQL = strSQL & rsEntry("ClientID") & ", " & rsEntry("ClientID")
											strSQL = strSQL & ", " & DateSep & cBillingDate & DateSep
											Call SetLocale(session("mvarLocaleStr"))
												intlAmt100 = cdbl(request.form("requiredtxtAmount"&rsEntry("ClientID")))*100
											Call SetLocale("en-us")
											strSQL = strSQL & ", " & intlAmt100/100
											strSQL = strSQL & ", " & request.form("optPayMethod"&rsEntry("ClientID"))
											strSQL = strSQL & ", -6" 	'Payment on Account Product
											strSQL = strSQL & ", null"	'ClassID
											strSQL = strSQL & ", " & request.form("optLocation"&rsEntry("ClientID"))
											strSQL = strSQL & ", 1"	'''Status Code 1 - Scheduled
											strSQL = strSQL & ", 'Scheduled'"
											if session("Admin")="sa" OR session("Admin")="owner" then	'sa or owner
												strSQL = strSQL & ", 0"
											elseif session("empID")<>"" then 	''Reg Business Mode
												strSQL = strSQL & ", " & session("empID")
											else
												strSQL = strSQL & ", null"
											end if
											strSQL = strSQL & ", " & DateSep & DateAdd("n", Session("tzOffset"),Now) & DateSep
											strSQL = strSQL & ", 1" 'AccountAutoPay
											strSQL = strSQL & ")"
										response.write debugSQL(strSQL, "SQL")
											cnWS.execute strSQL
										end if	''Schedule AutoPays

										rsClientID=rsEntry("ClientID")
                    rsRssid=rsEntry("RSSID")
										rsBalance=rsEntry("AccountBal")
										if NOT isNull(rsEntry("Bal30")) then
											tmpBal30 = rsEntry("Bal30")
										else
											tmpBal30 = 0
										end if
										if NOT isNull(rsEntry("Bal60")) then
											tmpBal60 = rsEntry("Bal60")
										else
											tmpBal60 = 0
										end if
										if NOT isNull(rsEntry("Bal90")) then
											tmpBal90 = rsEntry("Bal90")
										else
											tmpBal90 = 0
										end if
										if NOT isNull(rsEntry("Bal90Plus")) then
											tmpBal90Plus = rsEntry("Bal90Plus")
										else
											tmpBal90Plus = 0
										end if
										
										
										pCltPhone = ""
            							if NOT isNULL(rsEntry("CellPhone")) then
                                            pCltPhone = FmtPhoneNum(rsEntry("CellPhone"))
                                        elseif NOT isNULL(rsEntry("HomePhone")) then
                                            pCltPhone = FmtPhoneNum(rsEntry("HomePhone"))
                                        elseif NOT isNULL(rsEntry("WorkPhone")) then
                                            pCltPhone = FmtPhoneNum(rsEntry("WorkPhone"))
                                        end if
										
										' total balances
										balanceTotal = balanceTotal + rsEntry("AccountBal")
										totalBal30 = totalBal30 + tmpBal30
										totalBal60 = totalBal60 + tmpBal60
										totalBal90 = totalBal90 + tmpBal90
										totalBal90Plus = totalBal90Plus + tmpBal90Plus
										
										if isNull(rsEntry("UnpaidRem")) OR ABS(rsEntry("UnpaidRem")) = 0 then		
											rsUnpaid = ""
										elseif ABS(rsEntry("UnpaidRem")) = 1 then
											if NOT request.form("frmExpReport")="true" then 
												rsUnpaid = ABS(rsEntry("UnpaidRem")) & " Unpaid"
											else
												rsUnpaid = ABS(rsEntry("UnpaidRem"))
											end if
											unpaidTotal = unpaidTotal + ABS(rsEntry("UnpaidRem"))
										else
											if NOT request.form("frmExpReport")="true" then 
												rsUnpaid = ABS(rsEntry("UnpaidRem")) & " Unpaids"
											else
												rsUnpaid = ABS(rsEntry("UnpaidRem"))
											end if
											unpaidTotal = unpaidTotal + ABS(rsEntry("UnpaidRem")) 
										end if										
%>
									<tr align="left">
									<% if NOT request.form("frmExpReport")="true" then %>
									  <% if ShowAutoPaySchedule then %>
									  <td style="white-space: nowrap;">
									    <input class="filecheck" name="optAutoPay<%=rsEntry("ClientID")%>" type="checkbox">
									  	<input onBlur="validateCurrency(this);updateHidden(this);" type="text" id="requiredtxtAmount<%=rsEntry("ClientID")%>" name="requiredtxtAmount<%=rsEntry("ClientID")%>" maxlength="10" size="7" value="<%if request.form("requiredtxtAmount"&rsEntry("ClientID"))<>"" then response.write FmtNumber(request.form("requiredtxtAmount"&rsEntry("ClientID"))) else response.write FmtNumber(rsBalance*-1) end if%>">
									  	<input type="hidden" id="hiddentxtAmount<%=rsEntry("ClientID")%>" value="<%if request.form("requiredtxtAmount"&rsEntry("ClientID"))<>"" then response.write (request.form("requiredtxtAmount"&rsEntry("ClientID"))*100) else response.write (rsBalance*-1*100) end if%>"/>
										<select name="optPayMethod<%=rsEntry("ClientID")%>" id="optPayMethod<%=rsEntry("ClientID")%>">
										  <option value="1" <%if request.form("optPayMethod"&rsEntry("ClientID"))="1" then response.write "selected" end if%>>Credit Card</option>
										<%if ss_EnableACH AND NOT isNULL(rsEntry("ACHAccountNum")) then %>
										  <option value="3" <%if NOT isNULL(rsEntry("ACHAccountNum")) AND request.form("optPayMethod"&rsEntry("ClientID"))<>"1" then response.write "selected" end if%>>ACH</option>
										<%end if%>
										</select>
										<%if session("numLocations")>1 then%>
											<input type="hidden" name="optLocation<%=rsEntry("ClientID")%>" value="<%if NOT isNULL(rsEntry("HomeStudio")) AND rsEntry("HomeStudio")<>"0" then response.write rsEntry("HomeStudio") else response.write "1" end if%>">
										<%else%>
											<input type="hidden" name="optLocation<%=rsEntry("ClientID")%>" value="<%=session("curLocation")%>">
										<%end if%>
									  </td>
									  <% end if %>
									<% end if %>
									<% if request.form("frmExpReport")="true" then %>
										<td class="right nowrap"><%=rsRssid%></td>
									<% end if %>
									<% if NOT request.form("frmExpReport")="true" then %>
										<td class="nowrap"><div><a href="adm_clt_ph.asp?ID=<%=rsClientID%>"><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></a></div></td>
									<% else %>
										<td><div><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></div></td>
									<% end if %>
									<% if NOT request.form("frmExpReport")="true" then %>
										<td><div><a href="mailto:<%=rsEntry("EmailName")%>"><%=rsEntry("EmailName")%></a></div></td>
									<% else %>
										<td><div><%=rsEntry("EmailName")%></div></td>
									<% end if %>
									<td nowrap><div><%=pCltPhone%></div></td>
									<% if NOT request.form("frmExpReport")="true" then %>
									  <td><div class="right"><%if rsBalance<0 then %><span style="color:#990000;"><%=FmtCurrency(rsBalance)%></span><%else%><%=FmtCurrency(rsBalance)%><%end if%>&nbsp;&nbsp;</div></td>
									<% else %>
									  <td><div class="right"><%if rsBalance<0 then %><span style="color:#990000;"><%=FmtNumber(rsBalance)%></span><%else%><%=FmtNumber(rsBalance)%><%end if%></div></td>
									<% end if %>
									<% if NOT request.form("frmExpReport")="true" then %>
									
									  <td class="nowrap"><%'JM-55_3182-if NOT (request.form("optEventBal")="on") then%><div>&nbsp;<!--<a href="adm_rpt_acct_bal_statmnt.asp?ClientID=<%=rsClientID%>">--><a href="javascript:createInvoice(<%=rsClientID%>);">[Create Statement]</a></div><%'end if%></td>
									<% else %>
									  <td></td>
									<% end if %>
									<% if NOT request.form("frmExpReport")="true" then %>  
									  <td><div class="right"><%if tmpBal30<0 then %><span style="color:#990000;"><%=FmtCurrency(tmpBal30)%></span><%else%><%=FmtCurrency(tmpBal30)%><%end if%>&nbsp;&nbsp;</div></td>
									<% else %>
									  <td><div class="right"><%if tmpBal30<0 then %><span style="color:#990000;"><%=FmtNumber(tmpBal30)%></span><%else%><%=FmtNumber(tmpBal30)%><%end if%></div></td>
									<% end if %>
									<% if NOT request.form("frmExpReport")="true" then %>  
									  <td><div class="right"><%if tmpBal60<0 then %><span style="color:#990000;"><%=FmtCurrency(tmpBal60)%></span><%else%><%=FmtCurrency(tmpBal60)%><%end if%>&nbsp;&nbsp;</div></td>
									<% else %>
									  <td><div class="right"><%if tmpBal60<0 then %><span style="color:#990000;"><%=FmtNumber(tmpBal60)%></span><%else%><%=FmtNumber(tmpBal60)%><%end if%></div></td>
									<% end if %>
									<% if NOT request.form("frmExpReport")="true" then %>  
									  <td><div class="right"><%if tmpBal90<0 then %><span style="color:#990000;"><%=FmtCurrency(tmpBal90)%></span><%else%><%=FmtCurrency(tmpBal90)%><%end if%>&nbsp;&nbsp;</div></td>
									<% else %>
									  <td><div class="right"><%if tmpBal90<0 then %><span style="color:#990000;"><%=FmtNumber(tmpBal90)%></span><%else%><%=FmtNumber(tmpBal90)%><%end if%></div></td>
									<% end if %>
									<% if NOT request.form("frmExpReport")="true" then %>  
									  <td><div class="right"><%if tmpBal90Plus<0 then %><span style="color:#990000;"><%=FmtCurrency(tmpBal90Plus)%></span><%else%><%=FmtCurrency(tmpBal90Plus)%><%end if%>&nbsp;&nbsp;</div></td>
									<% else %>
									  <td><div class="right"><%if tmpBal90Plus<0 then %><span style="color:#990000;"><%=FmtNumber(tmpBal90Plus)%></span><%else%><%=FmtNumber(tmpBal90Plus)%><%end if%></div></td>
									<% end if %>
									<% if NOT request.form("frmExpReport")="true" then %>
									  <td><div class="center-ch"><a href="main_retail.asp?cltID=<%=rsClientID%>"><%=rsUnpaid%></a></div></td>
									<% else %>
									  <td><div class="center-ch"><%=rsUnpaid%></div></td>
									<% end if %>
									</tr>
<%
										rsEntry.MoveNext
									loop
									rsEntry.close
									set rsEntry = nothing
									
									if request.form("frmGenAutoPaySch")="true" then	'forward to AutoPay Detail for Today
										response.redirect "adm_eft_det.asp?category=PaymentProcessing&AccountAP=true&eDate=" & FmtDateShort(cBillingDate)
									end if
%>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="2">
									<td colspan="11" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
								</tr>
								<% end if %>
								</tr>
								<tr align="left">
								<% if ShowAutoPaySchedule then %>
								  <td>
								      <input type="button" name="RunBtn" value="Schedule Selected AutoPays" onClick="GenerateAutoPaySchedules();">
								      <br />On: <input type="text" name="optAutopayScheduleDate" id="optAutopayScheduleDate" value="<%=FmtDateShort(cBillingDate)%>" onBlur="validateDate(this, '<%=FmtDateShort(cStDate)%>', true);" class="date">
								      <script type="text/javascript">
						                  var cal1 = new tcal({'formname':'frmParameter', 'controlname':'optAutopayScheduleDate'});
						                  cal1.a_tpl.yearscroll = true;
						              </script>
						          </td>
								<% end if %>
								<% if request.form("frmExpReport")="true" then %>
								  <td nowrap colspan=4 class="right">Total:</td>
								<% else %>
								  <td nowrap colspan=3 class="right">Total:</td>
								<% end if %>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td><div class="right"><%if balanceTotal<0 then %><span style="color:#990000;"><%=FmtCurrency(balanceTotal)%></span><%else%><%=FmtCurrency(balanceTotal)%><%end if%>&nbsp;&nbsp;</div></td>
								<% else %>
								  <td><div class="right"><%if balanceTotal<0 then %><span style="color:#990000;"><%=FmtNumber(balanceTotal)%></span><%else%><%=FmtNumber(balanceTotal)%><%end if%></div></td>
								<% end if %>
								  <td>&nbsp;</td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td><div class="right"><%if totalBal30<0 then %><span style="color:#990000;"><%=FmtCurrency(totalBal30)%></span><%else%><%=FmtCurrency(totalBal30)%><%end if%>&nbsp;&nbsp;</div></td>
								<% else %>
								  <td><div class="right"><%if totalBal30<0 then %><span style="color:#990000;"><%=FmtNumber(totalBal30)%></span><%else%><%=FmtNumber(totalBal30)%><%end if%></div></td>
								<% end if %>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td><div class="right"><%if totalBal60<0 then %><span style="color:#990000;"><%=FmtCurrency(totalBal60)%></span><%else%><%=FmtCurrency(totalBal60)%><%end if%>&nbsp;&nbsp;</div></td>
								<% else %>
								  <td><div class="right"><%if totalBal60<0 then %><span style="color:#990000;"><%=FmtNumber(totalBal60)%></span><%else%><%=FmtNumber(totalBal60)%><%end if%></div></td>
								<% end if %> 
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td><div class="right"><%if totalBal90<0 then %><span style="color:#990000;"><%=FmtCurrency(totalBal90)%></span><%else%><%=FmtCurrency(totalBal90)%><%end if%>&nbsp;&nbsp;</div></td>
								<% else %>
								  <td><div class="right"><%if totalBal90<0 then %><span style="color:#990000;"><%=FmtNumber(totalBal90)%></span><%else%><%=FmtNumber(totalBal90)%><%end if%></div></td>
								<% end if %> 
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td><div class="right"><%if totalBal90Plus<0 then %><span style="color:#990000;"><%=FmtCurrency(totalBal90Plus)%></span><%else%><%=FmtCurrency(totalBal90Plus)%><%end if%>&nbsp;&nbsp;</div></td>
								<% else %>
								  <td><div class="right"><%if totalBal90Plus<0 then %><span style="color:#990000;"><%=FmtNumber(totalBal90Plus)%></span><%else%><%=FmtNumber(totalBal90Plus)%><%end if%></div></td>
								<% end if %> 
								  <td><div class="center-ch"><%=unpaidTotal%> Unpaid<% if unpaidTotal > 1 then %>s<% end if %></div></td>
								</tr>
<%
								end if	'eof
%>
							<% end if ' end gen report %>
						  </table>
						  </td>
						</tr>
						</table>
						</td>
						</tr></form>
					</table>
				</td>
				</tr>
</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%
	
end if
%>
