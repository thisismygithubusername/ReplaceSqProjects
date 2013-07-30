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
	dim rsEntry, rsEntry2, rsClient,firstPM, rsEntry3
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	set rsEntry2 = Server.CreateObject("ADODB.Recordset")
	set rsEntry3 = Server.CreateObject("ADODB.Recordset")
	set rsClient = Server.CreateObject("ADODB.Recordset")
	%>
	<!-- #include file="inc_accpriv.asp" -->
	<!-- #include file="inc_acct_balance.asp" -->
	<!-- #include file="inc_rpt_tagging.asp" -->
	<!-- #include file="inc_hotword.asp" -->
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
		Dim showDetails, cSttDate, cSDate, cEDate, tmpCurDate, CurClientName, clientID, ss_ShowInsurance
		Dim rowColor, Balance, Amount, stLineItem, stLineItemDetail, BeginBal, isInvoice, curBalance, BalByCategory, AmountByCategory, AcctCredit
		Dim studioAddress, pStudioPhone, pPayMethod, contText, pStudioURL, pCltAddrLine1, pCltAddrLine2, pCltPhone, pRecLn1, pRecLn2, pCltInsComp, pCltInsPolicy, pCltBirthdate, pActLineCount, pPurchLineCount, optEventBal

		ss_ShowInsurance = checkStudioSetting("tblGenOpts", "InsuranceFields")
		Dim ss_InvoiceShowCC : ss_InvoiceShowCC = checkStudioSetting("tblGenOpts", "InvoiceShowCC")
		Dim ss_InvoiceAskAutoPay : ss_InvoiceAskAutoPay = checkStudioSetting("tblGenOpts", "InvoiceAskAutoPay")
		Dim ss_IncPurchActInStmt : ss_IncPurchActInStmt = checkStudioSetting("tblGenOpts", "IncPurchActInStmt") 'JM-51_2782
		Dim ss_ShowReceiptMessage : ss_ShowReceiptMessage = checkStudioSetting("tblGenOpts",  "InvoiceShowReceiptMsg")
		Dim	receiptMsg1 : receiptMsg1 = checkStudioSetting("tblGenOpts",  "receiptLn1") 
		Dim	receiptMsg2 : receiptMsg2 = checkStudioSetting("tblGenOpts",  "receiptLn2") 
		
		dim rsLogo, cltcount
		set rsLogo = Server.CreateObject("ADODB.Recordset")
		strSQL = "SELECT tblAppearance.LogoHeight, tblAppearance.LogoWidth, Studios.StudioURL FROM tblAppearance INNER JOIN Studios ON Studios.StudioID = tblAppearance.StudioID WHERE (tblAppearance.StudioID = " & session("StudioID") & ")"
		rsLogo.CursorLocation = 3
		rsLogo.open strSQL, cnWS
		Set rsLogo.ActiveConnection = Nothing
			''Standard
			dim logoH, logoW, mvStudioURL
			mvStudioURL = TRIM(rsLogo("StudioURL")) & session("rtnURLqs")
			logoH = rsLogo("LogoHeight")
			logoW = rsLogo("LogoWidth")
		rsLogo.close%>
<!--JM-54_2836-->
<%
dim phraseDictionary
set phraseDictionary = LoadPhrases("BusinessmodebalancestatementprintingPage", 76)

%>		
<%		if request.form("requiredtxtDateStatement")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cSttDate = CDATE(request.form("requiredtxtDateStatement"))
			Call SetLocale("en-us")
		else
			cSttDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		end if

		if request.form("requiredtxtDateStart")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cSDate = CDATE(request.form("requiredtxtDateStart"))
			Call SetLocale("en-us")
		else
			cSDate = DateAdd("m",-1,cSttDate)
		end if
	
		'if request.form("requiredtxtDateEnd")<>"" then
		'	Call SetLocale(session("mvarLocaleStr"))
		'		cEDate = CDATE(request.form("requiredtxtDateEnd"))
		'	Call SetLocale("en-us")
		'else
		'	cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		'end if
		
		CurClientName=""
		
		if request.querystring("ClientID")<>"" then
			clientID=CLNG(request.querystring("ClientID"))
		end if
		
'JM-55_3182
		if request.QueryString("optEventBal")="on" then
		    optEventBal = request.QueryString("optEventBal")
		elseif request.form("optEventBal")="on" then
		    optEventBal = request.form("optEventBal")
		else
		    optEventBal = ""
		end if
		
		if session("curLocation")<>"0" then		
			strSQL = "SELECT Address, AddressLn2, Phone From Location WHERE LocationID=" & session("curLocation")
		else 
			strSQL = "SELECT TOP 1 Address, AddressLn2, Phone From Location ORDER BY LocationID"
		end if
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing

		if not rsEntry.eof then
			studioAddress = rsEntry("Address") & "<br />"
			if NOT isNULL(rsEntry("AddressLn2")) then
				studioAddress = studioAddress & rsEntry("AddressLn2") & "<br />"
			end if
			pStudioPhone = rsEntry("Phone")
		else
			studioAddress = ""
			pStudioPhone = ""
		end if
		rsEntry.close
		
		strSQL = "SELECT Studios.StudioURL FROM Studios WHERE STUDIOS.StudioID=" & session("StudioID")
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing

		if NOT isNULL(rsEntry("StudioURL")) then
			pStudioURL = rsEntry("StudioURL") & "<br />"
		else
			pStudioURL = ""		
		end if
		pStudioURL = Replace(pStudioURL, "http://", "")
		rsEntry.close
	
		if NOT request.form("frmExpReport")="true" then
		%>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_acct_bal_statmnt_mult")) %>
			<script type="text/javascript">
			function genReport() {
				document.frmParameter.frmGenReport.value = "true";
				document.frmParameter.frmExpReport.value = "false";
				//document.frmParameter.action = 'adm_rpt_acct_bal_statmnt.asp?ClientID=<%=clientID%>';
				document.frmParameter.submit();
			}
			function exportReport() {
				document.frmParameter.frmExpReport.value = "true";
				document.frmParameter.frmGenReport.value = "true";
				<% iframeSubmit "frmParameter", "adm_rpt_acct_bal_statmnt_mult.asp" %>
			}
			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
		<%
		end if
		
		%>
		


<% pageStart %>
		<% if request.form("frmExpReport")<>"true" AND (request.form("frmGenReport")<>"true" OR request.form("chkPrinterFriendly")<>"on") then %>

			<table height="100%" width="<%=strPageWidth%>" cellspacing="0" style="margin :0 auto">    
				<tr> 
				<td  valign="top" height="100%" width="100%"> 
				<table cellspacing="0" width="90%" height="100%" style="margin :0 auto">
					<tr>
					<td class="headText" align="left" valign="top">
					<table width="100%" cellspacing="0">
						<tr>
						<td class="headText" valign="bottom"><b> <%=DisplayPhrase(phraseDictionary, "Balancestatement")  %></b></td>
						<td valign="bottom" class="right" height="26"> </td>
						</tr>
					</table>
					</td>
					</tr>
					<tr> 
					<td height="30"  valign="bottom" class="center-ch headText">
					<table class="mainText border4" cellspacing="0">
						<form name="frmParameter" action="adm_rpt_acct_bal_statmnt_mult.asp" method="POST">
						<input type="hidden" name="frmGenReport" value="">
						<input type="hidden" name="frmExpReport" value="">
						

						<tr> 
						<td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b><span style="color:<%=session("pageColor4")%>;"></span>
						&nbsp;<%=xssStr(allHotWords(77))%>: 
						<input onBlur="document.frmParameter.submit();" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
						<script type="text/javascript">
						var cal1 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStart'});
						cal1.a_tpl.yearscroll = true;
						</script>
						&nbsp;<%=DisplayPhrase(phraseDictionary, "Statementdate")  %>: 
						<input onBlur="document.frmParameter.submit();" type="text"  name="requiredtxtDateStatement" value="<%=FmtDateShort(cSttDate)%>" class="date">
						<script type="text/javascript">
						var cal0 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStatement'});
						cal0.a_tpl.yearscroll = true;
						</script>
						<!--
						&nbsp;<%=xssStr(allHotWords(79))%>: 
						<input onBlur="document.frmParameter.submit();" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
						<script type="text/javascript">
						/*
						var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
						cal2.a_tpl.yearscroll = true;
						*/
						</script>
						-->
						&nbsp;<%=DisplayPhrase(phraseDictionary, "Show") %>:  
						<select name="chkNegOnly">
							<option value="0" <% if request.form("chkNegOnly")="0" then response.write "selected" end if %>><%=DisplayPhrase(phraseDictionary, "Allbalances")  %></option>
							<option value="-1" <% if request.form("chkNegOnly")="-1" then response.write "selected" end if %>><%=DisplayPhrase(phraseDictionary, "Negativebalancesonly")  %></option>
							<option value="1" <% if request.form("chkNegOnly")="1" then response.write "selected" end if %>><%=DisplayPhrase(phraseDictionary, "Positivebalancesonly")  %></option>
							<!--JM-bug#321-->
							<option value="2" <% if request.form("chkNegOnly")="2" then response.write "selected" end if %>><%=DisplayPhrase(phraseDictionary, "Zerobalancesonly")  %></option>
						</select>
						
						<%=DisplayPhrase(phraseDictionary, "Showclientsowing")  %>
						<input type="text" size="2" name="txtBalanceMax" value="<% if isNumeric(request.form("txtBalanceMax")) then response.write request.form("txtBalanceMax") end if %>"><br /><!--JM - 55_3182 -->
						&nbsp;<%=DisplayPhrase(phraseDictionary, "Showeventbalancesonly")  %>:
						<input type="checkbox" name="optEventBal" <% if optEventBal="on" then response.write " checked" end if %>>&nbsp;
						
&nbsp;&nbsp;
						&nbsp;
						<%=DisplayPhrase(phraseDictionary, "Nostorebillinginfo")  %>:
						<input type="checkbox" name="optNoBillingInfo" <% if request.form("optNoBillingInfo")="on" then response.write " checked" end if %>>
						&nbsp;<% taggingFilter %>
						&nbsp;<%=DisplayPhrase(phraseDictionary, "Makeprinterfriendly")  %>
						<input name="chkPrinterFriendly" type="checkbox" <% if request.form("chkPrinterFriendly")="on" then response.write " checked" end if %>> <!--JM-52_2729-->&nbsp; <%=DisplayPhrase(phraseDictionary, "Hidebalancedetails")  %>
						<input name="OptHideBalDet" type="checkbox" <%if request.form("OptHideBalDet")="on" then %>checked<%end if %> >
						<!--JM-54_2836-->&nbsp; <%= DisplayPhrase(phraseDictionary, "Removebalanceforward") %>
						<input name="OptRemoveBalForward" type="checkbox" <%if request.form("OptRemoveBalForward")="on" then %>checked<%end if %> >&nbsp;<img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" title="Removing the balance forward will only display purchases between the start date and statement date. Otherwise the report will display all purchases between the start and end dates, along with any negative account balance from the past." align="middle">
						<br />
						<input type="button" name="Button" value="<%= getHotWord(226)%>" onClick="genReport();"></b>
						<% exportToExcelButton %>
						</td>
						</tr>
						
						</form>
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" class="mainTextBig" > 
					
					<table class="mainText" width="95%" cellspacing="0" style="margin :0 auto">
						<tr > 
						<td class="mainTextBig" colspan="2" valign="top">
		<% 
		end if			'end of frmExpreport value check before /head line	  
		%>
						
							<% 
		if request.form("frmGenReport")="true" then 
			if request.form("frmExpReport")="true" then
				Dim stFilename
				stFilename="attachment; filename=Account Balance Statement - " & curClientName & " " & Replace(cSDate,"/","-") & " to " & Replace(cSttDate,"/","-") & ".xls" 
				Response.ContentType = "application/vnd.ms-excel" 
				Response.AddHeader "Content-Disposition", stFilename 
			end if
			
			
			strSQL = "SELECT Clients.RSSID, CLIENTS.LastName, CLIENTS.FirstName, CLIENTS.Address, CLIENTS.Address2, CLIENTS.City, CLIENTS.State, CLIENTS.PostalCode, [PAYMENT DATA].ClientID, SUM([PAYMENT DATA].ClientCredit) AS AccountBal, CLIENTS.InsuranceCompany, CLIENTS.InsurancePolicyNum, CLIENTS.Birthdate, CLIENTS.HomeStudio, Location.Address AS LocAdd, Location.AddressLn2 AS LocAdd2, Location.Phone AS LocPhone "
			strSQL = strSQL & "FROM [PAYMENT DATA] INNER JOIN CLIENTS ON [PAYMENT DATA].ClientID = CLIENTS.ClientID LEFT OUTER JOIN Location ON Location.LocationID = CLIENTS.HomeStudio LEFT OUTER JOIN tblCCNumbers ON CLIENTS.ClientID = tblCCNumbers.ClientID "
			
			if request.form("optFilterTagged")="on" then
				strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
				if session("mVarUserID")<>"" then
					strSQL = strSQL & " AND smodeID = " & session("mVarUserID")
				end if
				strSQL = strSQL & " ) "
			end if
					
			strSQL = strSQL & "WHERE " 
			if request.form("optLocation")<>"0" and request.form("optLocation")<>"" then
				strSQL = strSQL & " (Clients.HomeStudio=" & request.form("optLocation") & " OR CLIENTS.HomeStudio=0) AND "
			end if
			strSQL = strSQL & " ([PAYMENT DATA].PaymentDate <= " & DateSep &  cSttDate & DateSep & ") AND ([PAYMENT DATA].ClientID <> 1) "
			'JM - BUG#756
			if optEventBal="on" then
				strSQL = strSQL & " AND NOT ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
			else
				strSQL = strSQL & " AND ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
			end if
			strSQL = strSQL & " AND ([PAYMENT DATA].ClientContractID IS NULL OR NOT([PAYMENT DATA].DepositReleaseDate IS NULL)) AND ([PAYMENT DATA].Returned = 0) "
			'AND ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) AND ([PAYMENT DATA].ClientContractID IS NULL OR NOT([PAYMENT DATA].DepositReleaseDate IS NULL)) AND ([PAYMENT DATA].Returned = 0) "
			'JM-54_2836
			if request.form("OptRemoveBalForward")="on" then 
		        if cSDate<>"" then
                    strSQL = strSQL & " AND PaymentDate>=" & DateSep & cSDate & DateSep
                end if
            end if    
			if request.form("optNoBillingInfo")="on" then
				strSQL = strSQL & "AND (CreditCardNo IS NULL) AND (ACHAccountNum IS NULL) "
			end if
			strSQL = strSQL & "GROUP BY Clients.RSSID, CLIENTS.LastName, CLIENTS.FirstName, [PAYMENT DATA].ClientID, CLIENTS.Address, CLIENTS.Address2, CLIENTS.City, CLIENTS.State, CLIENTS.PostalCode, CLIENTS.InsuranceCompany, CLIENTS.InsurancePolicyNum, CLIENTS.Birthdate, CLIENTS.HomeStudio, Location.Address, Location.AddressLn2, Location.Phone "
			if request.form("chkNegOnly")="0" then
				strSQL = strSQL & "HAVING SUM([PAYMENT DATA].ClientCredit)<>0 " 
			elseif request.form("chkNegOnly")="-1" then
				strSQL = strSQL & "HAVING SUM([PAYMENT DATA].ClientCredit)<0 " 
			elseif request.form("chkNegOnly")="1" then
				strSQL = strSQL & "HAVING SUM([PAYMENT DATA].ClientCredit)>0 " 
			'JM-Bug#321
			elseif request.form("chkNegOnly")="2" then
				strSQL = strSQL & "HAVING SUM([PAYMENT DATA].ClientCredit)=0 " 
			end if
			if isNumeric(request.form("txtBalanceMax")) then
				strSQL = strSQL & " AND SUM([PAYMENT DATA].ClientCredit)<-" & request.form("txtBalanceMax") & " "
			end if
			
			if request.form("frmTagClients")<>"true" then
				if request.form("optSortBy")="0" then
					strSQL = strSQL & "ORDER BY CLIENTS.LastName, SUM([PAYMENT DATA].ClientCredit) "
				else
					strSQL = strSQL & "ORDER BY SUM([PAYMENT DATA].ClientCredit), CLIENTS.LastName "
				end if
			end if
			
		response.write debugSQL(strSQL, "SQL")
			
			rsClient.CursorLocation = 3
			rsClient.open strSQL, cnWS
			Set rsClient.ActiveConnection = Nothing
			cltcount = 0
			do while not rsClient.EOF
			cltcount = cltcount + 1%>
			<table  class="mainText" width="<% if request.form("chkPrinterFriendly")="on" then response.write "90%" else response.write "100%" end if%>" cellspacing="0">
			<%if request.form("chkPrinterFriendly") = "on" then
					%>
						
						<!--<tr>
							<td>
							<p>
								<table class="mainText" width="100%" border="1" cellspacing="0">-->
									
					<%
					end if
				clientID = rsClient("ClientID")
				curClientName = UCASE(TRIM(rsClient("FirstName"))) & " " & UCASE(TRIM(rsClient("LastName")))
				if NOT isNULL(rsClient("Address")) then
					pCltAddrLine1 = rsClient("Address") & "<br />"
				else
					pCltAddrLine1 = ""	
				end if
                 if checkStudioSetting("tblGenOpts","ReceiptShowCltInfo") AND NOT isNULL(rsClient("Address2")) then 
                    pCltAddrLine1 = pCltAddrLine1 & rsClient("Address2") & "<br />"
                 end if
				if isNULL(rsClient("City")) then
					pCltAddrLine2 = ""
				else
					pCltAddrLine2 = rsClient("City") & ", " & rsClient("State") & " " & rsClient("PostalCode") & "<br />"
				end if
				if rsClient("HomeStudio")<>0 AND rsClient("HomeStudio")<>"" AND NOT IsNull(rsClient("HomeStudio")) then
			        studioAddress = rsClient("LocAdd") & "<br />"
			        if NOT isNULL(rsClient("LocAdd2")) then
				        studioAddress = studioAddress & rsClient("LocAdd2") & "<br />"
			        end if
			        pStudioPhone = rsClient("LocPhone")
		        end if
				pCltInsComp = rsClient("InsuranceCompany")
				pCltInsPolicy = rsClient("InsurancePolicyNum")
				pCltBirthdate = rsClient("Birthdate")
				
					BeginBal = 0
					Balance = 0
					BalByCategory = 0
					AmountByCategory = 0
					curBalance = getAccountBalance(clientID, "","")
					'JM-54_2836
					if request.form("OptRemoveBalForward")<>"on" then 
					    BeginBal = (getAccountBalance(clientID, DateAdd("d", -1, cSDate),"")) * (-1)
					    Balance = BeginBal
					end if
					'if Balance < 0 then
					'	isInvoice = false
					'else
					'	isInvoice = true
					'end if 

					if curBalance < 0 then
						isInvoice = true
					else
						isInvoice = false
					end if 
					
					'Recordset for the statement line items
					strSQL = "SELECT DISTINCT [PAYMENT DATA].PmtRefNo, [PAYMENT DATA].ClientID, [PAYMENT DATA].PaymentDate, [PAYMENT DATA].TypePurch, " 
					strSQL = strSQL & "[PAYMENT DATA].PaymentMethod, [PAYMENT DATA].ClientCredit, [PAYMENT DATA].SaleID, LOC.LocationName , LOC.LocationName "
					strSQL = strSQL & "FROM [PAYMENT DATA] LEFT OUTER JOIN ( SELECT DISTINCT Location.LocationName, Sales.SaleID FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN Location ON [Sales Details].Location = Location.LocationID) LOC ON [PAYMENT DATA].SaleID=LOC.SaleID " 
					strSQL = strSQL & " WHERE ([PAYMENT DATA].ClientCredit <> 0) AND ([PAYMENT DATA].ClientID = " & clientID & ") " 
					strSQL = strSQL & "AND ([PAYMENT DATA].PaymentDate >= " & DateSep & cSDate & DateSep & " AND [PAYMENT DATA].PaymentDate <= " & DateSep & cSttDate & DateSep & ") " 
					'strSQL = strSQL & "AND ([PAYMENT DATA].ClassID Is Null OR [PAYMENT DATA].ClassID=0) 
					''JM-55_3182
					if optEventBal ="on" then
						strSQL = strSQL & " AND NOT ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
					else
						strSQL = strSQL & " AND ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
					end if
					strSQL = strSQL & " AND ([PAYMENT DATA].ClientContractID IS NULL OR NOT([PAYMENT DATA].DepositReleaseDate IS NULL)) " 
					strSQL = strSQL & " AND ([PAYMENT DATA].[Returned]=0) " 
					strSQL = strSQL & "ORDER BY [PAYMENT DATA].PaymentDate, [PAYMENT DATA].PmtRefNo"
				response.write debugSQL(strSQL, "SQL")
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing
					
								'''if rsEntry.EOF then			'EOF%>
								
								<%'else
									'rowColor = "#F2F2F2"
									'Balance=(getAccountBalance(clientID, cSttDate,"")) * (-1)	'Final balance for Amount Due at the top
								%>
										<tr>
										<td class="mainText">
											<!--
											<span class="mainTextBig"><b><%=session("StudioName")%></b></span>
											-->
											<!-- #include file="inc_logo.asp" -->											
										</td>
										<td  class="mainText right"><span class="mainTextBig"><b><%=DisplayPhrase(phraseDictionary, "Balancestatement")  %></b></span></td>
										</tr>
										<tr>
										<td class="mainText"><span class="mainTextBig"><b><%if studioAddress<>"" then response.write studioAddress else response.write "[Studio Address Here2]" end if%></b></span></td>
										<td  valign="top" nowrap class="right mainText"><span class="mainTextBig"><b><%=DisplayPhrase(phraseDictionary, "Statementdate")  %>: <%=FmtdateShort(CSttDate)%></b></span></td>
										</tr>
										<tr><td colspan="2" class="mainText"><span class="mainTextBig"><b><%=FmtPhoneNum(pStudioPhone)%></b></span></td></tr>
										<tr><td colspan="2" class="mainText"><span class="mainTextBig"><b><%=pstudioURL%></b></span></td></tr>
										<tr height="30"><td colspan="2">&nbsp;
								<% if request.form("chkPrinterFriendly") = "on" then %>
									<br />
								<% end if %>
										</td></tr>
										<tr><td width="52%" class="mainText"><%=DisplayPhrase(phraseDictionary, "To")  %>: <%=CurClientName%></td>
										<td rowspan="3" class="right">
											<table border="1" cellspacing="0" bordercolor="#666666" class="mainText" width="45%" >
											<tr>
										 
											<td width="55%" nowrap style="background-color:#CCCCCC;"> <% 	if isInvoice then %>&nbsp;<%=DisplayPhrase(phraseDictionary, "Amountdue")  %>: <%	else %>&nbsp;<%= getHotWord(29)%>: <% 	end if %></td>
									<%
									if rsEntry.EOF then
									    Balance = Balance
									'JM-54_2836
									elseif request.form("OptRemoveBalForward")="on" then 
								        strSQL = " SELECT Sum(ClientCredit) AS ClientBalance FROM [PAYMENT DATA] WHERE ClientID=" & clientID & " AND ClientCredit<>0 AND Returned=0 "
								        'AND ((ClassID IS NULL) OR ClassID=0) 
								        ''JM-55_3182
					                    if optEventBal ="on" then
						                    strSQL = strSQL & " AND NOT ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
					                    else
						                    strSQL = strSQL & " AND ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
					                    end if
								        strSQL = strSQL & " AND ([PAYMENT DATA].ClientContractID IS NULL OR NOT([PAYMENT DATA].DepositReleaseDate IS NULL)) "
		                                if cSDate<>"" then
			                                strSQL = strSQL & " AND PaymentDate>=" & DateSep & cSDate & DateSep
		                                end if
		                                if cSttDate<>"" then
			                                strSQL = strSQL & " AND PaymentDate<=" & DateSep & cSttDate & DateSep
		                                end if
		                               'response.write debugSQL(strSQL, "SQL")
		                                rsEntry2.CursorLocation = 3
								        rsEntry2.open strSQL, cnWS
								        Set rsEntry2.ActiveConnection = Nothing
								        if NOT rsEntry2.EOF then
								            Balance = rsEntry2("ClientBalance")
								        end if
								        rsEntry2.close
									else
									    if optEventBal ="on" then 
							                strSQL = "SELECT Sum(ClientCredit) AS ClientBalance FROM [PAYMENT DATA] WHERE "
							                strSQL = strSQL & " ([PAYMENT DATA].ClientID = " & clientID & ") " 
							                strSQL = strSQL & "AND ([PAYMENT DATA].PaymentDate >= " & DateSep & cSDate & DateSep & " AND [PAYMENT DATA].PaymentDate <= " & DateSep & cSttDate & DateSep & ") " 
							                strSQL = strSQL & " AND NOT ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
							                strSQL = strSQL & " AND ([PAYMENT DATA].ClientContractID IS NULL OR NOT([PAYMENT DATA].DepositReleaseDate IS NULL)) " 
							                strSQL = strSQL & "AND ([PAYMENT DATA].[Returned]=0) AND ClientCredit<>0   " 

							                'response.write debugSQL(strSQL, "SQL")
	                                        rsEntry2.CursorLocation = 3
							                rsEntry2.open strSQL, cnWS
							                Set rsEntry2.ActiveConnection = Nothing
							                if NOT rsEntry2.EOF then
							                    Balance = rsEntry2("ClientBalance")
							                end if
							                rsEntry2.close
								        else
										    Balance=(getAccountBalance(clientID, cSttDate,"")) * (-1)	'Final balance for Amount Due at the top
										end if
									end if%>
											<td width="109"><div class="right"><% if NOT request.form("frmExpReport")="true" then %><%=FmtCurrency(abs(Balance))%><% else %><%=(abs(Balance))%><% end if %></div></td>

											</tr>
											<tr>
											<td style="background-color:#CCCCCC;"> &nbsp;<%=DisplayPhrase(phraseDictionary, "Enclosed")  %>:</td>
											<td>&nbsp;</td>
											</tr>
										  </table>
										</td>
										</tr>
										<tr><td class="mainText"><%=pCltAddrLine1%></td></tr>
										<tr><td class="mainText"><%=pCltAddrLine2%></td></tr>
									<% if ss_ShowInsurance AND (pCltInsComp<>"" OR pCltInsPolicy<>"") then %>
										<tr><td colspan="2" class="mainText">
											<%= getHotWord(122)%>:&nbsp;<%=pCltInsComp%><br />
											<%= getHotWord(123)%>:&nbsp;<%=pCltInsPolicy%><br />
										<% if true=false then 'CB: 45_2335 removed per HIPAA%>
											<%=DisplayPhrase(phraseDictionary, "Birthdate")  %>: <%=pCltBirthdate%>
										<% end if %>
										</td></tr>
									<% end if %>
										<tr height="30"><td colspan="2">&nbsp;</td></tr>
<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="1">
											<td colspan="2" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
<% end if %>
										<tr height="20">
											<td colspan="2">&nbsp;</td>
										</tr>
										<tr height="30">
											<td colspan="2"><strong><%= getHotWord(161)%>&nbsp;<%=FmtDateShort(cSDate)%>&nbsp;<%= getHotWord(162)%>&nbsp;<%=FmtDateShort(cSttDate)%></strong></td>
										</tr>
										<tr>
										
									<%if	ss_IncPurchActInStmt then'JM-51_2782%>
									<td width="52%" valign="top">
									<table  class="mainText" width="100%">
									  <tr class="whiteHeader" style="background-color:#666666;">
									  <td width="10%"><div align="left"><strong>&nbsp;<%= getHotWord(57)%></strong></div></td>
									  <td width="50%" align="left"><strong>&nbsp;&nbsp;&nbsp;&nbsp;<%= getHotWord(164)%> </strong></td>
									  <td width="5%"><div class="center-ch"><strong><%= getHotWord(70)%> </strong></div></td>
									  <!-- <td width="10%"><div class="right"><strong>Price </strong></div></td>  JM_4/16/2009 hidden as per Aspen's Req %>-->
									  <td width="10%"><div class="right"><strong><%= getHotWord(35)%> </strong></div></td>
									  <td width="20%"><div class="center-ch"><strong><%= getHotWord(59)%> </strong></div></td>
									 	
									 <!-- MB bug#3233 added extra empty column to Purchases(left) table -->
									 <td width="5%"><div class="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div></td>
									 	
										</tr>
<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="2">
											  <td colspan="6" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
										
										<!-- MB bug#3233 added empty row to Purchases(left) table -->
										<tr>
										 <td colspan="6" class="center-ch"> &nbsp;</td>
										</tr>
<% end if %>
										<%strSQL ="SELECT DISTINCT Sales.SaleDate, Location.LocationName, Sales.SaleID, Sales.SaleDate, [Sales Details].Returned, CASE WHEN ([Returns].ID IS NULL) THEN NULL ELSE 1 END AS ID, [Sales Details].Description, [Sales Details].UnitPrice, [Sales Details].Quantity, (CASE WHEN (NOT ([Sales Details].SDPaymentAmtB) Is Null) THEN [Sales Details].SDPaymentAmtB ELSE 0 END + CASE WHEN (NOT ([Sales Details].SDPaymentAmt) Is Null) THEN [Sales Details].SDPaymentAmt ELSE 0 END) as PaymentTotal FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN Location ON [Sales Details].Location = Location.LocationID LEFT OUTER JOIN [Returns] ON [Sales Details].SaleID = [Returns].SaleID WHERE Sales.ClientID = " & clientID &  " AND Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cSttDate & DateSep &" ORDER BY Sales.SaleDate"
										rsEntry2.CursorLocation = 3
										rsEntry2.open strSQL, cnWS
										Set rsEntry2.ActiveConnection = Nothing

											If NOT rsEntry2.EOF then
											pPurchLineCount = 0
												do while NOT rsEntry2.EOF
												%>
											<tr height="20" style="background-color:<%=rowColor%>;">
											 <td width="10%"><%=FmtDateShort(rsEntry2("SaleDate"))%></td>
										  <td width="50%"><div align="left">&nbsp;&nbsp;&nbsp;
											 	<%if rsEntry2("Returned") then
													response.write "<i><span style=""color:#990000;"">RETURNED: </span></i>"
													pPurchLineCount = pPurchLineCount + 1
												elseif NOT isNULL(rsEntry2("ID")) then 
													response.write "<span style=""color:#990000;"">RETURN SALE:  "
													pPurchLineCount = pPurchLineCount + 1
												end if%>
												<%=rsEntry2("Description")%>&nbsp;</div></td>
												<%if len(rsEntry2("Description")) > 15 then%>
												    	<%pPurchLineCount = pPurchLineCount + 2%>
													<%else %>
												    	<%pPurchLineCount = pPurchLineCount + 1%>
													<%end if %>
											 <td width="5%"><div class="center-ch"><%=rsEntry2("Quantity")%></div></td>
											 <!-- <td width="10%"><div class="right"> JM_4/16/2009 hidden as per Aspen's Req 
											 <%=FmtCurrency(rsEntry2("UnitPrice"))%>
											 <%if rsEntry2("UnitPrice")>0 then response.write "&nbsp;" end if%></div></td> -->
											 <td width="10%"><div class="right">
											 <%=FmtCurrency(rsEntry2("PaymentTotal"))%>
											 <%if rsEntry2("PaymentTotal")>0 then response.write "&nbsp;" end if%></div></td>
											  <td width="20%"><div class="center-ch">&nbsp;
											<%		firstPM = true
													strSQL = " SELECT [Item#], PmtTypes, PaymentNotes FROM [Payment Types] INNER JOIN tblPayments ON tblPayments.PaymentMethod = [Payment Types].[Item#] WHERE tblPayments.SaleID = " & rsEntry2("SaleID")
												response.write debugSQL(strSQL, "SQL")
													rsEntry3.CursorLocation = 3
													rsEntry3.open strSQL, cnWS
													Set rsEntry3.ActiveConnection = Nothing
													if NOT rsEntry3.EOF then
														do while NOT rsEntry3.EOF
															if firstPM then
																firstPM = false
															else
																response.Write "&nbsp;|&nbsp;"
															end if
															response.Write rsEntry3("PmtTypes")
															rsEntry3.MoveNext
														loop
													end if
													rsEntry3.close %>
												</div></td>
										  </tr>
											<% rsEntry2.movenext
												loop
												else%>
											<tr height="20" style="background-color:<%=rowColor%>;">
											 <td colspan="5" class="center-ch"><%=DisplayPhrase(phraseDictionary, "Nopurchaseactivity")  %></td></tr>
											<%end if
											
											rsEntry2.close %>

</table>
											</td>
										
										
										<%end if'JM-51_2782 %>
										
										
											<td <%if not ss_IncPurchActInStmt then%>width="100%"<%else%>width="48%"<%end if%> valign="top"   <%if	not ss_IncPurchActInStmt then%>colspan="2"<%end if%>>
											<table  class="mainText" width="100%" style="margin :0 auto">
										<tr class="whiteHeader" style="background-color:#666666;">
										<td width="15%"><div align="left"><strong>&nbsp;<%= getHotWord(57)%></strong></div></td>
										  
										  <td width="55%" align="left"><strong><%=DisplayPhrase(phraseDictionary, "Balanceactivity")  %> </strong></td>
										    <%if NOT ss_IncPurchActInStmt then'JM_4/16/2009 hidden as per Aspen's Req %>
										  <td width="15%"><div class="right"><strong><%= getHotWord(35)%>&nbsp;</strong></div></td>
										  <%end if%>
										  <td width="15%"><div class="right"><strong><%= getHotWord(29)%>&nbsp;</strong></div></td>
									 	
										</tr>
<% if NOT request.form("frmExpReport")="true" then %>
										<tr height="2">
											<!--  <td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>-->
										 <td <%if not ss_IncPurchActInStmt then%> colspan="4" <%else%> colspan="3"<%end if%> style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
										
										</tr>
<% end if %>

                                <%'JM-54_2836
                                    if request.form("OptRemoveBalForward")<>"on" then%>
										<tr height="20" style="background-color:<%=rowColor%>;">
										  <td width="15%"><div align="left"><%=FmtDateShort(DateAdd("d",-1, DateValue(cSDate)))%></div></td>
										  <td width="55%"><%=DisplayPhrase(phraseDictionary, "Balanceforward")  %> </td>
										   <%if NOT ss_IncPurchActInStmt then'JM_4/16/2009 hidden as per Aspen's Req %>
										  <td width="15%"><div class="right">&nbsp;</div></td>
										  <%end if%>

										  <td width="15%"><div class="right">
										 <% if NOT request.form("frmExpReport")="true" then %>
										  <% if isInvoice then response.write FmtCurrency(BeginBal) else response.write FmtCurrency(BeginBal*-1) end if%>
										  <% else %>
										  <% if isInvoice then response.write (BeginBal) else response.write (BeginBal*-1) end if%>
										  <% end if %>
										  </div></td>

										 

										
										
										</tr>
									<%end if ' request.form("OptRemoveBalForward")<>"on" %>
								<%if NOT rsEntry.EOF then %>
										
							<%'JM-52_2729 %>
                            <% if request.form("OptHideBalDet")="on" then %>
                            <%
                            strSQL =" SELECT DISTINCT Cat.CategoryID, Cat.CategoryName, SUM(Cat.PaymentTotal) as PaymentTotal "
                            strSQL = strSQL & " FROM  [PAYMENT DATA] INNER JOIN (SELECT Sales.SaleID, Categories.CategoryName, "
                            strSQL = strSQL & " [Sales Details].CategoryID, tblSDPayments.SDPaymentAmount AS PaymentTotal "
                            strSQL = strSQL & " FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID "
                            strSQL = strSQL & " INNER JOIN tblPayments ON tblPayments.SaleID = Sales.SaleID "
                            strSQL = strSQL & " INNER JOIN tblSDPayments ON tblSDPayments.PaymentID = tblPayments.PaymentID AND tblSDPayments.SDID = [Sales Details].SDID "
                            strSQL = strSQL & " INNER JOIN Categories on Categories.CategoryID=[Sales Details].CategoryID "
                            strSQL = strSQL & " WHERE Sales.ClientID = " & clientID &  " AND  Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND "
                            strSQL = strSQL & " Sales.SaleDate <= " & DateSep & cSttDate & DateSep &"  ) Cat ON Cat.saleID=[PAYMENT DATA].SaleID "
                            strSQL = strSQL & " WHERE ([PAYMENT DATA].ClientCredit < 0) AND ([PAYMENT DATA].ClientID = " & clientID &  ") AND "
                            strSQL = strSQL & " ([PAYMENT DATA].PaymentDate >= " & DateSep & cSDate & DateSep & " AND "
                            strSQL = strSQL & " [PAYMENT DATA].PaymentDate <= " & DateSep & cSttDate & DateSep &")  "
                           ' strSQL = strSQL & " ([PAYMENT DATA].ClassID Is Null OR [PAYMENT DATA].ClassID=0) AND "
                            ''JM-55_3182
							if optEventBal ="on" then
								strSQL = strSQL & " AND NOT ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
							else
								strSQL = strSQL & " AND ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
							end if
                            strSQL = strSQL & " AND ([PAYMENT DATA].ClientContractID IS NULL OR  NOT([PAYMENT DATA].DepositReleaseDate IS NULL)) AND "
                            strSQL = strSQL & " ([PAYMENT DATA].[Returned]=0) GROUP BY Cat.CategoryID, Cat.CategoryName"
                           response.write debugSQL(strSQL, "SQL")
                            rsEntry2.CursorLocation = 3
                            rsEntry2.open strSQL, cnWS
                            Set rsEntry2.ActiveConnection = Nothing
                                BalByCategory = BeginBal
                                If NOT rsEntry2.EOF then
                                    pActLineCount = 0
                                    
                                    do while NOT rsEntry2.EOF
                                    pActLineCount = pActLineCount + 1
                                    AmountByCategory = rsEntry2("PaymentTotal")  
                                    BalByCategory = BalByCategory +  AmountByCategory
                                    'BalByCategory = BalByCategory + rsEntry2("PaymentTotal")%>
                                    <tr height="20" style="background-color:<%=rowColor%>;">
                                        <td valign="top" width="15%">
                                            <div align="left">
                                                &nbsp;</div>
                                        </td>
                                        <td width="55%">
                                            <%=rsEntry2("CategoryName") %>
                                        </td>
                                        <%if NOT ss_IncPurchActInStmt then%>
                                        <td width="15%">
                                            <div class="right">
                                         <% if NOT request.form("frmExpReport")="true" then %>
										  	<% if isInvoice then response.write FmtCurrency(AmountByCategory) else response.write FmtCurrency(AmountByCategory*-1) end if%>
										<% else %>	
											<% if isInvoice then response.write (AmountByCategory) else response.write (AmountByCategory*-1) end if%>
										<%end if%>
										</div>
                                        </td>
                                        <%end if%>
                                        <td width="15%">
                                            <div class="right">
                                               <% if NOT request.form("frmExpReport")="true" then %>
												<% if isInvoice then response.write FmtCurrency(BalByCategory) else response.write FmtCurrency(BalByCategory*-1) end if%>
											<% else %>	
												<% if isInvoice then response.write (BalByCategory) else response.write (BalByCategory*-1) end if%>
											<%end if%> 
                                            </div>
                                        </td>
                                    </tr>
                                    <% rsEntry2.movenext
                                    loop
                                    %>
                                    <%end if
                                
                                rsEntry2.close %>
                                
                                  <%
                                  
                                  strSQL =" SELECT SUM([PAYMENT DATA].ClientCredit) AS AccountCredit FROM  [PAYMENT DATA] WHERE ([PAYMENT DATA].ClientCredit > 0) AND  ([PAYMENT DATA].ClientID =  " & clientID &  ") AND  ([PAYMENT DATA].PaymentDate >= " & DateSep & cSDate & DateSep & " AND [PAYMENT DATA].PaymentDate <= " & DateSep & cSttDate & DateSep &" ) "
                                  
                                  'AND  ([PAYMENT DATA].ClassID Is Null OR [PAYMENT DATA].ClassID=0) "
                                  ''JM-55_3182
							    if optEventBal ="on" then
								    strSQL = strSQL & " AND NOT ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
							    else
								    strSQL = strSQL & " AND ([PAYMENT DATA].ClassID IS NULL OR [PAYMENT DATA].ClassID = 0) "
							    end if
                                 strSQL = strSQL & " AND ([PAYMENT DATA].ClientContractID IS NULL OR  NOT([PAYMENT DATA].DepositReleaseDate IS NULL)) AND ([PAYMENT DATA].[Returned]=0)  "
                              
                           response.write debugSQL(strSQL, "SQL")
                            rsEntry2.CursorLocation = 3
                            rsEntry2.open strSQL, cnWS
                            Set rsEntry2.ActiveConnection = Nothing

                                If NOT rsEntry2.EOF then 
                                AcctCredit = rsEntry2("AccountCredit") * (-1) 
                                BalByCategory = BalByCategory +  AcctCredit %>
                                    <tr height="20" style="background-color:<%=rowColor%>;">
                                        <td valign="top" width="15%">
                                            <div align="left">
                                                &nbsp;</div>
                                        </td>
                                        <td width="55%">
                                            <%=DisplayPhrase(phraseDictionary, "Accountcredit")  %>
                                        </td>
                                        <%if NOT ss_IncPurchActInStmt then%>
                                        <td width="15%">
                                            <div class="right">
                                         <% if NOT request.form("frmExpReport")="true" then %>
										  	<% if isInvoice then response.write FmtCurrency(AcctCredit) else response.write FmtCurrency(AcctCredit * -1) end if%>
										<% else %>	
											<% if isInvoice then response.write (AcctCredit) else response.write (AcctCredit * -1) end if%>
										<%end if%>
										</div>
                                        </td>
                                        <%end if%>
                                        <td width="15%">
                                            <div class="right">
                                               <% if NOT request.form("frmExpReport")="true" then %>
												<% if isInvoice then response.write FmtCurrency(BalByCategory) else response.write FmtCurrency(BalByCategory*-1) end if%>
											<% else %>	
												<% if isInvoice then response.write (BalByCategory) else response.write (BalByCategory*-1) end if%>
											<%end if%> 
                                            </div>
                                        </td>
                                    </tr>
                                    <% 
                                    %>
                                    <%end if
                                
                                rsEntry2.close %>
                                
                                <%end if 'request.form("OptHideBalDet")="on"  %>
                                
									<%
									pActLineCount = 0
									Balance=BeginBal		'Starting balance for the activity list
									do while NOT rsEntry.EOF
										if rowColor = "#F2F2F2" then
											rowColor = "#FAFAFA"
										else
											rowColor = "#F2F2F2"
										end if
										
										
										
										Amount = rsEntry("ClientCredit") * (-1)
									
										Balance = Balance + Amount				
										'stLineItem = getHotWord(37) & rsEntry("PmtRefNo")  & ": " &  rsEntry("TypePurch")
										'JM - 45_2329
										'stLineItem = "Sale " 
										'if NOT request.form("frmExpReport")="true" then
											'stLineItem = stLineItem & "<a href=""adm_tlbx_voidedit.asp?saleno=" & rsEntry("SaleID") & """>"
										'end if
										'stLineItem = stLineItem & "#" & rsEntry("SaleID")
										stLineItem = getHotWord(37) 
										if NOT request.form("frmExpReport")="true" then
											stLineItem = stLineItem & "<a href=""adm_tlbx_voidedit.asp?pmtno=" & rsEntry("PmtRefNo") & """>"
										end if
										stLineItem = stLineItem & rsEntry("PmtRefNo")
										if NOT request.form("frmExpReport")="true" then
											stLineItem = stLineItem & "</a>"
										end if
										
										if NOT isNULL(rsEntry("LocationName")) then
											stLineItem = stLineItem & " at " & rsEntry("LocationName") 
										end if										
									%>
									<% if NOT request.form("OptHideBalDet")="on" then %>
										<tr height="20" style="background-color:<%=rowColor%>;">
										  <td valign="top" width="15%"><div align="left" ><%=FmtDateShort(rsEntry("PaymentDate"))%></div></td>

										  <td width="55%"><%=stLineitem%>
										  <%pActLineCount=pActLineCount + 1%>
											<%
											strSQL = "SELECT [Sales Details].SaleID, [Sales Details].Description, [PAYMENT DATA].ClientID as RecipClientID, CLIENTS.FirstName, CLIENTS.LastName, (IsNull(SDPaymentAmt,0) + IsNull(SDPaymentAmtB,0)) AS AmtPaid FROM [Sales Details] LEFT OUTER JOIN [PAYMENT DATA] ON [Sales Details].PmtRefno = [PAYMENT DATA].PmtRefNo LEFT OUTER JOIN CLIENTS ON [PAYMENT DATA].ClientID = CLIENTS.ClientID "
										if isNULL(rsEntry("SaleID")) then
											strSQL = strSQL & " WHERE 1=0"
										else
											strSQL = strSQL & "WHERE [Sales Details].SaleID = " & rsEntry("SaleID") & " "
										end if
										strSQL = strSQL & "ORDER BY [Sales Details].Description" 
										rsEntry2.CursorLocation = 3
										rsEntry2.open strSQL, cnWS
										Set rsEntry2.ActiveConnection = Nothing

										
										stLineItemDetail=""
											If NOT rsEntry2.EOF then
												do while NOT rsEntry2.EOF
													if NOT request.form("frmExpReport")="true" then
														stLineItemDetail = rsEntry2("Description") '& " - " & FmtCurrency(rsEntry2("AmtPaid")) 
													else
														stLineItemDetail = rsEntry2("Description") '& " - " & rsEntry2("AmtPaid")
													end if
													' added 6/18/8 to add recipient client name if different than billed client
													if NOT isNull(rsEntry2("RecipClientID")) then 
														if CSTR(rsEntry2("RecipClientID"))<>CSTR(clientID) then
															stLineItemDetail = stLineItemDetail & " for " & rsEntry2("FirstName") & " " & rsEntry2("LastName")
														end if
													end if
													If Amount>0 then
													%>
													<br />&nbsp;&nbsp;&nbsp;<%=stLineitemDetail%>
													<%if len(stLineitemDetail) > 53 then%>
												    	<%pActLineCount=pActLineCount + 2%>
													<%else %>
												    	<%pActLineCount=pActLineCount + 1%>
													<%end if %>
													<%
													end if
													rsEntry2.movenext
												loop
											end if
											rsEntry2.close
											%>
										  </td>

										<%if request.form("OptHideBalDet")="on" OR (NOT ss_IncPurchActInStmt) then'JM_4/16/2009 hidden as per Aspen's Req %>
										  <td width="15%"><div class="right">
										  <% if NOT request.form("frmExpReport")="true" then %>
										  	<% if isInvoice then response.write FmtCurrency(Amount) else response.write FmtCurrency(Amount*-1) end if%>
										<% else %>	
											<% if isInvoice then response.write (Amount) else response.write (Amount*-1) end if%>
										<%end if%>
											</div></td>
											<%end if%>
										  <td width="15%"><div class="right">
										    <% if NOT request.form("frmExpReport")="true" then %>
												<% if isInvoice then response.write FmtCurrency(Balance) else response.write FmtCurrency(Balance*-1) end if%>
											<% else %>	
												<% if isInvoice then response.write (Balance) else response.write (Balance*-1) end if%>
											<%end if%>
												</div></td>

										
										</tr>
										 <%end if 'request.form("OptHideBalDet")="on" then %>
									<%
									
										rsEntry.MoveNext
										
									loop
									%>
									<%end if 'NOT EOF%>
									</table>
											</td>
											
										</tr>
<% if NOT request.form("frmExpReport")="true" then %>


										<tr height="2">
											  <td colspan="2" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
<% end if %>
										<tr height="20">
										
										 
										  <td colspan="2"><div class="right"><strong> <% 	if isInvoice then %><%=DisplayPhrase(phraseDictionary, "Amountdue")  %>:<%	else %><%= getHotWord(29)%>:<% 	end if %>&nbsp;&nbsp;&nbsp;&nbsp;</strong>
										  <% if NOT request.form("frmExpReport")="true" then %>
											 <strong><%=FmtCurrency(abs(Balance))%></strong>
										  <% else %>
											  <strong><%=(abs(Balance))%></strong>
										  <% end if %>
										  </div>
										  </td>
										

										</tr>
                           				<% if ss_ShowReceiptMessage then %>
										  <% if NOT request.form("frmExpReport")="true" then %>
										    <tr><td colspan="2" ><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="50" width="100%"></td></tr>
											 <tr height="2"><td colspan="2" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td></tr>
										  <% end if %>
										  <tr height="40"><td colspan="2"><%=receiptMsg1%></td></tr>
										  <tr height="20"><td colspan="2"><%=receiptMsg2%></td></tr>
                                          <tr><td colspan="2">&nbsp;</td></tr>
										<% end if %>
									<%
									
								

			''' *************** PAYMENT COUPON ********************
%>
<%if (pActLineCount > 20 and pActLineCount < 30) OR (pPurchLineCount > 20 and pPurchLineCount < 30) and request.form("chkPrinterFriendly")="on" then%>
<tr><td class="center-ch" colspan="2"><div style="page-break-before:always">&nbsp;</div></td></tr><%end if%>
						<tr><td class="center-ch" colspan="2"><%=DisplayPhrase(phraseDictionary, "Pleasetearoff")  %></td></tr>
						<tr><td colspan="2"><table style="border-top: 2px dashed #666666;" width="100%"><tr><td>&nbsp;</td></tr></table></td></tr>
						<tr>
							<td colspan="2"><b><%=DisplayPhrase(phraseDictionary, "Paymnetcoupon")  %></b></td>
						</tr>
						<tr>
							<td width="52%" class="mainText"><%=CurClientName%></td>
							<td rowspan="3" class="right">
								<table border="1" cellspacing="0" bordercolor="#666666" class="mainText" width="45%" >
								<tr>
									<td style="background-color:#CCCCCC;" width="60%" > &nbsp;<%=DisplayPhrase(phraseDictionary, "Statementdate")  %>:</td>
									<td class="right"><%=FmtdateShort(CSttDate)%></td>
								</tr>
								<tr>
								  
									<td width="112" nowrap style="background-color:#CCCCCC;"><% 	if isInvoice then %>&nbsp;<%=DisplayPhrase(phraseDictionary, "Amountdue")  %>:<%	else %>&nbsp;<%= getHotWord(29)%>: <% 	end if %></td>
	
									<td width="109"><div class="right"><% if NOT request.form("frmExpReport")="true" then %><%=FmtCurrency(abs(Balance))%><% else %><%=(abs(Balance))%><% end if %></div></td>
	
								</tr>
								<tr>
									<td style="background-color:#CCCCCC;"> &nbsp;<%=DisplayPhrase(phraseDictionary, "Enclosed")  %>:</td>
									<td>&nbsp;</td>
								</tr>
							  </table>
							</td>
						</tr>
						<tr><td class="mainText"><%=pCltAddrLine1%></td></tr>
						<tr><td class="mainText"><%=pCltAddrLine2%></td></tr>
					<% if ss_ShowInsurance AND (pCltInsComp<>"" OR pCltInsPolicy<>"") then %>
						<tr><td colspan="2" class="mainText">
							<%= getHotWord(122)%>:&nbsp;<%=pCltInsComp%><br />
							<%= getHotWord(123)%>:&nbsp;<%=pCltInsPolicy%><br />
						<% if true=false then 'CB: 45_2335 removed per HIPAA%>
							<%=DisplayPhrase(phraseDictionary, "Birthdate")  %>: <%=pCltBirthdate%>
						<% end if %>
						</td></tr>
					<% end if %>
					<%if ss_InvoiceShowCC then 
					    Dim ss_Amex, ss_Visa, ss_Mastercard, ss_Discover '55_3292, only show accepted methods, CCP 10/6/09
					    ss_Amex = checkStudioSetting("tblCCOpts","ccAmericanExpress")
					    ss_Visa = checkStudioSetting("tblCCOpts","ccVisa")
					    ss_Mastercard = checkStudioSetting("tblCCOpts","ccMasterCard")
					    ss_Discover = checkStudioSetting("tblCCOpts","ccDiscover")
					%>
						<tr>
							<td >
								<strong><%=DisplayPhrase(phraseDictionary, "Topaybycreditcard")  %>:</strong><br />
								<table style="background-color:#CCCCCC;"  class="mainText">
									<tr><td><%= getHotWord(146)%>:&nbsp;_______________________________</td></tr>
									<tr><td><%= getHotWord(46)%>:&nbsp;______________________________________</td></tr>
									<tr><td><%= getHotWord(47)%>:&nbsp;______________________________</td></tr>
									<tr><td><%= getHotWord(48)%>:&nbsp;________________________</td></tr>
									<tr><td><%= getHotWord(49)%>:&nbsp;________________________</td></tr>
									<tr><td height="4"></td></tr>
									<tr><td><%= getHotWord(50)%>&nbsp;<span style="font-size:9px;"><%= DisplayPhrase(phraseDictionary, "Checkone") %></span>:&nbsp;&nbsp;<% if ss_Amex then%>__AMEX&nbsp;&nbsp;<%end if %><% if ss_Visa then%>__VISA&nbsp;&nbsp;<%end if %><% if ss_Mastercard then%>__MC&nbsp;&nbsp;<%end if %><% if ss_Discover then%>__DISC<%end if %></td></tr>
									<tr><td><%= getHotWord(51)%>&nbsp;:&nbsp;_________________________________</td></tr>
									<tr><td><%= getHotWord(52)%>&nbsp;<span style="font-size:9px;">(MM/YYYY)</span>:&nbsp;______________________</td></tr>
								</table>
							</td>
							<td class="right">
								<b><%=DisplayPhrase(phraseDictionary, "Topaybycheck")  %>: <u><%=session("StudioName")%></u></b><br /><br />
								<%=DisplayPhrase(phraseDictionary, "Iauthorize")  %><br /><span class="textSmall"><%=DisplayPhrase(phraseDictionary, "Areceipt")  %></span><br /><br />
								<strong>X</strong>_____________________________________
							</td>
						</tr>
						<%end if 'ss_InvoiceShowCC%>

						<%if ss_InvoiceAskAutoPay then %>
						<tr >
							<td colspan="2"><strong>
								<input type="checkbox" name="optIAgree"> 
								<%=DisplayPhrase(phraseDictionary, "Pleasecontactme")  %>	<%= getHotWord(93)%>: _______________
								</strong>
							</td>
						</tr>
						<% end if 'ss_InvoiceAskAutoPay%>				
					
					<%rsEntry.close%>
							 </table>	<%'if cltcount > 1 then%><div style="page-break-after:always">&nbsp;</div><%'end if%>
				<%rsClient.moveNext
				loop
				
			end if		'end of generate report if statement
							%>
						 
						  </td>
							</tr>
						</table>
						</td>
						</tr>
						
					</table>
				</td>
				</tr>
				
			  	</table>
				
<!--% pageEndShowFooter(request.form("chkPrinterFriendly")<>"on") %-->
<% pageEnd %>
<!-- #include file="post.asp" -->

<%
	
end if
%>
