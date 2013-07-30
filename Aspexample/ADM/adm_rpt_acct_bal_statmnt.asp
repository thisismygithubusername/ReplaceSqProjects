<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<!-- #include file="../json/JSON.asp" -->
<%
Server.ScriptTimeout = 300    '5 min (value in seconds)
If request.querystring("pdf") = "true" Then
'    if request.form("sid")<>"" then
'        Response.Cookies("SessionFarmGUID") = request.form("sid")
'    end if
end if


'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%

	dim rsEntry, rsEntry2, rsEntry3
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	set rsEntry2 = Server.CreateObject("ADODB.Recordset")
	set rsEntry3 = Server.CreateObject("ADODB.Recordset")
	%>
	<!-- #include file="inc_accpriv.asp" -->
	<!-- #include file="inc_acct_balance.asp" -->
	<!-- #include file="inc_rpt_pdf.asp" -->
        <!-- #include file="inc_rpt_tagging.asp" -->
	<!-- #include file="inc_hotword.asp" -->
	<!-- #include file="../inc_post.asp" -->
	<!-- #include file="../inc_tinymcesetup.asp" -->
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
		Dim showDetails, cSttDate, cSDate, cEDate, tmpCurDate, CurClientName, clientID
		Dim rowColor, Balance, Amount, stLineItem, stLineItemDetail, BeginBal, isInvoice, curBalance, BalByCategory, AmountByCategory, AcctCredit
		Dim studioAddress, pStudioPhone, pPayMethod, contText, pStudioURL, pCltAddrLine1, pCltAddrLine2
		Dim pCltPhone, pRecLn1, pRecLn2, pCltInsComp, pCltInsPolicy, pCltBirthdate, pActLineCount, pPurchLineCount, firstPM

		'--------------------------------------------------
		' Check Studio Settings
		'--------------------------------------------------
		Dim ss_ApplyAccountPaymentsByLocation : ss_ApplyAccountPaymentsByLocation = checkStudioSetting("tblGenOpts", "ApplyAccountPaymentsByLocation")
		Dim ss_ShowInsurance, ss_InvoiceShowCC, ss_InvoiceAskAutoPay, ss_IncPurchActInStmt, ss_ShowReceiptMessage, receiptMsg1, receiptMsg2, optEventBal
				
		strSQL = "SELECT tblGenOpts.InsuranceFields, tblGenOpts.InvoiceShowCC, tblGenOpts.InvoiceAskAutoPay, tblGenOpts.IncPurchActInStmt, tblGenOpts.InvoiceShowReceiptMsg, tblGenOpts.receiptLn1, tblGenOpts.receiptLn2 FROM Studios INNER JOIN tblGenOpts ON Studios.StudioID = tblGenOpts.StudioID WHERE Studios.StudioID=" & session("StudioID")
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
			ss_ShowInsurance = rsEntry("InsuranceFields")
			ss_InvoiceShowCC = rsEntry("InvoiceShowCC")
			ss_InvoiceAskAutoPay = rsEntry("InvoiceAskAutoPay")
			ss_IncPurchActInStmt = rsEntry("IncPurchActInStmt") 'JM-51_2782
			ss_ShowReceiptMessage = rsEntry("InvoiceShowReceiptMsg")
			receiptMsg1 = HtmlPurifyForDisplay(rsEntry("receiptLn1"))
			receiptMsg2 = HtmlPurifyForDisplay(rsEntry("receiptLn2"))
		rsEntry.close
		'--------------------------------------------------
		'JM-55_3182
		if request.QueryString("optEventBal")="on" then
		    optEventBal = request.QueryString("optEventBal")
		elseif request.form("optEventBal")="on" then
		    optEventBal = request.form("optEventBal")
		else
		    optEventBal = ""
		end if

		dim rsLogo
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
		rsLogo.close
%>
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
	
		Dim ap_view_all_locs : ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
		Dim optSaleLocation : optSaleLocation = 0
		Dim tblClientAccountStr : tblClientAccountStr = "tblClientAccount"
		if isNumeric(request.form("optSaleLocation")) then
			optSaleLocation = CInt(request.Form("optSaleLocation"))
			if NOT optSaleLocation=0 then
				dim jsonParams : set jsonParams = JSON.parse("{}")
				jsonParams.set "LocationID",optSaleLocation
				CallMethodWithJSON "mb.Core.BLL.SitesPoco.ClientAccount.GetPerLocationView",jsonParams
			'	dim clientAccountBLL : set clientAccountBLL = Server.CreateObject("mb.Core.BLL.ClientAccountBLLCOM")	
			'	tblClientAccountStr = clientAccountBLL.GetPerLocationView(optSaleLocation)
			end if

		end if
		'if request.form("requiredtxtDateEnd")<>"" then
		'	Call SetLocale(session("mvarLocaleStr"))
			'	cEDate = CDATE(request.form("requiredtxtDateEnd"))
			'Call SetLocale("en-us")
		'else
		'	cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		'end if
		if request.form("optRetAdd")="0" OR NOT isNum(request.form("optRetAdd")) then
		    if session("curLocation")<>"0" then		
			    strSQL = "SELECT Address, AddressLn2, Phone From Location WHERE LocationID=" & session("curLocation")
		    else 
			    strSQL = "SELECT TOP 1 Address, AddressLn2, Phone From Location ORDER BY LocationID"
		    end if
		else 
		    strSQL = "SELECT Address, AddressLn2, Phone From Location WHERE LocationID=" & request.form("optRetAdd")
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
		
		
		CurClientName=""

                if request.querystring("ClientID")<>"" then
                    clientID = CLNG(request.querystring("ClientID"))
                elseif request.form("ClientID")<>"" then
                    clientID = CLNG(request.form("ClientID"))
                else
                    clientID = 0 'This is probably an error case
                end if

		strSQL = "SELECT FirstName, LastName, CLIENTS.Address, Address2, CLIENTS.City, CLIENTS.State, CLIENTS.PostalCode, InsuranceCompany, InsurancePolicyNum, Birthdate, HomeStudio, Location.Address AS LocAdd, Location.AddressLn2 AS LocAdd2, Location.Phone AS LocPhone  FROM CLIENTS LEFT OUTER JOIN Location ON CLIENTS.HomeStudio=Location.LocationID WHERE CLIENTID = " & clientID
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing

		if NOT rsEntry.EOF then
			curClientName = UCASE(TRIM(rsEntry("FirstName"))) & " " & UCASE(TRIM(rsEntry("LastName")))
			if NOT isNULL(rsEntry("Address")) then
				pCltAddrLine1 = rsEntry("Address") & "<br />"
			else
				pCltAddrLine1 = ""	
			end if
             if checkStudioSetting("tblGenOpts","ReceiptShowCltInfo") AND NOT isNULL(rsEntry("Address2")) then 
                    pCltAddrLine1 = pCltAddrLine1 & rsEntry("Address2") & "<br />"
            end if
			if isNULL(rsEntry("City")) then
				pCltAddrLine2 = ""
			else
				pCltAddrLine2 = rsEntry("City") & ", " & rsEntry("State") & " " & rsEntry("PostalCode") & "<br />"
			end if
			pCltInsComp = rsEntry("InsuranceCompany")
			pCltInsPolicy = rsEntry("InsurancePolicyNum")
			pCltBirthdate = rsEntry("Birthdate")
			if request.form("optRetAdd")="0" OR request.form("optRetAdd")="" then
		        if rsEntry("HomeStudio")<>0 AND rsEntry("HomeStudio")<>"" AND NOT IsNull(rsEntry("HomeStudio")) then
		            studioAddress = rsEntry("LocAdd") & "<br />"
		            if NOT isNULL(rsEntry("LocAdd2")) then
			            studioAddress = studioAddress & rsEntry("LocAdd2") & "<br />"
		            end if
		            pStudioPhone = rsEntry("LocPhone")
	            end if
	        end if
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
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode)) %>
			<script type="text/javascript">
			function genReport() {
				document.frmParameter.frmGenReport.value = "true";
				document.frmParameter.frmExpReport.value = "false";
                document.frmParameter.frmGenPdf.value = "false";
				document.frmParameter.action = 'adm_rpt_acct_bal_statmnt.asp?ClientID=<%=clientID%>&optEventBal=<%=optEventBal %>';
				document.frmParameter.submit();
			}
			function exportReport() {
				document.frmParameter.frmExpReport.value = "true";
				document.frmParameter.frmGenReport.value = "true";
                document.frmParameter.frmGenPdf.value = "false";
				<% iframeSubmit "frmParameter", "adm_rpt_acct_bal_statmnt.asp" %>
			}
			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
		<%
		else
		%>
			<base href="http://<%=getLocalhostString()%>"  />
		<%
		end if
		%>		
<script type="text/javascript">
	$(function ()
	{	// add active class to account detail sub link
		$('a[href*="adm_clt_ph.asp"]').parent().addClass("active");
	});

</script>
<style type="text/css">
	#main-content
	{
		margin: 20px 20px 0; 
	}
</style>
<% pageStart %>
	<!-- #include file="inc_sub_links.asp" -->
		<%if (request.form("chkPrinterFriendly") = "on") then  %>	
        <style type="text/css">
         #footer, #setupNav {display:none}
        </style>
        <%end if %>
		<% if request.form("frmExpReport")<>"true" and request.form("chkPrinterFriendly")<>"on" then %>
            
			<table height="100%" width="<%=strPageWidth%>" cellspacing="0" style="margin:0 auto">    
				<tr> 
				<td  valign="top" height="100%" width="100%"> 
				<table cellspacing="0" width="90%" height="100%"style="margin:0 auto" >
					<tr>
					<td class="headText" align="left" valign="top">
					<table width="100%" cellspacing="0">
						<tr>
						<td class="headText" valign="bottom"><b> <%=DisplayPhrase(phraseDictionary, "Balancestatement") %> for <%=curClientname%></b></td>
						<td valign="bottom" height="26"> </td>
						</tr>
					</table>
					</td>
					</tr>
					<tr> 
					<td height="30"  valign="bottom" class="headText">
					<table class="mainText border4 center" cellspacing="0">
						<form name="frmParameter" action="adm_rpt_acct_bal_statmnt.asp?ClientID=<%=clientID%>&optEventBal=<%=optEventBal %>" method="POST">
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
						<!--&nbsp;<%=xssStr(allHotWords(79))%>: 
						<input onBlur="document.frmParameter.submit();" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
						<script type="text/javascript">
						/*
						var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
						cal2.a_tpl.yearscroll = true;
						*/
						</script>-->
						&nbsp;<%= DisplayPhrase(phraseDictionary, "Statementdate")%>: 
						<input onBlur="document.frmParameter.submit();" type="text"  name="requiredtxtDateStatement" value="<%=FmtDateShort(cSttDate)%>" class="date">
						<script type="text/javascript">
						var cal0 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStatement'});
						cal0.a_tpl.yearscroll = true;
						</script>
						<br />
						&nbsp; <%= DisplayPhrase(phraseDictionary, "Makeprinterfriendly")%>
						<input name="chkPrinterFriendly" type="checkbox"><!--JM-52_2729-->&nbsp; <%= DisplayPhrase(phraseDictionary, "Hidebalancedetails")%>
						<input name="OptHideBalDet" type="checkbox" <%if request.form("OptHideBalDet")="on" then %>checked<%end if %> ><!--JM-54_2836-->&nbsp; <%= DisplayPhrase(phraseDictionary, "Removebalanceforward")%>
						<input name="OptRemoveBalForward" type="checkbox" <%if request.form("OptRemoveBalForward")="on" then %>checked<%end if %> >&nbsp;<img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" title="Removing the balance forward will only display purchases between the start date and statement date. Otherwise the report will display all purchases between the start and end dates, along with any negative account balance from the past." align="middle">
						<br />
                        <% if ss_ApplyAccountPaymentsByLocation then %>
                        &nbsp;
                        Sale Location: <select name="optSaleLocation" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
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
					    <script type="text/javascript">
					  	document.frmParameter.optSaleLocation.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' + " Sale <%=jsEscDouble(allHotWords(8))%>s";
					    </script>						
					    <% end if %>
						&nbsp;Return Address:
						<select name="optRetAdd" onChange="genReport();">
						<% if session("numLocations") > 1 then %>
					  	<option value="0" <% if request.form("optRetAdd")="0" then response.write "selected" end if %>><%=xssStr(allHotWords(43))%></option>
					  	<% end if %>
<%

								strSQL = "SELECT LocationID, LocationName from Location WHERE wsShow=1 ORDER BY LocationName " 
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								Set rsEntry.ActiveConnection = Nothing


								do While NOT rsEntry.EOF
%>
                        <option value="<%response.write rsEntry("LocationID")%>" <%if request.form("optRetAdd")=cstr(rsEntry("LocationID")) then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
                        <%
									rsEntry.MoveNext
								loop
								rsEntry.close
%>
                      </select>
						<input type="button" name="Button" value="<%= getHotWord(226)%>" onClick="genReport();"></b>
						<% 
                                                exportToExcelButton 
                                                pdfExportButton "frmParameter", "Balance_Statement_" & Replace(cSDate, "/", "-") & "_to_" & Replace(cEDate, "/", "-") & ".pdf"
                                                %>
                                                <input type="hidden" name="clientID" value="<%=xssStr(clientID)%>" />
                                                 <input type="hidden" name="optEventBal" value="<%=xssStr(optEventBal)%>" />
						</td>
						</tr>
						
						</form>
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" class="mainTextBig" > 
					
					<table class="mainText" width="95%" border="0" cellspacing="0" cellpadding="0" style="margin :0 auto">
						<tr > 
						<td class="mainTextBig" colspan="2" valign="top" >
		<% 
		end if			'end of frmExpreport value check before /head line	  
		%>
						<table  class="mainText" width="<% if request.form("chkPrinterFriendly")="on" then response.write "90%" else response.write "100%" end if%>" border="0" cellpadding="0" cellspacing="0">
							<% 
							if request.form("frmGenReport")="true" then 
								if request.form("frmExpReport")="true" then
									Dim stFilename
									stFilename="attachment; filename=Account Balance Statement - " & curClientName & " " & Replace(cSDate,"/","-") & " to " & Replace(cSttDate,"/","-") & ".xls" 
									Response.ContentType = "application/vnd.ms-excel" 
									Response.AddHeader "Content-Disposition", stFilename 
								end if

								BeginBal = 0
								Balance = 0
								BalByCategory = 0
								AmountByCategory = 0
								curBalance = getAccountBalance(clientID, "","")
								'JM-54_2836
								if request.form("OptRemoveBalForward")<>"on" then 
'//								    BeginBal = (getAccountBalance(clientID, DateAdd("d", -1, cSDate),"")) * (-1)
								    BeginBal = (getAccountBalanceNew(clientID, DateAdd("d", -1, cSDate),"","",optEventBal="on")) * (-1)
								    Balance = BeginBal
								end if
								'response.Write getAccountBalance(clientID, DateAdd("d", -1, cSDate),"")
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
'isInvoice = false ' this view is fine
'isInvoice = true

								'Recordset for the statement line items
								strSQL = "SELECT DISTINCT tblClientAccount.ClientAccountID, tblClientAccount.ClientID, tblClientAccount.EntryDate, " 
          						strSQL = strSQL & "tblClientAccount.Amount, loc.SaleID, LOC.LocationName , LOC.LocationName "
								strSQL = strSQL & "FROM " & tblClientAccountStr & " LEFT OUTER JOIN ("
								strSQL = strSQL & "SELECT DISTINCT Location.LocationName, Sales.SaleID, tblPayments.PaymentID, [Sales Details].SDID FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN tblPayments ON Sales.SaleID = tblPayments.SaleID INNER JOIN Location ON [Sales Details].Location = Location.LocationID"
								strSQL = strSQL & ") LOC ON (tblClientAccount.PaymentID=LOC.PaymentID OR tblClientAccount.SDID=LOC.SDID)" 
								strSQL = strSQL & " WHERE (tblClientAccount.Amount <> 0) AND (tblClientAccount.ClientID = " & clientID & ") " 
								strSQL = strSQL & "AND (tblClientAccount.EntryDate >= " & DateSep & cSDate & DateSep & " AND tblClientAccount.EntryDate < " & DateSep & CDATE(DateAdd("d",1,cSttDate)) & DateSep & ") " 
								'strSQL = strSQL & "AND (tblClientAccount.ClassID Is Null OR tblClientAccount.ClassID=0) "
								''JM-55_3182
								if optEventBal ="on" then
									strSQL = strSQL & " AND NOT (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
								else
									strSQL = strSQL & " AND (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
								end if
								strSQL = strSQL & " AND (tblClientAccount.ClientContractID IS NULL OR NOT(tblClientAccount.DepositReleaseDate IS NULL)) " 
								strSQL = strSQL & " ORDER BY tblClientAccount.EntryDate, tblClientAccount.ClientAccountID"
								logIt strSQL
							response.write debugSQL(strSQL, "SQL")
							'response.Write strSQL
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS
								Set rsEntry.ActiveConnection = Nothing
								if request.form("chkPrinterFriendly") = "on" then
								%>
									<tr>
									<td colspan="2"><br /></td>
									</tr>
								<%
								end if%>
								

								
										<tr>
										<td class="mainText" nowrap>
											<!--
											<span class="mainTextBig"><b><%=session("StudioName")%></b></span>
											-->
											<!-- #include file="inc_logo.asp" -->											
										</td>
										<td  class="mainText" align="right" nowrap><span class="mainTextBig"><b><%= DisplayPhrase(phraseDictionary, "Balancestatement") %></b></span></td>
										</tr>
										<tr>
										<td class="mainText" nowrap><span class="mainTextBig"><b><%if studioAddress<>"" then response.write studioAddress else response.write "[Studio Address Here2]" end if%></b></span></td>
										<td align="right" valign="top" nowrap class="mainText"><span class="mainTextBig"><b><%= DisplayPhrase(phraseDictionary, "Statementdate") %>: <%=FmtdateShort(CSttDate)%></b></span></td>
										</tr>
										<tr><td colspan="2" class="mainText" nowrap><span class="mainTextBig"><b><%=FmtPhoneNum(pStudioPhone)%></b></span></td></tr>
										<tr><td colspan="2" class="mainText" nowrap><span class="mainTextBig"><b><%=pstudioURL%></b></span></td></tr>
										<tr height="30"><td colspan="2">&nbsp;
								<% if request.form("chkPrinterFriendly") = "on" then %>
									<br />
								<% end if %>
										</td></tr>
										<tr><td width="52%" class="mainText"><%= DisplayPhrase(phraseDictionary, "To") %>: <%=CurClientName%></td>
										<td rowspan="3" align="right">
											<table border="1" cellspacing="0" bordercolor="#666666" class="mainText" width="45%">
											<tr>
										 
											
										 
									<%
									if rsEntry.EOF then
									    Balance = Balance
									'JM-54_2836
									elseif request.form("OptRemoveBalForward")="on" then 
								        strSQL = " SELECT Sum(Amount) AS ClientBalance FROM " & tblClientAccountStr & " "
								        strSQL = strSQL & " WHERE ClientID=" & clientID & " AND Amount<>0 "
								        ''JM-55_3182
								        if optEventBal ="on" then
									        strSQL = strSQL & " AND NOT (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
								        else
									        strSQL = strSQL & " AND (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
								        end if
								           ' AND ((ClassID IS NULL) OR ClassID=0) 
							            strSQL = strSQL & " AND (tblClientAccount.ClientContractID IS NULL OR NOT(tblClientAccount.DepositReleaseDate IS NULL)) "
	                                    if cSDate<>"" then
		                                    strSQL = strSQL & " AND EntryDate>=" & DateSep & cSDate & DateSep
	                                    end if
	                                    if cSttDate<>"" then
		                                    strSQL = strSQL & " AND EntryDate<=" & DateSep & cSttDate & DateSep
	                                    end if
		                               'response.write debugSQL(strSQL, "SQL")
		                                rsEntry2.CursorLocation = 3
		                                logIt strSQL
								        rsEntry2.open strSQL, cnWS
								        Set rsEntry2.ActiveConnection = Nothing
								        if NOT rsEntry2.EOF then
								            Balance = rsEntry2("ClientBalance")
								        end if
								        rsEntry2.close
							        else
							            if optEventBal ="on" then 
							                strSQL = "SELECT Sum(Amount) AS ClientBalance FROM " & tblClientAccountStr & " " 
							                strSQL = strSQL & " WHERE "
							                strSQL = strSQL & " (tblClientAccount.ClientID = " & clientID & ") " 
							                strSQL = strSQL & "AND (tblClientAccount.EntryDate >= " & DateSep & cSDate & DateSep & " AND tblClientAccount.EntryDate <= " & DateSep & cSttDate & DateSep & ") " 
							                strSQL = strSQL & " AND NOT (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
							                strSQL = strSQL & " AND (tblClientAccount.ClientContractID IS NULL OR NOT(tblClientAccount.DepositReleaseDate IS NULL)) " 
							                strSQL = strSQL & "AND Amount<>0   " 

							                logIt strSQL
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
									end if
									'if Balance < 0 then
									'	isInvoice = false
									'end if
											%>		
										
											<td width="55%" nowrap style="background-color:#CCCCCC;"> <% 	if isInvoice then %>&nbsp;<%=DisplayPhrase(phraseDictionary, "Amountdue") %>: <%	else %>&nbsp;<%= getHotWord(29)%>:  <% 	end if %></td>

											<td width="109"><div align="right"><%=FmtCurrency(abs(Balance))%></div></td>

										

											</tr>
											<tr>
											<td style="background-color:#CCCCCC;"> &nbsp;<%= DisplayPhrase(phraseDictionary, "Enclosed")%>:</td>
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
											<%= DisplayPhrase(phraseDictionary, "Birthdate")%>: <%=pCltBirthdate%>
										<% end if %>
										</td></tr>
									<% end if %>
										<tr height="30"><td colspan="2">&nbsp;</td></tr>
                                    <% if NOT request.form("frmExpReport")="true" then %>
										<tr height="1">
											<td colspan="2" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
										</tr>
                                    <% end if %>
										<tr height="20">
											<td colspan="2">&nbsp;</td>
										</tr>
										<tr height="30">
											<td colspan="2"><strong><%= getHotWord(161)%>&nbsp;<%=FmtDateShort(cSDate)%>&nbsp;<%= getHotWord(162)%>&nbsp;<%=FmtDateShort(cSttDate)%></strong></td>
										</tr>
										<tr>
										
										<%if ss_IncPurchActInStmt then'JM-51_2782%>
										<td width="52%" valign="top">
									<table  class="mainText" width="100%">
									  <tr class="whiteHeader" style="background-color:#666666;">
									  <td width="10%"><div align="left"><strong>&nbsp;<%= getHotWord(57)%></strong></div></td>
									  <td width="50%" align="left"><strong>&nbsp;&nbsp;&nbsp;&nbsp;<%=xssStr(allHotWords(149))%>&nbsp;<%= getHotWord(164)%></strong></td>
									  <td width="5%"><div align="center"><strong><%= getHotWord(70)%> </strong></div></td>
									  <!-- JM_4/16/2009 hidden as per Aspen's Req -->
									  <!-- <td width="10%"><div align="right"><strong>Price </strong></div></td> -->
									  <td width="10%"><div align="right"><strong><%= getHotWord(35)%> </strong></div></td>
									  <td width="20%"><div align="center"><strong><%= getHotWord(59)%> </strong></div></td>
									
									  <!-- MB bug#3233 added extra empty column to Purchases(left) table -->
									 <td width="5%"><div align="right">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</div></td>
										</tr>
                                    <%  if NOT request.form("frmExpReport")="true" then %>
										<tr height="2">
											  <td colspan="6" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
										</tr>
										
										<!-- MB bug#3233 added empty row to Purchases(left) table -->
										<tr>
										 <td colspan="6" align="center"> &nbsp;</td>
										</tr>
                                    <%  end if 

                                   
										strSQL ="SELECT DISTINCT [Sales Details].SDID, Sales.SaleDate, Location.LocationName, Sales.SaleID, Sales.SaleDate, [Sales Details].Returned, CASE WHEN ([Returns].ID IS NULL) THEN NULL ELSE 1 END AS ID, [Sales Details].Description, [Sales Details].UnitPrice, [Sales Details].Quantity, (CASE WHEN (NOT ([Sales Details].SDPaymentAmtB) Is Null) THEN [Sales Details].SDPaymentAmtB ELSE 0 END + CASE WHEN (NOT ([Sales Details].SDPaymentAmt) Is Null) THEN [Sales Details].SDPaymentAmt ELSE 0 END) as PaymentTotal FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN Location ON [Sales Details].Location = Location.LocationID LEFT OUTER JOIN [Returns] ON [Sales Details].SaleID = [Returns].SaleID WHERE Sales.ClientID = " & clientID &  " AND Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cSttDate & DateSep &" ORDER BY Sales.SaleDate"
										rsEntry2.CursorLocation = 3
										rsEntry2.open strSQL, cnWS
										response.write debugSQL(strSQL, "SQL")
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
											 <td width="5%"><div align="center"><%=rsEntry2("Quantity")%></div></td>
											<!--  <td width="10%"><div align="right">
											 <%=FmtCurrency(rsEntry2("UnitPrice"))%>
											 <%if rsEntry2("UnitPrice")>0 then response.write "&nbsp;" end if%></div></td> -->
											 <td width="10%"><div align="right">
											 <%=FmtCurrency(rsEntry2("PaymentTotal"))%>
											 <%if rsEntry2("PaymentTotal")>0 then response.write "&nbsp;" end if%></div></td>
											  <td width="20%"><div align="center">&nbsp;
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
											 <td colspan="5" align="center"><%=DisplayPhrase(phraseDictionary, "Nopurchaseactivity") %></td></tr>
											<%end if
											
											rsEntry2.close %>

</table>
											</td>
										
										
										<%end if'JM-51_2782 %>
										
										
											<td <%if not ss_IncPurchActInStmt then%>width="100%"<%else%>width="48%"<%end if%> valign="top"   <%if	not ss_IncPurchActInStmt then%>colspan="2"<%end if%>>
											<table  class="mainText" width="100%">
										<tr class="whiteHeader" style="background-color:#666666;">
										<td width="15%"><div align="left"><strong>&nbsp;<%= getHotWord(57)%></strong></div></td>
										  
										  <td width="55%" align="left"><strong><%= DisplayPhrase(phraseDictionary, "Balanceactivity")%></strong></td>
										  <%if NOT ss_IncPurchActInStmt then'JM_4/16/2009 hidden as per Aspen's Req %>
										  <td width="15%"><div align="right"><strong><%= getHotWord(35)%>&nbsp;</strong></div></td>
										  <%end if%>
										  <td width="15%"><div align="right"><strong><%= getHotWord(29)%>&nbsp;</strong></div></td>
									 	
										</tr>
										
<%                                  if NOT request.form("frmExpReport")="true" then %>
										<tr height="2">
										 <td <%if not ss_IncPurchActInStmt then%> colspan="4" <%else%> colspan="3"<%end if%> style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
											   
										</tr>
<%                                  end if 
                                    
                                    'JM-54_2836
                                    if request.form("OptRemoveBalForward")<>"on" then%>
										<tr height="20" style="background-color:<%=rowColor%>;">
										  <td width="15%"><div align="left"><%=FmtDateShort(DateAdd("d",-1, DateValue(cSDate)))%></div></td>
										  <td width="55%"><%= DisplayPhrase(phraseDictionary, "Balanceforward")%> </td>
<%                                      if NOT ss_IncPurchActInStmt then'JM_4/16/2009 hidden as per Aspen's Req %>
										  <td width="15%"><div align="right">&nbsp;</div></td>
<%                                      end if%>

										  <td width="15%"><div align="right"><%'=FmtCurrency(BeginBal*-1) %><% if isInvoice then response.write FmtCurrency(BeginBal) else response.write FmtCurrency(BeginBal*-1) end if%></div></td>
										</tr>
									<%end if 'request.form("OptRemoveBalForward")="on"%>
							<%if NOT rsEntry.EOF then %>			
                            <%'JM-52_2729 %>
                            <% if request.form("OptHideBalDet")="on" then %>
                            <% '
                            strSQL =" SELECT DISTINCT Cat.CategoryID, Cat.CategoryName, SUM(Cat.PaymentTotal) as PaymentTotal "
                            strSQL = strSQL & " FROM  tblClientAccount "
                            strSQL = strSQL & " INNER JOIN (SELECT tblPayments.PaymentID, Categories.CategoryName, [Sales Details].CategoryID, " 
                            strSQL = strSQL & " tblSDPayments.SDPaymentAmount AS PaymentTotal FROM Sales INNER JOIN [Sales Details] ON " 
                            strSQL = strSQL & " Sales.SaleID = [Sales Details].SaleID INNER JOIN Categories on Categories.CategoryID=[Sales Details].CategoryID "
                            strSQL = strSQL & " INNER JOIN tblPayments ON Sales.SaleID = tblPayments.SaleID "
                            strSQL = strSQL & " INNER JOIN tblSDPayments on tblSDPayments.PaymentID = tblPayments.PaymentID AND tblSDPayments.SDID = [Sales Details].SDID "
                            strSQL = strSQL & " WHERE Sales.ClientID = " & clientID &  " AND  Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND " 
                            strSQL = strSQL & " Sales.SaleDate <= " & DateSep & cSttDate & DateSep &"  ) Cat ON Cat.PaymentID=tblClientAccount.PaymentID " 
                            strSQL = strSQL & " WHERE (tblClientAccount.Amount < 0) AND (tblClientAccount.ClientID = " & clientID &  ") " 
                            strSQL = strSQL & " AND  (tblClientAccount.EntryDate >= " & DateSep & cSDate & DateSep & " " 
                            strSQL = strSQL & " AND tblClientAccount.EntryDate <= " & DateSep & cSttDate & DateSep &") " 
                             ''JM-55_3182
					        if optEventBal ="on" then
						        strSQL = strSQL & " AND NOT (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
					        else
						        strSQL = strSQL & " AND (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
					        end if
                           ' strSQL = strSQL & " AND  (tblClientAccount.ClassID Is Null OR tblClientAccount.ClassID=0) " 
                            strSQL = strSQL & " AND (tblClientAccount.ClientContractID IS NULL OR  NOT(tblClientAccount.DepositReleaseDate IS NULL)) " 
                            strSQL = strSQL & " GROUP BY Cat.CategoryID, Cat.CategoryName"
                           logIt strSQL
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
                                        <td width="15%"><div align="right"><%'=FmtCurrency(AmountByCategory*-1) %><% if isInvoice then response.write FmtCurrency(AmountByCategory) else response.write FmtCurrency(AmountByCategory*-1) end if%></div></td>
                                        <%end if%>
                                        
                                        <td width="15%"><div align="right"><%'=FmtCurrency(BalByCategory*-1) %><% if isInvoice then response.write FmtCurrency(BalByCategory) else response.write FmtCurrency(BalByCategory*-1) end if%></div></td>
                                    </tr>
                                    <% rsEntry2.movenext
                                    loop
                                    %>
                                    <%end if
                                
                                rsEntry2.close %>
                                
                                  <%
                                  
                                  strSQL =" SELECT SUM(tblClientAccount.Amount) AS AccountCredit FROM " & tblClientAccountStr & " "
                                  strSQL = strSQL & " WHERE (tblClientAccount.Amount > 0) AND  (tblClientAccount.ClientID =  " & clientID &  ") AND  (tblClientAccount.EntryDate >= " & DateSep & cSDate & DateSep & " AND tblClientAccount.EntryDate <= " & DateSep & cSttDate & DateSep &" ) "
                                  
                                  'AND  (tblClientAccount.ClassID Is Null OR tblClientAccount.ClassID=0) 
                               ''JM-55_3182
						        if optEventBal ="on" then
							        strSQL = strSQL & " AND NOT (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
						        else
							        strSQL = strSQL & " AND (tblClientAccount.ClassID IS NULL OR tblClientAccount.ClassID = 0) "
						        end if
                                strSQL = strSQL & " AND (tblClientAccount.ClientContractID IS NULL OR  NOT(tblClientAccount.DepositReleaseDate IS NULL)) "
                              
                            logIt strSQL
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
                                            <%=DisplayPhrase(phraseDictionary, "Accountcredit") %>
                                        </td>
                                        <%if NOT ss_IncPurchActInStmt then%>
                                        <td width="15%"><div align="right"><%'=FmtCurrency(AcctCredit*-1) %><% if isInvoice then response.write FmtCurrency(AcctCredit) else response.write FmtCurrency(AcctCredit * -1) end if%></div></td>
                                        <%end if%>
                                        <td width="15%"><div align="right"><%'=FmtCurrency(BalByCategory*-1) %><% if isInvoice then response.write FmtCurrency(BalByCategory) else response.write FmtCurrency(BalByCategory*-1) end if%></div></td>
                                    </tr>
                                    <% 
                                    %>
                                    <%end if
                                
                                rsEntry2.close %>
                                
                                <%end if 'request.form("OptHideBalDet")="on"  %>
										 
									<%
									Balance=BeginBal		'Starting balance for the activity list
									do while NOT rsEntry.EOF
										if rowColor = "#F2F2F2" then
											rowColor = "#FAFAFA"
										else
											rowColor = "#F2F2F2"
										end if
										
										
										
										Amount = rsEntry("Amount") * (-1)
										
										Balance = Balance + Amount				
										'stLineItem = getHotWord(37) & rsEntry("PmtRefNo")  & ": " &  rsEntry("TypePurch")
										'JM - 45_2329
										'stLineItem = "Sale " 
										'if NOT request.form("frmExpReport")="true" then
											'stLineItem = stLineItem & "<a href=""adm_tlbx_voidedit.asp?saleno=" & rsEntry("SaleID") & """>"
										'end if
										'stLineItem = stLineItem & "#" & rsEntry("SaleID")
										'stLineItem = getHotWord(37)
										stLineItem = "Account Activity #"
										if NOT request.form("frmExpReport")="true" then
											stLineItem = stLineItem & "<a href=""adm_tlbx_voidedit.asp?saleno=" & rsEntry("SaleID") & """>"
										end if
										stLineItem = stLineItem & rsEntry("ClientAccountID")
										if NOT request.form("frmExpReport")="true" then
											stLineItem = stLineItem & "</a>"
										end if
										
										if NOT isNULL(rsEntry("LocationName")) then
											stLineItem = stLineItem & " at " & rsEntry("LocationName") 
										end if										
									%>
									     <% if NOT request.form("OptHideBalDet")="on" then %>
										 
										
										<tr height="20" style="background-color:<%=rowColor%>;">
										  <td valign="top" width="15%"><div align="left" ><%=FmtDateShort(rsEntry("EntryDate"))%></div></td>

										  <td width="55%" ><%=stLineitem%>
										    <%pActLineCount=pActLineCount + 1%>
											
											<%
											strSQL = "SELECT SD.SaleID,SD.Description,SD.RecipClientID, CLIENTS.FirstName, CLIENTS.LastName FROM "
											strSQL = strSQL & "(SELECT [Sales Details].SaleID, [Sales Details].Description, IsNull([Sales Details].RecClientID,[Payment Data].ClientID) as RecipClientID FROM [Sales Details] LEFT OUTER JOIN [Payment Data] ON [Sales Details].PmtRefNo = [Payment Data].PmtRefNo "
										if isNULL(rsEntry("SaleID")) then
											strSQL = strSQL & " WHERE 1=0"
										else
											strSQL = strSQL & "WHERE [Sales Details].SaleID = " & rsEntry("SaleID") & " "
										end if
										strSQL = strSQL & ") SD LEFT OUTER JOIN CLIENTS ON SD.RecipClientID = CLIENTS.ClientID "
										strSQL = strSQL & "ORDER BY SD.Description" 
                                        response.write debugSQL(strSQL, "SQL")
										rsEntry2.CursorLocation = 3
										logIt strSQL
										rsEntry2.open strSQL, cnWS
										Set rsEntry2.ActiveConnection = Nothing

										
										stLineItemDetail=""
											If NOT rsEntry2.EOF then
												do while NOT rsEntry2.EOF
													stLineItemDetail = rsEntry2("Description")
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

										<% if NOT ss_IncPurchActInStmt then'JM_4/16/2009 hidden as per Aspen's Req %>
										  <td width="15%"><div align="right"><%'=FmtCurrency(Amount*-1) %><% if isInvoice then response.write FmtCurrency(Amount) else response.write FmtCurrency(Amount*-1) end if%></div></td>
										<% end if %>
										  <td width="15%"><div align="right"><%'=FmtCurrency(Balance*-1)%><% if isInvoice then response.write FmtCurrency(Balance) else response.write FmtCurrency(Balance*-1) end if%></div></td>
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
<%                                  if NOT request.form("frmExpReport")="true" then %>
										<tr height="2">
											  <td colspan="2" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
										</tr>
<%                                  end if %>
										<tr height="20">
										  <td colspan="2"><div align="right"><strong> <% 	if isInvoice then %><%=DisplayPhrase(phraseDictionary, "Amountdue") %>:<%	else %><%= getHotWord(29)%>:<% 	end if %>&nbsp;&nbsp;&nbsp;&nbsp;<%=FmtCurrency(abs(Balance))%></strong></div></td>
										</tr>
                           				<% if ss_ShowReceiptMessage then %>
										  <% if NOT request.form("frmExpReport")="true" then %>
										    <tr><td colspan="2" ><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="50" width="100%"></td></tr>
											 <tr height="2"><td colspan="2" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td></tr>
										  <% end if %>
										  <tr height="40"><td colspan="2" class="userHTML"><%=receiptMsg1%></td></tr>
										  <tr height="20"><td colspan="2" class="userHTML"><%=receiptMsg2%></td></tr>
                                          <tr><td colspan="2">&nbsp;</td></tr>
										<%end if %>
									<%
									rsEntry.close
									set rsEntry = nothing
									set rsEntry2 = nothing
								'end if	'eof

			''' *************** PAYMENT COUPON ********************
%>
<%'if (pActLineCount > 15 and pActLineCount < 30) OR (pPurchLineCount > 15 and pPurchLineCount < 30) and request.form("chkPrinterFriendly")="on" then %>

<%        'MB 7863 fit everything to one page if there are less then 10 lines and show pp_PleaseTearOff
          'push Payment Coupon to a separate page otherwise, works OK for pdf export, 
          'page-break-before does not work in FF as expected for printing
        if (pActLineCount > 10 ) then %>
            <tr><td align="center" colspan="2"><div style="page-break-before:always!important;">&nbsp;</div></td></tr>
        <%else%>
            <tr><td align="center" colspan="2"><div style="page-break-before:auto!important;">&nbsp;</div></td></tr>
            <tr><td align="center" colspan="2"><%=DisplayPhrase(phraseDictionary, "Pleasetearoff") %></td></tr>
        <%end if%>
<%'end if%>
						<tr><td colspan="2"><table style="border-top: 2px dashed #666666;" width="100%"><tr><td>&nbsp;</td></tr></table></td></tr>
						<tr>
							<td colspan="2"><b><%=DisplayPhrase(phraseDictionary, "Paymentcoupon") %></b></td>
						</tr>
						<tr>
							<td width="45%" class="mainText"><%=CurClientName%></td>
							<td rowspan="3" align="right">
								<table border="1" cellspacing="0" bordercolor="#666666" class="mainText" width="45%">
								<tr>
									<td width= "60%" style="background-color:#CCCCCC;"> &nbsp;<%=DisplayPhrase(phraseDictionary, "Statementdate") %>:</td>
									<td align="right"><%=FmtdateShort(CSttDate)%></td>
								</tr>
								<tr>
									<td width="112" nowrap style="background-color:#CCCCCC;"><% 	if isInvoice then %>&nbsp;<%=DisplayPhrase(phraseDictionary, "Amountdue") %>: <%	else %>&nbsp;<%= getHotWord(29)%>: <% 	end if %></td>
								    <td width="109"><div align="right"><%=FmtCurrency(abs(Balance))%></div></td>
								</tr>
								<tr>
									<td style="background-color:#CCCCCC;"> &nbsp;<%=DisplayPhrase(phraseDictionary, "Enclosed") %>:</td>
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
								<strong><%=DisplayPhrase(phraseDictionary, "Topaybycreditcard") %>:</strong><br />
								<table border="0" style="background-color:#CCCCCC;"  class="mainText">
									<tr><td><%= getHotWord(146)%>:&nbsp;_______________________________</td></tr>
									<tr><td><%= getHotWord(46)%>:&nbsp;______________________________________</td></tr>
									<tr><td><%= getHotWord(47)%>:&nbsp;______________________________</td></tr>
									<tr><td><%= getHotWord(48)%>:&nbsp;________________________</td></tr>
									<tr><td><%= getHotWord(49)%>:&nbsp;________________________</td></tr>
									<tr><td height="4"></td></tr>
									<tr><td><%= getHotWord(50)%>&nbsp;<span style="font-size:9px;"><%=DisplayPhrase(phraseDictionary, "Checkone") %></span>:&nbsp;&nbsp;<% if ss_Amex then%>__AMEX&nbsp;&nbsp;<%end if %><% if ss_Visa then%>__VISA&nbsp;&nbsp;<%end if %><% if ss_Mastercard then%>__MC&nbsp;&nbsp;<%end if %><% if ss_Discover then%>__DISC<%end if %></td></tr>
									<tr><td><%= getHotWord(51)%>&nbsp;:&nbsp;_________________________________</td></tr>
									<tr><td><%= getHotWord(52)%>&nbsp;<span style="font-size:9px;">(MM/YYYY)</span>:&nbsp;______________________</td></tr>
								</table>
							</td>
							<td align="right">
								<b><%=DisplayPhrase(phraseDictionary, "Topaybycheck") %>: <u><%=session("StudioName")%></u></b><br /><br />
								<%=DisplayPhrase(phraseDictionary, "Iauthorize") %><br /><span class="textSmall"><%=DisplayPhrase(phraseDictionary, "Areceipt") %></span><br /><br />
								<strong>X</strong>_____________________________________
							</td>
						</tr>
						<%end if 'ss_InvoiceShowCC%>

						<%if ss_InvoiceAskAutoPay then %>
						<tr >
							<td colspan="2"><strong>
								<input type="checkbox" name="optIAgree"> 
								<%=DisplayPhrase(phraseDictionary, "Pleasecontactme") %>&nbsp;<%= getHotWord(93)%>: _______________
								</strong>
							</td>
						</tr>
						<% end if 'ss_InvoiceAskAutoPay%>						
<%
							end if		'end of generate report if statement
							%>
						  </table>
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
