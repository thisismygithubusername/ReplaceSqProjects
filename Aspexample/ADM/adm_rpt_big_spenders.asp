<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
	 dim rsEntry, rsEntry2
	 set rsEntry = Server.CreateObject("ADODB.Recordset")
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
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_SALES") then 
		%>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
		<%
	else
		%>
			<!-- #include file="../inc_i18n.asp" -->
			<!-- #include file="inc_hotword.asp" -->
		<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

		Dim showDetails, cSDate, cEDate, tmpCurDate, rsPrice, rsQty, rsDiscount, rsTax, rsPaid, cLoc, tmpSubTotQty, tmpSubSaleTotal, tmpTotQty, tmpIsProduct, colWidth
		Dim tmpPMTotal, rowColor, cashBasis, splitFactorA, splitFactorB, tmpClientName, rptSummary, tmpSaleTotal, tmpQTY
		Dim tmpClientID, tmpCatName, tmpSubCat, rsTmpSubCat, tmpTotSub, tmpTotDiscount, tmpTotTax, tmpTotal, tmpTotalDisc, tmpTotCashEq, tmpTotNonCashEq
		Dim TotSub, TotDiscount, TotTax, GrandTotal, TotCashEq, TotNonCashEq, GrandTotalDisc, dd_paymeth, printedProdServ, rangeSalesTotal, rangeProductsTotal, rangeServicesTotal
		Dim tmpProductQTY, tmpServiceQTY, tmpServicesTotal, tmpProductTotal, tmpSalesQTY, tmpSalesTotal, rowCount, prodServ, ap_view_all_locs
		
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

		if request.form("optSummary") = "1" then
			rptSummary = true
		else
			rptSummary = false
		end if

		prodServ = request.form("optProdServ")

		if request.form("optBasis")="0" then
			cashBasis = false
		else
			cashBasis = true
		end if

		if request.form("optTG")<>"" and request.form("optTG")<>"0" then
			prodServ = "0"
		end if
		
		if NOT request.form("frmExpReport")="true" then
		%>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "adm/sorttable", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_big_spenders", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %>
			<script type="text/javascript">
			function exportReport() {
				document.frmSales.frmExpReport.value = "true";
				document.frmSales.frmGenReport.value = "true";
				<% iframeSubmit "frmSales", "adm_rpt_big_spenders.asp" %>
			}
			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
		<%
		end if
		
		%>
		


		<% if NOT request.form("frmExpReport")="true" then %>
<% pageStart %>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div class="headText breadcrumbs-old">
			<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
			<%if category<>"" then%>
			<span class="breadcrumb-item">&raquo;</span>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<% end if %>
			<span class="breadcrumb-item">&raquo;</span>
			<%=  DisplayPhrase(reportPageTitlesDictionary, "Bigspenders") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
			</div>
<%end if %>
			<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
			<tr> 
			  <td valign="top" height="100%" width="100%"> 
				<table cellspacing="0" width="90%" height="100%" style="margin: 0 auto;">
				<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<tr>
					<td class="headText"  align="left" valign="top">
					  <table width="100%" cellspacing="0">
						<tr>
						  <td class="headText" valign="bottom"><b id="bigSpenderHeader"> <%= pp_PageTitle("Big Spenders") %></b></td>
						  <td valign="bottom" class="right" height="26"> </td>
						</tr>
					  </table>
					</td>
				  </tr>
				  <%end if %>
				  <tr> 
					<td height="30" valign="bottom" class="headText">
						  <form name="frmSales" action="adm_rpt_big_spenders.asp" method="POST">
							<input type="hidden" name="frmGenReport" value="">
							<input type="hidden" name="frmExpReport" value="">
							<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
								<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
								<input type="hidden" name="category" value="<%=category%>">
							<% end if %>
						<table class="mainText border4 center-block" cellspacing="0">
							<tr> 
							  <td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b><span style="color:<%=session("pageColor4")%>">&nbsp;</span><%=xssStr(allHotWords(77))%> 
								<input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
				<script type="text/javascript">
					var cal1 = new tcal({'formname':'frmSales', 'controlname':'requiredtxtDateStart'});
					cal1.a_tpl.yearscroll = true;
				</script>
								&nbsp;<%=xssStr(allHotWords(79))%> 
								<input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
				<script type="text/javascript">
					var cal2 = new tcal({'formname':'frmSales', 'controlname':'requiredtxtDateEnd'});
					cal2.a_tpl.yearscroll = true;
				</script>
								&nbsp;
								<select name="optSummary" onChange="singleSubmit()">
									<option value="0"<% if request.form("optSummary")="0" then %> selected<% end if %>>Detail</option>
									<option value="1"<% if request.form("optSummary")="1" or request.form("optSummary")="" then %> selected<% end if %>>Summary</option>
								</select>
						<% if request.form("optSummary") = "1" or request.form("optSummary")="" then %>
								&nbsp;Top: <input type="text" size="3" value="<% if request.form("optTopNum")="" then %>100<% else response.write(request.form("optTopNum")) end if%>" name="optTopNum">
						<% end if %>
								&nbsp;
								<select name="optProdServ" onChange="if (this.value != '0') { document.frmSales.optTG.value = '0' }; singleSubmit()">
									<option value="0"<% if prodServ="0" then %> selected<% end if %>>Services Only</option>
									<option value="1"<% if prodServ="1" then %> selected<% end if %>>Products Only</option>
									<option value="2"<% if prodServ="2" or prodServ="" then %> selected<% end if %>>Products and Services</option>
								</select>
								&nbsp;

						<select name="optTG" onchange="if (this.value != '0') { document.frmSales.optProdServ.value='0' }; singleSubmit();"<% if prodServ<>"0" then %> style="visibility:hidden"<% end if%>>
						<option value="0" <%if request.form("optTG")="0" then response.write "selected" end if%>>All Type Groups</option>
						<%
							strSQL = "SELECT TypegroupID, Typegroup FROM tblTypegroup "
							strSQL = strSQL & "WHERE [Active]=1 "
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
<br />
								&nbsp;Who spent more than: $<input type="text" name="optMinValue" size="3" value="<% if request.form("optMinValue")<>"" then response.write(request.form("optMinValue")) else response.write("0.00") end if %>">
								<!-- &nbsp;Sort By:
								<select name="optSortBy" onChange="singleSubmit()">
									<option value="0"<% if request.form("optSortBy")="0" then %> selected<% end if %>>Sales Total</option>
									<option value="1"<% if request.form("optSortBy")="1" then %> selected<% end if %>><%=session("ClientHW")%> Name</option>
									<option value="2"<% if request.form("optSortBy")="2" then %> selected<% end if %>>Quantity Sold</option>
								</select-->
								
								&nbsp;
								<%=xssStr(allHotWords(8))%>:
								<select name="optSaleLoc" onChange="singleSubmit()" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
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
					document.frmSales.optSaleLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
				</script> 
								&nbsp;
								<select name="optNewClientOnly" onChange="singleSubmit()">
									<option value="0"<% if request.form("optNewClientOnly")="0" then %> selected<% end if %>>All <%=session("ClientHW")%>s</option>
									<option value="1"<% if request.form("optNewClientOnly")="1" then %> selected<% end if %>>With First Sales During</option>
									<option value="2"<% if request.form("optNewClientOnly")="2" then %> selected<% end if %>>With First Sales Before</option>
								</select>
								&nbsp;
								<select name="optBasis">
								  <option value="0" <%if request.form("optBasis")="0" then response.write "selected" end if%>>Accrual Basis</option>
								  <option value="1" <%if request.form("optBasis")="1" then response.write "selected" end if%>>Cash Basis</option>
								</select>
								Hide QTY Columns: <input type="checkbox" name="optHideQTY" <% if request.form("optHideQTY")="on" then response.write "checked" end if %>>

								<br />
								<% showDateArrows("frmSales") %>
								<% taggingFilter %>
								<input type="button" name="Button" value="Generate" onClick="genReport();">
								<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
								else%>
									<% exportToExcelButton %>
								<% end if %>
								<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
								 else%>
									<% taggingButtons("frmSales") %>
								<%end if%>
							<% savingButtons "frmSales", "Big Spenders" %>
							</b>&nbsp;
							</td>
							</tr>
							  </form>
					  </table>
					</td>
				  </tr>
				  <tr> 
					<td valign="top" id="bigSpendersGenTag" class="mainTextBig"> 
					  <table class="mainText center" width="95%" cellspacing="0" style="margin: 0 auto;">
						<tr >
						  <td  colspan="2" valign="top" class="mainTextBig center-ch">&nbsp;</td>
						</tr>
						<tr > 
						  <td class="mainTextBig" colspan="2" valign="top">
		<% end if			'end of frmExpreport value check before /head line	  
	%>	<table width="100%"  cellspacing="0"<% if rptSummary then response.write " class=""sortable"" " else response.write " class=""mainText"" " end if %> id="sortable_table" style="margin: 0 auto;"> 
		<% if request.form("frmGenReport")="true" then
			if request.form("frmExpReport")="true" then
				Dim stFilename
				stFilename="attachment; filename=Big Spender Report " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
				Response.ContentType = "application/vnd.ms-excel" 
				Response.AddHeader "Content-Disposition", stFilename 
			end if
			tmpClientID = ""
			TotSub = 0
			TotDiscount = 0
			TotTax = 0
			GrandTotal = 0
			GrandTotalDisc = 0
			TotCashEq = 0
			TotNonCashEq = 0
			
			if rptSummary then 
				strSQL = "SELECT "

				if isNum(request.form("optTopNum")) then
				    if CINT(request.form("optTopNum")) >= 1 then
					    strSQL = strSQL & " TOP " & request.form("optTopNum")
					end if
				end if

				strSQL = strSQL & " Clients.ClientID, Clients.LastName, Clients.FirstName, Clients.RSSID, SalesTotal.SalesTotal, SalesTotal.SalesServicesTotal, SalesTotal.SalesProductsTotal, "
				strSQL = strSQL & " SUM(SDP.Quantity) as SalesQTY, SUM(SDP.SDPaymentAmount) as SaleTotal "
				
				if request.form("optNewClientOnly")<>"" then 
					strSQL = strSQL & ", ALLSALES.FirstSale "
				end if
				
				strSQL = strSQL & ", SUM(CASE WHEN (([Sales Details].CategoryID >25)) THEN SDP.Quantity ELSE 0 END) as ProductQTY "
				strSQL = strSQL & ", SUM(CASE WHEN (([Sales Details].CategoryID<=25)) THEN SDP.Quantity ELSE 0 END) as ServicesQTY "

				strSQL = strSQL & ", SUM(CASE WHEN [Sales Details].CategoryID >25 THEN SDP.SDPaymentAmount ELSE 0 END) as ProductTotal "
				strSQL = strSQL & ", SUM(CASE WHEN [Sales Details].CategoryID<=25 THEN SDP.SDPaymentAmount ELSE 0 END) as ServicesTotal "

				strSQL = strSQL & " FROM [Sales Details] INNER JOIN Sales ON [Sales Details].SaleID = Sales.SaleID INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID " 
					
					strSQL = strSQL & " INNER JOIN ( " 
					strSQL = strSQL & " SELECT [Sales Details].SDID, [Sales Details].Quantity, SUM(tblSDPayments.SDPaymentAmount) as SDPaymentAmount "
					strSQL = strSQL & " FROM [Sales Details] INNER JOIN tblSDPayments ON [Sales Details].SDID = tblSDPayments.SDID "
					strSQL = strSQL & " INNER JOIN tblPayments ON tblPayments.PaymentID = tblSDPayments.PaymentID "
					strSQL = strSQL & " INNER JOIN [Payment Types] ON tblPayments.PaymentMethod = [Payment Types].Item# "
					if cashBasis then
        				strSQL = strSQL & " WHERE [Payment Types].[CashEQ]=1 AND [Sales Details].CategoryID<>21 "
					else ''Accrual Basis
						strSQL = strSQL & " WHERE (NOT ([Sales Details].CategoryID BETWEEN 21 AND 23)) "
					end if
					strSQL = strSQL & " GROUP BY [Sales Details].SDID, [Sales Details].Quantity "
					strSQL = strSQL & " ) SDP ON [Sales Details].SDID = SDP.SDID " 
								
				' BQL 50_2000 added cross join of sales summary data to use for % column
					strSQL = strSQL & " CROSS JOIN ( SELECT SUM(tblSDPayments.SDPaymentAmount) as SalesTotal " 
					strSQL = strSQL & ", SUM(CASE WHEN [Sales Details].CategoryID  > 25 THEN tblSDPayments.SDPaymentAmount ELSE 0 END) as SalesProductsTotal "
					strSQL = strSQL & ", SUM(CASE WHEN [Sales Details].CategoryID <= 25 THEN tblSDPayments.SDPaymentAmount ELSE 0 END) as SalesServicesTotal "
					strSQL = strSQL & " FROM Sales INNER JOIN [Sales Details] ON [Sales Details].SaleID = Sales.SaleID INNER JOIN tblSDPayments ON [Sales Details].SDID = tblSDPayments.SDID INNER JOIN tblPayments ON tblSDPayments.PaymentID = tblPayments.PaymentID INNER JOIN [Payment Types] ON tblPayments.PaymentMethod = [Payment Types].Item# " 
					if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
						strSQL = strSQL & " INNER JOIN [Payment Data] ON  [Sales Details].PmtRefNo = [PAYMENT DATA].PmtRefNo "
					end if
					strSQL = strSQL & "WHERE (Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") "
					if cLoc<>0 then
						strSQL = strSQL & "AND ([Sales Details].Location=" & cLoc & ") "
					end if
					if prodServ="1" then ' Products Only
						strSQL = strSQL & " AND ([Sales Details].CategoryID>25) "
					elseif prodServ="0" then ' Services Only
						strSQL = strSQL & " AND ([Sales Details].CategoryID<=25) "
					end if
					
					if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
						strSQL = strSQL & " AND [Payment Data].TypeGroup = " & request.form("optTG")
					end if
	
					if cashBasis then
        				strSQL = strSQL & " AND [Payment Types].[CashEQ]=1 AND [Sales Details].CategoryID<>21 "
					else ''Accrual Basis
						strSQL = strSQL & " AND (NOT ([Sales Details].CategoryID BETWEEN 21 AND 23)) "
					end if

					strSQL = strSQL & ") SalesTotal "
				' end 50_2000
				
				if request.form("optNewClientOnly")<>"" then
					strSQL = strSQL & " INNER JOIN (SELECT MIN(SALES.SaleDate) as FirstSale, SALES.ClientID FROM SALES GROUP BY Sales.ClientID) ALLSALES ON ALLSALES.ClientID = Clients.ClientID "
				end if

				if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
					strSQL = strSQL & " INNER JOIN [Payment Data] ON  [Sales Details].PmtRefNo = [PAYMENT DATA].PmtRefNo "
				end if
				
				if request.form("optFilterTagged")="on" then
					strSQL = strSQL & " INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
				end if
				
				strSQL = strSQL & "WHERE (Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") AND Clients.ClientID <> 1 "
				if cLoc<>0 then
					strSQL = strSQL & "AND ([Sales Details].Location=" & cLoc & ") "
				end if
				if prodServ="1" then ' Products Only
					strSQL = strSQL & " AND ([Sales Details].CategoryID>25) "
				elseif prodServ="0" then ' Services Only
					strSQL = strSQL & " AND ([Sales Details].CategoryID<=25) "
				end if
				
				if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
					strSQL = strSQL & " AND [Payment Data].TypeGroup = " & request.form("optTG")
				end if

				if cashBasis then
    				strSQL = strSQL & " AND [Sales Details].CategoryID<>21 "
				else ''Accrual Basis
					strSQL = strSQL & " AND (NOT ([Sales Details].CategoryID BETWEEN 21 AND 23)) "
				end if
				
				if request.form("optFilterTagged")="on" then
					if session("mvaruserID")<>"" then
						strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
					else
						strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
					end if
				end if
				
				strSQL = strSQL & " GROUP BY Clients.ClientID, Clients.LastName, Clients.FirstName, Clients.RSSID, ALLSALES.FirstSale, SalesTotal.SalesTotal, SalesTotal.SalesServicesTotal, SalesTotal.SalesProductsTotal "

				strSQL = strSQL & " HAVING 1=1 " 
				if request.form("optMinValue")<>"" then
					strSQL = strSQL & " AND SUM(SDP.SDPaymentAmount) >= " & request.form("optMinValue")
				end if

				if request.form("optNewClientOnly")="1" then
					strSQL = strSQL & " AND (ALLSALES.FirstSale >= " & DateSep & cSDate & DateSep & ") AND (ALLSALES.FirstSale <= " & DateSep & cEDate & DateSep & ") "
				elseif request.form("optNewClientOnly")="2" then
					strSQL = strSQL & " AND (ALLSALES.FirstSale < " & DateSep & cSDate & DateSep & " ) "
				end if

				'if request.form("frmTagClients")<>"true" then
					'if request.form("optSortBy")="0" then
						strSQL = strSQL & " ORDER BY SaleTotal DESC, Clients.LastName, SalesQTY DESC "
					'elseif request.form("optSortBy")="1" then
						'strSQL = strSQL & " ORDER BY Clients.LastName, SaleTotal DESC, SalesQTY DESC "
					'else 
						'strSQL = strSQL & " ORDER BY SalesQTY DESC, SaleTotal DESC, Clients.LastName "
					'end if
				'end if
				
			else 
				strSQL = "SELECT SALESTOTALS.SaleTotal, SALESTOTALS.SalesQTY, Clients.ClientID, Clients.LastName, Clients.FirstName, Clients.RSSID, "
				strSQL = strSQL & " [Sales Details].Description, [Sales Details].Quantity, Sales.SaleID, Sales.SaleDate, Location.LocationName "

				if request.form("optNewClientOnly")<>"" then 
					strSQL = strSQL & ", ALLSALES.FirstSale "
				end if

				strSQL = strSQL & ", tblSDPayments.SDPaymentAmount as SDTotal "
				strSQL = strSQL & ", CASE WHEN ([Sales Details].CategoryID > 25) THEN 1 ELSE 0 END as IsProduct "
				strSQL = strSQL & " FROM tblSDPayments INNER JOIN [Sales Details] INNER JOIN Sales ON [Sales Details].SaleID = Sales.SaleID INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID INNER JOIN Location ON Location.LocationID = [Sales Details].Location ON tblSDPayments.SDID = [Sales Details].SDID INNER JOIN [Payment Types] INNER JOIN tblPayments ON [Payment Types].Item# = tblPayments.PaymentMethod ON tblSDPayments.PaymentID = tblPayments.PaymentID "
				if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
					strSQL = strSQL & " INNER JOIN [Payment Data] ON  [Sales Details].PmtRefNo = [PAYMENT DATA].PmtRefNo "
				end if
				if request.form("optNewClientOnly")<>"" then
					strSQL = strSQL & " INNER JOIN (SELECT MIN(SALES.SaleDate) as FirstSale, SALES.ClientID FROM SALES GROUP BY Sales.ClientID) ALLSALES ON ALLSALES.ClientID = Clients.ClientID "
				end if
				if request.form("optFilterTagged")="on" then
					strSQL = strSQL & " INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
				end if
				strSQL = strSQL & " INNER JOIN (SELECT Sales.ClientID, SUM(tblSDPayments.SDPaymentAmount) as SaleTotal, Sum([Sales Details].Quantity) as SalesQTY "
				strSQL = strSQL & " FROM tblSDPayments INNER JOIN [Sales Details] INNER JOIN Sales ON [Sales Details].SaleID = Sales.SaleID ON tblSDPayments.SDID = [Sales Details].SDID INNER JOIN [Payment Types] INNER JOIN tblPayments ON [Payment Types].Item# = tblPayments.PaymentMethod ON tblSDPayments.PaymentID = tblPayments.PaymentID " 
				if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
					strSQL = strSQL & " INNER JOIN [Payment Data] ON  [Sales Details].PmtRefNo = [PAYMENT DATA].PmtRefNo "
				end if
								
				strSQL = strSQL & " WHERE Sales.ClientID <> 1 AND (Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") "
				if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
					strSQL = strSQL & " AND [Payment Data].TypeGroup = " & request.form("optTG")
				end if
				if prodServ="1" then
					strSQL = strSQL & " AND ([Sales Details].CategoryID>25) "
				elseif prodServ="0" then
					strSQL = strSQL & " AND ([Sales Details].CategoryID<=25) "
				end if
				if cLoc<>0 then
					strSQL = strSQL & "AND ([Sales Details].Location=" & cLoc & ") "
				end if
				if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
					strSQL = strSQL & " AND [Payment Data].TypeGroup = " & request.form("optTG")
				end if
				
				if cashBasis then
    				strSQL = strSQL & " AND [Payment Types].[CashEQ]=1 AND [Sales Details].CategoryID<>21 "
				else ''Accrual Basis
					strSQL = strSQL & " AND (NOT ([Sales Details].CategoryID BETWEEN 21 AND 23)) "
				end if
				
				strSQL = strSQL & " GROUP BY Sales.ClientID  HAVING 1=1 "
				if request.form("optMinValue")<>"" then
					strSQL = strSQL & " AND SUM(tblSDPayments.SDPaymentAmount) >= " & request.form("optMinValue")
				end if
				
				strSQL = strSQL & ") SALESTOTALS ON Sales.ClientID = SALESTOTALS.ClientID "
				strSQL = strSQL & "WHERE (Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") AND Clients.ClientID <> 1 "
				if cLoc<>0 then
					strSQL = strSQL & "AND ([Sales Details].Location=" & cLoc & ") "
				end if
				
				if prodServ="1" then
					strSQL = strSQL & " AND ([Sales Details].CategoryID>25) "
				elseif prodServ="0" then
					strSQL = strSQL & " AND ([Sales Details].CategoryID<=25) "
				end if

				if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
					strSQL = strSQL & " AND [Payment Data].TypeGroup = " & request.form("optTG")
				end if
				
				if cashBasis then
    				strSQL = strSQL & " AND [Payment Types].[CashEQ]=1 AND [Sales Details].CategoryID<>21 "
				else ''Accrual Basis
					strSQL = strSQL & " AND (NOT ([Sales Details].CategoryID BETWEEN 21 AND 23)) "
				end if
				
				if request.form("optNewClientOnly")="1" then
					strSQL = strSQL & " AND (ALLSALES.FirstSale >= " & DateSep & cSDate & DateSep & ") AND (ALLSALES.FirstSale <= " & DateSep & cEDate & DateSep & ") "
				elseif request.form("optNewClientOnly")="2" then
					strSQL = strSQL & " AND (ALLSALES.FirstSale < " & DateSep & cSDate & DateSep & " ) "
				end if
				
				if request.form("optFilterTagged")="on" then
					if session("mvaruserID")<>"" then
						strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
					else
						strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
					end if
				end if
	
				'if request.form("frmTagClients")<>"true" then
					'if request.form("optSortBy")="0" then
						strSQL = strSQL & " ORDER BY SALESTOTALS.SaleTotal DESC, Clients.LastName, IsProduct, Sales.SaleDate, SALESTOTALS.SalesQTY DESC"
					'elseif request.form("optSortBy")="1" then
						'strSQL = strSQL & " ORDER BY Clients.LastName, SALESTOTALS.SaleTotal DESC, SALESTOTALS.SalesQTY DESC, IsProduct "
					'else 
						'strSQL = strSQL & " ORDER BY SALESTOTALS.SalesQTY DESC, SALESTOTALS.SaleTotal DESC, Clients.LastName, IsProduct "
					'end if
				'end if

			end if

		response.write debugSQL(strSQL, "SQL")
			'response.end
			
			if request.form("frmTagClients")="true" then
				if rptSummary then
					if request.form("frmTagClientsNew")="true" then
						clearAndTagQuery(strSQL)
					else
						tagQuery(strSQL)
					end if
				else %>
					<script>
						alert("Detail report results can't be tagged");
					</script>
				<% end if
				strSQL = "SELECT StudioID FROM Studios WHERE 1=0 "
			end if

			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing

			if rptSummary then ' Summary view
				if request.form("frmTagClients")="true" then				
					' old condition
				else 
					rowCount = 1

					if request.form("optHideQTY")="on" then
						if prodServ="" or prodServ="2" then ' Products and Services 
							colWidth = "20%" ' 8 columns
						else
							colWidth = "33%" ' 4 columns
						end if
					else
						if prodServ="" or prodServ="2" then ' Products and Services 
							colWidth = "12%" ' 8 columns
						else
							colWidth = "25%" ' 4 columns
						end if
					end if 
						
					if NOT rsEntry.EOF then %>
								<tr <% if NOT request.form("frmExpReport")="true" then %>style="background-color:<%=session("pageColor4")%>;" class="whiteHeader" <% end if %> class="right">
				 	<% 	if request.form("frmExpReport")="true" then %>
									<th><%= getHotWord(134)%></th>
					<% 	end if %>
						  
									<th width="22%" align="left"><%=session("ClientHW")%></th>
					<%	if prodServ<>"1" then ' Products Only %>
						<%	if NOT request.form("optHideQTY")="on" then %>
									<th width="<%=colWidth%>" class="right"><strong>&nbsp;Services&nbsp;<%= getHotWord(70)%></strong></th>
						<%	end if %>
									<th width="<%=colWidth%>" class="right"><strong>&nbsp;Services&nbsp;<%= getHotWord(22)%></strong></th>
					<%	end if %>
					<%	if prodServ<>"0" then ' Services Only %>
						<%	if NOT request.form("optHideQTY")="on" then %>
									<th width="<%=colWidth%>" class="right"><strong>&nbsp;Products&nbsp;<%= getHotWord(70)%></strong></th>
						<%	end if %>
									<th width="<%=colWidth%>" class="right"><strong>&nbsp;Products&nbsp;<%= getHotWord(22)%></strong></th>
					<%	end if %>
					<%	if prodServ="" or prodServ="2" then ' Products and Services %>
						<%	if NOT request.form("optHideQTY")="on" then %>
									<th width="<%=colWidth%>" class="right"><strong>&nbsp;<%= getHotWord(22)%>&nbsp;<%= getHotWord(70)%></strong></th>
						<%	end if %>
									<th width="<%=colWidth%>" class="right"><strong>&nbsp;Sales&nbsp;<%= getHotWord(22)%></strong></th>
					<%	end if %>
									<th width="<%=colWidth%>"><strong>&nbsp;%&nbsp;</strong></th>
								</tr>
					<%	rangeSalesTotal = rsEntry("SalesTotal")
						rangeProductsTotal = rsEntry("SalesProductsTotal")
						rangeServicesTotal = rsEntry("SalesServicesTotal")
						do while NOT rsEntry.EOF
							' Print Totals for Last Client, then new client header %>
								<tr>
								<% if request.form("frmExpReport")="true" then %>
									<td class="mainText"><%=rsEntry("RSSID")%></td>
								<% end if %>
								<% if NOT request.form("frmExpReport")="true" then %>
									<td class="mainText"><a href="adm_clt_purch.asp?ID=<%=rsEntry("ClientID")%>"><%=rsEntry("LastName")%>, <%=rsEntry("FirstName")%></a></td>
								<% else %>
									<td class="mainText"><%=rsEntry("LastName")%>, <%=rsEntry("FirstName")%></td>
								<% end if %>
						<%	if prodServ<>"1" then ' NOT Products Only %>
							<%	if NOT request.form("optHideQTY")="on" then %>
									<td class="mainText right"><%=rsEntry("ServicesQTY")%></td>
							<%	end if %>
								<% if NOT request.form("frmExpReport")="true" then %>
									<td class="mainText right"><strong><%=FmtCurrency(rsEntry("ServicesTotal"))%></strong></td>
								<% else %>
									<td class="mainText right"><strong><%=FmtNumber(rsEntry("ServicesTotal"))%></strong></td>
								<% end if %>
						<%	end if %>
						<%	if prodServ<>"0" then ' NOT Services Only %>
							<%	if NOT request.form("optHideQTY")="on" then %>
									<td class="mainText right"><%=rsEntry("ProductQTY")%></td>
							<%	end if %>
								<% if NOT request.form("frmExpReport")="true" then %>
									<td class="mainText right"><strong><%=FmtCurrency(rsEntry("ProductTotal"))%></strong></td>
								<% else %>
									<td class="mainText right"><strong><%=FmtNumber(rsEntry("ProductTotal"))%></strong></td>
								<% end if %>
						<%	end if %>
						<%	if prodServ="" or prodServ="2" then ' Products and Services %>
							<%	if NOT request.form("optHideQTY")="on" then %>
									<td class="mainText right"><%=rsEntry("SalesQTY")%></td>
							<%	end if %>
								<% if NOT request.form("frmExpReport")="true" then %>
									<td class="mainText right"><strong><%=FmtCurrency(rsEntry("SaleTotal"))%></strong></td>
								<% else %>
									<td class="mainText right"><strong><%=FmtNumber(rsEntry("SaleTotal"))%></strong></td>
								<% end if %>
						<%	end if %>
									<td width="<%=colWidth%>" class="mainText" class="right">&nbsp;<% if rangeSalesTotal<>0 then response.write FmtNumber(rsEntry("SaleTotal") * 100 / rangeSalesTotal) & "%" else response.write "---" end if %>&nbsp;</td>
								</tr>
					<%		tmpProductQTY = tmpProductQTY + rsEntry("ProductQTY")
							tmpServiceQTY = tmpServiceQTY + rsEntry("ServicesQTY")
							tmpServicesTotal = tmpServicesTotal + rsEntry("ServicesTotal")
							tmpProductTotal = tmpProductTotal + rsEntry("ProductTotal")
							tmpSalesQTY = tmpSalesQTY + rsEntry("SalesQTY")
							tmpSalesTotal = tmpSalesTotal + rsEntry("SaleTotal")
							rowCount = rowCount + 1
							rsEntry.MoveNext
						loop
					%>
							</table>
							<br />
							<table class="mainText" width="100%"  cellspacing="0">
								<tr>
							<% if request.form("frmExpReport")="true" then %>
							  	<td>&nbsp;</td>
							<% end if %>
								  <td width="22%"><strong><%= getHotWord(22)%></strong>:</td>
							  
					<%	if prodServ<>"1" then ' Products Only %>
						<%	if NOT request.form("optHideQTY")="on" then %>
								  <td width="<%=colWidth%>" class="right"><%=tmpServiceQTY%></td>
						<%	end if %>
						<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtCurrency(tmpServicesTotal)%></strong></td>
						<% else %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtNumber(tmpServicesTotal)%></strong></td>
						<% end if %>
					<%	end if %>
					<%	if prodServ<>"0" then ' Services Only %>
						<%	if NOT request.form("optHideQTY")="on" then %>
								  <td width="<%=colWidth%>" class="right"><%=tmpProductQTY%></td>
						<%	end if %>
						<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtCurrency(tmpProductTotal)%></strong></td>
						<% else %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtNumber(tmpProductTotal)%></strong></td>
						<% end if %>
					<%	end if %>
					<%	if prodServ="" or prodServ="2" then ' Products and Services %>
						<%	if NOT request.form("optHideQTY")="on" then %>
								  <td width="<%=colWidth%>" class="right"><%=tmpSalesQTY%></td>
						<%	end if %>
						<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtCurrency(tmpSalesTotal)%></strong></td>
						<% else %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtNumber(tmpSalesTotal)%></strong></td>
						<% end if %>
					<%	end if %>
									<td width="<%=colWidth%>" class="right">&nbsp;&nbsp;</td>
								</tr>
								<tr>
							<% if request.form("frmExpReport")="true" then %>
							  	<td>&nbsp;</td>
							<% end if %>
								  <td width="22%"><strong>Total Revenue</strong>:</td>
							  
					<%	if prodServ<>"1" then ' Products Only %>
						<%	if NOT request.form("optHideQTY")="on" then %>
								  <td width="<%=colWidth%>" class="right">&nbsp;</td>
						<%	end if %>
						<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtCurrency(rangeServicesTotal)%></strong></td>
						<% else %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtNumber(rangeServicesTotal)%></strong></td>
						<% end if %>
					<%	end if %>
					<%	if prodServ<>"0" then ' Services Only %>
						<%	if NOT request.form("optHideQTY")="on" then %>
								  <td width="<%=colWidth%>" class="right">&nbsp;</td>
						<%	end if %>
						<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtCurrency(rangeProductsTotal)%></strong></td>
						<% else %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtNumber(rangeProductsTotal)%></strong></td>
						<% end if %>
					<%	end if %>
					<%	if prodServ="" or prodServ="2" then ' Products and Services %>
						<%	if NOT request.form("optHideQTY")="on" then %>
								  <td width="<%=colWidth%>" class="right">&nbsp;</td>
						<%	end if %>
						<% if NOT request.form("frmExpReport")="true" then %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtCurrency(rangeSalesTotal)%></strong></td>
						<% else %>
								  <td width="<%=colWidth%>" class="right"><strong><%=FmtNumber(rangeSalesTotal)%></strong></td>
						<% end if %>
					<%	end if %>
									<td width="<%=colWidth%>" class="right">&nbsp;&nbsp;</td>
								</tr>
								<tr>
							<% if request.form("frmExpReport")="true" then %>
							  	<td>&nbsp;</td>
							<% end if %>
								  <td width="22%"><strong>% of Total Revenue</strong>:</td>
							  
					<%	if prodServ<>"1" then ' Products Only %>
						<%	if NOT request.form("optHideQTY")="on" then %>
								  <td width="<%=colWidth%>" class="right">&nbsp;</td>
						<%	end if %>
								  <td width="<%=colWidth%>" class="right"><strong><% if rangeServicesTotal<>0 then response.write FmtNumber(tmpServicesTotal * 100 / rangeServicesTotal) & "%" else response.write "---" end if %></strong></td>
					<%	end if %>
					<%	if prodServ<>"0" then ' Services Only %>
						<%	if NOT request.form("optHideQTY")="on" then %>
								  <td width="<%=colWidth%>" class="right">&nbsp;</td>
						<%	end if %>
								  <td width="<%=colWidth%>" class="right"><strong><% if rangeProductsTotal<>0 then response.write FmtNumber(tmpProductTotal * 100 / rangeProductsTotal) & "%" else response.write "---" end if %></strong></td>
					<%	end if %>
					<%	if prodServ="" or prodServ="2" then ' Products and Services %>
						<%	if NOT request.form("optHideQTY")="on" then %>
								  <td width="<%=colWidth%>" class="right">&nbsp;</td>
						<%	end if %>
								  <td width="<%=colWidth%>" class="right"><strong><% if rangeSalesTotal<>0 then response.write FmtNumber(tmpSalesTotal * 100 / rangeSalesTotal) & "%" else response.write "---" end if %></strong></td>
					<%	end if %>
									<td width="<%=colWidth%>" class="right">&nbsp;&nbsp;</td>
								</tr>
				<%	end if '
				end if ' request.form("frmTagClients")="true"
			else ' Detail View
				if NOT rsEntry.EOF then
				%>
				
				<% ' First print the first client's header %>
					<tr class="right whiteSmallText" style="background-color:<%=Session("pageColor4")%>;">
					  <td width="12%" align="left">&nbsp;<strong><a class="whiteSmallText" href="adm_clt_purch.asp?ID=<%=rsEntry("ClientID")%>"><%=rsEntry("LastName")%>, <%=rsEntry("FirstName")%></a></strong></td>
					  <td nowrap width="12%" align="left"><strong><%= getHotWord(66)%></strong></td>
					  <td nowrap width="40%" class="center-ch"><strong><%= getHotWord(65)%></strong></td>
					  <td nowrap align="left" width="12%"><strong><%= getHotWord(8)%></strong></td>
					  <td nowrap width="12%"><strong><%= getHotWord(70)%></strong></td>
					  <td nowrap width="12%"><strong>Sales&nbsp;<%= getHotWord(22)%></strong></td>
					</tr>
					<% if NOT request.form("frmExpReport")="true" then %>
					<tr height="1">
						  <td colspan="6" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
					</tr>
					<% end if %>
				<%	if prodServ="" or prodServ="2" then ' Products and Services %>
					<tr>
					<% if rsEntry("IsProduct")="1" then %>
						  <td colspan="6" align="left"><b>Products</b></td>
					<% else %>
						  <td colspan="6" align="left"><b>Services</b></td>
					<% end if %>
					</tr>
				<% end if %>
				<%	tmpClientID = rsEntry("ClientID") %> 
				<%	tmpIsProduct = rsEntry("IsProduct") %> 
				<%	
					do while NOT rsEntry.EOF
						if (cstr(tmpClientID) <> cstr(rsEntry("ClientID"))) then
							' Print Totals for Last Client, then new client header
				%>
						<%	if tmpSubTotQty <> tmpTotQty or tmpSubSaleTotal <> tmpSaleTotal then 
								'Don't print this if the subtotal is the same as the totals  %>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="6" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% end if %>								
								<tr>
								  <td colspan="4" nowrap class="right"><b><%= getHotWord(118)%></b>:</td>	
								  <td class="right"><%=tmpSubTotQty%></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><%=FmtCurrency(tmpSubSaleTotal)%></strong></td>
								<% else %>
								  <td class="right"><strong><%=FmtNumber(tmpSubSaleTotal)%></strong></td>
								<% end if %>
								</tr>
								
						<%
							end if
							tmpSubTotQty = 0
							tmpSubSaleTotal = 0
						%>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="6" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% else %>
								<tr height="1">
									<td colspan="6">&nbsp;</td>
								</tr>
								<% end if %>
								<tr>
								  <td colspan="4" nowrap class="right"><b><%= getHotWord(22)%></b>:</td>	
								  <td class="right"><%=tmpTotQty%></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><%=FmtCurrency(tmpSaleTotal)%></strong></td>
								<% else %>
								  <td class="right"><strong><%=FmtNumber(tmpSaleTotal)%></strong></td>
								<% end if %>
								</tr>
								<tr height="10">
									<td colspan="6">&nbsp;</td>
								</tr>

				<%
							' Print header for new client
				%>
								<tr class="right whiteSmallText" style="background-color:<%=Session("pageColor4")%>;">
								  <td align="left">&nbsp;<strong><a class="whiteSmallText" href="adm_clt_purch.asp?ID=<%=rsEntry("ClientID")%>"><%=rsEntry("LastName")%>, <%=rsEntry("FirstName")%></a></strong></td>
								  <td nowrap align="left"><strong><%= getHotWord(66)%></strong></td>
								  <td nowrap class="center-ch"><strong><%= getHotWord(65)%></strong></td>
								  <td nowrap align="left"><strong><%= getHotWord(8)%></strong></td>
								  <td><strong><%= getHotWord(70)%></strong></td>
								  <td><strong>Sales&nbsp;<%= getHotWord(22)%></strong></td>
								</tr>

								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="6" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% end if %>
						<%	if prodServ="" or prodServ="2" then ' Products and Services %>
								<tr>
								<% if rsEntry("IsProduct")="1" then %>
									  <td colspan="6" align="left"><b>Products</b></td>
								<% else %>
									  <td colspan="6" align="left"><b>Services</b></td>
								<% end if 
								   printedProdServ = true
								%>
								</tr>
						<% end if %>
				<%
							tmpTotQty = 0
							tmpSaleTotal = 0

							tmpClientID = rsEntry("ClientID")
							
							
						end if ' end New client (old footer, new header)
				%>
							<% if tmpIsProduct <> rsEntry("IsProduct") and NOT printedProdServ then %>
								
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="6" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% end if %>								
								<tr>
								  <td colspan="4" nowrap class="right"><b><%= getHotWord(118)%></b>:</td>	
								  <td class="right"><%=tmpSubTotQty%></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><%=FmtCurrency(tmpSubSaleTotal)%></strong></td>
								<% else %>
								  <td class="right"><strong><%=FmtNumber(tmpSubSaleTotal)%></strong></td>
								<% end if %>
								</tr>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="6" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% end if 
		   							tmpSubTotQty = 0
									tmpSubSaleTotal = 0
								%>
								<tr>
								<% if rsEntry("IsProduct")="1" then %>
									  <td colspan="6" align="left"><b>Products</b></td>
								<% else %>
									  <td colspan="6" align="left"><b>Services</b></td>
								<% end if %>
								</tr>
							<% end if %>
								<tr class="right">
								  <td align="left"><a href="adm_tlbx_voidedit.asp?saleno=<%=rsEntry("SaleID")%>"><%=Right(rsEntry("SaleID"),4)%></a></td>
								  <td align="left"><%=rsEntry("SaleDate")%></td>
								  <td align="left">&nbsp;&nbsp;<%=rsEntry("Description")%></td>
								  <td align="left"><%=rsEntry("LocationName")%></td>
								  <td><%=rsEntry("Quantity")%></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td><%=FmtCurrency(rsEntry("SDTotal"))%></td>
								<% else %>
								  <td><%=FmtNumber(rsEntry("SDTotal"))%></td>
								<% end if %>
								</tr>
			
				<%
						tmpSubTotQty = tmpSubTotQty + rsEntry("Quantity")
						tmpSubSaleTotal = tmpSubSaleTotal + rsEntry("SDTotal")
						tmpIsProduct = rsEntry("IsProduct")
						tmpSaleTotal = rsEntry("SaleTotal")
						tmpTotQty = rsEntry("SalesQTY")
						rsEntry.MoveNext
						printedProdServ = false
					loop
				end if
			
				if tmpClientID<>"" then
	
			%>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="6" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% else %>
								<tr height="1">
									<td colspan="6">&nbsp;</td>
								</tr>
								<% end if %>
								<tr>
								  <td colspan="4" nowrap class="right"><b><%= getHotWord(22)%></b>:</td>	
								  <td class="right"><%=tmpTotQty%></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td class="right"><strong><%=FmtCurrency(tmpSaleTotal)%></strong></td>
								<% else %>
								  <td class="right"><strong><%=FmtNumber(tmpSaleTotal)%></strong></td>
								<% end if %>
								</tr>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="6" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% else %>
								<tr height="1">
									<td colspan="6">&nbsp;</td>
								</tr>
								<% end if %>
								<tr height="10">
									<td colspan="6">&nbsp;</td>
								</tr>
			<%	end if	' Last footer
			end if ' Summary vs Detail view
			rsEntry.close
			set rsEntry = nothing

		end if 	'First Load
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
