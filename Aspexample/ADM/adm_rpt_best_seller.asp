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
	%>
	<!-- #include file="inc_accpriv.asp" -->
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
			<!-- #include file="inc_rpt_tagging.asp" -->
			<!-- #include file="inc_utilities.asp" -->
			<!-- #include file="inc_rpt_save.asp" -->
			<%	dim doRefresh : doRefresh = false %>
			<!-- #include file="inc_date_arrows.asp" -->
			<!-- #include file="../inc_ajax.asp" --> 
			<!-- #include file="../inc_val_date.asp" --> 
			<!-- #include file="inc_hotword.asp" -->
		<%
		Dim showDetails, cSDate, cEDate, tmpCurDate, rsPrice, rsQty, rsDiscount, rsTax, rsPaid, cLoc, tmpSubTotQty, tmpSubSaleTotal, tmpTotQty, tmpIsProduct, tmpMargin, tmpCost, tmpCategory
		Dim tmpPMTotal, rowColor, cashBasis, splitFactorA, splitFactorB, tmpClientName, rptSummary, tmpSaleTotal, tmpQTY
		Dim tmpClientID, tmpCatName, tmpSubCat, rsTmpSubCat, tmpTotSub, tmpTotDiscount, tmpTotTax, tmpTotal, tmpTotalDisc, tmpTotCashEq, tmpTotNonCashEq
		Dim TotSub, TotDiscount, TotTax, GrandTotal, TotCashEq, TotNonCashEq, GrandTotalDisc, dd_paymeth, printedProdServ, tmpCOGSTotal, tmpCOGSSubtotal
		Dim tmpProductQTY, tmpServiceQTY, tmpServicesTotal, tmpProductTotal, tmpSalesQTY, tmpSalesTotal, rowCount, prodServ, ap_view_all_locs, MarginTooltip
		
		MarginToolTip = "Net profit divided by gross sales"
		
		ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
		
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
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_best_seller")) %>
			<script type="text/javascript">
			function exportReport() {
				document.frmSales.frmExpReport.value = "true";
				document.frmSales.frmGenReport.value = "true";
				<% iframeSubmit "frmSales", "adm_rpt_best_seller.asp" %>
			}
			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
		<%
		end if
		
		%>
		
        <style type="text/css" >
        .tooltip {
            cursor:default;
            }
        </style>

		<% if NOT request.form("frmExpReport")="true" then %>
<% pageStart %>
			<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
			<tr> 
			  <td valign="top" height="100%" width="100%"> 
				<table class="center" cellspacing="0" width="90%" height="100%">
				  <tr>
					<td class="headText" align="left" valign="top">
					  <table width="100%" cellspacing="0">
						<tr>
						  <td class="headText" valign="bottom"><b> <%= pp_PageTitle("Best Sellers") %></b>
						   <%if session("Admin")="sa" then %>
                                 <a class="mainText" href="/Report/Sales/BestSellers">Current version</a>                                 
                           <%end if %>
						  </td>
						  <td valign="bottom" class="right" height="26"> </td>
						</tr>
					  </table>
					</td>
				  </tr>
				  <tr> 
					<td height="30"  valign="bottom" class="headText">
						<table class="mainText border4 center" cellspacing="0">
						  <form name="frmSales" action="adm_rpt_best_seller.asp" method="POST">
							<input type="hidden" name="frmGenReport" value="">
							<input type="hidden" name="frmExpReport" value="">
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
								&nbsp;Top: <input type="text" size="3" value="<% if request.form("optTopNum")="" then %>20<% else response.write(request.form("optTopNum")) end if%>" name="optTopNum">
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
								&nbsp;Which earned more than: $<input type="text" name="optMinValue" size="3" value="<% if request.form("optMinValue")<>"" then response.write(request.form("optMinValue")) else response.write("0.00") end if %>">
								&nbsp;Sort By:
								<select name="optSortBy" onChange="singleSubmit()">
									<option value="0"<% if request.form("optSortBy")="0" then %> selected<% end if %>>Sales Total</option>
									<option value="1"<% if request.form("optSortBy")="1" then %> selected<% end if %>>Product Description</option>
									<option value="2"<% if request.form("optSortBy")="2" then %> selected<% end if %>>Quantity Sold</option>
									<option value="3"<% if request.form("optSortBy")="3" then %> selected<% end if %>>Category</option>
								</select>
								
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
								<select name="optBasis">
								  <option value="0" <%if request.form("optBasis")="0" then response.write "selected" end if%>>Accrual Basis</option>
								  <option value="1" <%if request.form("optBasis")="1" then response.write "selected" end if%>>Cash Basis</option>
								</select>
<% if request.form("optProdServ")="1" then %>
								<br />
						<select name="optSupplier"><option value="0">All Suppliers</option>
						<%
							strSQL = "SELECT Suppliers.SupplierID, Suppliers.CompanyName FROM Suppliers "
							'strSQL = strSQL & " INNER JOIN PRODUCTS ON PRODUCTS.SupplierID = Suppliers.SupplierID INNER JOIN [Sales Details] ON [Sales Details].ProductID = Products.ProductID INNER JOIN Sales ON Sales.SaleID=[Sales Details].SaleID  "
							'strSQL = strSQL & "WHERE ([Sales Details].CategoryID >25) AND Suppliers.Active = 1 and Suppliers.[Delete] = 0 AND "
							'strSQL = strSQL & "(Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") "
							'strSQL = strSQL & " GROUP BY Suppliers.SupplierID, Suppliers.CompanyName "
							strSQL = strSQL & "WHERE Suppliers.Active = 1 and Suppliers.[Delete] = 0 "
							strSQL = strSQL & "ORDER BY Suppliers.SupplierID, Suppliers.CompanyName "
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>	
									<option value="<%=rsEntry("SupplierID")%>" <%if request.form("optSupplier")=CSTR(rsEntry("SupplierID")) then response.write "selected" end if%>><%=rsentry("CompanyName")%></option>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
						</select>
						&nbsp;
						<select name="optColor"><option value="0">All Colors</option>
						<%
							strSQL = "SELECT Colors.ColorID, Colors.ColorName FROM Colors "
							'strSQL = strSQL & " INNER JOIN [Sales Details] ON [Sales Details].ColorID = Colors.ColorID INNER JOIN Sales ON Sales.SaleID=[Sales Details].SaleID  "
							'strSQL = strSQL & "WHERE ([Sales Details].CategoryID >25) AND Colors.Active = 1 AND Colors.[Delete] = 0 AND "
							'strSQL = strSQL & "(Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") "
							'strSQL = strSQL & " GROUP BY Colors.ColorID, Colors.ColorName "
							strSQL = strSQL & "WHERE Colors.Active = 1 AND Colors.[Delete] = 0 "
							strSQL = strSQL & "ORDER BY Colors.SortOrderID, Colors.ColorName "
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>	
									<option value="<%=rsEntry("ColorID")%>" <%if request.form("optColor")=CSTR(rsEntry("ColorID")) then response.write "selected" end if%>><%=rsentry("ColorName")%></option>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
						</select>
						&nbsp;
						<select name="optSize"><option value="0">All Sizes</option>
						<%
							strSQL = "SELECT Sizes.SizeID, Sizes.SizeName FROM Sizes "
							'strSQL = strSQL & " INNER JOIN [Sales Details] ON [Sales Details].SizeID = Sizes.SizeID INNER JOIN Sales ON Sales.SaleID=[Sales Details].SaleID  "
							'strSQL = strSQL & "WHERE ([Sales Details].CategoryID >25) AND Sizes.Active = 1 and Sizes.[Delete] = 0 AND "
							'strSQL = strSQL & "(Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") "
							'strSQL = strSQL & " GROUP BY Sizes.SizeID, Sizes.SizeName "
							strSQL = strSQL & "WHERE Sizes.Active = 1 and Sizes.[Delete] = 0 "
							strSQL = strSQL & "ORDER BY Sizes.SortOrderID, Sizes.SizeName "
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>	
									<option value="<%=rsEntry("SizeID")%>" <%if request.form("optSize")=CSTR(rsEntry("SizeID")) then response.write "selected" end if%>><%=rsentry("SizeName")%></option>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
						</select>
						&nbsp;
						<input type="checkbox" name="optColorSizeSplit" <% if request.form("optColorSizeSplit")="on" then response.write " checked" end if %>> Detailed Color/Size Info

</span>	
<% end if ' optProdServ="1" %>
								<br />
								<% taggingFilter %><img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" width="14" height="15" title="Runs the report looking only at sales data from tagged clients." align="middle">
								<br />
								<% showDateArrows("frmSales") %>
								<input type="button" name="Button" value="Generate" onClick="genReport();">
								<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
								else%>
									<% exportToExcelButton %>
								<% end if %>
							<% savingButtons "frmSales", "Best Sellers" %>
							</b>&nbsp;
							</td>
							</tr>
							  </form>
					  </table>
					</td>
				  </tr>
				  <tr> 
					<td valign="top" class="mainTextBig center-ch"> 
					  <table class="mainText" width="95%" cellspacing="0">
						<tr >
						  <td  colspan="2" valign="top" class="mainTextBig center-ch">&nbsp;</td>
						</tr>
						<tr > 
						  <td  colspan="2" valign="top" class="mainTextBig center-ch">
		<% end if			'end of frmExpreport value check before /head line	  %>
		
		<table class="mainText" width="100%"  cellspacing="0">
		<% if request.form("frmGenReport")="true" then
			if request.form("frmExpReport")="true" then
				Dim stFilename
				stFilename="attachment; filename=Best Seller Report " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
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

			if request.form("optSummary") = "1" then
				rptSummary = true
			else
				rptSummary = false
			end if
			
			if rptSummary then 
				strSQL = "SELECT "

				if request.form("optTopNum")<>"" and CINT(request.form("optTopNum")) >= 1 then
					strSQL = strSQL & " TOP " & request.form("optTopNum")
				end if

				strSQL = strSQL & " Products.ProductID, Products.Description, Products.OurCost, Categories.CategoryName,"
				strSQL = strSQL & " Products.ColorID, Products.SizeID, Colors.ColorName, Sizes.SizeName, "
				strSQL = strSQL & " SUM(SDP.Quantity) as SalesQTY, SUM(SDP.SDPaymentAmount) as SaleTotal, "
				
				' Profit Margin
				
			    ' 55_3286, Updated margin calc: AVG((Price - Discount - Cost)/(Price - Discount), CCP 10/20/09
				strSQL = strSQL & " AVG((((ISNULL(SDP.UnitPrice, 0) - (ISNULL(SDP.DiscAmt / (CASE WHEN SDP.Quantity = 0 THEN NULL ELSE SDP.Quantity END), 0)) - CASE WHEN (PRODUCTS.OurCost = 0 OR PRODUCTS.OurCost IS NULL) THEN 1 ELSE (PRODUCTS.OurCost) END)) / (CASE WHEN(ISNULL(SDP.UnitPrice, 0) - (ISNULL(SDP.DiscAmt / (CASE WHEN SDP.Quantity = 0 THEN NULL ELSE SDP.Quantity END), 0)))=0 THEN NULL ELSE (ISNULL(SDP.UnitPrice, 0) - (ISNULL(SDP.DiscAmt / (CASE WHEN SDP.Quantity = 0 THEN NULL ELSE SDP.Quantity END), 0))) END))) AS Margin, "
				' Cost of Goods Sold (COGS)
				strSQL = strSQL & "SUM([Sales Details].Quantity * ISNULL(PRODUCTS.OurCost, 0)) as COGS, "
				
				strSQL = strSQL & "SUM(CASE WHEN [Sales Details].CategoryID<=20 THEN SDP.Quantity ELSE 0 END) as ProductQTY "
				strSQL = strSQL & ", SUM(CASE WHEN [Sales Details].CategoryID>=25 THEN SDP.Quantity ELSE 0 END) as ServicesQTY "
				strSQL = strSQL & ", SUM(CASE WHEN [Sales Details].CategoryID>=25 THEN SDP.SDPaymentAmount ELSE 0 END) as ServicesTotal "
				strSQL = strSQL & ", SUM(CASE WHEN [Sales Details].CategoryID<=25 THEN SDP.SDPaymentAmount ELSE 0 END) as ProductTotal "

				strSQL = strSQL & "FROM [Sales Details] INNER JOIN Colors ON Colors.ColorID = [Sales Details].ColorID INNER JOIN Sizes ON Sizes.SizeID = [Sales Details].SizeID INNER JOIN Sales ON [Sales Details].SaleID = Sales.SaleID INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID INNER JOIN Categories ON [Sales Details].CategoryID = Categories.CategoryID "

					strSQL = strSQL & " INNER JOIN ( " 
					strSQL = strSQL & " SELECT [Sales Details].SDID, [Sales Details].UnitPrice, [Sales Details].DiscAmt, [Sales Details].Quantity, SUM(tblSDPayments.SDPaymentAmount) as SDPaymentAmount "
					strSQL = strSQL & " FROM [Sales Details] INNER JOIN tblSDPayments ON [Sales Details].SDID = tblSDPayments.SDID "
					strSQL = strSQL & " INNER JOIN tblPayments ON tblPayments.PaymentID = tblSDPayments.PaymentID "
					strSQL = strSQL & " INNER JOIN [Payment Types] ON tblPayments.PaymentMethod = [Payment Types].Item# "
					if cashBasis then
						strSQL = strSQL & " WHERE [Payment Types].[CashEQ]=1 AND [Sales Details].CategoryID<>21 "
					else ''Accrual Basis
						strSQL = strSQL & " WHERE (NOT ([Sales Details].CategoryID BETWEEN 21 AND 23)) "
					end if
					strSQL = strSQL & " GROUP BY [Sales Details].SDID, [Sales Details].UnitPrice, [Sales Details].DiscAmt, [Sales Details].Quantity "
					strSQL = strSQL & " ) SDP ON [Sales Details].SDID = SDP.SDID " 

				if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
					strSQL = strSQL & " INNER JOIN [Payment Data] ON  [Sales Details].PmtRefNo = [PAYMENT DATA].PmtRefNo "
				end if
				
				if request.form("optFilterTagged")="on" then
					strSQL = strSQL & "INNER JOIN tblClientTag ON Sales.ClientID = tblClientTag.clientID "
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
				if request.form("optColor")<>"" AND request.form("optColor")<>"0" and request.Form("optProdServ")="1" then
					strSQL = strSQL & " AND [Sales Details].ColorID = " & request.form("optColor")
				end if
				if request.form("optSize")<>"" AND request.form("optSize")<>"0" and request.Form("optProdServ")="1" then
					strSQL = strSQL & " AND [Sales Details].SizeID = " & request.form("optSize")
				end if
				if request.form("optSupplier")<>"" AND request.form("optSupplier")<>"0" and request.Form("optProdServ")="1" then
					strSQL = strSQL & " AND Products.SupplierID = " & request.form("optSupplier")
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
								
				strSQL = strSQL & " GROUP BY Products.ProductID, Products.Description, Products.OurCost "
				strSQL = strSQL & ", Products.ColorID, Products.SizeID, Colors.ColorName, Sizes.SizeName, Categories.CategoryName "

				strSQL = strSQL & " HAVING 1=1 " 
				if request.form("optMinValue")<>"" then
					strSQL = strSQL & " AND SUM(SDP.SDPaymentAmount) >= " & request.form("optMinValue") & " "
				end if

				if request.form("optSortBy")="0" then
					strSQL = strSQL & " ORDER BY SaleTotal DESC, Products.Description, SalesQTY DESC "
				elseif request.form("optSortBy")="1" then
					strSQL = strSQL & " ORDER BY Products.Description, SaleTotal DESC, SalesQTY DESC "
				elseif request.form("optSortBy")="2" then
					strSQL = strSQL & " ORDER BY SalesQTY DESC, SaleTotal DESC, Products.Description "
				else 
					strSQL = strSQL & " ORDER BY Categories.CategoryName, SaleTotal DESC, SalesQTY DESC "	
				end if
				
			else    'Detail View
				strSQL = "SELECT SALESTOTALS.SaleTotal, SALESTOTALS.Margin, SALESTOTALS.COGS, PRODUCTS.OurCost, SALESTOTALS.SalesQTY, [Sales Details].ProductID, CLIENTS.FirstName, CLIENTS.LastName, CLIENTS.ClientID, Categories.CategoryName, "
				strSQL = strSQL & " Products.Description, [Sales Details].Quantity, Sales.SaleID, Sales.SaleDate, Location.LocationName, Colors.ColorName, Sizes.SizeName "
				strSQL = strSQL & ", tblSDPayments.SDPaymentAmount as SDTotal "
				strSQL = strSQL & ", CASE WHEN ([Sales Details].CategoryID<=25) THEN 0 ELSE 1 END as IsProduct "
				
				strSQL = strSQL & "FROM [Payment Types] INNER JOIN tblPayments ON [Payment Types].Item# = tblPayments.PaymentMethod INNER JOIN tblSDPayments SDP ON tblPayments.PaymentID = SDP.PaymentID INNER JOIN [Sales Details] ON SDP.SDID = [Sales Details].SDID INNER JOIN Colors ON Colors.ColorID = [Sales Details].ColorID INNER JOIN Sizes ON Sizes.SizeID = [Sales Details].SizeID INNER JOIN Sales ON [Sales Details].SaleID = Sales.SaleID INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID INNER JOIN Categories ON [Sales Details].CategoryID = Categories.CategoryID INNER JOIN Location ON Location.LocationID = [Sales Details].Location INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID "
				if request.form("optTG")<>"" and request.form("optTG")<>"0" then 
					strSQL = strSQL & "INNER JOIN [Payment Data] ON  [Sales Details].PmtRefNo = [PAYMENT DATA].PmtRefNo "
				end if
				strSQL = strSQL & "INNER JOIN (SELECT [Sales Details].ProductID, SUM(CASE WHEN ((NOT (tblSDPayments.SDPaymentAmount) IS NULL)) THEN tblSDPayments.SDPaymentAmount ELSE 0 END) AS SaleTotal, "
				'Profit Margin, 55_3286, Updated margin calc: AVG((Price - Discount - Cost)/(Price - Discount), CCP 10/20/09
				strSQL = strSQL & "AVG((((ISNULL([Sales Details].UnitPrice, 0) - (ISNULL([Sales Details].DiscAmt / (CASE WHEN [Sales Details].Quantity = 0 THEN NULL ELSE [Sales Details].Quantity END), 0)) - CASE WHEN (PRODUCTS.OurCost = 0 OR PRODUCTS.OurCost IS NULL) THEN 1 ELSE (PRODUCTS.OurCost) END)) / (CASE WHEN(ISNULL([Sales Details].UnitPrice, 0) - (ISNULL([Sales Details].DiscAmt / (CASE WHEN [Sales Details].Quantity = 0 THEN NULL ELSE [Sales Details].Quantity END), 0)))=0 THEN NULL ELSE (ISNULL([Sales Details].UnitPrice, 0) - (ISNULL([Sales Details].DiscAmt / (CASE WHEN [Sales Details].Quantity = 0 THEN NULL ELSE [Sales Details].Quantity END), 0))) END))) AS Margin, "
				strSQL = strSQL & "SUM([Sales Details].Quantity * ISNULL(PRODUCTS.OurCost, 0)) AS COGS, SUM([Sales Details].Quantity) AS SalesQTY "
				strSQL = strSQL & "FROM Sales AS Sales INNER JOIN [Sales Details] AS [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN PRODUCTS AS PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID INNER JOIN tblSDPayments ON [Sales Details].SDID = tblSDPayments.SDID INNER JOIN tblPayments ON tblSDPayments.PaymentID = tblPayments.PaymentID INNER JOIN [Payment Types] ON tblPayments.PaymentMethod = [Payment Types].Item# "
				if request.form("optProdServ")="0" then
				    strSQL =strSQL & " INNER JOIN [PAYMENT DATA] ON [Sales Details].PmtRefNo = [PAYMENT DATA].PmtRefNo "
				end if
			    strSQL = strSQL & " WHERE (Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") "
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
				if request.form("optFilterTagged")="on" then
					if session("mvaruserID")<>"" then
						strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
					else
						strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
					end if
				end if
				strSQL = strSQL & " GROUP BY [Sales Details].ProductID  HAVING 1=1 "
				if request.form("optMinValue")<>"" then
					strSQL = strSQL & " AND SUM(tblSDPayments.SDPaymentAmount) >= " & request.form("optMinValue")
				end if
				strSQL = strSQL & ") AS SALESTOTALS ON [Sales Details].ProductID = SALESTOTALS.ProductID INNER JOIN tblSDPayments ON tblSDPayments.SDID = [Sales Details].SDID "
				
				
				strSQL = strSQL & "WHERE (Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") "
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
				if request.form("optSortBy")="0" then
					strSQL = strSQL & " ORDER BY SALESTOTALS.SaleTotal DESC, Products.Description, SALESTOTALS.SalesQTY DESC, IsProduct "
				elseif request.form("optSortBy")="1" then
					strSQL = strSQL & " ORDER BY Products.Description, SALESTOTALS.SaleTotal DESC, SALESTOTALS.SalesQTY DESC, IsProduct "
				elseif request.form("optSortBy")="2" then
				    strSQL = strSQL & " ORDER BY SALESTOTALS.SalesQTY DESC, SALESTOTALS.SaleTotal DESC, Products.Description, IsProduct "
				else  
					strSQL = strSQL & " ORDER BY CategoryName, SALESTOTALS.SaleTotal DESC, Products.Description, SALESTOTALS.SalesQTY DESC, IsProduct "
				end if
			end if

			rsEntry.CursorLocation = 3
		   response.write debugSQL(strSQL, "SQL") 
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			
		if rptSummary then ' Summary view
				rowCount = 1
				%>
					<tr class="right">
					</tr>
				<%


				if NOT rsEntry.EOF then
				%>
					<tr class="right">
					  <td align="left"><strong>Product</strong></td>
				<%	if prodServ<>"1" then ' Not Products Only %>
					  <td><strong>&nbsp;Services&nbsp;<%= getHotWord(70)%></strong></td>
					  <td><strong>&nbsp;Services&nbsp;<%= getHotWord(22)%></strong></td>
				<%	end if %>
				<%	if prodServ<>"0" then ' Not Services Only %>
					  <td><strong>&nbsp;Products&nbsp;<%= getHotWord(70)%></strong></td>
					  <td><strong>&nbsp;Products&nbsp;<%= getHotWord(22)%></strong></td>
					  <td><strong>&nbsp;<%= getHotWord(67)%></strong></td>
					  <td><strong>&nbsp;<%= getHotWord(75)%></strong></td>
				<%	end if %>
				<%	if prodServ="" or prodServ="2" then ' Products and Services %>
					  <td><strong>&nbsp;<%= getHotWord(22)%>&nbsp;<%= getHotWord(70)%></strong></td>
					  <td><strong>&nbsp;Sales&nbsp;<%= getHotWord(22)%></strong></td>
				<%	end if %>
				      <td><strong>&nbsp;Category</strong></td>
				      <td><strong>&nbsp;COGS</strong></td>
					  <td><span class="tooltip" title="<%=MarginTooltip %>"><strong>&nbsp;Margin</strong></span></td>						
					</tr>
				<%	if NOT request.form("frmExpReport")="true" then %>
					<tr height="1">
						  <td colspan="12" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
					</tr>
				<% 	end if %>
				<%	do while NOT rsEntry.EOF
						' Print Totals for Last Client, then new client header %>
							<tr>
							<% if NOT request.form("frmExpReport")="true" then %>
							  <td><%=rowCount%>.&nbsp;<%=rsEntry("Description")%></td>
							<% else %>
							  <td><%=rowCount%>.&nbsp;<%=rsEntry("Description")%></td>
							<% end if %>
				<%	if prodServ<>"1" then ' Products Only %>
							  <td class="right"><%=rsEntry("ProductQTY")%></td>
							<% if NOT request.form("frmExpReport")="true" then %>
							  <td class="right"><strong><%=FmtCurrency(rsEntry("ProductTotal"))%></strong></td>
							<% else %>
							  <td class="right"><strong><%=FmtNumber(rsEntry("ProductTotal"))%></strong></td>
							<% end if %>
				<%	end if %>
				<%	if prodServ<>"0" then ' Services Only %>
							  <td class="right"><%=rsEntry("ServicesQTY")%></td>
							<% if NOT request.form("frmExpReport")="true" then %>
							  <td class="right"><strong><%=FmtCurrency(rsEntry("ServicesTotal"))%></strong></td>
							<% else %>
							  <td class="right"><strong><%=FmtNumber(rsEntry("ServicesTotal"))%></strong></td>
							<% end if %>
							  <td class="right"><% if rsEntry("ColorName")<>"None" then response.write rsEntry("ColorName") else response.write "&nbsp;" end if %></td>
							  <td class="right"><% if rsEntry("SizeName")<>"None" then response.write rsEntry("SizeName") else response.write "&nbsp;" end if %></td>
				<%	end if %>
				<%	if prodServ="" or prodServ="2" then ' Products and Services %>
							  <td class="right"><%=rsEntry("SalesQTY")%></td>
							<% if NOT request.form("frmExpReport")="true" then %>
							  <td class="right"><strong><%=FmtCurrency(rsEntry("SaleTotal"))%></strong></td>
							<% else %>
							  <td class="right"><strong><%=FmtNumber(rsEntry("SaleTotal"))%></strong></td>
							<% end if %>

				<%	end if %>
				              <td class="right"><strong>&nbsp;<% if ( NOT IsNull(rsEntry("CategoryName")) AND rsEntry("CategoryName")<>"") then response.write rsEntry("CategoryName") else response.write "---" end if %></strong></td>
							  <td class="right"><strong>&nbsp;<% if ( NOT IsNull(rsEntry("OurCost")) AND rsEntry("OurCost")<>0) then response.write FmtCurrency(rsEntry("COGS")) else response.write "---" end if%></strong></td>
							  <td class="right"><strong>&nbsp;<% if NOT IsNull(rsEntry("Margin")) then response.write FmtNumber((rsEntry("Margin") * 100)) & "%" else response.write "---" end if %></strong></td>
							</tr>
							<% if NOT request.form("frmExpReport")="true" then %>
							<tr height="1">
								  <td colspan="12" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>								  
							</tr>
						<%	end if
					tmpProductQTY = tmpProductQTY + rsEntry("ProductQTY")
					tmpServiceQTY = tmpServiceQTY + rsEntry("ServicesQTY")
					tmpServicesTotal = tmpServicesTotal + rsEntry("ServicesTotal")
					tmpProductTotal = tmpProductTotal + rsEntry("ProductTotal")
					tmpSalesQTY = tmpSalesQTY + rsEntry("SalesQTY")
					tmpCOGSTotal = tmpCOGSTotal + rsEntry("COGS")
					tmpSalesTotal = tmpSalesTotal + rsEntry("SaleTotal")
					rowCount = rowCount + 1
					rsEntry.MoveNext
					loop
				%>
							<tr>
							  <td><strong><%= getHotWord(22)%></strong>:</td>
				<%	if prodServ<>"1" then ' Products Only %>
							  <td class="right"><%=tmpProductQTY%></td>
					<% if NOT request.form("frmExpReport")="true" then %>
							  <td class="right"><strong><%=FmtCurrency(tmpProductTotal)%></strong></td>
					<% else %>
							  <td class="right"><strong><%=FmtNumber(tmpProductTotal)%></strong></td>
					<% end if %>
				<%	end if %>
				<%	if prodServ<>"0" then ' Services Only %>
							  <td class="right"><%=tmpServiceQTY%></td>
					<% if NOT request.form("frmExpReport")="true" then %>
							  <td class="right"><strong><%=FmtCurrency(tmpServicesTotal)%></strong></td>
					<% else %>
							  <td class="right"><strong><%=FmtNumber(tmpServicesTotal)%></strong></td>
					<% end if %>
							  <td class="right" colspan="2">&nbsp;</td>
				<%	end if %>
				<%	if prodServ="" or prodServ="2" then ' Products and Services %>
							  <td class="right"><%=tmpSalesQTY%></td>
					<% if NOT request.form("frmExpReport")="true" then %>
							  <td class="right"><strong><%=FmtCurrency(tmpSalesTotal)%></strong></td>
					<% else %>
							  <td class="right"><strong><%=FmtNumber(tmpSalesTotal)%></strong></td>
					<% end if %>
				<%	end if %>
				              <td>&nbsp;</td>
				              <td class="right"><strong>&nbsp;<%=FmtCurrency(tmpCOGSTotal) %></strong></td>
							</tr>
			<%	end if '
			else ' Detail View

				if NOT rsEntry.EOF then
				%>
				
				<% ' First print the first client's header %>
					<tr class="right whiteSmallText" style="background-color:<%=Session("pageColor4")%>;">
					  <td width="10%" align="left">&nbsp;<strong><%=rsEntry("Description")%><% if rsEntry("ColorName")<>"None" then response.write ", " & rsEntry("ColorName") end if %><% if rsEntry("SizeName")<>"None" then response.write ", " & rsEntry("SizeName") else response.write "&nbsp;" end if %></strong>&nbsp;&nbsp;&nbsp;</td>
					  <td nowrap width="10%" align="left"><strong><%= getHotWord(66)%></strong></td>
					  <td nowrap width="28%" align="left">&nbsp;<strong><%=session("ClientHW")%>&nbsp;<%= getHotWord(40)%></strong></td>
					  <td nowrap align="left" width="12%"><strong><%= getHotWord(8)%></strong></td>
					  <td nowrap width="7%"><strong><%= getHotWord(70)%></strong></td>
					  <td nowrap width="7%"><strong>Sales&nbsp;<%= getHotWord(22)%></strong></td>
					  <td nowrap width="7%" class="right"><strong>Category</strong>&nbsp;</td>
					  <td nowrap width="7%" class="right"><strong>COGS</strong>&nbsp;</td>
					  <td nowrap width="7%" class="right"><span class="tooltip" title="<%=MarginTooltip %>"><strong>Margin</strong>&nbsp;</span></td>
					</tr>
					<% if NOT request.form("frmExpReport")="true" then %>
					<tr height="1">
						  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
					</tr>
					<% end if %>
				<%	if prodServ="" or prodServ="2" then ' Products and Services %>
					<tr>
					<% if rsEntry("IsProduct")="1" then %>
						  <!--td colspan="8" align="left"><b>Products</b></td-->
					<% else %>
						  <!--td colspan="8" align="left"><b>Services</b></td-->
					<% end if %>
					</tr>
				<% end if %>
				<%	tmpClientID = rsEntry("ProductID") %> 
				<%	tmpIsProduct = rsEntry("IsProduct") %> 
				<%	
					do while NOT rsEntry.EOF
						if (cstr(tmpClientID) <> cstr(rsEntry("ProductID"))) then
							' Print Totals for Last Product, then new product header
				%>
						<%	if tmpSubTotQty <> tmpTotQty or tmpSubSaleTotal <> tmpSaleTotal then 
								'Don't print this if the subtotal is the same as the totals  %>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
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
								  <td class="right"><strong><%=rsEntry("CategoryName")%></strong>&nbsp;</td>								
								  <td class="right"><strong><% if ( NOT IsNull(tmpCost) AND tmpCost<>0) then response.write FmtCurrency(tmpCOGSSubtotal) else response.write "---" end if %></strong>&nbsp;</td>
								  <td class="right"><strong><% if NOT IsNull(tmpMargin) then response.write FmtNumber(tmpMargin * 100) & "%" else response.write "---" end if %></strong>&nbsp;</td>
								</tr>
						<%
							end if
							tmpSubTotQty = 0
							tmpSubSaleTotal = 0
						%>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% else %>
								<tr height="1">
									<td colspan="10">&nbsp;</td>
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
								  <td class="right"><strong><% if ( NOT IsNull(tmpCategory) AND tmpCategory<>"") then response.write tmpCategory else response.write "---" end if %></strong>&nbsp;</td>								
								  <td class="right"><strong><% if ( NOT IsNull(tmpCost) AND tmpCost<>0) then response.write FmtCurrency(tmpCOGSSubtotal) else response.write "---" end if %></strong>&nbsp;</td>
								  <td class="right"><strong><% if  NOT IsNull(tmpMargin) then response.write FmtNumber(tmpMargin * 100) & "%" else response.write "---" end if %></strong>&nbsp;</td>
								</tr>
								<tr height="10">
									<td colspan="10">&nbsp;</td>
								</tr>

				<%
							' Print header for new client
				%>
								<tr class="right whiteSmallText" style="background-color:<%=Session("pageColor4")%>;">
								  <td width="12%" align="left">&nbsp;<strong><%=rsEntry("Description")%><% if rsEntry("ColorName")<>"None" then response.write ", " & rsEntry("ColorName") end if %><% if rsEntry("SizeName")<>"None" then response.write ", " & rsEntry("SizeName") else response.write "&nbsp;" end if %></strong>&nbsp;&nbsp;&nbsp;</td>
								  <td nowrap align="left"><strong><%= getHotWord(66)%></strong></td>
								  <td nowrap align="left">&nbsp;<strong><%=session("ClientHW")%> <%= getHotWord(40)%></strong></td>
								  <td nowrap align="left"><strong><%= getHotWord(8)%></strong></td>
								  <td><strong><%= getHotWord(70)%></strong></td>
								  <td><strong>Sales&nbsp;<%= getHotWord(22)%></strong></td>
								  <td><strong>Category</strong>&nbsp;</td>
								  <td><strong>COGS</strong>&nbsp;</td>
								  <td><span class="tooltip" title="<%=MarginTooltip %>"><strong>Margin</strong>&nbsp;</span></td>
								</tr>

								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% end if %>
						<%	if prodServ="" or prodServ="2" then ' Products and Services %>
								<tr>
								<% if rsEntry("IsProduct")="1" then %>
									  <!--td colspan="8" align="left"><b>Products</b></td-->
								<% else %>
									  <!--td colspan="8" align="left"><b>Services</b></td-->
								<% end if 
								   printedProdServ = true
								%>
								</tr>
						<% end if %>
				<%
							tmpTotQty = 0
							tmpSaleTotal = 0
							tmpCOGSSubtotal = 0

							tmpClientID = rsEntry("ProductID")
							
							
						end if ' end New product (old footer, new header)
				%>
							<% if tmpIsProduct <> rsEntry("IsProduct") and NOT printedProdServ then %>
								
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
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
								  <td class="right"><strong><% if ( NOT IsNull(tmpCost) AND tmpCost<>0) then response.write FmtNumber(tmpMargin * 100) & "%" else response.write "---" end if %></strong>&nbsp;</td>
								</tr>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% end if 
									tmpSubTotQty = 0
									tmpSubSaleTotal = 0
									tmpCOGSSubtotal = 0
								%>
								<tr>
								<% if rsEntry("IsProduct")="1" then %>
									  <td colspan="10" align="left"><b>Products</b></td>
								<% else %>
									  <td colspan="10" align="left"><b>Services</b></td>
								<% end if %>
								</tr>
							<% end if %>
								<tr class="right">
								  <td align="left"><a href="adm_tlbx_voidedit.asp?saleno=<%=rsEntry("SaleID")%>"><%=Right(rsEntry("SaleID"),4)%></a></td>
								  <td align="left"><%=rsEntry("SaleDate")%></td>
								  <td align="left">&nbsp;<% if rsEntry("ClientID")<>"1" then %><a href="adm_clt_purch.asp?ID=<%=rsEntry("ClientID")%>"><% end if %><%=rsEntry("LastName")%>, <%=rsEntry("FirstName")%><% if rsEntry("ClientID")<>"1" then %></a><% end if %></td>
								  <td align="left"><%=rsEntry("LocationName")%></td>
								  <td><%=rsEntry("Quantity")%></td>
								<% if NOT request.form("frmExpReport")="true" then %>
								  <td><%=FmtCurrency(rsEntry("SDTotal"))%></td>
								<% else %>
								  <td><%=FmtNumber(rsEntry("SDTotal"))%></td>
								<% end if %>
								  <!--<td><%=rsEntry("CategoryName") %></td>-->
								  <td class="right"><!--<% if ( NOT IsNull(rsEntry("OurCost")) AND rsEntry("OurCost")<>0) then response.write FmtNumber(rsEntry("Margin") * 100) & "%" else response.write "---" end if %>--></td>
								</tr>
			
				<%
						tmpSubTotQty = tmpSubTotQty + rsEntry("Quantity")
						tmpSubSaleTotal = tmpSubSaleTotal + rsEntry("SDTotal")
						tmpIsProduct = rsEntry("IsProduct")
						tmpSaleTotal = rsEntry("SaleTotal")
						tmpTotQty = rsEntry("SalesQTY")
						tmpMargin = rsEntry("Margin")
						tmpCategory = rsEntry("CategoryName")
						tmpCost = rsEntry("OurCost")
						tmpCOGSSubtotal = rsEntry("COGS")
						rsEntry.MoveNext
						printedProdServ = false
					loop
				else
				response.Write "NO RECORDS"
				end if
			
				if tmpClientID<>"" then
	
			%>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% else %>
								<tr height="1">
									<td colspan="10">&nbsp;</td>
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
								  <td class="right"><strong><% if ( NOT IsNull(tmpCategory) AND tmpCategory<>"") then response.write tmpCategory else response.write "---" end if %></strong></td>
								  <td class="right"><strong><% if ( NOT IsNull(tmpCost) AND tmpCost<>0) then response.write FmtCurrency(tmpCOGSSubtotal) else response.write "---" end if %></strong></td>
								  <td class="right"><strong><% if ( NOT IsNull(tmpCost) AND tmpCost<>0) then response.write FmtNumber(tmpMargin * 100) & "%" else response.write "---" end if %></strong>&nbsp;</td>
								</tr>
								<% if NOT request.form("frmExpReport")="true" then %>
								<tr height="1">
									  <td colspan="10" style="background-color:#666666;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"></td>
								</tr>
								<% else %>
								<tr height="1">
									<td colspan="10">&nbsp;</td>
								</tr>
								<% end if %>
								<tr height="10">
									<td colspan="10">&nbsp;</td>
								</tr>

			<%
				end if	' Last footer
			end if ' Summary vs Detail view

			rsEntry.close
			set rsEntry = nothing

		end if 	'First Load
		%>				
						  </table></td>
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
