<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")

dim phraseDictionary
set phraseDictionary = LoadPhrases("BusinessmodecloseoutdataPage", 149)

dim rsEntry, rsEntryB, rsEntry2, rsEntry2B, rsEntry3
set rsEntry     = Server.CreateObject("ADODB.Recordset")
set rsEntryB    = Server.CreateObject("ADODB.Recordset")
set rsEntry2    = Server.CreateObject("ADODB.Recordset")
set rsEntry2B   = Server.CreateObject("ADODB.Recordset")
set rsEntry3    = Server.CreateObject("ADODB.Recordset")
%>
<!-- #include file="inc_accpriv.asp" -->
<!-- #include file="inc_rpt_tagging.asp" -->
<%

if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_CASH_CLOSEOUT") then
%>
<script type="text/javascript">
	alert("<%=DisplayPhraseJS(systemMessagesErrorsDictionary,"Notauthorizedtoviewpage")%>");
	javascript: history.go(-1);
</script>
<%
else
%>
<!-- #include file="../inc_i18n.asp" -->
<%
	Public Function showCurrencyOrNumber(amount)
		if request.form("frmExpResults")="true" then
			RW FmtNumber(amount)
		else
			RW FmtCurrency(amount)
		end if
	End Function

	Dim ChangeSalesAmt, ChangeSalesAmtCash, ChangeSalesAmtCheck, DeferredSalesAmt, currencySymbol
	Dim StartingCash, TipsPaidOut
	Dim ChangeSalesTaxAmt
	Dim amtActCash, amtActCheck, amtCashOverShort, amtCheckOverShort, ss_Fitlink
	Dim SalesArr(), cloc, intCloseID, oldCloseDate, ap_view_all_locs

	ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
	ss_IncludeTips = checkStudioSetting("tblGenOpts", "IncludeTipsInPayroll")

	strSQL = "SELECT FitLink FROM Studios WHERE Studios.StudioID=" & session("StudioID")
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing

	ss_Fitlink = rsEntry("FitLink")
	rsEntry.close

	if ss_Fitlink then
%>
<!-- #include file="../inc_dbconn_fitlink.asp" -->
<%
	end if

	NewStartingCash   = 0
	amtCashOverShort  = 0
	amtCheckOverShort = 0
	ChangeSalesAmt    = 0
	TipsPaidOut       = 0

	currencySymbol = Replace(FmtCurrency(0), FmtNumber(0), "")

	if request.form("optLoc")<>"" then
		cLoc = CINT(request.form("optLoc"))
	else
		if session("numLocations")>1 then
			if session("UserLoc") <> 0 then
				cLoc = CINT(session("UserLoc"))
			else
				if session("curLocation") <> 0 then
					cLoc = CINT(session("curLocation"))
				else
					cLoc = 1
				end if
			end if
		else
			strSQL= "SELECT LocationID from Location WHERE wsShow = 1 "
			rsEntry.open strSQL, cnWS, 3
			if NOT rsEntry.EOF then
				cLoc = rsEntry("LocationID")
			else
				cLoc = 1
			end if 
			rsEntry.close
		end if
	end if

	Dim i
	i=0
	strSQL= "SELECT [Payment Types].[Item#], [Payment Types].[PmtTypes] FROM [Payment Types] "
	rsEntry.open strSQL, cnWS, 3
	Do while not rsEntry.eof
		ReDim Preserve SalesArr(4,i+1)
		SalesArr(0,i)=Cint(rsEntry("Item#"))
		SalesArr(1,i)=rsEntry("PmtTypes")
		i=i+1
		rsEntry.movenext
	Loop
	rsEntry.close

	Dim cEDate, dtLastClose, dateForDB

	strSQL=	"SELECT ISNULL(Max(tblClosedData.CloseDate), '1/1/1970') AS MaxOfCloseDate " &_
					"FROM tblClosedData " &_
					"WHERE (((tblClosedData.Location)=" & cloc & "))"
	rsEntry.open strSQL, cnWS, 3
	Do while not rsEntry.eof
		dtLastClose=rsEntry("MaxofCloseDate")
		oldCloseDate = rsEntry("MaxofCloseDate")
		Call SetLocale("en-us")
		rsEntry.movenext
	Loop
	rsEntry.close

	Call SetLocale(session("mvarLocaleStr"))

	if request.form("requiredtxtDateEnd")<>"" then
		if cDate(dtLastClose)>=cDate(request.form("requiredtxtDateEnd")) then
			cEDate = CDATE(request.form("requiredtxtDateEnd"))
			'This is for bug 15411, to keep track of multiple closeouts per day.
			'If they've changed the date on the form, we'll just append the current
			'time of day in order to differentiate this closeout from another
			'closeout they may make for the same date.
			dateForDB = CDATE(DateValue(cEDate) & " " & TimeValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
%>
<script type="text/javascript">
	//alert("Please select a date later than Last Close Date.");
    //document.frmCloseOut.frmShowResults.value = "false";
</script>
<%
		else
			cEDate = CDATE(request.form("requiredtxtDateEnd"))
			dateForDB = CDATE(DateValue(cEDate) & " " & TimeValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))

		end if
	else
		cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		dateForDB = CDATE(DateAdd("n", Session("tzOffset"), Now))
	end if
	Call SetLocale("en-us")

	if request.form("txtActualCash")<>"" then
		if isNumeric(request.form("txtActualCash")) then
			amtActCash = request.form("txtActualCash")
			amtCashOverShort = amtActCash - ChangeSalesAmtCash
		end if
	end if

	if request.form("txtActualCheck")<>"" then
		if isNumeric(request.form("txtActualCheck")) then
			amtActCheck = request.form("txtActualCheck")
			amtCheckOverShort = amtActCheck - ChangeSalesAmtCheck
		end if
	end if

	if request.form("StartingCash")<>"" then
		if isNumeric(request.form("StartingCash")) then
			amtActCheck = request.form("StartingCash")
		end if
	end if

	function setrowcolor()
		if rowCount = 0 then
			rowCount = 1
			setrowcolor = "#F2F2F2"
		else
			rowCount = 0
			setrowcolor = "#FAFAFA"
		end if
	end function

	if NOT request.form("frmExpResults")="true" then
%>
<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->

<%= js(array("mb", "calendar" & dateFormatCode, "adm/adm_rpt_cashdrawer_closeout")) %>
<script type="text/javascript">
	function exportReport() {
		document.frmCloseOut.frmExpResults.value = "true";
		document.frmCloseOut.frmShowResults.value = "true";
		<% iframeSubmit "frmCloseOut", "adm_rpt_cashdrawer_closeout.asp" %>
	}
</script>
<style>
.mbo-hidden-text {
	border-style:none;
	color:<%=session("pageColor3")%>;
	text-align:right;
}
</style>

<!-- #include file="../inc_date_ctrl.asp" -->
<!-- #include file="inc_help_content.asp" -->
<!-- #include file="../inc_ajax.asp" -->
<!-- #include file="../inc_val_date.asp" -->
<script type="text/javascript">
	function dateValidated() {
		document.frmCloseOut.submit();
	}
</script>
<%
	end if 'NOT request.form("frmExpResults")="true"

	if NOT request.form("frmExpResults")="true" then
%>
<% pageStart %>
<table id="table1" height="100%" width="<%=strPageWidth%>" cellspacing="0">
	<tr>
		<td valign="top" height="100%" width="100%">
			<form name="frmCloseOut" action="adm_rpt_cashdrawer_closeout.asp" method="POST">
			<input type="hidden" name="frmShowResults" value="" />
			<input type="hidden" name="frmExpResults" value="" />
			<input type="hidden" name="runCloseOut" value="" />
			<table id="table2" class="center" cellspacing="0" width="90%" height="100%">
				<tr>
					<td class="headText" align="left" valign="top">
						<table id="table3" width="100%" cellspacing="0">
							<tr>
								<td class="headText" valign="bottom">
									<b>
										<%=DisplayPhrase(pageTitlesDictionary,"Closeoutdata")%>
									</b>
									<%'JM - 49_2447%>
									<% showNewHelpContentIcon("daily-closeout-report") %>
								</td>
								<td valign="bottom" class="right" height="26">
								</td>
							</tr>
						</table> <!--#table3-->
					</td>
				</tr>
				<tr>
					<td valign="top" class="mainText right" height="6">
					</td>
				</tr>
				<tr>
					<td class="headText">
						<table id="table4" class="mainText border4 center" cellspacing="0">
							<tr>
								<td class="center-ch" valign="bottom" style="background-color: #F2F2F2;">
									<b>
										<span style="color: <%=session("pageColor4")%>;"></span>&nbsp;
<%
		if session("numLocations")>1 then
%>
										&nbsp;<%=xssStr(allHotWords(8))%>
										<select name="optLoc" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
<%
			strSQL = "SELECT LocationID, LocationName FROM Location WHERE Active=1 AND LocationID<>98 ORDER BY LocationName "
			rsEntry2.open strSQL, cnWS, 3
			do while not rsEntry2.EOF
				'if request.form("optLoc")="" then
%>
											<option value="<%=rsEntry2("LocationID")%>" <%if cLoc=CINT(rsEntry2("LocationID")) then response.write "selected" end if%>>
												<%=rsEntry2("LocationName")%>
											</option>
<%
				'else
%>
											<!-- <option value="<%=rsEntry2("LocationID")%>" <%if request.form("optLoc")=CSTR(rsEntry2("LocationID")) then response.write "selected" end if%>>
												<%=rsEntry2("LocationName")%>
											</option> -->
<%
				'end if
				rsEntry2.MoveNext
			loop
			rsEntry2.close
%>
										</select>
<%
		end if 'session("numLocations")>1
%>
										&nbsp;<%=DisplayPhrase(phraseDictionary,"Closedby")%>:&nbsp;<span style="color: <%=session("pageColor4")%>;"><%=Session("mvarNameFirst") & " " & Trim(Session("mvarNameLast"))%></span><br />
										&nbsp;<%=DisplayPhrase(phraseDictionary,"Lastclosedate")%>
										<input type="text" size="8" name="requiredtxtDateStart" value="<%=FmtDateShort(dtLastClose)%>" disabled />
										&nbsp;<%=DisplayPhrase(phraseDictionary,"Closedate")%>
										<input type="text" name="requiredtxtDateEnd" onblur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" value="<%=FmtDateShort(cEDate)%>" class="date" />
										<script type="text/javascript">
											var cal2 = new tcal({ 'formname': 'frmCloseOut', 'controlname': 'requiredtxtDateEnd' });
											cal2.a_tpl.yearscroll = true;
										</script>
										&nbsp;
										<input type="button" name="Button" value="<%=DisplayPhraseAttr(phraseDictionary,"Previewcloseamounts")%>" onclick="showReport();" />
									</b>
								</td>
							</tr>
						</table> <!--#table4-->
					</td>
				</tr>
				<tr>
					<td valign="top" class="mainText right" height="8"></td>
				</tr>
				<tr>
					<td valign="top" class="mainTextBig">
						<table id="table5" class="mainText" width="95%" cellspacing="0">
							<tr>
								<td class="mainTextBig" colspan="2" valign="top" align="left">
									<table id="table6" width="85%" cellspacing="0" style="margin: 0 auto;">
										<tr>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Pennies")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Pennies" value="<%if request.form("Pennies")<>"" then response.write(request.form("Pennies")) end if %>" />
											</td>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Quarters")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Quarters" value="<%if request.form("Quarters")<>"" then response.write(request.form("Quarters")) end if %>" />
											</td>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Fives")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Fives" value="<%if request.form("Fives")<>"" then response.write(request.form("Fives")) end if %>" />
											</td>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Fifties")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Fifties" value="<%if request.form("Fifties")<>"" then response.write(request.form("Fifties")) end if %>" />
											</td>
										</tr>
										<tr>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Nickels")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Nickels" value="<%if request.form("Nickels")<>"" then response.write(request.form("Nickels")) end if %>" />
											</td>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Halfdollars")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="HalfDollars" value="<%if request.form("HalfDollars")<>"" then response.write(request.form("HalfDollars")) end if %>" />
											</td>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Tens")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Tens" value="<%if request.form("Tens")<>"" then response.write(request.form("Tens")) end if %>" />
											</td>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Hundreds")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Hundreds" value="<%if request.form("Hundreds")<>"" then response.write(request.form("Hundreds")) end if %>" />
											</td>
										</tr>
										<tr>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Dimes")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Dimes" value="<%if request.form("Dimes")<>"" then response.write(request.form("Dimes")) end if %>" />
											</td>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Ones")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Ones" value="<%if request.form("Ones")<>"" then response.write(request.form("Ones")) end if %>" />
											</td>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Twenties")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Twenties" value="<%if request.form("Twenties")<>"" then response.write(request.form("Twenties")) end if %>" />
											</td>
											<td class="right">
												<%=DisplayPhrase(phraseDictionary,"Other")%>
											</td>
											<td>
												&nbsp;&nbsp;<%=currencySymbol%>
												<input type="text" size="5" name="Other" value="<%if request.form("Other")<>"" then response.write(request.form("Other")) end if %>" />
											</td>
										</tr>
									</table> <!--#table6-->
<%
	end if	'end of frmExpResults value check before /head line

	if request.form("frmShowResults")="true" then
		if request.form("frmExpResults")="true" then
			Dim stFilename
			stFilename="attachment; filename=Close Out for " & Replace(cEDate,"/","-") & ".xls"
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", stFilename
		end if

		ChangeSalesAmt=0
		DeferredSalesAmt=0
		ChangeSalesTaxAmt=0

		strSQL =	"SELECT Sum(tblSDPayments.SDPaymentAmount) AS SumOfSDPaymentAmount " &_
							"FROM tblSDPayments " &_
								"INNER JOIN Sales " &_
									"INNER JOIN [Sales Details] " &_
										"ON Sales.SaleID = [Sales Details].SaleID " &_
									"ON tblSDPayments.SDID = [Sales Details].SDID " &_
								"INNER JOIN [Payment Types] " &_
									"INNER JOIN tblPayments " &_
										"ON [Payment Types].Item# = tblPayments.PaymentMethod " &_
									"ON tblSDPayments.PaymentID = tblPayments.PaymentID " &_
							"WHERE (Sales.SaleDate<= " & DateSep & cEDate & DateSep & ") AND (Sales.Closed=0) " &_
								"AND (([Sales Details].Location)=" & cloc & ") "

		if NOT ss_IncludeTips then
			strSQL = strSQL & "AND [Sales Details].CategoryID <> 21 "
		end if
		debugSQL strSQL, "SumOfSDPaymentAmount"
		rsEntry.open strSQL, cnWS, 3
		Do while not rsEntry.eof
			if not IsNull(rsEntry("SumOfSDPaymentAmount")) then
				ChangeSalesAmt = rsEntry("SumOfSDPaymentAmount")
			else
				ChangeSalesAmt = 0
			end if
			rsEntry.movenext
		Loop
		rsEntry.close

		strSQL =	"SELECT " &_
								"Sum(tblSDPayments.ItemTax1 + tblSDPayments.ItemTax2 + tblSDPayments.ItemTax3 + tblSDPayments.ItemTax4 + tblSDPayments.ItemTax5) AS SumOfTax " &_
							"FROM Sales " &_
								"INNER JOIN [Sales Details] " &_
									"ON Sales.SaleID = [Sales Details].SaleID " &_
								"INNER JOIN tblSDPayments " &_
									"ON [Sales Details].SDID = tblSDPayments.SDID " &_
							"WHERE (Sales.SaleDate<= " & DateSep & cEDate & DateSep & ") " &_
								"AND (Sales.Closed=0) " &_
								"AND (([Sales Details].Location)=" & cloc & ") "
		if NOT ss_IncludeTips then
			strSQL = strSQL & "AND [Sales Details].CategoryID <> 21 "
		end if
		debugSQL strSQL, "SumOfTax"
		rsEntry.open strSQL, cnWS, 3
		if NOT rsEntry.eof then
			ChangeSalesTaxAmt = rsEntry("SumOfTax")
		end if
		rsEntry.close

		ChangeSalesAmt = ChangeSalesAmt - ChangeSalesTaxAmt

		TipsPaidOut = 0
		if NOT ss_IncludeTips then
			strSQL =	"SELECT Sum(tblSDPayments.SDPaymentAmount) AS TipsPaidOut " &_
								"FROM tblSDPayments " &_
									"INNER JOIN Sales " &_
										"INNER JOIN [Sales Details] " &_
											"ON Sales.SaleID = [Sales Details].SaleID " &_
										"ON tblSDPayments.SDID = [Sales Details].SDID " &_
									"INNER JOIN tblPayments " &_
										"INNER JOIN [Payment Types] " &_
											"ON tblPayments.PaymentMethod = [Payment Types].Item# " &_
										"ON tblSDPayments.PaymentID = tblPayments.PaymentID " &_
								"WHERE (Sales.SaleDate<= " & DateSep & cEDate & DateSep & ") " &_
									"AND (Sales.Closed=0) " &_
									"AND (([Sales Details].Location)=" & cloc & ") " &_
									"AND [Sales Details].CategoryID = 21 "
			debugSQL strSQL, "TipsPaidOut"
			rsEntry.open strSQL, cnWS, 3
			if NOT rsEntry.EOF then
				if not isNull(rsEntry("TipsPaidOut")) then
					TipsPaidOut = rsEntry("TipsPaidOut")
				end if
			end if
			rsEntry.close
		end if 'NOT ss_IncludeTips

		strSQL =	"SELECT tblPayments.PaymentMethod, [Payment Types].PmtTypes, " &_
								"Sum(tblSDPayments.SDPaymentAmount) AS SumOfSDPaymentAmount " &_
							"FROM tblSDPayments " &_
								"INNER JOIN Sales " &_
									"INNER JOIN [Sales Details] " &_
										"ON Sales.SaleID = [Sales Details].SaleID " &_
									"ON tblSDPayments.SDID = [Sales Details].SDID " &_
								"INNER JOIN tblPayments " &_
									"INNER JOIN [Payment Types] " &_
										"ON tblPayments.PaymentMethod = [Payment Types].Item# " &_
									"ON tblSDPayments.PaymentID = tblPayments.PaymentID " &_
							"WHERE (Sales.SaleDate<= " & DateSep & cEDate & DateSep & ") " &_
								"AND (Sales.Closed=0) " &_
								"AND (([Sales Details].Location)=" & cloc & ") "
		if NOT ss_IncludeTips then
			strSQL = strSQL & "AND [Sales Details].CategoryID <> 21 "
		end if
		strSQL = strSQL &_
							"GROUP BY tblPayments.PaymentMethod, [Payment Types].PmtTypes " &_
							"ORDER BY tblPayments.PaymentMethod"
		debugSQL strSQL, "PaymentMethod"
		rsEntry.open strSQL, cnWS, 3

		strSQL =	"SELECT [Payment Types].[Item#], " &_
								"SUM(tblSDPayments.ItemTax1 + tblSDPayments.ItemTax2 + tblSDPayments.ItemTax3 + tblSDPayments.ItemTax4 + tblSDPayments.ItemTax5) AS CategoryTaxTotal " &_
							"FROM tblSDPayments " &_
								"INNER JOIN Sales " &_
									"INNER JOIN [Sales Details] " &_
										"ON Sales.SaleID = [Sales Details].SaleID " &_
									"ON tblSDPayments.SDID = [Sales Details].SDID " &_
								"INNER JOIN [Payment Types] " &_
									"INNER JOIN tblPayments " &_
										"ON [Payment Types].Item# = tblPayments.PaymentMethod " &_
									"ON tblSDPayments.PaymentID = tblPayments.PaymentID " &_
							"WHERE (Sales.SaleDate<= " & DateSep & cEDate & DateSep & ") " &_
								"AND (Sales.Closed=0) " &_
								"AND [Sales Details].Location = " & cloc & " "
		if NOT ss_IncludeTips then
			strSQL = strSQL & "AND [Sales Details].CategoryID <> 21 "
		end if
		strSQL = strSQL &_
							"GROUP BY [Payment Types].Item# ORDER BY [Payment Types].Item# "
		debugSQL strSQL, "CategoryTaxTotal"
		rsEntry2B.open strSQL, cnWS, 3
%>
									<table id="table8" class="mainText" width="80%" cellspacing="0" style="margin: 0 auto;">
										<tr>
											<td>
												&nbsp;
											</td>
										</tr>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Salessincelastclose")%></strong>
											</td>
											<td nowrap class="right">&nbsp;
												<strong><% showCurrencyOrNumber ChangeSalesAmt %></strong>
											</td>
										</tr>
										<tr>
											<td>&nbsp;
												<strong><%=DisplayPhrase(phraseDictionary,"Salestaxsincelastclose")%></strong>
											</td>
											<td nowrap class="right">&nbsp;
												<strong><% showCurrencyOrNumber ChangeSalesTaxAmt %></strong>
											</td>
										</tr>
<%
        dim pp_paymentMethodSales : pp_paymentMethodSales = phraseDictionary("Paymentmethodsales")
        if isNull(pp_paymentMethodSales) then
            pp_paymentMethodSales = " "
        end if
		For i=0 to (UBound(SalesArr,2)-1)
			if NOT rsEntry.EOF then
				Do While NOT rsEntry.EOF
					If SalesArr(0,i)=rsEntry("PaymentMethod") then
						If Not isnull(rsEntry("SumOfSDPaymentAmount")) then
							SalesArr(2,i)=rsEntry("SumOfSDPaymentAmount")
						else
							SalesArr(2,i)=0
						end if
					end if
					rsEntry.movenext
				Loop
				rsEntry.movefirst
			end if

			if NOT rsEntry2B.EOF then
				Do While NOT rsEntry2B.EOF
					If SalesArr(0,i)=rsEntry2B("Item#") then
						If Not isnull(rsEntry2B("CategoryTaxTotal")) then
							SalesArr(3,i)=SalesArr(3,i) + rsEntry2B("CategoryTaxTotal")
						end if
					end if
					rsEntry2B.movenext
				Loop
				rsEntry2B.movefirst
			end if

			if SalesArr(2,i)<>0 then
%>
										<tr>
											<td>
												&nbsp;
											</td>
										</tr>
										<tr>
											<td class="whiteHeader" colspan="9" style="background-color: <%=session("pageColor4")%>;">
												&nbsp;<%=SalesArr(1,i)%>
											</td>
										</tr>
<%
			end if

			If SalesArr(0,i)="1" then	'"1" - Cash Section
				ChangeSalesAmtCash=SalesArr(2,i)
				' Get starting Cash Amount

				'DateAdd("d", 1, cEDate) was added for bug 15411; now that we are
				'recording actual timestamps in the CloseDate field, we need to pull
				'all records that happened before a day after the close date, to make
				'sure we get include records that happend on the close date
				strSQL =	"SELECT tblClosedData.StartingCash " &_
									"FROM tblClosedData " &_
										"INNER JOIN ( " &_
											"SELECT MAX(CloseDate) AS LastCloseDate " &_
											"FROM tblClosedData " &_
											"WHERE tblClosedData.Location = " & cLoc & " " &_
												"AND CloseDate <" & DateSep & DateAdd("d", 1, cEDate) & DateSep &_
										") LastClose " &_
											"ON LastClose.LastCloseDate = tblClosedData.CloseDate " &_
									"WHERE tblClosedData.Location = " & cLoc & " " &_
									"ORDER BY tblClosedData.CloseId DESC" 'bug 15411; this is a fix for historical data that does not record the time w/ the date
				rsEntry2.CursorLocation = 3
				rsEntry2.open strSQL, cnWS
				Set rsEntry2.ActiveConnection = Nothing
				debugSQL strSQL, "StartingCash"
				if NOT rsEntry2.EOF and NOT isNull(rsEntry2("StartingCash")) then
					StartingCash = rsEntry2("StartingCash")
				else
					StartingCash = 0
				end if

				TotalCash = ChangeSalesAmtCash+StartingCash

				dim totalCashTips : totalCashTips = 0
				dim totalNonCashTips : totalNonCashTips = 0

				if NOT ss_IncludeTips then
					strSQL =	"SELECT Sum(tblSDPayments.SDPaymentAmount) AS TotalCashTips " &_
										"FROM tblSDPayments " &_
											"INNER JOIN Sales " &_
												"INNER JOIN [Sales Details] " &_
													"ON Sales.SaleID = [Sales Details].SaleID " &_
												"ON tblSDPayments.SDID = [Sales Details].SDID " &_
											"INNER JOIN tblPayments " &_
												"INNER JOIN [Payment Types] " &_
													"ON tblPayments.PaymentMethod = [Payment Types].Item# " &_
												"ON tblSDPayments.PaymentID = tblPayments.PaymentID " &_
										"WHERE (Sales.SaleDate <= " & DateSep & cEDate & DateSep & " ) " &_
											"AND (Sales.Closed=0) " &_
											"AND ([Sales Details].Location = " & cLoc & " ) " &_
											"AND [Sales Details].CategoryID = 21  AND [Payment Types].Item# = 1"

					rsEntry3.CursorLocation = 3
					rsEntry3.open strSQL, cnWS
					set rsEntry3.ActiveConnection = nothing

					if not rsEntry3.EOF then
						if not isNull(rsEntry3("TotalCashTips")) then
							totalCashTips = rsEntry3("TotalCashTips")
						end if
					end if
					rsEntry3.close

					strSQL =	"SELECT Sum(tblSDPayments.SDPaymentAmount) AS TotalNonCashTips " &_
										"FROM tblSDPayments " &_
											"INNER JOIN Sales " &_
												"INNER JOIN [Sales Details] " &_
													"ON Sales.SaleID = [Sales Details].SaleID " &_
												"ON tblSDPayments.SDID = [Sales Details].SDID " &_
											"INNER JOIN tblPayments " &_
												"INNER JOIN [Payment Types] " &_
													"ON tblPayments.PaymentMethod = [Payment Types].Item# " &_
												"ON tblSDPayments.PaymentID = tblPayments.PaymentID " &_
										"WHERE (Sales.SaleDate <= " & DateSep & cEDate & DateSep & " ) " &_
											"AND (Sales.Closed=0) " &_
											"AND ([Sales Details].Location = " & cLoc & " ) " &_
											"AND [Sales Details].CategoryID = 21  AND [Payment Types].Item# <> 1 "

					rsEntry3.CursorLocation = 3
					rsEntry3.open strSQL, cnWS
					set rsEntry3.ActiveConnection = nothing

					if not rsEntry3.EOF then
						if not isNull(rsEntry3("TotalNonCashTips")) then
							totalNonCashTips = rsEntry3("TotalNonCashTips")
						end if
					end if
					rsEntry3.close
				end if 'NOT ss_IncludeTips

				if NOT ss_IncludeTips then
					TotalCash = TotalCash + totalCashTips - TipsPaidOut
				end if

				ActualCash = 0
               
				if request.form("Pennies")<>"" then
                    ActualCash = ActualCash + request.form("Pennies")
				end if
				if request.form("Quarters")<>"" then
					ActualCash = ActualCash + request.form("Quarters")
				end if
				if request.form("Fives")<>"" then
					ActualCash = ActualCash + request.form("Fives")
				end if
				if request.form("Fifties")<>"" then
					ActualCash = ActualCash + request.form("Fifties")
				end if
				if request.form("Nickels")<>"" then
					ActualCash = ActualCash + request.form("Nickels")
				end if
				if request.form("HalfDollars")<>"" then
					ActualCash = ActualCash + request.form("HalfDollars")
				end if
				if request.form("Tens")<>"" then
					ActualCash = ActualCash + request.form("Tens")
				end if
				if request.form("Hundreds")<>"" then
					ActualCash = ActualCash + request.form("Hundreds")
				end if
				if request.form("Dimes")<>"" then
					ActualCash = ActualCash + request.form("Dimes")
				end if
				if request.form("Ones")<>"" then
					ActualCash = ActualCash + request.form("Ones")
				end if
				if request.form("Twenties")<>"" then
					ActualCash = ActualCash + request.form("Twenties")
				end if
				if request.form("Other")<>"" then
					ActualCash = ActualCash + request.form("Other")
				end if

				if ActualCash = 0 then
					ActualCash = TotalCash
				end if

				amtCashOverShort = ActualCash-TotalCash

				if request.form("InDrawer")<>"" then
					dim getInDrawerAmount
					if isnumeric(request.form("InDrawer")) then
						getInDrawerAmount = request.form("InDrawer")
					else
						getInDrawerAmount = request.form("hidInDrawer")
					end if

					Call SetLocale(session("mvarLocaleStr"))
						NewStartingCash = cdbl(getInDrawerAmount)
					Call SetLocale("en-us")
				else
					NewStartingCash = 0
				end if
%>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Startingcashindrawer")%></strong>
											</td>
											<td nowrap class="right">
												<strong>
													<input class="mainText" style="border-style: none; text-align: right; font-weight: bold" readonly="true" value="<% showCurrencyOrNumber StartingCash %>" type="text" name="StartingCash" />
												</strong>
											</td>
										</tr>
<%
				rsEntry2.close
%>
										<tr>
                                            <%pp_paymentMethodSales = Replace(phraseDictionary.Item("Paymentmethodsales"), "<PAYMENTMETHOD>", SalesArr(1,i))%>
											<td>
												&nbsp;<strong><%=DisplaySinglePhrase(pp_paymentMethodSales,"Paymentmethodsales")%>:</strong>
											</td>
											<td nowrap class="right">
												<strong>
													<input class="mainText" style="border-style: none; text-align: right; font-weight: bold" readonly="true" value="<% showCurrencyOrNumber ChangeSalesAmtCash %>" type="text" name="ChangeSalesAmtCash" />
												</strong>
											</td>
										</tr>
<%
				if NOT ss_IncludeTips then
%>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Cashtips")%></strong>
											</td>
											<td nowrap class="right">
												<strong><% showCurrencyOrNumber totalCashTips %></strong>
											</td>
										</tr>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Tipspaidout")%></strong>
											</td>
											<td nowrap class="right">
												<strong><% showCurrencyOrNumber TipsPaidOut %></strong>
											</td>
										</tr>
										<tr style="font-style: italic;">
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Othertips")%></strong>
											</td>
											<td nowrap class="right">
												<strong>
													<% showCurrencyOrNumber totalNonCashTips %></strong>
											</td>
										</tr>
<%
				end if
%>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Totalcashexpectedindrawer")%></strong>
											</td>
											<td nowrap class="right">
												<strong>
													<%showCurrencyOrNumber TotalCash %>
												</strong>
											</td>
										</tr>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Actualcashcountedindrawer")%></strong>
											</td>
											<td nowrap class="right">
												<strong>
													<% if ActualCash = 0 then showCurrencyOrNumber TotalCash else showCurrencyOrNumber ActualCash end if %>
												</strong>
											</td>
										</tr>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Overshortamount")%></strong>
											</td>
											<td nowrap class="right">
												<strong>
													<span id="spanCashOverShort">
														<% if amtCashOverShort < 0 then %><span style="color: red;"><% end if %>
															<% showCurrencyOrNumber amtCashOverShort %>
														<% if amtCashOverShort < 0 then %></span><% end if %>
													</span>
												</strong>
											</td>
										</tr>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Amounttokeepindrawer")%></strong>
											</td>
											<td nowrap class="right">
												<input type="hidden" name="hidInDrawer" value="<%= NewStartingCash %>" />
												<input class="mainText right" style="font-weight: bold" type="text" name="InDrawer" size="16" maxlength="30" value="<% showCurrencyOrNumber NewStartingCash %>" onchange="showReport();" onclick="javascript:this.focus();this.select();" />
											</td>
										</tr>
<%
				'END "1" Cash Section

			elseIf SalesArr(0,i)="2" then
				ChangeSalesAmtCheck=SalesArr(2,i)

				if request.form("txtActualCheck")<>"" then

					dim getTxtActualCheck
					if isnumeric(request.form("txtActualCheck")) then
						getTxtActualCheck = CSTR(request.form("txtActualCheck"))
					else 
						getTxtActualCheck = CSTR(request.form("hidtxtActualCheck"))
					end if

					Call SetLocale(session("mvarLocaleStr"))
						ActualCheck = cdbl(getTxtActualCheck)
					Call SetLocale("en-us")
				else
					ActualCheck = ChangeSalesAmtCheck
				end if

				amtCheckOverShort = ActualCheck - ChangeSalesAmtCheck
%>
										<tr>
                                            <%pp_paymentMethodSales = Replace(phraseDictionary.Item("Paymentmethodsales"), "<PAYMENTMETHOD>", SalesArr(1,i))%>
											<td>
												&nbsp;<strong><%=DisplaySinglePhrase(pp_paymentMethodSales,"Paymentmethodsales")%>:</strong>
											</td>
											<td nowrap class="right">
												&nbsp;<strong>&nbsp;<% showCurrencyOrNumber ChangeSalesAmtCheck %></strong>
												<input type="hidden" name="ChangeSalesAmtCheck" value="<%=ChangeSalesAmtCheck%>" />
											</td>
										</tr>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Actualcheckamount")%></strong>
											</td>
											<td nowrap class="right">
												<input type="hidden" name="hidtxtActualCheck" value="<% if ActualCheck = 0 then response.write ChangeSalesAmtCheck else response.write ActualCheck end if %>" />
												<input class="mainText" style="text-align: right; font-weight: bold" type="text" name="txtActualCheck" size="16" maxlength="30" align="right" value="<% if ActualCheck = 0 then showCurrencyOrNumber ChangeSalesAmtCheck else showCurrencyOrNumber ActualCheck end if %>" onchange="showReport();" onclick="javascript:this.focus();this.select();" />
											</td>
										</tr>
										<tr>
											<td>
												&nbsp;<strong><%=DisplayPhrase(phraseDictionary,"Overshortamount")%></strong>
											</td>
											<td nowrap class="right">
												<strong>
													<span id="spanCheckOverShort">
														<% if amtCheckOverShort < 0 then %><span style="color: red;"><% end if %>
															<% showCurrencyOrNumber amtCheckOverShort %>
														<% if amtCheckOverShort < 0 then %></span><% end if %>
													</span>
												</strong>
											</td>
										</tr>
<%
			elseif SalesArr(2,i)<>0 then
%>
										<tr>
                                            <%pp_paymentMethodSales = Replace(phraseDictionary.Item("Paymentmethodsales"), "<PAYMENTMETHOD>", SalesArr(1,i))%>
											<td>
												&nbsp;<strong><%=DisplaySinglePhrase(pp_paymentMethodSales,"Paymentmethodsales")%>:</strong>
											</td>
											<td nowrap class="right">
												&nbsp;<strong>&nbsp;<% showCurrencyOrNumber SalesArr(2,i) %></strong>
											</td>
										</tr>
<%
			end if
		Next 'For i=0 to (UBound(SalesArr,2)-1)
		rsEntry.close

		if request.form("runCloseOut")="true" then

			''Do CloseOut
			If isnull(ChangeSalesAmt) or ChangeSalesAmt=Null then
				ChangeSalesAmt=0
			End if
			If isnull(NewStartingCash) or NewStartingCash=Null then
				NewStartingCash=0
			End if

			If isnull(amtCashOverShort) or amtCashOverShort=Null then
				amtCashOverShort=0
			End if
			If isnull(amtCheckOverShort) or amtCheckOverShort=Null then
				amtCheckOverShort=0
			End if

			strSQL =	"INSERT INTO tblClosedData (" &_
									"CloseDate, " &_
									"Location, " &_
									"Sales, " &_
									"StartingCash, " &_
									"ShortCash, " &_
									"ShortCheck, " &_
									"Base, " &_
									"Exported, " &_
									"Notes, " &_
									"ClosedBy, " &_
									"ActualCash, " &_
									"ActualCheck, " &_
									"OldCloseDate, " &_
									"SalesTax" &_
								") VALUES (" &_
									DateSep & dateForDB & DateSep & ", " &_
									cLoc & ", " &_
									ChangeSalesAmt & ", " &_
									NewStartingCash & ", " &_
									amtCashOverShort & ", " &_
									amtCheckOverShort & ", " &_
									"0, " &_
									"0, " &_
									"'Closeout from WS', "
			if session("Admin")="sa" OR session("Admin")="owner" then
				strSQL = strSQL & "0, "
			else
				strSQL = strSQL & session("empID") & ", "
			end if
			strSQL = strSQL &_
									FormatNumber(ActualCash-StartingCash, 2, -1, 0, 0) & ", " &_
									FormatNumber(ActualCheck, 2, -1, 0, 0) & ", " &_
									DateSep & oldCloseDate & DateSep & ", " &_
									ChangeSalesTaxAmt &_
								")"
			debugSQL strSQL, "SQL"
			'response.end
			cnWS.execute strSQL	'write to sql db

			strSQL =	"SELECT CloseID " &_
								"FROM tblClosedData " &_
								"WHERE Location = " & cloc & " " &_
									"AND CloseDate = " & DateSep & dateForDB & DateSep
			rsEntry.open strSQL, cnWS, 3
			if Not rsEntry.eof then
				Do while not rsEntry.eof
					intCLoseID = rsEntry("CloseID")
					rsEntry.movenext
				Loop
			end if
			rsEntry.close

			if ss_Fitlink then
				strSQL =	"INSERT INTO tblClosedData (" &_
										"CloseID, " &_
										"CloseDate, " &_
										"Location, " &_
										"Sales, " &_
										"StartingCash, " &_
										"ShortCash, " &_
										"ShortCheck, " &_
										"Base, " &_
										"Exported, " &_
										"Notes, " &_
										"ClosedBy, " &_
										"SalesTax" &_
									") VALUES (" &_
										intCloseID & ", " &_
										DateSepFit & dateForDB & DateSepFit & ", " &_
										cLoc & ", " &_
										ChangeSalesAmt & ", " &_
										NewStartingCash & ", " &_
										amtCashOverShort & ", " &_
										amtCheckOverShort & ", " &_
										"0, " &_
										"0, " &_
										"'Closeout from WS', "
				if session("Admin")="sa" OR session("Admin")="owner" then
					strSQL = strSQL & "0" & ", "
				else
					strSQL = strSQL & session("empID") & ", "
				end if
				strSQL = strSQL &_
										ChangeSalesTaxAmt &_
									")"
				debugSQL strSQL, "SQL"
				cnWSFit.execute strsql	'write to fitlink closeddata.mdb
			end if 'ss_Fitlink

			For i=0 to (UBound(SalesArr,2)-1)
				If isnull(SalesArr(2,i)) or SalesArr(2,i)="" then
					SalesArr(2,i)=0
				end if
				strSQL = "INSERT INTO tblClosedSalesPmtType (CloseID, PmtTypeID, Amt, Tax) VALUES "
				strSQL = strSQL & "(" & intCloseID & ", " & SalesArr(0,i) & ", " & (SalesArr(2,i)-SalesArr(3,i)) & ", "
				if isnumeric(SalesArr(3,i)) and SalesArr(3,i)<>"" then
					strSQL = strSQL & SalesArr(3,i)
				else
					strSQL = strSQL & "0"
				end if
				strSQL = strSQL & ") "
				debugSQL strSQL, "SQL"
				'response.write "<br />: " & SalesArr(3,i) & " :<br />"
				cnWS.execute strSQL	'write to sql db
				if ss_Fitlink then
					strSQL = "INSERT INTO tblClosedSalesPmtType (CloseID, PmtTypeID, Amt) VALUES "
					strSQL = strSQL & "(" & intCloseID & ", " & SalesArr(0,i) & ", " & (SalesArr(2,i)-SalesArr(3,i)) & ")"
					cnWSFit.execute strsql	'write to fitlink closeddata.mdb
				end if
			Next

			Dim rsSalesbyCat, rsCatCount, rsCategory, counter, numCat, catArray() 
			set rsSalesbyCat = Server.CreateObject("ADODB.Recordset")
			set rsCatCount = Server.CreateObject("ADODB.Recordset")
			set rsCategory = Server.CreateObject("ADODB.Recordset")
			strSQL=	"SELECT tblPayments.PaymentMethod, " &_
								"Sum(tblSDPayments.SDPaymentAmount) AS SumOfSDPaymentAmount, " &_
								"[Sales Details].CategoryID, " &_
								"SUM(tblSDPayments.ItemTax1 + tblSDPayments.ItemTax2 + tblSDPayments.ItemTax3 + tblSDPayments.ItemTax4 + tblSDPayments.ItemTax5) AS CategoryTaxTotal " &_
							"FROM tblSDPayments " &_
								"INNER JOIN Sales " &_
									"INNER JOIN [Sales Details] " &_
										"ON Sales.SaleID = [Sales Details].SaleID " &_
									"ON tblSDPayments.SDID = [Sales Details].SDID " &_
								"INNER JOIN tblPayments " &_
									"ON tblSDPayments.PaymentID = tblPayments.PaymentID " &_
							"WHERE ([Sales Details].Location=" & cloc & ") " &_
								"AND (Sales.SaleDate<= " & DateSep & cEDate & DateSep & ") " &_
								"AND (Sales.Closed=0) " &_
								"AND [Sales Details].CategoryID <> 21 " &_
							"GROUP BY tblPayments.PaymentMethod, [Sales Details].CategoryID"
			debugSQL strSQL, "SQL"
			rsSalesbyCat.open strSQL, cnWS, 3

			strSQL= "SELECT Count(CategoryID) as cntCatID FROM Categories"
			rsCatCount.open strSQL, cnWS, 3
			numCat = rsCatCount("cntCatID")
			ReDim catArray(3, numCat)

			strSQL= "SELECT CategoryID FROM Categories"
			rsCategory.open strSQL, cnWS, 3
			counter=0
			If not rsCategory.eof then
				Do While not rsCategory.eof
					catArray(0, counter) = rsCategory("CategoryID")
					counter = counter + 1
					rsCategory.MoveNext
				Loop
			end if
			rsCategory.close

			catArray(1, counter)=0
			If not rsSalesbyCat.eof then
				Do while not rsSalesbyCat.eof
					'If rsSalesbyCat("CashEQ")=True then
						counter = 0
						Do While (counter < numCat)
							If rsSalesbyCat("CategoryID") = catArray(0, counter) Then
								If not isnull(rsSalesbyCat("SumOfSDPaymentAmount")) then
									catArray(1, counter) = catArray(1, counter) + rsSalesbyCat("SumOfSDPaymentAmount")
								end if
								if NOT isNull(rsSalesbyCat("CategoryTaxTotal")) then
									catArray(2, counter) = catArray(2, counter) + rsSalesbyCat("CategoryTaxTotal")
								end if
								Exit Do
							Else
								counter = counter + 1
							End If
						Loop
					'end if

					rsSalesbyCat.movenext
				loop
			end if 'not rsSalesbyCat.eof
			rsSalesbyCat.close

			For i=0 to (UBound(catArray,2)-1)
				If catArray(1,i)<>0 and not isnull(catArray(1,i)) and catArray(1,i)<>"" then
					strSQL = "INSERT INTO tblClosedSalesCategory (CloseID, CategoryID, Amt, Tax) VALUES "
					strSQL = strSQL & "(" & intCloseID & ", " & catArray(0,i) & ", " & (catArray(1,i)-catArray(2,i)) & ", "
					if isnumeric(catArray(2,i)) then
						strSQL = strSQL & catArray(2,i)
					else
						strSQL = strSQL & "0"
					end if
					strSQL = strSQL & ") "
					cnWS.execute strSQL	'write to sql db
					if ss_Fitlink then
						strSQL = "INSERT INTO tblClosedSalesCategory (CloseID, CategoryID, Amt) VALUES "
						strSQL = strSQL & "(" & intCloseID & ", " & catArray(0,i) & ", " & (catArray(1,i)-catArray(2,i)) & ")"
						cnWSFit.execute strsql	'write to fitlink closeddata.mdb
					end if
				end if
			Next

			strSQL =	"UPDATE sales SET " &_
									"sales.Closed = 1, CloseID = " & intCloseID & " " &_
								"FROM sales " &_
									"INNER JOIN [Sales Details] " &_
										"ON sales.SaleID = [Sales Details].SaleID " &_
								"WHERE (sales.SaleDate<=" & DateSep & cEDate & DateSep & ") " &_
									"AND ([Sales Details].Location=" & cloc & ") " &_
									"AND Sales.Closed = 0"
			cnWS.execute strSQL

			strSQL =	"UPDATE [Sales Details] SET " &_
									"[Sales Details].Closed = 1, CloseID = " & intCloseID & " " &_
								"FROM sales " &_
									"INNER JOIN [Sales Details] " &_
										"ON sales.SaleID = [Sales Details].SaleID " &_
								"WHERE (sales.SaleDate<=" & DateSep & cEDate & DateSep & ") " &_
									"AND ([Sales Details].Location=" & cloc & ") " &_
									"AND [Sales Details].Closed = 0 "
			cnWS.execute strSQL

			strSQL =	"UPDATE [Payment Data] SET " &_
									"[Payment Data].Closed = 1, CloseID = " & intCloseID & " " &_
								"WHERE ([Payment Data].PaymentDate<=" & DateSep & cEDate & DateSep & ") " &_
									"AND ([Payment Data].Location=" & cloc & ") " &_
									"AND ([Payment Data].[Type]<>9) AND [Payment Data].Closed = 0 "
			cnWS.execute strSQL
%>
										<script type="text/javascript">
											alert("<%=DisplayPhraseJS(phraseDictionary,"Closeoutsuccessful")%>");
										</script>
<%
		else 'request.form("runCloseOut")<>"true"
%>
										<tr>
											<td colspan="2" class="center-ch" valign="middle" height="30">
												<input onclick="closeOut();" type="button" name="CloseOutBtn" value="<%=DisplayPhraseAttr(phraseDictionary,"Closeout")%>" />
											</td>
										</tr>
<%
		end if 'request.form("runCloseOut")="true"
	end if 'show report
%>
									</table> <!--#table8-->
								</td>
							</tr>
						</table> <!--#table5-->
					</td>
				</tr>
			</table> <!--#table2-->
			</form>
		</td>
	</tr>
</table> <!--#table1-->
<% pageEnd %>
<%
	if NOT request.form("frmExpResults")="true" then
%>
<!-- #include file="post.asp" -->
<%
	end if

end if
%>
