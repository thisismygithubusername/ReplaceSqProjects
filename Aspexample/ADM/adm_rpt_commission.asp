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

	Dim rsEntry
	Set rsEntry = Server.CreateObject("ADODB.Recordset")
	
	dim trainerID, clockStatus, tmpDate, tmpTrn, cEDate, cSDate, totalTime, tmpRate, ap_ipay_trn, ap_view_all_locs, totalSales
	 
	ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
	tmpDate = ""
	tmpTrn = ""
	totalTime = 0
	tmpRate = 0.0

	ap_ipay = validAccessPriv("RPT_IPAY")
	ap_ipay_trn = validAccessPriv("RPT_IPAY_TRN")

	if not Session("Pass") OR Session("Admin")="false" OR (NOT ap_ipay AND NOT ap_ipay_trn) then 
		%>
		<script type="text/javascript">
			alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
			javascript:history.go(-1);
		</script>
		<%
	else

	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if
		
		' Load up das H0tWurdz
		Dim hw_Date			:	hw_Date			= xssStr(allHotWords(57))
		Dim hw_ItemName		:	hw_ItemName		= xssStr(allHotWords(370))
		Dim hw_ID			:	hw_ID			= xssStr(allHotWords(134))
		Dim hw_ItemPrice	:	hw_ItemPrice	= xssStr(allHotWords(363))
		Dim hw_ExcludingTax	:	hw_ExcludingTax	= Replace(xssStr(allHotWords(341)), " ", "&nbsp;")
		Dim hw_Client		:	hw_Client		= xssStr(allHotWords(12))
		Dim hw_All			:	hw_All			= xssStr(allHotWords(149))
		Dim hw_View			:	hw_View			= xssStr(allHotWords(159))
		Dim hw_Location		:	hw_Location		= xssStr(allHotWords(8))
		Dim hw_Total		:	hw_Total		= xssStr(allHotWords(22))
		Dim hw_Commission	:	hw_Commission	= xssStr(allHotWords(208))
		Dim hw_Commission2	:	hw_Commission2	= xssStr(allHotWords(209))
		Dim hw_StartDate	:	hw_StartDate	= xssStr(allHotWords(77))
		Dim hw_EndDate		:	hw_EndDate		= xssStr(allHotWords(79))
		
		dim category : category = ""
		if (RQ("category"))<>"" then
			category = RQ("category")
		elseif (RF("category"))<>"" then
			category = RF("category") 
		end if

		if isNum(request.QueryString("trainerID")) then
			trainerID = request.QueryString("trainerID")
		elseif isNum(request.form("trainerID")) then
			trainerID = request.form("trainerID")
		else
			trainerID = 0
		end if
		
		if request.form("requiredtxtDateStart")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cSDate = CDATE(request.form("requiredtxtDateStart"))
			Call SetLocale("en-us")
		else
			cSDate = DateAdd("y",-14,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
		end if
	
		if request.form("requiredtxtDateEnd")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cEDate = CDATE(request.form("requiredtxtDateEnd"))
			Call SetLocale("en-us")
		else
			cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		end if
		
		if isNum(request.form("optLocation")) then
			cLoc = CSTR(request.form("optLocation"))
		else
			if session("numLocations")>1 then
				if session("UserLoc") <> 0 then
					cLoc = CSTR(session("UserLoc"))
				else
					cLoc = CSTR(session("curLocation"))
				end if
			else
				cLoc = "0"
			end if
		end if
	
		%>
	
<META http-equiv="Content-Type" content="text/html; charset=UTF-8">
	
	<% if NOT request.form("frmExpReport")="true" then %>	
	
		<!-- #include file="pre.asp" -->
		<!-- #include file="frame_bottom.asp" -->
			
		<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_commission", "reportFavorites", "plugins/jquery.SimpleLightBox" )) %>	
		<%= css(array("SimpleLightBox")) %>
		<script type="text/javascript">
			function exportReport() {
				document.frmTimeClock.frmExpReport.value = "true";
				document.frmTimeClock.frmGenReport.value = "true";
				<% iframeSubmit "frmTimeClock", "adm_rpt_commission.asp" %>
			}
		</script>		
		<style type="text/css">
			.ColumnHeader
			{
				font-weight: bold !important;
				text-transform: capitalize !important;
			}
		 </style>
		<!-- #include file="../inc_date_ctrl.asp" -->
	
	<% end if 'frmExpReport="true" before <html>%>


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
	<%=DisplayPhrase(reportPageTitlesDictionary,"Commission") %>
	<div id="add-to-favorites">
		<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
	</div>
	</div>
<%end if %>

	<table height="100%" width="<%=strPageWidth%>" cellspacing="0">
		<tr> 
		 <td valign="top" height="100%" width="100%"> <br />
			<table class="center" cellspacing="0" width="90%">
			<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<tr> 
				<td class="headText" align="left" valign="top"><b><%= pp_PageTitle("Commission") %></b></td>
			  </tr>
			<%end if %>
			  <tr>
				<td valign="top">
					<table id="commisionReport" class="center" width="85%" cellspacing="0">
						<form name="frmTimeClock" action="adm_rpt_commission.asp" method="post">
							<input type="hidden" name="frmGenReport" value="<%=xssStr(request.form("frmGenReport"))%>">
							<input type="hidden" name="frmExpReport" value="">
							<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
								<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
								<input type="hidden" name="category" value="<%=category%>">
							<% end if %>
						<tr>
							<td colspan="14" class="center-ch headText">
								<table class="mainText border4 center-block" cellspacing="0">
									<tr>
										<td class="center-ch" valign="bottom" style="background-color:#F2F2F2;">
										<b>
							&nbsp;<%= hw_View %>:&nbsp;<select name="trainerID" onChange="document.frmTimeClock.submit();">

		<%	if ap_ipay then  %>
								<option value="0">All Staff Members - Summary</option>
								<option value="-1" <% if trainerID = "-1" then response.write "selected" end if %>>All Staff Members - Detail</option>
		<% 		strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName FROM TRAINERS INNER JOIN [Sales Details] ON ([Sales Details].commissionTrnID = TRAINERS.TrainerID OR  [Sales Details].commissionTrn2ID = TRAINERS.TrainerID)"
		 		strSQL = strSQL & " INNER JOIN SALES ON ([Sales Details].SaleID = SALES.SaleID "
				strSQL = strSQL & " AND SALES.SaleDate >= " & DateSep & cSDate & DateSep & " AND SALES.SaleDate <= " & DateSep & cEDate & DateSep & ")"
				strSQL = strSQL & " WHERE TRAINERS.[Delete]=0 " 
				if request.form("optInactive")="" then
					strSQL = strSQL & " AND TRAINERS.[Active]=1 " 
				end if
				
				strSQL = strSQL & " AND TRAINERS.TrainerID>0 AND TRAINERS.isSystem=0 "
				strSQL = strSQL & " ORDER BY TRAINERS.TrLastName;"
			else
				strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName, TRAINERS.DisplayName FROM TRAINERS "
				strSQL = strSQL & " WHERE TRAINERS.[Delete]=0 " 
				if request.form("optInactive")="" then
					strSQL = strSQL & " AND TRAINERS.[Active]=1 " 
				end if
				strSQL = strSQL & " AND TRAINERS.TrainerID = " & session("EmpID")
			end if			
		response.write debugSQL(strSQL, "SQL")
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			
			if NOT rsEntry.EOF then
				Do While NOT rsEntry.EOF
			%>
									<option value="<%=rsEntry("TrainerID")%>" <%if trainerID=CSTR(rsEntry("TrainerID")) then response.write "selected" end if%>><%=FmtTrnNameNew(rsEntry, true)%></option>
			<%
					rsEntry.MoveNext
				Loop
			else
			end if

			rsEntry.close
		
		%>
							  </select>

&nbsp;Include Inactive Staff:<input onClick="document.frmTimeClock.submit();" type="checkbox" name="optInactive" <% if request.form("optInactive")="on" then %>checked<% end if %>>&nbsp;
&nbsp;Show Split&nbsp;<%= hw_Commission%>&nbsp;Detail:<input onClick="document.frmTimeClock.submit();" type="checkbox" name="showCommissionSplitDetail" <% if request.form("showCommissionSplitDetail")="on" then %>checked<% end if %>>&nbsp;

			<select name="optProdServ">
				<option value="" <%if request.form("optProdServ")="" then response.write "selected" end if%>>All Products & Services</option>
				<option value="1" <%if request.form("optProdServ")="1" then response.write "selected" end if%>>Products Only</option>
				<option value="2" <%if request.form("optProdServ")="2" then response.write "selected" end if%>>Services Only</option>
			</select>				
							  
					<br />&nbsp;<%=hw_Location %>:&nbsp;<select name="optLocation" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
							<option value="0" ><%=hw_All%>&nbsp;<%= hw_Location %>s</option>
	<%
		strSQL = "SELECT LocationID, LocationName FROM Location WHERE Active=1 AND LocationID<>98 ORDER BY LocationName "
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
	
		do while not rsEntry.EOF
	%>
	
								<option value="<%=rsEntry("LocationID")%>" <%if cstr(rsEntry("LocationID"))=cLoc then response.write "selected" end if%>><%=rsEntry("LocationName")%></option>
	<%
			rsEntry.MoveNext
		loop
		rsEntry.close						
	%>
					  </select>
	
				   &nbsp;<span class="ColumnHeader"><%= hw_StartDate %>:</span>
						  <input type="text"  onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
			<script type="text/javascript">
				var cal1 = new tcal({'formname':'frmTimeClock', 'controlname':'requiredtxtDateStart'});
				cal1.a_tpl.yearscroll = true;
			</script>
					   &nbsp;<span class="ColumnHeader"><%= hw_EndDate %>:</span>
						  <input type="text"  onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
			<script type="text/javascript">
				var cal2 = new tcal({'formname':'frmTimeClock', 'controlname':'requiredtxtDateEnd'});
				cal2.a_tpl.yearscroll = true;
			</script>
			 &nbsp;
			 <%'MB 59_3810 %>
			<% if request.form("trainerID")="0" OR request.form("trainerID")="" then %>		
		        Show $0 Commission:<input type="checkbox" name="optZeroCommission" <% if request.form("optZeroCommission")="on" then %>checked<% end if %>>	&nbsp; 			   <%end if %>
			  <br />
						<% showDateArrows("frmTimeClock") %>
						  <input type="button" name="Button" value="Generate" onClick="showReport();">
						  <% exportToExcelButton %>
						<% savingButtons "frmTimeClock", "Commissions" %>
	
										</b></td>
									</tr>
								</table>
							</td>
						</tr>
						<tr>
							<td colspan="14">&nbsp;</td>
						</tr>
						</form>
					</table>
<% end if			'end of frmExpreport value check before /head line
							if request.form("frmGenReport")="true" then
								if request.form("frmExpReport")="true" then
									Dim stFilename
									if request.form("TrainerID")=0 then 
										stFilename="attachment; filename=Commission-Summary-All-Staff " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									elseif request.form("TrainerID")=-1 then
										stFilename="attachment; filename=Commission-Detail-All-Staff " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									else
										stFilename="attachment; filename=Commission-Detail " & Replace(cSDate,"/","-") & " to " & Replace(cEDate,"/","-") & ".xls" 
									end if
									Response.ContentType = "application/vnd.ms-excel" 
									Response.AddHeader "Content-Disposition", stFilename 
								end if 
							
								dim strSQLCommission, rsCommission
								strSQLCommission = ""
								set rsCommission=Server.createobject("ADODB.Recordset")
								
								if request.form("trainerID")="0" then ' Summary View
								
									
									strSQLCommission = " SELECT TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, ISNULL(TR2.CommissionTotal2, 0) as Commission1Total, ISNULL(TR3.CommissionTotal3, 0) as Commission2Total, SUM(ISNULL(TR2.CommissionTotal2, 0) + ISNULL(TR3.CommissionTotal3, 0)) as CommissionTotal FROM TRAINERS LEFT OUTER JOIN ( "
									strSQLCommission = strSQLCommission & " SELECT TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, "
									strSQLCommission = strSQLCommission & " SUM( (CASE WHEN NOT PRODUCTS.StdCommissionPercRate IS NULL THEN ROUND((PRODUCTS.StdCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.StdCommissionFlatRate IS NULL THEN ROUND(PRODUCTS.StdCommissionFlatRate * [Sales Details].Quantity, 2) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.PromoCommissionPercRate IS NULL THEN ROUND((PRODUCTS.PromoCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.PromoCommissionFlatRate IS NULL THEN ROUND(PRODUCTS.PromoCommissionFlatRate * [Sales Details].Quantity, 2 ) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.StdTrnCommissionPercRate IS NULL THEN ROUND((TRAINERS.StdTrnCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.StdTrnCommissionFlatRate IS NULL THEN ROUND(TRAINERS.StdTrnCommissionFlatRate * [Sales Details].Quantity, 2 ) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.PromoTrnCommissionPercRate IS NULL THEN ROUND((TRAINERS.PromoTrnCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END + "
									' BQL Bug 4682 - removed condition from query in the division CASE statement - we already calculate both CommissTrn1 and 2 seperately, so we need to divide by 2 even if it's the same trainer, since we'll pick up the other half in the other query
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.PromoTrnCommissionFlatRate IS NULL THEN ROUND(TRAINERS.PromoTrnCommissionFlatRate * [Sales Details].Quantity, 2 ) ELSE 0 END) / CASE WHEN [Sales Details].commissionTrn2ID IS NULL THEN 1 ELSE 2 END) as CommissionTotal2 "
									strSQLCommission = strSQLCommission & " FROM [Sales Details] INNER JOIN TRAINERS ON [Sales Details].commissionTrnID = TRAINERS.TrainerID INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID INNER JOIN SALES on [Sales Details].SaleID = SALES.SaleID "
									strSQLCommission = strSQLCommission & " WHERE TRAINERS.TrainerID > 0  AND TRAINERS.isSystem=0 AND TRAINERS.[Delete]=0 AND TRAINERS.[Active]=1 "
									strSQLCommission = strSQLCommission & "	AND ( NOT (PRODUCTS.StdCommissionPercRate IS NULL AND PRODUCTS.StdCommissionFlatRate IS NULL AND PRODUCTS.PromoCommissionPercRate IS NULL AND PRODUCTS.PromoCommissionFlatRate IS NULL)) "

									if request.form("optProdServ")="1" then	'products only
										strSQLCommission = strSQLCommission & " AND [Sales Details].CategoryID > 25 "
									elseif request.form("optProdServ")="2" then
										strSQLCommission = strSQLCommission & " AND [Sales Details].CategoryID < 21 "
									end if

									strSQLCommission = strSQLCommission & " AND Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cEDate & DateSep
									'strSQLCommission = strSQLCommission & " AND [Sales Details].Returned = 0 "
									'if request.form("optLocation")<>"0" then
									if cLoc<>"0" then
										strSQLCommission = strSQLCommission & " AND [Sales Details].Location = " & cLoc
									end if
									strSQLCommission = strSQLCommission & " GROUP BY TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, "
									strSQLCommission = strSQLCommission & " TRAINERS.StdTrnCommissionPercRate, TRAINERS.StdTrnCommissionFlatRate, TRAINERS.PromoTrnCommissionPercRate, TRAINERS.PromoTrnCommissionFlatRate " 
									strSQLCommission = strSQLCommission & ") TR2 ON TRAINERS.TrainerID = TR2.TrainerID LEFT OUTER JOIN ("
									strSQLCommission = strSQLCommission & " SELECT TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, "
									strSQLCommission = strSQLCommission & " SUM((CASE WHEN NOT PRODUCTS.StdCommissionPercRate IS NULL THEN ROUND((PRODUCTS.StdCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.StdCommissionFlatRate IS NULL THEN ROUND(PRODUCTS.StdCommissionFlatRate * [Sales Details].Quantity, 2 ) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.PromoCommissionPercRate IS NULL THEN ROUND((PRODUCTS.PromoCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.PromoCommissionFlatRate IS NULL THEN ROUND(PRODUCTS.PromoCommissionFlatRate * [Sales Details].Quantity, 2 ) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.StdTrnCommissionPercRate IS NULL THEN ROUND((TRAINERS.StdTrnCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.StdTrnCommissionFlatRate IS NULL THEN ROUND(TRAINERS.StdTrnCommissionFlatRate * [Sales Details].Quantity, 2 ) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.PromoTrnCommissionPercRate IS NULL THEN ROUND((TRAINERS.PromoTrnCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END + "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.PromoTrnCommissionFlatRate IS NULL THEN ROUND(TRAINERS.PromoTrnCommissionFlatRate * [Sales Details].Quantity, 2 ) ELSE 0 END) / CASE WHEN [Sales Details].commissionTrnID IS NULL THEN 1 ELSE 2 END) as CommissionTotal3 "
									strSQLCommission = strSQLCommission & " FROM [Sales Details] INNER JOIN TRAINERS ON [Sales Details].commissionTrn2ID = TRAINERS.TrainerID INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID INNER JOIN SALES on [Sales Details].SaleID = SALES.SaleID "
									strSQLCommission = strSQLCommission & " WHERE TRAINERS.TrainerID > 0 AND TRAINERS.isSystem=0 AND TRAINERS.[Delete]=0 AND TRAINERS.[Active]=1 "
									strSQLCommission = strSQLCommission & "	AND ( NOT (PRODUCTS.StdCommissionPercRate IS NULL AND PRODUCTS.StdCommissionFlatRate IS NULL AND PRODUCTS.PromoCommissionPercRate IS NULL AND PRODUCTS.PromoCommissionFlatRate IS NULL)) "

									if request.form("optProdServ")="1" then	'products only
										strSQLCommission = strSQLCommission & " AND [Sales Details].CategoryID > 25 "
									elseif request.form("optProdServ")="2" then
										strSQLCommission = strSQLCommission & " AND [Sales Details].CategoryID < 21 "
									end if

									strSQLCommission = strSQLCommission & " AND Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cEDate & DateSep
									'strSQLCommission = strSQLCommission & " AND [Sales Details].Returned = 0 "
									'if request.form("optLocation")<>"0" then
									if cLoc<>"0" then
										strSQLCommission = strSQLCommission & " AND [Sales Details].Location = " & cLoc
									end if
									strSQLCommission = strSQLCommission & " GROUP BY TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, "
									strSQLCommission = strSQLCommission & " TRAINERS.StdTrnCommissionPercRate, TRAINERS.StdTrnCommissionFlatRate, TRAINERS.PromoTrnCommissionPercRate, TRAINERS.PromoTrnCommissionFlatRate " 
									strSQLCommission = strSQLCommission & ") TR3 "
									strSQLCommission = strSQLCommission & " ON TRAINERS.TrainerID = TR3.TrainerID "	
									'MB 59_3810
									if request.form("optZeroCommission")<>"on" then								
									    strSQLCommission = strSQLCommission & "	WHERE ISNULL(TR2.CommissionTotal2, 0) + ISNULL(TR3.CommissionTotal3, 0) <> 0" 
									else
									     strSQLCommission = strSQLCommission & " WHERE (TR2.CommissionTotal2 is NOT Null OR TR3.CommissionTotal3 is NOT Null) "
									end if
									strSQLCommission = strSQLCommission & " GROUP BY TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, TR2.CommissionTotal2, TR3.CommissionTotal3 "
									strSQLCommission = strSQLCommission & " ORDER BY TRAINERS.TrLastName "
								response.write debugSQL(strSQLCommission, "SQL Summary view")
									rsCommission.CursorLocation = 3
									rsCommission.open strSQLCommission, cnWS
									Set rsCommission.ActiveConnection = Nothing
									
									tmpTotal = 0
									com1Total = 0
									com2Total = 0
									if NOT rsCommission.EOF then %>
								<table width="85%" cellspacing="0" style="margin: 0 auto;">										
									<tr class="whiteHeader" style="background-color:<%=session("pageColor4")%>;">
											<td colspan="4" class="whiteHeader" align="left">&nbsp;</td>
											<td colspan="2" class="whiteHeader" align="left"><strong><%=xssStr(allHotWords(6))%></strong></td>
										<%	if request.form("showCommissionSplitDetail")="on" then %>
											<td colspan="2" class="whiteHeader right"><strong><%= hw_Commission %></strong></td>
											<td colspan="2" class="whiteHeader right"><strong><%= hw_Commission2 %></strong></td>
										<%	end if %>
											<td colspan="2" class="whiteHeader right"><strong><%= hw_Commission %></strong></td>
											<td colspan="3" class="whiteHeader" align="left">&nbsp;</td>
									</tr>
										
									<%	do while NOT rsCommission.EOF %>
		
		
										<tr>
											<td colspan="4" class="mainText" align="left">&nbsp;</td>
											<td colspan="2" class="mainText" align="left"><strong><%=FmtTrnNameNew(rsCommission, false)%></strong></td>
										<%	if request.form("showCommissionSplitDetail")="on" then %>
											<td colspan="2" class="mainText right"><%=FmtCurrency(rsCommission("Commission1Total"))%></td>
											<td colspan="2" class="mainText right"><%=FmtCurrency(rsCommission("Commission2Total"))%></td>
										<%	end if %>
										<% if NOT request.form("frmExpReport")="true"  then %>
											<td colspan="2" class="mainText right"><%=FmtCurrency(rsCommission("CommissionTotal"))%></td>
										<% else %>
											<td colspan="2" class="mainText right"><%=FmtNumber(rsCommission("CommissionTotal"))%></td>
										<% end if %>
											<td colspan="3" class="mainText" align="left">&nbsp;</td>
										</tr>
										
										<%	
											com1Total = com1Total + rsCommission("Commission1Total")
											com2Total = com2Total + rsCommission("Commission2Total")
											tmpTotal = tmpTotal + rsCommission("CommissionTotal")
											rsCommission.MoveNext
										loop %>
									<% if NOT request.form("frmExpReport")="true"  then %>
										<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="20"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="1"></td></tr>
									<% end if %>
										<tr>
											<td colspan="4" class="mainText" align="left">&nbsp;</td>
											<td colspan="2" class="mainText" align="left"><strong><%= hw_Total %></strong></td>
										<% if NOT request.form("frmExpReport")="true"  then %>
											<%	if request.form("showCommissionSplitDetail")="on" then %>
											<td colspan="2" class="mainText right"><%=FmtCurrency(com1Total)%></td>
											<td colspan="2" class="mainText right"><%=FmtCurrency(com2Total)%></td>
											<%	end if %>
											<td colspan="2" class="mainText right"><%=FmtCurrency(tmpTotal)%></td>
										<% else %>
											<%	if request.form("showCommissionSplitDetail")="on" then %>
											<td colspan="2" class="mainText right"><%=FmtNumber(com1Total)%></td>
											<td colspan="2" class="mainText right"><%=FmtNumber(com2Total)%></td>
											<%	end if %>
											<td colspan="2" class="mainText right"><%=FmtNumber(tmpTotal)%></td>
										<% end if %>
											<td colspan="3" class="mainText" align="left">&nbsp;</td>
										</tr>
										<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="20"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="1"></td></tr>
									</table>
								<%	end if
								else ' detail view
									strSQLCommission = "SELECT TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, Sales.SaleDate, [Sales Details].Description, Clients.FirstName, Clients.LastName, Clients.ClientID, Clients.RSSID, (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - ISNULL([Sales Details].DiscAmt, 0)) as SalePrice, "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.StdCommissionPercRate IS NULL THEN ROUND((PRODUCTS.StdCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END / CASE WHEN [Sales Details].commissionTrnID IS NULL OR [Sales Details].commissionTrn2ID IS NULL OR [Sales Details].commissionTrn2ID = [Sales Details].commissionTrnID THEN 1 ELSE 2 END as ProdStdPerc, "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.StdCommissionFlatRate IS NULL THEN ROUND(PRODUCTS.StdCommissionFlatRate * [Sales Details].Quantity, 2)  ELSE 0 END / CASE WHEN [Sales Details].commissionTrnID IS NULL OR [Sales Details].commissionTrn2ID IS NULL OR [Sales Details].commissionTrn2ID = [Sales Details].commissionTrnID THEN 1 ELSE 2 END  as ProdStdFlat, "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.PromoCommissionPercRate IS NULL THEN ROUND((PRODUCTS.PromoCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END / CASE WHEN [Sales Details].commissionTrnID IS NULL OR [Sales Details].commissionTrn2ID IS NULL OR [Sales Details].commissionTrn2ID = [Sales Details].commissionTrnID THEN 1 ELSE 2 END  as ProdPromoPerc, "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT PRODUCTS.PromoCommissionFlatRate IS NULL THEN ROUND(PRODUCTS.PromoCommissionFlatRate * [Sales Details].Quantity, 2)  ELSE 0 END / CASE WHEN [Sales Details].commissionTrnID IS NULL OR [Sales Details].commissionTrn2ID IS NULL OR [Sales Details].commissionTrn2ID = [Sales Details].commissionTrnID THEN 1 ELSE 2 END  as ProdPromoFlat, "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.StdTrnCommissionPercRate IS NULL THEN ROUND((TRAINERS.StdTrnCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END / CASE WHEN [Sales Details].commissionTrnID IS NULL OR [Sales Details].commissionTrn2ID IS NULL OR [Sales Details].commissionTrn2ID = [Sales Details].commissionTrnID THEN 1 ELSE 2 END  as TrnStdPerc, "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.StdTrnCommissionFlatRate IS NULL THEN ROUND(TRAINERS.StdTrnCommissionFlatRate * [Sales Details].Quantity, 2)  ELSE 0 END / CASE WHEN [Sales Details].commissionTrnID IS NULL OR [Sales Details].commissionTrn2ID IS NULL OR [Sales Details].commissionTrn2ID = [Sales Details].commissionTrnID THEN 1 ELSE 2 END  as TrnStdFlat, "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.PromoTrnCommissionPercRate IS NULL THEN ROUND((TRAINERS.PromoTrnCommissionPercRate *.01) * (([Sales Details].UnitPrice * CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE [Sales Details].Quantity END) - [Sales Details].DiscAmt), 2) ELSE 0 END / CASE WHEN [Sales Details].commissionTrnID IS NULL OR [Sales Details].commissionTrn2ID IS NULL OR [Sales Details].commissionTrn2ID = [Sales Details].commissionTrnID THEN 1 ELSE 2 END  as TrnPromoPerc, "
									strSQLCommission = strSQLCommission & " CASE WHEN NOT TRAINERS.PromoTrnCommissionFlatRate IS NULL THEN ROUND(TRAINERS.PromoTrnCommissionFlatRate * [Sales Details].Quantity, 2)  ELSE 0 END / CASE WHEN [Sales Details].commissionTrnID IS NULL OR [Sales Details].commissionTrn2ID IS NULL OR [Sales Details].commissionTrn2ID = [Sales Details].commissionTrnID THEN 1 ELSE 2 END  as TrnPromoFlat, "
									strSQLCommission = strSQLCommission & " CASE WHEN TRAINERS.TrainerID = [Sales Details].commissionTrnID THEN 1 WHEN TRAINERS.TrainerID = [Sales Details].commissionTrn2ID THEN 2 ELSE 0 END as CommissionType "
									strSQLCommission = strSQLCommission & " FROM [Sales Details] INNER JOIN TRAINERS ON ([Sales Details].commissionTrnID = TRAINERS.TrainerID OR [Sales Details].commissionTrn2ID = TRAINERS.TrainerID) INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID INNER JOIN SALES on [Sales Details].SaleID = SALES.SaleID INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID"
									strSQLCommission = strSQLCommission & " WHERE ( NOT (PRODUCTS.StdCommissionPercRate IS NULL AND PRODUCTS.StdCommissionFlatRate IS NULL AND PRODUCTS.PromoCommissionPercRate IS NULL AND PRODUCTS.PromoCommissionFlatRate IS NULL)) "

									if request.form("optProdServ")="1" then	'products only
										strSQLCommission = strSQLCommission & " AND [Sales Details].CategoryID > 25 "
									elseif request.form("optProdServ")="2" then
										strSQLCommission = strSQLCommission & " AND [Sales Details].CategoryID < 21 "
									end if

									strSQLCommission = strSQLCommission & " AND Sales.SaleDate >= " & DateSep & cSDate & DateSep & " AND Sales.SaleDate <= " & DateSep & cEDate & DateSep
									'strSQLCommission = strSQLCommission & " AND [Sales Details].Returned = 0 "
									if cLoc<>"0" then
										strSQLCommission = strSQLCommission & " AND [Sales Details].Location = " & cLoc
									end if
									if request.form("trainerID")<>"-1" then
										strSQLCommission = strSQLCommission & " AND TRAINERS.TrainerID = " & request.form("trainerID")
									end if
									if request.form("optInactive")="" then
										strSQL = strSQL & " AND TRAINERS.[Active]=1 " 
									end if
									strSQLCommission = strSQLCommission & " ORDER BY TRAINERS.TrainerID, Sales.SaleDate "
	
								response.write debugSQL(strSQLCommission, "SQL detail view")

									rsCommission.CursorLocation = 3
									rsCommission.open strSQLCommission, cnWS
									Set rsCommission.ActiveConnection = Nothing
									
									totalSales = 0
									grandTotal = 0
									
									if NOT rsCommission.EOF then %>
					<table width="100%" cellspacing="0" class="center-ch">
									<%	do while NOT rsCommission.EOF 
		
											if tmpTrn="" or tmpTrn<>cstr(rsCommission("TrainerID")) then 
												if tmpTrn<>"" then %>
						<% if NOT request.form("frmExpReport")="true"  then %>
							<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="16"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="1"></td></tr>				
						<% end if %>
							<tr class="mainText">
								<td colspan="3"><strong><%=hw_Total %></strong>&nbsp;</td>
    						<% if NOT request.form("frmExpReport")="true"  then %>
								<td align="left">&nbsp;<strong><%=FmtCurrency(tmpSalePrice)%></strong></td>
							<% else %>
								<td align="left"><strong><%=FmtNumber(tmpSalePrice)%></strong></td>
							<% end if %>
								<td colspan="5">&nbsp;</td>
								<td colspan="2" nowrap class="right"><strong>Total Sales:&nbsp;</strong></td>
								<td class="center-ch"><strong><%=totalSales%></strong></td>
								<td colspan="2" class="right"><strong><%= hw_Commission %>:&nbsp;</strong></td>
							<% if NOT request.form("frmExpReport")="true"  then %>
								<td class="right">&nbsp;<strong><%=FmtCurrency(tmpTotal)%></strong></td>
							<% else %>
								<td class="right"><strong><%=FmtNumber(tmpTotal)%></strong></td>
							<% end if %>
							</tr>
							<tr><td colspan="16">&nbsp;</td></tr>
											<%		tmpTotal = 0
													totalSales = 0
													tmpSalePrice = 0
												end if %>
							<tr class="whiteHeader" style="background-color:<%=session("pageColor4")%>;">
								<td colspan="16" align="left">&nbsp;&nbsp;<strong><%=FmtTrnNameNew(rsCommission, false)%></strong></td>
							</tr>
							<tr valign="bottom" class="center-ch">
								<td class="mainText ColumnHeader"><%= hw_Date %></td>
								<td class="mainText ColumnHeader"><%= hw_ItemName %></td>
								<td class="mainText ColumnHeader"><%= hw_ID %></td>
								<%	if request.form("showCommissionSplitDetail")="on" then %>
								<td>&nbsp;</td>
								<%	end if %>
							<% if session("Unamerican") then %>							
								<td class="mainText ColumnHeader"><%= hw_ItemPrice %><br />(<%=hw_ExcludingTax %>)&nbsp;&nbsp;&nbsp;</td>
							<% else %>
								<td class="mainText ColumnHeader"><%= hw_ItemPrice %></td>
							<% end if %>
								
								

								<td class="mainText ColumnHeader"><%=hw_Client %></td>
								<td>&nbsp;</td>
								<td class="mainText ColumnHeader">Item Standard&nbsp;%&nbsp;</td>
								<td class="mainText ColumnHeader">Item Standard&nbsp;Flat&nbsp;</td>
								<td class="mainText ColumnHeader">Item Promo&nbsp;%&nbsp;</td>
								<td class="mainText ColumnHeader">Item Promo&nbsp;Flat&nbsp;</td>
								<td class="mainText ColumnHeader">Staff Standard&nbsp;%&nbsp;</td>
								<td class="mainText ColumnHeader">Staff Standard&nbsp;Flat&nbsp;</td>
								<td class="mainText ColumnHeader">Staff Promo&nbsp;%&nbsp;</td>
								<td class="mainText ColumnHeader">Staff Promo&nbsp;Flat&nbsp;</td>
								<td class="mainText center-ch ColumnHeader"><%= hw_Total %></td>
							</tr>
										<%	end if %>
						<% if NOT request.form("frmExpReport")="true"  then %>
							<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="16"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="1"></td></tr>
						<% end if %>
										<% 	tmpTrn = cstr(rsCommission("TrainerID")) %>
										<tr>
											<td class="mainText"><%=FmtDateShort(rsCommission("SaleDate"))%>&nbsp;</td>
											<td class="mainText"><%=rsCommission("Description")%>&nbsp;</td>
											<td class="mainText"><%=rsCommission("RSSID")%>&nbsp;&nbsp;</td>
										<%	if request.form("showCommissionSplitDetail")="on" then %>
											<td class="mainText">
												&nbsp;
												<% 	if rsCommission("CommissionType")="1" then %>
													<%= hw_Commission %>
												<%	elseif rsCommission("CommissionType")="2" then %>
													<%= hw_Commission2 %>
												<%	end if %>
											</td>
										<%	end if %>
										<% if NOT request.form("frmExpReport")="true"  then %>
											<td class="mainText">&nbsp;<%=FmtCurrency(cdbl(rsCommission("SalePrice")))%></td>
											<td class="mainText"><a href="main_info.asp?id=<%=rsCommission("ClientID")%>&fl=true" title="Click Here to View <%=session("ClientHW")%> Information"><%=rsCommission("FirstName")%>&nbsp;<%=rsCommission("LastName")%></a>&nbsp;</td>
											<td class="mainText">&nbsp;<a href="adm_clt_conlog.asp?clientid=<%=rsCommission("ClientID")%>">[View&nbsp;Logs]</a>&nbsp;</td>
											<td class="mainText">&nbsp;<%=FmtCurrency(rsCommission("ProdStdPerc"))%>&nbsp;</td>
											<td class="mainText">&nbsp;<%=FmtCurrency(rsCommission("ProdStdFlat"))%>&nbsp;</td>
											<td class="mainText">&nbsp;<%=FmtCurrency(rsCommission("ProdPromoPerc"))%>&nbsp;</td>
											<td class="mainText">&nbsp;<%=FmtCurrency(rsCommission("ProdPromoFlat"))%>&nbsp;</td>
											<td class="mainText">&nbsp;<%=FmtCurrency(rsCommission("TrnStdPerc"))%>&nbsp;</td>
											<td class="mainText">&nbsp;<%=FmtCurrency(rsCommission("TrnStdFlat"))%>&nbsp;</td>
											<td class="mainText">&nbsp;<%=FmtCurrency(rsCommission("TrnPromoPerc"))%>&nbsp;</td>
											<td class="mainText">&nbsp;<%=FmtCurrency(rsCommission("TrnPromoFlat"))%>&nbsp;</td>
											<td class="mainText right">&nbsp;<strong><%=FmtCurrency(rsCommission("ProdStdPerc")+rsCommission("ProdStdFlat")+rsCommission("ProdPromoPerc")+rsCommission("ProdPromoFlat")+rsCommission("TrnStdPerc")+rsCommission("TrnStdFlat")+rsCommission("TrnPromoPerc")+rsCommission("TrnPromoFlat"))%></strong></td>
										<% else %>
											<td class="mainText"><%=FmtNumber(rsCommission("SalePrice"))%></td>
											<td class="mainText"><%=rsCommission("FirstName")%>&nbsp;<%=rsCommission("LastName")%></td>
											<td class="mainText">&nbsp;</td>
											<td class="mainText"><%=FmtNumber(rsCommission("ProdStdPerc"))%></td>
											<td class="mainText"><%=FmtNumber(rsCommission("ProdStdFlat"))%></td>
											<td class="mainText"><%=FmtNumber(rsCommission("ProdPromoPerc"))%></td>
											<td class="mainText"><%=FmtNumber(rsCommission("ProdPromoFlat"))%></td>
											<td class="mainText"><%=FmtNumber(rsCommission("TrnStdPerc"))%></td>
											<td class="mainText"><%=FmtNumber(rsCommission("TrnStdFlat"))%></td>
											<td class="mainText"><%=FmtNumber(rsCommission("TrnPromoPerc"))%></td>
											<td class="mainText"><%=FmtNumber(rsCommission("TrnPromoFlat"))%></td>
											<td class="mainText right"><strong><%=FmtNumber(rsCommission("ProdStdPerc")+rsCommission("ProdStdFlat")+rsCommission("ProdPromoPerc")+rsCommission("ProdPromoFlat")+rsCommission("TrnStdPerc")+rsCommission("TrnStdFlat")+rsCommission("TrnPromoPerc")+rsCommission("TrnPromoFlat"))%></strong></td>
										<% end if %>
										</tr>
								<% 			
											tmpTotal = tmpTotal + rsCommission("ProdStdPerc")+rsCommission("ProdStdFlat")+rsCommission("ProdPromoPerc")+rsCommission("ProdPromoFlat")+rsCommission("TrnStdPerc")+rsCommission("TrnStdFlat")+rsCommission("TrnPromoPerc")+rsCommission("TrnPromoFlat")
											tmpSalePrice = tmpSalePrice + rsCommission("SalePrice")
											grandTotal = grandTotal + rsCommission("ProdStdPerc")+rsCommission("ProdStdFlat")+rsCommission("ProdPromoPerc")+rsCommission("ProdPromoFlat")+rsCommission("TrnStdPerc")+rsCommission("TrnStdFlat")+rsCommission("TrnPromoPerc")+rsCommission("TrnPromoFlat")	
											
											totalSalePrice = totalSalePrice + rsCommission("SalePrice")
											totalSales = totalSales + 1
											grandTotalSales = grandTotalSales + 1

											rsCommission.MoveNext
										loop %>
						<% if NOT request.form("frmExpReport")="true"  then %>
							<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="16"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="1"></td></tr>
						<% end if %>
							<tr class="mainText">
								<td colspan="3"><strong><%= hw_Total %></strong>&nbsp;</td>
    						<% if NOT request.form("frmExpReport")="true"  then %>
								<td align="left">&nbsp;<strong><%=FmtCurrency(tmpSalePrice)%></strong></td>
							<% else %>
								<td align="left"><strong><%=FmtNumber(tmpSalePrice)%></strong></td>
							<% end if %>
								<td colspan="5">&nbsp;</td>
                                <td colspan="2" nowrap class="right"><strong>Sales:&nbsp;</strong></td>
								<td class="center-ch"><strong><%=totalSales%></strong></td>
								<td colspan="2" class="right"><strong><%= hw_Commission %>:&nbsp;</strong></td>
							<% if NOT request.form("frmExpReport")="true"  then %>
								<td class="right">&nbsp;<strong><%=FmtCurrency(tmpTotal)%></strong></td>
							<% else %>
								<td class="right"><strong><%=FmtNumber(tmpTotal)%></strong></td>
							<% end if %>
							</tr>
			<% 	if NOT request.form("frmExpReport")="true" then %>
						<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="16"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="1"></td></tr>				
						<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="16"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="1"></td></tr>				
			<%	end if %>
							<tr class="mainText">
								<td colspan="3"><strong>GRAND TOTAL:</strong>&nbsp;</td>
							<% if NOT request.form("frmExpReport")="true"  then %>
								<td align="left">&nbsp;<strong><%=FmtCurrency(totalSalePrice)%></strong></td>
							<% else %>
								<td align="left"><strong><%=FmtNumber(totalSalePrice)%></strong></td>
							<% end if %>
								<td colspan="5">&nbsp;</td>
								<td colspan="2" nowrap class="right"><strong>SALES:&nbsp;</strong></td>
								<td class="center-ch"><strong><%=grandTotalSales%></strong></td>
								<td colspan="2" class="right"><strong><%= Ucase(hw_Commission)%>:&nbsp;</strong></td>
							<% if NOT request.form("frmExpReport")="true"  then %>
								<td class="right">&nbsp;<strong><%=FmtCurrency(grandTotal)%></strong></td>
							<% else %>
								<td class="right"><strong><%=FmtNumber(grandTotal)%></strong></td>
							<% end if %>
							</tr>
						<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="16"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="1"></td></tr>				
						</table>
								<%	end if
								end if							
								rsCommission.close
	%>
			<% if request.form("frmExpReport")="true" then %>		

			<% end if %>
			<% if NOT request.form("frmExpReport")="true" then %>		

						<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="16"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="1"></td></tr>				
						<tr><td colspan="16">&nbsp;</td></tr>
						<%	end if %>


	<%
			end if ' if NOT request.form("frmExpReport")="true" then
	%>

						<tr>
							<td colspan="16">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="16">&nbsp;</td>
						</tr>
					</table>
				</td>
	
			</table>
<% pageEnd %>
<%end if ' access to see page %>
<!-- #include file="post.asp" -->

	<%
		

%>
