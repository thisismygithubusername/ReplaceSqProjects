<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
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
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_GIFT_CARDS") then 
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

		Dim cLoc, totNumGCs, totGCAmt, totGCSaleAmt, rowColor, cSDate, cEDate, ap_view_all_locs
		dim rsEntry
		set rsEntry = Server.CreateObject("ADODB.Recordset")
		
		ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
			
		dim category : category = ""
		if (RQ("category"))<>"" then
			category = RQ("category")
		elseif (RF("category"))<>"" then
			category = RF("category") 
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

		if NOT request.form("frmExpReport")="true" then
		%>
			<!-- #include file="../inc_date_ctrl.asp" -->
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
			<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_gc_outstanding", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
			<%= css(array("SimpleLightBox")) %> 
			<script type="text/javascript">
			function exportReport() {
				document.frmSales.frmExpReport.value = "true";
				document.frmSales.frmGenReport.value = "true";
				<% iframeSubmit "frmSales", "adm_rpt_gc_outstanding.asp" %>
			}
			</script>
		<%
		end if
		
		%>
		
		

		<% if NOT request.form("frmExpReport")="true" then %>
<% pageStart %>
	<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div class="headText breadcrumbs-old" align="left">
			<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
			<%if category<>"" then%>
			<span class="breadcrumb-item">&raquo;</span>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<% end if %>
			<span class="breadcrumb-item">&raquo;</span>
			<%= DisplayPhrase(reportPageTitlesDictionary,"Giftcards") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
			</div>
	<%end if %>
			<table height="100%" width="<%=strPageWidth%>" cellspacing="0">    
			<tr> 
			  <td valign="top" height="100%" width="100%"> 
				<table class="center" cellspacing="0" width="90%" height="100%">
	<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
				<tr>
					<td class="headText" align="left" valign="top">
					  <table width="100%" cellspacing="0">
						<tr>
						  <td class="headText" valign="bottom"><b><%= pp_PageTitle("Gift Cards") %></b></td>
						  <td valign="bottom" class="right" height="26"> </td>
						</tr>
					  </table>
					</td>
				  </tr>
	<%end if %>
				  <tr>
				  	<td class="smallTextBlack">
							<ul style="width:940px; margin:0 auto 15px auto">
							<span id="gctypetext0">
								<li><strong>Unassigned Gift Cards:</strong> The gift card value is associated with the <%=LCASE(session("ClientHW"))%>'s account, not the actual card ID. These have NOT been assigned to a <%=LCASE(session("ClientHW"))%>'s account.</li>
							</span>
							<span id="gctypetext1">
								<li><strong>Assignable Gift Card:</strong> The gift card value is associated with the <%=LCASE(session("ClientHW"))%>'s account, not the actual card ID. These HAVE been assigned to a <%=LCASE(session("ClientHW"))%>'s account.</li>
							</span>
							<span id="gctypetext2">
								<li><strong>Unassigned &  Assigned:</strong> Both assigned and unassigned gift cards.</li>
							</span>
							<%if Session("CR_GC") = 0 then %>
							<span id="gctypetext3">
								<li><strong>Prepaid Gift Cards:</strong> The gift card value is associated with the card's ID, not the <%=LCASE(session("ClientHW"))%>'s account. If the card is lost, the money is lost.</li>
							</span>
							<%end if %>
							</ul>					
					</td>
				  </tr>
				  <tr> 
					<td height="30"  valign="bottom" class="headText">
						<table class="mainText border4 center" cellspacing="0">
							  <form name="frmSales" action="adm_rpt_gc_outstanding.asp" method="POST">
								<input type="hidden" name="frmGenReport" value="">
								<input type="hidden" name="frmExpReport" value="">
								<input type="hidden" name="frmTagClients" value="false">
								<input type="hidden" name="frmTagClientsNew" value="false">
								<input type="hidden" name="frmTagBuyers" value="false">
								<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
									<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
									<input type="hidden" name="category" value="<%=category%>">
								<% end if %>
								<tr> 
								  <td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b>
									&nbsp;<%=xssStr(allHotWords(8))%>:</b>&nbsp;<select name="optSaleLoc" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
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
									
				&nbsp;&nbsp;
				<select name="optGiftCertType" onChange="document.frmSales.submit();">
					<option value="0"<% if request.form("optGiftCertType")="0" then response.write(" selected") end if %>>Unassigned</option>
					<option value="1"<% if request.form("optGiftCertType")="1" then response.write(" selected") end if %>>Assigned</option>
					<option value="2"<% if request.form("optGiftCertType")="2" then response.write(" selected") end if %>>Assigned & Unassigned</option>
					<%if Session("CR_GC") = 0 then %>
					<option value="3"<% if request.form("optGiftCertType")="3" then response.write(" selected") end if %>>Prepaid</option>
					<%end if %>
				</select>
				&nbsp;&nbsp;
				&nbsp; <b>Sort By:</b>&nbsp;
									<select name="optSortBy">
                                          <option value="0" <%if request.form("optSortBy")="0" or request.form("optSortBy")="" then response.write "selected" end if%>><%=xssStr(allHotWords(66))%></option>
                                          <option value="1" <%if request.form("optSortBy")="1" then response.write "selected" end if%>>Purchaser Name</option>
                                          <%if request.form("optGiftCertType")<>"3" then  'CCP 8/25/09 error log fix, prepaids don't have defined recipients%>
                                          <option value="2" <% if request.form("optSortBy")="2" then response.write "selected" end if%>>Recipient Name</option>
                                          <% end if %>
                                          <option value="3" <%if request.form("optSortBy")="3" then response.write "selected" end if%>>Gift Card ID (Assigned)</option>
                                          <option value="4" <%if request.form("optSortBy")="4" then response.write "selected" end if%>>Gift Card ID (System)</option>
                                    </select>							
							</td>
							  </tr>
							<tr valign="middle">
							  <td class="center-ch"  style="background-color:#F2F2F2;">
			<%'if CINT(request.form("optGiftCertType"))>0 then%>
				<span style="color:<%=session("pageColor4")%>;">&nbsp;</span><b>With 
				<% if request.Form("optGiftCertType")="1" then %>
				<select name="optSaleRedeem"><option value="0"<%if request.form("optSaleRedeem")="0" then response.write " selected" end if%>><%=xssStr(allHotWords(66))%></option><option value="1" <%if request.form("optSaleRedeem")="1" then response.write "selected" end if%>>Assign Date</option></select>
				<% else %>
				<%=xssStr(allHotWords(66))%>
				<% end if %>
                <input type="radio" id="optRange1" name="optRange" value="range" <%if request.form("optRange")<>"all" then response.write "checked" end if %> /><label for="optRange1">Range</label>
                <input type="radio" id="optRange2" name="optRange" value="all" <%if request.form("optRange")="all" then response.write "checked" end if %>/><label for="optRange2">All Dates</label>
                <br />
                <div id="rangeOptions" <%if request.form("optRange")="all" then response.write "style=""display:none;""" end if%>>
				Between - <%=xssStr(allHotWords(77))%>: 
				<input onBlur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" type="text"  name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date">
							<script type="text/javascript">
								var cal1 = new tcal({'formname':'frmSales', 'controlname':'requiredtxtDateStart'});
								cal1.a_tpl.yearscroll = true;
							</script>
				<input onBlur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date">
							<script type="text/javascript">
								var cal2 = new tcal({'formname':'frmSales', 'controlname':'requiredtxtDateEnd'});
								cal2.a_tpl.yearscroll = true;
							</script>


				<br />
			<% 'end if %>
			 &nbsp;			  <b>
				<script type="text/javascript">
					document.frmSales.optSaleLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
				</script>
				<% showDateArrows("frmSales") %>
                </div>
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
				 else%>
						<input type="button"  name="TagRecvAdd" value="Tag Receiving <%=session("ClientHW")%>s (Add)" onClick="tagRecvAdd();" <%if request.form("optFilterTagged")="on" then response.write "disabled" end if%>>
						<input type="button"  name="TagRecvNew" value="Tag Receiving <%=session("ClientHW")%>s (New)" onClick="tagRecvNew();" <%if request.form("optFilterTagged")="on" then response.write "disabled" end if%>>
						<input type="button"  name="TagPurchAdd" value="Tag Purchasing <%=session("ClientHW")%>s (Add)" onClick="tagBuyersAdd();" <%if request.form("optFilterTagged")="on" then response.write "disabled" end if%>>
						<input type="button"  name="TagPurchNew" value="Tag Purchasing <%=session("ClientHW")%>s (New)" onClick="tagBuyersNew();" <%if request.form("optFilterTagged")="on" then response.write "disabled" end if%>>
				<%end if%>
				<br />
				&nbsp;&nbsp;<% taggingFilter %>&nbsp;&nbsp;
				<input type="button" name="Button" value="Generate" onClick="genReport();">
				<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
				else%>
						<% exportToExcelButton %>
				<%end if%>
				<% savingButtons "frmSales", "Gift Cards" %>
				</b>				
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
						  <td  colspan="2" valign="top" class="mainTextBig center-ch">
					<table width="100%" cellspacing="0" class="mainText">
					  <tr><td valign="top" align=center>
					<table class="mainText" cellspacing="0" width="85%">
					  <tr> 
						<td valign="top" align=center> 
						  </TD>
					  </TR>
					</TABLE>
					  </TD>
					  </TR>
						<tr>
						  <td  colspan="2" valign="top" class="mainTextBig center-ch">
		<% 
		end if			'end of frmExpreport value check before /head line	  
		%>
		<br /><table id="giftCardsReport" class="mainText" width="90%"  cellspacing="0">
		<%
		if request.form("frmTagClients")="true" then

			if CINT(request.form("optGiftCertType"))=3 then 'prepaid
				strSQL = "SELECT Sales.ClientID AS PurchaserID, Sales.ClientID AS RecipientID FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN tblGiftDebitCardData ON [Sales Details].ItemDebitCardID = tblGiftDebitCardData.DebitCardID WHERE (tblGiftDebitCardData.Amount > 0) "
                if request.Form("optRange")="range" then
				    strSQL = strSQL & "AND (Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") "
    				strSQL = strSQL & "AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") " 				
                end if
				if cLoc<>0 then
					strSQL = strSQL & " AND ([Sales Details].Location = " & cLoc & ") "
				end if
			else	'assignable gift cards
				strSQL = "SELECT DISTINCT CLIENTS.ClientID, Sales.ClientID AS PurchaserID, [PAYMENT DATA].ClientID AS RecipientID "
				'Fixed join from PD.SaleID->Sales to SD.PmtRefNo->PD
				strSQL = strSQL & "FROM Sales INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN [PAYMENT DATA] ON [Sales Details].PmtRefNo = [PAYMENT DATA].PmtRefNo INNER JOIN Location ON [Sales Details].Location = Location.LocationID INNER JOIN CLIENTS AS CLIENTS_1 ON [PAYMENT DATA].ClientID = CLIENTS_1.ClientID INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID "
				strSQL = strSQL & " WHERE (PRODUCTS.GiftCertificate = 1) AND [PAYMENT DATA].Returned=0 "
				if request.form("optGiftCertType")="0" then  
					strSQL = strSQL & " AND ([PAYMENT DATA].ClientID = 1) "
				elseif request.form("optGiftCertType")="1" then
					strSQL = strSQL & " AND ([PAYMENT DATA].ClientID <> 1) "
				end if
                if request.Form("optRange")="range" then
				    if CINT(request.form("optGiftCertType"))=1 then 'assigned
					    if CINT(request.Form("optSaleRedeem"))=0 then 'Look by Sale Date
						    strSQL = strSQL & "AND (Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") "
    						strSQL = strSQL & "AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") " 				
	    				else ' Look by assigned Date
		    				strSQL = strSQL & "AND ([Payment Data].PaymentDate >= " & DateSep & cSDate & DateSep & ") "
			    			strSQL = strSQL & "AND ([Payment Data].PaymentDate <= " & DateSep & cEDate & DateSep & ") " 				
				    	end if
    				else 'Unassigned / Both
	    				strSQL = strSQL & "AND (Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") "
		    			strSQL = strSQL & "AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") " 
			    	end if
                end if
				if cLoc<>0 then
					strSQL = strSQL & " AND ([Sales Details].Location = " & cLoc & ") "
				end if
			end if
		response.write debugSQL(strSQL, "SQL")
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing

			if request.form("frmTagClientsNew")="true" then
				strSQLTag = "DELETE FROM tblClientTag "
				if session("mvarUserID")<>"" then
					strSQLTag = strSQLTag & "WHERE smodeID = " & session("mvarUserID") & ""
				else
					strSQLTag = strSQLTag & "WHERE smodeID = 0"
				end if
				cnWS.execute strSQLTag
			end if

			Do while not rsEntry.eof
				on error resume next
				
				if request.form("frmTagBuyers")="true" AND CLNG(rsEntry("PurchaserID"))<>1 then
					strSQLTag = "INSERT INTO tblClientTag (ClientID, smodeID) VALUES "
					strSQLTag = strSQLTag & "(" & rsEntry("PurchaserID") & ", "
					if session("mvarUserID")<>"" then
						strSQLTag = strSQLTag & session("mvarUserID") & ")"
					else
						strSQLTag = strSQLTag & "0)"
					end if
					cnWS.execute strSQLTag
				elseif request.form("frmTagBuyers")<>"true" AND CLNG(rsEntry("RecipientID"))<>1 then
					strSQLTag = "INSERT INTO tblClientTag (ClientID, smodeID) VALUES "
					strSQLTag = strSQLTag & "(" & rsEntry("RecipientID") & ", "
					if session("mvarUserID")<>"" then
						strSQLTag = strSQLTag & session("mvarUserID") & ")"
					else
						strSQLTag = strSQLTag & "0)"
					end if
					cnWS.execute strSQLTag
				end if
				
				rsEntry.movenext
			loop
        %>
			<script type="text/javascript">
				alert("Resulting <%=jsEscDouble(allHotWords(12))%>s are tagged.");
				document.getElementById('TaggedCount').innerHTML = "<%=getTaggedCount()%>";
			</SCRIPT>
<%
			rsEntry.close
		end if
		''End Tag Clients


		if request.form("frmGenReport")="true" then
			if request.form("frmExpReport")="true" then
				Dim stFilename
				stFilename="attachment; filename=Gift Card Report" & ".xls" 
				Response.ContentType = "application/vnd.ms-excel" 
				Response.AddHeader "Content-Disposition", stFilename 
			end if

			if CINT(request.form("optGiftCertType"))=3 then 'prepaid

				'CB Bug #1869 - SQL Timeout Errors
                'Query for Prepaid Gift Card Payments
                strSQL = "SELECT tblGiftDebitCard.DebitCardExtID, tblGiftDebitCard.DebitCardID, tblGiftDebitCardData.Amount, tblGiftDebitCard.DateIssued, Sales.LocationID, Location.LocationName, Sales.SaleID, Sales.SaleDate, CLIENTS.ClientID, CLIENTS.FirstName, CLIENTS.LastName "
                strSQL = strSQL & "FROM tblGiftDebitCard INNER JOIN tblGiftDebitCardData ON tblGiftDebitCard.DebitCardID = tblGiftDebitCardData.DebitCardID INNER JOIN tblPayments ON tblGiftDebitCardData.DebitCardID = tblPayments.DebitCardID AND tblGiftDebitCardData.SaleID = tblPayments.SaleID INNER JOIN Sales ON tblPayments.SaleID = Sales.SaleID INNER JOIN Location ON Sales.LocationID = Location.LocationID INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID "
				if request.form("optFilterTagged")="on" then
					strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
				end if
                ' = strSQL & "WHERE (tblGiftDebitCardData.Amount < 0) "
                strSQL = strSQL & "WHERE 1=1 "
                if request.Form("optRange")="range" then
				    strSQL = strSQL & "AND (Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") "
				    strSQL = strSQL & "AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") " 
                end if
				if cLoc<>0 then
					strSQL = strSQL & " AND (Location.LocationID = " & cLoc & ") "
				end if
				if request.form("optFilterTagged")="on" then
					if session("mvaruserID")<>"" then
						strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
					else
						strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
					end if
				end if

                strSQL = strSQL & "UNION "

                'Query for Prepaid Gift Cards Sold
                strSQL = strSQL & "SELECT tblGiftDebitCard.DebitCardExtID, tblGiftDebitCard.DebitCardID, tblGiftDebitCardData.Amount, tblGiftDebitCard.DateIssued, Location.LocationID, Location.LocationName, Sales.SaleID, Sales.SaleDate, CLIENTS.ClientID, CLIENTS.FirstName, CLIENTS.LastName "
                strSQL = strSQL & "FROM tblGiftDebitCard INNER JOIN tblGiftDebitCardData ON tblGiftDebitCard.DebitCardID = tblGiftDebitCardData.DebitCardID INNER JOIN [Sales Details] ON tblGiftDebitCardData.SDID = [Sales Details].SDID INNER JOIN Sales ON [Sales Details].SaleID = Sales.SaleID INNER JOIN Location ON [Sales Details].Location = Location.LocationID INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID "
				if request.form("optFilterTagged")="on" then
					strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
				end if
                'strSQL = strSQL & "WHERE (tblGiftDebitCardData.Amount > 0) "
                strSQL = strSQL & "WHERE 1=1 "
                if request.Form("optRange")="range" then
						    strSQL = strSQL & "AND (Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") "
    						strSQL = strSQL & "AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") " 				
                end if
				if cLoc<>0 then
					strSQL = strSQL & " AND (Location.LocationID = " & cLoc & ") "
				end if
				if request.form("optFilterTagged")="on" then
					if session("mvaruserID")<>"" then
						strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
					else
						strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
					end if
				end if

			else	'assignable
			
				strSQL = "SELECT DISTINCT Sales.SaleID, [PAYMENT DATA].PaymentAmount, [PAYMENT DATA].PaymentDate, tblGCMsg.PmtRefNo as IsPrintable, [PAYMENT DATA].PmtRefNo, Sales.ClientID AS PurchaserID, CLIENTS.ClientID, CLIENTS.FirstName, CLIENTS.LastName, Sales.SaleDate, [PAYMENT DATA].ClientCredit, [PAYMENT DATA].ClientID AS RecipientID, [PAYMENT DATA].GiftCardExtID, [PAYMENT DATA].GiftCardNumber, [Sales Details].Location, Location.LocationName, CLIENTS_1.FirstName AS RFirstName, CLIENTS_1.LastName AS RLastName, PRODUCTS.GiftCertificate "	
				'Prior to 8/15/06
				'strSQL = strSQL & "FROM Sales INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID INNER JOIN [PAYMENT DATA] ON Sales.SaleID = [PAYMENT DATA].SaleID INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN Location ON [Sales Details].Location = Location.LocationID INNER JOIN CLIENTS CLIENTS_1 ON [PAYMENT DATA].ClientID = CLIENTS_1.ClientID INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID "
				'Fixed join from PD to Prod to be SD to Prod
				'strSQL = strSQL & "FROM Sales INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID INNER JOIN [PAYMENT DATA] ON Sales.SaleID = [PAYMENT DATA].SaleID INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN Location ON [Sales Details].Location = Location.LocationID INNER JOIN CLIENTS CLIENTS_1 ON [PAYMENT DATA].ClientID = CLIENTS_1.ClientID INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID "
				
				'Fixed join from PD.SaleID->Sales to SD.PmtRefNo->PD
				strSQL = strSQL & "FROM Sales INNER JOIN CLIENTS ON Sales.ClientID = CLIENTS.ClientID INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN [PAYMENT DATA] ON [Sales Details].PmtRefNo = [PAYMENT DATA].PmtRefNo INNER JOIN Location ON [Sales Details].Location = Location.LocationID INNER JOIN CLIENTS AS CLIENTS_1 ON [PAYMENT DATA].ClientID = CLIENTS_1.ClientID INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID "
				strSQL = strSQL & " LEFT OUTER JOIN tblGCMsg ON [PAYMENT DATA].PmtRefNo = tblGCMsg.PmtRefNo "
				if request.form("optFilterTagged")="on" then
					strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
				end if
				
				strSQL = strSQL & " WHERE (PRODUCTS.GiftCertificate = 1) AND (PRODUCTS.DebitCard = 0) AND [PAYMENT DATA].Returned=0 "
				
				if request.form("optGiftCertType")="0" then  
					strSQL = strSQL & " AND ([PAYMENT DATA].ClientID = 1) "
				elseif request.form("optGiftCertType")="1" then
					strSQL = strSQL & " AND ([PAYMENT DATA].ClientID <> 1) "
				end if
				
                if request.Form("optRange")="range" then
				    if CINT(request.form("optGiftCertType"))=1 then 'assigned
					    if CINT(request.Form("optSaleRedeem"))=0 then 'Look by Sale Date
						    strSQL = strSQL & "AND   (Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") "
    						strSQL = strSQL & "AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") " 				
	    				else ' Look by assigned Date
		    				strSQL = strSQL & "AND ([Payment Data].PaymentDate >= " & DateSep & cSDate & DateSep & ") "
			    			strSQL = strSQL & "AND ([Payment Data].PaymentDate <= " & DateSep & cEDate & DateSep & ") " 				
				    	end if
    				else 'Unassigned / Both
	    				strSQL = strSQL & "AND (Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") "
		    			strSQL = strSQL & "AND (Sales.SaleDate <= " & DateSep & cEDate & DateSep & ") " 
			    	end if
                end if
				if cLoc<>0 then
					strSQL = strSQL & " AND ([Sales Details].Location = " & cLoc & ") "
				end if
				if request.form("optFilterTagged")="on" then
					if session("mvaruserID")<>"" then
						strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
					else
						strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
					end if
				end if
			
			end if	'prepaid vs assignable
			
			if request.form("optSortBy")="0" then
				strSQL = strSQL & " ORDER BY SaleDate, CLIENTS.LastName"
			elseif request.form("optSortBy")="1" then
				strSQL = strSQL & " ORDER BY CLIENTS.LastName, SaleDate"
			elseif request.form("optSortBy")="2" then
				strSQL = strSQL & " ORDER BY CLIENTS_1.LastName, SaleDate"
			elseif request.form("optSortBy")="3" then
				if CINT(request.form("optGiftCertType"))=3 then
					strSQL = strSQL & " ORDER BY tblGiftDebitCard.DebitCardExtID, CLIENTS.LastName"
				else
					strSQL = strSQL & " ORDER BY GiftCardExtID, CLIENTS.LastName"
				end if
			elseif request.form("optSortBy")="4" then
				if CINT(request.form("optGiftCertType"))=3 then
					strSQL = strSQL & " ORDER BY tblGiftDebitCard.DebitCardID, CLIENTS.LastName"
				else
					strSQL = strSQL & " ORDER BY GiftCardNumber, CLIENTS.LastName"
				end if
			end if
		
		response.write debugSQL(strSQL, "SQL")
		
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			if NOT rsEntry.EOF then

				if CINT(request.form("optGiftCertType"))=3 then 'prepaid
%>
				<tr>
					<td><strong><%= getHotWord(66)%></strong></td>
					<td><strong><%= getHotWord(115)%></strong></td>
					<td><strong><%=session("ClientHW")%></strong></td>
				<% if cLoc=0 then %>
					<td><strong><%=xssStr(allHotWords(8))%></strong></td>
				<% end if %>
					<td><strong>Gift&nbsp;Card&nbsp;ID&nbsp;(Assigned)</strong></td>
					<td nowrap class="right"><strong><%= getHotWord(35)%>&nbsp;</strong></td>
				</tr>
				<%if request.form("frmExpReport")<>"true" then%>
				<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="8"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" width="100%" height="1"></td></tr>
				<%end if%>
<%
				else	'assignable
%>
				<tr>
					<td><strong><%= getHotWord(66)%></strong></td>
					<td><strong><%= getHotWord(115)%></strong></td>
					<td><strong>Purchased By</strong></td>

				<% if cLoc=0 then %>
					<td><strong><%=xssStr(allHotWords(8))%></strong></td>
				<% end if %>
					<td><strong>Gift&nbsp;Card&nbsp;ID&nbsp;(Assigned)</strong></td>
					<td nowrap class="right"><strong>Gift Card Amount&nbsp;</strong></td>
				<% if request.form("optGiftCertType")<>"0" then %> 
					<td><strong>Recipient&nbsp;</strong></td>
					<td><strong><%= getHotWord(60)%>&nbsp;</strong></td>
					<td><strong>Assigned Date&nbsp;</strong></td>
				<% end if %>
					<td nowrap class="right"><strong>Sale Amount&nbsp;</strong></td>
				</tr>
				<%if request.form("frmExpReport")<>"true" then%>
				<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="11"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" width="100%" height="1"></td></tr>
				<%end if%>
<%
				end if	'prepaid vs assignable

				do while NOT rsEntry.EOF
					totNumGCs = totNumGCs + 1
					if CINT(request.form("optGiftCertType"))=3 then 'prepaid
						totGCAmt = totGCAmt + rsEntry("Amount")
					else
						totGCAmt = totGCAmt + rsEntry("ClientCredit")
						totGCSaleAmt = totGCSaleAmt + rsEntry("PaymentAmount")
					end if
					if rowColor = "#F2F2F2" then
						rowColor = "#FAFAFA"
					else
						rowColor = "#F2F2F2"
					end if

					if CINT(request.form("optGiftCertType"))=3 then 'prepaid
%>
				<tr style="background-color:<%=rowColor%>;">
					<td><%=FmtDateShort(rsEntry("SaleDate"))%></td>
					<td><%if request.form("frmExpReport")<>"true" then%>
							<a href="adm_tlbx_voidedit.asp?saleno=<%=rsEntry("SaleID")%>"><%=Right(rsEntry("SaleID"),4)%></a>
						<%else%>
							<%=rsEntry("SaleID")%>
						<%end if%>
					</td>
					<td>
						<% if request.form("frmExpReport")<>"true" then %><a href="main_info.asp?id=<%=rsEntry("ClientID")%>&fl=true"><% end if %><% response.write(rsEntry("LastName") & ", " & rsEntry("FirstName")) %><% if request.form("frmExpReport")<>"true" then %></a><% end if %>
					</td>
				<% if cLoc=0 then %>
					<td><%=rsEntry("LocationName")%></td>
				<% end if %>
					<td><%=rsEntry("DebitCardExtID")%></td>
					<td nowrap class="right"><%=FmtCurrency(rsEntry("Amount"))%></td>
				</tr>
<%
					else	'assignable
%>
				<tr style="background-color:<%=rowColor%>;">
					<td><%=FmtDateShort(rsEntry("SaleDate"))%></td>
					<td><%if request.form("frmExpReport")<>"true" then%>
							<a href="adm_tlbx_voidedit.asp?saleno=<%=rsEntry("SaleID")%>"><%=Right(rsEntry("SaleID"),4)%></a><% if NOT isNull(rsEntry("IsPrintable")) then %>&nbsp;<a href="../print_gift_card_pdf.asp?sid=<%=getSessionGUID()%>&pmtrefno=<%=rsEntry("PmtRefNo")%>"><img src="<%= contentUrl("/asp/images/printer-20px.png") %>" title="Print Gift Card"></a><% end if %>
						<%else%>
							<%=rsEntry("SaleID")%>
						<%end if%>
					</td>
					<td>
						<%if request.form("frmExpReport")<>"true" then%>
						<a href="main_info.asp?id=<%=rsEntry("ClientID")%>&fl=true"><%=rsEntry("LastName")%>, <%=rsEntry("FirstName")%></a>
						<%else%>
						<%=rsEntry("LastName")%>, <%=rsEntry("FirstName")%>
						<%end if%>
					</td>
				<% if cLoc=0 then %>
					<td><%=rsEntry("LocationName")%></td>
				<% end if %>
					<td><%=rsEntry("GiftCardExtID")%></td>
				<%	if request.form("frmExpReport")<>"true" then %>
					<td class="right"><%=FmtCurrency(rsEntry("ClientCredit"))%>&nbsp;</td>
				<% else %>
					<td class="right"><%=FmtNumber(rsEntry("ClientCredit"))%></td>
				<% end if %>
				<% if request.form("optGiftCertType")<>"0" then %> 
					<td>
						<% if rsEntry("RecipientID") <> "1" then %>
							<% if request.form("frmExpReport")<>"true" then %><a href="main_info.asp?id=<%=rsEntry("RecipientID")%>&fl=true"><% end if %><% response.write(rsEntry("RLastName") & ", " & rsEntry("RFirstName")) %><% if request.form("frmExpReport")<>"true" then %></a><% end if %>
						<% end if %>&nbsp;
					</td>
					<td><% if rsEntry("RecipientID") = "1" then response.write("Unassigned") else response.write("Assigned") end if %></td>
					<td><% if rsEntry("RecipientID") <> "1" then response.Write(FmtDateShort(rsEntry("PaymentDate"))) end if %>&nbsp;</td>
				<% end if %>
				<%	if request.form("frmExpReport")<>"true" then %>
					<td class="right"><%=FmtCurrency(rsEntry("PaymentAmount"))%>&nbsp;</td>
				<% else %>
					<td class="right"><%=FmtNumber(rsEntry("PaymentAmount"))%></td>
				<% end if %>
				</tr>
			<% end if	'prepaid vs assignable %>
<%					
					rsEntry.MoveNext
				loop
				'''Print Totals
%>

			<%if request.form("frmExpReport")<>"true" then%>
				<tr style="background-color:#CCCCCC;"><td colspan="11"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" width="100%" height="1"></td></tr>
				<tr style="background-color:<%=session("pageColor4")%>;"><td colspan="11"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" width="100%" height="1"></td></tr>
			<%end if%>

			<%	if CINT(request.form("optGiftCertType"))<>3 then 'assignable %>

				<tr>
					<td><strong><%= getHotWord(22)%>: <%=totNumGCs%></strong></td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				<% if cLoc=0 then %>
					<td><br /></td>
				<% end if %>
					<td>&nbsp;</td>
				<% if NOT request.form("frmExpReport")="true" then %>
					<td class="right"><strong><%=FmtCurrency(totGCAmt)%>&nbsp;</strong></td>
				<% else %>
					<td class="right"><strong><%=FmtNumber(totGCAmt)%></strong></td>
				<% end if %>
				<% if request.form("optGiftCertType")<>"0" then %> 
					<td>&nbsp;</td>
					<td>&nbsp;</td>
					<td>&nbsp;</td>
				<% end if %>
				<% if NOT request.form("frmExpReport")="true" then %>
					<td class="right"><strong><%=FmtCurrency(totGCSaleAmt)%>&nbsp;</strong></td>
				<% else %>
					<td class="right"><strong><%=FmtNumber(totGCSaleAmt)%></strong></td>
				<% end if %>
				</tr>
			<% else 'CB 49_2633 %>
				<tr>
					<td class="right" colspan="7"><strong><%=FmtCurrency(totGCAmt)%></td>
				</tr>				
			<% end if 'assignable%>

<%				
			else	'EOF / No Results
%>
				<tr>
				  <td colspan="8" class="center-ch"><span style="color:#990000;">There are currently no results for that search.</span></td>
				</tr>
<%
			end if
			rsEntry.close
		end if	'gen Report true
%>
						  </table>
						  </td>
						</tr>
					</TABLE>
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
