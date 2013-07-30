<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
	dim rsEntry, strSQLIns, strSQLDel, rsARC, strSQLARC
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	set rsARC = Server.CreateObject("ADODB.Recordset")
	%>
	<!-- #include file="inc_accpriv.asp" -->
	<!-- #include file="inc_utilities.asp" -->
	<!-- #include file="inc_rpt_save.asp" -->
	<!-- #include file="inc_rpt_tagging.asp" -->
	<%
	if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_ANALYSIS") then 
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

		Dim showDetails, cSDate, cEDate, cLoc, curPrice, tmpPmtRefCount, tmpVisitCount, tmpAvgVisit, tmpARCperVisit
		Dim rowColor, curPackage, curProductID, curTG, curTGID, AvgOverall, RevPackage, totVisitCount, totRev, tmpTotAmtSold
				
		dim category : category = ""
		if (RQ("category"))<>"" then
			category = RQ("category")
		elseif (RF("category"))<>"" then
			category = RF("category") 
		end if

		Function ARCDate()
			strSQLARC = "SELECT Max(DateCreated) as LastArcDate from dyntblAvgRevContr"
			rsARC.CursorLocation = 3
			rsARC.open strSQLARC, cnWS
			Set rsARC.ActiveConnection = Nothing
			If not rsARC.eof then
				if NOT isNULL(rsARC("LastArcDate")) then
					ARCDate=Cstr(rsARC("LastArcDate"))
				else
					ARCDate = "N/A"
				end if
			else
				ARCDate="N/A"
			%>
			<script type="text/javascript">
				alert("Please run Average Revenue Contribution by Series report first to create contribution numbers.");
				javascript:history.go(-1);
			</script>
			<%
			end if
			rsARC.close
		end function

		if RF("requiredtxtDateStart")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cSDate = CDATE(RF("requiredtxtDateStart"))
			Call SetLocale("en-us")
		else
			cSDate = DateAdd("yyyy",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
		end if
	
		if RF("requiredtxtDateEnd")<>"" then
			Call SetLocale(session("mvarLocaleStr"))
				cEDate = CDATE(RF("requiredtxtDateEnd"))
			Call SetLocale("en-us")
		else
			cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
		end if

		If RF("optSaleLoc")<>"" then
			cLoc = CLNG(RF("optSaleLoc"))
		else
			if session("numLocations")>1 then
				cLoc = CLNG(session("curLocation"))
			else
				cLoc = 0
			end if
		end if
		
		showDetails = true
		if NOT RF("frmExpReport")="true" then
		%>
<!-- #include file="pre.asp" -->
			<!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_avg_rev_contr", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %> 
	<style type="text/css">
		table.results-table th
		{
			text-transform: capitalize;
		}
		table.results-table td
		{
			padding: 5px;
		}
	</style>
			<script type="text/javascript">
			function exportReport() { 
				document.frmSales.frmExpReport.value = "true";
				document.frmSales.frmGenReport.value = "true";
				<% iframeSubmit "frmSales", "adm_rpt_avg_contr.asp" %>
			}
			</script>
			
			<!-- #include file="../inc_date_ctrl.asp" -->
			<!-- #include file="inc_help_content.asp" -->
		<%
		end if
		
		%>
		
		
		<% if NOT RF("frmExpReport")="true" then %>
<% pageStart %>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div class="headText breadcrumbs-old" align="left">
			<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
			<%if category<>"" then%>
			<span class="breadcrumb-item">&raquo;</span>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<% end if %>
			<span class="breadcrumb-item">&raquo;</span>
			<%= DisplayPhrase(reportPageTitlesDictionary, "Averagerevenueanalysis") %>
			<% showNewHelpContentIcon("average-revenue-analysis-report") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div></div>
<%end if %>
				<table class="center" cellspacing="0" width="<%=strPageWidth%>" height="100%">
					<tr>
					<td class="headText" align="left" valign="top">
					<table width="100%" cellspacing="0">
						<tr class="headText" height="30"  valign="middle">
						<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
						<td><b><%= pp_PageTitle("Average Revenue Analysis") %></b>
						<!--JM - 49_2447-->
						<% showNewHelpContentIcon("average-revenue-analysis-report") %>
						</td>
						<%end if %>
						<td class="right"><b>Last Run Date: <%=ARCDate()%></b>
						
						</td>
						</tr>
					</table>
					</td>
					</tr>
					<tr> 
					<td height="30"  valign="bottom" class="headText">
					<table class="mainText border4 center" cellspacing="0">
						<form name="frmSales" action="adm_rpt_avg_rev_contr.asp" method="POST">
						<input type="hidden" name="frmGenReport" value="">
						<input type="hidden" name="frmExpReport" value="">
						<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
							<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
							<input type="hidden" name="category" value="<%=category%>">
						<% end if %>
						<tr> 
						<td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b><span style="color:<%=session("pageColor4")%>;">&nbsp;</span>Avg. Revenue Contribution as of: 
						<input onBlur="document.frmSales.submit();" type="text"  name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" disabled>
                        &nbsp;<%=xssStr(allHotWords(7))%>:
						<select name="optTG" onchange="document.frmSales.submit();">
						<option value="0" <%if RF("optTG")="0" then RW "selected" end if%>>All Type Groups</option>
						<%
							strSQL = "SELECT TypegroupID, Typegroup FROM tblTypegroup "
							strSQL = strSQL & "WHERE [Active]=1 "
							strSQL = strSQL & "ORDER BY Typegroup"
							rsEntry.CursorLocation = 3
							rsEntry.open strSQL, cnWS
							Set rsEntry.ActiveConnection = Nothing

							do while NOT rsEntry.EOF
								%>	
									<option value="<%=rsEntry("TypegroupID")%>" <%if RF("optTG")=CSTR(rsEntry("TypegroupID")) then RW "selected" end if%>><%=rsEntry("Typegroup")%></option>
								<%
								rsEntry.MoveNext
							loop
							rsEntry.close
						%>
						</select>		
						<script type="text/javascript">
							document.frmSales.optTG.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(503))%>";
						</script>
						&nbsp; Sort By 
						<select name="optSortBy">
							  <option value="0" <%if RF("optSortBy")="0" then RW "selected" end if%>>Alphabetically</option>
							  <option value="1" <%if RF("optSortBy")="1" then RW "selected" end if%>>Avg Revenue</option>
						</select> 
						&nbsp;
						<input type="button" name="Button" value="Generate" onClick="genReport();"></b>
						<%if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_EXPORT") then
						else%>
						<% exportToExcelButton %>
						<%end if%>
						<% savingButtons "frmSales", "Average Revenue Analysis" %>
						<br />
						&nbsp; * Data created here will be used for Revenue by Class report. <br />
						&nbsp; * Clicking Generate or Export buttons will erase existing revenue contribution data and create new values based on the date.<br />
						&nbsp; * This report is based on series sold in the last 12 months that are expired or used up (cannot be used any more).
						</td>
						</tr>
						</form>
						<script type="text/javascript">
						</script>
					</table>			
					</td>
					</tr>
					<tr> 
					<td valign="top" class="mainTextBig center-ch"> 
					
					
		<% 
		end if			'end of frmExpreport value check before /head line	  
		%>
						<div style="padding:30px;">
						<table id="averageRevenueAnalysisGenTag" class="mainText results-table" width="100%"  cellspacing="0">
		<% 
							if RF("frmGenReport")="true" then 
								if RF("frmExpReport")="true" then
									Dim stFilename
									stFilename="attachment; filename=Average Revenue Contribution by Series as of " & Replace(cEDate,"/","-") & ".xls" 
									Response.ContentType = "application/vnd.ms-excel" 
									Response.AddHeader "Content-Disposition", stFilename 
								end if

								'if cDate(ARCDate())<>CDate(cEdate) then		' If it's run in the same day don't update the values
									strSQL = "SELECT [PRODUCTS].Description, [PAYMENT DATA].ProductID, [PAYMENT DATA].PmtRefNo, ([Sales Details].UnitPrice - IsNull([Sales Details].DiscAmt,0)) as UnitPrice, tblTypegroup.TypegroupID, "
									strSQL = strSQL & "tblTypegroup.Typegroup, [PAYMENT DATA].PaymentAmount, Count([VISIT DATA].VisitRefNo) AS CountOfVisitRefNo "
									strSQL = strSQL & "FROM [PAYMENT DATA] INNER JOIN [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo "
									strSQL = strSQL & "LEFT OUTER JOIN [VISIT DATA] ON [PAYMENT DATA].PmtRefNo = [VISIT DATA].PmtRefNo "
									strSQL = strSQL & "LEFT OUTER JOIN PRODUCTS ON [PAYMENT DATA].ProductID = PRODUCTS.ProductID "
									strSQL = strSQL & "INNER JOIN tblTypegroup ON [PAYMENT DATA].Typegroup = tblTypegroup.TypegroupID "
									strSQL = strSQL & " WHERE ([PAYMENT DATA].ClientCredit=0)  "
                                    strSQL = strSQL & " AND ([PAYMENT DATA].ClientContractID IS NULL OR NOT([PAYMENT DATA].DepositReleaseDate IS NULL)) AND ([PAYMENT DATA].Returned = 0) "
									strSQL = strSQL & "AND ([PAYMENT DATA].PaymentDate >=  " & DateSep & cSDate & DateSep & ") AND ([PAYMENT DATA].PaymentDate <=  " & DateSep & cEDate & DateSep & ") " 
									strSQL = strSQL & "AND ([PAYMENT DATA].Type <> 9) "
									strSQL = strSQL & "AND (([PAYMENT DATA].Remaining <= 0) OR ([PAYMENT DATA].Expdate <= " & DateSep & cEDate & DateSep & ")) "
									strSQL = strSQL & "GROUP BY [PRODUCTS].Description, [PAYMENT DATA].ProductID, [PAYMENT DATA].PmtRefNo, [Sales Details].UnitPrice, [Sales Details].DiscAmt, "
									strSQL = strSQL & "tblTypegroup.TypegroupID, tblTypegroup.Typegroup, [PAYMENT DATA].PaymentAmount "
									strSQL = strSQL & "ORDER BY tblTypegroup.Typegroup, [PRODUCTS].Description, [PAYMENT DATA].ProductID"
									RW debugSQL(strSQL, "SQL")
									rsEntry.CursorLocation = 3
									rsEntry.open strSQL, cnWS, 0, 1
									Set rsEntry.ActiveConnection = Nothing
		
									curProductID=0
									curPackage=""
									tmpPmtRefCount=0
									tmpVisitCount=0
									tmpAvgVisit=0
									tmpARCperVisit=0
									curPrice=0
									curTG=""
									
									'*************First write data into dyntblAvgRevContr table***************************
									if NOT rsEntry.EOF then			'EOF
										strSQLDel = "DELETE FROM dyntblAvgRevContr"
										cnWS.Execute strSQLDel
										
										do while NOT rsEntry.EOF
											if curProductID<>rsEntry("ProductID") then
												If curPackage<>"" then
													if tmpVisitCount<>0 then
														tmpAvgVisit=ROUND(tmpVisitCount / tmpPmtRefCount,2)
														if tmpAvgVisit < 1 then
															tmpAvgVisit = 1
														end if
														if tmpAvgVisit<>0 then
															'tmpARCperVisit = ROUND(curPrice / tmpAvgVisit,2)
															'CB 4/15/2008 Updated to use Average Price
															tmpARCperVisit = ROUND(tmpTotAmtSold / tmpPmtRefCount / tmpAvgVisit,2)
														else
															tmpARCperVisit = 0
														end if
														strSQLIns = "INSERT INTO dyntblAvgRevContr (DateCreated, ProductID, TypePurch, Typegroup, TotalSold, TotalVisits, AvgVisits, CurPrice, AvgRevContr) "
														strSQLIns = strSQLIns & "VALUES (" & DateSep & cEDate & DateSep & ", " &  curProductID & ", N'" & Replace(curPackage, "'", "''") & "', " & curTGID & ", "
														strSQLIns = strSQLIns & tmpPmtRefCount & ", " & tmpVisitCount & ", " & tmpAvgVisit & ", " & tmpTotAmtSold / tmpPmtRefCount & ", " & tmpARCperVisit & ")"
														cnWS.Execute strSQLIns
													end if
												end if
												tmpPmtRefCount=0
												tmpVisitCount=0
												tmpAvgVisit=0
												tmpARCperVisit=0
												curPrice=0
												tmpTotAmtSold = 0
											end if
											curPackage=rsEntry("Description")
											curProductID=rsEntry("ProductID")
											If isnull(rsEntry("UnitPrice")) or rsEntry("UnitPrice")=0 then
												curPrice=rsEntry("PaymentAmount")
											else
												curPrice=rsEntry("UnitPrice")
											end if
											curTg=rsEntry("Typegroup")
											curTGID=rsEntry("TypegroupID")
											tmpTotAmtSold = tmpTotAmtSold + curPrice
											tmpPmtRefCount = tmpPmtRefCount + 1
											tmpVisitCount = tmpVisitCount + rsEntry("CountOfVisitRefNo")
											rsEntry.MoveNext
										loop
										if tmpVisitCount<>0 then ' in case last productID has never been used
											tmpAvgVisit=ROUND(tmpVisitCount / tmpPmtRefCount,2)
											if tmpAvgVisit < 1 then
												tmpAvgVisit = 1
											end if
											if tmpAvgVisit<>0 then
												'tmpARCperVisit = ROUND(curPrice / tmpAvgVisit,2)
												'CB 4/15/2008 Updated to use Average Price
												tmpARCperVisit = ROUND(tmpTotAmtSold / tmpPmtRefCount / tmpAvgVisit,2)
											else
												tmpARCperVisit = 0
											end if
										
											strSQLIns = "INSERT INTO dyntblAvgRevContr (DateCreated, ProductID, TypePurch, Typegroup, TotalSold, TotalVisits, AvgVisits, CurPrice, AvgRevContr) "
											strSQLIns = strSQLIns & "VALUES (" & DateSep & cEDate & DateSep & ", " &  curProductID & ", N'" & Replace(curPackage, "'", "''") & "', " & curTGID & ", "
											strSQLIns = strSQLIns & tmpPmtRefCount & ", " & tmpVisitCount & ", " & tmpAvgVisit & ", " & tmpTotAmtSold / tmpPmtRefCount & ", " & tmpARCperVisit & ")"
											cnWS.Execute strSQLIns
										end if
									end if	'eof
									rsEntry.close
								'end if		'ArcDate()<>cEDate
								
								'Second Read data from dyntblAvgRevContr table that we wrote above
								strSQL = "SELECT dyntblAvgRevContr.DateCreated, dyntblAvgRevContr.ProductID, dyntblAvgRevContr.TypePurch, tblTypeGroup.TypeGroup, tblTypeGroup.TypeGroupID,"
								strSQL = strSQL & "dyntblAvgRevContr.TotalSold, dyntblAvgRevContr.TotalVisits, dyntblAvgRevContr.AvgVisits, dyntblAvgRevContr.CurPrice, dyntblAvgRevContr.AvgRevContr "
								strSQL = strSQL & "FROM dyntblAvgRevContr "
								strSQL = strSQL & "INNER JOIN tblTypeGroup ON dyntblAvgRevContr.Typegroup = tblTypeGroup.TypeGroupID "
								if RF("optTG")<>0 and RF("optTG")<>"" then
									strSQL = strSQL & "WHERE tblTypegroup.TypegroupID = " & CLNG(RF("optTG"))
								end if
								if RF("optSortBy")=0 or RF("optSortBy")="" then
									strSQL = strSQL & "ORDER BY tblTypegroup.Typegroup, dyntblAvgRevContr.TypePurch"
								else
									strSQL = strSQL & "ORDER BY dyntblAvgRevContr.AvgRevContr DESC, dyntblAvgRevContr.TypePurch"
								end if
								
								RW debugSQL(strSQL, "SQL")
								
								rsEntry.CursorLocation = 3
								rsEntry.open strSQL, cnWS, 0, 1
								Set rsEntry.ActiveConnection = Nothing
								
								AvgOverall=0
								RevPackage=0
								totVisitCount=0
								totRev=0
		
								if NOT rsEntry.EOF then			'EOF
									%>
										<% if RF("frmExpReport")="true" then %>
										<tr>
											<td colspan="7" class="maintextbig"><strong>AVERAGE REVENUE CONTRIBUTION BY SERIES as of <%=FmtDateShort(cEDate)%></strong></td>
										</tr>
										<%end if%>
										<tr>
											<td colspan="100%">&nbsp; </td>
										</tr>
										<tr>
											<th  class="left">
											<% if Session("wsType")<>"appt" then %>
												<%=xssStr(allHotWords(7))%>
											<% else %>
												<%=xssStr(allHotWords(19))%>
											<% end if %>
											</th>
											<th  class="left"><%=xssStr(allHotWords(61))%></th>
											<th class="right">Total Sold</th>
											<th class="right">Total Visits</th>
											<th class="right">Avg Visits</th>
											<th class="right"><%=allHotWords(368) %>&nbsp;<%if session("Unamerican") then RW("(" & allHotWords(341) & ")") end if%></th>
											<th class="right">Avg Rev per Visit</th>
										</tr>
<% if NOT RF("frmExpReport")="true" then %>
										<tr>
											<td colspan="7" style="background-color:#666666; padding:0;"><% if NOT RF("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
<% end if %>
									<%
									
									do while NOT rsEntry.EOF
										if rowColor = "#F2F2F2" then
											rowColor = "#FAFAFA"
										else
											rowColor = "#F2F2F2"
										end if
										
										RevPackage=rsEntry("AvgRevContr") * rsEntry("TotalVisits")
										totVisitCount=totVisitCount + rsEntry("TotalVisits")
										totRev=totRev + RevPackage
										
											%>
												<tr style="background-color:<%=rowcolor%>;">
												  <td class="left"><%=rsEntry("Typegroup")%></td>
												  <td class="left"><%=rsEntry("TypePurch")%></td>
												  <td class="right">
													  <%=rsEntry("TotalSold")%>
													</td>
													<td class="right">
														<%=rsEntry("TotalVisits")%>
													</td>
													<td class="right"><%=FormatNumber((rsEntry("TotalVisits")/rsEntry("TotalSold")),1)%></td>
												<% if NOT RF("frmExpReport")="true" then %>
												  <td class="right"><%=FmtCurrency(rsEntry("curPrice"))%></td>
												  <td class="right"><%=FmtCurrency(rsEntry("AvgRevContr"))%></td>
												<% else %>
												  <td><%=FmtNumber(rsEntry("curPrice"))%></td>
												  <td><%=FmtNumber(rsEntry("AvgRevContr"))%></td>
												<% end if %>
												</tr>
											<%
										rsEntry.MoveNext
									loop
									AvgOverall=Round(totRev/totVisitCount,2)
									%>
										<tr>
											<td colspan="100%">&nbsp; </td>
										</tr>
<% if NOT RF("frmExpReport")="true" then %>
										<tr>
											<td colspan="100%" style="background-color:#666666;padding:0;"><% if NOT RF("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
<% end if %>
										<tr height="20" valign="middle">
										  <td colspan="6" class="right"><b>Overall Average Revenue Per Visit:</b></td>
<% if NOT RF("frmExpReport")="true" then %>
										  <td class="right"><b><%=FmtCurrency(AvgOverall)%></b></td>
<% else %>
										  <td class="right"><b><%=FmtNumber(AvgOverall)%></b></td>
<% end if %>
										</tr>
<% if NOT RF("frmExpReport")="true" then %>
										<tr >
											<td colspan="100%" style="background-color: #666666; padding: 0;"><% if NOT RF("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
										</tr>
<% end if %>
									<%
								end if			'eof
								rsEntry.close
								set rsEntry = nothing
							end if		'end of generate report if statement
							%>
						  </table>
						</div>
						  </td>
							</tr>
						</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%'end of frmexport value check
	
end if
%>
