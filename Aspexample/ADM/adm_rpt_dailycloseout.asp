<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%
Server.ScriptTimeout = 300    '5 min (value in seconds)

if request.querystring("pdf") = "true" Then
'  if request.form("sid")<>"" then
'    Response.Cookies("SessionFarmGUID") = request.form("sid")
'  end if
end if

'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
        dim rsEntry, rsEntryB, rsEntry2, rsEntry3
        set rsEntry =  Server.CreateObject("ADODB.Recordset")
        set rsEntryB = Server.CreateObject("ADODB.Recordset")
        set rsEntry2 = Server.CreateObject("ADODB.Recordset")
        set rsEntry3 = Server.CreateObject("ADODB.Recordset")
        %>
        <!-- #include file="inc_rpt_pdf.asp" -->
        <!-- #include file="inc_accpriv.asp" -->
<%      dim doRefresh, mboFormName
        'BQL - 45_1395 added doRefresh, set to true if changing the date should refresh the report
        doRefresh = false
        mboFormName = "document.frmCloseOut"
%>
        <!-- #include file="inc_date_arrows.asp" -->
        <!-- #include file="inc_rpt_tagging.asp" -->
        <!-- #include file="inc_utilities.asp" -->
        <!-- #include file="inc_rpt_save.asp" -->
        <!-- #include file="inc_row_colors.asp" -->
        <%

        if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_CASH_CLOSEOUT") then 
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
        
Dim ChangeSalesAmt, ChangeSalesAmtCash, ChangeSalesAmtCheck, authTotalSwiped, authTotalKeyed, settledTotalSwiped, settledTotalKeyed
Dim ChangeSalesTaxAmt, closeDate, oldCloseDate, salesTotal, locName
Dim amtActCash, amtActCheck, amtCashOverShort, amtCheckOverShort, ss_Fitlink, ss_IncludeTips, ss_Category2, cSDate, tipTotal, memRev, memDisc
Dim SalesArr(), strLocChk, strLoc, intCloseID, ap_view_all_locs, gcName

ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")

strSQL = "SELECT FitLink, IncludeTipsInPayroll, UseCategory2 FROM Studios INNER JOIN tblGenOpts ON Studios.StudioID = tblGenOpts.StudioID WHERE Studios.StudioID=" & session("StudioID")
rsEntry.CursorLocation = 3
rsEntry.open strSQL, cnWS
Set rsEntry.ActiveConnection = Nothing

ss_Category2 = rsEntry("UseCategory2")
ss_IncludeTips = rsEntry("IncludeTipsInPayroll")
ss_Fitlink = rsEntry("FitLink")
rsEntry.close

'RI 58_3377 setting variables to use
strSQL = "SELECT PmtTypes FROM [Payment Types] WHERE Item# = 98"
rsEntry.CursorLocation = 3
rsEntry.open strSQL, cnWS
Set rsEntry.ActiveConnection = Nothing
	gcName = rsEntry("PmtTypes")
rsEntry.Close

if ss_Fitlink then
%>      <!-- #include file="../inc_dbconn_fitlink.asp" -->
<%
end if
		
		dim category : category = ""
		if (RQ("category"))<>"" then
			category = RQ("category")
		elseif (RF("category"))<>"" then
			category = RF("category") 
		end if

        if request.querystring("CloseID")<>"" AND isNum(request.querystring("CloseID")) then
                intCloseID = request.querystring("CloseID")
        elseif request.form("optCloseID")<>"" AND isNum(request.form("optCloseID")) then
                intCloseID = request.form("optCloseID")
        else
                intCloseID = 0
        end if

        if request.form("requiredtxtDateStart")<>"" then
                Call SetLocale(session("mvarLocaleStr"))
                        cSDate = CDATE(request.form("requiredtxtDateStart"))
                Call SetLocale("en-us")
        else
                cSDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
        end if

        if request.form("requiredtxtDateEnd")<>"" then
                Call SetLocale(session("mvarLocaleStr"))
                        cEDate = CDATE(request.form("requiredtxtDateEnd"))
                Call SetLocale("en-us")
        else
                cEDate = DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
        end if
        
        ' If optLocation is provided, use it. If it's blank, no selection has been made
		' or the field is disabled. In that case, if the user has a location, use just
		' that location. Otherwise, use all locations
        if request.form("optLocation")<>"" then
			strLoc = sqlInjectStr(request.Form("optLocation"))
		else
			if session("numLocations")>1 then
				if session("UserLoc") <> 0 then
					strLoc = session("UserLoc")
				else
					strLoc = "0"
				end if
			else
				strLoc = "0"
			end if
		end if
		strLocChk = "," & Replace(strLoc, " ", "") & ","
        
        if intCloseID<>0 then
                dim SalesAmt, StartingCash, ShortCash, ShortCheck, ActualCash, ActualCheck, closeEmpID
        
                strSQL = "SELECT * FROM tblClosedData WHERE CloseID = " & intCloseID
                rsEntry.open strSQL, cnWS, 3
                if NOT rsEntry.EOF then
                        strLoc = rsEntry("Location")
                        closeDate = rsEntry("CloseDate")
                        oldCloseDate = rsEntry("OldCloseDate")
                        SalesAmt = rsEntry("Sales")
                        StartingCash = rsEntry("StartingCash")
                        ShortCash = rsEntry("ShortCash")
                        ShortCheck = rsEntry("ShortCheck")
                        ActualCash = rsEntry("ActualCash")
                        ActualCheck = rsEntry("ActualCheck")
                        closeEmpID = rsEntry("ClosedBy")
                end if
                
                rsEntry.close
        end if ' intCloseID <> 0

        Dim j
        j=0
        strSQL= "SELECT [Payment Types].[Item#], [Payment Types].[PmtTypes] FROM [Payment Types] "
        strSQL = strSQL & "WHERE ((([Payment Types].[CashEQ])=1))"
        rsEntry.open strSQL, cnWS, 3
        Do while not rsEntry.eof
                ReDim Preserve SalesArr(3,j+1)
                SalesArr(0,j)=Cint(rsEntry("Item#"))
                SalesArr(1,j)=rsEntry("PmtTypes")
                j=j+1
                rsEntry.movenext
        Loop
        rsEntry.close
        
        if request.form("optView")="0" then ' use explicit date range instead of closeID
                Call SetLocale(session("mvarLocaleStr"))
                closeDate = cdate(request.form("requiredtxtDateEnd"))
                oldCloseDate = cdate(request.form("requiredtxtDateStart"))
                Call SetLocale("en-us")
        end if

        Dim cEDate, dtLastClose

        function setrowcolor()
                if rowCount = 0 then
                        rowCount = 1
                        setrowcolor = "#F2F2F2"
                else
                        rowCount = 0
                        setrowcolor = "#FAFAFA"
                end if
        end function

        if request.form("frmExpReport")<>"true" then
                %>
<!-- #include file="pre.asp" -->
                <!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "calendar" & dateFormatCode, "adm/adm_rpt_dailycloseout", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %>
                <script type="text/javascript">
                        function exportReport() {
                                document.frmCloseOut.frmExpReport.value = "true";
                                document.frmCloseOut.frmGenReport.value = "true";
                                document.frmCloseOut.frmGenPdf.value = "false";
                                <% iframeSubmit "frmCloseOut", "adm_rpt_dailycloseout.asp" %>
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
                
                function startMinus() {
                        document.frmCloseOut.requiredtxtDateStart.value = '<%=CDATE(dateadd("d", -1, cSDate))%>';
                        dateValidated();
                };
                function startPlus() {
                        document.frmCloseOut.requiredtxtDateStart.value = '<%=CDATE(dateadd("d", 1, cSDate))%>';
                        dateValidated();
                };
                function endMinus() {
                        document.frmCloseOut.requiredtxtDateEnd.value = '<%=CDATE(dateadd("d", -1, cEDate))%>';
                        dateValidated();
                };
                function endPlus() {
                        document.frmCloseOut.requiredtxtDateEnd.value = '<%=CDATE(dateadd("d", 1, cEDate))%>';
                        dateValidated();
                };
                </script>
        <%
        end if
        
        %>
        
        <% if request.form("frmExpReport")<>"true" then %>
<% pageStart %>
<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div class="headText breadcrumbs-old" align="left">
			<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
			<%if category<>"" then%>
			<span class="breadcrumb-item">&raquo;</span>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<% end if %>
			<span class="breadcrumb-item">&raquo;</span>
			<%= DisplayPhrase(reportPageTitlesDictionary,"Dailycloseout") %>
            <% if request.form("frmGenReport")="true" then %>
                 - <%=FmtDateShort(closeDate)%>
            <% end if %>
            <% showNewHelpContentIcon("daily-closeout-report") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
			</div>
<%end if %>

                <table height="100%" width="<%=strPageWidth%>" border="0" cellspacing="0" cellpadding="0">    
                        <tr> 
                                <td valign="top" width="100%"> 
                                        <table align="center" border="0" cellspacing="0" cellpadding="0" width="100%">
										<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
										<tr>
                                                        <td class="headText" align="left" valign="top">
                                                                <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                                        <tr>
                                                                                <td class="headText" valign="bottom"><b><%= pp_PageTitle("Daily Closeout") %>
                                                                                <% if request.form("frmGenReport")="true" then %>
                                                                                        - <%=FmtDateShort(closeDate)%>
                                                                                <% end if %>
                                                                                </b>
                                                                                <!--JM - 49_2447-->
                                                                                <% showNewHelpContentIcon("daily-closeout-report") %>

                                                                                </td>
                                                                        </tr>
                                                                </table>
                                                        </td>
                                                </tr>
										<%end if%>
                                                <tr><td>&nbsp;</td></tr>
                                                <tr> 
                                                        <td valign="top" class="mainText">
                                                                <table class="mainText center border4">
                                                                <form name="frmCloseOut" id="frmCloseOut" action="adm_rpt_dailycloseout.asp" method="POST">
                                                                        <input type="hidden" name="frmGenReport" value="">
                                                                        <input type="hidden" name="frmExpReport" value="">
																		<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
																			<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
																			<input type="hidden" name="category" value="<%=category%>">
																		<% end if %>
                                                                        <tr> 
                                                                                <td align="center" valign="bottom" style="background-color:#F2F2F2;" nowrap>
                                                                                        &nbsp;<%=xssStr(allHotWords(77))%>: 
                                                                                        <input type="text"  name="requiredtxtDateStart" onblur="validateDate(this, '<%=FmtDateShort(cSDate)%>', true)" value="<%=FmtDateShort(cSDate)%>" class="date"/>
                                                                                        <img border=0 src="<%= contentUrl("/asp/adm/images/trans_arrow_grey_lt.gif") %>" width="10" height="10" style="cursor:pointer" onClick="startMinus();">
                                                                        <script type="text/javascript">
                                                                                var cal1 = new tcal({'formname':'frmCloseOut', 'controlname':'requiredtxtDateStart'});
                                                                                cal1.a_tpl.yearscroll = true;
                                                                        </script>
                                                                                        <img border=0 src="<%= contentUrl("/asp/adm/images/trans_arrow_grey_rt.gif") %>" width="10" height="10" style="cursor:pointer" onClick="startPlus();">
                                                                                        &nbsp;<%=xssStr(allHotWords(79))%>: 
                                                                                        <input type="text"  name="requiredtxtDateEnd" onblur="validateDate(this, '<%=FmtDateShort(cEDate)%>', true)" value="<%=FmtDateShort(cEDate)%>" class="date"/>
                                                                                        <img border=0 src="<%= contentUrl("/asp/adm/images/trans_arrow_grey_lt.gif") %>" width="10" height="10" style="cursor:pointer" onClick="endMinus();">
                                                                        <script type="text/javascript">
                                                                                var cal2 = new tcal({'formname':'frmCloseOut', 'controlname':'requiredtxtDateEnd'});
                                                                                cal2.a_tpl.yearscroll = true;
                                                                        </script>
                                                                                        <img border=0 src="<%= contentUrl("/asp/adm/images/trans_arrow_grey_rt.gif") %>" width="10" height="10" style="cursor:pointer" onClick="endPlus();">
                                                                                        &nbsp;
                                                                                        <b><span style="color:<%=session("pageColor4")%>;"></span>&nbsp;                      
                                                                                        <% if session("numLocations")>1 then %>
                                                                                                          &nbsp;<%= xssStr(allHotWords(8)) %>:
                                                                                                                        <select name="optLocation" size="4" multiple="multiple" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
                                                                                                                                <option value="0" <% if strLoc = "" OR InStr(strLocChk, ",0,") > 0 then response.write "selected" end if %>>All Locations</option>
                                                                                                                                
                                                                                        <%      strSQL = "SELECT LocationID, LocationName FROM Location WHERE Active=1 " 
                                                                                        
                                                                                                strSQL = strSQL & " ORDER BY LocationName "
                                                                                                rsEntry2.open strSQL, cnWS, 3
                                                                                                do while not rsEntry2.EOF
                                                                                        %>
                                                                                                                                <option value="<%=rsEntry2("LocationID")%>" <%if InStr(strLocChk, "," & rsEntry2("LocationID") & ",") <> 0 then response.write "selected" end if %>><%=rsEntry2("LocationName")%></option>
                                                                                        <%              rsEntry2.MoveNext
                                                                                                loop
                                                                                                rsEntry2.close                                          
                                                                                        %>
                                                                                                                        </select>
                                                                                        <% end if ' numLocations > 1 %>
                                                                                        &nbsp;&nbsp;&nbsp;
                                                                                        <select name="optView" onchange="dateValidated()">
                                                                                                <option value="0"<% if request.form("optView")="0" then response.write " selected" end if %>>Use Date Range</option>
                                                                                                <option value="1"<% if request.form("optView")="1" then response.write " selected" end if %>>Use Closed Data</option>
                                                                                        </select>
                                                                                        <img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" title="Use Data Range shows all data between the two dates.  Use Closed Data only shows data from a specific close that was performed during that time.">
                                                                                        &nbsp;&nbsp;
                                                                                        <%      if request.form("optView")="1" then %>
                                                                                                                        Closed Data:
                                                                                                &nbsp;<select name="optCloseID" class="textSmall"
                                                                                        <%      strSQL = "SELECT CloseID, CloseDate, Location.LocationName FROM tblClosedData "
                                                                                                strSQL = strSQL & " INNER JOIN Location ON tblClosedData.Location = Location.LocationID "
                                                                                                strSQL = strSQL & " WHERE CAST([tblClosedData].CloseDate AS Date) >= " & DateSep & cSDate & DateSep 
                                                                                                strSQL = strSQL & " AND CAST([tblClosedData].CloseDate AS Date) <= " & DateSep & cEDate & DateSep
                                                                                                
                                                                                                if strLoc <> "" AND strLoc <> "0" then
									                                                                strSQL = strSQL & " AND (tblClosedData.Location IN (" & strLoc & ")) "
								                                                                end if
                                                                                                
                                                                                                strSQL = strSQL & " ORDER BY CloseID "
                                                                                               
																								response.write debugSQL(strSQL, "SQL")
                                                                                                rsEntry2.open strSQL, cnWS, 3
                                                                                                
                                                                                                if NOT rsEntry2.EOF  then
                                                                                                        do while not rsEntry2.EOF
                                                                                        %>                      >
                                                                                                                <option value="<%=rsEntry2("CloseID")%>" <%if cint(intCloseID)=rsEntry2("CloseID") then response.write "selected" end if%>>Close #<%=rsEntry2("CloseID")%> on <%=FmtDateShort(rsEntry2("CloseDate"))%> at <%=rsEntry2("LocationName")%></option>
                                                                                        <%                      rsEntry2.MoveNext
                                                                                                        loop
                                                                                        %>
                                                                                        <%      else ' rsEntry2.EOF %>
                                                                                                                disabled>
                                                                                                                <option value="">No Closed Data for this date range</option>
                                                                                        <%      end if
                                                                                                rsEntry2.close
                                                                                        %>

                                                                                                        </select>
                                                                                        <%      end if ' optView = 1 %>
                                                                                                                        <select name="optEmployeeID">
                                                                                                                                <option value="-1" <% if request.form("optEmployeeID")="-1" or request.form("optEmployeeID")="" then response.write "selected"%>>All Employees</option>
                                                                                                                                <option value="0" <% if request.form("optEmployeeID")="0" then response.write "selected"%>>Owner</option>
                                                                                        <%      strSQL = "SELECT DISTINCT Sales.EmployeeID, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisplayName, TRAINERS.DisplayName, TRAINERS.TrLastName, TRAINERS.TrFirstName  "
                                                                                                strSQL = strSQL & " FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID INNER JOIN TRAINERS ON Sales.EmployeeID = TRAINERS.TrainerID "
                                                                                                strSQL = strSQL & " WHERE TRAINERS.TrainerID<> 0 AND Sales.SaleDate <= " & DateSep & cEDate & DateSep & " AND Sales.SaleDate >= " & DateSep & cSDate & DateSep
                                                                                                'MB bug#6184
													                                            'if request.Form("optLocation") <> "" AND request.Form("optLocation") <> "0" then
                                                                                                if strLoc <> "" AND strLoc <> "0" then
									                                                                strSQL = strSQL & " AND ([Sales Details].Location IN (" & strLoc & ")) "
								                                                                end if
                                                                                                if intCloseID<>0 then
                                                                                                        strSQL = strSQL & " AND Sales.CloseID = " & intCloseID
                                                                                                end if
                                                                                                if NOT ss_IncludeTips then
                                                                                                        strSQL = strSQL & " AND [Sales Details].CategoryID <> 21 " 
                                                                                                end if
                                                                                                strSQL = strSQL & " ORDER BY " & GetTrnOrderBy()
                                                                                               'response.write debugSQL(strSQL, "SQL")
                                                                                                'response.end
                                                                                                rsEntry.open strSQL, cnWS, 3
                                                                                                if NOT rsEntry.EOF then %>
                                                                                                <%      do while NOT rsEntry.EOF %>
                                                                                                                                <option value="<%=rsEntry("EmployeeID")%>" <%if CLNG(request.form("optEmployeeID"))=CLNG(rsEntry("EmployeeID")) then response.write "selected" end if%>><%=FmtTrnNameNew(rsEntry, false)%></option>
                                                                                        <%                      rsEntry.MoveNext
                                                                                                        loop %>
                                                                                                                        
                                                                                        <%      end if %>
																																							</select>						
                                                                                             <% rsEntry.close                                           
                                                                                        %>
                                                                                        <br />
                                                                                        
                                                                                        &nbsp;Include Subcategories&nbsp;<input type="checkbox" name="optSubCat" <%if request.form("optSubCat")="on" then response.write "checked" end if%> />
                                                                                   		<% if ss_Category2 then %>
                                                                                        &nbsp;Include Secondary Categories&nbsp;<input type="checkbox" name="optIncludeCat2" <%if request.form("optIncludeCat2")="on" then response.write "checked" end if%> />
                                                                                       <% end if %>
                                                                                       <br />
                                                                                       &nbsp;
                                                </td></tr><tr class="mainText">       
                                                                                <td align="center" valign="bottom" style="background-color:#F2F2F2;" nowrap>
                                                                                        <% showDateArrows("frmCloseOut") %>
                                                                                        
																																												<table style="margin: 0pt auto;">
																																													<tr>
																																														<td>
																																															<b>Include <%= gcName %> Sales</b>&nbsp;
																																															<input type="checkbox" id="optIncludePrepaid" name="optIncludePrepaid" <%if request.form("optIncludePrepaid")="on" then response.write "checked" end if%> />&nbsp;
																																														</td>
																																														<td><input type="button" name="Button" value="View Closed Data" onClick="showReport();"></td>
																																														<td><input type="button" name="closeDataButton" value="Close Data" onClick="document.location='adm_rpt_cashdrawer_closeout.asp';"></b></td>
																																														<td><% pdfExportButton "frmCloseOut", "DailyCloseout_" & Replace(cSDate, "/", "-") & "_to_" & Replace(cEDate, "/", "-") & ".pdf" %></td>
																																														<td><% savingButtons "frmCloseOut", "Daily Closeout" %></td>
																																													</tr>
																																												</table>
                                                                                </td>
                                                                        </tr>
																																				
                                                                        </form>
                                                                 </table>
                                                        </td>
                                                </tr>
                <% 
                end if                  'end of frmExpreport value check before /head line        
                %>
                                                <tr><td>&nbsp;</td></tr>
<%      if request.form("frmGenReport") = "true" then  ' then show results, duh! %>
                                                <tr> 
                                                        <td valign="top" class="mainTextBig"> 
                                                                <table class="mainText center" width="85%" border="0" cellspacing="0" cellpadding="0">
                                                                  <tr>
                                      <td colspan="18"> <!-- Header for exports -->
<%                    if request.Form("optLocation") = "0" OR request.Form("optLocation") = "" then
                        locName = "All Locations"
                      elseif request.Form("optLocation") = "98" then
                        locName = "Online Store"
                      else
                          strSQL = "SELECT LocationName FROM Location WHERE LocationID IN (" & strLoc & ")"
                          rsEntry.CursorLocation = 3
                          rsEntry.open strSQL, cnWS
                          Set rsEntry.ActiveConnection = Nothing

                          do while NOT rsEntry.EOF
                              locName = rsEntry("LocationName") & ", "
                              rsEntry.MoveNext
                          loop
                          rsEntry.close
                      end if
          
                      if request.Form("frmExpReport")="true" then 
%>                      <strong><%=FmtDateShort(cSDate)%> - <%=FmtDateShort(cEDate)%>&nbsp;&nbsp;<%=locName%></strong>
<%                    end if 
%>                    &nbsp;
        </td>
                        </tr>

                                                                        <tr> 
                                                                                <td class="mainTextBig" valign="top" align="left">
        
        
<%      ChangeSalesAmt=0
        ChangeSalesTaxAmt=0
        
        dim totalDrawer, totalCount, prodLinePrinted, diffUERTotal, trueEarnedRevTotal
        dim categoryTotal, returnsTotal, returnsTaxTotal, tipReturnsTotal, nonreturnsTotal, GCreturnsTotal, GCreturnsTaxTotal, GCnonreturnsTotal, taxTotal, GCTaxTotal, tipsTotal, countTotal, returnCountTotal, GCCountTotal, GCReturnCountTotal
        
        prodLinePrinted = false
        
                if request.form("optView")="1" then 
        
                        strSQL = "SELECT [Payment Types].PmtTypes, tblClosedSalesPmtType.Amt, [Payment Types].Item#, tblClosedData.ActualCash, tblClosedData.ActualCheck, tblClosedSalesPmtType.Tax, "
                        strSQL = strSQL & " Short.ShortCash, Short.ShortCheck "
                        strSQL = strSQL & " FROM tblClosedData INNER JOIN tblClosedSalesPmtType ON tblClosedData.CloseId = tblClosedSalesPmtType.CloseID "
                        strSQL = strSQL & " INNER JOIN [Payment Types] ON tblClosedSalesPmtType.PmtTypeID = [Payment Types].Item# "
                        strSQL = strSQL & " INNER JOIN (SELECT CloseID, ShortCash, ShortCheck FROM tblClosedData) Short ON Short.CloseID = tblClosedData.CloseID "
                        strSQL = strSQL & " WHERE tblClosedData.CloseId = " & intCloseID
                        strSQL = strSQL & " ORDER BY [Payment Types].Item# "
                       'response.write debugSQL(strSQL, "SQL")
                        rsEntry.open strSQL, cnWS, 3
                
                        if NOT rsEntry.EOF then %>

                                                                                        <table class="mainText" width="70%" border="0" cellspacing="0" cellpadding="0" align="center">
                                                                                                <tr>
                                                                                                  <td class="mainText">Closed By:</td>
                                                                                                  <td align="right" class="mainText">&nbsp;</td>
                                                                                                  <td align="right" class="mainText">
                                                                                                        <%      if CLNG(closeEmpID) = 0 then %>
                                                                                                                        Owner
                                                                                                        <%      else 
                                                                                                                        response.write FmtTrnName(closeEmpID)
                                                                                                                end if %>
                                                                                                                
                                                                                                  </td>
                                                                                                </tr>
                                                                                                <tr style="background-color:<%=session("pageColor4")%>;">
                                                                                                  <td>&nbsp; </td>
                                                                                                  <td align="right" class="whiteHeader">Cash Drawer Count</td>
                                                                                                  <td align="right" class="whiteHeader">Your Count</td>
                                                                                                  <td>&nbsp;</td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                  <td class="mainText">Amount in Drawer:</td>
                                                                                                  <td align="right" class="mainText"><%=FmtCurrency(StartingCash)%></td>
                                                                                                  <td align="right" class="mainText"><%=FmtCurrency(StartingCash)%></td>
                                                                                                </tr>
                                
                <%              totalDrawer = totalDrawer + StartingCash
                                totalCount = totalCount + StartingCash
                                do while NOT rsEntry.EOF  
                                        if rsEntry("Item#")= 1 or rsEntry("Item#")=2 then
                                                totalDrawer = totalDrawer + rsEntry("Amt") + Cdbl(rsEntry("Tax"))
                                                if rsEntry("Item#") = 1 then 
                                                        totalCount = totalCount + (rsEntry("Amt") + Cdbl(rsEntry("Tax")) + rsEntry("ShortCash"))
                                                elseif rsEntry("Item#") = 2 then
                                                        totalCount = totalCount + (rsEntry("Amt") + Cdbl(rsEntry("Tax")) + rsEntry("ShortCheck"))
                                                else
                                                        totalCount = totalCount + rsEntry("Amt") + Cdbl(rsEntry("Tax"))
                                                end if  
                                         %>
                                                                                                <tr>
                                                                                                        <td class="mainText"><%=rsEntry("PmtTypes")%></td>
                                                                                                        <td align="right" class="mainText"><%=FmtCurrency(rsEntry("Amt") + Cdbl(rsEntry("Tax")))%></td>
                <%                              if rsEntry("Item#") = 1 then %>
                                                                                                        <td align="right" class="mainText"><%=FmtCurrency(rsEntry("Amt") + Cdbl(rsEntry("Tax")) + rsEntry("ShortCash"))%></td>
                <%                              elseif rsEntry("Item#") = 2 then %>
                                                                                                        <td align="right" class="mainText"><%=FmtCurrency(rsEntry("Amt") + Cdbl(rsEntry("Tax")) + rsEntry("ShortCheck"))%></td>                               
                <%                              else %>
                                                                                                        <td align="right" class="mainText"><%=FmtCurrency(rsEntry("Amt") + Cdbl(rsEntry("Tax")))%></td>
                <%                              end if %>
                                                                                                        <td>&nbsp;</td>
                                                                                                </tr>                           
                <%                      end if
                                        rsEntry.MoveNext
                                loop %>
                                                                                                <tr><td colspan="4" style="height: 1px; line-height:1px;font-size:1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                                                                <tr>
                                                                                                        <td align="left"><b>Difference:</b></td>
                                                                                                        <td colspan="2" align="right"><%=FmtCurrency(totalCount - totalDrawer)%></td>
                                                                                                        <td>&nbsp;</td>
                                                                                                </tr>
                                                                                        </table>&nbsp;
        <%              end if ' rsEntry.EOF 
                        rsEntry.close  %>
                                                                                </td>
                                                                        </tr>
        <%      strSQL = "SELECT SUM(ISNULL(PaymentAmount, 0) + ISNULL(PaymentAmountB, 0)) as SalesTotal, SaleDate FROM SALES WHERE CloseID = " & intCloseID & " GROUP BY Sales.SaleDate "
               'response.write debugSQL(strSQL, "SQL")
                'response.end
                rsEntry.open strSQL, cnWS, 3
                if NOT rsEntry.EOF then
                
                        if rsEntry.RecordCount > 1 then
                                salesTotal = 0 %>

                                                                        <tr>
                                                                                <td class="mainTextBig" valign="top" align="left">&nbsp;
                                                        
                                                                                        <table class="mainText" width="70%" border="0" cellspacing="0" cellpadding="0" align="center">
                                                                                                <tr style="background-color:<%=session("pageColor4")%>;">
                                                                                                        <td colspan="10" align="center" class="whiteHeader"><strong>Sales by Date</strong></td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                        <td align="left" valign="bottom"><strong><%= getHotWord(66)%>&nbsp;</strong></td>
                                                                                                        <td align="right" valign="bottom"><strong><%= getHotWord(22)%></strong>&nbsp;</td>
                                                                                                </tr>
                        <%      do while NOT rsEntry.EOF
                                        salesTotal = salesTotal + rsEntry("SalesTotal") %>
                                                                                                <tr>
                                                                                                        <td align="left" valign="bottom"><%=FmtDateShort(rsEntry("SaleDate"))%>&nbsp;</td>
                                                                                                        <td align="right" valign="bottom"><%=FmtCurrency(rsEntry("SalesTotal"))%>&nbsp;</td>
                                                                                                </tr>
                        <%              rsEntry.MoveNext
                                loop %>
                                                                                                <tr><td colspan="10" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                                                                <tr>
                                                                                                        <td align="left" valign="bottom"><strong><%= getHotWord(22)%>:</strong>&nbsp;</td>
                                                                                                        <td align="right" valign="bottom"><strong><%=FmtCurrency(salesTotal)%></strong>&nbsp;</td>
                                                                                                </tr>
                                                                                        </table>
                                                                                </td>
                                                                        </tr>
<%                      end if ' recordcount > 1
                end if ' rsentry.eof
                rsEntry.close
        end if ' optView = 1 %>
                                                                        <tr>
                                                                                <td class="mainTextBig" valign="top" align="left">&nbsp;
<%      strSQL = "SELECT [Payment Types].CashEQ, [Payment Types].PmtTypes, [Payment Types].Item#, SUM(tblSDPayments.SDPaymentAmount) AS SumOfSDPaymentAmount, "
        strSQL = strSQL & " SUM(CASE WHEN R.SaleID IS NULL THEN 0 ELSE -1*tblSDPayments.SDPaymentAmount END) AS CategoryReturnsAmount, "
        strSQL = strSQL & " SUM(CASE WHEN R.SaleID IS NULL THEN 0 ELSE -1 * (tblSDPayments.ItemTax1 + tblSDPayments.ItemTax2 + tblSDPayments.ItemTax3 + tblSDPayments.ItemTax4 + tblSDPayments.ItemTax5) END) AS CategoryReturnsTaxAmount, "
        strSQL = strSQL & " SUM(CASE WHEN [Sales Details].CategoryID <> 21 THEN 0 ELSE tblSDPayments.SDPaymentAmount END) AS CategoryTipAmount, "
        strSQL = strSQL & " SUM(CASE WHEN [Sales Details].CategoryID = 21 AND R.SaleID IS NOT NULL THEN tblSDPayments.SDPaymentAmount ELSE 0 END) AS CategoryTipReturnsAmount, "
        strSQL = strSQL & " SUM(CASE WHEN R.SaleID IS NULL THEN (tblSDPayments.ItemTax1 + tblSDPayments.ItemTax2 + tblSDPayments.ItemTax3 + tblSDPayments.ItemTax4 + tblSDPayments.ItemTax5) ELSE 0 END) AS CategoryTaxTotal "

    'CB 54_3168 - New Sales Tables    
   strSQL = strSQL & " FROM tblSDPayments INNER JOIN Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID ON tblSDPayments.SDID = [Sales Details].SDID INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID INNER JOIN tblPayments INNER JOIN [Payment Types] ON tblPayments.PaymentMethod = [Payment Types].Item# ON tblSDPayments.PaymentID = tblPayments.PaymentID "

        strSQL = strSQL & " LEFT OUTER JOIN (SELECT DISTINCT [Returns].SaleID FROM [Returns]) R ON Sales.SaleID = R.SaleID "
        strSQL = strSQL & " WHERE "
        if intCloseID<>0 then
                strSQL = strSQL & " SALES.CloseID = " & intCloseID
        else
                strSQL = strSQL & " Sales.SaleDate <= " & DateSep & closeDate & DateSep & " AND Sales.SaleDate >= " & DateSep & oldCloseDate & DateSep
        end if
        'MB bug#6184
        'if request.Form("optLocation") <> "" AND request.Form("optLocation") <> "0" then
        if strLoc <> "" AND strLoc <> "0" then
            strSQL = strSQL & " AND ([Sales Details].Location IN (" & strLoc & ")) "
        end if
        if NOT ss_IncludeTips then
                strSQL = strSQL & " AND [Sales Details].CategoryID <> 21 " 
        end if
        if request.form("optEmployeeID")<>"-1" AND isNum(request.form("optEmployeeID")) then
                strSQL = strSQL & " AND Sales.EmployeeID = " & request.form("optEmployeeID")
        end if
				if request.Form("optIncludePrepaid")<>"on" then
					strSQL = strSQL & " AND PRODUCTS.DebitCard <> 1 "
				end if
        strSQL = strSQL & " GROUP BY [Payment Types].Item#, [Payment Types].PmtTypes, [Payment Types].CashEQ ORDER BY [Payment Types].CashEQ DESC, [Payment Types].Item# " 
       response.write debugSQL(strSQL, "Sales by Payment Type")
      ' response.write strSQL
        'response.end
        rsEntry.open strSQL, cnWS, 3

        if NOT rsEntry.EOF then  %>
        
                                        <table id="dailyCloseoutReport" class="mainText" width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
                                                <tr style="background-color:<%=session("pageColor4")%>;">
                                                        <td colspan="10" align="center" class="whiteHeader"><strong>Sales by Payment Type</strong></td>
                                                </tr>
                                                <tr>
                                                        <td align="left" valign="bottom"><strong>Payment Type&nbsp;</strong></td>
                                                        <td align="right" valign="bottom"><strong>Receipts&nbsp;</strong></td>
                                                        <td align="right" valign="bottom"><strong><%= getHotWord(71)%>&nbsp;</strong></td>
                                                        <td align="right" valign="bottom"><strong>Returns&nbsp;</strong></td>
                                                        <td align="right" valign="bottom"><strong>Return Tax&nbsp;</strong></td>
                                                <%      if ss_IncludeTips then %>
                                                        <td align="right" valign="bottom"><strong>Tips</strong></td>
                                                        <td align="right" valign="bottom"><strong>Return Tips</strong></td>
                                                <%      end if %>
                                                        <td align="right" valign="bottom"><strong><%= getHotWord(22)%></strong>&nbsp;</td>
                                                </tr>
                                
                                <%              dim paymentTotal, CashEQFlag, nonCashEQTotal, nonCashEQReturns, nonCashEQTaxReturns, nonCashEQtax, nonCashEQNonReturns, nonCashEQTips, nonCashEQTipReturns
                                                nonCashEQTotal = 0
                                                CashEQFlag = false
                                                do while NOT rsEntry.EOF

                                                        if CashEQFlag then
                                                                nonCashEQTotal = nonCashEQTotal + rsEntry("SumOfSDPaymentAmount")
                                                                nonCashEQtax = nonCashEQtax + rsEntry("CategoryTaxTotal")
                                                                nonCashEQReturns = nonCashEQReturns + rsEntry("CategoryReturnsAmount") + rsEntry("CategoryTipReturnsAmount")
                                                                nonCashEQTaxReturns = nonCashEQTaxReturns + rsEntry("CategoryReturnsTaxAmount")
                                                                nonCashEQTips = nonCashEQTips + rsEntry("CategoryTipAmount") - rsEntry("CategoryTipReturnsAmount")
                                                                nonCashEQTipReturns = nonCashEQTipReturns + rsEntry("CategoryTipReturnsAmount")
                                                        end if
                                                        if rsEntry("CashEQ")=0 and NOT CashEQFlag then ' print non-CashEQ line
                                                                CashEQFlag = true
                                                                nonCashEQTotal = nonCashEQTotal + rsEntry("SumOfSDPaymentAmount")
                                                                nonCashEQtax = nonCashEQtax + rsEntry("CategoryTaxTotal")
                                                                nonCashEQReturns = nonCashEQReturns + rsEntry("CategoryReturnsAmount") + rsEntry("CategoryTipReturnsAmount")
                                                                nonCashEQTaxReturns = nonCashEQTaxReturns + rsEntry("CategoryReturnsTaxAmount")
                                                                nonCashEQTips = nonCashEQTips + rsEntry("CategoryTipAmount") - rsEntry("CategoryTipReturnsAmount")
                                                                nonCashEQTipReturns = nonCashEQTipReturns + rsEntry("CategoryTipReturnsAmount")
                                %>
                                                <tr><td colspan="10" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                <tr>
                                                        <td align="left"><b>Cash Equivalent Receipts:</b></td>
                                                        <td align="right"><% if paymentTotal+returnsTotal-taxTotal-tipTotal-tipReturnsTotal<>0 then %><strong><%=FmtCurrency(paymentTotal+returnsTotal-taxTotal-tipTotal-tipReturnsTotal)%></strong><% end if %>&nbsp;</td>
                                                        <td align="right"><% if taxTotal<>0 then %><strong><%=FmtCurrency(taxTotal)%></strong><% end if %>&nbsp;</td>
                                                        <td align="right"><% if returnsTotal-returnsTaxTotal<>0 then %><strong><span style="color:#FF0000;"><%=FmtCurrency(returnsTotal-returnsTaxTotal)%></span></strong><% end if %>&nbsp;</td>
                                                        <td align="right"><% if returnsTaxTotal<>0 then %><strong><span style="color:#FF0000;"><%=FmtCurrency(returnsTaxTotal)%></span></strong><% end if %>&nbsp;</td>
                                                <%      if ss_IncludeTips then %>
                                                        <td align="right"><% if tipTotal<>0 then %><strong><%=FmtCurrency(tipTotal)%></strong><% end if %>&nbsp;</td>
                                                        <td align="right"><% if tipReturnsTotal<>0 then %><strong><span style="color:#FF0000;"><%=FmtCurrency(-1*tipReturnsTotal)%></span></strong><% end if %>&nbsp;</td>
                                                <%      end if %>
                                                        <td align="right"><strong><%=FmtCurrency(paymentTotal)%></strong>&nbsp;</td>
                                                </tr>           
                                                <tr><td colspan="10">&nbsp;</td></tr>
                                <%                      end if %>
                                                <tr>
                                                        <td><%=rsEntry("PmtTypes")%></td>
                                                        <td align="right"><% if rsEntry("SumOfSDPaymentAmount")+rsEntry("CategoryReturnsAmount")-rsEntry("CategoryTaxTotal")-rsEntry("CategoryTipAmount")+rsEntry("CategoryTipReturnsAmount")<>0 then %><%=FmtCurrency(rsEntry("SumOfSDPaymentAmount")+rsEntry("CategoryReturnsAmount")-rsEntry("CategoryTaxTotal")-rsEntry("CategoryTipAmount")+rsEntry("CategoryTipReturnsAmount"))%><% end if %>&nbsp;</td>
                                                        <td align="right"><% if rsEntry("CategoryTaxTotal")<>0 then %><%=FmtCurrency(rsEntry("CategoryTaxTotal"))%><% end if %>&nbsp;</td>
                                                        <td align="right"><% if rsEntry("CategoryReturnsAmount")-rsEntry("CategoryReturnsTaxAmount")+rsEntry("CategoryTipReturnsAmount")<>0 then %><span style="color:#FF0000;"><%=FmtCurrency(rsEntry("CategoryReturnsAmount")-rsEntry("CategoryReturnsTaxAmount")+rsEntry("CategoryTipReturnsAmount"))%></span><% end if %>&nbsp;</td>
                                                        <td align="right"><% if rsEntry("CategoryReturnsTaxAmount")<>0 then %><span style="color:#FF0000;"><%=FmtCurrency(rsEntry("CategoryReturnsTaxAmount"))%></span><% end if %>&nbsp;</td>
                                                <%      if ss_IncludeTips then %>
                                                        <td align="right"><% if rsEntry("CategoryTipAmount")-rsEntry("CategoryTipReturnsAmount")<>0 then %><%=FmtCurrency(rsEntry("CategoryTipAmount")-rsEntry("CategoryTipReturnsAmount"))%><% end if %>&nbsp;</td>
                                                        <td align="right"><% if rsEntry("CategoryTipReturnsAmount")<>0 then %><span style="color:#FF0000;"><%=FmtCurrency(-1*rsEntry("CategoryTipReturnsAmount"))%></span><% end if %>&nbsp;</td>
                                                <%      end if %>
                                                        <td align="right"><%=FmtCurrency(rsEntry("SumOfSDPaymentAmount"))%>&nbsp;</td>
                                                </tr>                           
                                <%                      paymentTotal = paymentTotal + rsEntry("SumOfSDPaymentAmount")
                                                        taxTotal = taxTotal + rsEntry("CategoryTaxTotal")
                                                        returnsTotal = returnsTotal + rsEntry("CategoryReturnsAmount") + rsEntry("CategoryTipReturnsAmount")
                                                        returnsTaxTotal = returnsTaxTotal + rsEntry("CategoryReturnsTaxAmount")
                                                        tipTotal = tipTotal + rsEntry("CategoryTipAmount") - rsEntry("CategoryTipReturnsAmount")
                                                        tipReturnsTotal = tipReturnsTotal + rsEntry("CategoryTipReturnsAmount")

                                                        rsEntry.MoveNext
                                                loop
                                                if CashEQFlag then %>
                                                <tr><td colspan="10" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                <tr>
                                                        <td align="left"><b>Non-Cash Equivalent Receipts:</b></td>
                                                        <td align="right"><strong><%=FmtCurrency(nonCashEQTotal+nonCashEQReturns-nonCashEQtax-nonCashEQTips-nonCashEQTipReturns)%></strong>&nbsp;</td>
                                                        <td align="right"><% if nonCashEQtax<>0 then %><strong><%=FmtCurrency(nonCashEQtax)%></strong><% end if %>&nbsp;</td>
                                                        <td align="right"><% if nonCashEQReturns-nonCashEQTaxReturns<>0 then %><strong><span style="color:#FF0000;"><%=FmtCurrency(nonCashEQReturns-nonCashEQTaxReturns)%></span></strong><% end if %>&nbsp;</td>
                                                        <td align="right"><% if nonCashEQTaxReturns<>0 then %><strong><span style="color:#FF0000;"><%=FmtCurrency(nonCashEQTaxReturns)%></span></strong><% end if %>&nbsp;</td>
                                                <%      if ss_IncludeTips then %>
                                                        <td align="right"><% if nonCashEQTips<>0 then %><strong><%=FmtCurrency(nonCashEQTips)%></strong><% end if %>&nbsp;</td>
                                                        <td align="right"><% if nonCashEQTipReturns<>0 then %><strong><span style="color:#FF0000;"><%=FmtCurrency(-1*nonCashEQTipReturns)%></span></strong><% end if %>&nbsp;</td>
                                                <%      end if %>
                                                        <td align="right"><strong><%=FmtCurrency(nonCashEQTotal)%></strong>&nbsp;</td>
                                                        <%      if ccTransExist then %>
                                                        <td align="right">&nbsp;</td>
                                                        <td align="right">&nbsp;</td>
                                                        <%      end if %>
                                                </tr>           
                                                <tr><td colspan="10">&nbsp;</td></tr>
                                        <%      end if %>
                                                <tr><td colspan="10" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                <tr>
                                                        <td align="left"><b>Total Receipts:</b></td>
                                                        <td align="right"><strong><%=FmtCurrency(paymentTotal+returnsTotal-taxTotal-tipTotal-tipReturnsTotal)%></strong>&nbsp;</td>
                                                        <td align="right"><strong><%=FmtCurrency(taxTotal)%></strong>&nbsp;</td>
                                                        <td align="right"><% if returnsTotal-returnsTaxTotal<>0 then %><strong><span style="color:#FF0000;"><%=FmtCurrency(returnsTotal-returnsTaxTotal)%></span></strong><% end if %>&nbsp;</td>
                                                        <td align="right"><% if returnsTaxTotal<>0 then %><strong><span style="color:#FF0000;"><%=FmtCurrency(returnsTaxTotal)%></span></strong><% end if %>&nbsp;</td>
                                                <%      if ss_IncludeTips then %>
                                                        <td align="right"><% if tipTotal<>0 then %><strong><%=FmtCurrency(tipTotal)%></strong><% end if %>&nbsp;</td>
                                                        <td align="right"><% if tipReturnsTotal<>0 then %><strong><span style="color:#FF0000;"><%=FmtCurrency(-1*tipReturnsTotal)%></span></strong><% end if %>&nbsp;</td>
                                                <%      end if %>
                                                        <td align="right"><strong><%=FmtCurrency(paymentTotal)%></strong>&nbsp;</td>
                                                        <%      if ccTransExist then %>
                                                        <td align="right">&nbsp;</td>
                                                        <td align="right">&nbsp;</td>
                                                        <%      end if %>
                                                </tr>           
                                                <tr><td colspan="8">&nbsp;</td></tr>
                                        </table>
                <%      end if 
                        rsEntry.close 

                        %>
                                                                                </td>
                                                                        </tr>
                                                                        <tr>
                                                                                <td class="mainTextBig" valign="top" align="left">&nbsp;
                                <%     taxTotal = 0
													countTotal = 0 
													returnCountTotal = 0
													returnsTotal = 0
													returnsTaxTotal = 0
													nonreturnsTotal = 0
													tipReturnsTotal = 0
													diffUERTotal = 0
													trueEarnedRevTotal = 0

													dim paymentsExist, GCpaymentsExist
													paymentsExist = false
													GCpaymentsExist = false

                                                    ''''''''''''''''''''''''''''''
                                                    'CB TFS Bug 3150 - Optimized SQL before had 2 sub queries that were inner join, made into single main query
                                                    ''''''''''''''''''''''''''''''    
'													strSQL = " SELECT CATS.CategoryID, CATS.CategoryName, " 
'													if request.form("optSubCat")="on" then 
'														strSQL = strSQL & " CATS.SubCategoryName, CATS.SubCategoryID, "
'													end if
'													if request.form("optIncludeCat2")="on" then
'														strSQL = strSQL & " CATS.Cat2Name, CATS.Cat2ID, "
'													end if
'													strSQL = strSQL & " S.CategoryAmount, S.CategoryReturnsAmount, S.CategoryReturnsTaxAmount, S.CategoryTaxTotal, S.CategoryReturnCount, S.CategoryCount " 
'
'													strSQL = strSQL & " FROM (SELECT Categories.CategoryID, Categories.CategoryName " 
'													if request.form("optSubCat")="on" then 
'														strSQL = strSQL & ", SubCategory.SubCategoryID, SubCategory.SubCategoryName "
'													end if
'													if request.form("optIncludeCat2")="on" then
'														strSQL = strSQL & ", CAT2.CategoryID as Cat2ID, CAT2.CategoryName as Cat2Name "
'													end if
'													strSQL = strSQL & " FROM SALES INNER JOIN [Sales Details] ON sales.SaleID = [Sales Details].SaleID "
'													strSQL = strSQL & " INNER JOIN Categories ON [Sales Details].CategoryID = Categories.CategoryID "
'													if request.form("optSubCat")="on" then 
'														strSQL = strSQL & " LEFT OUTER JOIN SubCategory ON [Sales Details].SubCategoryID = SubCategory.SubCategoryID AND Categories.CategoryID = SubCategory.CategoryID "
'													end if
'													if request.form("optIncludeCat2")="on" then
'														strSQL = strSQL & " LEFT OUTER JOIN Categories CAT2 ON CAT2.CategoryID = [Sales Details].Category2ID "
'													end if
'													strSQL = strSQL & " WHERE "
'													strSQL = strSQL & " (Sales.SaleDate <= " & DateSep & cEDate & DateSep & " AND Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") "
'													'MB bug#6184
'													'if request.Form("optLocation") <> "" AND request.Form("optLocation") <> "0" then
'                                                    if strLoc <> "" AND strLoc <> "0" then
'                                                        strSQL = strSQL & " AND ([Sales Details].Location IN (" & strLoc & ")) "
'                                                    end if
'													if request.form("optEmployeeID")<>"0" AND isNum(request.form("optEmployee")) then
'														strSQL = strSQL & " AND Sales.EmployeeID = " & sqlInjectStr(request.form("optEmployee"))
'													end if
'													strSQL = strSQL & " GROUP BY Categories.CategoryID, Categories.CategoryName "
'													if request.form("optSubCat")="on" then 
'														strSQL = strSQL & ", SubCategory.SubCategoryID, SubCategory.SubCategoryName "
'													end if
'													if request.form("optIncludeCat2")="on" then
'														strSQL = strSQL & ", CAT2.CategoryID, CAT2.CategoryName "
'													end if
													

													' The actual category totals
'													strSQL = strSQL & " ) CATS INNER JOIN ("
													strSQL = " SELECT Categories.CategoryID, Categories.CategoryName, " 
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & "ISNULL(SubCategory.SubCategoryID, 0) as SubCategoryID, SubCategory.SubCategoryName, "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & " ISNULL(CAT2.CategoryID, 0) as Cat2ID, CAT2.CategoryName as Cat2Name, "
													end if
													strSQL = strSQL & " SUM(ISNULL(tblSDPayments.SDPaymentAmount, 0)) AS CategoryAmount, "
													strSQL = strSQL & " SUM(CASE WHEN R.SaleID IS NULL THEN 0 ELSE -1*(ISNULL(tblSDPayments.SDPaymentAmount, 0)) END) AS CategoryReturnsAmount, "
													strSQL = strSQL & " SUM(CASE WHEN R.SaleID IS NULL THEN 0 ELSE -1 * (tblSDPayments.ItemTax1 + tblSDPayments.ItemTax2 + tblSDPayments.ItemTax3 + tblSDPayments.ItemTax4 + tblSDPayments.ItemTax5) END) AS CategoryReturnsTaxAmount, "
													strSQL = strSQL & " SUM(CASE WHEN R.SaleID IS NULL THEN (tblSDPayments.ItemTax1 + tblSDPayments.ItemTax2 + tblSDPayments.ItemTax3 + tblSDPayments.ItemTax4 + tblSDPayments.ItemTax5) ELSE 0 END) AS CategoryTaxTotal, "
													strSQL = strSQL & " SUM(CASE WHEN [Sales Details].Quantity > 0 THEN [Sales Details].Quantity ELSE 0 END) as CategoryCount, SUM(CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE 0 END) as CategoryReturnCount "
													'CB 54_3168 - New Sales Tables    
													strSQL = strSQL & " FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID "
													strSQL = strSQL & " INNER JOIN tblSDPayments ON [Sales Details].SDID = tblSDPayments.SDID "
													strSQL = strSQL & " INNER JOIN Categories ON [Sales Details].CategoryID = Categories.CategoryID "
													strSQL = strSQL & " LEFT OUTER JOIN (SELECT DISTINCT [Returns].SaleID FROM [Returns]) R ON Sales.SaleID = R.SaleID "
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & " LEFT OUTER JOIN SubCategory ON [Sales Details].SubCategoryID = SubCategory.SubCategoryID AND Categories.CategoryID = SubCategory.CategoryID "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & " LEFT OUTER JOIN Categories CAT2 ON CAT2.CategoryID = [Sales Details].Category2ID "
													end if
													strSQL = strSQL & " WHERE  "
													if intCloseID<>0 then
														strSQL = strSQL & " SALES.CloseID = " & intCloseID
													else
														strSQL = strSQL & " Sales.SaleDate <= " & DateSep & closeDate & DateSep & " AND Sales.SaleDate >= " & DateSep & oldCloseDate & DateSep
													end if
													strSQL = strSQL & " AND "
													'MB bug#6184
													'if request.Form("optLocation") <> "" AND request.Form("optLocation") <> "0" then
                                                    if strLoc <> "" AND strLoc <> "0" then
													    strSQL = strSQL & " ([Sales Details].Location IN (" & strLoc & ")) AND "
                                                    end if
													if request.form("optEmployeeID")<>"-1" AND isNum(request.form("optEmployeeID")) then
														strSQL = strSQL & " Sales.EmployeeID = " & request.form("optEmployeeID") & " AND "
													end if
													strSQL = strSQL & " [Sales Details].CategoryID NOT BETWEEN 22 AND 23 GROUP BY Categories.CategoryID, Categories.CategoryName " 
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & ", ISNULL(SubCategory.SubCategoryID, 0), SubCategory.SubCategoryName"
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & ", ISNULL(CAT2.CategoryID, 0), CAT2.CategoryName"
													end if
													strSQL = strSQL & " ORDER BY CASE WHEN Categories.CategoryID<=20 THEN 0 WHEN Categories.CategoryID>=25 THEN 1 ELSE 2 END, Categories.CategoryName, Categories.CategoryID "
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & ", SubCategory.SubCategoryName "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & ", CAT2.CategoryName "
													end if
												    response.write debugSQL(strSQL, "Categories")
                                                    'response.Write strSQL
													'response.end
													rsEntry.open strSQL, cnWS, 3

													if NOT rsEntry.EOF then
														paymentsExist = true  %>
        
                                        <table class="mainText" width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
                                                <tr style="background-color:<%=session("pageColor4")%>;">
                                                        <td colspan="15" align="center" class="whiteHeader"><strong>Sales by Category</strong></td>
                                                </tr>
                                                <!--
                                                <tr>
                                                        <td align="center" <%	if request.form("optIncludeCat2")="on" then %>colspan="2"<% end if %>><strong>Service Categories</strong></td>
                                   						     <td align="center"  colspan="3"  style="background-color:#F0F0F0;background-color:#F0F0F0;"><strong>Gross Sales</strong>&nbsp;</td>
                                                        <td align="center" colspan="3"><strong>Returns</strong>&nbsp;</td>
                                                        <td align="center" style="background-color:#F0F0F0;"><strong style="background-color:#F0F0F0;">Net Sales</strong>&nbsp;</td>
																		<%	if false then 'request.form("optSubCat")<>"on" AND request.Form("optIncludeCat2")<>"on" then ' show unearned revenue %>
                                                        <td align="center"><strong>Change In</strong>&nbsp;</td>
                                                        <td align="center"><strong>True Earned</strong>&nbsp;</td>
																		<%	end if %>
                                                </tr>
                                                -->
                                                <tr>
                                                        <td align="left"><strong>Service Categories</strong></td>
                                   						<%	if request.form("optIncludeCat2")="on" then %>
																		  <td align="left"><strong>Category 2</strong></td>
																	<% end if %>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong>Sales</strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%= getHotWord(71)%></strong></td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong>Qty.</strong></td>
                                                        <td align="right"><strong>Returns</strong>&nbsp;</td>
                                                        <td align="right"><strong>Tax</strong>&nbsp;</td>
                                                        <td align="right"><strong>Qty.</strong></td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%= getHotWord(22)%></strong>&nbsp;</td>
																		<%	if false then 'request.form("optSubCat")<>"on" AND request.Form("optIncludeCat2")<>"on" then ' show unearned revenue %>
                                                        <td align="center"><strong>Deferred Revenue</strong>&nbsp;<% if request.form("frmExpReport")<>"true" then %><img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" title="The total amount of prepaid sessions sold, minus the  total amount of prepaid sessions used. Totals are only accurate for data starting May, 2010."><% end if %></td>
                                                        <td align="center"><strong>Revenue</strong>&nbsp;<% if request.form("frmExpReport")<>"true" then %><img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" title="This measures your actual earnings in each category of services, by subtracting the change in deferred revenue from your total sales. Totals are only accurate for data starting May, 2010."><% end if %></td>
																		<%	end if %>
                                                </tr>

                                <%              do while NOT rsEntry.EOF  
																	if CLng(rsEntry("CategoryID"))>25 AND NOT prodLinePrinted then
																		if cateGoryTotal<>0 then ' print service subtotal
																			%>
																				<tr><td colspan="15" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
																				<tr>
																				  <td align="left"><b>Services Subtotals:</b></td>
                       											<%	if request.form("optIncludeCat2")="on" then %>
																				  <td align="left">&nbsp;</td>
																		<% end if %>
																				  <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(categoryTotal+returnsTotal-taxTotal)%></strong>&nbsp;</td>
																				  <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(taxTotal)%></strong>&nbsp;</td>
																				  <td align="right" style="background-color:#F0F0F0;"><strong><%=countTotal%></strong>&nbsp;</td>
																				  <td align="right"><strong><% if returnsTotal-returnsTaxTotal<>0 then %><span style="color:red;"><%=FmtCurrency(returnsTotal-returnsTaxTotal)%></span><% end if %></strong>&nbsp;</td>
																				  <td align="right"><strong><% if returnsTaxTotal<>0 then %><span style="color:red;"><%=FmtCurrency(returnsTaxTotal)%></span><% end if %></strong>&nbsp;</td>
																				  <td align="right"><strong><% if returnCountTotal<>0 then %><span style="color:red;"><%=returnCountTotal%></span><% end if %></strong>&nbsp;</td>
																				  <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(categoryTotal)%></strong>&nbsp;</td>
																		<%	if false then 'request.form("optSubCat")<>"on" AND request.Form("optIncludeCat2")<>"on" then ' show unearned revenue %>
																				  <td align="right"><strong><%=FmtCurrency(diffUERTotal)%></strong></td>
																				  <td align="right"><strong><%=FmtCurrency(categoryTotal-diffUERTotal)%></strong>&nbsp;<% if request.form("frmExpReport")<>"true" then %><img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" title="Total earned service revenue. Totals are only accurate for data starting May, 2010."><% end if %></td>
																		<%	end if %>
																				</tr>
																		<%	if false then 'request.form("optSubCat")<>"on" AND request.Form("optIncludeCat2")<>"on" then ' show unearned revenue %>
																				<tr>
																				  <td colspan="9" align="right"><strong>Change in Deferred Revenue:&nbsp;&nbsp;</strong></td>
																				  <td align="right"><strong><%=FmtCurrency(diffUERTotal)%></strong></td>
																				</tr> 
																				<tr>
																				  <td colspan="9" align="right"><strong>Total Sales by Category:&nbsp;&nbsp;</strong></td>
																				  <td align="right"><strong><%=FmtCurrency(categoryTotal)%></strong></td>
																				</tr> 
																		<%	end if %>
																	<%	end if
																	end if
																	if rsEntry("CategoryID")<>"21" or ss_IncludeTips then
																		categoryTotal = categoryTotal + rsEntry("CategoryAmount")
																		taxTotal = taxTotal + rsEntry("CategoryTaxTotal")
																		returnsTotal = returnsTotal + rsEntry("CategoryReturnsAmount") 
																		returnsTaxTotal = returnsTaxTotal + rsEntry("CategoryReturnsTaxAmount")
																		countTotal = countTotal + rsEntry("CategoryCount")
																		returnCountTotal = returnCountTotal + rsEntry("CategoryReturnCount")
																	end if
																	if CLng(rsEntry("CategoryID"))>25 AND NOT prodLinePrinted then %>
																		<tr>
																				  <td align="right" colspan="15">&nbsp;</td>
																		</tr>
																		<tr>
																				  <td align="left"><strong>Product Categories</strong></td>
																				  <td colspan="8">&nbsp;</td>
																		</tr>
																	<%	prodLinePrinted = true
																	end if %>
																		<tr>
                                                        <td nowrap><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %>
																			<%	response.Write rsEntry("CategoryName")
																	if request.form("optSubCat")="on" AND rsEntry("CategoryID")>20 then
																		if NOT isNull(rsEntry("SubCategoryName")) then
																			response.write " - " & rsEntry("SubCategoryName")
																		end if
																	end if
																				
																						if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>
																			</td>
																<%	if request.form("optIncludeCat2")="on" then %>
																			  <td nowrap><%if rsEntry("CategoryID")=21 then response.write "<i>" end if%>
																	<%	if rsEntry("Cat2Name")<>"" then
																			response.write "&nbsp;" & rsEntry("Cat2Name")
																		end if %>
																			  </td>
																<%	end if %>
                                                        <td align="right" style="background-color:#F0F0F0;"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><%=FmtCurrency(rsEntry("CategoryAmount")+rsEntry("CategoryReturnsAmount")-rsEntry("CategoryTaxTotal"))%><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><%=FmtCurrency(rsEntry("CategoryTaxTotal"))%><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><% if rsEntry("CategoryCount")<>0 then %><%=rsEntry("CategoryCount")%><% end if %>&nbsp;</td>
                                                        <td align="right"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><% if rsEntry("CategoryReturnsAmount")-rsEntry("CategoryReturnsTaxAmount")<>0 then %><span style="color:red;"><%=FmtCurrency(rsEntry("CategoryReturnsAmount")-rsEntry("CategoryReturnsTaxAmount"))%></span><% end if %><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;</td>
                                                        <td align="right"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><% if rsEntry("CategoryReturnsTaxAmount")<>0 then %><span style="color:red;"><%=FmtCurrency(rsEntry("CategoryReturnsTaxAmount"))%></span><% end if %><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;</td>
                                                        <td align="right"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><% if rsEntry("CategoryReturnCount")<>0 then %><span style="color:red;"><%=rsEntry("CategoryReturnCount")%></span><% end if %>&nbsp;</td>
												        <td align="right" style="background-color:#F0F0F0;">
													        <% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><%=FmtCurrency(rsEntry("CategoryAmount"))%><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;
                                                        </td>
                                                </tr>
                                
																<%	rsEntry.MoveNext
                                                loop
                                        end if
                                        rsEntry.close
                                        
													dim AcctGCTotal
													AcctGCTotal = 0
													strSQL = " SELECT CATS.CategoryID, CATS.CateGoryName," 
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & " CATS.SubCategoryName, CATS.SubCategoryID, "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & " CATS.Cat2Name, CATS.Cat2ID, "
													end if
													strSQL = strSQL & " S.CategoryAmount, S.CategoryReturnsAmount, S.CategoryReturnsTaxAmount, S.CategoryTaxTotal, S.CategoryReturnCount, S.CategoryCount " 
													strSQL = strSQL & " FROM (SELECT Categories.CategoryID, "
													'RI 58_3377 
													strSQL = strSQL & "CASE WHEN DebitCard = 0 THEN Categories.CategoryName ELSE '" & sqlInjectStr(gcName) & "' END AS CategoryName "

													if request.form("optSubCat")="on" then 
														strSQL = strSQL & ", SubCategory.SubCategoryID, SubCategory.SubCategoryName "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & ", CAT2.CategoryID as Cat2ID, CAT2.CategoryName as Cat2Name "
													end if
													strSQL = strSQL & " FROM SALES INNER JOIN [Sales Details] ON sales.SaleID = [Sales Details].SaleID "
													strSQL = strSQL & " INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID "
													strSQL = strSQL & " INNER JOIN Categories ON [Sales Details].CategoryID = Categories.CategoryID "
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & " LEFT OUTER JOIN SubCategory ON [Sales Details].SubCategoryID = SubCategory.SubCategoryID AND Categories.CategoryID = SubCategory.CategoryID "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & " LEFT OUTER JOIN Categories CAT2 ON CAT2.CategoryID = [Sales Details].Category2ID "
													end if
													strSQL = strSQL & " WHERE "
													strSQL = strSQL & " (Sales.SaleDate <= " & DateSep & cEDate & DateSep & " AND Sales.SaleDate >= " & DateSep & cSDate & DateSep & ") "
													'MB bug#6184
													'if request.Form("optLocation") <> "" AND request.Form("optLocation") <> "0" then
                                                    if strLoc <> "" AND strLoc <> "0" then
                                                        strSQL = strSQL & " AND ([Sales Details].Location IN (" & strLoc & ")) "
                                                    end if
													if request.form("optEmployeeID")<>"0" AND isNum(request.form("optEmployee")) then
														strSQL = strSQL & " AND Sales.EmployeeID = " & sqlInjectStr(request.form("optEmployee"))
													end if
													if request.Form("optIncludePrepaid")<>"on" then
														strSQL = strSQL & " AND PRODUCTS.DebitCard <> 1 "
													end if
													strSQL = strSQL & " GROUP BY PRODUCTS.DebitCard, Categories.CategoryID, Categories.CategoryName "
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & ", SubCategory.SubCategoryID, SubCategory.SubCategoryName "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & ", CAT2.CategoryID, CAT2.CategoryName "
													end if
													strSQL = strSQL & " ) CATS INNER JOIN ("
													strSQL = strSQL & " SELECT Categories.CategoryID, " 
													strSQL = strSQL & "CASE WHEN DebitCard = 0 THEN Categories.CategoryName ELSE '" & sqlInjectStr(gcName) & "' END AS CategoryName, "
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & " ISNULL(SubCategory.SubCategoryID, 0) as SubCategoryID, "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & " ISNULL(CAT2.CategoryID, 0) as Cat2ID, "
													end if
													strSQL = strSQL & " SUM(ISNULL(tblSDPayments.SDPaymentAmount, 0)) AS CategoryAmount, "
													strSQL = strSQL & " SUM(CASE WHEN R.SaleID IS NULL THEN 0 ELSE -1*(ISNULL(tblSDPayments.SDPaymentAmount, 0)) END) AS CategoryReturnsAmount, "
													strSQL = strSQL & " SUM(CASE WHEN R.SaleID IS NULL THEN 0 ELSE -1 * (tblSDPayments.ItemTax1 + tblSDPayments.ItemTax2 + tblSDPayments.ItemTax3 + tblSDPayments.ItemTax4 + tblSDPayments.ItemTax5) END) AS CategoryReturnsTaxAmount, "
													strSQL = strSQL & " SUM(CASE WHEN R.SaleID IS NULL THEN (tblSDPayments.ItemTax1 + tblSDPayments.ItemTax2 + tblSDPayments.ItemTax3 + tblSDPayments.ItemTax4 + tblSDPayments.ItemTax5) ELSE 0 END) AS CategoryTaxTotal, "
													strSQL = strSQL & " SUM(CASE WHEN [Sales Details].Quantity > 0 THEN [Sales Details].Quantity ELSE 0 END) as CategoryCount, SUM(CASE WHEN [Sales Details].Quantity < 0 THEN [Sales Details].Quantity * -1 ELSE 0 END) as CategoryReturnCount "
													'CB 54_3168 - New Sales Tables    
													strSQL = strSQL & " FROM Sales INNER JOIN [Sales Details] ON Sales.SaleID = [Sales Details].SaleID "
													strSQL = strSQL & " INNER JOIN tblSDPayments ON [Sales Details].SDID = tblSDPayments.SDID "
													strSQL = strSQL & " INNER JOIN PRODUCTS ON [Sales Details].ProductID = PRODUCTS.ProductID "
													strSQL = strSQL & " INNER JOIN Categories ON [Sales Details].CategoryID = Categories.CategoryID "
													strSQL = strSQL & " LEFT OUTER JOIN (SELECT DISTINCT [Returns].SaleID FROM [Returns]) R ON Sales.SaleID = R.SaleID "
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & " LEFT OUTER JOIN SubCategory ON [Sales Details].SubCategoryID = SubCategory.SubCategoryID AND Categories.CategoryID = SubCategory.CategoryID "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & " LEFT OUTER JOIN Categories CAT2 ON CAT2.CategoryID = [Sales Details].Category2ID "
													end if
													strSQL = strSQL & " WHERE  "
													if intCloseID<>0 then
														strSQL = strSQL & " SALES.CloseID = " & intCloseID
													else
														strSQL = strSQL & " Sales.SaleDate <= " & DateSep & closeDate & DateSep & " AND Sales.SaleDate >= " & DateSep & oldCloseDate & DateSep
													end if
													strSQL = strSQL & " AND "
													if request.Form("optIncludePrepaid")<>"on" then
														strSQL = strSQL & " PRODUCTS.DebitCard <> 1 AND "
													end if
													'MB bug#6184
													'if request.Form("optLocation") <> "" AND request.Form("optLocation") <> "0" then
                                                    if strLoc <> "" AND strLoc <> "0" then
                                                        strSQL = strSQL & " ([Sales Details].Location IN (" & strLoc & ")) AND "
                                                    end if
													if request.form("optEmployeeID")<>"-1" AND isNum(request.form("optEmployeeID")) then
														strSQL = strSQL & " Sales.EmployeeID = " & request.form("optEmployeeID") & " AND "
													end if
													strSQL = strSQL & " [Sales Details].CategoryID BETWEEN 22 AND 23 GROUP BY DebitCard, Categories.CategoryID, Categories.CategoryName " 
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & ", ISNULL(SubCategory.SubCategoryID, 0) "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & ", ISNULL(CAT2.CategoryID, 0) "
													end if
													strSQL = strSQL & " ) S ON S.CategoryName = CATS.CategoryName AND S.CategoryID = CATS.CategoryID " 
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & " AND ISNULL(CATS.SubCategoryID, 0) = ISNULL(S.SubCategoryID, 0) "
													end if
													if request.form("optIncludeCat2")="on" then 
														strSQL = strSQL & " AND ISNULL(CATS.Cat2ID, 0) = ISNULL(S.Cat2ID, 0) "
													end if				
													strSQL = strSQL & " ORDER BY CASE WHEN CATS.CategoryID<=20 THEN 0 WHEN CATS.CategoryID>=25 THEN 1 ELSE 2 END, CATS.CategoryName, CATS.CategoryID "
													if request.form("optSubCat")="on" then 
														strSQL = strSQL & ", CATS.SubCategoryName "
													end if
													if request.form("optIncludeCat2")="on" then
														strSQL = strSQL & ", CATS.Cat2Name "
													end if
												    response.write debugSQL(strSQL, "Sales By Cat")
													'response.end
													rsEntry.open strSQL, cnWS, 3

													if NOT rsEntry.EOF then
                                                if paymentsExist then %>
                                                <tr><td colspan="15" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                <tr>
                                                        <td align="left"><b>Total Accrued Sales:</b></td>
                                   						<%	if request.form("optIncludeCat2")="on" then %>
																		  <td align="left">&nbsp;</td>
																	<% end if %>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(categoryTotal+returnsTotal-taxTotal)%></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(taxTotal)%></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=countTotal%></strong>&nbsp;</td>
                                                        <td align="right"><strong><% if returnsTotal-returnsTaxTotal<>0 then %><span style="color:red;"><%=FmtCurrency(returnsTotal-returnsTaxTotal)%></span><% end if %></strong>&nbsp;</td>
                                                        <td align="right"><strong><% if returnsTaxTotal<>0 then %><span style="color:red;"><%=FmtCurrency(returnsTaxTotal)%></span><% end if %></strong>&nbsp;</td>
                                                        <td align="right"><strong><% if returnCountTotal<>0 then %><span style="color:red;"><%=returnCountTotal%></span><% end if %></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(categoryTotal)%></strong>&nbsp;</td>
																		<%	if false then 'request.form("optSubCat")<>"on" AND request.Form("optIncludeCat2")<>"on" then ' show unearned revenue %>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(diffUERTotal)%></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(categoryTotal-diffUERTotal)%></strong>&nbsp;</td>
																		<% end if %>
                                                </tr>           
                                                <tr><td colspan="15">&nbsp;</td></tr>
                                <%              else %>
                                        <table class="mainText" width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
                                                <tr style="background-color:<%=session("pageColor4")%>;">
                                                        <td colspan="15" align="center" class="whiteHeader"><strong>Sales by Category</strong></td>
                                                </tr>

                                <%              end if
                                                GCpaymentsExist = true
                                                do while NOT rsEntry.EOF
                                                        if rsEntry("CategoryID")<>"21" or ss_IncludeTips then
                                                                categoryTotal = categoryTotal + rsEntry("CategoryAmount")
                                                                taxTotal = taxTotal + rsEntry("CategoryTaxTotal")
                                                                returnsTotal = returnsTotal + rsEntry("CategoryReturnsAmount")
                                                                returnsTaxTotal = returnsTaxTotal + rsEntry("CategoryReturnsTaxAmount")
                                                                countTotal = countTotal + rsEntry("CategoryCount")
                                                                returnCountTotal = returnCountTotal + rsEntry("CategoryReturnCount")
                                                                
                                                                AcctGCTotal = AcctGCTotal + rsEntry("CategoryAmount")
                                                                GCTaxTotal = GCTaxTotal + rsEntry("CategoryTaxTotal")
                                                                GCreturnsTotal = GCreturnsTotal + rsEntry("CategoryReturnsAmount")
                                                                GCreturnsTaxTotal = GCreturnsTaxTotal + rsEntry("CategoryReturnsTaxAmount")
                                                                GCCountTotal = GCCountTotal + rsEntry("CategoryCount")
                                                                GCReturnCountTotal = GCReturnCountTotal + rsEntry("CategoryReturnCount")
                                                        end if  %>
                                                        <tr>
																					  <td><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %>
																						<%	response.Write rsEntry("CategoryName")
																							if request.form("optSubCat")="on" AND rsEntry("CategoryID")>20 then
																								if NOT isNull(rsEntry("SubCategoryName")) then
																									response.write " - " & rsEntry("SubCategoryName")
																								end if
																							end if
																							
																							if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>
																						</td>
																					<%	if request.form("optIncludeCat2")="on" then %>
																						  <td nowrap><%if rsEntry("CategoryID")=21 then response.write "<i>" end if%>
																						<%	if rsEntry("Cat2Name")<>"" then
																								response.write "&nbsp;" & rsEntry("Cat2Name")
																							end if %>
																						  </td>
																					<%	end if %>
                                                                <td align="right" style="background-color:#F0F0F0;"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><%=FmtCurrency(rsEntry("CategoryAmount")+rsEntry("CategoryReturnsAmount")-rsEntry("CategoryTaxTotal"))%><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;</td>
                                                                <td align="right" style="background-color:#F0F0F0;"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><%=FmtCurrency(rsEntry("CategoryTaxTotal"))%><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;</td>
                                                                <td align="right" style="background-color:#F0F0F0;"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><% if rsEntry("CategoryCount")<>0 then %><%=rsEntry("CategoryCount")%><% end if %>&nbsp;</td>
                                                                <td align="right"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><% if rsEntry("CategoryReturnsAmount")-rsEntry("CategoryReturnsTaxAmount")<>0 then %><span style="color:red;"><%=FmtCurrency(rsEntry("CategoryReturnsAmount")-rsEntry("CategoryReturnsTaxAmount"))%></span><% end if %><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;</td>
                                                                <td align="right"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><% if rsEntry("CategoryReturnsTaxAmount")<>0 then %><span style="color:red;"><%=FmtCurrency(rsEntry("CategoryReturnsTaxAmount"))%></span><% end if %><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;</td>
                                                                <td align="right"><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><% if rsEntry("CategoryReturnCount")<>0 then %><span style="color:red;"><%=rsEntry("CategoryReturnCount")%></span><% end if %>&nbsp;</td>
                                                                <td align="right" style="background-color:#F0F0F0;">
																						<% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %><i><% end if %><%=FmtCurrency(rsEntry("CategoryAmount"))%><% if rsEntry("CategoryID")=21 and NOT ss_IncludeTips then %></i><% end if %>&nbsp;
                                                                </td>
                                                        </tr>
                                        
                                <%                      rsEntry.MoveNext
                                                loop %>
                                                <tr><td colspan="15" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                <tr>
                                                        <td align="left"><b>Total Payments on Account:</b></td>
                                   						<%	if request.form("optIncludeCat2")="on" then %>
																		  <td align="left">&nbsp;</td>
																	<% end if %>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(AcctGCTotal+GCreturnsTotal-GCTaxTotal)%></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(GCTaxTotal)%></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=GCCountTotal%></strong>&nbsp;</td>
                                                        <td align="right"><strong><% if GCreturnsTotal-GCreturnsTaxTotal<>0 then %><span style="color:red;"><%=FmtCurrency(GCreturnsTotal-GCreturnsTaxTotal)%></span><% end if %></strong>&nbsp;</td>
                                                        <td align="right"><strong><% if GCreturnsTaxTotal<>0 then %><span style="color:red;"><%=FmtCurrency(GCreturnsTaxTotal)%></span><% end if %></strong>&nbsp;</td>
                                                        <td align="right"><strong><% if GCReturnCountTotal<>0 then %><span style="color:red;"><%=GCReturnCountTotal%></span><% end if %></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(AcctGCTotal)%></strong>&nbsp;</td>
																		<%	if false then 'request.form("optSubCat")<>"on" AND request.Form("optIncludeCat2")<>"on" then ' show unearned revenue %>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(0)%></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(AcctGCTotal)%></strong>&nbsp;</td>
																		<% end if %>
                                                </tr>           
                                                <tr><td colspan="15">&nbsp;</td></tr>

                                <%      end if
                                        rsEntry.close
                                        
                                        if paymentsExist OR GCpaymentsExist then
                                                %>
                                                <tr><td colspan="15" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                <tr>
                                                        <td align="left"><b>Total Payments:</b></td>
                                   						<%	if request.form("optIncludeCat2")="on" then %>
																		  <td align="left">&nbsp;</td>
																	<% end if %>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(categoryTotal+returnsTotal-taxTotal)%></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(taxTotal)%></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=countTotal%>&nbsp;</strong></td>
                                                        <td align="right"><strong><% if returnsTotal-returnsTaxTotal<>0 then %><span style="color:red;"><%=FmtCurrency(returnsTotal-returnsTaxTotal)%></span><% end if %></strong>&nbsp;</td>
                                                        <td align="right"><strong><% if returnsTaxTotal<>0 then %><span style="color:red;"><%=FmtCurrency(returnsTaxTotal)%></span><% end if %></strong>&nbsp;</td>
                                                        <td align="right"><% if returnCountTotal<>0 then %><span style="color:red;"><strong><%=returnCountTotal%></strong><% end if %>&nbsp;</span></td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(categoryTotal)%></strong>&nbsp;</td>
																		<%	if false then 'request.form("optSubCat")<>"on" AND request.Form("optIncludeCat2")<>"on" then ' show unearned revenue %>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(diffUERTotal)%></strong>&nbsp;</td>
                                                        <td align="right" style="background-color:#F0F0F0;"><strong><%=FmtCurrency(categoryTotal-diffUERTotal)%></strong>&nbsp;</td>
																		<% end if %>
                                                </tr>
                                        </table>
                                <%      end if %>
                                                                                </td>
                                                                        </tr>
                                                                        <tr>
                                                                                <td class="mainTextBig" valign="top" align="left">&nbsp;
<% ' MERCHANT ACCOUNT PROCESSING %>
<% if NOT Session("mvarMIDs")=0 then %>
<%                              dim ccTransExist, ACHExist
                                
                                strSQL = "SELECT tblCCTrans.ccType, ISNULL(AuthTbl.AuthTotalSwiped, 0) AS AuthTotalSwiped, ISNULL(AuthTbl.AuthTotalKeyed, 0) AS AuthTotalKeyed, ISNULL(SettledTbl.SettledTotalSwiped, 0) AS SettledTotalSwiped, ISNULL(SettledTbl.SettledTotalKeyed, 0) AS SettledTotalKeyed "
                                strSQL = strSQL & " FROM tblCCTrans LEFT OUTER JOIN "
                                strSQL = strSQL & " (SELECT ccType, SUM(CASE WHEN CCSwiped=1 then ccAmt ELSE 0 END) AS AuthTotalSwiped, SUM(CASE WHEN CCSwiped=0 then ccAmt ELSE 0 END) AS AuthTotalKeyed  FROM tblCCTrans "
                                strSQL = strSQL & " WHERE Settled = 0 AND Status = 'Approved' "
                                'MB bug#6184
                                if strLoc <> "" AND strLoc <> "0" then
                                    strSQL = strSQL & " AND (LocationID IN (" & strLoc & ")) "
                                end if
                                strSQL = strSQL & " AND TransTime < " & DateSep & dateadd("d", 1, closeDate) & DateSep & " AND TransTime >= " & DateSep & oldCloseDate & DateSep 
                                strSQL = strSQL & " GROUP BY ccType) AuthTbl ON AuthTbl.ccType = tblCCTrans.ccType LEFT OUTER JOIN "
                                strSQL = strSQL & " (SELECT ccType, SUM(CASE WHEN CCSwiped=1 AND tblCCTrans.Status='Credit' THEN ccAmt * -1 WHEN CCSwiped=1 THEN ccAmt ELSE 0 END) AS SettledTotalSwiped, SUM(CASE WHEN CCSwiped=0 AND tblCCTrans.Status='Credit' THEN ccAmt * -1 WHEN CCSwiped=0 THEN ccAmt ELSE 0 END) AS SettledTotalKeyed  FROM tblCCTrans "
                                strSQL = strSQL & " WHERE Settled = 1 AND TransTime < " & DateSep & dateadd("d", 1, closeDate) & DateSep & " AND TransTime >= " & DateSep & oldCloseDate & DateSep 
                                'MB bug#6184
                                if strLoc <> "" AND strLoc <> "0" then                                
                                    strSQL = strSQL & " AND (LocationID IN (" & strLoc & ")) "
                                end if
                                strSQL = strSQL & " GROUP BY ccType) SettledTbl ON SettledTbl.ccType = tblCCTrans.ccType "
                                'MB bug#6184
                                if strLoc <> "" AND strLoc <> "0" then
                                    strSQL = strSQL & " WHERE (tblCCTrans.LocationID IN (" & strLoc & ")) "
                                end if
                                strSQL = strSQL & " GROUP BY tblCCTrans.ccType, AuthTbl.AuthTotalSwiped, AuthTbl.AuthTotalKeyed, SettledTbl.SettledTotalSwiped, SettledTbl.SettledTotalKeyed "
                                strSQL = strSQL & " HAVING (NOT tblCCTrans.ccType IS NULL) "
                        
                               'response.write debugSQL(strSQL, "SQL")
                                rsEntry2.open strSQL, cnWS, 3
                                
                                if rsEntry2.recordCount > 0 then
                                        ccTransExist = true
                                end if
                        
                                strSQL = "SELECT ISNULL(AuthTbl.AuthTotal, 0) as AuthTotal, ISNULL(SettledTbl.SettledTotal, 0) as SettledTotal "
                                strSQL = strSQL & " FROM "
                                strSQL = strSQL & " (SELECT SUM(ccAmt) AS AuthTotal FROM tblCCTrans "
                                strSQL = strSQL & " WHERE Settled = 0 AND (NOT (tblCCTrans.ACHName IS NULL)) "
                                'MB bug#6184
                                if strLoc <> "" AND strLoc <> "0" then
                                    strSQL = strSQL & " AND (LocationID IN (" & strLoc & ")) "
                                end if
                                strSQL = strSQL & " AND TransTime < " & DateSep & dateadd("d", 1, closeDate) & DateSep & " AND TransTime >= " & DateSep & oldCloseDate & DateSep 
                                strSQL = strSQL & " ) AuthTbl, "
                                strSQL = strSQL & " (SELECT SUM(ccAmt) AS SettledTotal FROM tblCCTrans "
                                strSQL = strSQL & " WHERE Settled = 1 AND Status = 'Funded' AND TransTime < " & DateSep & dateadd("d", 1, closeDate) & DateSep & " AND TransTime >= " & DateSep & oldCloseDate & DateSep 
                                'MB bug#6184
                                if strLoc <> "" AND strLoc <> "0" then
                                    strSQL = strSQL & " AND (LocationID IN (" & strLoc & ")) "
                                end if
                                strSQL = strSQL & " ) SettledTbl "
                        
                               'response.write debugSQL(strSQL, "SQL")
                                rsEntry3.open strSQL, cnWS, 3
                                
                                if rsEntry3.recordCount > 0 then
                                        ACHExist = true
                                end if

                                authTotalSwiped = 0
                                authTotalKeyed = 0
                                settledTotalSwiped = 0
                                settledTotalKeyed = 0

                                if ccTransExist OR ACHExist then %>
                                        <table class="mainText" width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
                                                <tr><td colspan="15">&nbsp;</td></tr>
                                                <tr style="background-color:<%=session("pageColor4")%>;">
                                                        <td colspan="15" align="center" class="whiteHeader"><strong>Merchant Account Processing</strong></td>
                                                </tr>
                                                <tr>
                                                        <td colspan="15">
                                                                <table class="mainText" width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
                                                                        <tr>
                                                                                <td align="left" valign="bottom"><strong>Payment Type</strong></td>
                                                                                <td align="right" valign="bottom">&nbsp;<strong>Approved&nbsp;Swiped</strong></td>
                                                                                <td align="right" valign="bottom">&nbsp;<strong>Approved&nbsp;Keyed</strong></td>
                                                                                <td align="right" valign="bottom">&nbsp;<strong>Settled&nbsp;Swiped</strong></td>
                                                                                <td align="right" valign="bottom">&nbsp;<strong>Settled&nbsp;Keyed</strong>&nbsp;
                                                                                  <% if request.form("frmExpReport")<>"true" then %><img src="<%= contentUrl("/asp/adm/images/tech-tip-16px.png") %>" title="Settled Totals are only accurate for data beyond October, 2006."></td><% end if %>
                                                                        </tr>
                                
                        <%      end if
                                if ccTransExist then
                                        do WHILE NOT rsEntry2.EOF %>
                                                                        <tr>
                                                                                <td align="left"><%=rsEntry2("ccType")%></td>
                                                                                <td align="right">&nbsp;<% if rsEntry2("AuthTotalSwiped")<>0 then response.write FmtCurrency(rsEntry2("AuthTotalSwiped")/100) end if %></td>
                                                                                <td align="right">&nbsp;<% if rsEntry2("AuthTotalKeyed")<>0 then response.write FmtCurrency(rsEntry2("AuthTotalKeyed")/100) end if %></td>
                                                                                <td align="right">&nbsp;<% if rsEntry2("SettledTotalSwiped")<>0 then response.write FmtCurrency(rsEntry2("SettledTotalSwiped")/100) end if %></td>
                                                                                <td align="right">&nbsp;<% if rsEntry2("SettledTotalKeyed")<>0 then response.write FmtCurrency(rsEntry2("SettledTotalKeyed")/100) end if %></td>
                                                                        </tr>
                                <%              
                                                authTotalSwiped = authTotalSwiped + rsEntry2("AuthTotalSwiped")
                                                authTotalKeyed = authTotalKeyed + rsEntry2("AuthTotalKeyed")
                                                settledTotalSwiped = settledTotalSwiped + rsEntry2("SettledTotalSwiped")
                                                settledTotalKeyed = settledTotalKeyed + rsEntry2("SettledTotalKeyed")
                                                rsEntry2.MoveNext
                                        loop %>
                        <%      end if  ' ccTrans %>
                        <%      if ACHExist then %>
                                                                        <tr>
                                                                                <td align="left">ACH</td>
                                                                                <td align="right">&nbsp;</td>
                                                                                <td align="right">&nbsp;<% if authTotalKeyed<>0 then response.write FmtCurrency(rsEntry3("AuthTotal")/100) end if %></td>
                                                                                <td align="right">&nbsp;</td>
                                                                                <td align="right">&nbsp;<% if settledTotalKeyed<>0 then response.write FmtCurrency(rsEntry3("SettledTotal")/100) end if %></td>
                                                                        </tr>
                        <%              
                                        authTotalKeyed = authTotalKeyed + rsEntry3("AuthTotal")
                                        settledTotalKeyed = settledTotalKeyed + rsEntry3("SettledTotal")
                                end if %>
                        <%      if ccTransExist OR ACHExist then %>
                                                                        <tr><td colspan="10" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                                        <tr>
                                                                                <td align="left"><strong><%= getHotWord(22)%>:</strong></td>
                                                                                <td align="right"><strong><%=FmtCurrency(authTotalSwiped/100)%></strong></td>
                                                                                <td align="right"><strong><%=FmtCurrency(authTotalKeyed/100)%></strong></td>
                                                                                <td align="right"><strong><%=FmtCurrency(settledTotalSwiped/100)%></strong></td>
                                                                                <td align="right"><strong><%=FmtCurrency(settledTotalKeyed/100)%></strong></td>
                                                                        </tr>
                                                                </table>
                                                        </td>
                                                </tr>
                                        </table>
                        <%      end if %>
<%      end if %>
                                                                                </td>
                                                                        </tr>
                                                                        <tr>
                                                                                <td class="mainTextBig" valign="top" align="left">&nbsp;
        <%      if intCloseID=0 and request.form("optEmployeeID")="-1" then

                        ' BQL 45_2327 added discount computations to the query and to the output
                        ' BQL 49_2616 Added unpaid column to query/output.
                        strSQL = " SELECT tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup, ISNULL(VDSum.CompVisits, 0) AS CompVisits, ISNULL(VDSum.UnpaidVisits, 0) as UnpaidVisits, ISNULL(VDSum.CountVisits, 0) AS CountSeriesVisits, ISNULL(VDSum.TimeVisits, 0) "
                        strSQL = strSQL & " AS TimeSeriesVisits, ISNULL(VDSum.MemberVisits, 0) AS MemberSeriesVisits, ISNULL(CountRev.RevPerTG, 0) AS CountRev, ISNULL(CountRev.CountDiscPerTG, 0) AS CountDisc, "
                        strSQL = strSQL & " ISNULL(TimeRev.RevPerTG, 0) AS TimeRev, ISNULL(TimeRev.TimeDiscPerTG, 0) AS TimeDisc, ISNULL(MemberRev.RevPerTG, 0) AS MemberRev, ISNULL(MemberRev.MemberDiscPerTG, 0) AS MemberDisc, ISNULL(ExpRev.ExpirationRev, 0) AS ExpRev, ISNULL(ExpDiscPerTG, 0) AS ExpDisc "
                        strSQL = strSQL & " FROM tblTypeGroup LEFT OUTER JOIN "
                        strSQL = strSQL & " (SELECT SUM(CASE WHEN [PAYMENT DATA].Type = 1 AND [VISIT DATA].Value = 1 THEN 1 ELSE 0 END) AS CountVisits, SUM(CASE WHEN [PAYMENT DATA].Type = 9 THEN 1 ELSE 0 END) AS UnpaidVisits, SUM(CASE WHEN [PAYMENT DATA].Type = 2 AND "
                        strSQL = strSQL & " [VISIT DATA].Value = 1 THEN 1 ELSE 0 END) AS TimeVisits, SUM(CASE WHEN (tblSeriesType.IsSystem=0) AND [VISIT DATA].Value = 1 THEN 1 ELSE 0 END) "
                        strSQL = strSQL & " AS MemberVisits, SUM(CASE WHEN [VISIT DATA].Value = 0 THEN 1 ELSE 0 END) AS CompVisits, [VISIT DATA].TypeGroup "
                        strSQL = strSQL & " FROM [VISIT DATA] INNER JOIN "
                        strSQL = strSQL & " [PAYMENT DATA] ON [VISIT DATA].PmtRefNo = [PAYMENT DATA].PmtRefNo "
                        strSQL = strSQL & " INNER JOIN tblSeriesType ON [Payment Data].Type = tblSeriesType.SeriesTypeID "
                        strSQL = strSQL & " WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") "
                        'BN Bug#3374 Don't display earned revenue for returned series
                        strSQL = strSQL & " AND ([PAYMENT DATA].Returned = 0) "
                        ' BQL 11/5/2008 - added to filter out XR payment rows for visits at another location
                        strSQL = strSQL & " AND [VISIT DATA].VisitType<>-1 "
                        'MB bug#6184
                                if strLoc <> "" AND strLoc <> "0" then
                            strSQL = strSQL & " AND ([VISIT DATA].Location IN (" & strLoc & ")) "
                        end if
                        strSQL = strSQL & " GROUP BY [VISIT DATA].TypeGroup) VDSum ON  "
                        strSQL = strSQL & " VDSum.TypeGroup = tblTypeGroup.TypeGroupID LEFT OUTER JOIN "
                        strSQL = strSQL & " (SELECT SUM(((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0))  "
                        strSQL = strSQL & " / [PAYMENT DATA].NumClasses) * [VISIT DATA].NumDeducted) AS RevPerTG, SUM(ISNULL([Sales Details].DiscAmt, 0)/ ([PAYMENT DATA].NumClasses) * [VISIT DATA].NumDeducted) as CountDiscPerTG, [PAYMENT DATA].TypeGroup "
                        strSQL = strSQL & " FROM [VISIT DATA] INNER JOIN "
                        strSQL = strSQL & " [PAYMENT DATA] ON [VISIT DATA].PmtRefNo = [PAYMENT DATA].PmtRefNo INNER JOIN "
                        strSQL = strSQL & " [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo "
                        strSQL = strSQL & " WHERE ([VISIT DATA].ClassDate >= " & DateSep & cSDate & DateSep & ") AND ([VISIT DATA].ClassDate <= " & DateSep & cEDate & DateSep & ") AND ([PAYMENT DATA].Type = 1) AND ([VISIT DATA].NumDeducted > 0) AND ([PAYMENT DATA].NumClasses > 0)"
                        'BN Bug#3374 Don't display earned revenue for returned series
                        strSQL = strSQL & " AND ([PAYMENT DATA].Returned = 0) "
                        ' BQL 11/5/2008 - added to filter out XR payment rows for visits at another location
                        strSQL = strSQL & " AND [VISIT DATA].VisitType<>-1 "
                        'MB bug#6184
                        if strLoc <> "" AND strLoc <> "0" then
                            strSQL = strSQL & " AND ([VISIT DATA].Location IN (" & strLoc & ")) "
                        end if
                        strSQL = strSQL & " GROUP BY [PAYMENT DATA].TypeGroup) CountRev ON CountRev.TypeGroup = tblTypeGroup.TypeGroupID LEFT OUTER JOIN "
                        strSQL = strSQL & " (SELECT SUM((((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, "
                        strSQL = strSQL & " CASE WHEN [PAYMENT DATA].ActiveDate < " & DateSep & cSDate & DateSep & " THEN " & DateSep & cSDate & DateSep & " ELSE [PAYMENT DATA].ActiveDate END,  "
                        strSQL = strSQL & " CASE WHEN [PAYMENT DATA].ExpDate > " & DateSep & cEDate & DateSep & " THEN " & DateSep & cEDate & DateSep & " ELSE [PAYMENT DATA].ExpDate END) + 1)) /  (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) AS RevPerTG,  "
                        strSQL = strSQL & " SUM((((ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, "
                        strSQL = strSQL & " CASE WHEN [PAYMENT DATA].ActiveDate < " & DateSep & cSDate & DateSep & " THEN " & DateSep & cSDate & DateSep & " ELSE [PAYMENT DATA].ActiveDate END,  "
                        strSQL = strSQL & " CASE WHEN [PAYMENT DATA].ExpDate > " & DateSep & cEDate & DateSep & " THEN " & DateSep & cEDate & DateSep & " ELSE [PAYMENT DATA].ExpDate END) + 1)) /  (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) as TimeDiscPerTG, [PAYMENT DATA].TypeGroup "
                        strSQL = strSQL & " FROM [PAYMENT DATA] INNER JOIN "
                        strSQL = strSQL & " [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo "
                        strSQL = strSQL & " WHERE ([PAYMENT DATA].ActiveDate <= " & DateSep & cEDate & DateSep & ") AND ([PAYMENT DATA].ExpDate >= " & DateSep & cSDate & DateSep & ") AND ([PAYMENT DATA].Type = 2) "
                        'BN Bug#3374 Don't display earned revenue for returned series
                        strSQL = strSQL & " AND ([PAYMENT DATA].Returned = 0) "
                        'MB bug#6184
						'if request.Form("optLocation") <> "" AND request.Form("optLocation") <> "0" then
                        if strLoc <> "" AND strLoc <> "0" then
                            strSQL = strSQL & " AND ([Sales Details].Location IN (" & strLoc & ")) "
                        end if
                        strSQL = strSQL & " GROUP BY [PAYMENT DATA].TypeGroup) TimeRev ON TimeRev.TypeGroup = tblTypeGroup.TypeGroupID LEFT OUTER JOIN "
                        
                        strSQL = strSQL & " (SELECT SUM((((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, "
                        strSQL = strSQL & " CASE WHEN [PAYMENT DATA].ActiveDate < " & DateSep & cSDate & DateSep & " THEN " & DateSep & cSDate & DateSep & " ELSE [PAYMENT DATA].ActiveDate END,  "
                        strSQL = strSQL & " CASE WHEN [PAYMENT DATA].ExpDate > " & DateSep & cEDate & DateSep & " THEN " & DateSep & cEDate & DateSep & " ELSE [PAYMENT DATA].ExpDate END) + 1)) /  (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) AS RevPerTG,  "
                        strSQL = strSQL & " SUM((((ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, "
                        strSQL = strSQL & " CASE WHEN [PAYMENT DATA].ActiveDate < " & DateSep & cSDate & DateSep & " THEN " & DateSep & cSDate & DateSep & " ELSE [PAYMENT DATA].ActiveDate END,  "
                        strSQL = strSQL & " CASE WHEN [PAYMENT DATA].ExpDate > " & DateSep & cEDate & DateSep & " THEN " & DateSep & cEDate & DateSep & " ELSE [PAYMENT DATA].ExpDate END) + 1)) /  (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) as MemberDiscPerTG, [PAYMENT DATA].TypeGroup "
                        strSQL = strSQL & " FROM [PAYMENT DATA] INNER JOIN "
                        strSQL = strSQL & " [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo " &_
                                          " INNER JOIN tblSeriesType ON [Payment Data].Type = tblSeriesType.SeriesTypeID "
                        strSQL = strSQL & " WHERE ([PAYMENT DATA].ActiveDate <= " & DateSep & cEDate & DateSep & ") AND ([PAYMENT DATA].ExpDate >= " & DateSep & cSDate & DateSep & ") " &_
                                          " AND (tblSeriesType.isSystem = 0) "
                        'BN Bug#3374 Don't display earned revenue for returned series
                        strSQL = strSQL & " AND ([PAYMENT DATA].Returned = 0) "
                        'MB bug#6184
						'if request.Form("optLocation") <> "" AND request.Form("optLocation") <> "0" then
                        if strLoc <> "" AND strLoc <> "0" then
                            strSQL = strSQL & " AND ([Sales Details].Location IN (" & strLoc & ")) "
                        end if
                        strSQL = strSQL & " GROUP BY [PAYMENT DATA].TypeGroup) MemberRev ON MemberRev.TypeGroup = tblTypeGroup.TypeGroupID LEFT OUTER JOIN "
                        
                        strSQL = strSQL & " (SELECT SUM((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0)) / [PAYMENT DATA].NumClasses * [PAYMENT DATA].Remaining) AS ExpirationRev, "
                        strSQL = strSQL & " SUM((ISNULL([Sales Details].DiscAmt, 0)) / [PAYMENT DATA].NumClasses * [PAYMENT DATA].Remaining) as ExpDiscPerTG, [PAYMENT DATA].TypeGroup "
                        strSQL = strSQL & " FROM [PAYMENT DATA] INNER JOIN "
                        strSQL = strSQL & " [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo " &_
                                          " INNER JOIN Products on Products.ProductId = [Payment Data].ProductId "
                        strSQL = strSQL & " WHERE ([PAYMENT DATA].ExpDate >= " & DateSep & cSDate & DateSep & ") AND ([PAYMENT DATA].ExpDate <= " & DateSep & cEDate & DateSep & ") AND ([PAYMENT DATA].Type = 1) AND "
                        strSQL = strSQL & " ([PAYMENT DATA].Remaining > 0)  "
                        'BN Bug#3374 Don't display earned revenue for returned series
                        strSQL = strSQL & " AND ([PAYMENT DATA].Returned = 0) " &_
                                          " AND (Products.ItemTypeID = 1) " ' 
                        'MB bug#6184
						'if request.Form("optLocation") <> "" AND request.Form("optLocation") <> "0" then
                        if strLoc <> "" AND strLoc <> "0" then
                            strSQL = strSQL & " AND ([Sales Details].Location IN (" & strLoc & ")) "
                        end if
                        strSQL = strSQL & " GROUP BY [PAYMENT DATA].TypeGroup) ExpRev ON ExpRev.TypeGroup = tblTypeGroup.TypeGroupID "
                        
                        strSQL = strSQL & " WHERE (tblTypeGroup.Active = 1) " 

                       response.write debugSQL(strSQL, "Monster")
                        rsEntry.open strSQL, cnWS, 3
                        
                        dim memberVisits, compVisits, timeVisits, countVisits, unpaidVisits, totalVisits, timeRev, countRev, expRev, totalRev, printFirst, discountTotal, timeDisc, countDisc, expDisc, totalDisc
                        memberVisits = 0
                        compVisits = 0
                        timeVisits = 0
                        countVisits = 0
                        unpaidVisits = 0
                        totalVisits = 0
                        timeRev = 0
                        countRev = 0
                        expRev = 0
                        memRev = 0
                        totalRev = 0
                        timeDisc = 0
                        countDisc = 0
                        expDisc = 0
                        memDisc = 0
                        totalDisc = 0
                        printFirst = false
                        
                        setRowColors "#F0F0F0", "#FFFFFF"
                        
                        if NOT rsEntry.EOF then %>
                <%              do while NOT rsEntry.EOF
                                        if rsEntry("MemberSeriesVisits")<>0 OR rsEntry("CompVisits")<>0 OR rsEntry("TimeSeriesVisits")<>0 OR rsEntry("CountSeriesVisits")<>0 then
                                                if NOT printFirst then ' first row hasn't been printed yet %>
                                        <table class="mainText" width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
                                                <tr><td colspan="18">&nbsp;</td></tr>
                                                <tr style="background-color:<%=session("pageColor4")%>;">
                                                        <td colspan="18" align="center" class="whiteHeader"><strong>Service Program Performance</strong></td>
                                                </tr>
                                                <tr><td colspan="18">&nbsp;</td></tr>
                                                <tr>
                                                        <td colspan="18" align="center" class="mainText"><strong>Attendance</strong></td>
                                                </tr>
                                                <tr>
                                                        <td align="left" colspan="2" width="25%"><b>Program</b></td>
                                                        <td align="right" colspan="2"><b>Membership</b></td>
                                                        <td align="right" colspan="2"><b>Comps</b></td>
                                                        <td align="right" colspan="2"><b>Unpaids</b></td>
                                                        <td align="right" colspan="2"><b>Unlimited Visit Pricing Option</b></td>
                                                        <td align="right" colspan="3"><b>Limited Visit Pricing Option</b></td>
                                                        <td align="right" colspan="3"><b>Total Attendance</b></td>
                                                </tr>
                                        <%              printFirst = true
                                                end if %>
                                                
                                                <tr style="background-color:<%=getRowColor(true)%>;">
                                                        <td align="left" colspan="2"><%=rsEntry("TypeGroup")%></td>
                                                        <td align="right" colspan="2"><% if rsEntry("MemberSeriesVisits")<>0 then %><%=rsEntry("MemberSeriesVisits")%><% end if %></td>
                                                        <td align="right" colspan="2"><% if rsEntry("CompVisits")<>0 then %><%=rsEntry("CompVisits")%><% end if %></td>
                                                        <td align="right" colspan="2"><% if rsEntry("UnpaidVisits")<>0 then %><%=rsEntry("UnpaidVisits")%><% end if %></td>
                                                        <td align="right" colspan="2"><% if rsEntry("TimeSeriesVisits")<>0 then %><%=rsEntry("TimeSeriesVisits")%><% end if %></td>
                                                        <td align="right" colspan="3"><% if rsEntry("CountSeriesVisits")<>0 then %><%=rsEntry("CountSeriesVisits")%><% end if %></td>
                                                        <td align="right" colspan="3"><% if (rsEntry("MemberSeriesVisits")+rsEntry("CompVisits")+rsEntry("UnpaidVisits")+rsEntry("TimeSeriesVisits")+rsEntry("CountSeriesVisits"))<>0 then %><%=(rsEntry("MemberSeriesVisits")+rsEntry("CompVisits")+rsEntry("TimeSeriesVisits")+rsEntry("UnpaidVisits")+rsEntry("CountSeriesVisits"))%><% end if %></td>
                                                </tr>
                                                <%              memberVisits = memberVisits + rsEntry("MemberSeriesVisits")
                                                                compVisits = compVisits + rsEntry("CompVisits")
                                                                timeVisits = timeVisits + rsEntry("TimeSeriesVisits")
                                                                countVisits = countVisits + rsEntry("CountSeriesVisits")
                                                                unpaidVisits = unpaidVisits + rsEntry("UnpaidVisits")
                                                                totalVisits = totalVisits + rsEntry("MemberSeriesVisits")+rsEntry("CompVisits")+rsEntry("UnpaidVisits")+rsEntry("TimeSeriesVisits")+rsEntry("CountSeriesVisits")
                                                                
                                        end if
                                        rsEntry.MoveNext
                                loop
                                if printFirst then      %>
                                                <tr><td colspan="18" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                <tr>
                                                        <td align="left" colspan="2"><strong><%= getHotWord(22)%>:</strong></td>
                                                        <td align="right" colspan="2"><strong><%=memberVisits%></strong></td>
                                                        <td align="right" colspan="2"><strong><%=compVisits%></strong></td>
                                                        <td align="right" colspan="2"><strong><%=unpaidVisits%></strong></td>
                                                        <td align="right" colspan="2"><strong><%=timeVisits%></strong></td>
                                                        <td align="right" colspan="3"><strong><%=countVisits%></strong></td>
                                                        <td align="right" colspan="3"><strong><%=totalVisits%></strong></td>
                                                </tr>
                        <%      end if 
                                printFirst = false
                                rsEntry.MoveFirst
                                
                                setRowColors "#F0F0F0", "#FFFFFF"

                                do while NOT rsEntry.EOF
                                        if rsEntry("CountRev")<>0 OR rsEntry("TimeRev")<>0 OR rsEntry("ExpRev")<>0 OR rsEntry("MemberRev")<>0 then 
                                                if NOT printFirst then %>
                                        </table>
                                        <br />
                                        <table class="mainText" width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
                                                <tr><td colspan="18">&nbsp;</td></tr>
                                                <tr>
                                                        <td colspan="18" align="center" class="mainText"><strong>Earned Revenue</strong></td>
                                                </tr>
                                                <tr>
                                                        <td align="left" width="25%"><b>&nbsp;</b></td>
                                                        <td>&nbsp;</td>
                                                        <td align="right"><b>Unlimited Visit Pricing Option</b></td>
                                                        <td align="right" ><b>Memberships</b></td>
                                                        <td align="right"><b>Limited Visit Pricing Option</b></td>
                                                        <td align="right"><b>Expired Series</b></td>
                                                        <td align="right"><b>Total</b></td>
                                                </tr>
                                                <!--tr>
                                                        <td align="left"><b>Program</b></td>
                                                        <td>&nbsp;</td>
                                                        <td align="right"><i>Subtotal</i></td>
                                                        <td align="right"><i>Discount</i></td>
                                                        <td align="right"><i>Earned</i></td>
                                                        <td align="right"><i>Subtotal</i></td>
                                                        <td align="right"><i>Discount</i></td>
                                                        <td align="right"><i>Earned</i></td>
                                                        <td align="right"><i>Subtotal</i></td>
                                                        <td align="right"><i>Discount</i></td>
                                                        <td align="right"><i>Earned</i></td>
                                                        <td align="right"><i>Subtotal</i></td>
                                                        <td align="right"><i>Discount</i></td>
                                                        <td align="right"><i>Earned</i></td>
                                                </tr-->
                                        <%              printFirst = true
                                                end if %>
                                                <tr style="background-color:<%=getRowColor(true)%>;">
                                                        <td align="left"><%=rsEntry("TypeGroup")%></td>
                                                        <td>&nbsp;</td>
                                                        <!--td align="right">&nbsp;<% if rsEntry("TimeDisc")<>0 then %><%=FmtCurrency(rsEntry("TimeRev")+rsEntry("TimeDisc"))%><% end if %></td-->
                                                        <!--td align="right">&nbsp;<% if rsEntry("TimeDisc")<>0 then %><%=FmtCurrency(rsEntry("TimeDisc"))%><% end if %></td-->
                                                        <td align="right">&nbsp;<% if rsEntry("TimeRev")<>0 OR rsEntry("TimeDisc")<>0 then %><%=FmtCurrency(rsEntry("TimeRev"))%><% end if %></td>
                                                        <!--td align="right">&nbsp;<% if rsEntry("MemberDisc")<>0 then %><%=FmtCurrency(rsEntry("MemberRev")+rsEntry("MemberDisc"))%><% end if %></td-->
                                                        <!--td align="right">&nbsp;<% if rsEntry("MemberDisc")<>0 then %><%=FmtCurrency(rsEntry("MemberDisc"))%><% end if %></td-->
                                                        <td align="right">&nbsp;<% if rsEntry("MemberRev")<>0 OR rsEntry("MemberDisc")<>0 then %><%=FmtCurrency(rsEntry("MemberRev"))%><% end if %></td>
                                                        <!--td align="right">&nbsp;<% if rsEntry("CountDisc")<>0 then %><%=FmtCurrency(rsEntry("CountRev")+rsEntry("CountDisc"))%><% end if %></td-->
                                                        <!--td align="right">&nbsp;<% if rsEntry("CountDisc")<>0 then %><%=FmtCurrency(rsEntry("CountDisc"))%><% end if %></td-->
                                                        <td align="right">&nbsp;<% if rsEntry("CountRev")<>0 OR rsEntry("CountDisc")<>0 then %><%=FmtCurrency(rsEntry("CountRev"))%><% end if %></td>
                                                        <!--td align="right">&nbsp<% if rsEntry("ExpDisc")<>0 then %><%=FmtCurrency(rsEntry("ExpRev")+rsEntry("ExpDisc"))%><% end if %>;</td-->
                                                        <!--td align="right">&nbsp;<% if rsEntry("ExpDisc")<>0 then %><%=FmtCurrency(rsEntry("ExpDisc"))%><% end if %></td-->
                                                        <td align="right">&nbsp;<% if rsEntry("ExpRev")<>0 OR rsEntry("ExpDisc")<>0 then %><%=FmtCurrency(rsEntry("ExpRev"))%><% end if %></td>
                                                        <!--td align="right">&nbsp;<% if rsEntry("CountDisc")+rsEntry("TimeDisc")+rsEntry("ExpDisc")<>0 then %><%=FmtCurrency(rsEntry("CountDisc")+rsEntry("TimeDisc")+rsEntry("ExpDisc")+rsEntry("CountRev")+rsEntry("TimeRev")+rsEntry("ExpRev"))%><% end if %></td-->
                                                        <!--td align="right">&nbsp;<% if rsEntry("CountDisc")+rsEntry("TimeDisc")+rsEntry("ExpDisc")<>0 then %><%=FmtCurrency(rsEntry("CountDisc")+rsEntry("TimeDisc")+rsEntry("ExpDisc"))%><% end if %></td-->
                                                        <td align="right">&nbsp;<% if rsEntry("CountRev")+rsEntry("TimeRev")+rsEntry("ExpRev")+rsEntry("MemberRev")<>0 then %><%=FmtCurrency(rsEntry("CountRev")+rsEntry("TimeRev")+rsEntry("ExpRev")+rsEntry("MemberRev"))%><% end if %></td>
                                                </tr>
                        <%              end if

                                        timeRev = timeRev + rsEntry("TimeRev")
                                        countRev = countRev + rsEntry("CountRev")
                                        expRev = expRev + rsEntry("ExpRev")
                                        memRev = memRev + rsEntry("MemberRev")
                                        totalRev = totalRev + rsEntry("CountRev")+rsEntry("TimeRev")+rsEntry("ExpRev") +rsEntry("MemberRev")
                                        timeDisc = timeDisc + rsEntry("TimeDisc")
                                        countDisc = countDisc + rsEntry("CountDisc")
                                        expDisc = expDisc + rsEntry("ExpDisc")
                                        memDisc = memDisc + rsEntry("MemberDisc")
                                        totalDisc = totalDisc + rsEntry("CountDisc")+rsEntry("TimeDisc")+rsEntry("ExpDisc") +rsEntry("MemberDisc")

                                        rsEntry.MoveNext
                                loop
                                if printFirst then %>
                                                <tr><td colspan="18" style="height: 1px; line-height: 1px; font-size: 1px;background-color:<%=session("pageColor4")%>;"><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1"></td></tr>
                                                <tr>
                                                        <td align="left"><strong><%= getHotWord(22)%>:</strong></td>
                                                        <td align="right">&nbsp;</td>
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(timeRev+timeDisc)%></td-->
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(timeDisc)%></td-->
                                                        <td align="right">&nbsp;<strong><%=FmtCurrency(timeRev)%></strong></td>
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(memRev+memDisc)%></td-->
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(memDisc)%></td-->
                                                        <td align="right">&nbsp;<strong><%=FmtCurrency(memRev)%></strong></td>
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(countRev+countDisc)%></td-->
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(countDisc)%></td-->
                                                        <td align="right">&nbsp;<strong><%=FmtCurrency(countRev)%></strong></td>
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(expRev+expDisc)%></td-->
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(expDisc)%></td-->
                                                        <td align="right">&nbsp;<strong><%=FmtCurrency(expRev)%></strong></td>
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(totalRev+totalDisc)%></td-->
                                                        <!--td align="right">&nbsp;<%=FmtCurrency(totalDisc)%></td-->
                                                        <td align="right">&nbsp;<strong><%=FmtCurrency(totalRev)%></strong></td>
                                                </tr>
                                        </table>
                        <%      end if
                        end if
                end if ' intCloseID = 0 %>
                                                                                </td>
                                                                        </tr>
                                                                </table>
                                                        </td>
                                                </tr>
<%      end if 'frmGenReport %>
                                                <tr>
                                                        <td colspan="2">&nbsp;</td>
                                                        <td>
                                                                <table class="mainText" width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
                                                                        <tr><td colspan="3">&nbsp;</td></tr>
                                                                        <tr><td colspan="3">&nbsp;</td></tr>
                                                                        <tr><td colspan="3">&nbsp;</td></tr>
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
%>
