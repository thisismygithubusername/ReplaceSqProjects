<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<%


'dim SessionFarm : set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
<%
    dim rsEntry : set rsEntry = Server.CreateObject("ADODB.Recordset")
    dim rsEntry2 : set rsEntry2 = Server.CreateObject("ADODB.Recordset")
    %>
    <!-- #include file="inc_accpriv.asp" -->
    <!-- #include file="inc_rpt_tagging.asp" -->
    <!-- #include file="inc_utilities.asp" -->
    <!-- #include file="inc_rpt_save.asp" -->
    <% dim doRefresh : doRefresh = false %>
    <!-- #include file="inc_date_arrows.asp" -->
    <!-- #include file="../inc_ajax.asp" -->
    <!-- #include file="../inc_val_date.asp" -->
    <%
    if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_ANALYSIS") then
        Response.Write "<script type=""text/javascript"">alert('You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.');javascript:history.go(-1);</script>"
    else
        %>
        <!-- #include file="../inc_i18n.asp" -->
        <!-- #include file="inc_hotword.asp" -->
        <%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

        dim showDetails, CSDate, CEDate, tmpCurDate, cLoc, rsPaid, DurationMonth, MonthRem, DurationDay, DayRem
        dim rowColor, tmpDefAmt, TotDefAmtCount, TotDefAmtTime, ap_view_all_locs, DaysUsed, EarnedAmount
        dim program, timeSeriesSubtotal, timeSeriesEarned, timeSeriesDiscount, countSeriesEarned, countSeriesDiscount
        dim countSeriesSubtotal, expiredSeriesEarned, expiredSeriesDiscount, expiredSeriesSubtotal, totalEarned, totalDiscount, totalSubtotal
        dim memberSeriesSubtotal, memberSeriesEarned, memberSeriesDiscount
        dim GrandTotal, TotEarned, TotEarnedTime
        dim remaining, used, earned, paymentAmtSum, deferredAmt, strOrder

        dim memberVisits : memberVisits = 0
        dim compVisits   : compVisits = 0
        dim timeVisits   : timeVisits = 0
        dim countVisits  : countVisits = 0
        dim unpaidVisits : unpaidVisits = 0
        dim totalVisits  : totalVisits = 0
        dim timeRev      : timeRev = 0
        dim countRev     : countRev = 0
        dim expRev       : expRev = 0
        dim memRev       : memRev = 0
        dim totalRev     : totalRev = 0
        dim timeDisc     : timeDisc = 0
        dim countDisc    : countDisc = 0
        dim expDisc      : expDisc = 0
        dim memDisc      : memDisc = 0
        dim totalDisc    : totalDisc = 0

        GrandTotal = 0
        tmpDefAmt=0
        TotDefAmtCount=0
        TotDefAmtTime=0

        ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
	
		dim category : category = ""
		if (RQ("category"))<>"" then
			category = RQ("category")
		elseif (RF("category"))<>"" then
			category = RF("category") 
		end if

        if request.form("requiredtxtDateStart")<>"" then
            Call SetLocale(session("mvarLocaleStr"))
            CSDate = CDATE(sqlInjectStr(request.form("requiredtxtDateStart")))
            Call SetLocale("en-us")
        else
            CSDate = DateValue(CDATE(DateAdd("m", -1, tzNow())))
        end if

        if request.form("requiredtxtDateEnd")<>"" then
            Call SetLocale(session("mvarLocaleStr"))
                CEDate = CDATE(sqlInjectStr(request.form("requiredtxtDateEnd")))
            Call SetLocale("en-us")
        else
            CEDate = DateValue(tzNow())
        end if

        If request.form("optSaleLoc")<>"" then
            cLoc = CINT(sqlInjectStr(request.form("optSaleLoc")))
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

        showDetails = true

        ' Hotwords
        'dim arrHW : arrHW = getHotWords(array(8,22,25,26,32,34,35,37,40,57,61,64,149,159,195,196))
            dim location_hw       : location_hw       = xssStr(allHotWords(8))
            dim total_hw          : total_hw          = xssStr(allHotWords(22))
            dim onlineStore_hw    : onlineStore_hw    = xssStr(allHotWords(25))
            dim currentSeries_hw  : currentSeries_hw  = xssStr(allHotWords(26))
            dim remaining_hw      : remaining_hw      = xssStr(allHotWords(32))
            dim purchased_hw      : purchased_hw      = xssStr(allHotWords(34))
            dim amount_hw         : amount_hw         = xssStr(allHotWords(35))
            dim paymentRefNum_hw  : paymentRefNum_hw  = xssStr(allHotWords(37))
            dim name_hw           : name_hw           = xssStr(allHotWords(40))
            dim date_hw           : date_hw           = xssStr(allHotWords(57))
            dim series_hw         : series_hw         = xssStr(allHotWords(61))
            dim expirationDate_hw : expirationDate_hw = xssStr(allHotWords(64))
            dim all_hw            : all_hw            = xssStr(allHotWords(149))
            dim view_hw           : view_hw           = xssStr(allHotWords(159))
            dim memberships_hw    : memberships_hw    = xssStr(allHotWords(195))
            dim and_hw            : and_hw            = xssStr(allHotWords(196))

        %>
<% if NOT request.form("frmExpReport")="true" then %>
<!-- #include file="pre.asp" -->
            <!-- #include file="frame_bottom.asp" -->
            <!-- #include file="../inc_date_ctrl.asp" -->
            <!-- #include file="inc_help_content.asp" -->
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_earned_revenue", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %>
            <script type="text/javascript">
              function exportReport() {
                  document.frmSales.frmExpReport.value = "true";
                  document.frmSales.frmGenReport.value = "true";
                  <% iframeSubmit "frmSales", "adm_rpt_earned_revenue.asp" %>
              }

            </script>
<% end if %>
            <!-- #include file="css/report.asp" -->
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
			<%= DisplayPhrase(reportPageTitlesDictionary,"Earnedrevenue") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
		  	</div>
		<%end if %>
            <div id="container">
		  <% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div id="head" class="headText" style="position:relative;left:2%;">
                <%=pp_pageTitle("Earned revenue")%>
			</div>
		<%end if %>
              <div id="options" class="mainText center">
                <form name="frmSales" action="adm_rpt_earned_revenue.asp" method="POST">
                  <input type="hidden" name="frmGenReport" value="" />
                  <input type="hidden" name="frmExpReport" value="" />
				 <% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
					 <input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
					 <input type="hidden" name="category" value="<%=category%>">
				 <% end if %>

                  <label for="requiredtxtDateStart">
                      Earned between:
                      <input type="text"  id="requiredtxtDateStart" name="requiredtxtDateStart" value="<%=FmtDateShort(CSDate)%>" class="date">
                   <script type="text/javascript">
                    var cal1 = new tcal({'formname':'frmSales', 'controlname':'requiredtxtDateStart'});
                    cal1.a_tpl.yearscroll = true;
                    </script>
                  </label>

                  <label for="requiredtxtDateEnd">
                      and:
                      <input type="text"  name="requiredtxtDateEnd" id="requiredtxtDateEnd" value="<%=FmtDateShort(CEDate)%>" class="date">
                   <script type="text/javascript">
                    var cal2 = new tcal({'formname':'frmSales', 'controlname':'requiredtxtDateEnd'});
                    cal2.a_tpl.yearscroll = true;
                    </script>
                   </label>
                   <label for="optSaleLoc">
                      <%=series_hw%> Purchased at:
                      <select id="optSaleLoc" name="optSaleLoc" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if%>>
                        <option value="0" <% if cLoc=0 then response.write "selected" end if %>><%=view_hw%></option>
                        <option value="98" <% if cLoc=98 then response.write "selected" end if %>><%=onlineStore_hw%></option>
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
                      <script type="text/javascript">document.frmSales.optSaleLoc.options[0].text = '<%=all_hw%>' +" " + '<%=location_hw%>' + "s";</script>
                  </label>

                  
                  
                  <!--
                  <label for="optIntegrated">
                    Only Integrated Payment Methods:
                    <input type="checkbox" id="optIntegrated" name="optIntegrated" <%if request.form("optIntegrated")="on" then response.write " checked" end if%> />
                  </label>
                  -->

                  <label for="optSortBy">
                    Sort By:
                    <select name="optSortBy">
                      <option value="0" <% if request.form("optSortBy")="" or request.form("optSortBy")="0" then %>selected<% end if %>>Payment Date</option>
                      <option value="1" <% if request.form("optSortBy")="1" then %>selected<% end if %>><%=expirationDate_hw%></option>
                      <option value="2" <% if request.form("optSortBy")="2" then %>selected<% end if %>><%=session("ClientHW")%>&nbsp;<%=name_hw%></option>
                      <option value="3" <% if request.form("optSortBy")="3" then %>selected<% end if %>>Earned Revenue Amount</option>
                    </select>
                  </label>

                  <% taggingFilter %>

                  <label for="optTG">
                    <%
                    strSQL = "SELECT DISTINCT tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup "
                    strSQL = strSQL & "FROM tblTypeGroup WHERE Active = 1 "
                    strSQL = strSQL & "ORDER BY tblTypeGroup.TypeGroup "
                    rsEntry.CursorLocation = 3
                    rsEntry.open strSQL, cnWS
                    set rsEntry.ActiveConnection = Nothing
                    %>
                    <select id="optTG" name="optTG">
                      <option value="0">All Programs</option>
                      <%
                      do while NOT rsEntry.EOF
                          %>
                          <option value="<%=rsEntry("TypeGroupID")%>" <% if request.form("optTG")=CSTR(rsEntry("TypeGroupID")) then response.write " selected" end if %>><%=rsEntry("TypeGroup")%></option>
                          <%
                          rsEntry.moveNext
                      loop
                      rsEntry.close
                      %>
                    </select>
                  </label>

                  <label for="optSummary">
                    View:
                    <select id="optSummary" name="optSummary">
                      <option value="summary" <% If request.Form("optSummary")="summary" Then Response.Write "selected" End If %>>Summary</option>
                      <option value="detail" <% If request.Form("optSummary")="detail" Then Response.Write "selected" End If %>>Detail</option>
                    </select>
                  </label>

                  <label for="optSeriesType"><%=series_hw%> Type:
                    <select id="optSeriesType" name="optSeriesType" disabled="disabled"><%
                    dim seriesTypeFilter
                    seriesTypeFilter = request.Form("optSeriesType")
                     %>
                      <option value=""><%=all_hw%>&nbsp;<%=xssStr(allHotWords(61))%></option>
                      <option value="1" <%if seriesTypeFilter="1" then response.write "selected" end if%>>Limited visit pricing option only</option>
                      <option value="2" <%if seriesTypeFilter="2" then response.write "selected" end if%>>Unlimited visit pricing options only</option>
                      <option value="-3" <%if seriesTypeFilter="-3" then response.write "selected" end if%>>Membership series only</option>
                      <%
                      strSQL = "SELECT SeriesTypeID, SeriesTypeName FROM tblSeriesType WHERE (IsSystem=0) AND (Active=1) ORDER BY SortOrder, SeriesTypeName "
                      rsEntry.CursorLocation = 3
                      rsEntry.open strSQL, cnWS
                      Set rsEntry.ActiveConnection = Nothing

                      do while NOT rsEntry.EOF
                          %>
                          <option value="<%=rsEntry("SeriesTypeID")%>" <% if seriesTypeFilter=CSTR(rsEntry("SeriesTypeID")) then response.write "selected" end if %>><%=rsEntry("SeriesTypeName")%></option>
                          <%
                          rsEntry.moveNext
                      loop
                      rsEntry.close

                      %>
                    </select>
                  </label>

                  <div class="center-ch">
                    <% showDateArrows("frmSales") %>
                  </div>

<div style ="text-align :center ">
<input type="button" name="Button"      value="Generate" onClick="report.generate();" />

                  <% if session("Pass") AND session("Admin")<>"false" AND validAccessPriv("RPT_EXPORT") then %>
                      <% exportToExcelButton %>
                  <% end if %>

                  <% if Session("Pass") AND Session("Admin")<>"false" AND validAccessPriv("RPT_TAG") then %>
                      <% taggingButtons("frmSales") %>
                  <% end if %>

                  <% savingButtons "frmSales", "Earned Revenue" %>

</div>

                  
                </form>
              </div>
<% end if 
if request.form("frmTagClients")="true" then
  strSQL = generateEarnedSummarySQL()
  response.write debugSQL(strSQL, "EarnedSummary")
  if request.form("frmTagClientsNew")="true" then
	  clearAndTagQuery(strSQL)
  else
	  tagQuery(strSQL)
	end if
else
%>
              <div id="report" class="mainText">
                <%
                if request.form("frmGenReport")="true" then
                    if request.form("frmExpReport")="true" then
                        Dim stFilename : stFilename="attachment; filename=Earned Revenue.xls"
                        Response.ContentType = "application/vnd.ms-excel"
                        Response.AddHeader "Content-Disposition", stFilename
                    end if
                    If request.Form("optSummary")="detail" then
                      ' Filter which series and membership details to display.
                      ' ""  : All Series and Memberships
                      ' "1" : Count Series Only
                      if seriesTypeFilter = "" OR seriesTypeFilter = "1" then
                        strSQL = generateOutstandingCountSQL()

                        rsEntry.CursorLocation = 3
                        rsEntry.open strSQL, cnWS
                        Set rsEntry.ActiveConnection = Nothing

                        if NOT rsEntry.EOF then                     'EOF
                            %>
                            <table class="sortable" id="countSeries">
                                <caption>Outstanding Limited Visit Pricing Option</caption>
                                <thead>
                                    <tr class="mainText">
                                        <th><%=session("ClientHW") & " " & name_hw%></th>
                                        <th><%=paymentRefNum_hw%></th>
                                        <th><%=date_hw%></th>
                                        <th><%=expirationDate_hw%></th>
                                        <!--<th>Last Visit</th>-->
                                        <th><%=series_hw%></th>
                                        <th>Paid <%=amount_hw%></th>
                                        <th><%=purchased_hw%></th>
                                        <th>Used</th>
                                        <% if request.form("frmExpReport")="true" then %><th>UnBooked</th><% end if %>
                                        <th><%=remaining_hw%></th>
                                        <th>Earned <%=amount_hw%></th>
                                        <th>Deferred <%=amount_hw%></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <%
                                    do while NOT rsEntry.EOF
                                        ' Calculated values
                                        'remaining = rsEntry("numClasses") - rsEntry("TotalNumDeducted")
                                        'CB: Bug #4975 - Show the current actual remaining on the series not just the remaining based off activity in the specified date range
                                        remaining = rsEntry("RealRemaining")
                                        used = rsEntry("TotalNumDeducted")
                                        paymentAmtSum = rsEntry("unitPrice") - rsEntry("DiscAmt")
                                        earned = ((used / rsEntry("NumClasses")) * paymentAmtSum)
                                        deferredAmt = paymentAmtSum * (remaining / rsEntry("NumClasses"))

                                        ' Calculated Totals
                                        TotDefAmtCount = TotDefAmtCount + deferredAmt
                                        TotEarned = TotEarned + earned
                                        GrandTotal = GrandTotal + paymentAmtSum
                                        %>
                                        <tr class="mainText">
                                             <% if NOT request.form("frmExpReport")="true"  then %>
                                                 <td><a href="adm_clt_ph.asp?ID=<%=rsEntry("ClientID")%>"><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></a></td>
                                             <% else %>
                                                 <td><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></td>
                                             <% end if %>
                                             <td><%=rsEntry("PmtRefNo")%></td>
                                             <td><%=FmtDateShort(rsEntry("PaymentDate"))%></td>
                                             <td><%=FmtDateShort(rsEntry("ExpDate"))%></td>
                                             <!--<td>FmtDateShort(rsEntry("LastVisit"))</td>-->
                                             <td><%=rsEntry("TypePurch")%></td>
                                             <td><%=exportableCurrency(paymentAmtSum)%></td>
                                             <td><%=rsEntry("NumClasses")%></td>
                                             <td><%=used%></td>
                                             <% if request.form("frmExpReport")="true"  then %><td><%=remaining%></td><% end if %>
                                             <td><%=remaining%></td>
                                             <td><%=exportableCurrency(earned)%></td>
                                             <td><%=exportableCurrency(deferredAmt)%></td>
                                         </tr>
                                         <%
                                         rsEntry.MoveNext
                                    loop
                                    %>
                                </tbody>
                                <tfoot>
                                    <tr class="mainText">
                                        <th><%=total_hw%></th>
                                        <th></th>
                                        <th></th>
                                        <th></th>
                                        <!--<th></th>-->
                                        <th></th>
                                        <th><%=exportableCurrency(GrandTotal)%></th>
                                        <th></th>
                                        <th></th>
                                        <% if request.form("frmExpReport")="true"  then %><th></th><% end if %>
                                        <th></th>
                                        <th><%=exportableCurrency(TotEarned)%></th>
                                        <th><%=exportableCurrency(TotDefAmtCount)%></th>
                                    </tr>
                                </tfoot>
                            </table>
                            <%
                        end if        'eof count series
                              
                        rsEntry.close
                      end if 'seriesTypeFilter :: Count Series display
                      
                      ' Filter which series and membership details to display.
                      ' ""  : All Series and Memberships
                      ' "2" : Time Series Only
                      if seriesTypeFilter = "" OR seriesTypeFilter = "2" then 
                        strSQL =  generateTimeSeriesSQL()
                        DisplayTimeSeries "Time"
                      end if 'seriesTypeFilter :: Time Series display
                      
                      ' Filter which series and membership details to display.
                      ' Everything but Time and Series is going to be a membership so just check for those two.
                      if seriesTypeFilter <> "1" AND seriesTypeFilter <> "2" then 
                        strSQL =  generateMembershipSeriesSQL()
                        DisplayTimeSeries "Membership"
                      end if 'seriesTypeFilter :: Time Series display
                    else 'Summary view
                        %>
                        <table class="maintext">
                            <caption>Earned Revenue</caption>
                            <thead>
                                <tr class="mainText">
                                    <th>&nbsp;</th>
                                    <th colspan="3" class="odd">Unlimited Visit Pricing Option</th>
                                    <th colspan="3">Membership Series</th>
                                    <th colspan="3" class="odd">Limited Visit Pricing Option</th>
                                    <th colspan="3">Expired Series</th>
                                    <th colspan="3" class="odd">Total</th>
                                </tr>
                                <tr class="mainText">
                                    <th>Program</th>
                                    <th class="odd"><em>Subtotal</em></th>
                                    <th class="odd"><em>Discount</em></th>
                                    <th class="odd"><em>Earned</em></th>
                                    
                                    <th><em>Subtotal</em></th>
                                    <th><em>Discount</em></th>
                                    <th><em>Earned</em></th>
                                    
                                    <th class="odd"><em>Subtotal</em></th>
                                    <th class="odd"><em>Discount</em></th>
                                    <th class="odd"><em>Earned</em></th>
                                    
                                    <th><em>Subtotal</em></th>
                                    <th><em>Discount</em></th>
                                    <th><em>Earned</em></th>
                                    
                                    <th class="odd"><em>Subtotal</em></th>
                                    <th class="odd"><em>Discount</em></th>
                                    <th class="odd"><em>Earned</em></th>
                                </tr>
                            </thead>
                            <tbody>
                                <%
                                strSQL = generateEarnedSummarySQL()
                                
                                rsEntry.CursorLocation = 3
                                rsEntry.open strSQL, cnWS
                                Set rsEntry.ActiveConnection = Nothing

                                'rsEntry.MoveFirst

                                do while NOT rsEntry.EOF

                                    if rsEntry("CountRev")<>0 OR rsEntry("CountDisc")<>0 OR rsEntry("TimeRev")<>0 OR rsEntry("TimeDisc")<>0 OR rsEntry("ExpRev")<>0 OR rsEntry("ExpDisc")<>0 OR rsEntry("MemberRev")<>0 OR rsEntry("MemberDisc")<>0 then

                                        ''Clear & Calculate Values

                                        timeSeriesSubtotal = ""
                                        timeSeriesDiscount = ""
                                        timeSeriesEarned = ""
                                        countSeriesSubtotal = ""
                                        countSeriesDiscount = ""
                                        countSeriesEarned = ""
                                        memberSeriesSubtotal = ""
                                        memberSeriesDiscount = ""
                                        memberSeriesEarned = ""
                                        expiredSeriesSubtotal = ""
                                        expiredSeriesDiscount = ""
                                        expiredSeriesEarned = ""
                                        totalSubtotal = ""
                                        totalDiscount = ""
                                        totalEarned = ""

                                        program = rsEntry("TypeGroup")

                                        'Time Series
                                        if rsEntry("TimeRev")<>0 OR rsEntry("TimeDisc")<>0 then 
                                            if rsEntry("TimeDisc")<>0 then
                                                timeSeriesDiscount = exportableCurrency(rsEntry("TimeDisc"))
                                            end if
                                            timeSeriesSubtotal = exportableCurrency(rsEntry("TimeRev") + rsEntry("TimeDisc"))
                                            timeSeriesEarned = exportableCurrency(rsEntry("TimeRev"))
                                        end if 
                              
                                        'Count Series
                                        if rsEntry("CountRev")<>0 OR rsEntry("CountDisc")<>0 then 
                                            if rsEntry("CountDisc")<>0 then 
                                                countSeriesDiscount = exportableCurrency(rsEntry("CountDisc"))
                                            end if 
                                            countSeriesSubtotal = exportableCurrency(rsEntry("CountRev")+rsEntry("CountDisc"))
                                            countSeriesEarned = exportableCurrency(rsEntry("CountRev"))
                                        end if 

                                        'Member Series
                                        if rsEntry("MemberRev")<>0 OR rsEntry("MemberDisc")<>0 then 
                                            if rsEntry("MemberDisc")<>0 then 
                                                memberSeriesDiscount = exportableCurrency(rsEntry("MemberDisc"))
                                            end if
                                            memberSeriesSubtotal = exportableCurrency(rsEntry("MemberRev")+rsEntry("MemberDisc"))
                                            memberSeriesEarned = exportableCurrency(rsEntry("MemberRev"))
                                        end if 

                                        'Expired Series
                                        if rsEntry("ExpRev")<>0 OR rsEntry("ExpDisc")<>0 then 
                                            if rsEntry("ExpDisc")<>0 then 
                                                expiredSeriesDiscount = exportableCurrency(rsEntry("ExpDisc"))
                                            end if 
                                            expiredSeriesSubtotal = exportableCurrency(rsEntry("ExpRev")+rsEntry("ExpDisc"))
                                            expiredSeriesEarned = exportableCurrency(rsEntry("ExpRev"))
                                        end if 

                                        'Totals
                                        if rsEntry("CountRev")+rsEntry("TimeRev")+rsEntry("MemberRev")+rsEntry("ExpRev") <> 0 OR rsEntry("CountDisc")+rsEntry("TimeDisc")+rsEntry("MemberDisc")+rsEntry("ExpDisc") <> 0 then 
                                            if rsEntry("CountDisc") + rsEntry("TimeDisc") + rsEntry("MemberDisc") + rsEntry("ExpDisc") <> 0 then 
                                                totalDiscount = exportableCurrency(rsEntry("CountDisc")+ rsEntry("TimeDisc")+ rsEntry("MemberDisc")+ rsEntry("ExpDisc"))
                                            end if 
                                            totalSubtotal = exportableCurrency(rsEntry("CountDisc")+ rsEntry("TimeDisc")+ rsEntry("MemberDisc")+ rsEntry("ExpDisc")+_
                                                                               rsEntry("CountRev") + rsEntry("TimeRev") + rsEntry("MemberRev") + rsEntry("ExpRev"))                                        
                                            totalEarned = exportableCurrency(rsEntry("CountRev")+rsEntry("TimeRev")+rsEntry("MemberRev")+rsEntry("ExpRev"))
                                        end if 

                                        'Calculate totals each iteration
                                        timeRev = timeRev + rsEntry("TimeRev")
                                        countRev = countRev + rsEntry("CountRev")
                                        expRev = expRev + rsEntry("ExpRev")
                                        memRev = memRev + rsEntry("MemberRev")
                                        totalRev = totalRev + rsEntry("CountRev")+rsEntry("TimeRev")+rsEntry("ExpRev")+rsEntry("MemberRev")
                                        timeDisc = timeDisc + rsEntry("TimeDisc")
                                        countDisc = countDisc + rsEntry("CountDisc")
                                        expDisc = expDisc + rsEntry("ExpDisc")
                                        memDisc = memDisc + rsEntry("MemberDisc")
                                        totalDisc = totalDisc + rsEntry("CountDisc")+rsEntry("TimeDisc")+rsEntry("ExpDisc")+rsEntry("MemberDisc")

                                        'Draw Values
                                        %>
                                        <tr class="mainText">
                                            <td><%=rsEntry("TypeGroup")%></td>
                                            <td class="odd"><%= timeSeriesSubtotal %></td>
                                            <td class="odd"><%= timeSeriesDiscount %></td>
                                            <td class="odd"><%= timeSeriesEarned %></td>
                                            
                                            <td><%= memberSeriesSubtotal %></td>
                                            <td><%= memberSeriesDiscount %></td>
                                            <td><%= memberSeriesEarned %></td>
                                            
                                            <td class="odd"><%= countSeriesSubtotal %></td>
                                            <td class="odd"><%= countSeriesDiscount %></td>
                                            <td class="odd"><%= countSeriesEarned %></td>
                                            
                                            <td><%= expiredSeriesSubtotal %></td>
                                            <td><%= expiredSeriesDiscount %></td>
                                            <td><%= expiredSeriesEarned %></td>
                                            
                                            <td class="odd"><%= totalSubtotal %></td>
                                            <td class="odd"><%= totalDiscount %></td>
                                            <td class="odd"><%= totalEarned %></td>
                                        </tr>
                                        <%                
                                    end if

                                    rsEntry.MoveNext
                                loop
                                %>
                            </tbody>
                            <tfoot>
                                <tr class="mainText">
                                    <th><%=total_hw%>:</th>
                                    
                                    <th class="odd"><%=exportableCurrency(timeRev+timeDisc)%></th>
                                    <th class="odd"><%=exportableCurrency(timeDisc)%></th>
                                    <th class="odd"><%=exportableCurrency(timeRev)%></th>
                                    
                                    <th><%=exportableCurrency(memRev+memDisc)%></th>
                                    <th><%=exportableCurrency(memDisc)%></th>
                                    <th><%=exportableCurrency(memRev)%></th>
                                    
                                    <th class="odd"><%=exportableCurrency(countRev+countDisc)%></th>
                                    <th class="odd"><%=exportableCurrency(countDisc)%></th>
                                    <th class="odd"><%=exportableCurrency(countRev)%></th>
                                    
                                    <th><%=exportableCurrency(expRev+expDisc)%></th>
                                    <th><%=exportableCurrency(expDisc)%></th>
                                    <th><%=exportableCurrency(expRev)%></th>
                                    
                                    <th class="odd"><%=exportableCurrency(totalRev+totalDisc)%></th>
                                    <th class="odd"><%=exportableCurrency(totalDisc)%></th>
                                    <th class="odd"><%=exportableCurrency(totalRev)%></th>
                                </tr>
                            </tfoot>
                        </table>
                        <%
                        rsEntry.close
                    end if 'Summary/Full view
                end if 'end of generate report if statement

                set rsEntry = nothing
                %>
            </div>
<%end if ' tagging %>
        </div>
        <% pageEnd %>
<!-- #include file="post.asp" -->

<%

end if

sub DisplayTimeSeries(timeSeriesType)
  rsEntry.CursorLocation = 3
  rsEntry.open strSQL, cnWS
  Set rsEntry.ActiveConnection = Nothing
  
  TotEarnedTime = 0
  TotDefAmtTime = 0

  if NOT rsEntry.EOF then                       'EOF
      %>
      <table class="sortable" id="<%=timeSeriesType %>Series">
          <caption><%if timeSeriesType = "Time" then RW "Outstanding Unlimited Visit Pricing Option" else RW "Outstanding Membership Pricing Option"%></caption>
          <thead>
              <tr class="mainText">
                  <th><%=session("ClientHW")%> <%= name_hw%></th>
                  <th><%=paymentRefNum_hw%></th>
                  <th><%=date_hw%></th>
                  <th><%=expirationDate_hw%></th>
                  <!--<th>Last Visit</th>-->
                  <th><%=series_hw%></th>
                  <th>Paid <%=amount_hw%></th>
                  <th>Duration (Days)</th>
                  <th>Days Used</th>
                  <th>Days <%=remaining_hw%></th>
                  <th>Earned <%=amount_hw%></th>
                  <th>Deferred <%=amount_hw%></th>
              </tr>
          </thead>
          <tbody>
              <%
              do while NOT rsEntry.EOF
                  PaymentAmtSum = rsEntry("UnitPrice") - rsEntry("DiscAmt")

                  if rsEntry("DurationUnit")=1 then   'in days
                      DurationDay = rsEntry("Duration")
                  else        'duration in months
                      DurationDay = rsEntry("Duration")*30
                  end if

                  if DurationDay = 0 then ' old data for "1 day"
                      DurationDay = 1
                  end if

                  DayRem= INT((rsEntry("ExpDate")-tzNow()))
                  if DayRem < 0 or isNull(DayRem) then
                      DayRem = 0
                  end if

                  DaysUsed = DurationDay - DayRem
                  'EarnedAmount = (DaysUsed / DurationDay) * PaymentAmtSum
                  EarnedAmount = rsEntry("RevPerTG")

                  'DeferredAmt = PaymentAmtSum * rsEntry("diffAsOfDate") / DurationDay
                  DeferredAmt = PaymentAmtSum - EarnedAmount

                  TotDefAmtTime=TotDefAmtTime + DeferredAmt
                  TotEarnedTime = TotEarnedTime + EarnedAmount
                  GrandTotal=GrandTotal + DeferredAmt
                  %>
                  <tr class="mainText">
                      <% if NOT request.form("frmExpReport")="true"  then %>
                          <td><a href="adm_clt_ph.asp?ID=<%=rsEntry("ClientID")%>"><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></a></td>
                      <% else %>
                          <td><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></td>
                      <% end if %>
                      <td><%=rsEntry("PmtRefNo")%></td>
                      <td><%=FmtDateShort(rsEntry("PaymentDate"))%></td>
                      <td><%=FmtDateShort(rsEntry("ExpDate"))%></td>
                      <!--<td>FmtDateShort(rsEntry("LastVisit"))</td>-->
                      <td><%=rsEntry("TypePurch")%></td>
                      <td><%=exportableCurrency(PaymentAmtSum)%></td>
                      <td><%=DurationDay%></td>
                      <td><%=DaysUsed%></td>
                      <td><%=DayRem%></td>
                      <td><%=exportableCurrency(EarnedAmount)%></td>
                      <td><%=exportableCurrency(DeferredAmt)%></td>
                  </tr>
                  <%
                  rsEntry.MoveNext
              loop
              %>
          </tbody>
          <tfoot>
              <tr class="mainText">
                  <th><%=total_hw%></th>
                  <th></th>
                  <th></th>
                  <th></th>
                  <!--<th></th>-->
                  <th></th>
                  <th></th>
                  <th></th>
                  <th></th>
                  <th></th>
                  <th><%=exportableCurrency(TotEarnedTime)%></th>
                  <th><%=exportableCurrency(TotDefAmtTime)%></th>
              </tr>
          </tfoot>
      </table>
      <%
  end if        'eof time series
  rsEntry.close
end sub 'DisplayTimeSeries

function exportableCurrency(currencyVal)
    if NOT request.form("frmExpReport") = "true" Then
        exportableCurrency = FmtCurrency(currencyVal)
    Else
        exportableCurrency = FmtNumber(currencyVal)
    end If
end function

function tzOffset(dateVal)
    tzOffset = CDATE(DateAdd("h", Session("tzOffset"), dateVal))
end Function

function tzNow()
    tzNow = tzOffset(Now)
end function

function generateOutstandingCountSQL()

    strSQL = "SELECT SUM(((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0)) / [PAYMENT DATA].NumClasses) * [VISIT DATA].NumDeducted) AS RevPerTG, SUM(ISNULL([Sales Details].DiscAmt, 0) / ([PAYMENT DATA].NumClasses) * [VISIT DATA].NumDeducted) as CountDiscPerTG, [PAYMENT DATA].TypeGroup, [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, [PAYMENT DATA].PmtRefNo, [PAYMENT DATA].PaymentDate, [PAYMENT DATA].ExpDate, [PAYMENT DATA].TypePurch, ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0) AS PaymentAmtSum, [PAYMENT DATA].NumClasses, [PAYMENT DATA].RealRemaining, [Sales Details].Location, [PAYMENT DATA].Type, [PAYMENT DATA].[Current Series], VisitData_REM.TotalNumDeducted, [Sales Details].UnitPrice, [Sales Details].DiscAmt "
    strSQL = strSQL & "FROM [VISIT DATA] "
    strSQL = strSQL & "INNER JOIN [PAYMENT DATA] ON [VISIT DATA].PmtRefNo = [PAYMENT DATA].PmtRefNo "
    strSQL = strSQL & "INNER JOIN [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo "
    strSQL = strSQL & "INNER JOIN CLIENTS ON CLIENTS.ClientID = [PAYMENT DATA].ClientID "
	if request.form("optFilterTagged")="on" then
		strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
		if session("mvaruserID")<>"" then
			strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
		else
			strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
		end if
	end if
    strSQL = strSQL & "INNER JOIN (SELECT PmtRefNo, SUM(NumDeducted) AS TotalNumDeducted FROM [VISIT DATA] WHERE (ClassDate >= '" & CSDate & "') AND (ClassDate <= '" & CEDate & "') GROUP BY PmtRefNo) VisitData_REM ON [PAYMENT DATA].PmtRefNo = VisitData_REM.PmtRefNo "
    strSQL = strSQL & "WHERE ([Sales Details].CategoryID <= 20) AND ([VISIT DATA].ClassDate >= '" & CSDate & "') "
    strSQL = strSQL & "AND ([VISIT DATA].ClassDate <= '" & CEDate & "') "
    strSQL = strSQL & "AND ([PAYMENT DATA].Type = 1) "
    strSQL = strSQL & "AND ([VISIT DATA].NumDeducted > 0) "
    strSQL = strSQL & "AND ([PAYMENT DATA].NumClasses > 0) " &_
                      "AND ([PAYMENT DATA].Returned = 0) "
    strSQL = strSQL & "AND [VISIT DATA].VisitType<>-1 "
    if cLoc<>"0" then
        strSQL = strSQL & "AND [VISIT DATA].Location = " & cLoc & " "
    end If
    if request.form("optTG")<>"0" Then
        strSQL = strSQL & "AND [PAYMENT DATA].TypeGroup = " & request.form("optTG") 'CCP 12/8/09 Bug # 2217
    end If
    
    strSQL = strSQL & "GROUP BY [PAYMENT DATA].TypeGroup, [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, [PAYMENT DATA].PmtRefNo, [PAYMENT DATA].PaymentDate, [PAYMENT DATA].ExpDate, [PAYMENT DATA].TypePurch, ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0), [PAYMENT DATA].NumClasses, [PAYMENT DATA].RealRemaining, [Sales Details].Location, [PAYMENT DATA].Type, [PAYMENT DATA].[Current Series], VisitData_REM.TotalNumDeducted, [Sales Details].UnitPrice, [Sales Details].DiscAmt, [PAYMENT DATA].ActiveDate "

    ' Order by
    select case request.form("optSortBy")
        case "0" strOrder = "[PAYMENT DATA].ActiveDate"
        case "1" strOrder = "[PAYMENT DATA].ExpDate"
        case "2" strOrder = "CLIENTS.LastName, CLIENTS.FirstName"
        case "3" strOrder = "RevPerTG"
    end Select
    if strOrder<>"" then
        strSQL = strSQL & "ORDER BY " & strOrder
    end if


       response.write debugSQL(strSQL, "OutstandingCount")
    generateOutstandingCountSQL = strSQL

end function

function generateTimeSeriesSQL()
  generateTimeSeriesSQL = generateTimeOrMembershipSeriesSQL("time")
end function

function generateMembershipSeriesSQL()
  generateMembershipSeriesSQL = generateTimeOrMembershipSeriesSQL("membership")
end function

function generateTimeOrMembershipSeriesSQL(reportType)
    strSQL = "SELECT SUM((((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, CASE WHEN [PAYMENT DATA].ActiveDate < '" & CSDate & "' THEN '" & CSDate & "' ELSE [PAYMENT DATA].ActiveDate END, CASE WHEN [PAYMENT DATA].ExpDate > '" & CEDate & "' THEN '" & CEDate & "' ELSE [PAYMENT DATA].ExpDate END) + 1)) / (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) AS RevPerTG, SUM((((ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, CASE WHEN [PAYMENT DATA].ActiveDate < '" & CSDate & "' THEN '" & CSDate & "' ELSE [PAYMENT DATA].ActiveDate END, CASE WHEN [PAYMENT DATA].ExpDate > '" & CEDate & "' THEN '" & CEDate & "' ELSE [PAYMENT DATA].ExpDate END) + 1)) / (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) as TimeDiscPerTG, [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, [PAYMENT DATA].TypeGroup, [PAYMENT DATA].PmtRefNo, [PAYMENT DATA].PaymentDate, [PAYMENT DATA].ExpDate, [PAYMENT DATA].TypePurch, ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0) AS PaymentAmtSum, [PAYMENT DATA].Duration, [PAYMENT DATA].DurationUnit, [Sales Details].Location, [PAYMENT DATA].Type, [PAYMENT DATA].[Current Series], [Sales Details].UnitPrice, [Sales Details].DiscAmt "
    strSQL = strSQL & "FROM [PAYMENT DATA] "
    strSQL = strSQL & "INNER JOIN [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo "
    strSQL = strSQL & "INNER JOIN CLIENTS ON CLIENTS.ClientID = [PAYMENT DATA].ClientID " &_
                      "INNER JOIN tblSeriesType ON tblSeriesType.SeriesTypeID = [PAYMENT DATA].Type "
    
	if request.form("optFilterTagged")="on" then
		strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
		if session("mvaruserID")<>"" then
			strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
		else
			strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
		end if
	end if
    strSQL = strSQL & "WHERE ([Sales Details].CategoryID <= 20) AND ([PAYMENT DATA].ActiveDate <= '" & CEDate & "') "
    strSQL = strSQL & "AND ([PAYMENT DATA].ExpDate >= '"& CSDate & "') "
    
    if reportType = "time" then
      strSQL = strSQL & "AND ([PAYMENT DATA].Type = 2) "
    else ' membership
      strSQL = strSQL & "AND (tblSeriesType.isSystem = 0) "
      ' If the filter is looking for a specific membership type
      if seriesTypeFilter <> "" AND seriesTypeFilter <> "1" AND seriesTypeFilter <> "2" AND seriesTypeFilter <> "3" AND seriesTypeFilter <> "-3" then 
        strSQL = strSQL & " AND ([PAYMENT DATA].TYPE = " & sqlInjectStr(seriesTypeFilter) & ") "
      end if
    end if
    strSQL = strSQL & "AND ([PAYMENT DATA].Returned = 0) "
        
    if cLoc<>"0" then
        strSQL = strSQL & "AND [Sales Details].Location = " & cLoc & " "
    end if
    if request.form("optTG")<>"0" Then
        strSQL = strSQL & "AND [PAYMENT DATA].TypeGroup = " & request.form("optTG")
    end If
    
    strSQL = strSQL & "GROUP BY [PAYMENT DATA].ClientID, CLIENTS.LastName, CLIENTS.FirstName, [PAYMENT DATA].TypeGroup, [PAYMENT DATA].PmtRefNo, [PAYMENT DATA].PaymentDate, [PAYMENT DATA].ExpDate, [PAYMENT DATA].TypePurch, ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0), [PAYMENT DATA].Duration, [PAYMENT DATA].DurationUnit, [Sales Details].Location, [PAYMENT DATA].Type, [PAYMENT DATA].[Current Series], [Sales Details].UnitPrice, [Sales Details].DiscAmt, [PAYMENT DATA].ActiveDate "

    ' Order by
    select case request.form("optSortBy")
        case "0" strOrder = "[PAYMENT DATA].ActiveDate"
        case "1" strOrder = "[PAYMENT DATA].ExpDate"
        case "2" strOrder = "CLIENTS.LastName, CLIENTS.FirstName"
        case "3" strOrder = "RevPerTG"
    end Select
    if strOrder<>"" then
        strSQL = strSQL & "ORDER BY " & strOrder
    end if

    response.write debugSQL(strSQL, reportType & "Series")

    generateTimeOrMembershipSeriesSQL = strSQL

end function

function generateEarnedSummarySQL()
    ''TODO make it work better with the filters in this report

    if request.form("frmTagClients") <> "true" then
      strSQL = " SELECT tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup, " &_
               " ISNULL(VDSum.CompVisits, 0) AS CompVisits, ISNULL(VDSum.UnpaidVisits, 0) as UnpaidVisits, " &_
               " ISNULL(VDSum.CountVisits, 0) AS CountSeriesVisits, ISNULL(VDSum.TimeVisits, 0) AS TimeSeriesVisits, " &_
               " ISNULL(VDSum.MemberVisits, 0) AS MemberSeriesVisits, ISNULL(CountRev.RevPerTG, 0) AS CountRev, " &_
               " ISNULL(CountRev.CountDiscPerTG, 0) AS CountDisc, ISNULL(TimeRev.RevPerTG, 0) AS TimeRev, " &_
               " ISNULL(TimeRev.TimeDiscPerTG, 0) AS TimeDisc, ISNULL(MemberRev.RevPerTG, 0) AS MemberRev, " &_
               " ISNULL(MemberRev.MemberDiscPerTG, 0) AS MemberDisc, ISNULL(ExpRev.ExpirationRev, 0) AS ExpRev, " &_
               " ISNULL(ExpDiscPerTG, 0) AS ExpDisc " &_
               " FROM tblTypeGroup "
    else
      strSQL = ""
    end if

    ' Visit Data Sum
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " LEFT OUTER JOIN " &_
                        " (SELECT SUM(CASE WHEN [PAYMENT DATA].Type = 1 AND [VISIT DATA].Value = 1 THEN 1 ELSE 0 END) AS CountVisits, " &_
                        " SUM(CASE WHEN [PAYMENT DATA].Type = 9 THEN 1 ELSE 0 END) AS UnpaidVisits, " &_
                        " SUM(CASE WHEN ([PAYMENT DATA].Type = 2 OR tblSeriesType.isSystem = 0) AND [VISIT DATA].Value = 1 THEN 1 ELSE 0 END) AS TimeVisits, " &_
                        " SUM(CASE WHEN (tblSeriesType.IsSystem=0) AND [VISIT DATA].Value = 1 THEN 1 ELSE 0 END)  AS MemberVisits, " &_
                        " SUM(CASE WHEN [VISIT DATA].Value = 0 THEN 1 ELSE 0 END) AS CompVisits, [VISIT DATA].TypeGroup "
    else
      strSQL = strSQL & " (SELECT [PAYMENT DATA].ClientId "
    end if

    strSQL = strSQL & " FROM [VISIT DATA] INNER JOIN "
    strSQL = strSQL & " [PAYMENT DATA] ON [VISIT DATA].PmtRefNo = [PAYMENT DATA].PmtRefNo "

	  if request.form("optFilterTagged")="on" then
		  strSQL = strSQL & "INNER JOIN tblClientTag ON [PAYMENT DATA].ClientID = tblClientTag.clientID "
		  if session("mvaruserID")<>"" then
			  strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
		  else
			  strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
		  end if
	  end if
    
    strSQL = strSQL & " INNER JOIN tblSeriesType ON [Payment Data].Type = tblSeriesType.SeriesTypeID "

    strSQL = strSQL & " WHERE ([VISIT DATA].ClassDate >= '" & cSDate & "') AND ([VISIT DATA].ClassDate <= '" & cEDate & "') "
    strSQL = strSQL & " AND [VISIT DATA].VisitType<>-1 " &_
                      "AND ([PAYMENT DATA].Returned = 0) "
    
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " GROUP BY [VISIT DATA].TypeGroup) VDSum ON  "
      strSQL = strSQL & " VDSum.TypeGroup = tblTypeGroup.TypeGroupID " 
    else
      strSQL = strSQL & ") "
    end if


    ' Count Revenue
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " LEFT OUTER JOIN " &_
                        " (SELECT SUM(((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0))  " &_
                        " / [PAYMENT DATA].NumClasses) * [VISIT DATA].NumDeducted) AS RevPerTG, SUM(ISNULL([Sales Details].DiscAmt, 0)/ ([PAYMENT DATA].NumClasses) * [VISIT DATA].NumDeducted) as CountDiscPerTG, [PAYMENT DATA].TypeGroup "
    else
      strSQL = strSQL & " UNION " &_ 
                        " (SELECT [PAYMENT DATA].ClientId "
    end if

    strSQL = strSQL & " FROM [VISIT DATA] INNER JOIN "
    strSQL = strSQL & " [PAYMENT DATA] ON [VISIT DATA].PmtRefNo = [PAYMENT DATA].PmtRefNo "
	if request.form("optFilterTagged")="on" then
		strSQL = strSQL & "INNER JOIN tblClientTag ON [PAYMENT DATA].ClientID = tblClientTag.clientID "
		if session("mvaruserID")<>"" then
			strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
		else
			strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
		end if
	end if
    strSQL = strSQL & " INNER JOIN [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo "

    strSQL = strSQL & " WHERE ([Sales Details].CategoryID <= 20) AND ([VISIT DATA].ClassDate >= '" & cSDate & "') AND ([VISIT DATA].ClassDate <= '" & cEDate & "') AND ([PAYMENT DATA].Type = 1) AND ([VISIT DATA].NumDeducted > 0) AND ([PAYMENT DATA].NumClasses > 0)" &_
                      "AND ([PAYMENT DATA].Returned = 0) "
    strSQL = strSQL & " AND [VISIT DATA].VisitType<>-1 "
    if cLoc<>"0" then
        strSQL = strSQL & "AND [Sales Details].Location = " & cLoc & " "
    end if
    if request.form("optTG")<>"0" Then
        strSQL = strSQL & "AND [PAYMENT DATA].TypeGroup = " & request.form("optTG") 'CCP 12/8/09 Bug #2217
    end If
    
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " GROUP BY [PAYMENT DATA].TypeGroup) CountRev ON CountRev.TypeGroup = tblTypeGroup.TypeGroupID "
    else
      strSQL = strSQL & ") "
    end if


    ' Time Revenue
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " LEFT OUTER JOIN " &_
                        " (SELECT SUM((((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, " &_
                        " CASE WHEN [PAYMENT DATA].ActiveDate < '" & cSDate & "' THEN '" & cSDate & "' ELSE [PAYMENT DATA].ActiveDate END,  " &_
                        " CASE WHEN [PAYMENT DATA].ExpDate > '" & cEDate & "' THEN '" & cEDate & "' ELSE [PAYMENT DATA].ExpDate END) + 1)) /  (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) AS RevPerTG,  " &_
                        " SUM((((ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, " &_
                        " CASE WHEN [PAYMENT DATA].ActiveDate < '" & cSDate & "' THEN '" & cSDate & "' ELSE [PAYMENT DATA].ActiveDate END,  " &_
                        " CASE WHEN [PAYMENT DATA].ExpDate > '" & cEDate & "' THEN '" & cEDate & "' ELSE [PAYMENT DATA].ExpDate END) + 1)) /  (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) as TimeDiscPerTG, [PAYMENT DATA].TypeGroup "
    else
      strSQL = strSQL & " UNION " &_ 
                        " (SELECT [PAYMENT DATA].ClientId "
    end if

    strSQL = strSQL & " FROM [PAYMENT DATA] "
	if request.form("optFilterTagged")="on" then
		strSQL = strSQL & "INNER JOIN tblClientTag ON [PAYMENT DATA].ClientID = tblClientTag.clientID "
		if session("mvaruserID")<>"" then
			strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
		else
			strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
		end if
	end if
    strSQL = strSQL & " INNER JOIN [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo "
    if cLoc<>"0" then
        strSQL = strSQL & "AND [Sales Details].Location = " & cLoc & " "
    end if
    if request.form("optTG")<>"0" Then
        strSQL = strSQL & "AND [PAYMENT DATA].TypeGroup = " & request.form("optTG")
    end If

    strSQL = strSQL & " WHERE ([Sales Details].CategoryID <= 20) AND ([PAYMENT DATA].ActiveDate <= '" & cEDate & "') " &_
                      "AND ([PAYMENT DATA].ExpDate >= '" & cSDate & "') AND ([PAYMENT DATA].Type = 2) " &_
                      "AND ([PAYMENT DATA].Returned = 0) "
    
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " GROUP BY [PAYMENT DATA].TypeGroup) TimeRev ON TimeRev.TypeGroup = tblTypeGroup.TypeGroupID "
    else
      strSQL = strSQL & ") "
    end if

    ' Member Revenue
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " LEFT OUTER JOIN " &_
                        " (SELECT SUM((((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, " &_
                        " CASE WHEN [PAYMENT DATA].ActiveDate < '" & cSDate & "' THEN '" & cSDate & "' ELSE [PAYMENT DATA].ActiveDate END,  " &_
                        " CASE WHEN [PAYMENT DATA].ExpDate > '" & cEDate & "' THEN '" & cEDate & "' ELSE [PAYMENT DATA].ExpDate END) + 1)) /  (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) AS RevPerTG,  " &_
                        " SUM((((ISNULL([Sales Details].DiscAmt, 0)) * (DATEDIFF(d, " &_
                        " CASE WHEN [PAYMENT DATA].ActiveDate < '" & cSDate & "' THEN '" & cSDate & "' ELSE [PAYMENT DATA].ActiveDate END,  " &_
                        " CASE WHEN [PAYMENT DATA].ExpDate > '" & cEDate & "' THEN '" & cEDate & "' ELSE [PAYMENT DATA].ExpDate END) + 1)) /  (CASE WHEN [PAYMENT DATA].ActiveDate > [PAYMENT DATA].ExpDate THEN 1 ELSE DATEDIFF(d, [PAYMENT DATA].ActiveDate, [PAYMENT DATA].ExpDate) + 1 END))) as MemberDiscPerTG, [PAYMENT DATA].TypeGroup "
    else
      strSQL = strSQL & " UNION " &_ 
                        " (SELECT [PAYMENT DATA].ClientId "
    end if

    strSQL = strSQL & " FROM [PAYMENT DATA] "
	if request.form("optFilterTagged")="on" then
		strSQL = strSQL & "INNER JOIN tblClientTag ON [PAYMENT DATA].ClientID = tblClientTag.clientID "
		if session("mvaruserID")<>"" then
			strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
		else
			strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
		end if
	end if
    strSQL = strSQL & " INNER JOIN [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo " &_
                      "INNER JOIN tblSeriesType ON tblSeriesType.SeriesTypeID = [PAYMENT DATA].Type "

    strSQL = strSQL & " WHERE ([Sales Details].CategoryID <= 20) AND ([PAYMENT DATA].ActiveDate <= '" & cEDate & "') " &_
                      "AND ([PAYMENT DATA].ExpDate >= '" & cSDate & "') " &_
                      "AND (tblSeriesType.isSystem = 0) AND ([PAYMENT DATA].Returned = 0) "
    if cLoc<>"0" then
        strSQL = strSQL & "AND [Sales Details].Location = " & cLoc & " "
    end if
    if request.form("optTG")<>"0" Then
        strSQL = strSQL & "AND [PAYMENT DATA].TypeGroup = " & request.form("optTG")
    end If
    
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " GROUP BY [PAYMENT DATA].TypeGroup) MemberRev ON MemberRev.TypeGroup = tblTypeGroup.TypeGroupID "
    else
      strSQL = strSQL & ") "
    end if


    ' Expired Revenue
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " LEFT OUTER JOIN " &_
                        " (SELECT SUM((ISNULL([Sales Details].UnitPrice, 0) - ISNULL([Sales Details].DiscAmt, 0)) / [PAYMENT DATA].NumClasses * [PAYMENT DATA].Remaining) AS ExpirationRev, " &_
                        " SUM((ISNULL([Sales Details].DiscAmt, 0)) / [PAYMENT DATA].NumClasses * [PAYMENT DATA].Remaining) as ExpDiscPerTG, [PAYMENT DATA].TypeGroup "
    else
      strSQL = strSQL & " UNION " &_ 
                        " (SELECT [PAYMENT DATA].ClientId "
    end if

    strSQL = strSQL & " FROM [PAYMENT DATA] "
	if request.form("optFilterTagged")="on" then
		strSQL = strSQL & "INNER JOIN tblClientTag ON [PAYMENT DATA].ClientID = tblClientTag.clientID "
		if session("mvaruserID")<>"" then
			strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
		else
			strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
		end if
	end if
    strSQL = strSQL & " INNER JOIN [Sales Details] ON [PAYMENT DATA].PmtRefNo = [Sales Details].PmtRefNo "

    strSQL = strSQL & " WHERE ([Sales Details].CategoryID <= 20) AND  ([PAYMENT DATA].ExpDate >= '" & cSDate & "') AND ([PAYMENT DATA].ExpDate <= '" & cEDate & "') AND ([PAYMENT DATA].Type = 1) AND "
    strSQL = strSQL & " ([PAYMENT DATA].Remaining > 0) AND ([PAYMENT DATA].Returned = 0) "
    if cLoc<>"0" then
        strSQL = strSQL & "AND [Sales Details].Location = " & cLoc & " "
    end if
    if request.form("optTG")<>"0" Then
        strSQL = strSQL & "AND [PAYMENT DATA].TypeGroup = " & request.form("optTG")
    end If
    
    if request.form("frmTagClients") <> "true" then
      strSQL = strSQL & " GROUP BY [PAYMENT DATA].TypeGroup) ExpRev ON ExpRev.TypeGroup = tblTypeGroup.TypeGroupID "
      
      strSQL = strSQL & " WHERE (tblTypeGroup.Active = 1) OR CountRev.RevPerTG > 0 " 
    else
      strSQL = strSQL & ") "
    end if

    'response.Write strSQL
    'response.end
    response.write debugSQL(strSQL, "EarnedSummary")
                
    generateEarnedSummarySQL = strSQL
end function
%>
