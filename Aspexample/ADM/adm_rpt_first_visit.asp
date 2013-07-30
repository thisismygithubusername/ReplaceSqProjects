<%@ CodePage=65001 %>
<!-- #include file="init.asp" -->
<%
'Dim SessionFarm : set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
Server.ScriptTimeout = 300    '5 min (value in seconds)
%>
<%
    dim rsEntry : set rsEntry = Server.CreateObject("ADODB.Recordset")
    %>
    <!-- #include file="inc_accpriv.asp" -->
    <!-- #include file="inc_rpt_tagging.asp" -->
    <!-- #include file="inc_utilities.asp" -->
    <!-- #include file="inc_rpt_save.asp" -->
    <%
    if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_MARKETING") then
        Response.Write "<script type=""text/javascript"">alert('You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.');javascript:history.go(-1);</script>"
    else
        %>
        <!-- #include file="../inc_i18n.asp" -->
        <!-- #include file="inc_hotword.asp" -->
        <%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

        Dim cEDate, cSDate, CltCount, cTG
        Dim rowColor, ap_view_all_locs, cloc
        CltCount = 0

        useTagSubtract = true
				
		dim category : category = ""
		if (RQ("category"))<>"" then
			category = RQ("category")
		elseif (RF("category"))<>"" then
			category = RF("category") 
		end if

        if request.form("requiredtxtDateStart")<>"" then
            Call SetLocale(session("mvarLocaleStr"))
                cSDate = CDATE(sqlInjectStr(request.form("requiredtxtDateStart")))
            Call SetLocale("en-us")
        else
            cSDate = DateAdd("m",-1,DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))))
        end if

        ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")

        if request.form("requiredtxtDateEnd")<>""  then
            Call SetLocale(session("mvarLocaleStr"))
                cEDate = CDATE(sqlInjectStr(request.form("requiredtxtDateEnd")))
            Call SetLocale("en-us")
        else
            cEDate =DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now)))
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

        if request.form("optPmtTG")<>"" then
            cTG = request.form("optPmtTG")
        else
            cTG = 0
        end if

        'JM - 48_2487
        DIM ss_UseRep2, ss_UseRep3, ss_UseRep4, ss_UseRep5, ss_UseRep6, ap_CLT_REP_VIEW_ALL
        strSQL = "SELECT tblGenOpts.CRMContactLogs, tblGenOpts.UseRep2, tblGenOpts.UseRep3, tblGenOpts.UseRep4, tblGenOpts.UseRep5, tblGenOpts.UseRep6 FROM tblGenOpts"
        rsEntry.CursorLocation = 3
        rsEntry.open strSQL, cnWS
        Set rsEntry.ActiveConnection = Nothing
        ss_UseRep2 = false
        ss_UseRep3 = false
        ss_UseRep4 = false
        ss_UseRep5 = false
        ss_UseRep6 = false

        rsEntry.close
        ap_CLT_REP_VIEW_ALL = validAccessPriv("CLT_REP_VIEW_ALL")

        ' Hotwords
        'dim arrHW : arrHW = getHotWords(array(1,2,7,8,39,64,93,110,111,119,120,121,149,201))
            dim classType_hw : classType_hw = allHotWords(1)
            dim sessionType_hw : sessionType_hw = allHotWords(2)
            dim programClassesEvents_hw : programClassesEvents_hw = allHotWords(7)
            dim location_hw : location_hw = allHotWords(8)
            dim emailAddress_hw : emailAddress_hw = allHotWords(39)
            dim expirationDate_hw : expirationDate_hw = allHotWords(64)
            dim phoneNum_hw : phoneNum_hw = allHotWords(93)
            dim repTwo_hw : repTwo_hw = allHotWords(110)
            dim repThree_hw : repThree_hw = allHotWords(111)
            dim repFour_hw : repFour_hw = allHotWords(119)
            dim repFive_hw : repFive_hw = allHotWords(120)
            dim repSix_hw : repSix_hw = allHotWords(121)
            dim all_hw : all_hw = allHotWords(149)
            dim repOne_hw : repOne_hw = allHotWords(201)

        %>
<% if NOT request.form("frmExpReport")="true" then %>
<!-- #include file="pre.asp" -->
              <!-- #include file="frame_bottom.asp" -->
              
              <!-- #include file="../inc_date_ctrl.asp" -->
              <!-- #include file="inc_help_content.asp" -->
              <!-- #include file="../inc_ajax.asp" -->
              <!-- #include file="../inc_val_date.asp" -->
              <!-- #include file="inc_date_arrows.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_first_visit", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %> 
              <script type="text/javascript">
              function exportReport() {
                  document.frmParameter.frmExpReport.value = "true";
                  document.frmParameter.frmGenReport.value = "true";
                  <% iframeSubmit "frmParameter", "adm_rpt_first.asp" %>
  
                  // wait for all iframes to load before submitting the page
                  var count = $('iframe').length;
                  $('iframe').load(function() {
                      count--;
                      if (count == 0) {
                          saveReport(true);
                      }
                  });
              }
              </script>

<% end if %>
              <!-- #include file="css/report.asp" -->
<% if NOT request.form("frmExpReport")="true" then %>
<%= pageStart %>            
<style type="text/css">
#options {
margin: 0 auto;
}
</style>
	<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			<div class="headText breadcrumbs-old">
			<span class="breadcrumb-item"><a href="/reportslandingpage/index"><%=DisplayPhrase(reportPageTitlesDictionary, "Reports")%></a></span>
			<%if category<>"" then%>
			<span class="breadcrumb-item">&raquo;</span>
			<span class="breadcrumb-item active"><a href="/reportslandingpage/<%=category%>category"><%=GetReportsCategoryPagePhrase(category)%></a></span>
			<% end if %>
			<span class="breadcrumb-item">&raquo;</span>
			<%= DisplayPhrase(reportPageTitlesDictionary,"Firstvisit") %>
			<div id="add-to-favorites">
				<span id="" class="favorite"></span><span class="text"><%= DisplayPhrase(reportPageTitlesDictionary,"Addtofavorites") %></span>
			</div>
			</div>
	<%end if %>
              <div id="container">
			<% if NOT UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
			  <div id="head" class="headText">
                  <%= pp_PageTitle("First Visit") %>
                </div>
			<%end if %>
                <div id="options" class="mainText">
                  <form name="frmParameter" action="adm_rpt_first_visit.asp" method="POST">
                    <input type="hidden" name="frmGenReport" value="" />
                    <input type="hidden" name="frmExpReport" value="" />
					<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
						<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
						<input type="hidden" name="category" value="<%=category%>">
					<% end if %>

                    <label>
                      First Visit Between
                    </label>
                      <input onChange="validateDate(this, '<%=FmtDateShort(cSDate)%>', true);" class="transForm" type="text" size="11" name="requiredtxtDateStart" value="<%=FmtDateShort(cSDate)%>" class="date" />
                      <script type="text/javascript">
	                  var cal1 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateStart'});
	                  cal1.a_tpl.yearscroll = true;
	                  </script>
                      and
                      <input onChange="validateDate(this, '<%=FmtDateShort(cEDate)%>', true);" class="transForm" type="text" size="11" name="requiredtxtDateEnd" value="<%=FmtDateShort(cEDate)%>" class="date" />
                      <script type="text/javascript">
	                  var cal2 = new tcal({'formname':'frmParameter', 'controlname':'requiredtxtDateEnd'});
	                  cal2.a_tpl.yearscroll = true;
                      </script>

                    <label>
                      <%=programClassesEvents_hw%>:
                      <select class="textSmall" name="optPmtTG" size="4" multiple>
                        <option value="0" <% if cTG="0" then response.write "selected" end if%>><%=all_hw & " " & programClassesEvents_hw & "s"%></option>
                        <%
                        strSQL = "SELECT TypeGroupID, TypeGroup FROM tblTypeGroup WHERE (Active = 1) AND ((wsReservation = 1) OR (wsAppointment = 1) OR (wsResource=1) OR (wsEnrollment=1) ) ORDER BY TypeGroup"
                        rsEntry.CursorLocation = 3
                        rsEntry.open strSQL, cnWS
                        Set rsEntry.ActiveConnection = Nothing

                        do while NOT rsEntry.EOF
                            %>
                            <option value="<%=rsEntry("TypeGroupID")%>" <%if inStr((","&cTG&","), (","&rsEntry("TypeGroupID")&",")) then response.write "selected" end if%>><%=rsEntry("TypeGroup")%></option>
                            <%
                            rsEntry.MoveNext
                        loop
                        rsEntry.close
                        %>
                      </select>
                    </label>

                    <% If request.form("optPmtTG")<>"0" and request.form("optPmtTG")<>"" then %>
                      <label>
                        <%
                        if classType_hw = sessionType_hw then
                            Response.Write classType_hw
                        else
                            Response.Write classType_hw & "/" & sessionType_hw
                        end if
                        %>
                        <select class="textSmall" name="optVT">
                          <option value="0" <%if request.form("optVT")="0" then response.write "selected" end if%>>All Class Type/Appointment Types</option>
                          <%
                          strSQL = "SELECT TypeID, TypeName FROM tblVisitTypes "
                          strSQL = strSQL & "WHERE [Active]=1 AND [Delete] = 0 and (typegroup= "& replace(cTG, ",", " OR typegroup=") & ")"
                          strSQL = strSQL & "ORDER BY TypeName"
                          rsEntry.CursorLocation = 3
                          rsEntry.open strSQL, cnWS
                          Set rsEntry.ActiveConnection = Nothing
 
                          do while NOT rsEntry.EOF
                              %>
                              <option value="<%=rsEntry("TypeID")%>" <%if request.form("optVT")=CSTR(rsEntry("TypeID")) then response.write "selected" end if%>><%=rsEntry("TypeName")%></option>
                              <%
                              rsEntry.MoveNext
                          loop
                          rsEntry.close
                          %>
                        </select>
                        <script type="text/javascript">
                        if ('<%=classType_hw%>' == '<%=sessionType_hw%>') {
                            document.frmParameter.optVT.options[0].text = '<%=all_hw%>' +" " + '<%=classType_hw%>' + "s";
                        } else {
                            document.frmParameter.optVT.options[0].text = '<%=all_hw%>' +" " + '<%=classType_hw%>' + "/" + '<%=sessionType_hw%>' + "s";
                        }
                        </script>
                      </label>
                    <% end if %>

                    <label>
                      <select name="optSaleLoc" class="textSmall" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
                        <option value="0" <% if cLoc=0 then response.write "selected" end if %>><%=all_hw & " " & location_hw & "s"%></option>
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
                    </label>

                    <div class="center-ch">
                    <label>
                      Sort By:
                      <select class="textSmall" name="optSortBy">
                        <option value="1" <%if request.form("optSortBy")="1" then response.write "selected" end if%>><%=session("ClientHW")%> Name</option>
                        <option value="2" <%if request.form("optSortBy")="2" then response.write "selected" end if%>># Visits</option>
                        <option value="3" <%if request.form("optSortBy")="3" then response.write "selected" end if%>>Trainer Name</option> %>
                      </select>
                    </label>

                    <label>
                      Include Inactive <%=session("ClientHW")%>s:
                      <input type="checkbox" name="chkIncInactive" <%If request.form("chkIncInactive")="on" then response.write "checked" end if%> />
                    </label>
                    </div>

                    <div class="center-ch">
                    <label>
                      Only show <%=session("ClientHW")%>s With No Visits After the First Visit:
                      <input type="checkbox" name="optOnlyFirst" <%If request.form("optOnlyFirst")="on" then response.write "checked" end if%> />
                    </label>

                    <label>
                    <% taggingFilter %>
                    </label>
                    </div>

                    <div class="center-ch">
                        <% showDateArrows "frmParameter" %>
                    </div>
                    
                    <div class="center-ch">
                    <input class="textSmall" type="button" name="Button" value="Generate" onClick="genReport();" />
                    <%
                    if Session("Pass") AND Session("Admin")<>"false" AND validAccessPriv("RPT_EXPORT") then
                        exportToExcelButton
                    end if
                    %>
                    <%
                    if not Session("Pass") OR Session("Admin")<>"false" OR NOT validAccessPriv("RPT_TAG") then
                        taggingButtons("frmParameter")
                    end if
                    %>

                    <% savingButtons "frmParameter", "First Visit" %>
                    </div>

                  </form>
                </div>
<% end if %>
                <div id="report" class="mainText">
                <br />
                <br />
                  <table class="mainText">
                    <thead>
                    <%
                    if request.form("frmGenReport")="true" then
                        if request.form("frmExpReport")="true" then
                            Dim stFilename
                            stFilename="attachment; filename=""First Visit Report " & Replace(cSDate,"/","-") & " - " & Replace(cEDate,"/","-") & ".xls"""
                            Response.ContentType = "application/vnd.ms-excel"
                            Response.AddHeader "Content-Disposition", stFilename
                        end if

                        strSQL = generateFirstVisitSQL()

                        rsEntry.CursorLocation = 3
                        rsEntry.open strSQL, cnWS
                        Set rsEntry.ActiveConnection = Nothing

                        if NOT rsEntry.EOF then                 'EOF
                            if request.form("frmExpReport")="true" then 
                                %>
                                <tr class="mainTextBig">
                                  <td>First Visit Report</td>
                                  <td>Between: <%=Replace(cSDate,"/","-") & " and " & Replace(cEDate,"/","-")%></td>
                                </tr>
                                <%
                            end if
                            %>
                              <tr>
                                <% if NOT request.form("frmExpReport")="true" then %><th></th><% end if %>
                                <th><%=session("ClientHW")%></th>
                                <th>First Visit</th>
                                <th>First Visit Description (Includes Program - Session Type)</th>
                                <th>Trainer</th>
                                <th style="text-align: right;"># Visits Since First Visit</th>
                                <th><%=phoneNum_hw%></th>
                                <th><%=emailAddress_hw%></th>
                              </tr>
                            </thead>
                            <tbody>
                              <%
                              do while NOT rsEntry.EOF
                                  CltCount = CltCount + 1
                                  %>
                                  <tr>
                                    <% if NOT request.form("frmExpReport")="true" then %><td width="1%"><%=CltCount%>.</td><% end if %>
                                    <td>
                                      <% if NOT request.form("frmExpReport")="true" then %>
                                          <%checkMembershipIcon(rsEntry("ClientID"))%>
                                      <% end if %>
                                      <a href="adm_clt_vh.asp?ID=<%=rsEntry("ClientID")%>"><%=TRIM(rsEntry("LastName")) & ",&nbsp;" & TRIM(rsEntry("FirstName"))%></a>
                                    </td>
                                    <td><%=FmtDateShort(rsEntry("ClassDate"))%></td>
                                    <td><%=rsEntry("TypePurch")%> - <%=rsEntry("TypeGroup")%>/sessionType</td>
                                    <td><%=FmtTrnName(rsEntry("TrainerID"))%></td>
                                    <td style="text-align: right;"><%=rsEntry("VisitCount") - 1%></td>
                                    <td style="white-space: nowrap;">
                                        <%
                                        if not isNULL(rsEntry("CellPhone")) then
                                                if NOT request.form("frmExpReport")="true" then
                                                        response.write "<img src=""" & contentUrl("/asp/adm/images/smart-phone-16px.png") & """ align=""absbottom"" title=""Mobile Phone""> "
                                                end if
                                                response.write FmtPhoneNum(rsEntry("CellPhone"))
                                        elseif NOT isNULL(rsEntry("HomePhone")) then
                                                response.write FmtPhoneNum(rsEntry("HomePhone"))
                                        elseif NOT isNull(rsEntry("WorkPhone")) then
                                                response.write FmtPhoneNum(rsEntry("WorkPhone"))
                                        end if
                                        %>
                                    </td>
                                    <td><%=rsEntry("EmailName")%></td>
                                  </tr>
                                  <%
                                  rsEntry.MoveNext
                              Loop
                              %>
                            </tbody>
                            <%
                        end if              'end of EOF check
                        rsEntry.close
                        set rsEntry = nothing
                    end if  'Generate report check
                    %>
                  </table>
<% if NOT request.form("frmExpReport")="true" then %>
                </div>
              </div>
			<% pageEnd %>
<!-- #include file="post.asp" -->

<% end If %> 
    <%

end if

function generateFirstVisitSQL()

    strSQL = "SELECT CLIENTS.ClientID, CLIENTS.LastName, CLIENTS.FirstName,  CLIENTS.CellPhone, CLIENTS.HomePhone, CLIENTS.WorkPhone, CLIENTS.EmailName, CLIENTS.Inactive, tblTypeGroup.TypeGroup, "
    strSQL = strSQL & "[PAYMENT DATA].PmtRefNo, [PAYMENT DATA].TypePurch, [PAYMENT DATA].ExpDate, [VISIT DATA].ClassDate, VD2.VisitCount, [PAYMENT DATA].NumClasses, [VISIT DATA].TrainerID, TRAINERS.TrLastName, TRAINERS.TrFirstName "
    strSQL = strSQL & " FROM [PAYMENT DATA] INNER JOIN "
    strSQL = strSQL & " [VISIT DATA] ON [PAYMENT DATA].PmtRefNo = [VISIT DATA].PmtRefNo INNER JOIN "
    strSQL = strSQL & " TRAINERS ON [VISIT DATA].TrainerID = TRAINERS.TrainerID INNER JOIN "
    strSQL = strSQL & " tblTypeGroup ON [VISIT DATA].TypeGroup = tblTypeGroup.TypeGroupID INNER JOIN "
    strSQL = strSQL & " (SELECT MIN((CONVERT(DATETIME, [Visit Data].ClassDate + CONVERT(varchar, [Visit Data].ClassTime, 108), 102)))  AS Firstvisit, ClientID, SUM(CASE WHEN [VISIT DATA].[ClassDate] >= '" & cSDate & "' THEN 1 ELSE 0 END) AS VisitCount "
    strSQL = strSQL & " FROM [VISIT DATA] "
    strSQL = strSQL & " WHERE 1=1 "

    ' HAVING clause fliters out clients without this WHERE
    if request.form("optOnlyFirst")<>"on" then 
        strSQL = strSQL & " AND ClassDate <= '" & Date & "' "
    end if

    ' Filter for location
    if cLoc<>0 then
        strSQL = strSQL & " AND ([VISIT DATA].Location=" & cLoc & ") "
    end if 

    ' Filter for typegroup/program
    if LEFT(cTG,1) <> "0" then 'Bug #3048, CCP 2/18/10
        strSQL = strSQL & " AND ([VISIT DATA].TypeGroup=" & replace(cTG, ","," OR [VISIT DATA].TypeGroup=")  & ") "
    end if

    ' Filter for Visittype
    if request.form("optVT")<>"0" and request.form("optVT")<>"" then
        strSQL = strSQL & " AND ([VISIT DATA].VisitType = " & request.form("optVT") & ") "
    end if

    strSQL = strSQL & " GROUP BY ClientID "

    'Having one or more class in date range
    strSQL = strSQL & " HAVING SUM(CASE WHEN [VISIT DATA].[ClassDate] >= '" & cSDate & "' THEN 1 ELSE 0 END) "
    if request.form("optOnlyFirst")="on" then 
        strSQL = strSQL & "= 1 "
    Else
        strSQL = strSQL & ">= 1 "
    end if


    strSQL = strSQL & ") VD2 ON VD2.Firstvisit = (CONVERT(DATETIME, [Visit Data].ClassDate + CONVERT(varchar, [Visit Data].ClassTime, 108), 102)) "
    strSQL = strSQL & " AND [VISIT DATA].ClientID = VD2.ClientID INNER JOIN CLIENTS ON VD2.ClientID = CLIENTS.ClientID "

    ' Filter by tagged clients only
    if request.form("optFilterTagged")="on" then
        strSQL = strSQL & " INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
    end if

    strSQL = strSQL & " WHERE 1=1 "

    ' Filter for location
    if cLoc<>0 then
        strSQL = strSQL & " AND ([VISIT DATA].Location=" & cLoc & ") "
    end if

    ' Filter for typegroup/program
		if LEFT(cTG,1) <> "0" then 'Bug #3048, CCP 2/18/10
        strSQL = strSQL & " AND ([VISIT DATA].TypeGroup=" & replace(cTG, ",", " OR [VISIT DATA].TypeGroup=") & ") "
    end if

    ' Filter for visittype
    if request.form("optVT")<>"0" and request.form("optVT")<>"" then
        strSQL = strSQL & " AND ([VISIT DATA].VisitType = " & request.form("optVT") & ") "
    end if

    ' Filter for include inactive
    if request.form("chkIncInactive")<>"on" then
        strSQL = strSQL & " AND (CLIENTS.Inactive=0) "
    end if

    ' First visit date is between date range
    strSQL = strSQL & " AND [VISIT DATA].[ClassDate] >='" & cSDate & "' "
    strSQL = strSQL & " AND [VISIT DATA].[ClassDate] <='" & cEDate & "' "

    ' Filter by tagged clients only
    if request.form("optFilterTagged")="on" then
        if session("mvaruserID")<>"" then
            strSQL = strSQL & " AND (tblClientTag.smodeID = " & session("mvaruserID") & ") "
        else
            strSQL = strSQL & " AND (tblClientTag.smodeID = 0) "
        end if
    end if

    ' Don't include the order by if tagging clients
    if request.form("frmTagClients")<>"true" then

        ' Filter for Sort by
        if request.form("optSortBy")="2" then
            strSQL = strSQL & " ORDER BY VisitCount DESC, CLIENTS.LastName, CLIENTS.FirstName, Firstvisit "
        else
            if request.form("optSortBy")="3" then
                strSQL = strSQL & "ORDER BY " & GetTrnOrderBy()
            else
                strSQL = strSQL & " ORDER BY CLIENTS.LastName, CLIENTS.FirstName, Firstvisit "
            end if
        end if

    end if

    if debugMode then
       response.write debugSQL(strSQL, "SQL")
    end if

    ' Tag clients
    if request.form("frmTagClients")="true" then
        if request.form("frmTagClientsNew")="true" then
            clearAndTagQuery(strSQL)
       'MB bug#4384 - fixed UnTag Clients button   
        elseif request.form("frmUnTagClients")="true" then
			tagSubtract(strSQL)
        else
            tagQuery(strSQL)
        end if
        strSQL = "SELECT StudioID FROM Studios WHERE 1=0 "
    end if

    generateFirstVisitSQL = strSQL

end function

%>

