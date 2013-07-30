<%@ CodePage=65001 %>
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
        dim rsEntry2
        set rsEntry2 = Server.CreateObject("ADODB.Recordset")
        %>
        <!-- #include file="inc_accpriv.asp" -->
        <!-- #include file="inc_rpt_tagging.asp" -->
        <!-- #include file="inc_utilities.asp" -->
        <!-- #include file="inc_rpt_save.asp" -->
        <!-- #include file="inc_row_colors.asp" -->
<%
        if not Session("Pass") OR Session("Admin")="false" then 'OR NOT validAccessPriv("RPT_DAY") then 
%>
                <script type="text/javascript">
                        alert("You are not authorized to view this page.\nPlease login to an account with Access Privileges granted for this feature.");
                        javascript:history.go(-1);
                </script>
        <% else ' has permission to view page %>
                <!-- #include file="../inc_i18n.asp" -->
                <!-- #include file="../inc_date_ctrl.asp" -->
                <!-- #include file="../inc_jquery.asp" -->
                <!-- #include file="inc_help_content.asp" -->
<%
	if useVersionB(TEST_NAVIGATION_REDESIGN) then
		Session("tabID") = "97"
	end if

                Dim showDetails, cLoc, curIndexID, curIndexValueID
                Dim showHeader, rowcolor, barcolor
                Dim totalClients, totalUnassigned, percent, ap_view_all_locs, ss_CRMContactLogs, ap_CLT_REP_VIEW_ALL, frmOptRep
                
                ap_view_all_locs = validAccessPriv("TB_V_RPT_ALL_LOC")
                ss_CRMContactLogs = checkStudioSetting("tblGenOpts", "CRMContactLogs")
                ap_CLT_REP_VIEW_ALL = validAccessPriv("CLT_REP_VIEW_ALL")
                
                showDetails = request.form("optDisMode")
        
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
                
                if request.form("optIndex")<>"" then
                        curIndexID = CINT(request.form("optIndex"))
                else
                        curIndexID = 0
                end if
                
                if request.form("optIndexValue")<>"" then
                        curIndexValueID = CINT(request.form("optIndexValue"))
                else
                        curIndexValueID = 0
                end if
                
                if request.form("optRep")<>"" then
                        frmOptRep = sqlInjectStr(request.form("optRep"))
                else
                        if NOT ap_CLT_REP_VIEW_ALL then
                                if session("empID")<>"" then
                                        frmOptRep = session("empID")
                                else
                                        frmOptRep = "-3" ' no login assigned, should not be able to view any clients, so only ones assigned to 'System Generated'.
                                end if
                        else
                                frmOptRep = "0"
                        end if
                end if
        
                if NOT request.form("frmExpReport")="true" then
%>
<!-- #include file="pre.asp" -->
                        <!-- #include file="frame_bottom.asp" -->
			
<%= js(array("mb", "MBS", "calendar" & dateFormatCode, "adm/adm_rpt_clt_index", "reportFavorites", "plugins/jquery.SimpleLightBox")) %>
<%= css(array("SimpleLightBox")) %>
                        <script type="text/javascript">
                        function exportReport() {
                                document.frmParameter.frmGenReport.value = "true";
                                document.frmParameter.frmExpReport.value = "true";
                                <% iframeSubmit "frmParameter", "adm_rpt_clt_index.asp" %>
                        }
                        </script>
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
			<%= DisplayPhrase(reportPageTitlesDictionary,"Clientindexes") %>
			<% showNewHelpContentIcon("client-indexes-report") %>
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
										<td class="headText" align="left" valign="top">
												<table width="100%" cellspacing="0">
														<tr>
														<td class="headText" valign="bottom"><b> <%= pp_PageTitle("Client Indexes") %> </b>
														<!--JM - 49_2447-->
														<% showNewHelpContentIcon("client-indexes-report") %>

														</td>
														<td valign="bottom" class="right" height="26"> </td>
														</tr>
												</table>
										</td>
								</tr>
					<%end if %>
                                        <tr> 
                                        <td height="30" valign="bottom" class="headText">
                                        <table class="mainText border4 center-block" cellspacing="0">
                                                <form name="frmParameter" action="adm_rpt_clt_index.asp" method="POST">
                                                <input type="hidden" name="frmGenReport" value="">
                                                <input type="hidden" name="frmExpReport" value="">
												<% if UseVersionB(TEST_REPORT_LANDING_PAGE) then %>
													<input type="hidden" name="reportUrl" id="reportUrl" value="<%=Request.ServerVariables("SCRIPT_NAME") %>" />
													<input type="hidden" name="category" value="<%=category%>">
												<% end if %>
                                                <tr> 
                                                <td class="center-ch" valign="middle" style="background-color:#F2F2F2;"><b>&nbsp;&nbsp;Index:&nbsp;
                                                
                                                <select name="optIndex" onchange="document.frmParameter.submit();">
                                                        <option value="0">Select An Index</option>
<%
                                                        strSQL = "SELECT ClientIndexName, ClientIndexID FROM tblClientIndex "
                                                        if request.form("optIncInactiveIndex")="" then
                                                                strSQL = strSQL & "WHERE Active=1 "
                                                        end if
                                                        strSQL = strSQL & "ORDER BY SortOrderID, ClientIndexName"
                                                        rsEntry.CursorLocation = 3
                                                        rsEntry.open strSQL, cnWS
                                                        Set rsEntry.ActiveConnection = Nothing
                                                        
                                                        do While NOT rsEntry.EOF                        
%>
                                                                <option value="<%=rsEntry("ClientIndexID")%>" <%if curIndexID=rsEntry("ClientIndexID") then response.write "selected" end if%>><%=rsEntry("ClientIndexName")%></option>
<%
                                                                rsEntry.MoveNext
                                                        loop
                                                        rsEntry.close
%>
                                                </select>&nbsp;&nbsp;
                                                
                                                <% if request.form("optIndex")<>"0" AND request.form("optIndex")<>"" then %>
                                                        Index Value:&nbsp;
                                                        <select name="optIndexValue">
                                                                <option value="0">All Values</option>
<%
                                                                strSQL = "SELECT ClientIndexValueID, ClientIndexValueName FROM tblClientIndexValue WHERE (ClientIndexID = " & curIndexID & ") AND Active=1 ORDER BY ClientIndexValueName "
                                                                rsEntry.CursorLocation = 3
                                                                rsEntry.open strSQL, cnWS
                                                                Set rsEntry.ActiveConnection = Nothing
                                                                
                                                                do While NOT rsEntry.EOF                        
%>
                                                                        <option value="<%=rsEntry("ClientIndexValueID")%>" <%if curIndexValueID=rsEntry("ClientIndexValueID") then response.write "selected" end if%>><%=rsEntry("ClientIndexValueName")%></option>
<%
                                                                        rsEntry.MoveNext
                                                                loop
                                                                rsEntry.close
                                                                
                                                                if request.form("optIncUnassigned")="on" then
%>
                                                                        <option value="-1" <% if request.form("optIndexValue")="-1" then response.write " selected" end if %>>Unassigned</option>
<%
                                                                end if
%>
                                                        
                                                        </select>
                                                <% end if %>
                                                
                                                &nbsp;&nbsp;Home&nbsp;<%=xssStr(allHotWords(8))%>:&nbsp;<select name="optSaleLoc" <% if session("numlocations") > 1 and session("userLoc") <> 0 and not ap_view_all_locs then response.write "disabled" end if %>>
                                                <option value="0" <% if cLoc=0 then response.write "selected" end if %>>All</option>
                                                <!-- CB 3/12/2009 <option value="98" <%if cLoc=98 then response.write "selected" end if%>>Online Store</option> -->
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
                                                        document.frmParameter.optSaleLoc.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' +" <%=jsEscDouble(allHotWords(8))%>s";
                                                </script>
                                                &nbsp;&nbsp;
                                                
                                                &nbsp;&nbsp;View By:&nbsp;
                                                <select name="optOrderBy">
                                                        <option value="1" <% if request.form("optOrderBy")="1" then response.write " selected" end if %>>Alphabetical</option>
                                                <% if curIndexValueID=0 then 'summary view %>
                                                        <option value="2" <% if request.form("optOrderBy")="2" then response.write " selected" end if %>>Descending</option>
                                                <% end if %>
                                                </select>
                                                <%
                                                if ss_CRMContactLogs then
                                                        strSQL = "SELECT TrainerID, TrLastName, TrFirstName, DisplayName FROM TRAINERS WHERE (Rep = 1 OR Rep2=1 OR Rep3=1 OR Rep4=1 OR Rep5=1 OR Rep6=1) AND (Active = 1) AND ([Delete] = 0) "
                                                        strSQL = strSQL & " ORDER BY "
                                                        strSQL = strSQL & GetTrnOrderBy()
                                                        rsEntry.CursorLocation = 3
                                                        rsEntry.open strSQL, cnWS
                                                        Set rsEntry.ActiveConnection = Nothing
                                                        if NOT rsEntry.EOF then
                                                                %>
                                                                &nbsp;&nbsp;<%=xssStr(allHotWords(108))%>:&nbsp;
                                                                <select name="optRep"  <%if NOT ap_CLT_REP_VIEW_ALL then response.write "disabled" end if%>>
                                                                        <option value="0">All</option>
                                                                        <option value="-1" <% if frmOptRep = "-1" then response.write " selected" end if %>>Unassigned</option>
                                                                        <%                                                      
                                                                        do while NOT rsEntry.EOF
                                                                                %>
                                                                                <option value="<%=rsEntry("TrainerID")%>" <%if CSTR(rsEntry("TrainerID"))=frmOptRep then response.write "selected" end if%>><%=FmtTrnNameNew(rsEntry, true)%></option>
                                                                                <%
                                                                                rsEntry.MoveNext
                                                                        loop
                                                                        %>
                                                                </select>
                                                                <script type="text/javascript">document.frmParameter.optRep.options[0].text = '<%=jsEscSingle(allHotWords(149))%>' + " <%=jsEscDouble(allHotWords(108))%>s";</script>
                                                                <%
                                                        end if  'Reps in System
                                                        rsEntry.close
                                                end if  'CRM
                                                %>
                                                <br />
                                                &nbsp;&nbsp;Include Inactive Indexes:
                                                <input type="checkbox" name="optIncInactiveIndex" <% if request.form("optIncInactiveIndex")="on" then response.write " checked" end if %> onClick="document.frmParameter.submit();">
                                                &nbsp;&nbsp;Include Unassigned <%=session("ClientHW")%>s:
                                                <input type="checkbox" name="optIncUnassigned" <% if request.form("optIncUnassigned")="on" then response.write " checked" end if %> onClick="document.frmParameter.submit();">
                                                &nbsp;&nbsp;Include Inactive <%=session("ClientHW")%>s:
                                                <input type="checkbox" name="optIncInactive" <% if request.form("optIncInactive")="on" then response.write " checked" end if %>>
                                                &nbsp;&nbsp;<% taggingFilter %>&nbsp;&nbsp;
                                                <br />
                                                <input type="button" name="Button" value="Generate" onClick="genReport();" <% if curIndexID=0 then response.write " disabled" end if %>>
                                                <% if curIndexID<>0 then %>
                                                <span class="icon-button" style="vertical-align: middle;" title="Export to Excel" ><a onClick="exportReport();" ><img src="<%= contentUrl("/asp/adm/images/export-to-excel-20px.png") %>" /></a></span>
                                                <% else %>
                                                <span style="vertical-align: middle;" title="Export to Excel" ><img src="<%= contentUrl("/asp/adm/images/export-to-excel-grey-20px.png") %>" /></span>
                                                <% end if %>
                                                <% if not Session("Pass") OR Session("Admin")="false" OR NOT validAccessPriv("RPT_TAG") then
                                                 else 
                                                        ' show disabled buttons if no index is selected
                                                        if curIndexID=0 then %>
                                                                <span id="spTagAddGray" title="Tag <%=session("ClientHW")%>s (Add)" style="vertical-align: middle;" ><img src="<%= contentUrl("/asp/adm/images/tag-add-grey-20px.png") %>" /></span>
                                                                <span id="spTagNewGray" title="Tag <%=session("ClientHW")%>s (New)" style="vertical-align: middle;" ><img src="<%= contentUrl("/asp/adm/images/icon_tag_new_gray.png") %>" /></span>
                                                        <% else
                                                                taggingButtons("frmParameter")
                                                        end if 
                                                end if %>
                                                <% savingButtons "frmParameter", "Client Indexes" %>
                                                </b>&nbsp;&nbsp;
                                                </td>
                                                </tr>
                                                
                                                </form>
                                                
                                        </table>                        
                                        </td>
                                        </tr>
                                        <tr> 
                                        <td valign="top" id="clientIndexesGenTag" class="mainTextBig"> 
                                        
                                        <table class="mainText" width="100%" cellspacing="0" style="margin: 0 auto;">
                                                <tr>
                                                <td class="mainTextBig" colspan="2" valign="top">&nbsp;</td>
                                                </tr>
                                                <tr > 
                                                <td class="mainTextBig" colspan="2" valign="top">
<% 
                end if                  'end of frmExpreport value check before /head line        
        
                setRowColors "#F2F2F2", "#FAFAFA"
        
                if request.form("frmGenReport")="true" then 
                        if request.form("frmTagClients")="true" then
                                
                                ' tagging sql
                                if curIndexValueID>0 then ' detail
                                        
                                        strSQL = "SELECT CLIENTS.ClientID "
                                        strSQL = strSQL & "FROM tblClientIndexData INNER JOIN CLIENTS ON tblClientIndexData.ClientID = CLIENTS.ClientID "
                                        if request.form("optFilterTagged")="on" then
                                                strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
                                                if session("mVarUserID")<>"" then
                                                        strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
                                                end if
                                                strSQL = strSQL & " ) "
                                        end if
                                        strSQL = strSQL & "WHERE (ClientIndexValueID = " & curIndexValueID & ") AND CLIENTS.IsSystem=0 "
                                        
                                        
                                elseif curIndexValueID<0 then ' detail - unassigned clients only
																	strSQL = "SELECT CLIENTS.ClientID "
																	strSQL = strSQL & "FROM CLIENTS LEFT OUTER JOIN "
																	strSQL = strSQL & "(SELECT tblClientIndexData.ClientID, tblClientIndexValue.ClientIndexValueID, tblClientIndexValue.ClientIndexValueName, tblClientIndexValue.Active "
																	strSQL = strSQL & "FROM tblClientIndexData INNER JOIN tblClientIndexValue ON tblClientIndexValue.ClientIndexValueID = tblClientIndexData.ClientIndexValueID "
																	strSQL = strSQL & "WHERE tblClientIndexValue.ClientIndexID = " & curIndexID & " AND (tblClientIndexValue.Active = 1)) CID ON CID.ClientID = CLIENTS.ClientID "
																	if request.form("optFilterTagged")="on" then
																					strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
																	end if
																	strSQL = strSQL & "WHERE (CID.ClientIndexValueID IS NULL) AND CLIENTS.IsSystem=0 "
																	if request.form("optFilterTagged")="on" then
																					if session("mvaruserID")<>"" then
																									strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
																					else
																									strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
																					end if
																	end if
																	if cLoc<>0 then
																					strSQL = strSQL & "AND (CLIENTS.HomeStudio = " & cLoc & " OR CLIENTS.HomeStudio=0) "
																	end if
																	if NOT request.form("optIncInactive")="on" then
																					strSQL = strSQL & "AND CLIENTS.Inactive = 0 "
																	end if

																	if frmOptRep<>"0" and frmOptRep<>"" then
																					if frmOptRep="-1" then  'unassigned
																									strSQL = strSQL & " AND (Clients.RepID Is Null) "
																					else
																									strSQL = strSQL & " AND (Clients.RepID=" & frmOptRep & " OR Clients.RepID2=" & frmOptRep & " OR Clients.RepID3=" & frmOptRep & " OR Clients.RepID4=" & frmOptRep & " OR Clients.RepID5=" & frmOptRep & " OR Clients.RepID6=" & frmOptRep & ")"
																					end if
																	end if      
                                else ' summary *****
                                        strSQL = "SELECT CLIENTS.ClientID "
                                        strSQL = strSQL & "FROM CLIENTS LEFT OUTER JOIN "
                                                strSQL = strSQL & "(SELECT tblClientIndexData.ClientID, tblClientIndexValue.ClientIndexValueID, tblClientIndexValue.ClientIndexValueName, tblClientIndexValue.Active "
                                                strSQL = strSQL & "FROM tblClientIndexData INNER JOIN tblClientIndexValue ON tblClientIndexValue.ClientIndexValueID = tblClientIndexData.ClientIndexValueID "
                                                strSQL = strSQL & "WHERE tblClientIndexValue.ClientIndexID = " & curIndexID & " AND (tblClientIndexValue.Active = 1)) CID ON CID.ClientID = CLIENTS.ClientID "
                                        if request.form("optFilterTagged")="on" then
                                                strSQL = strSQL & " INNER JOIN tblClientTag ON (tblClientTag.ClientID = CLIENTS.ClientID "
                                                if session("mVarUserID")<>"" then
                                                        strSQL = strSQL & " AND tblClientTag.smodeID = " & session("mVarUserID")
                                                end if
                                                strSQL = strSQL & " ) "
                                        end if
                                        strSQL = strSQL & "WHERE CLIENTS.IsSystem=0 "
                                        if NOT request.form("optIncUnassigned")="on" then
                                                strSQL = strSQL & "AND (CID.ClientIndexValueID IS NOT NULL) "
                                        end if


                                end if
                                
                                ' all versions share these
                                if cLoc<>0 then
                                        strSQL = strSQL & "AND (CLIENTS.HomeStudio = " & cLoc & " OR CLIENTS.HomeStudio=0) "
                                end if
                                if NOT request.form("optIncInactive")="on" then
                                        strSQL = strSQL & "AND CLIENTS.Inactive = 0 "
                                end if

                                if frmOptRep<>"0" and frmOptRep<>"" then
                                        if frmOptRep="-1" then  'unassigned
                                                strSQL = strSQL & " AND (Clients.RepID Is Null) "
                                        else
                                                strSQL = strSQL & " AND (Clients.RepID=" & frmOptRep & " OR Clients.RepID2=" & frmOptRep & " OR Clients.RepID3=" & frmOptRep & " OR Clients.RepID4=" & frmOptRep & " OR Clients.RepID5=" & frmOptRep & " OR Clients.RepID6=" & frmOptRep & ")"
                                        end if
                                end if
                                
                                
                               response.write debugSQL(strSQL, "SQL")
                                'response.end
                                'rsEntry.CursorLocation = 3
                                'rsEntry.open strSQL, cnWS
                                'Set rsEntry.ActiveConnection = Nothing
        
                                
                                if request.form("frmTagClientsNew")="true" then
                                        clearAndTagQuery(strSQL)
                                else
                                        tagQuery(strSQL)
                                end if
                                
                                'rsEntry.close
                        end if  'End Tag Clients
        
        
                        if request.form("frmExpReport")="true" then
                                Dim stFilename
                                
                                stFilename = "attachment; filename=Client Index Report.xls"
                                Response.ContentType = "application/vnd.ms-excel" 
                                Response.AddHeader "Content-Disposition", stFilename 
                        end if
                        
                        showHeader = "false"
                        
                        ' variables
                        totalClients = 0
                        totalUnassigned = 0
                        percent=0
                        
                        ' client index sql - detail view
                        if curIndexValueID>0 then
                        
                                strSQL = "SELECT CLIENTS.ClientID, CLIENTS.LastName, CLIENTS.FirstName "
                                strSQL = strSQL & "FROM tblClientIndexData INNER JOIN CLIENTS ON tblClientIndexData.ClientID = CLIENTS.ClientID "
                                if request.form("optFilterTagged")="on" then
                                        strSQL = strSQL & "INNER JOIN tblClientTag ON tblClientIndexData.ClientID = tblClientTag.clientID "
                                end if
                                strSQL = strSQL & "WHERE (ClientIndexValueID = " & curIndexValueID & ") AND CLIENTS.IsSystem=0 "
                                if request.form("optFilterTagged")="on" then
                                        if session("mvaruserID")<>"" then
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
                                        else
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
                                        end if
                                end if
                                if cLoc<>0 then
                                        strSQL = strSQL & "AND (CLIENTS.HomeStudio = " & cLoc & " OR CLIENTS.HomeStudio=0) "
                                end if
                                if NOT request.form("optIncInactive")="on" then
                                        strSQL = strSQL & "AND CLIENTS.Inactive = 0 "
                                end if

                                if frmOptRep<>"0" and frmOptRep<>"" then
                                        if frmOptRep="-1" then  'unassigned
                                                strSQL = strSQL & " AND (Clients.RepID Is Null) "
                                        else
                                                strSQL = strSQL & " AND (Clients.RepID=" & frmOptRep & " OR Clients.RepID2=" & frmOptRep & " OR Clients.RepID3=" & frmOptRep & " OR Clients.RepID4=" & frmOptRep & " OR Clients.RepID5=" & frmOptRep & " OR Clients.RepID6=" & frmOptRep & ")"
                                        end if
                                end if
                                if request.form("optOrderBy")="1" then 'alphebetical
                                        strSQL = strSQL & "ORDER BY LastName, FirstName ASC "
                                end if
                                
                        elseif curIndexValueID<0 then ' detail - unassigned clients only

                                strSQL = "SELECT CLIENTS.ClientID, CLIENTS.LastName, CLIENTS.FirstName "
                                strSQL = strSQL & "FROM CLIENTS LEFT OUTER JOIN "

                'CB 6/15/09
                                'strSQL = strSQL & "(SELECT tblClientIndexValue.ClientIndexID, tblClientIndexData.ClientID "
                                'strSQL = strSQL & "FROM tblClientIndexValue LEFT OUTER JOIN tblClientIndexData ON tblClientIndexValue.ClientIndexValueID = tblClientIndexData.ClientIndexValueID "
                'strSQL = strSQL & "WHERE (tblClientIndexValue.ClientIndexID = " & curIndexID & ")) ClientIndex ON CLIENTS.ClientID = ClientIndex.ClientID "
                                strSQL = strSQL & "(SELECT tblClientIndexData.ClientID, tblClientIndexValue.ClientIndexValueID, tblClientIndexValue.ClientIndexValueName, tblClientIndexValue.Active "
                                strSQL = strSQL & "FROM tblClientIndexData INNER JOIN tblClientIndexValue ON tblClientIndexValue.ClientIndexValueID = tblClientIndexData.ClientIndexValueID "
                                strSQL = strSQL & "WHERE tblClientIndexValue.ClientIndexID = " & curIndexID & " AND (tblClientIndexValue.Active = 1)) CID ON CID.ClientID = CLIENTS.ClientID "

                                if request.form("optFilterTagged")="on" then
                                        strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
                                end if

                'CB 6/15/09
                                'strSQL = strSQL & "WHERE (ClientIndex.ClientID IS NULL) AND CLIENTS.IsSystem=0 "
                                strSQL = strSQL & "WHERE (CID.ClientIndexValueID IS NULL) AND CLIENTS.IsSystem=0 "
                                
                                if request.form("optFilterTagged")="on" then
                                        if session("mvaruserID")<>"" then
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
                                        else
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
                                        end if
                                end if
                                if cLoc<>0 then
                                        strSQL = strSQL & "AND (CLIENTS.HomeStudio = " & cLoc & " OR CLIENTS.HomeStudio=0) "
                                end if
                                if NOT request.form("optIncInactive")="on" then
                                        strSQL = strSQL & "AND CLIENTS.Inactive = 0 "
                                end if

                                if frmOptRep<>"0" and frmOptRep<>"" then
                                        if frmOptRep="-1" then  'unassigned
                                                strSQL = strSQL & " AND (Clients.RepID Is Null) "
                                        else
                                                strSQL = strSQL & " AND (Clients.RepID=" & frmOptRep & " OR Clients.RepID2=" & frmOptRep & " OR Clients.RepID3=" & frmOptRep & " OR Clients.RepID4=" & frmOptRep & " OR Clients.RepID5=" & frmOptRep & " OR Clients.RepID6=" & frmOptRep & ")"
                                        end if
                                end if
                                if request.form("optOrderBy")="1" then 'alphebetical
                                        strSQL = strSQL & "ORDER BY LastName, FirstName ASC "
                                end if
                        
                        else ' summary view (all values)
                                
                                strSQL = "SELECT CID.ClientIndexValueID, CASE WHEN CID.ClientIndexValueName IS NULL THEN 'Unassigned' ELSE CID.ClientIndexValueName END AS ClientIndexValueName, COUNT(*) AS NumClients "
                                strSQL = strSQL & "FROM CLIENTS LEFT OUTER JOIN "
                                        strSQL = strSQL & "(SELECT tblClientIndexData.ClientID, tblClientIndexValue.ClientIndexValueID, tblClientIndexValue.ClientIndexValueName, tblClientIndexValue.Active "
                                        strSQL = strSQL & "FROM tblClientIndexData INNER JOIN tblClientIndexValue ON tblClientIndexValue.ClientIndexValueID = tblClientIndexData.ClientIndexValueID "
                                        strSQL = strSQL & "WHERE tblClientIndexValue.ClientIndexID = " & curIndexID & " AND (tblClientIndexValue.Active = 1)) CID ON CID.ClientID = CLIENTS.ClientID "
                                if request.form("optFilterTagged")="on" then
                                        strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
                                end if
                                strSQL = strSQL & "WHERE CLIENTS.IsSystem=0 "
                                if NOT request.form("optIncUnassigned")="on" then
                                        strSQL = strSQL & "AND (CID.ClientIndexValueID IS NOT NULL) "
                                end if
                                if request.form("optFilterTagged")="on" then
                                        if session("mvaruserID")<>"" then
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
                                        else
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
                                        end if
                                end if
                                if cLoc<>0 then
                                        strSQL = strSQL & "AND (CLIENTS.HomeStudio = " & cLoc & " OR CLIENTS.HomeStudio=0) "
                                end if
                                if NOT request.form("optIncInactive")="on" then
                                        strSQL = strSQL & "AND CLIENTS.Inactive = 0 "
                                end if

                                if frmOptRep<>"0" and frmOptRep<>"" then
                                        if frmOptRep="-1" then  'unassigned
                                                strSQL = strSQL & " AND (Clients.RepID Is Null) "
                                        else
                                                strSQL = strSQL & " AND (Clients.RepID=" & frmOptRep & " OR Clients.RepID2=" & frmOptRep & " OR Clients.RepID3=" & frmOptRep & " OR Clients.RepID4=" & frmOptRep & " OR Clients.RepID5=" & frmOptRep & " OR Clients.RepID6=" & frmOptRep & ")"
                                        end if
                                end if
                                strSQL = strSQL & "GROUP BY CID.ClientIndexValueID, CID.ClientIndexValueName "
                                if request.form("optOrderBy")="1" then 'alphebetical
                                        strSQL = strSQL & "ORDER BY CID.ClientIndexValueName ASC "
                                elseif request.form("optOrderBy")="2" then 'descending
                                        strSQL = strSQL & "ORDER BY NumClients DESC, CID.ClientIndexValueName ASC "
                                end if
                                
                        end if ' end detail vs summary
                        
                       response.write debugSQL(strSQL, "SQL")
                        'response.end
                        rsEntry.CursorLocation = 3
                        rsEntry.open strSQL, cnWS
                        Set rsEntry.ActiveConnection = Nothing
                        
                        ' find total number with selected index - required in summary only
                        if curIndexValueID=0 then
                                strSQL = "SELECT COUNT(*) AS TotalClients "
                                strSQL = strSQL & "FROM tblClientIndexValue INNER JOIN tblClientIndexData ON tblClientIndexValue.ClientIndexValueID = tblClientIndexData.ClientIndexValueID INNER JOIN CLIENTS ON CLIENTS.ClientID = tblClientIndexData.ClientID "
                                if request.form("optFilterTagged")="on" then
                                        strSQL = strSQL & "INNER JOIN tblClientTag ON tblClientIndexData.ClientID = tblClientTag.clientID "
                                end if
                                strSQL = strSQL & "WHERE (tblClientIndexValue.Active = 1) AND CLIENTS.IsSystem=0 "
                                if request.form("optFilterTagged")="on" then
                                        if session("mvaruserID")<>"" then
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
                                        else
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
                                        end if
                                end if
                                if cLoc<>0 then
                                        strSQL = strSQL & "AND (CLIENTS.HomeStudio=" & cLoc & " OR CLIENTS.HomeStudio=0) "
                                end if
                                if NOT request.form("optIncInactive")="on" then
                                        strSQL = strSQL & "AND CLIENTS.Inactive = 0 "
                                end if

                                if frmOptRep<>"0" and frmOptRep<>"" then
                                        if frmOptRep="-1" then  'unassigned
                                                strSQL = strSQL & " AND (Clients.RepID Is Null) "
                                        else
                                                strSQL = strSQL & " AND (Clients.RepID=" & frmOptRep & " OR Clients.RepID2=" & frmOptRep & " OR Clients.RepID3=" & frmOptRep & " OR Clients.RepID4=" & frmOptRep & " OR Clients.RepID5=" & frmOptRep & " OR Clients.RepID6=" & frmOptRep & ")"
                                        end if
                                end if
                                strSQL = strSQL & "GROUP BY tblClientIndexValue.ClientIndexID HAVING (tblClientIndexValue.ClientIndexID = " & curIndexID & ")"
                                
                               response.write debugSQL(strSQL, "SQL")
                                rsEntry2.CursorLocation = 3
                                rsEntry2.open strSQL, cnWS
                                Set rsEntry2.ActiveConnection = Nothing
                                
                                if NOT rsEntry2.EOF then
                                        totalClients = rsEntry2("TotalClients")
                                else
                                        totalClients = 0
                                end if
                                
                                rsEntry2.close
                                
                        end if 'end find total - summary view only
                        
                        ' eventually replace with total clients *******************************************************
                        ' count number of unassigned if the option is on AND it is summary view
                        if request.form("optIncUnassigned")="on" AND curIndexValueID=0 then
                                strSQL = "SELECT COUNT(*) AS TotalUnassigned "
                                        strSQL = strSQL & "FROM CLIENTS LEFT OUTER JOIN "
                                        strSQL = strSQL & "(SELECT tblClientIndexData.ClientID, tblClientIndexValue.ClientIndexValueID, tblClientIndexValue.ClientIndexValueName, tblClientIndexValue.Active "
                                        strSQL = strSQL & "FROM tblClientIndexData INNER JOIN tblClientIndexValue ON tblClientIndexValue.ClientIndexValueID = tblClientIndexData.ClientIndexValueID "
                                        strSQL = strSQL & "WHERE tblClientIndexValue.ClientIndexID = " & curIndexID & " AND (tblClientIndexValue.Active = 1)) CID ON CID.ClientID = CLIENTS.ClientID "

                                if request.form("optFilterTagged")="on" then
                                        strSQL = strSQL & "INNER JOIN tblClientTag ON CLIENTS.ClientID = tblClientTag.clientID "
                                end if

                                strSQL = strSQL & "WHERE (CID.ClientIndexValueID IS NULL) "
                                if NOT request.form("optIncInactive")="on" then
                                        strSQL = strSQL & "AND CLIENTS.Inactive = 0 AND CLIENTS.IsSystem=0 "
                                end if

                                if request.form("optFilterTagged")="on" then
                                        if session("mvaruserID")<>"" then
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = " & session("mvaruserID") & ") " 
                                        else
                                                strSQL = strSQL & "AND (tblClientTag.smodeID = 0) " 
                                        end if
                                end if

                                if cLoc<>0 then
                                        strSQL = strSQL & "AND (CLIENTS.HomeStudio = " & cLoc & " OR CLIENTS.HomeStudio=0) "
                                end if

                                if frmOptRep<>"0" and frmOptRep<>"" then
                                        if frmOptRep="-1" then  'unassigned
                                                strSQL = strSQL & " AND (Clients.RepID Is Null) "
                                        else
                                                strSQL = strSQL & " AND (Clients.RepID=" & frmOptRep & " OR Clients.RepID2=" & frmOptRep & " OR Clients.RepID3=" & frmOptRep & " OR Clients.RepID4=" & frmOptRep & " OR Clients.RepID5=" & frmOptRep & " OR Clients.RepID6=" & frmOptRep & ")"
                                        end if
                                end if
                                
                               response.write debugSQL(strSQL, "SQL")
                                'response.end
                                rsEntry2.CursorLocation = 3
                                rsEntry2.open strSQL, cnWS
                                Set rsEntry2.ActiveConnection = Nothing
                                
                                if NOT rsEntry2.EOF then
                                        totalUnassigned = rsEntry2("TotalUnassigned")
                                else
                                        totalUnassigned = 0
                                end if
                                
                        end if
        
%>
                        <table class="mainText"  cellspacing="0" style="margin: 0 auto;">
<% 
                        if curIndexValueID<>0 then ' if detail view
                                totalClients = 0 ' reset totalClients to zero - detail counts the rows
                        
                                if NOT rsEntry.EOF then                 'EOF
                                        do while NOT rsEntry.EOF
                                                If showHeader = "false" then %>
                                                        <tr>
                                                                <td colspan="4">&nbsp;</td>
                                                        </tr>
                                                        <tr class="right">
                                                                <td width="25%"><strong><%=session("ClientHW")%></strong></td>
                                                                <td width="25%"></td>
                                                                <td width="50%">&nbsp;</td>
                                                        </tr>
                                                        <% if NOT request.form("frmExpReport")="true" then %>
                                                                <tr height="2">
                                                                        <td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
                                                                </tr>
<% 
                                                        end if  
                                                end if ' end showHeader=false
                                                showHeader = "true"
%>
                                                        <tr class="right" style="background-color:<%=getRowColor(true)%>;">
                                                                <td nowrap><a href="main_info.asp?id=<%=rsEntry("ClientID")%>&fl=true" title="Click Here to View <%=session("ClientHW")%> Information"><%=rsEntry("FirstName")%>&nbsp;<%=rsEntry("LastName")%></a></td>
                                                                <td><a href="adm_clt_conlog.asp?clientid=<%=rsEntry("ClientID")%>">[View Logs]</a></td>
                                                                <td>&nbsp;</td>
                                                        </tr>
<%              
                                                totalClients = totalClients + 1
                                                rsEntry.MoveNext
                                        loop 'end rsEntry.EOF loop
%>
                                        <% if NOT request.form("frmExpReport")="true" then %>
                                                <tr height="2">
                                                        <td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
                                                </tr>
                                        <% end if %>
                                                <tr class="right">
                                                        <td>Total <%=session("ClientHW")%>s:</td>
                                                        <td><%=totalClients%></td>
                                                </tr>
                                                <tr>
                                                        <td colspan="4">&nbsp;</td>
                                                </tr>
<% 
                                end if  ' end if detail eof
                                
                        else ' display summary view
                        
                                ' add unassigned to the total if marked to include unassigned
                                if request.form("optIncUnassigned")="on" then                                           
                                        totalClients = totalClients + totalUnassigned
                                end if
                        
                                if NOT rsEntry.EOF then                 'EOF
                                        do while NOT rsEntry.EOF
                                                If showHeader = "false" then %>
                                                        <tr>
                                                                <td colspan="4">&nbsp;</td>
                                                        </tr>
                                                        <tr>
                                                                <th width="20%"><strong>Index Value</strong></th>
                                                                <th width="20%"><strong>&nbsp;&nbsp;<%=session("ClientHW")%>s</strong></th>
                                                                <th width="10%"><strong>&nbsp;&nbsp;Percent</strong></th>
                                                                <th width="50%">&nbsp;</th>
                                                        </tr>
                                                        <% if NOT request.form("frmExpReport")="true" then %>
                                                                <tr height="2">
                                                                        <td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
                                                                </tr>
<% 
                                                        end if  
                                                end if ' end showHeader=false
                                                showHeader = "true"
                                                        
                                                ' calculate percent of total
                                                if totalClients = 0 then
                                                        percent = 0
                                                else
                                                        percent=FormatNumber((rsEntry("NumClients")/totalClients), 3) * 100
                                                end if
                                                
                                                if barColor = session("pageColor4") then
                                                        barColor = session("pageColor3")
                                                elseif barColor = session("pageColor3") then
                                                        barColor = session("pageColor2")
                                                else
                                                        barColor = session("pageColor4")
                                                end if
%>
                                                        <tr class="right" style="background-color:<%=getRowColor(true)%>;">
                                                                <td><%=rsEntry("ClientIndexValueName")%></td>
                                                                <td><%=rsEntry("NumClients")%></td>
                                                                <td><%=percent%>%</td>
                                                                <td>
                                                                <% if NOT request.form("frmExpReport")="true" then%>
                                                                        <table align="left" style="background-color:<%=barcolor%>;" width="<%if totalClients = 0 then response.write "0" else response.write (rsEntry("NumClients")/totalClients)*360 end if%>">
                                                                                <tr height="9">
                                                                                        <td><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>"></td>
                                                                                </tr>
                                                                        </table>
                                                                <%end if%>
                                                                </td>
                                                        </tr>
<%              
                                                rsEntry.MoveNext
                                        loop 'end rsEntry.EOF loop
%>
                                        
                                        <% if NOT request.form("frmExpReport")="true" then %>
                                                <tr height="2">
                                                        <td colspan="4" style="background-color:#666666;"><% if NOT request.form("frmExpReport")="true" then %><img src="<%= contentUrl("/asp/adm/images/trans.gif") %>" height="1" width="100%"><% end if %></td>
                                                </tr>
                                        <% end if %>
                                                <tr class="right">
                                                        <td>Total <%=session("ClientHW")%>s:</td>
                                                        <td><%=totalClients%></td>
                                                </tr>
                                                <tr>
                                                        <td colspan="4">&nbsp;</td>
                                                </tr>
<% 
                                end if  ' end if eof
                        end if ' end display summary vs detail view
%>
                        </table>
<%
                        rsEntry.close
                        set rsEntry = nothing
                end if          'end of generate report if statement
%>
                                        </table></table>
                        </td>
                </tr>
</table>
<% pageEnd %>
<!-- #include file="post.asp" -->

<%
        
end if
%>
