<%
%>
<!---  <%=session("TabID")%> -->

<%
dim SchDisplayName, curSchDisplayName, rsTabs
dim ap_APPT_SCH_TAB, ap_VIEW_OWN_SCHEDULE, ap_RETAIL_TAB, ap_RETAIL_OPEN_TICKETS
dim  rsRewTabs

	
set rsTabs = Server.CreateObject("ADODB.Recordset")
set rsRewTabs = Server.CreateObject("ADODB.Recordset")
dim tabType, numTabs, tabName, currentTab, lastTabID, hideRewardsTab 'Req for bizmode tabs
	
	%>

<%if session("Admin")="false" then %>
	<div id="tab-scroll-l" class="tab-scroll-l">&nbsp;</div>
	<div id="tab-scroll-r" class="tab-scroll-r">&nbsp;</div>
	<div id="tab-bar" class="tab-bar">
	<div id="tab-padding">
	<div id="tab-bar-padding-cm">
  
		<table id="tab-table" class="tab-table"><tr>
	<%	

		strSQL = " SELECT tblTabs.TabID, tblTypeGroupTab.TabID AS FK, tblTabs.TabNameCltMode as TabName, tblTabs.tabData, tblTabs.SortOrder, tblTabs.wsReservation, tblTabs.wsEnrollment, "&_
				 " tblTabs.wsAppointment, tblTabs.wsResource, tblTabs.wsMedia, tblTabs.wsLink FROM tblTabs LEFT OUTER JOIN "&_
				 " tblTypeGroupTab ON tblTabs.TabID = tblTypeGroupTab.TabID WHERE showCltMode = 1 AND "&_
				 " tblTabs.TabID <> 4 AND "&_
				 " ((tblTabs.isSystem = 1) OR (wsLink = 1) OR "&_
				 " ((tblTypeGroupTab.TypeGroupID IS NOT NULL) AND ( (1 = 0) OR (wsMedia = 1) "
		if session("wsType")="res" or session("wsType")="prem" then ' classes
			strSQL = strSQL & " OR (wsReservation = 1) OR (wsEnrollment = 1) "
		end if
		if (session("wsType")="appt" or session("wsType")="prem") AND true then
			strSQL = strSQL & " OR (wsAppointment = 1) "
		end if
		if session("useResrcResv") OR session("useResrcAppt") then
			strSQL = strSQL & " OR (wsResource = 1) "
		end if
		strSQL = strSQL & "))) GROUP BY tblTabs.TabID, tblTypeGroupTab.TabID, tblTabs.TabNameCltMode, tblTabs.tabData, tblTabs.SortOrder, tblTabs.wsReservation, tblTabs.wsEnrollment, tblTabs.wsAppointment,  "&_
						  " tblTabs.wsResource, tblTabs.wsMedia, tblTabs.wsLink ORDER BY tblTabs.SortOrder  "
	'response.write debugSQL(strSQL, "SQL")
		rsTabs.CursorLocation = 3
		rsTabs.open strSQL, cnWS
		
		if NOT rsTabs.EOF then
			numTabs =  rsTabs.RecordCount
			currentTab = 0
			lastTabID = -999
			do while NOT rsTabs.EOF
				currentTab = currentTab + 1

				if currentTab = (numTabs - 1) then 'penultimate tab
					rsTabs.MoveLast
					lastTabID = rsTabs("TabID")
					rsTabs.MovePrevious

					hideRewardsTab = true
					if CBOOL(session("Pass")) then
						hideRewardsTab = GetClientRewardOptIn(session("mvarUserID"))
					end if

					if lastTabID = 100 and (not ss_RewardsOptInRequired or hideRewardsTab) then ' rewards tab is hidden so make current tab the last one
						numTabs = numTabs - 1
					end if
				end if

				if ss_UpperCaseTabs then
					tabName =  ucase(rsTabs("TabName"))
				else 
					tabName = rsTabs("TabName")
				end if
				
				
				select case rsTabs("TabID")
				'JM-51_2567
				case 100 ' rewards tab
					if ss_RewardsOptInRequired AND session("mvarUserID")<> "" then
						strSQL = "SELECT RewardsOptIn FROM CLIENTS WHERE CLIENTS.ClientID=" & session("mvarUserID")
						rsRewTabs.CursorLocation = 3
						rsRewTabs.open strSQL, cnWS
						Set rsRewTabs.ActiveConnection = Nothing
			
						if NOT rsRewTabs.EOF then
							if NOT rsRewTabs("RewardsOptIn") then
	    				
									tabType = showTab(rsTabs("TabID"), "rewards", tabName, tabType, false, currentTab, numTabs, false, "") 
		    			
		    				end if'if NOT rsEntry("RewardsOptIn") then	
						end if   'if NOT rsRewTabs.EOF
				   rsRewTabs.Close
				   end if
				case 2 ' Client Info Tab - special case
					tabType = showTab(rsTabs("TabID"), "info", tabName, tabType, false, currentTab, numTabs, false, "") 
					
				case 3, 4 ' Retail Tab - special case
					if session("onlineProds") OR session("PartnersEnabled") then
						if NOT ss_ClientModeLockTGRemoveBuy then
							tabType = showTab(rsTabs("TabID"), "shop", tabName, tabType, false, currentTab, numTabs, false, "") 
						end if
					end if
				case 6 ' help tab
					tabType = showTab(rsTabs("TabID"), "help", tabName, tabType, false, currentTab, numTabs, true, "help/")
				case else ' normal tab
					if rsTabs("wsReservation") then
						strSQL = "SELECT TOP 1 tblTypeGroup.wsReservation, tblClasses.ClassActive "&_
								 "FROM tblTypeGroup "&_
								 "INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & rsTabs("tabID") &_
								 " INNER JOIN tblClassDescriptions ON tblTypeGroup.TypeGroupID = tblClassDescriptions.ClassPayment "&_
								 "INNER JOIN tblClasses ON tblClassDescriptions.ClassDescriptionID = tblClasses.DescriptionID "&_
								 "WHERE (tblTypeGroup.Active = 1) AND (tblTypeGroup.wsReservation = 1) AND (tblTypeGroup.wsEnrollment = 0) AND (tblClasses.ClassActive = 1) "&_
								 "AND (tblClasses.ClassDateEnd >= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ")"
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						Set rsEntry.ActiveConnection = Nothing
			
						if NOT rsEntry.EOF then		

							tabType = showTab(rsTabs("TabID"), "class", tabName, tabType, false, currentTab, numTabs, false, "") 
						end if
						rsEntry.close
					elseif rsTabs("wsEnrollment") then
						strSQL = "SELECT TOP 1 tblTypeGroup.wsReservation, tblClasses.ClassActive "&_
								 "FROM tblTypeGroup "&_
								 "INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & rsTabs("tabID") &_
								 " INNER JOIN tblClassDescriptions ON tblTypeGroup.TypeGroupID = tblClassDescriptions.ClassPayment "&_
								 "INNER JOIN tblClasses ON tblClassDescriptions.ClassDescriptionID = tblClasses.DescriptionID "&_
								 "WHERE (tblTypeGroup.Active = 1) AND (tblTypeGroup.wsEnrollment = 1) AND (tblClasses.ClassActive = 1) "&_
								 "AND (tblClasses.ClassDateEnd >= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ")"
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						Set rsEntry.ActiveConnection = Nothing
			
						if NOT rsEntry.EOF then		
							tabType = showTab(rsTabs("TabID"), "enroll", tabName, tabType, false, currentTab, numTabs, false, "") 
						end if
						rsEntry.close
					elseif rsTabs("wsAppointment") then
						strSQL = "SELECT wsAppointment "&_
								 "FROM tblTypeGroup "&_
								 "INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & rsTabs("tabID") &_
								 " WHERE Active=1 AND wsAppointment=1 AND wsDisable=0"
						rsEntry.CursorLocation = 3
						rsEntry.open strSQL, cnWS
						Set rsEntry.ActiveConnection = Nothing
			
						if NOT rsEntry.EOF then
							tabType = showTab(rsTabs("TabID"), "appts", tabName, tabType, false, currentTab, numTabs, false, "") 
						end if
						rsEntry.close
					elseif rsTabs("wsMedia") then
						tabType = showTab(rsTabs("TabID"), "media", tabName, tabType, false, currentTab, numTabs, false, "") 
					elseif rsTabs("wsLink") then 
						tabType = showTab(rsTabs("TabID"), rsTabs("tabData"), tabName, tabType, true, currentTab, numTabs, false, "") 
					end if
				end select
		
				rsTabs.MoveNext
			loop
		end if %>
		</tr></table>
    </div>
    </div>
	</div>
<%else 'BizMode tabs %>
<div id="tab-scroll-l" class="tab-scroll-l">&nbsp;</div>
<div id="tab-scroll-r" class="tab-scroll-r">&nbsp;</div>
<div id="tab-bar" class="tab-bar tab-bar-biz">
<div id="tab-bar-padding" class="tab-bar-padding">
<table id="tab-table" class="tab-table"><tr>
<%
	
	if Session("Admin")="sa" OR Session("Admin")="owner" then
		ap_APPT_SCH_TAB = true
		ap_VIEW_OWN_SCHEDULE = true
		ap_RETAIL_TAB = true
		ap_RETAIL_OPEN_TICKETS = true
	else
		strSQL = "SELECT APPT_SCH_TAB, VIEW_OWN_SCHEDULE, RETAIL_TAB, RETAIL_OPEN_TICKETS FROM tblAccessPriv WHERE StudioID=" & session("StudioID") & " AND status=N'" & sqlInjectStr(Session("Admin")) & "'"
		rsTabs.CursorLocation = 3
		rsTabs.open strSQL, cnWS
		if NOT rsTabs.EOF then
			ap_APPT_SCH_TAB = rsTabs("APPT_SCH_TAB")
			ap_VIEW_OWN_SCHEDULE = rsTabs("VIEW_OWN_SCHEDULE")
			ap_RETAIL_TAB = rsTabs("RETAIL_TAB")
			ap_RETAIL_OPEN_TICKETS = rsTabs("RETAIL_OPEN_TICKETS")
		end if
		rsTabs.close
	end if

	strSQL = " SELECT tblTabs.TabID, tblTypeGroupTab.TabID AS FK, tblTabs.TabName, tblTabs.SortOrder, tblTabs.wsReservation, tblTabs.wsEnrollment, "
	strSQL = strSQL & " tblTabs.wsAppointment, tblTabs.wsResource, tblTabs.wsMedia, tblTabs.tabData, tblTabs.wsLink FROM tblTabs LEFT OUTER JOIN "
	strSQL = strSQL & " tblTypeGroupTab ON tblTabs.TabID = tblTypeGroupTab.TabID WHERE tblTabs.TabID <> 6 AND tblTabs.showBizMode = 1 AND "
	if ss_FullPOS then
		strSQL = strSQL & " tblTabs.TabID <> 4 AND "
	else
		strSQL = strSQL & " tblTabs.TabID <> 3 AND "
	end if
	strSQL = strSQL & " ((tblTabs.isSystem = 1) OR (tblTabs.wsLink = 1) OR "
	strSQL = strSQL & " ((tblTypeGroupTab.TypeGroupID IS NOT NULL) AND ( (1 = 0) OR (wsMedia = 1) "
	if session("wsType")="res" or session("wsType")="prem" then ' classes
		strSQL = strSQL & " OR (wsReservation = 1) OR (wsEnrollment = 1) "
	end if
	if (session("wsType")="appt" or session("wsType")="prem") AND (ap_APPT_SCH_TAB or ap_VIEW_OWN_SCHEDULE) then
		strSQL = strSQL & " OR (wsAppointment = 1) "
	end if
	if session("useResrcResv") OR session("useResrcAppt") then
		strSQL = strSQL & " OR (wsResource = 1) "
	end if
	strSQL = strSQL & "))) GROUP BY tblTabs.TabID, tblTypeGroupTab.TabID, tblTabs.TabName, tblTabs.SortOrder, tblTabs.wsReservation, tblTabs.wsEnrollment, tblTabs.wsAppointment,  "
	strSQL = strSQL & " tblTabs.wsResource, tblTabs.wsMedia, tblTabs.tabData, tblTabs.wsLink ORDER BY tblTabs.SortOrder  "
	'response.write debugSQL(strSQL, "SQL")
	rsTabs.CursorLocation = 3
	rsTabs.open strSQL, cnWS
	
	
	if NOT rsTabs.EOF then
		do while NOT rsTabs.EOF
			if ss_UpperCaseTabs then
				tabName =  ucase(rsTabs("TabName"))
			else 
				tabName = rsTabs("TabName")
			end if
		
			select case rsTabs("TabID")
			case -1, 99 ' Dash Tab - special case
				tabType = showTab(rsTabs("TabID"), "dash", tabName, tabType, false, 0, 0, false, "") 
			case 98 'Home tab - special case
				'do NOTHING :)
			case 97 'Reports tab - special case
				'do NOTHING :)
			case 1 ' Sign In Tab - special case
				tabType = showTab(rsTabs("TabID"), "signin", tabName, tabType, false, 0, 0, false, "") 
			case 2 ' Client Info Tab - special case
				tabType = showTab(rsTabs("TabID"), "info", tabName, tabType, false, 0, 0, false, "") 
			case 3, 4 ' Retail Tab - special case
				if ap_RETAIL_TAB OR ap_RETAIL_OPEN_TICKETS then
					tabType = showTab(rsTabs("TabID"), "retail", tabName, tabType, false, 0, 0, false, "") 
				end if	'retail tab AP
			case else ' normal tab
				if rsTabs("wsReservation") then
					tabType = showTab(rsTabs("TabID"), "class", tabName, tabType, false, 0, 0, false, "") 
				elseif rsTabs("wsEnrollment") then
					tabType = showTab(rsTabs("TabID"), "enroll", tabName, tabType, false, 0, 0, false, "") 
				elseif rsTabs("wsAppointment") then
					tabType = showTab(rsTabs("TabID"), "appts", tabName, tabType, false, 0, 0, false, "") 
				elseif rsTabs("wsResource") then
					if ss_WSPremSch then
						tabType = showTab(rsTabs("TabID"), "resrc", tabName, tabType, false, 0, 0, false, "") 
					end if
				elseif rsTabs("wsMedia") then
					tabType = showTab(rsTabs("TabID"), "media", tabName, tabType, false, 0, 0, false, "") 
				elseif rsTabs("wsLink") then
					 tabType = showTab(rsTabs("TabID"), rsTabs("tabData"), tabName, tabType, true, 0, 0, false, "") 
				end if
			end select
		
			rsTabs.MoveNext
		loop
	end if
%>
	</tr></table>
</div>
</div>
<%end if %>
