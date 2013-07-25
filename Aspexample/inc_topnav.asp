<%
'CB This file is legacy and not used anymore fine to remove any dependencies
'   Replaced by inc_topnav_div.asp
%>
<%

dim phraseDictionary
set phraseDictionary = LoadPhrases("ConsumermodetopnavigationPage", 9)

%>

<!---  <%=session("TabID")%> -->
<!-- #include file="adm/inc_chk_ss.asp" -->
<table cellspacing="0">
  <tr valign="bottom">
	<td class="mainTopNav"><%=DisplayPhrase(phraseDictionary,"Goto")%>:&nbsp;</td>
	<td>
	<table cellspacing="0">
		<tr height="14" valign="top">

<%	dim SchDisplayName, curSchDisplayName, rsTabs, rsRewTabs
	
	set rsTabs = Server.CreateObject("ADODB.Recordset")
	set rsRewTabs = Server.CreateObject("ADODB.Recordset")

	strSQL = " SELECT tblTabs.TabID, tblTypeGroupTab.TabID AS FK, tblTabs.TabNameCltMode as TabName, tblTabs.tabData, tblTabs.SortOrder, tblTabs.wsReservation, tblTabs.wsEnrollment, "
	strSQL = strSQL & " tblTabs.wsAppointment, tblTabs.wsResource, tblTabs.wsMedia, tblTabs.wsLink FROM tblTabs LEFT OUTER JOIN "
	strSQL = strSQL & " tblTypeGroupTab ON tblTabs.TabID = tblTypeGroupTab.TabID WHERE showCltMode = 1 AND "
	strSQL = strSQL & " tblTabs.TabID <> 4 AND "
	strSQL = strSQL & " ((tblTabs.isSystem = 1) OR (wsLink = 1) OR "
	strSQL = strSQL & " ((tblTypeGroupTab.TypeGroupID IS NOT NULL) AND ( (1 = 0) OR (wsMedia = 1) "
	if session("wsType")="res" or session("wsType")="prem" then ' classes
		strSQL = strSQL & " OR (wsReservation = 1) OR (wsEnrollment = 1) "
	end if
	if (session("wsType")="appt" or session("wsType")="prem") AND true then
		strSQL = strSQL & " OR (wsAppointment = 1) "
	end if
	if session("useResrcResv") OR session("useResrcAppt") then
		strSQL = strSQL & " OR (wsResource = 1) "
	end if
	strSQL = strSQL & "))) GROUP BY tblTabs.TabID, tblTypeGroupTab.TabID, tblTabs.TabNameCltMode, tblTabs.tabData, tblTabs.SortOrder, tblTabs.wsReservation, tblTabs.wsEnrollment, tblTabs.wsAppointment,  "
	strSQL = strSQL & " tblTabs.wsResource, tblTabs.wsMedia, tblTabs.wsLink ORDER BY tblTabs.SortOrder  "
'response.write debugSQL(strSQL, "SQL")
	rsTabs.CursorLocation = 3
	rsTabs.open strSQL, cnWS
	
	if NOT rsTabs.EOF then
		do while NOT rsTabs.EOF
			select case rsTabs("TabID")
			'JM-51_2567
			case 100 ' rewards tab
				if checkStudioSetting("tblGenOpts", "RewardsOptInRequired") AND session("mvarUserID")<> "" then
				    strSQL = "SELECT RewardsOptIn FROM CLIENTS WHERE CLIENTS.ClientID=" & session("mvarUserID")
				    rsRewTabs.CursorLocation = 3
				    rsRewTabs.open strSQL, cnWS
				    Set rsRewTabs.ActiveConnection = Nothing
			
                    if NOT rsRewTabs.EOF then
				        if NOT rsRewTabs("RewardsOptIn") then
    				        if session("tabID")=CSTR(rsTabs("TabID")) then %>
					<td class="center-ch">
						<div id="ncnav">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=session("pageColor4")%>;"><tr><td>
							<a href="javascript: document.search2.tabID.value = <%=rsTabs("TabID")%>; document.search2.submit();" class="mainTopNav"><b><span style="color:#FFFFFF;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></b></a>
							</td></tr></table>
						</div>			
					</td>
	    		        <%	else %>
					<td class="center-ch">
						<div id="ncnav<%=rsTabs("TabID")%>">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=topBGClr%>;"><tr><td>
							<a href="javascript:goTo('rewards', <%=rsTabs("TabID")%>);" class="mainTopNav"><b><span style="color:<%=session("pageColor")%>;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></b></a>
							</td></tr></table>
						</div>			
					</td>					
		    	        <%  end if
		    	        end if'if NOT rsEntry("RewardsOptIn") then	
		            end if   'if NOT rsRewTabs.EOF
			   rsRewTabs.Close
			   end if 'checkstudiosetting
			case 2 ' Client Info Tab - special case
				if session("tabID")=CSTR(rsTabs("TabID")) then  %>
					<td>
						<div id="ncnav">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=session("pageColor4")%>;"><tr><td>
							<a href="javascript: document.search2.tabID.value = <%=rsTabs("TabID")%>; document.search2.submit();" class="mainTopNav"><b><span style="color:#FFFFFF;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></b></a>
							</td></tr></table>
						</div>			
					</td>
			<%	else %>
					<td>
						<div id="ncnav<%=rsTabs("TabID")%>">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=topBGClr%>;"><tr><td>
							<a href="javascript:goTo('info', <%=rsTabs("TabID")%>);" class="mainTopNav"><b><span style="color:<%=session("pageColor")%>;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></b></a>
							</td></tr></table>
						</div>
					</td>
			<%	end if 
			case 3, 4 ' Retail Tab - special case
				if session("onlineProds") OR session("PartnersEnabled") then
					if NOT ss_ClientModeLockTGRemoveBuy then
						if session("tabID")=CSTR(rsTabs("TabID")) then %>
					<td>
						<div id="ncnav">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=session("pageColor4")%>;"><tr><td>
							<a href="javascript: document.search2.tabID.value = <%=rsTabs("TabID")%>; document.search2.submit();" class="mainTopNav"><b><span style="color:#FFFFFF;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></b></a>
							</td></tr></table>
						</div>
					</td>
					<%	else %>
					<td>
						<div id="ncnav<%=rsTabs("TabID")%>">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=topBGClr%>;"><tr><td>					
							<a href="javascript:goTo('shop', <%=rsTabs("TabID")%>);" class="mainTopNav"><b><span style="color:<%=session("pageColor")%>;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></b></a>
							</td></tr></table>
						</div>			
					</td>
					<%	end if
					end if
				end if
				
				
			case 6 ' help tab
				if session("tabID")=CSTR(rsTabs("TabID")) then %>
					<td class="center-ch">
						<div id="ncnav">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=session("pageColor4")%>;"><tr><td>
							<a href="javascript: document.search2.tabID.value = <%=rsTabs("TabID")%>; document.search2.submit();" class="mainTopNav"><b><span style="color:#FFFFFF;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></b></a>
							</td></tr></table>
						</div>			
					</td>
			<%	else %>
					<td class="center-ch">
						<div id="ncnav<%=rsTabs("TabID")%>">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=topBGClr%>;"><tr><td>
							<a href="javascript:goTo('help', <%=rsTabs("TabID")%>);" class="mainTopNav"><b><span style="color:<%=session("pageColor")%>;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></b></a>
							</td></tr></table>
						</div>			
					</td>
			<%	end if
			case else ' normal tab
				if rsTabs("wsReservation") then
					'strSQL = "SELECT wsReservation FROM tblTypeGroup WHERE Active=1 AND wsReservation=1 AND wsEnrollment=0"
					strSQL = "SELECT TOP 1 tblTypeGroup.wsReservation, tblClasses.ClassActive FROM tblTypeGroup INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & rsTabs("tabID") & " INNER JOIN tblClassDescriptions ON tblTypeGroup.TypeGroupID = tblClassDescriptions.ClassPayment INNER JOIN tblClasses ON tblClassDescriptions.ClassDescriptionID = tblClasses.DescriptionID WHERE (tblTypeGroup.Active = 1) AND (tblTypeGroup.wsReservation = 1) AND (tblTypeGroup.wsEnrollment = 0) AND (tblClasses.ClassActive = 1) AND (tblClasses.ClassDateEnd >= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ")"
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing
			
					if NOT rsEntry.EOF then		

						if session("tabID")=CSTR(rsTabs("TabID")) then %>
						<td class="center-ch">
							<div id="ncnav">
								<table height="14" class="mainText" cellspacing="0" style="background-color:<%=session("pageColor4")%>;"><tr><td>
								<a href="javascript: document.search2.tabID.value = <%=rsTabs("TabID")%>; document.search2.submit();" class="mainTopNav"><span style="color:#FFFFFF;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></a>
								</td></tr></table>
							</div>
						</td>
					<%	else %>
						<td class="center-ch">
							<div id="ncnav<%=rsTabs("TabID")%>">
								<table height="14" class="mainText" cellspacing="0" style="background-color:<%=topBGClr%>;"><tr><td>
								<a href="javascript:goTo('class', <%=rsTabs("TabID")%>);" class="mainTopNav"><span style="color:<%=session("pageColor")%>;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></a>
								</td></tr></table>
							</div>
						</td>
					<%	end if 
					end if
					rsEntry.close
				elseif rsTabs("wsEnrollment") then
					'strSQL = "SELECT wsReservation FROM tblTypeGroup WHERE Active=1 AND wsReservation=1 AND wsEnrollment=0"
					strSQL = "SELECT TOP 1 tblTypeGroup.wsReservation, tblClasses.ClassActive FROM tblTypeGroup INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & rsTabs("tabID") & " INNER JOIN tblClassDescriptions ON tblTypeGroup.TypeGroupID = tblClassDescriptions.ClassPayment INNER JOIN tblClasses ON tblClassDescriptions.ClassDescriptionID = tblClasses.DescriptionID WHERE (tblTypeGroup.Active = 1) AND (tblTypeGroup.wsEnrollment = 1) AND (tblClasses.ClassActive = 1) AND (tblClasses.ClassDateEnd >= " & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ")"
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing
			
					if NOT rsEntry.EOF then		
						if session("tabID")=CSTR(rsTabs("TabID")) then %>
						<td class="center-ch">
							<div id="ncnav">
								<table height="14" class="mainText" cellspacing="0" style="background-color:<%=session("pageColor4")%>;"><tr><td>
								<a href="javascript: document.search2.tabID.value = <%=rsTabs("TabID")%>; document.search2.submit();" class="mainTopNav"><span style="color:#FFFFFF;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></a>
								</td></tr></table>
							</div>
						</td>
					<% 	else %>
						<td class="center-ch">
							<div id="ncnav<%=rsTabs("TabID")%>">
								<table height="14" class="mainText" cellspacing="0" style="background-color:<%=topBGClr%>;"><tr><td>
								<a href="javascript:goTo('enroll', <%=rsTabs("TabID")%>);" class="mainTopNav"><span style="color:<%=session("pageColor")%>;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%></span>&nbsp;&nbsp;</a>
								</td></tr></table>
							</div> 
						</td>
					<%	end if
					end if
					rsEntry.close
				elseif rsTabs("wsAppointment") then
					strSQL = "SELECT wsAppointment FROM tblTypeGroup INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & rsTabs("tabID") & " WHERE Active=1 AND wsAppointment=1 AND wsDisable=0"
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing
			
					if NOT rsEntry.EOF then
						if session("tabID")=CSTR(rsTabs("TabID")) then %>
						<td class="center-ch">
							<div id="ncnav">
								<table height="14" class="mainText" cellspacing="0" style="background-color:<%=session("pageColor4")%>;"><tr><td>
								<a href="javascript: document.search2.tabID.value = <%=rsTabs("TabID")%>; document.search2.submit();" class="mainTopNav"><span style="color:#FFFFFF;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></a>
								</td></tr></table>
							</div>
						</td>
					<%	else %>
						<td>
							<div id="ncnav<%=rsTabs("TabID")%>">
								<table height="14" class="mainText" cellspacing="0" style="background-color:<%=topBGClr%>;"><tr><td>
								<a href="javascript:goTo('appts', <%=rsTabs("TabID")%>);" class="mainTopNav"><span style="color:<%=session("pageColor")%>;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></a>
								</td></tr></table>
							</div>
						</td>
					<%	end if 
					end if
					rsEntry.close
				elseif rsTabs("wsMedia") then
					if session("tabID")=CSTR(rsTabs("TabID")) then %>
					<td class="center-ch">
						<div id="ncnav">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=session("pageColor4")%>;"><tr><td>
							<a href="javascript: document.search2.tabID.value = <%=rsTabs("TabID")%>; document.search2.submit();" class="mainTopNav"><span style="color:#FFFFFF;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></a>
							</td></tr></table>
						</div>
					</td>
				<%	else %>
					<td>
						<div id="ncnav<%=rsTabs("TabID")%>">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=topBGClr%>;"><tr><td>
							<a href="javascript:goTo('media', <%=rsTabs("TabID")%>);" class="mainTopNav"><span style="color:<%=session("pageColor")%>;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></a>
							</td></tr></table>
						</div>
					</td>
				<%	end if 
				elseif rsTabs("wsLink") then %>
					<td>
						<div id="ncnav<%=rsTabs("TabID")%>">
							<table height="14" class="mainText" cellspacing="0" style="background-color:<%=topBGClr%>;"><tr><td>
							<a href="<%=rsTabs("tabData")%>" class="mainTopNav" target="_blank"><span style="color:<%=session("pageColor")%>;">&nbsp;&nbsp;<%if ss_UpperCaseTabs then response.write ucase(rsTabs("TabName")) else response.write rsTabs("TabName") end if%>&nbsp;&nbsp;</span></a>
							</td></tr></table>
						</div>
					</td>
					
					
			<%	end if
			end select
		
			rsTabs.MoveNext
		loop
	end if %>
					</tr>
				</table>
			</td>
		</tr>
	</table>
<DIV></DIV>
