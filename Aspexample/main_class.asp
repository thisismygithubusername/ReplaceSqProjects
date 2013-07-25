<%@ CodePage=65001 %>
<% Option Explicit %>
<!-- #include file="init.asp" -->
<!-- #include file="json/JSON.asp" -->
<%
'Dim SessionFarm
'set SessionFarm = Server.CreateObject("SessionFarm.SFSession")
response.charset="utf-8"
%>
		
	<!-- #include file="inc_internet_guest.asp" -->
    	<!-- #include file="adm/inc_hotword.asp" -->
	<!-- #include file="inc_tinymcesetup.asp" -->
<%
    session("pageID")="_class"
    if isNum(request.form("tabID")) then
	    session("TabID") = request.form("tabID")
    elseif isNum(request.querystring("tabID")) then
	    session("TabID") = request.querystring("tabID")
    end if
    
	Dim rsEntry
	set rsEntry = Server.CreateObject("ADODB.Recordset")

	Dim SchOpenDOMCloseDate : SchOpenDOMCloseDate = CDATE("1/1/2200")
	if Session("SchOpenDOMCloseDate")<>"" then
		SchOpenDOMCloseDate = CDATE(Session("SchOpenDOMCloseDate"))
	end if

	'Start STEVE 1/30
	if request.QueryString("classid") <> "" AND request.QueryString("classid") <> "0" then
		dim tclassID, tclassDate, ttypeID, tloc
		tclassID = request.QueryString("classid")
		tclassDate = request.QueryString("date")
		ttypeID = request.QueryString("tg")
		tloc = request.QueryString("loc")

		'CB 1/27/09 - Check for Free Class
		if session("Pass") AND isNum(tclassID) then
			strSQL = "SELECT [Free] FROM tblClasses WHERE ClassID=" & tclassID
			rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
			if NOT rsEntry.EOF then
				if rsEntry("Free") and NOT ss_reserveForOtherClient  then
					response.Redirect("res_debf.asp?classID=" & tclassID & "&classDate=" & tclassDate & "&tg=" & ttypeID & "&clsLoc=" & tloc)
				end if
			end if 
			rsEntry.close
		end if

        if DateValue(tclassDate) <= SchOpenDOMCloseDate then
	    	response.Redirect("res_a.asp?classID=" & tclassID & "&classDate=" & tclassDate & "&tg=" & ttypeID & "&clsLoc=" & tloc)
        end if
	end if
	'End Steve 1/30

	dim varLink, classStartTime, classEndTime, firstOne, allTGsDisabled, intDays, strDays, tmpDate, StartOnMonday, strSubFootnotes, subCount, strClassOpenSlots, ssClassShowAsst1, ssClassShowAsst2, ss_ForwardToFirstClass
    dim cResVT, curTG, curTrn, curTime, prtDayOne, prtDayTwo, rowcount, curTypeName, tmpBioStr, viewAllDayOn, tmpintdays, ShowCltCount, GroupByVT, ssAnchor, cView, SchHrsClosed, SchHrsStartTime, SchHrsEndTime
    Dim schSDate, schEDate, clscount, cont, weekDaystblStr, ssShowRsrc, numTGs, tmpCount, SameDaySchCutoff, ss_SubsInRed, ss_CltModeHideSubNotes, isMember, curLevel, schNumLocs, ss_CltModeAllowDupRes, autoFwdDate
    dim  ss_CltModeHideSignUpNoCurrent, cltCurrentTypegroups, ss_reserveForOtherClient, ss_ResvSchAllViewDay, ss_cltPackageSharing, filterByClsSch
    dim tmpFilter : tmpFilter = -999
    dim prevFilterByClsSch : prevFilterByClsSch = -1
    dim prevFilterByClsSch2 : prevFilterByClsSch2 = -2
    dim filterMode : filterMode = 0 ' 0 = reset, 1 = asc, 2 = desc
    Dim ssSchedOffset()

	strSQL = "SELECT tblResvOpts.ResvSchAllViewDay, tblResvOpts.ClientShowClsCount, tblResvOpts.UseWeeklyClassView, tblResvOpts.ClassSchStartOnMon, tblResvOpts.anchorToDay, tblResvOpts.CltModeShowRsrc, "&_
             "tblResvOpts.SubsInRed,tblResvOpts.CltModeHideSubNotes, tblResvOpts.CltModeAllowDupRes, tblResvOpts.CltModeHideSignUpNoCurrent, tblGenOpts.cltPackageSharing, tblGenOpts.CustSchHours, "&_
             "tblGenOpts.CustSchHrsStart, tblGenOpts.CustSchHrsEnd, tblGenOpts.SameDaySchCutoff, tblGenOpts.ClassShowAsst1, tblGenOpts.ClassShowAsst2, tblGenOpts.reserveForOtherClient, tblResvOpts.ForwardToFirstClass "&_
             "FROM tblResvOpts "&_
             "INNER JOIN tblGenOpts ON tblResvOpts.StudioID = tblGenOpts.StudioID "&_
             "WHERE tblResvOpts.StudioID=" & session("StudioID")
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing

	ss_ResvSchAllViewDay = rsEntry("ResvSchAllViewDay")
	ShowCltCount = rsEntry("ClientShowClsCount")
	GroupByVT = rsEntry("UseWeeklyClassView")
	StartOnMonday = rsEntry("ClassSchStartOnMon")
	ssAnchor = rsEntry("anchorToDay")
	ss_SubsInRed = rsEntry("SubsInRed")
	ss_CltModeHideSubNotes = rsEntry("CltModeHideSubNotes")
	ss_CltModeAllowDupRes = rsEntry("CltModeAllowDupRes")
	ss_CltModeHideSignUpNoCurrent = rsEntry("CltModeHideSignUpNoCurrent")
	if session("useResrcResv") then
		ssShowRsrc = rsEntry("CltModeShowRsrc")
	else
		ssShowRsrc = false
	end if
	ssClassShowAsst1=rsEntry("ClassShowAsst1")
	ssClassShowAsst2=rsEntry("ClassShowAsst2")
	ss_reserveForOtherClient = rsEntry("reserveForOtherClient")
	ss_ForwardToFirstClass = CBOOL(rsEntry("ForwardToFirstClass"))
    ss_cltPackageSharing = rsEntry("cltPackageSharing")
	SchHrsClosed = rsEntry("CustSchHours")
	if SchHrsClosed then	'If NOT 24 Hr scheduling
		''Validate not null and stime < end time
		if NOT isNULL(rsEntry("CustSchHrsStart")) AND NOT isNULL(rsEntry("CustSchHrsEnd")) then
			if TimeValue(rsEntry("CustSchHrsStart")) < TimeValue(rsEntry("CustSchHrsEnd")) then
				'''Check if cur time is closed
				if TimeValue(DateAdd("n", Session("tzOffset"),Time)) < TimeValue(rsEntry("CustSchHrsStart")) OR TimeValue(DateAdd("n", Session("tzOffset"),Time)) > TimeValue(rsEntry("CustSchHrsEnd")) then
					'' Schedule is currently closed
					SchHrsStartTime = rsEntry("CustSchHrsStart")
					SchHrsEndTime = rsEntry("CustSchHrsEnd")
				else
					SchHrsClosed = false
				end if
			else
				SchHrsClosed = false
			end if
		else 'Null Time Val
			SchHrsClosed = false
		end if
	end if

	if isNULL(rsEntry("SameDaySchCutoff")) then
		SameDaySchCutoff = false
	else
		if TimeValue(DateAdd("n", Session("tzOffset"),Time)) > TimeValue(rsEntry("SameDaySchCutoff")) then
			SameDaySchCutoff = true
		else
			SameDaySchCutoff = false
		end if
	end if
	rsEntry.close

   'MB setting location is moved to inc_loc_list.asp

    if request.form("optVT")<>"" then
	    cResVT = request.form("optVT")
    elseif isNum(request.querystring("vt")) then
	    cResVT = request.querystring("vt")
	    if cResVT<>"0" then
		    strSQL = "SELECT TypeName FROM tblVisitTypes WHERE TypeID=" & cResVT
		    rsEntry.CursorLocation = 3
		    rsEntry.open strSQL, cnWS
		    Set rsEntry.ActiveConnection = Nothing
		    if NOT rsEntry.EOF then
			    cResVT = rsEntry("TypeName")
		    else
			    cResVT="0"
		    end if			
		    rsEntry.close
	    end if
    else
	    cResVT="0"
    end if

    ' Pull out the current Class Level
    if isNum(request.querystring("lvl")) then
	    curLevel = CLNG(request.querystring("lvl"))
    elseif isNum(request.form("optLevel")) then
	    curLevel = CLNG(request.form("optLevel"))
    else
	    curLevel=0
    end if

    ' Pull out the current TG
    if isNum(request.form("optTG")) then
	    curTG = CINT(request.form("optTG"))
    elseif isNum(request.querystring("tg")) then
	    curTG = CINT(request.querystring("tg"))
    else
	    curTG=0
    end if

    ' Pull out the instructor
    if isNum(request.form("optInstructor")) then
	    curTrn = CLNG(request.form("optInstructor"))
    elseif isNum(request.querystring("trn")) then
	    curTrn = CLNG(request.querystring("trn"))
    else
	    curTrn = 0
    end if

    ' Pull out the header filter sort
    if isNum(request.Form("filterByClsSch")) then
        tmpFilter = CLNG(request.Form("filterByClsSch"))
        dim orderFilter : orderFilter = ""
        dim tmpPrevFilter : tmpPrevFilter = -1
        dim tmpPrevFilter2 : tmpPrevFilter2 = -2

        if isNum(request.Form("prevFilterByClsSch")) then
            tmpPrevFilter = CLNG(request.Form("prevFilterByClsSch")) 
        end if

        if isNum(request.Form("prevFilterByClsSch2")) then
            tmpPrevFilter2 = CLNG(request.Form("prevFilterByClsSch2"))
        end if

        filterMode = 1 ' 1 asc
        if tmpFilter = tmpPrevFilter and tmpPrevFilter <> tmpPrevFilter2 then 'descending sort
            orderFilter = " desc"
            filterMode = 2
        end if
        if tmpFilter = tmpPrevFilter and tmpPrevFilter = tmpPrevFilter2 then 'reset sort
            tmpFilter = -999
            filterMode = 0
        end if

        prevFilterByClsSch = tmpFilter
        prevFilterByClsSch2 = tmpPrevFilter

        select case tmpFilter
            case 1 ' start time
                filterByClsSch = "tblClassSch.StartTime" & orderFilter
            case 3 ' class 
                filterByClsSch = "tblClassDescriptions.ClassName" & orderFilter
            case 4 ' trainer 
                filterByClsSch = "TRAINERS.TrFirstName" & orderFilter & ", TRAINERS.TrLastName " & orderFilter
            case 5 ' assistant 1 
                filterByClsSch = "Asst1FirstName" & orderFilter & ", Asst1LastName " & orderFilter
            case 6 ' assistant 2
                filterByClsSch =  "Asst2FirstName" & orderFilter & ", Asst2LastName " & orderFilter
            case 7 ' location
                filterByClsSch = "Location.LocationName" & orderFilter
            case 8 ' resource
                filterByClsSch = "tblResources.ResourceName" & orderFilter
            case 9 ' duration
                filterByClsSch = "DurationHours " & orderFilter & ", DurationMinutes" & orderFilter
            case else ' default
                filterByClsSch = "tblClassSch.StartTime, tblClassDescriptions.ClassName"
        end select
    else
        filterByClsSch = "tblClassSch.StartTime, tblClassDescriptions.ClassName"
    end if
    ' end filter sort
%>
	<!-- #include file="inc_i18n.asp" -->
	<!-- #include file="adm/inc_chk_holiday.asp" -->
	<% if session("CR_Memberships") <> 0 then %>
		<!-- #include file="inc_dbconn_regions.asp" -->
		<!-- #include file="inc_dbconn_wsMaster.asp" -->
		<!-- #include file="adm/inc_masterclients_util.asp" -->
	<% end if %>
	<!-- #include file="adm/inc_acct_balance.asp" -->
	<!-- #include file="adm/inc_crypt.asp" -->
	<!-- #include file="inc_loading.asp" -->

<% ' Load the phrases for this page and set them to local variables
dim phraseDictionary
set phraseDictionary = LoadPhrases("ConsumermodeclassschedulePage", 11)

dim welcomeDictionary
set welcomeDictionary = LoadPhrases("ConsumermodeloginPage", 40)

dim sTime

%>
	<!-- #include file="pre.asp" -->
	<!-- #include file="frame_bottom.asp" -->
    <!-- #include file="inc_date_ctrl.asp" -->
    <!-- #include file="inc_fixed_bar.asp" -->
    <!-- #include file="adm/inc_hotword.asp" -->
    <%= js(array("mb", "MBS")) %>
    <!-- begin client alerts -->
    <%
    'client alert context vars
    focusFrmElement = ""
    cltAlertList = setClientAlertsList(session("mvarUserID"))
    %>
    <!-- #include file="inc_ajax.asp" -->
    <!-- #include file="adm/inc_alert_js.asp" -->
    <!-- end client alerts  -->

    <%= js(array("calendar" & dateFormatCode, "plugins/jquery.selectBox", "VCC2", "main_class")) %>
    <%= css(array("main_class", "day_week_navigator", "jquery.selectBox")) %>
    <style type="text/css">
        #classSchedule-mainTable tr.headerRow td.header
        {
            color: <%=session("pageColor3")%>;
        }
    </style>
    <%
     'MB setting location is moved to inc_loc_list.asp
     
    %>
    <script type="text/javascript">
        $(function() {
            <%
            dim tmpAnchorDate : tmpAnchorDate = Date
            if StartOnMonday then
                do while weekday(tmpAnchorDate) > 1 ' get to this sunday
                    tmpAnchorDate = DateAdd("y", 1, tmpAnchorDate)
                loop
            else
                do while weekday(tmpAnchorDate) < 7 ' get to this saturday
                    tmpAnchorDate = DateAdd("y", 1, tmpAnchorDate)
                loop
            end if
            
            if ssAnchor and cFrmDate < tmpAnchorDate then %>
                 // anchor to the date
                if ($('#an<%=WeekDay(cFrmDate)%>').offset() != null)
                {
                    var scrollValue = $('#an<%=WeekDay(cFrmDate)%>').offset().top - $('#classSchedule-header').offset().top - 30;
					<% if cltAlertList="" then %>
	                $(document).scrollTop(scrollValue);
					<% end if %>
                }
            <% end if%>
        });
    </script>

    <!-- #include file="adm/inc_alert_content.asp" -->
	<%
	dim ss_HideClsTGs, ss_HideClsVTs, ss_showLevels, lockLoc, boolShowCur, trnName
	set rsEntry = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT Studios.StudioURL, tblGenOpts.StudioLinkTab, tblGenOpts.HideCltHelp, tblGenOpts.HideCltForgotPwd, tblGenOpts.UpperCaseTabs, tblAppearance.LogoHeight, tblAppearance.LogoWidth, "&_
             "tblAppearance.topBGColor, tblGenOpts.ClientModeLockTGRemoveBuy, tblGenOpts.CltModeSigup, tblGenOpts.CltModeSigupExisting, tblGenOpts.BBenabled, tblGenOpts.ClientModeLockLoc, tblResvOpts.hideClsTGs, "&_
             "tblResvOpts.hideClsVTs, tblResvOpts.UseClassLeves, tblResvOpts.HideClsLevels "&_
             "FROM Studios "&_
             "INNER JOIN tblGenOpts ON Studios.StudioID = tblGenOpts.StudioID "&_
             "INNER JOIN tblAppearance ON Studios.StudioID = tblAppearance.StudioID "&_
             "INNER JOIN tblResvOpts ON Studios.StudioID = tblResvOpts.StudioID "&_
             "INNER JOIN tblApptOpts ON Studios.StudioID = tblApptOpts.StudioID "&_
             "WHERE (Studios.StudioID = " & session("StudioID") & ")"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing

	'top_class specific
	ss_HideClsTGs = rsEntry("hideClsTGs")
	ss_HideClsVTs = rsEntry("hideClsVTs")
	if rsEntry("UseClassLeves") AND NOT rsEntry("HideClsLevels") then
		ss_showLevels = true
	else
		ss_showLevels = false
	end if
	rsEntry.close
%>
  
    <script type="text/javascript">
        function updateView() {
	        document.search2.action = "main_class.asp";
	        document.search2.submit();
	        <% if Session("Pass") then %>
		        parent.mainFrame.focus();
	        <% else %>
		        loginFocus();
	        <% end if %>
        }
    </script>
      
    <form id="ClassScheduleSearch2Form" name="search2" method="post" action="main_class.asp">
      <input type="hidden" name="pageNum" value="1" />
      <input type="hidden" name="requiredtxtUserName" value="" />
      <input type="hidden" name="requiredtxtPassword" value="" />
      <input type="hidden" name="optForwardingLink" value="" />
	  <input type="hidden" name="optRememberMe" value="" />
	  <input type="hidden" name="tabID" value="<%=session("tabID")%>" />
	  <input type="hidden" id="optView" name="optView" class="bf-filter" value="<%=ifb_view%>" />
	  <input type="hidden" id="useClassLogic" name="useClassLogic" value="" />
      <input type="hidden" id="filterByClsSch" name="filterByClsSch" value="" />
      <input type="hidden" id="prevFilterByClsSch" name="prevFilterByClsSch" value="<%=prevFilterByClsSch %>" />
      <input type="hidden" id="prevFilterByClsSch2" name="prevFilterByClsSch2" value="<%=prevFilterByClsSch2 %>" />

    <div class="wrapperTop">
      <div class="wrapperTopInner">
        <div class="pageTop">
            <div class="pageTopLeft">
                &nbsp; 
            </div>
            <div class="pageTopRight">
                &nbsp; 
            </div>
            <h1 class="classScheduleHeader"><%=DisplayPhrase(phraseDictionary,"Classschedule")%></h1> 
            <div id="dateControls">
                <div class="leftSide">
                    <% dayAndWeekControls %>
                </div>
                <div class="rightSide">
                    <input id="txtDate" type="text" name="txtDate" size="10" maxlength="14" value="<%=FmtDateShort(topDate)%>" class="date bf-filter" />
                    &nbsp;
                    <script type="text/javascript">
                        var cal1 = new tcal({'id': 'dayInfo', 'formname':'search2', 'controlname':'txtDate'});
                        cal1.a_tpl.yearscroll = true;
                        $('#txtDate').change(<%if InStr(Request.ServerVariables("HTTP_USER_AGENT"), "MSIE" )>0 then response.write "chkDate" else response.write "subForm" end if%>);
						$('#tcalico_dayInfo').attr('onclick', '$("#txtDate").focus()');
                    </script>
                </div>
            </div>
            </div>
        </div>
    </div>

    <div class="fixedHeader">
        <div class="topFilters group">
            <div class="topFiltersInner group">
                <div class="filterList">
                <!-- #include file="inc_loc_list.asp" -->
                <% setLocationSessionVar(true)
				
				if not ss_HideClsTGs then
					boolShowCur = false
					''Check for > 1 TG otherwise don't display TG selector
					strSQL = "SELECT DISTINCT tblTypeGroup.TypeGroupID, tblTypeGroup.TypeGroup from tblClasses, tblClassDescriptions, tblTypeGroup "
					if session("tabID")<>"" AND isNumeric(session("tabID")) then
						strSQL = strSQL & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
					end if 
					strSQL = strSQL & "  WHERE tblClasses.DescriptionID=tblClassDescriptions.ClassDescriptionID AND tblClassDescriptions.ClassPayment=tblTypeGroup.TypeGroupID "
					strSQL = strSQL & " AND wsReservation=1 AND wsEnrollment=0 AND " & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & " <= tblClasses.ClassDateEnd"
					strSQL = strSQL & " ORDER BY TypeGroup"
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing

					if NOT rsEntry.EOF then
						rsEntry.MoveNext
						if NOT rsEntry.EOF then
							boolShowCur = true
							rsEntry.MoveFirst
						end if
					end if

					if boolShowCur then
%>
                        <select name="optTG" id="optTG" class="bf-filter" onChange="filterClasses(this);">
                          <option value="0"><%=xssStr(allHotWords(516))%></option>
					        <%	Do While NOT rsEntry.EOF %>
                          <option value="<%=rsEntry("TypeGroupID")%>" title="<%=stringIfLengthOverLimit(rsEntry("TypeGroup"),18)%>" <%if curTG=rsEntry("TypeGroupID") then response.write "selected" end if%>>
														<%=truncateString(rsEntry("TypeGroup"),18)%>
													</option>
						        <%	rsEntry.MoveNext
							        Loop %>							
                        </select>
				        <script type="text/javascript">
				            document.search2.optTG.options[0].text = "<%=jsEsc(allHotWords(516))%>";
				        </script>
<%
					end if	''1 or less hide
					rsEntry.close
				end if	''showTG selector

				if not ss_HideClsVTs then
					boolShowCur = false
					
					' Class Type drop down ----------------------------------------------------------------------
            	    ' New, optimized query provided by Chet, implemented by CWS
                    strSQL = "SELECT DISTINCT tblVisitTypes.SortOrder, tblVisitTypes.TypeName, tblVisitTypes.TypeID "&_
                             "FROM tblVisitTypes "&_
                             "INNER JOIN tblClassDescriptions ON tblVisitTypes.TypeID = tblClassDescriptions.VisitTypeID "

                    if session("tabID")<>"" AND isNumeric(session("tabID")) then
                        strSQL = strSQL & "INNER JOIN tblTypeGroupTab ON tblVisitTypes.Typegroup = tblTypeGroupTab.TypeGroupID AND tblTypeGroupTab.TabID = '" & session("tabID") & "' "
                    end if

                    strSQL = strSQL & "INNER JOIN tblClasses ON tblClassDescriptions.ClassDescriptionID = tblClasses.DescriptionID "&_
                             "WHERE (tblClasses.ClassDateEnd >= " & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & ") "&_
                             "AND (tblClasses.ClassDateStart<= " & DateSep & DateAdd("yyyy", 1, DateAdd("n", Session("tzOffset"),Now)) & DateSep & ") "

                    ' Filter by location
	                if session("curLocation") <> "0" then
	                    strSQL = strSQL & "AND (tblClasses.LocationID = " & session("curLocation") & ") "
	                end if
                	
                    ' Filter by program
	                if curTG <> 0 then
	                    strSQL = strSQL & "AND (tblVisitTypes.TypeGroup = " & curTG & ") "
	                end if

                    strSQL = strSQL & "ORDER BY tblVisitTypes.SortOrder, tblVisitTypes.TypeName "
                    rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing

					' Check if we have more than 1 class type
					boolShowCur = rsEntry.RecordCount > 1

					if boolShowCur then
%>							
                        <select name="optVT" id="optVT" class="bf-filter" onChange="filterClasses(this);">
                          <option value="0"><%=xssStr(allHotWords(518))%></option>
<%
							Do While NOT rsEntry.EOF
%>
                                <option value="<%=rsEntry("TypeName")%>" title="<%=stringIfLengthOverLimit(rsEntry("TypeName"),18)%>" <%if cResVT = rsEntry("TypeName") then response.write "selected" end if%>>
																	<%=truncateString(rsEntry("TypeName"),18)%>
																</option>
<%
								rsEntry.MoveNext
							Loop
%>
                        </select>
				        <script type="text/javascript">
				            document.search2.optVT.options[0].text = "<%=xssStr(allHotWords(518))%>";
				        </script>
<%
					end if ''show VT >1
					rsEntry.close
				end if	''SS show VT

			    if ss_showLevels then
					boolShowCur = false
					''check for more than 1 Class Level
					strSQL = "SELECT DISTINCT tblClassLevels.LevelID, tblClassLevels.LevelName FROM tblTypeGroup "
					if session("tabID")<>"" AND isNumeric(session("tabID")) then
						strSQL = strSQL & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
					end if 
					strSQL = strSQL & "  INNER JOIN tblClassDescriptions ON tblTypeGroup.TypeGroupID = tblClassDescriptions.ClassPayment "&_
                             "INNER JOIN tblClasses ON tblClassDescriptions.ClassDescriptionID = tblClasses.DescriptionID "&_
                             "INNER JOIN tblClassLevels ON tblClassDescriptions.LevelID = tblClassLevels.LevelID "&_
					         "WHERE (" & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & " <= tblClasses.ClassDateEnd) AND (tblClasses.ClassActive = 1) "&_
                             "AND (tblTypeGroup.wsReservation = 1) AND (tblTypeGroup.wsEnrollment = 0) ORDER BY tblClassLevels.LevelName"
					rsEntry.CursorLocation = 3
					rsEntry.open strSQL, cnWS
					Set rsEntry.ActiveConnection = Nothing

					if NOT rsEntry.EOF then
						rsEntry.MoveNext
						if NOT rsEntry.EOF then
							boolShowCur = true
							rsEntry.MoveFirst
						end if
					end if

				    if boolShowCur then
%>				
                    <select name="optLevel" id="optLevel" class="bf-filter" onChange="filterClasses(this);">
                      <option value="0"><%=allHotWords(761)%></option>
<%							

							Do While NOT rsEntry.EOF
%>
                            <option value="<%=rsEntry("LevelID")%>" title="<%=stringIfLengthOverLimit(rsEntry("LevelName"),18)%>" <%if curLevel=rsEntry("LevelID") then response.write "selected" end if%>>
															<%=truncateString(rsEntry("LevelName"),18)%>
														</option>
<%
								rsEntry.MoveNext
							Loop
%>
                    </select>
				    <script type="text/javascript">
				        document.search2.optLevel.options[0].text = "<%=jsEsc(allHotWords(761))%>";
				    </script>
<%
				end if ''show VT > 1
				rsEntry.close
			end if 'SS Levels
			
			strSQL = "SELECT DISTINCT TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, "&_
                     "CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END AS TrnDisName "&_
			         "FROM tblClassDescriptions "&_
                     "INNER JOIN tblTypeGroup ON tblClassDescriptions.ClassPayment = tblTypeGroup.TypeGroupID  "

			if session("tabID")<>"" AND isNumeric(session("tabID")) then
				strSQL = strSQL & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
			end if 

			strSQL = strSQL & " INNER JOIN tblClassSch ON tblClassDescriptions.ClassDescriptionID = tblClassSch.DescriptionID "&_
                     "INNER JOIN TRAINERS ON tblClassSch.TrainerID = TRAINERS.TrainerID "&_
                     "INNER JOIN tblClasses ON tblClassSch.ClassID = tblClasses.ClassID "&_
			         "WHERE (tblTypeGroup.wsReservation = 1) AND (tblTypeGroup.wsEnrollment = 0) AND (TRAINERS.[Delete] = 0) AND (TRAINERS.ReservationTrn = 1) AND (tblClasses.MaskTrainer = 0) "&_
                     "AND (tblClassSch.ClassDate >=" & DateSep & DateValue(DateAdd("n", Session("tzOffset"),Now)) & DateSep & ") AND (tblClassSch.ClassDate <= " & DateSep & DateValue(DateAdd("yyyy", 1, DateAdd("n", Session("tzOffset"),Now))) & DateSep & ") "&_
			         "ORDER BY " & GetTrnOrderBy()
		    'response.write debugSQL(strSQL, "SQL")

            rsEntry.CursorLocation = 3
			rsEntry.open strSQL, cnWS
			Set rsEntry.ActiveConnection = Nothing
%>
            <select name="optInstructor" id="optInstructor" class="bf-filter" onChange="filterClasses(this);">
                <option value="0"><%=xssStr(allHotWords(519))%></option>
                  <%
				do While NOT rsEntry.EOF
					trnName = FmtTrnNameNew(rsEntry, true)
%>
                  <option value="<%=rsEntry("TrainerID")%>" title="<%=stringIfLengthOverLimit(trnName,24)%>" <%if curTrn=CLNG(rsEntry("TrainerID")) then response.write "selected" end if%>>
										<%=truncateString(trnName,24)%>
									</option>
<%
					rsEntry.MoveNext
				Loop
				rsEntry.close
%>
                </select>
				<script type="text/javascript">
				    document.search2.optInstructor.options[0].text = "<%=jsEsc(allHotWords(519))%>";
				</script>
                </div>
	        </div>
        </div>
    </div>
    </form>
    <div style="clear:both"></div>
   <% pageStart %>
    <table height="100%" width="<%=strPageWidth%>" cellspacing="0" id="classSchedule">
	       <!-- #include file="inc_res_sch.asp" -->
    </table>
<% pageEnd %>
<!-- #include file="post.asp" -->
