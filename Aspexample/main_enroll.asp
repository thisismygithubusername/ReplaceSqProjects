<%@ CodePage=65001 %> 
<%Option Explicit%>
<!-- #include file="init.asp" -->
<!-- #include file="json/JSON.asp" -->
<%
response.charset="utf-8"
%>

<%
session("pageID")="_enroll"
if isNum(request.form("tabID")) then
	session("TabID") = request.form("tabID")
elseif isNum(request.querystring("tabID")) then
	session("TabID") = request.querystring("tabID")
end if

%>
        <!-- #include file="inc_internet_guest.asp" -->
		<!-- #include file="inc_enroll_info.asp" -->		
		<!-- #include file="inc_tinymcesetup.asp" -->
<%

dim phraseDictionary
set phraseDictionary = LoadPhrases("BusinessconsumermodeworkshopschedulePage", 10)

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
		strSQL = "SELECT [Free] FROM tblClasses WHERE ClassID=" & tclassID & " AND ClassDateStart = ClassDateEnd "
		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
		if NOT rsEntry.EOF then
			if rsEntry("Free") then
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

%>
			<!-- #include file="inc_i18n.asp" -->
            <% if session("CR_Memberships") <> 0 then %>
                <!-- #include file="inc_dbconn_regions.asp" -->
                <!-- #include file="inc_dbconn_wsMaster.asp" -->
                <!-- #include file="adm/inc_masterclients_util.asp" -->
            <% end if %>
			<!-- #include file="adm/inc_acct_balance.asp" -->
			<!-- #include file="adm/inc_crypt.asp" -->
			<!-- #include file="inc_loading.asp" -->
			<!-- #include file="adm/inc_hotword.asp" -->
            <!-- #include file="inc_capacity.asp" -->

<% 
%>
<!-- #include file="pre.asp" -->
<!-- #include file="frame_bottom.asp" -->
<!-- #include file="inc_date_ctrl.asp" -->
<!-- #include file="inc_fixed_bar.asp" -->
<%= js(array("mb", "MBS")) %>
<%
'client alert context vars
focusFrmElement = ""
cltAlertList = setClientAlertsList(session("mvarUserID"))
%>
<!-- #include file="inc_ajax.asp" -->
<!-- #include file="adm/inc_alert_js.asp" -->

<%
Dim cView, SchHrsClosed, SchHrsStartTime, SchHrsEndTime, SameDaySchCutoff, ssCltModeEnrollEndDates, isMember, ss_ShowAsst1, ss_ShowAsst2, ss_EnrollSchedSortBy', roomInClass
dim bookingStr1, bookingStr2, trainerName, dateStr, timeStr, resourceStr, dayHW
Dim ssSchedOffset()
Dim lockLoc : lockLoc = false

	set rsEntry = Server.CreateObject("ADODB.Recordset")
	strSQL = "SELECT tblGenOpts.CustSchHours, tblGenOpts.CustSchHrsStart, tblGenOpts.CustSchHrsEnd, tblGenOpts.SameDaySchCutoff, tblResvOpts.EnrollSchedSortBy, tblResvOpts.CltModeEnrollEndDates, tblGenOpts.EnrollShowAsst1, tblGenOpts.EnrollShowAsst2 FROM tblGenOpts INNER JOIN tblResvOpts ON tblGenOpts.StudioID = tblResvOpts.StudioID WHERE tblGenOpts.StudioID=" & session("StudioID")
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	SchHrsClosed = rsEntry("CustSchHours")
	ssCltModeEnrollEndDates = rsEntry("CltModeEnrollEndDates")
    ss_EnrollSchedSortBy = rsEntry("EnrollSchedSortBy")
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
	ss_ShowAsst1 = rsEntry("EnrollShowAsst1")
	ss_ShowAsst2 = rsEntry("EnrollShowAsst2")
	rsEntry.close
%>

<%= js(array("calendar" & dateFormatCode, "VCC2", "plugins/jquery.selectBox", "main_enroll")) %>
<%= css(array("main_class", "main_enroll", "day_week_navigator", "jquery.selectBox")) %>
<style type="text/css">
			.description, .dates, .times, .resourceName {
				padding: 5px;
			}
</style>
<!--[if lte IE 7]>
	<style type="text/css">
		#main-content {margin-top: 120px;}
	</style>
<![endif]-->
</head>
<body>
<!-- #include file="adm/inc_alert_content.asp" -->
<%
dim ss_HideEnrollTGs, ss_HideEnrollVTs, ss_showLevels, curTrn, curVT, curTG, curLevel, trnName, numTGs, subCount, tmpDate
dim curPageDate, cResVT, cResTG, cTrnID, rowcount, multiDays, calcSDate, first, intCount, tmpcount
dim courseFirstClassID, courseFirstClassDate, curCourseID, cont, ss_AllowOpenEnrollment

strSQL = "SELECT tblGenOpts.ClientModeLockLoc, tblResvOpts.hideEnrollTGs, tblResvOpts.hideEnrollVTs, tblResvOpts.UseClassLeves, tblResvOpts.HideEnrollLevels, tblResvOpts.AllowOpenEnrollment FROM Studios INNER JOIN tblGenOpts ON Studios.StudioID = tblGenOpts.StudioID INNER JOIN tblAppearance ON Studios.StudioID = tblAppearance.StudioID INNER JOIN tblResvOpts ON Studios.StudioID = tblResvOpts.StudioID INNER JOIN tblApptOpts ON Studios.StudioID = tblApptOpts.StudioID WHERE (Studios.StudioID = " & session("StudioID") & ")"
rsEntry.CursorLocation = 3
rsEntry.open strSQL, cnWS
Set rsEntry.ActiveConnection = Nothing

ss_HideEnrollTGs = rsEntry("hideEnrollTGs")
ss_HideEnrollVTs = rsEntry("hideEnrollVTs")
if rsEntry("UseClassLeves") AND NOT rsEntry("HideEnrollLevels") then
        ss_showLevels = true
else
        ss_showLevels = false
end if
ss_AllowOpenEnrollment = rsEntry("AllowOpenEnrollment")
rsEntry.close

' Pull out the current TG
if isNum(request.form("optTG")) then
	curTG = CINT(request.form("optTG"))
elseif isNum(request.querystring("tg")) then
	curTG = CINT(request.querystring("tg"))
else
	curTG=0
end if

' Pull out the current Visit Type
if isNum(request.Form("optVT")) then
	curVT = CLNG(request.Form("optVT"))
elseif isNum(request.QueryString("vt")) then
	curVT = CLNG(request.QueryString("vt"))
else
	curVT=0
end if

' Pull out the instructor
if isNum(request.form("optInstructor")) then
	curTrn = CLNG(request.form("optInstructor"))
elseif isNum(request.querystring("trn")) then
	curTrn = CLNG(request.querystring("trn"))
else
	curTrn = 0
end if

' Pull out the current Class Level
if isNum(request.querystring("lvl")) then
	curLevel = CLNG(request.querystring("lvl"))
elseif isNum(request.form("optLevel")) then
	curLevel = CLNG(request.form("optLevel"))
else
	curLevel=0
end if

%>
    <form name="search2" method="post" target="mainFrame" action="main_enroll.asp">
      <input type="hidden" name="pageNum" value="1" />
      <input type="hidden" name="requiredtxtUserName" value="" />
      <input type="hidden" name="requiredtxtPassword" value="" />
      <input type="hidden" name="optForwardingLink" value="" />
      <input type="hidden" name="optRememberMe" value="" />
      <input type="hidden" name="tabID" value="<%=session("tabID")%>" />
<div class="fixedHeader">
    <div class="topFilters group"> 
      <div class="topFiltersInner filterList group">
                <!-- #include file="inc_loc_list.asp" -->
                <% setLocationSessionVar(true) %>
                <%
                    Dim boolShowTG
                    if ss_hideEnrollTGs then
                    else
                            boolShowTG = false

                                    ''Check for > 1 TG otherwise don't display TG selector
                                    strSQL = "SELECT DISTINCT tbltypeGroup.TypeGroupID, tbltypeGroup.TypeGroup from tblClasses, tblClassDescriptions, tblTypeGroup "
                                    if session("tabID")<>"" AND isNumeric(session("tabID")) then
                                            strSQL = strSQL & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
                                    end if
                                    strSQL = strSQL & "  WHERE tblClasses.DescriptionID=tblClassDescriptions.ClassDescriptionID AND tblClassDescriptions.ClassPayment=tblTypeGroup.TypeGroupID AND tblTypeGroup.wsEnrollment=1 "
                                    strSQL = strSQL & " AND " & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & " <= tblClasses.ClassDateEnd"
                                    strSQL = strSQL & " ORDER BY TypeGroup"
                                    rsEntry.CursorLocation = 3
                                    rsEntry.open strSQL, cnWS
                                    Set rsEntry.ActiveConnection = Nothing

                                    boolShowTG = rsEntry.RecordCount > 1

                            if boolShowTG then
%>
                <select name="optTG" id="optTG" class="bf-filter" onChange="subForm();">
                  <option value="0"><%=xssStr(allHotWords(516))%></option>
<%
                    Do While NOT rsEntry.EOF
%>
                        <option value="<%=rsEntry("TypeGroupID")%>" title="<%=stringIfLengthOverLimit(rsEntry("TypeGroup"),18)%>" <%if curTG=rsEntry("TypeGroupID") then response.write "selected" end if%>>
													<%=truncateString(rsEntry("TypeGroup"),18)%>
												</option>
<%
                        rsEntry.MoveNext
                    Loop
%>
                </select>
                                <script type="text/javascript">
                                	document.search2.optTG.options[0].text = "<%=jsEsc(allHotWords(516))%>";
                                </script>
<%

                                        end if	''1 or less hide
                                rsEntry.close
                                end if	''show/hide TG

                                if ss_hideEnrollVTs then
                                else
                                        boolShowTG = false

                                                ''Check for > 1 TG otherwise don't display TG selector
                                                strSQL = "SELECT DISTINCT TypeID, TypeName, SortOrder FROM tblVisitTypes, tblClasses, tblClassDescriptions, tblTypeGroup "
                                                if session("tabID")<>"" AND isNumeric(session("tabID")) then
                                                        strSQL = strSQL & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
                                                end if
                                                strSQL = strSQL & "  WHERE tblClasses.DescriptionID=tblClassDescriptions.ClassDescriptionID AND tblClassDescriptions.ClassPayment=tblTypeGroup.TypeGroupID AND tblTypeGroup.wsEnrollment=1 AND tblClassDescriptions.VisitTypeID=tblVisitTypes.TypeID "
                                                strSQL = strSQL & " AND " & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & " <= tblClasses.ClassDateEnd AND tblClasses.ClassActive=1"
                                                strSQL = strSQL & " ORDER BY SortOrder, TypeName"
                                                rsEntry.CursorLocation = 3
                                                rsEntry.open strSQL, cnWS
                                                Set rsEntry.ActiveConnection = Nothing

                                                if NOT rsEntry.EOF then
                                                        rsEntry.MoveNext
                                                        if NOT rsEntry.EOF then
                                                                boolShowTG = true
                                                                rsEntry.MoveFirst
                                                        end if
                                                end if

                                        if boolShowTG then
%>
                <select name="optVT" id="optVT" class="bf-filter" onChange="subForm();">
                  <option value="0"><%=xssStr(allHotWords(518))%></option>
<%
                                        Do While NOT rsEntry.EOF
                                            response.Write("    <option value=""" & rsEntry("TypeID") & """ title=""" & stringIfLengthOverLimit(rsEntry("TypeName"),18) & """")
                                            if curVT = rsEntry("TypeID") then 
                                                response.Write(" selected") 
                                            end if
                                            response.Write(">" & truncateString(rsEntry("TypeName"),18) & "</option>")
                                            rsEntry.MoveNext
                                        Loop
%>
                                <script type="text/javascript">
                                	document.search2.optVT.options[0].text = "<%=jsEsc(allHotWords(518))%>";
                                </script>
                </select>
<%
                                        end if
                                        rsEntry.close
                                end if	'ss hide VTs

                        if ss_showLevels then
                                boolShowTG = false
                                ''check for more than 1 Class Level
                                strSQL = "SELECT DISTINCT tblClassLevels.LevelID, tblClassLevels.LevelName, tblClassLevels.SortOrder FROM tblTypeGroup "
                                if session("tabID")<>"" AND isNumeric(session("tabID")) then
                                        strSQL = strSQL & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
                                end if
                                strSQL = strSQL & "  INNER JOIN tblClassDescriptions ON tblTypeGroup.TypeGroupID = tblClassDescriptions.ClassPayment INNER JOIN tblClasses ON tblClassDescriptions.ClassDescriptionID = tblClasses.DescriptionID INNER JOIN tblClassLevels ON tblClassDescriptions.LevelID = tblClassLevels.LevelID "
                                strSQL = strSQL & "WHERE (" & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & " <= tblClasses.ClassDateEnd) AND (tblClasses.ClassActive = 1) AND (tblTypeGroup.wsEnrollment = 1) ORDER BY tblClassLevels.SortOrder, tblClassLevels.LevelName"
                                rsEntry.CursorLocation = 3
                                rsEntry.open strSQL, cnWS
                                Set rsEntry.ActiveConnection = Nothing

                                ' Only show the drop down if there is more than 1 option
                                boolShowTG = rsEntry.RecordCount > 1

                                if boolShowTG then
%>
                <select name="optLevel" id="optLevel" class="bf-filter" onChange="subForm();">
                  <option value="0"><%=xssStr(allHotWords(518))%></option>
<%
                                    Do While NOT rsEntry.EOF
%>
                  <option value="<%=rsEntry("LevelID")%>" title="<%=stringIfLengthOverLimit(rsEntry("LevelName"),18)%>"<%if curLevel=rsEntry("LevelID") then response.write "selected" end if%>>
										<%=truncateString(rsEntry("LevelName"),18)%>
									</option>
<%
                                            rsEntry.MoveNext
                                    Loop
%>
                </select>
                                <script type="text/javascript">
                                	document.search2.optLevel.options[0].text = "<%=jsEsc(allHotWords(518))%>";
                                </script>
<%
                                end if ''Show Levels > 1
                                rsEntry.close
                        end if 'SS show Levels
%>
                </b>
                <select name="optInstructor" id="optInstructor" class="bf-filter" onChange="subForm();">
                  <option value="0"><%=xssStr(allHotWords(705))%></option>
<%
                        strSQL = "SELECT DISTINCT TrainerID, TrFirstName, TrLastName, TRAINERS.DisplayName, CASE WHEN NOT TRAINERS.DisplayName IS NULL THEN TRAINERS.DisplayName ELSE TRAINERS.TrLastName END from Trainers, tblClasses, tblClassDescriptions, tblTypeGroup "
                        if session("tabID")<>"" AND isNumeric(session("tabID")) then
                                strSQL = strSQL & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
                        end if
                        strSQL = strSQL & "  WHERE tblClasses.DescriptionID=tblClassDescriptions.ClassDescriptionID AND tblClassDescriptions.ClassPayment=tblTypeGroup.TypeGroupID AND tblTypeGroup.wsEnrollment=1 AND Trainers.TrainerID=tblClasses.ClassTrainerID AND TRAINERS.[Delete]=0 AND TRAINERS.ReservationTrn=1 AND tblClasses.ClassDateEnd >=" & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & " AND tblClasses.MaskTrainer = 0 "
                        strSQL = strSQL & "ORDER BY " & GetTrnOrderBy()
                        rsEntry.CursorLocation = 3
                        rsEntry.open strSQL, cnWS
                        Set rsEntry.ActiveConnection = Nothing

                        do While NOT rsEntry.EOF
                            trnName = Replace(FmtTrnNameNew(rsEntry, true), "&nbsp;","")
    %>
                            <option value="<%=rsEntry("TrainerID")%>" title="<%=stringIfLengthOverLimit(trnName,24)%>"<%if curTrn=CLNG(rsEntry("TrainerID")) then response.write "selected" end if%>>
															<%=truncateString(trnName,24)%>
														</option>
    <%
                           rsEntry.MoveNext
                        Loop
                        rsEntry.close
%>
                </select>
                                <script type="text/javascript">
                                	document.search2.optInstructor.options[0].text = "<%=jsEsc(allHotWords(705))%>";
                                </script>

        </div>
</div>
</div>

<div class="wrapperTop">
    <div class="wrapperTopInner">
        <div class="pageTop">
            <div class="pageTopLeft">
                &nbsp; 
            </div>
            <div class="pageTopRight">
                &nbsp; 
            </div>
             <h1><%=DisplayPhrase(phraseDictionary,"Eventschedule")%></h1> 
              <div id="dateControls">
                    <div class="leftSide">
                        <% dayAndWeekControls %>
                    </div>
                    <div class="rightSide">
                        <input id="txtDate" type="text" name="txtDate" size="10" maxlength="14" value="<%=FmtDateShort(topDate)%>" class="date bf-filter" /><!--
                        -->&nbsp;
                        <script type="text/javascript">
                            var cal1 = new tcal({'id' : 'dayInfo', 'formname':'search2', 'controlname':'txtDate'});
                            cal1.a_tpl.yearscroll = true;
                            $('#txtDate').change(<%if InStr(Request.ServerVariables("HTTP_USER_AGENT"), "MSIE" )>0 then response.write "chkDate" else response.write "subForm" end if%>);
							$('#tcalico_dayInfo').attr('onclick', '$("#txtDate").focus()');
                        </script>
                    </div>
                </div>    
          </div>
    </div>
</div>
</form>

<% pageStart %>
  <table width="<%=strPageWidth%>" cellspacing="0" class="main-enroll-table" id="cm-m-enroll-tbl">
	<tr> 
		<td class="center" valign="top" height="100%" width="100%"> 
<%
	dim strCurWeekDay, dayCounter, failureReason, contactEmail, ss_EnrollHideTimes, irs_cMembershipID, clsStartTime, ss_EnrollCltModeShowRsrc, ss_reserveForOtherClient
	
	Dim Image
	Set Image = Server.CreateObject("csImageFile.Manage")

	ss_EnrollHideTimes = checkStudioSetting("tblResvOpts","EnrollHideTimes")
    ss_EnrollCltModeShowRsrc = checkStudioSetting("tblResvOpts","EnrollCltModeShowRsrc")
    ss_reserveForOtherClient = checkStudioSetting("tblGenOpts","reserveForOtherClient")
    
	'check membership
	if session("mvarUserId")<>"" then
		isMember = checkMembership(Session("mvarUserId"),"")
		if isMember then
			irs_cMembershipID = MemSeriesTypeID
		else
			irs_cMembershipID = -1
		end if
	else
		isMember = false
		irs_cMembershipID = -1
	end if

    set rsEntry = Server.CreateObject("ADODB.Recordset")

	'setup schedule start/end window for all tg's disabled
	strSQL = "SELECT TypeGroupID, SchedOffset, SchedOffsetEnd FROM tblTypeGroup WHERE wsEnrollment=1 AND [Active]=1"
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	if NOT rsEntry.EOF then
		numTGs = rsEntry.RecordCount
		ReDim ssSchedOffset(numTGs-1, 3)
		subCount = 0
		do while NOT rsEntry.EOF
			ssSchedOffset (subCount,0) = rsEntry("TypeGroupID")
			ssSchedOffset (subCount,1) = rsEntry("SchedOffset")
			ssSchedOffset (subCount,2) = rsEntry("SchedOffsetEnd")
			subCount = subCount + 1
			rsEntry.MoveNext
		loop
	else
		numTGs = 0
		ReDim ssSchedOffset(1, 3)
		ssSchedOffset (0,0) = 0
		ssSchedOffset (0,1) = 0
		ssSchedOffset (0,2) = 0
	end if
	subCount = 0
	rsEntry.close
	
	
	strSQL = "SELECT ClientContactEmail FROM tblGenOpts "
	
	rsEntry.CursorLocation = 3
	rsEntry.open strSQL, cnWS
	Set rsEntry.ActiveConnection = Nothing
	
	if NOT rsEntry.EOF then
		contactEmail = rsEntry("ClientContactEmail")
	end if
	
	rsEntry.close





	function getSchedWin(tgid, schwintype)
		getSchedWin = 0
		tmpcount = 0
		do while tmpcount < numTGs
			if ssSchedOffset(tmpCount,0) = tgid then
				if schwintype=1 then	'Start Window
					getSchedWin = ssSchedOffset(tmpCount, 1)
				else	'End Window
					getSchedWin = ssSchedOffset(tmpCount, 2)
				end if
			end if
			tmpcount = tmpcount + 1
		loop
		exit function
	end function

    tmpDate = CDate(FormatDateTime(DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))),2))
	curPageDate = CDate(FormatDateTime(DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))),2))
	
	

		if request.form("optVT")<>"" then
			cResVT = CLNG(sqlInjectStr(request.form("optVT")))
		elseif request.querystring("vt")<>"" then
			cResVT = CLNG(request.querystring("vt"))
		else
			cResVT=0
		end if
		if isNumeric(cResVT) then
			if cResVT<>0 then
				strSQL = "SELECT tblVisitTypes.TypeName FROM tblVisitTypes INNER JOIN tblTypeGroup ON tblVisitTypes.Typegroup = tblTypeGroup.TypeGroupID WHERE (tblTypeGroup.wsEnrollment = 1) AND tblVisitTypes.TypeID=" & cResVT
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
		end if

		if request.form("optLevel")<>"" AND IsNumeric(request.form("optLevel")) then
			curLevel = CLNG(sqlInjectStr(request.form("optLevel")))
		elseif request.querystring("lvl")<>"" AND IsNumeric(request.querystring("lvl")) then
			curLevel = CLNG(request.querystring("lvl"))
		else
			curLevel=0
		end if
		if request.form("optTG")<>"" AND IsNumeric(request.form("optTG")) then
			cResTG = CINT(sqlInjectStr(request.form("optTG")))
		elseif request.querystring("tg")<>"" AND IsNumeric(request.querystring("tg")) then
			cResTG = CINT(request.querystring("tg"))
		else
			cResTG=0
		end if
		if request.form("optInstructor")<>"" AND IsNumeric(request.form("optInstructor")) then
			cTrnID = CLNG(sqlInjectStr(request.form("optInstructor")))
		elseif request.querystring("trn")<>"" AND IsNumeric(request.querystring("trn")) then
			cTrnID = CLNG(request.querystring("trn"))
		else
			cTrnID = -1
		end if

		' BQL 50_2718  Added CourseLocation, CourseLocationName fields to queries and code below to differentiate between class location and course location
		' BN Bug#4219  Classes with multiple pre-reqs showed up as many times as there were pre-reqs.  Since we just chec PreReqType for isNull it doesn't matter
		'              how many pre-reqs there are so just look to see if there is at least one.
		strSQL =" SELECT PreReqs.ClassDescriptionID AS PreReqType, tblClassDescriptions.ClassDescriptionID, tblClasses.ClassID, tblCourses.CourseID, tblCourses.CourseName, "&_
				" tblCourses.CourseDescription, tblClasses.ClassStartTime, tblClasses.ClassEndTime, tblClasses.ClassDateStart, tblClasses.ClassDateEnd, "&_
				" tblClasses.DaySunday, tblClasses.DayMonday, tblClasses.DayTuesday, tblClasses.DayWednesday, tblClasses.DayThursday, tblClasses.DayFriday, "&_
				" tblClasses.DaySaturday, tblClasses.ClassCapacity, tblClasses.Free, tblClasses.MaskTrainer, "&_
				" ISNULL(tblCourses.AllowOpenEnrollment, tblClasses.AllowOpenEnrollment) as AllowOpenEnrollment, "&_
				" ISNULL(tblCourses.AllowDateForwardEnrollment, tblClasses.AllowDateForwardEnrollment) as AllowDateForwardEnrollment, "&_
				" tblClassDescriptions.ClassName, tblClassDescriptions.ClassPayment, tblClassDescriptions.ClassDescription, tblClassDescriptions.ClassNotes, "&_
				" tblCourses.LocationID as CourseLocationID, tblClasses.LocationID, CL.LocationName as CourseLocationName, Location.LocationName, "&_
				" TRAINERS.TrainerID, TRAINERS.TrFirstName, TRAINERS.TrLastName, TRAINERS.DisplayName, TRAINERS.Bio, tblTypeGroup.wsDisable, "&_
				" ISNULL(ClassCount.ClassCount, 0) AS ClassCount, NextClass.NextClassDate, tblClasses.TrainerID2, tblClasses.TrainerID3, "&_
				" Asst1.TrFirstName AS Asst1FirstName, Asst1.TrLastName AS Asst1LastName, Asst1.DisplayName AS Asst1DisplayName, Asst2.TrFirstName AS Asst2FirstName, "&_
				" Asst2.TrLastName AS Asst2LastName, Asst2.DisplayName AS Asst2DisplayName, MemLevel.TypeGroupID AS HasTGAccess, "&_
				" IsNull(MemRestrict.NumRestrictions, 0) AS NumRestrictions, tblResources.ResourceName, " &_
				" CASE WHEN Asst1.Bio is null or DATALENGTH(RTRIM(LTRIM(CONVERT(nvarchar(max), Asst1.bio)))) = 0 THEN 0 ELSE 1 END AS Asst1HasBio, "&_
				" CASE WHEN Asst2.Bio is null or DATALENGTH(RTRIM(LTRIM(CONVERT(nvarchar(max), Asst2.bio)))) = 0 THEN 0 ELSE 1 END AS Asst2HasBio, "&_
				" CASE WHEN "
				if ss_AllowOpenEnrollment then
					strSQL = strSQL & " (ISNULL(tblCourses.AllowDateForwardEnrollment, tblClasses.AllowDateForwardEnrollment) = 1"
					strSQL = strSQL & " OR"
					strSQL = strSQL & " ISNULL(tblCourses.AllowOpenEnrollment, tblClasses.AllowOpenEnrollment) = 1) OR"
				end if
				strSQL = strSQL & " NOT EXISTS (SELECT 1 FROM tblClasses innerClasses WHERE CASE WHEN tblClasses.CourseID IS NOT NULL THEN innerClasses.CourseID ELSE innerClasses.ClassID END  = ISNULL(tblClasses.CourseID, tblClasses.ClassID) AND innerClasses.ClassDateStart <"
				strSQL = strSQL &  DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep & ") THEN 1 ELSE 0 END AS AllowToEnroll "&_
				" FROM tblTypeGroup "&_
				" INNER JOIN tblClassDescriptions ON tblTypeGroup.TypeGroupID = tblClassDescriptions.ClassPayment " &_
		        " LEFT OUTER JOIN ( " &_
			        "SELECT DISTINCT tblClassPrereq.ClassDescriptionID " &_
			        "FROM tblClassPrereq " &_
		        " ) AS PreReqs ON tblClassDescriptions.ClassDescriptionID = PreReqs.ClassDescriptionID "
		
		if isNum(session("tabID")) then
			strSQL = strsql & " INNER JOIN tblTypeGroupTab ON tblTypeGroupTab.TypeGroupID = tblTypeGroup.TypeGroupID AND tblTypeGroupTab.TabID = " & session("tabID")
		end if 
		
		strSQL = strSQL &_
				" LEFT OUTER JOIN ("&_
				"	SELECT TypeGroupID FROM tblTypeGroupMembership WHERE (SeriesTypeID = " & irs_cMembershipID & ")"&_
				" ) AS MemLevel ON tblTypeGroup.TypeGroupID = MemLevel.TypeGroupID "&_
				" LEFT OUTER JOIN ("&_
				"	SELECT COUNT(*) AS NumRestrictions, TypeGroupID FROM tblTypeGroupMembership GROUP BY TypeGroupID"&_
				" ) AS MemRestrict ON tblTypeGroup.TypeGroupID = MemRestrict.TypeGroupID "&_
				" INNER JOIN tblClasses "&_
				"	INNER JOIN TRAINERS ON tblClasses.ClassTrainerID = TRAINERS.TrainerID "&_
				" ON tblClassDescriptions.ClassDescriptionID = tblClasses.DescriptionID "&_
				" LEFT OUTER JOIN TRAINERS Asst1 ON tblClasses.TrainerID2 = Asst1.TrainerID "&_
				" LEFT OUTER JOIN TRAINERS Asst2 ON tblClasses.TrainerID3 = Asst2.TrainerID "&_
				" LEFT OUTER JOIN ("&_
				"	SELECT MIN(classDateStart) AS courseDateStart, CourseID "&_
				"	FROM tblClasses "&_
				"	WHERE CourseID IS NOT NULL GROUP BY CourseID"&_
				" ) STARTDATE ON STARTDATE.CourseID = tblClasses.CourseID "&_
				" LEFT OUTER JOIN ("&_
				"	SELECT MIN(ClassDate) AS trueStart, ClassID "&_
				"	FROM tblClassSch "&_
				"	WHERE ClassDate > " & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep &_
				"		OR (ClassDate = " & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep &_
				"		AND EndTime > " & DateSep & "1899-12-30 " & FormatDateTime(DateAdd("n", Session("tzOffset"),Now), 4) & DateSep & ") "&_
				"   GROUP BY ClassID "&_
				" ) AS CLASSSTART ON CLASSSTART.ClassID = tblClasses.ClassID " &_
				" LEFT OUTER JOIN tblResources "&_
				"	INNER JOIN tblResourceSchedules ON tblResources.ResourceID = tblResourceSchedules.ResourceID "&_
				" ON tblClasses.ClassID = tblResourceSchedules.RefClass AND tblResourceSchedules.StartDate=tblClasses.ClassDateStart "
		
		if session("curLocation")<>"0" then
			strSQL = strSQL &	" LEFT OUTER JOIN ( "&_
								"	SELECT tblCourses.CourseID, tblClasses.LocationID "&_
								"	FROM tblCourses "&_
								"	INNER JOIN tblClasses ON tblCourses.CourseID = tblClasses.CourseID "&_
								"	GROUP BY tblClasses.LocationID, tblCourses.CourseID "&_
								"	HAVING (tblClasses.LocationID = " & session("curLocation") & ") "&_
								" ) CrsLoc ON CrsLoc.CourseID = tblClasses.CourseID "
		end if
		
		strSQL = strSQL &	" INNER JOIN Location ON tblClasses.LocationID = Location.LocationID " &_
							" INNER JOIN tblVisitTypes ON tblClassDescriptions.VisitTypeID = tblVisitTypes.TypeID "
		
		if cResVT<>"0" then
			strSQL = strSQL & " AND tblVisitTypes.TypeName=N'" & Replace(cResVT,"'","''") & "' "
		end if
		if cResTG=0 then
			strSQL = strSQL & " AND tblTypeGroup.wsEnrollment=1 "
		else
			strSQL = strSQL & " AND tblClassDescriptions.ClassPayment=" & cResTG & " "
		end if
		if curLevel<>0 then
			strSQL = strSQL & " AND tblClassDescriptions.LevelID=" & curLevel & " "
		end if	
		if cTrnID>0 then
			strSQL = strSQL & " AND TRAINERS.TrainerID=" & cTrnID & " AND tblClasses.MaskTrainer = 0 "
		end if
		if ssCltModeEnrollEndDates then
			strSQL = strSQL & " AND classDateEnd >= " & DateSep & cFrmDate & DateSep & " " 
		else
			strSQL = strSQL & " AND classDateStart >= " & DateSep & cFrmDate & DateSep & " " 
		end if
		strSQL = strSQL &	" LEFT OUTER JOIN ("&_
							"	SELECT ClassID, COUNT(ClassID) AS ClassCount "&_
							"	FROM tblClassSch "&_
							"	WHERE TrainerID <> - 1 "&_
							"	GROUP BY ClassID "&_
							" ) ClassCount ON ClassCount.ClassID = tblClasses.ClassID " &_
							" LEFT OUTER JOIN ( "&_
							"	SELECT ClassID, MIN(ClassDate) AS NextClassDate " &_
							"	FROM tblClassSch WHERE ClassDate > " & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep &_
							"	OR (ClassDate = " & DateSep & DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))) & DateSep &_
							"		AND EndTime > " & DateSep & "1899-12-30 " & FormatDateTime(DateAdd("n", Session("tzOffset"),Now), 4) & DateSep &  ") "&_
							"	GROUP BY ClassID "&_
							" ) NextClass ON NextClass.ClassID = tblClasses.ClassID "&_
							" LEFT OUTER JOIN tblCourses ON tblCourses.CourseID = tblClasses.CourseID "&_
							" LEFT OUTER JOIN Location CL ON CL.LocationID = tblCourses.LocationID " &_
							" WHERE ISNULL(tblCourses.ShowToPublic, classActive) = 1 "
		if session("curLocation")<>"0" then
			strSQL = strSQL & " AND (tblClasses.LocationID=" & session("curLocation") & " OR CrsLoc.LocationID IS NOT NULL) "
		end if
		 ' filter out any enrollments that are fully cancelled.
		strSQL = strSQL &_
			" AND EXISTS ("&_
				" SELECT 1 FROM tblClassSch AS t0 INNER JOIN tblClasses AS t1 ON t0.ClassID = t1.ClassID "&_
				" WHERE t0.TrainerID <> -1 AND t0.ClassID = tblClasses.ClassID ) "


        if ss_EnrollSchedSortBy=0 then 'by name
			strSQL = strSQL &	" ORDER BY CASE WHEN tblClasses.CourseID IS NULL THEN tblClassDescriptions.ClassName ELSE tblCourses.CourseName END, "&_
								" CASE WHEN tblClasses.CourseID IS NULL THEN CLASSSTART.trueStart ELSE STARTDATE.CourseDateStart END, "&_
								" tblClasses.CourseID, tblClasses.ClassDateStart, ClassStartTime"
		elseif ss_EnrollSchedSortBy=1 then 'by day/time
            strSQL = strSQL &	" ORDER BY CASE WHEN tblClasses.CourseID IS NULL THEN CLASSSTART.trueStart ELSE STARTDATE.CourseDateStart END, "&_
								" tblClasses.CourseID, tblClasses.ClassDateStart, ClassStartTime"
		end if      

		'response.write debugSQL(strSQL, "SQL")
		'response.write "SQL: " & strSQL

		rsEntry.CursorLocation = 3
		rsEntry.open strSQL, cnWS
		Set rsEntry.ActiveConnection = Nothing
%>

<table class="center mainText" border="0" cellspacing="0" cellpadding="0" width="100%" >
	<% if Session("SchOpenDOMCloseDate")<>"" then %>
	<tr>
		<td>
			<div class="m-enroll-sch-restrictions-ctr">
				<span class="m-enroll-sch-restrictions-txt">
                    <%ReplaceInPhrase phraseDictionary, "Schedulingforclassnameopens","<MONTH>", MonthName(Month(DateAdd("m", 1, Session("SchOpenDOMCloseDate"))))%>
                    <%ReplaceInPhrase phraseDictionary, "Schedulingforclassnameopens","<DATE>", FmtDateShort(CDATE(Month(Session("SchOpenDOMCloseDate")) & "/" & Session("SchOpenDOM") & "/" & Year(Session("SchOpenDOMCloseDate"))))%>
					<%=DisplayPhrase(phraseDictionary,"Schedulingforclassnameopens")%>
				</span>
			</div>
		</td>
	</tr>
	<% elseif SchHrsClosed then %>
	<tr>
		<td>
			<div class="m-enroll-sch-restrictions-ctr">
				<span class="m-enroll-sch-restrictions-txt m-enroll-sch-restrictions-txt-closed">
					<%=DisplayPhraseJS(phraseDictionary,"Schedclosed") & ": " & FmtTimeShort(SchHrsStartTime) & " - " & FmtTimeShort(SchHrsEndTime)%>
				</span>
			</div>
		</td>
	</tr>
	<% end if %>



  <%
	if rsEntry.EOF then

%>
  <tr> 
    <td class="mainText" align="left" valign="top"><br />
      &nbsp;<%=DisplayPhrase(phraseDictionary,"Noscheduledworkshops")%></td>
  </tr>
<%	else	%>

  <%
	end if

dim dWeekCount, dayCharNum, pbgcolor
dim tmpBioStr, trnHREF

	rowcount = 1

	Do While NOT rsEntry.EOF
%>


  <tr> 
    <td align="left" valign="top" height="2"> 
	    <%= drawEnrollmentDescription() %>
    </td>
  </tr>
  
  <%
		rsEntry.MoveNext
	Loop
rsEntry.close
set rsEntry = nothing
%>
  <tr> 
    <td align="left" valign="top">&nbsp;</td>
  </tr>
</table>
		</td>
    </tr>
</table>
<!-- #include file="inc_login_content.asp" -->
<!-- #include file="inc_fb_confirm_login_lb.asp" -->
<!-- #include file="inc_desc_lightbox.asp" -->
<% pageEnd %>
<!-- #include file="post.asp" -->
<%



function drawEnrollmentDescription()	
	dim mainLocId, mainLocName
	%>
	<div class="mainText enrollment group">
		<div class="bb4 mainTextBig">
            <span style="float: left">
			<%
			' instructor bio link
			tmpBioStr = rsEntry("Bio")

			' start and end dates
			if CDATE(rsEntry("ClassDateStart")) < CDATE(rsEntry("ClassDateEnd")) then
				multiDays = true
			else
				multiDays = false
			end if
				
			dWeekCount = 0
				
			if multiDays then
				'''''Count number of Days of Week Checked
				
				if rsEntry("DaySunday") then dWeekCount = dWeekCount + 1 end if
				if rsEntry("DayMonday") then dWeekCount = dWeekCount + 1 end if
				if rsEntry("DayTuesday") then dWeekCount = dWeekCount + 1 end if
				if rsEntry("DayWednesday") then dWeekCount = dWeekCount + 1 end if
				if rsEntry("DayThursday") then dWeekCount = dWeekCount + 1 end if
				if rsEntry("DayFriday") then dWeekCount = dWeekCount + 1 end if
				if rsEntry("DaySaturday") then dWeekCount = dWeekCount + 1 end if
			
			end if
			
			if dWeekCount >2 then
				dayCharNum=2
			elseif dWeekCount = 2 then
				dayCharNum=3
			else
				dayCharNum=10
			end if
			
			if isNull(rsEntry("CourseID")) then
				response.write("<strong>" & rsEntry("ClassName") & "</strong>")
				if NOT rsEntry("MaskTrainer") then
					response.Write("<span class=""mainText"">&nbsp;" & DisplayPhrase(phraseDictionary, "With") & "&nbsp;</span>")
                    
                    if NOT isNull(tmpBioStr) AND TRIM(tmpBioStr)<>"" then
                        response.Write("<a href=""javascript:;"" name=""bio" & rsEntry("trainerID") & """ class=""modalBio"" >" & FmtTrnNameNew(rsEntry, false) & "</a>")
                    else
                        response.write(FmtTrnNameNew(rsEntry, false))
                    end if
					
                    if ss_ShowAsst1 then
						if NOT isNull(rsEntry("TrainerID2")) then
                                if NOT isNull(rsEntry("TrainerID3")) then
							        response.write("<span class=""mainText"">, </span>")
							    else
								    response.Write("<span class=""mainText"">&nbsp;" & DisplayPhrase(phraseDictionary, "And") & "&nbsp;</span>")
							    end if

                                if rsEntry("Asst1HasBio") then
							        response.write("<a class=""modalBio"" name=""bio" & rsEntry("TrainerID2") & """ href=""javascript:;"">" & FmtAsstNameRS(rsEntry, false, 1) & "</a>")
                                else
                                    response.write(FmtAsstNameRS(rsEntry, false, 1))
                                end if
						end if
					end if
					
                    if ss_ShowAsst2 then
						if NOT isNull(rsEntry("TrainerID3")) then
							response.write("<span class=""mainText"">&nbsp;" & DisplayPhrase(phraseDictionary, "And") & "&nbsp;</span>")
                            if rsENtry("Asst2HasBio") then
                                response.Write("<a class=""modalBio"" name=""bio" & rsEntry("TrainerID3") &""" href=""javascript:;"">" & FmtAsstNameRS(rsEntry, false, 2) & "</a>")
                            else
                                response.write(FmtAsstNameRS(rsEntry, false, 2))
                            end if
						end if
					end if
				end if

				mainLocId = rsEntry("LocationId") 
				mainLocName = rsEntry("LocationName")
			else 
				courseFirstClassID = rsEntry("classID")
				courseFirstClassDate = rsEntry("ClassDateStart")
				calcSDate = courseFirstClassDate 
				mainLocId = rsEntry("CourseLocationId") 
				mainLocName = rsEntry("CourseLocationName")

				response.write("<b>" & rsEntry("CourseName") & "</a></b>")
			end if %>
            </span>
			<span class="locationLink">
				<% if session("numLocations")>1 then %>
					<b>
						<%=allHotWords(68) %>:
						<a class="modalLocationInfo" name="loc<%=mainLocId%>" href="javascript:;"><%=mainLocName%></a>
					</b> 
				<% end if %>
			</span>
			<div class="clear"></div>
		</div>
		
<%		 if isNull(rsEntry("CourseID")) then
		'NOT A COURSE
	        
            if rowcount=0 then
               rowcount = 1
               pbgcolor = "#F2F2F2"
            else
	           rowcount = 0
               pbgcolor ="#FAFAFA"
            end if 
			%>
						
			<div class="dates-and-time group">
			<%				
				'build date/time string....... lots o' code to do this
				dateStr = ""
				timeStr = ""
				resourceStr = ""
				if NOT (rsEntry("AllowOpenEnrollment") OR rsEntry("AllowDateForwardEnrollment")) OR isNull(rsEntry("NextClassDate")) then 'normal
					calcSDate = CDATE(rsEntry("ClassDateStart"))
				else
					calcSDate = CDATE(rsEntry("NextClassDate"))
				end if

				first = true

				intCount = 0
				dayCounter = WeekDay(calcSDate)
				tmpDate = calcSDate

				Do While intCount < 7
					strCurWeekDay = WeekdayName(dayCounter)

					if rsEntry("Day" & strCurWeekDay) then
						if first=false then
							dateStr = dateStr & ", "
						else
							first=false
							calcSDate = tmpDate
						end if
						
						
						if strCurWeekDay = "Sunday" then
							dayHW = allHotWords(445)
						elseif strCurWeekDay = "Monday" then 
							dayHW = allHotWords(446)
						elseif strCurWeekDay = "Tuesday" then
							dayHW = allHotWords(447)
						elseif strCurWeekDay = "Wednesday" then
							dayHW = allHotWords(448)
						elseif strCurWeekDay = "Thursday" then
							dayHW = allHotWords(449)
						elseif strCurWeekDay = "Friday" then
							dayHW = allHotWords(450)
						elseif strCurWeekDay = "Saturday" then
							dayHW = allHotWords(451)
						else
							dayHW = WeekDayName(WeekDay(tmpCount))
						end if
						first = false

						dateStr = dateStr & dayHW
					end if

					if dayCounter <> 7 then
						dayCounter = dayCounter + 1 
					else
						dayCounter = 1
					end if
					intCount = intCount + 1
					tmpDate = tmpDate + 1
				Loop
				
				'add start date
				dateStr = dateStr & " " & FmtDateShort(calcSDate)
				
				'add end date, if different
				if calcSDate <> CDATE(rsEntry("ClassDateEnd")) then
					dateStr = " " & dateStr & " - " & FmtDateShort(rsEntry("ClassDateEnd"))
				end if
				%>			
                  
		<% if NOT ss_EnrollHideTimes then %>
		        <%= drawClassDates(rsEntry("ClassDateStart"), rsEntry("ClassDateEnd")) %>
				<%= drawClassTimes(rsEntry("ClassStartTime"), rsEntry("ClassEndTime")) %>
               
		<% end if %>
            </div> <!-- .dates-and-time -->
		<%if session("useResrcResv") AND ss_EnrollCltModeShowRsrc AND not isNull(rsEntry("ResourceName")) then%>
			<div class="resourceName">
             <%=allHotWords(0) %>:
              <b>
	               <%=rsEntry("ResourceName")%>
              </b>
			</div>
		<%end if%>
		<div class="description group userHTML">
            <div class="enrollment-image">
				<%= conditionalImage("\reservations\" & rsEntry("ClassDescriptionID") & ".jpg") %>
            </div>
			<%=HtmlPurifyForDisplay(rsEntry("ClassDescription"))%>
		</div>
		<% if NOT isNull(rsEntry("ClassNotes")) then %>
			<div class="notes userHTML">
				<%=HtmlPurifyForDisplay(rsEntry("ClassNotes"))%>
			</div>
		<% end if %>
		
<%	else ' course
		curCourseID = CINT(rsEntry("CourseID"))
		cont = true %>
			<div class="mainText">
            
			
		<%	
		' BQL 49_2544 - Added CourseDescription to query and page output
		if NOT isNull(rsEntry("CourseDescription")) AND rsEntry("CourseDescription")<>"" then 
			%>
			<div class="description group">
                <div class="enrollment-image">
                    <%= conditionalImage("\courses\" & rsEntry("CourseID") & ".jpg") %>
                </div>
				<%=rsEntry("CourseDescription")%>
			</div>
			<%	
		end if 
		%>
       
	<%	do while cont %>
		<div class="course group" style="margin-left:20px;">
			<div class="courseName">
			<b><%=rsEntry("ClassName")%></b>

			<% if NOT rsEntry("MaskTrainer") then
				response.write("<span class=""mainText"">&nbsp;"& DisplayPhrase(phraseDictionary, "With") & "&nbsp;</span>")
				tmpBioStr = rsEntry("Bio")
				if NOT isNull(tmpBioStr) AND TRIM(tmpBioStr)<>"" then
                    response.write("<a class=""modalBio"" name=""bio" & rsEntry("trainerID") & """ href=""javascript:;"" >" & FmtTrnNameNew(rsEntry, false) & "</a>")
				else
					response.write(FmtTrnNameNew(rsEntry, false))
                end if
				if NOT isNull(rsEntry("TrainerID2")) then
					if  NOT isNull(rsEntry("TrainerID3")) then
						response.write("<span class=""mainText"">,&nbsp;</span>")
					else
						response.write("<span class=""mainText"">&nbsp;" & DisplayPhrase(phraseDictionary, "And") & "&nbsp;</span>")
					end if
                    
                    if rsEntry("Asst1HasBio") then
					    response.Write("<a class=""modalBio"" name=""bio" & rsEntry("trainerID2") &""" href=""javascript:;"">" & FmtAsstNameRS(rsEntry, false, 1) & "</a>")
                    else
                        response.Write(FmtAsstNameRS(rsEntry, false, 1) & " ")
                    end if
				end if
				if NOT isNull(rsEntry("TrainerID3")) then
					response.write("<span class=""mainText"">&nbsp;" & DisplayPhrase(phraseDictionary, "And") & "&nbsp;</span>")
                    if rsEntry("Asst2HasBio") then
                        response.Write("<a class=""modalBio"" name=""bio" & rsEntry("trainerID3") & """ href=""javascript:;"">" & FmtAsstNameRS(rsEntry, false, 2) & "</a>")
                    else
                        response.Write(FmtAsstNameRS(rsEntry, false, 2))
                    end if
				end if
				if session("numLocations")>1 then
					response.write("&nbsp;" & DisplayPhrase(phraseDictionary, "At") & "&nbsp;")
                    response.Write("<a class=""modalLocationInfo"" name=""loc" & rsEntry("LocationID") &""" href=""javascript:;"">" &rsEntry("LocationName") & "</a></b>")
				end if
            end if ' MaskTrainer %>
			</div>
			
			<div class="dates-and-time group">
				<%= drawClassDates(rsEntry("ClassDateStart"), rsEntry("ClassDateEnd")) %>
				<%= drawClassTimes(rsEntry("ClassStartTime"), rsEntry("ClassEndTime")) %>			
			</div>
				
<%				if NOT isNull(rsEntry("ClassDescription")) AND rsEntry("ClassDescription") <> "" then %>
					<div class="description userHTML">
						<%= HtmlPurifyForDisplay(rsEntry("ClassDescription")) %>
					</div>
<%				end if  %>
				
<%				if NOT isNull(rsEntry("ClassNotes")) AND rsEntry("ClassNotes") <> ""then 
					if rsEntry("ClassNotes")<>"" then%>
						<div class="notes userHTML">
							<%= HtmlPurifyForDisplay(rsEntry("ClassNotes"))%>
						</div>
<%					end if 
				end if
			
				rsEntry.MoveNext
				if rsEntry.EOF then
					cont = false
				elseif isNull(rsEntry("CourseID")) then
					cont = false
				elseif CINT(rsEntry("CourseID"))<>curCourseID then
					cont = false
				end if
%>
		</div>
<%		'end of <div class="course"> 

		
		loop
		rsEntry.MovePrevious
	end if
		
    'roomInClass =  canEnroll(rsEntry("ClassID"), DateValue(CDATE(DateAdd("n", Session("tzOffset"),Now))), false, true, true )  
 	if (rsEntry("ClassCount")>0) AND (NOT rsEntry("wsDisable") AND (rsEntry("HasTGAccess") OR rsEntry("NumRestrictions")=0) )  AND NOT SchHrsClosed AND (DateValue(rsEntry("ClassDateStart")) <= SchOpenDOMCloseDate) AND (NOT SameDaySchCutoff OR DateValue(rsEntry("ClassDateStart")) <> DateValue(DateAdd("n", Session("tzOffset"),Now)) ) AND rsEntry("ClassCapacity") <> 0 then'AND roomInClass then 

		'CB - 11_5_07 patched live change where sign up button provided when open enrollment is enabled.
		if isNull(rsEntry("ClassStartTime")) then 
			clsStartTime = "" 
		else 
			clsStartTime = TimeValue(rsEntry("ClassStartTime")) 
		end if
		
		if ((CDATE(calcSDate & " " & clsStartTime) < DateAdd("n", getSchedWin(rsEntry("ClassPayment"), 1), DateAdd("n", Session("tzOffset"),Now))) OR ((CDATE(calcSDate & " " & clsStartTime) > DateAdd("y", getSchedWin(rsEntry("ClassPayment"), 2), DateAdd("n", Session("tzOffset"),Now))) AND getSchedWin(rsEntry("ClassPayment"), 2)<>0)) AND NOT(rsEntry("AllowOpenEnrollment") OR rsEntry("AllowDateForwardEnrollment")) OR rsEntry("AllowToEnroll") = "0" then
			%>
			<div class="registration-closed">
				<%=DisplayPhrase(phraseDictionary,"Registrationclosed")%>
				<br />
			  	<% if NOT isNull(contactEmail) then %>
                    <%ReplaceInPhrase phraseDictionary, "Contactstudio", "<STUDIONAME>", session("StudioName")%>
			  		<a href="mailto:<%=contactEmail%>"><img align="absbottom" src="<%= contentUrl("/asp/adm/images/icon_email.png") %>"><%=DisplayPhrase(phraseDictionary,"Contactstudio")%></a>
				<% end if %>
			</div>
			<%
		else

			if UseVersionB(1) OR Session("pass") then 'AB TEST
				'REDIRECT TO RES_X

				if (rsEntry("Free") AND NOT ss_reserveForOtherClient) AND rsEntry("ClassDateStart")=rsEntry("ClassDateEnd") AND isNull(rsEntry("PreReqType")) then 
					%>
					<input class="sign-up-now" onClick="document.location='res_debf.asp?tg=<%=rsEntry("ClassPayment")%>&classId=<%=rsEntry("ClassID")%>&classDate=<%=calcSDate%>&clsLoc=<%=rsEntry("LocationID")%>';" type="button" name="signupBtn" value='<%=DisplayPhraseAttr(phraseDictionary,"Signup")%>'>
					<% 
				else 
					%>
					<input class="sign-up-now" onClick="document.location='res_a.asp?tg=<%=rsEntry("ClassPayment")%>&classId=<% if isNull(rsEntry("CourseID")) then response.write rsEntry("ClassID") else response.write courseFirstClassID end if %>&courseID=<%=rsEntry("CourseID")%>&classDate=<% if isNull(rsEntry("CourseID")) then response.write CSTR(calcSDate) else response.write courseFirstClassDate end if %>&clsLoc=<%=rsEntry("LocationID")%>';" type="button" name="signupBtn" value='<%=DisplayPhraseAttr(phraseDictionary,"Signup")%>'>
					<% 
				end if
				'build forwarding link for successful login submits
				dim varLink
			else 
			'B Test -- SHOW LIGHTBOX, LOOK PRETTY
				if (rsEntry("Free") AND NOT ss_reserveForOtherClient) AND rsEntry("ClassDateStart")=rsEntry("ClassDateEnd") AND isNull(rsEntry("PreReqType")) then 				
					'free class
					varLink = "/ASP/res_debf.asp?tg=" & rsEntry("ClassPayment") & "&classId=" & rsEntry("ClassID") & "&classDate=" & calcSDate & "&clsLoc=" & rsEntry("LocationID")
				else 
					'not free class
					varLink = "/ASP/res_a.asp?tg=" & rsEntry("ClassPayment") & "&classId="
					if isNull(rsEntry("CourseID")) then 
						varLink = varLink & rsEntry("ClassID") 
					else 
						varLink = varLink & courseFirstClassID 
					end if
				
					varLink = varLink & "&courseID=" & rsEntry("CourseID") & "&classDate=" 
					if isNull(rsEntry("CourseID")) then 
						varLink = varLink & CSTR(calcSDate) 
					else 
						varLink = varLink & courseFirstClassDate 
					end if 
				
					varLink = varLink & "&clsLoc=" & rsEntry("LocationID")
				end if
			
				'Class name/time/date info string
				if NOT isNull(rsEntry("CourseName")) then
					'bookingStr1 = rsEntry("CourseName")
					bookingStr2 =  rsEntry("CourseName")

				else
					'bookingStr1 = rsEntry("ClassName") 
					bookingStr2 = rsEntry("ClassName") & " on " & dateStr & timeStr & resourceStr
					'tack on trainer's name for classes	
					trainerName = Trim(FmtTrnNameNew(rsEntry, false)) 'FmtTrnName takes the current row to extract the proper teacher name
					'add trainer's name if there is one available
					if trainerName <> "" And NOT rsEntry("MaskTrainer") then
						bookingStr2 = bookingStr2 & ", " & DisplayPhrase(phraseDictionary, "With") & " " & trainerName
					end if
				end if

				Response.write("<div style=""margin:5px 0px"">")
					Response.write("<input class=""sign-up-now"" onClick=""promptLogin('" & jsEscSingle(xssStr(bookingStr1)) & "', '" & jsEscSingle(xssStr(bookingStr2)) & "', '" & varLink & "');"" ")
					Response.write " name=""signupBtn"" type=""button"" value="""& DisplayPhraseAttr(phraseDictionary,"Signup") &""">"
				Response.write("</div>")
			end if 'a/b test
		end if
        
	else ' the user can't sign up - notify them of the reason.
		if rsEntry("ClassCount")<=0 then
			failureReason = DisplayPhrase(phraseDictionary,"Enrollmentfull")
		elseif rsEntry("wsDisable") OR rsEntry("AllowToEnroll") = "0" then 'OR NOT roomInClass then
			failureReason = DisplayPhrase(phraseDictionary,"Registrationunavailable")
		elseif IsNull(rsEntry("HasTGAccess")) AND rsEntry("NumRestrictions")>0 then
			failureReason = DisplayPhrase(phraseDictionary,"Registrationmembersonly")
		elseif SchHrsClosed then
			failureReason = DisplayPhrase(phraseDictionary,"Schedulingclosednow")
		elseif SameDaySchCutoff AND DateValue(rsEntry("ClassDateStart")) = DateValue(DateAdd("n", Session("tzOffset"),Now)) then
			failureReason = DisplayPhrase(phraseDictionary,"Schedulingclosednow")
		else
			failureReason = DisplayPhrase(phraseDictionary,"Registrationclosed")
		end if
                %>
                <div class="clear"></div>
		<div class="registration-closed" style="width:477px">
			<%=failureReason%>
			<br />
			<% if NOT isNull(contactEmail) then %>
                <%ReplaceInPhrase phraseDictionary,"Contactstudio","<STUDIONAME>",session("StudioName") %>
				<a href="mailto:<%=contactEmail%>"><img align="absbottom" src="<%= contentUrl("/asp/adm/images/icon_email.png") %>"><%=DisplayPhrase(phraseDictionary,"Contactstudio") %></a>
			<% end if %>
		</div>
<% end if %>

<%
end function

function conditionalImage(src)
	if Image.FileExists(studio_path & "\" & session("studioShort") & src) then 
		conditionalImage = "<img style=""float:right;"" src=""" & "/studios" & session("ClusterID") & "/" & Session("studioShort") & src & "?imageVersion=" & session("imageVersion") & """ />"
	end if 
end function %>

